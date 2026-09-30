import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, odataGuid, odataString, or } from "./query.js";
import { CATALOGS, REGISTERS, resolveEntity } from "../config/mapping.js";
import { requireEntity } from "./publication.js";
import { buildQuery } from "./query.js";
import {
  AggregateOverflowError,
  addMeta,
  emptyMeta,
  fetchAllForAggregation,
  type ScanMeta,
} from "./aggregate.js";
import type { ODataEntity } from "../types/odata.js";
import { InputError } from "../errors.js";

/**
 * Аналитика для 1С:Бухгалтерия 3.0 строится на регистре бухгалтерии
 * «Хозрасчетный» и его виртуальной таблице Balance. И дебиторка (сч. 62),
 * и остатки товаров (сч. 41) — это сальдо по соответствующему счёту,
 * с контрагентом/номенклатурой в ExtDimension1.
 */

const CHART_CANDIDATES = ["ChartOfAccounts_Хозрасчетный"] as const;

export interface Account {
  key: string;
  code: string;
  description: string;
}

/** Возвращает счета, код которых начинается с одного из префиксов (напр. "62", "41"). */
export async function resolveAccounts(conn: Connection, prefixes: readonly string[]): Promise<Account[]> {
  const chart = await requireEntity(conn, CHART_CANDIDATES, "План счетов «Хозрасчётный»");
  const filter = or(...prefixes.map((p) => `startswith(Code, ${odataString(p)})`));
  const { rows } = await fetchAll(
    conn.client,
    chart,
    { filter, select: ["Ref_Key", "Code", "Description"], orderby: "Code" },
    conn.behavior.pageSize,
    conn.behavior.maxRows,
  );
  return rows.map((r) => ({
    key: String(r["Ref_Key"] ?? ""),
    code: String(r["Code"] ?? ""),
    description: String(r["Description"] ?? ""),
  }));
}

/**
 * Возвращает карту «код счёта → Ref_Key» для заданных точных кодов
 * (напр. ["41.01","60.01","90.01.1"]). Нужно для заполнения счетов учёта
 * в документах — через OData автозаполнение 1С не срабатывает.
 */
export async function accountsByCode(
  conn: Connection,
  codes: readonly string[],
): Promise<Map<string, string>> {
  const result = new Map<string, string>();
  if (codes.length === 0) return result;
  const chart = await requireEntity(conn, CHART_CANDIDATES, "План счетов «Хозрасчётный»");
  const filter = or(...codes.map((c) => cmp("Code", "eq", odataString(c))));
  const { rows } = await fetchAll(
    conn.client,
    chart,
    { filter, select: ["Ref_Key", "Code"] },
    codes.length,
    codes.length,
  );
  for (const r of rows) result.set(String(r["Code"]), String(r["Ref_Key"]));
  return result;
}

/** Счета учёта номенклатуры (Ref_Key каждого; undefined если не задан). */
export interface NomAccounts {
  goods?: string; // СчетУчета (41.xx)
  incomingVat?: string; // НДС по приобретённым ценностям (19.xx)
  outgoingVat?: string; // НДС по реализации (90.03)
  income?: string; // доходы от реализации (90.01.x)
  expense?: string; // расходы от реализации / себестоимость (90.02.x)
}

const NOM_ACCOUNTS_REG = ["InformationRegister_СчетаУчетаНоменклатуры"] as const;
const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";
const norm = (g: unknown): string | undefined => {
  const s = typeof g === "string" ? g : "";
  return s && s !== EMPTY_GUID ? s : undefined;
};

/**
 * Берёт счета учёта для номенклатуры из регистра «Счета учёта номенклатуры»,
 * выбирая самую специфичную подходящую запись (пустые измерения — «джокер»).
 * Возвращает undefined, если регистр не опубликован (вызывающий откатится на коды).
 */
export async function nomenclatureAccounts(
  conn: Connection,
  orgKey: string,
  nomRef: string,
): Promise<NomAccounts | undefined> {
  const available = await conn.available();
  const reg = resolveEntity(NOM_ACCOUNTS_REG, available);
  if (!reg) return undefined;

  // Вид номенклатуры нужен для матчинга измерения ВидНоменклатуры в регистре.
  let vidRef: string | undefined;
  const nomSet = resolveEntity(CATALOGS.nomenclature, available);
  if (nomSet) {
    try {
      const item = await conn.client.getEntity(
        `${nomSet}(guid'${nomRef.replace(/[{}']/g, "")}')${buildQuery({ select: ["ВидНоменклатуры_Key"] })}`,
      );
      vidRef = norm(item["ВидНоменклатуры_Key"]);
    } catch {
      // нет такого поля/объекта — матчим без вида
    }
  }

  const { rows } = await fetchAll(
    conn.client,
    reg,
    {
      select: [
        "Организация_Key",
        "Номенклатура_Key",
        "ВидНоменклатуры_Key",
        "СчетУчета_Key",
        "СчетУчетаНДСПоПриобретеннымЦенностям_Key",
        "СчетУчетаНДСПоРеализации_Key",
        "СчетДоходовОтРеализации_Key",
        "СчетРасходовОтРеализации_Key",
      ],
    },
    conn.behavior.pageSize,
    500,
  );
  let best: ODataEntity | undefined;
  let bestScore = -1;
  for (const r of rows) {
    const nm = norm(r["Номенклатура_Key"]);
    const og = norm(r["Организация_Key"]);
    const vd = norm(r["ВидНоменклатуры_Key"]);
    if (nm && nm !== nomRef) continue;
    if (og && og !== orgKey) continue;
    if (vd && vd !== vidRef) continue; // запись для другого вида номенклатуры
    const score = (nm ? 4 : 0) + (vd ? 2 : 0) + (og ? 1 : 0);
    if (score > bestScore) {
      bestScore = score;
      best = r;
    }
  }
  if (!best) return undefined;
  return {
    goods: norm(best["СчетУчета_Key"]),
    incomingVat: norm(best["СчетУчетаНДСПоПриобретеннымЦенностям_Key"]),
    outgoingVat: norm(best["СчетУчетаНДСПоРеализации_Key"]),
    income: norm(best["СчетДоходовОтРеализации_Key"]),
    expense: norm(best["СчетРасходовОтРеализации_Key"]),
  };
}

/**
 * Тянет строки сальдо регистра Хозрасчетный, отфильтрованные по набору счетов
 * и (необязательно) по организации.
 */
export async function balanceByAccounts(
  conn: Connection,
  accountKeys: string[],
  orgKey?: string,
  asOf?: string,
): Promise<ODataEntity[]> {
  if (accountKeys.length === 0) return [];
  const reg = await requireEntity(conn, REGISTERS.accounting, "Регистр бухгалтерии «Хозрасчётный»");
  const filter =
    and(
      or(...accountKeys.map((k) => cmp("Account_Key", "eq", odataGuid(k)))),
      orgKey ? cmp("Организация_Key", "eq", odataGuid(orgKey)) : undefined,
    ) || undefined;
  // Параметр Period — НЕ через $filter, а path-параметром у виртуальной таблицы:
  // .../AccountingRegister_Хозрасчетный/Balance(Period=datetime'YYYY-MM-DDT23:59:59').
  // Без него виртуальная таблица возвращает текущее сальдо.
  const balancePath = asOf ? `${reg}/Balance(Period=datetime'${asOf}T23:59:59')` : `${reg}/Balance`;
  // Через безопасную выборку: сальдо берём ПОЛНОСТЬЮ (иначе дебиторка/остатки
  // занижаются), с громким переполнением вместо тихой обрезки.
  const { rows } = await fetchAllForAggregation(
    conn,
    balancePath,
    { filter },
    `сальдо на ${asOf ?? "сейчас"}`,
  );
  return rows;
}

/**
 * Виртуальная таблица «Остатки и обороты» регистра бухгалтерии через OData.
 *
 * ВСЁ, что зависит от точного синтаксиса 1С, собрано здесь, в одном месте, чтобы
 * после живого теста править одну константу, а не инструмент:
 *  - имя виртуальной таблицы (по документации платформы — BalanceAndTurnovers;
 *    если конкретная база ответит 404/400 — поменять name, напр. на BalanceAndTurnover);
 *  - имена path-параметров периода (StartPeriod/EndPeriod — как Period у Balance);
 *  - имена ресурсных полей (ресурс «Сумма» + суффиксы виртуальной таблицы).
 */
export const BALANCE_AND_TURNOVERS = {
  name: "BalanceAndTurnovers",
  startParam: "StartPeriod",
  endParam: "EndPeriod",
  fields: {
    openingDr: "СуммаOpeningBalanceDr",
    openingCr: "СуммаOpeningBalanceCr",
    turnoverDr: "СуммаTurnoverDr",
    turnoverCr: "СуммаTurnoverCr",
    closingDr: "СуммаClosingBalanceDr",
    closingCr: "СуммаClosingBalanceCr",
  },
  /** Поле периода: появляется в строках, только если 1С разбила результат по периодичности. */
  periodField: "Period",
} as const;

/**
 * Следующий календарный день для YYYY-MM-DD — чистая арифметика по компонентам,
 * без Date/таймзоны (new Date("…") + toISOString() может «съехать» на сутки).
 * Учитывает конец месяца, года и високосный февраль. Несуществующую дату
 * (напр. 2025-02-30 — формат её пропускает) отвергает, а не «перекатывает».
 */
export function nextDay(ymd: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ymd);
  if (!m) throw new InputError(`Дата должна быть в формате YYYY-MM-DD: ${ymd}`);
  let y = Number(m[1]);
  let mo = Number(m[2]);
  let d = Number(m[3]);
  const leap = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const dim = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][mo - 1];
  if (dim === undefined || d < 1 || d > dim) throw new InputError(`Несуществующая дата: ${ymd}`);
  if (d < dim) d += 1;
  else if (mo < 12) {
    mo += 1;
    d = 1;
  } else {
    y += 1;
    mo = 1;
    d = 1;
  }
  const pad = (n: number, w = 2): string => String(n).padStart(w, "0");
  return `${pad(y, 4)}-${pad(mo)}-${pad(d)}`;
}

/**
 * Путь к виртуальной таблице «Остатки и обороты» за [from,to] (YYYY-MM-DD, оба включительно).
 * Параметры — path-параметрами (как Period у Balance), а не через $filter.
 * EndPeriod — начало СЛЕДУЮЩЕГО дня после `to` (а не to 23:59:59), чтобы не терять
 * движения последней секунды дня:
 *   AccountingRegister_Хозрасчетный/BalanceAndTurnovers(StartPeriod=datetime'2026-08-01T00:00:00',EndPeriod=datetime'2026-09-01T00:00:00')
 */
export function balanceAndTurnoversPath(reg: string, from: string, to: string): string {
  const vt = BALANCE_AND_TURNOVERS;
  return (
    `${reg}/${vt.name}(` +
    `${vt.startParam}=datetime'${from}T00:00:00',` +
    `${vt.endParam}=datetime'${nextDay(to)}T00:00:00')`
  );
}

/**
 * Проверяет, что строки виртуальной таблицы содержат все шесть ресурсных полей.
 * Иначе (другое имя таблицы/ресурса в конкретной базе) num() тихо дал бы нули,
 * и неизвестная структура выдала бы «ОСВ из нулей» с consistent:true.
 * В ошибку идут только ИМЕНА полей строки, не значения.
 */
export function assertTurnoverFields(rows: readonly ODataEntity[]): void {
  const expected = Object.values(BALANCE_AND_TURNOVERS.fields);
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] as Record<string, unknown>;
    const missing = expected.filter((f) => !Object.prototype.hasOwnProperty.call(r, f));
    if (missing.length === 0) continue;
    throw new Error(
      `Виртуальная таблица ${BALANCE_AND_TURNOVERS.name} вернула строки несовместимой структуры ` +
        `(строка #${i + 1}): нет полей ${missing.join(", ")}. ` +
        `Фактические поля строки: ${Object.keys(r).join(", ") || "(нет)"}. ` +
        `Проверьте BALANCE_AND_TURNOVERS в src/odata/accounting.ts.`,
    );
  }
}

/** $filter по набору счетов и (необязательно) организации — фильтрует сама 1С. */
export function accountsOrgFilter(accountKeys: readonly string[], orgKey?: string): string | undefined {
  return (
    and(
      or(...accountKeys.map((k) => cmp("Account_Key", "eq", odataGuid(k)))),
      orgKey ? cmp("Организация_Key", "eq", odataGuid(orgKey)) : undefined,
    ) || undefined
  );
}

/**
 * Сколько GUID-ов счетов кладём в один $filter. Широкий префикс (напр. «9») даёт
 * десятки субсчетов — длинный OR упирается в лимит длины URL веб-сервера.
 */
export const ACCOUNT_FILTER_BATCH = 40;

/**
 * Строки «Остатков и оборотов» Хозрасчетного за период по набору счетов (+орг).
 *
 * Безопасность итога: каждая пачка счетов — через fetchAllForAggregation (громкое
 * переполнение), и суммарное число строк по всем пачкам тоже сверяется с тем же
 * потолком analyticsMaxRows. Частичной выборки наружу не уходит — только ошибка.
 */
export async function turnoversByAccounts(
  conn: Connection,
  accountKeys: readonly string[],
  from: string,
  to: string,
  orgKey?: string,
): Promise<{ rows: ODataEntity[]; meta: ScanMeta }> {
  if (accountKeys.length === 0) return { rows: [], meta: emptyMeta() };
  const reg = await requireEntity(conn, REGISTERS.accounting, "Регистр бухгалтерии «Хозрасчётный»");
  const path = balanceAndTurnoversPath(reg, from, to);
  const cap = conn.behavior.analyticsMaxRows;
  const period = `${from}..${to}`;

  let rows: ODataEntity[] = [];
  let meta = emptyMeta();
  for (let i = 0; i < accountKeys.length; i += ACCOUNT_FILTER_BATCH) {
    const batch = accountKeys.slice(i, i + ACCOUNT_FILTER_BATCH);
    // Без $select: 1С группирует виртуальную таблицу по выбранным полям, и
    // $select=Account_Key,… свернул бы сальдо по субконто. Нужны строки по всем
    // измерениям/субконто — как развёрнутое сальдо в стандартной ОСВ.
    const part = await fetchAllForAggregation(
      conn,
      path,
      { filter: accountsOrgFilter(batch, orgKey) },
      period,
    );
    rows = rows.concat(part.rows);
    meta = addMeta(meta, part.meta);
    if (rows.length > cap) throw new AggregateOverflowError(path, cap, period);
  }

  // Структура ответа должна совпадать с ожидаемой — до любой агрегации.
  assertTurnoverFields(rows);

  // Защита от молча неверного сальдо: если 1С вернула разбивку по периодичности,
  // начальные/конечные остатки подпериодов сложились бы с задвоением.
  const periods = new Set(
    rows.map((r) => r[BALANCE_AND_TURNOVERS.periodField]).filter((p) => p !== undefined && p !== null),
  );
  if (periods.size > 1) {
    throw new Error(
      `Виртуальная таблица ${BALANCE_AND_TURNOVERS.name} вернула разбивку по периодам ` +
        `(${periods.size} значений ${BALANCE_AND_TURNOVERS.periodField}) — остатки нельзя сложить. ` +
        `Проверьте параметры виртуальной таблицы (periodicity) в src/odata/accounting.ts.`,
    );
  }
  return { rows, meta };
}

/**
 * Резолвит GUID-ы в наименования (Description) из справочника батчами.
 * Используется, чтобы показать имена контрагентов/номенклатуры вместо GUID.
 */
export async function resolveNames(
  conn: Connection,
  entitySet: string,
  keys: Iterable<string>,
): Promise<Map<string, string>> {
  const unique = [...new Set([...keys].filter(Boolean))];
  const result = new Map<string, string>();
  const available = await conn.available();
  if (!available.has(entitySet)) return result;

  const BATCH = 20;
  for (let i = 0; i < unique.length; i += BATCH) {
    const batch = unique.slice(i, i + BATCH);
    const filter = or(...batch.map((k) => cmp("Ref_Key", "eq", odataGuid(k))));
    const { rows } = await fetchAll(
      conn.client,
      entitySet,
      { filter, select: ["Ref_Key", "Description"] },
      BATCH,
      BATCH,
    );
    for (const r of rows) result.set(String(r["Ref_Key"]), String(r["Description"] ?? ""));
  }
  return result;
}

/** Число из поля сальдо (OData может отдавать строкой). */
export function num(v: unknown): number {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
}

export { and, cmp };
