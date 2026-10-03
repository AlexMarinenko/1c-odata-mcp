/**
 * Диагностика подключения и сверка взаиморасчётов двух сторон по ОСВ (только чтение).
 *
 * Запуск на сервере MCP, из каталога проекта:
 *   node --env-file=.env dist/scripts/probe-settlements.js                # только диагностика
 *   node --env-file=.env dist/scripts/probe-settlements.js \
 *     --org-a "<клиника>" --org-b "<ИП>" \
 *     --cp-in-a "<ИП как контрагент>" --cp-in-b "<клиника как контрагент>" \
 *     --from 2025-01-01 --to 2025-01-31 [--monthly] [--database buh]
 *
 * Сторона A учитывает сторону B как поставщика (счёт 60), сторона B учитывает сторону A как
 * покупателя (счёт 62). --org-* — часть названия организации в базе; --cp-in-a — как сторона B
 * названа в справочнике контрагентов (часть названия или Ref_Key), --cp-in-b — наоборот.
 * Если сторона B не ведётся организацией в этой базе, --org-b и --cp-in-b не указывайте:
 * будет показана только сторона A.
 *
 * Секретов скрипт не печатает: адрес базы выводится без учётных данных.
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { loadConfig } from "../config/env.js";
import { CATALOGS, REGISTERS, resolveEntity } from "../config/mapping.js";
import { ServerContext, type Connection } from "../context.js";
import { createServer } from "../mcp/server.js";
import { BALANCE_AND_TURNOVERS, GUID_RE, balanceAndTurnoversPath } from "../odata/accounting.js";
import { splitByMonth } from "../odata/aggregate.js";
import { listOrganizations } from "../odata/orgs.js";
import { fetchAll } from "../odata/pagination.js";
import { buildQuery, contains } from "../odata/query.js";

const TOOL = "read.accounting.get_account_turnover";
/** Имена, которые запрашивал клиент и которых в этом сервере нет. */
const MISSING_NAMES = ["get_accounting_balance", "get_accounting_register"];

const line = (s = ""): void => void process.stdout.write(`${s}\n`);

function parseArgs(argv: string[]): Record<string, string> {
  const out: Record<string, string> = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]!;
    if (!a.startsWith("--")) continue;
    const next = argv[i + 1];
    if (next === undefined || next.startsWith("--")) out[a.slice(2)] = "true";
    else {
      out[a.slice(2)] = next;
      i++;
    }
  }
  return out;
}

const cents = (v: unknown): number => Math.round(Number(v ?? 0) * 100);
const rub = (c: number): string => (c / 100).toFixed(2);

interface Sums {
  openingDebit: number;
  openingCredit: number;
  debitTurnover: number;
  creditTurnover: number;
  closingDebit: number;
  closingCredit: number;
}
interface Turnover extends Sums {
  consistent: boolean;
  dimensionFilter?: string;
  rowsMatched?: number;
  accounts: Array<Sums & { code: string }>;
  contracts?: Array<Sums & { name: string | null; ref: string | null }>;
  scan: { rowsScanned: number; elapsedMs: number };
}

const sumsLine = (s: Sums): string =>
  `Сн Дт ${rub(cents(s.openingDebit))} Кт ${rub(cents(s.openingCredit))} | ` +
  `Об Дт ${rub(cents(s.debitTurnover))} Кт ${rub(cents(s.creditTurnover))} | ` +
  `Ск Дт ${rub(cents(s.closingDebit))} Кт ${rub(cents(s.closingCredit))}`;

/** Шаг 1. Какие базы настроены и похожи ли они на бухгалтерию. */
async function diagnoseDatabases(ctx: ServerContext): Promise<void> {
  line("== 1. Настроенные базы (адрес без учётных данных)");
  const cfg = loadConfig();
  for (const c of cfg.connections) {
    const url = new URL(c.baseUrl);
    const mark = c.name === cfg.defaultName ? " [по умолчанию]" : "";
    line(`- ${c.name}${mark}: ${url.host}${url.pathname}`);
    try {
      const available = await ctx.db(c.name).available();
      const has = (candidates: readonly string[]): string =>
        resolveEntity(candidates, available) ? "да" : "НЕТ";
      const accounting = resolveEntity(REGISTERS.accounting, available) !== undefined;
      line(
        `    сущностей: ${available.size}; регистр «Хозрасчетный»: ${accounting ? "да" : "НЕТ"}; ` +
          `контрагенты: ${has(CATALOGS.counterparties)}; договоры: ${has(CATALOGS.contracts)}`,
      );
      line(
        accounting
          ? "    вывод: бухгалтерская база, ОСВ доступна"
          : "    вывод: ЭТО НЕ БУХГАЛТЕРИЯ (нет регистра бухгалтерии) — ОСВ по счетам 60/62 здесь получить нельзя",
      );
    } catch (e) {
      line(`    метаданные не получены: ${e instanceof Error ? e.message : String(e)}`);
    }
  }
  if (cfg.connections.length === 1) {
    line("  Настроена одна база: параметр database инструментов ничего не переключает.");
  }
}

/** Шаг 2. Какие инструменты реально есть у этого сервера. */
async function diagnoseTools(client: Client): Promise<void> {
  line("\n== 2. Инструменты этого сервера");
  const { tools } = await client.listTools();
  const names = tools.map((t) => t.name);
  line(
    `всего: ${names.length}; бухгалтерские: ${names.filter((n) => n.startsWith("read.accounting.")).join(", ")}`,
  );
  for (const missing of MISSING_NAMES) {
    line(
      `- ${missing}: ${names.includes(missing) ? "есть" : "НЕТ в этом сервере (и не было ни в одной версии)"}`,
    );
  }
  line(
    `- инструменты произвольного OData-запроса (query, register_query): ${names.includes("query") ? "есть" : "НЕТ"}`,
  );
  line(
    "  Они принадлежат другому пакету — @1c-odata/mcp. У него тот же исполняемый файл `1c-odata-mcp`,\n" +
      "  но свои подключения (`1c-odata-mcp list`) и нет бухгалтерских инструментов.",
  );
  const tool = tools.find((t) => t.name === TOOL);
  const props = Object.keys((tool?.inputSchema as { properties?: object } | undefined)?.properties ?? {});
  line(`- ${TOOL}: параметры ${props.join(", ")}`);
  line(
    `  отбор по контрагенту и договору: ${props.includes("counterpartyRef") && props.includes("breakdown") ? "есть" : "НЕТ (старая сборка)"}`,
  );
}

/** Шаг 3. Структура строки «Остатков и оборотов»: только имена полей, без значений. */
async function diagnoseRowShape(conn: Connection): Promise<void> {
  line("\n== 3. Структура виртуальной таблицы «Остатки и обороты»");
  const reg = resolveEntity(REGISTERS.accounting, await conn.available());
  if (!reg) return line("регистр бухгалтерии не опубликован — шаг пропущен");
  try {
    const path = balanceAndTurnoversPath(reg, "2025-01-01", "2025-01-31") + buildQuery({ top: 1 });
    const page = await conn.client.getCollection(path);
    const row = (page.value?.[0] ?? undefined) as Record<string, unknown> | undefined;
    if (!row) return line("за январь 2025 строк нет — структуру проверить не на чем");
    const keys = Object.keys(row);
    line(`поля: ${keys.join(", ")}`);
    const resources = Object.values(BALANCE_AND_TURNOVERS.fields).every((f) => keys.includes(f));
    line(`шесть ресурсов остатков и оборотов: ${resources ? "есть" : "НЕТ — ОСВ работать не будет"}`);
    line(
      `поля типа субконто (ExtDimensionN_Type): ${keys.some((k) => /^ExtDimension\d_Type$/.test(k)) ? "есть" : "НЕТ — разбивка по договорам невозможна"}`,
    );
  } catch (e) {
    line(`запрос не выполнен: ${e instanceof Error ? e.message : String(e)}`);
  }
}

async function findCounterparty(conn: Connection, query: string, label: string): Promise<string | undefined> {
  if (GUID_RE.test(query)) return query;
  const set = resolveEntity(CATALOGS.counterparties, await conn.available());
  if (!set) {
    line(`${label}: справочник контрагентов не опубликован`);
    return undefined;
  }
  const { rows } = await fetchAll(
    conn.client,
    set,
    {
      filter: contains("Description", query),
      select: ["Ref_Key", "Description", "IsFolder", "DeletionMark"],
    },
    50,
    50,
  );
  const items = rows.filter((r) => r["IsFolder"] !== true);
  line(`${label}: по «${query}» найдено ${items.length}`);
  for (const r of items) {
    line(
      `    ${String(r["Ref_Key"])}  ${String(r["Description"])}${r["DeletionMark"] === true ? "  [помечен на удаление]" : ""}`,
    );
  }
  if (items.length === 1) return String(items[0]!["Ref_Key"]);
  if (items.length > 1) line("    найдено несколько — повторите запуск с нужным Ref_Key вместо названия");
  return undefined;
}

async function turnover(client: Client, args: Record<string, unknown>): Promise<Turnover | undefined> {
  const res = await client.callTool({ name: TOOL, arguments: args });
  if (res.isError) {
    line(`    ОШИБКА: ${(res.content as Array<{ text?: string }>)[0]?.text ?? "без текста"}`);
    return undefined;
  }
  return res.structuredContent as unknown as Turnover;
}

function printTurnover(title: string, t: Turnover): void {
  line(`  ${title}`);
  line(`    итого   ${sumsLine(t)}`);
  for (const a of t.accounts) line(`    ${a.code.padEnd(7)} ${sumsLine(a)}`);
  for (const c of t.contracts ?? []) line(`    договор «${c.name ?? c.ref ?? "?"}»: ${sumsLine(c)}`);
  line(
    `    контроль Сн + ОбДт − ОбКт = Ск: ${t.consistent ? "сходится" : "НЕ СХОДИТСЯ"}; ` +
      `отбор по контрагенту: ${t.dimensionFilter ?? "нет"}; строк: ${t.rowsMatched ?? 0} из ${t.scan.rowsScanned}; ${t.scan.elapsedMs} мс`,
  );
}

/** Зеркальная сверка: что у A по счёту 60, то у B по счёту 62 с обратной стороной. */
function printMirror(a: Turnover, b: Turnover): void {
  const pairs: Array<[string, number, number]> = [
    ["начальный долг A перед B (60 Кт ↔ 62 Дт)", cents(a.openingCredit), cents(b.openingDebit)],
    ["начальные авансы A у B (60 Дт ↔ 62 Кт)", cents(a.openingDebit), cents(b.openingCredit)],
    ["оплаты и зачёты (60 ОбДт ↔ 62 ОбКт)", cents(a.debitTurnover), cents(b.creditTurnover)],
    ["поставки и возвраты (60 ОбКт ↔ 62 ОбДт)", cents(a.creditTurnover), cents(b.debitTurnover)],
    ["конечный долг A перед B (60 Кт ↔ 62 Дт)", cents(a.closingCredit), cents(b.closingDebit)],
    ["конечные авансы A у B (60 Дт ↔ 62 Кт)", cents(a.closingDebit), cents(b.closingCredit)],
  ];
  line("  зеркальная сверка (сторона A ↔ сторона B):");
  for (const [name, x, y] of pairs) {
    line(
      `    ${x === y ? "совпало " : "РАЗНИЦА "} ${name}: ${rub(x)} и ${rub(y)}${x === y ? "" : `, разница ${rub(x - y)}`}`,
    );
  }
  // Развёрнутые остатки зависят от того, как каждая сторона разнесла оплаты по документам,
  // поэтому дополнительно сравнивается свёрнутое сальдо расчётов.
  const netA = cents(a.closingCredit) - cents(a.closingDebit);
  const netB = cents(b.closingDebit) - cents(b.closingCredit);
  line(
    `    ${netA === netB ? "совпало " : "РАЗНИЦА "} свёрнутое конечное сальдо (A должна B): ${rub(netA)} и ${rub(netB)}`,
  );
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  const ctx = new ServerContext(loadConfig());
  const server = createServer(ctx);
  const [s, c] = InMemoryTransport.createLinkedPair();
  await server.connect(s);
  const client = new Client({ name: "probe-settlements", version: "0" });
  await client.connect(c);

  await diagnoseDatabases(ctx);
  await diagnoseTools(client);

  const database = args["database"];
  let conn: Connection;
  try {
    conn = ctx.db(database);
  } catch (e) {
    line(`\nБаза не выбрана: ${e instanceof Error ? e.message : String(e)}`);
    return void (await client.close());
  }
  line(`\nРабочая база для шагов 3–5: ${conn.cfg.name}`);
  await diagnoseRowShape(conn);

  line("\n== 4. Организации базы");
  try {
    for (const o of await listOrganizations(conn)) line(`- ${o.name}`);
  } catch (e) {
    line(`не получены: ${e instanceof Error ? e.message : String(e)}`);
  }

  if (!args["org-a"] || !args["cp-in-a"]) {
    line("\nСверка не запускалась: укажите --org-a и --cp-in-a (см. шапку скрипта).");
    return void (await client.close());
  }

  const from = args["from"] ?? "2025-01-01";
  const to = args["to"] ?? "2025-01-31";
  line(`\n== 5. Сверка взаиморасчётов, ${from}…${to}`);
  const cpInA = await findCounterparty(conn, args["cp-in-a"], "сторона B как контрагент в учёте A");
  const hasB = Boolean(args["org-b"] && args["cp-in-b"]);
  const cpInB = hasB
    ? await findCounterparty(conn, args["cp-in-b"]!, "сторона A как контрагент в учёте B")
    : undefined;
  if (!cpInA || (hasB && !cpInB)) {
    line("Контрагент не определён однозначно — сверка остановлена.");
    return void (await client.close());
  }

  const windows = args["monthly"] === "true" ? splitByMonth(from, to) : [{ from, to }];
  const dbArg = database ? { database } : {};
  let differences = 0;
  for (const w of windows) {
    line(`\n— период ${w.from}…${w.to}`);
    const a = await turnover(client, {
      ...dbArg,
      organization: args["org-a"],
      account: "60",
      from: w.from,
      to: w.to,
      counterpartyRef: cpInA,
      breakdown: "contract",
    });
    if (a) printTurnover(`A: «${args["org-a"]}», счёт 60, контрагент — сторона B`, a);
    if (!hasB) continue;
    const b = await turnover(client, {
      ...dbArg,
      organization: args["org-b"],
      account: "62",
      from: w.from,
      to: w.to,
      counterpartyRef: cpInB,
      breakdown: "contract",
    });
    if (b) printTurnover(`B: «${args["org-b"]}», счёт 62, контрагент — сторона A`, b);
    if (a && b) {
      printMirror(a, b);
      if (cents(a.closingCredit) - cents(a.closingDebit) !== cents(b.closingDebit) - cents(b.closingCredit))
        differences++;
    }
  }
  if (hasB)
    line(`\nИТОГ: периодов с расхождением свёрнутого конечного сальдо: ${differences} из ${windows.length}`);
  await client.close();
}

main().catch((e) => {
  line(`Скрипт прерван: ${e instanceof Error ? e.message : String(e)}`);
  process.exit(1);
});
