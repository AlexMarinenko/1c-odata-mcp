import { describe, it, expect } from "vitest";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import type { ServerContext } from "../src/context.js";
import {
  BALANCE_AND_TURNOVERS,
  dimensionRefOf,
  extDimensionFilter,
  rowMatchesDimension,
} from "../src/odata/accounting.js";
import { ODataError } from "../src/odata/errors.js";
import { aggregateTurnoverByContract, registerRegisterTools } from "../src/tools/registers.js";

const REG = "AccountingRegister_Хозрасчетный";
const CHART = "ChartOfAccounts_Хозрасчетный";
const ORGS = "Catalog_Организации";
const CP = "Catalog_Контрагенты";
const CONTRACTS = "Catalog_ДоговорыКонтрагентов";
const DOC = "Document_ПоступлениеТоваровУслуг";
const T = (set: string): string => `StandardODATA.${set}`;

const guid = (n: number, p = "00000000"): string => `${p}-0000-0000-0000-${String(n).padStart(12, "0")}`;
const A6001 = guid(1);
const A6002 = guid(2);
const A6201 = guid(3);
const A6202 = guid(4);
const CLINIC = guid(1, "aaaaaaaa"); // организация
const IP = guid(1, "bbbbbbbb"); // контрагент — ИП
const OTHER = guid(2, "bbbbbbbb"); // другой контрагент
const D1 = guid(1, "cccccccc");
const D2 = guid(2, "cccccccc");
const D_OTHER = guid(9, "cccccccc");

const f = BALANCE_AND_TURNOVERS.fields;
type Sums = Partial<Record<keyof typeof f, number>>;

/** Строка «Остатков и оборотов»: счёт × контрагент × договор × документ расчётов. */
function row(acc: string, cp: string, contract: string | null, doc: number, v: Sums, typed = true) {
  return {
    Account_Key: acc,
    Организация_Key: CLINIC,
    ExtDimension1: cp,
    ExtDimension2: contract ?? "00000000-0000-0000-0000-000000000000",
    ExtDimension3: guid(doc, "dddddddd"),
    ...(typed
      ? { ExtDimension1_Type: T(CP), ExtDimension2_Type: T(CONTRACTS), ExtDimension3_Type: T(DOC) }
      : {}),
    ...Object.fromEntries(Object.values(f).map((name) => [name, 0])),
    ...Object.fromEntries(Object.entries(v).map(([k, val]) => [f[k as keyof typeof f], val])),
  } as Record<string, unknown>;
}

/**
 * Расчёты клиники с ИП за январь по счёту 60 (и «шум» — другой контрагент):
 *  договор D1: долг по документу 1 (60.01, Кт) и аванс по документу 2 (60.02, Дт);
 *  договор D2: долг по документу 3 (60.01, Кт);
 *  без договора: аванс по документу 4 (60.02, Дт).
 */
const ROWS = [
  row(A6001, IP, D1, 1, { openingCr: 1000, turnoverDr: 400, turnoverCr: 700, closingCr: 1300 }),
  row(A6002, IP, D1, 2, { openingDr: 250, turnoverDr: 150, closingDr: 400 }),
  row(A6001, IP, D2, 3, { turnoverCr: 90.1, closingCr: 90.1 }),
  row(A6002, IP, null, 4, { openingDr: 30, turnoverCr: 10, closingDr: 20 }),
  row(A6001, OTHER, D_OTHER, 5, { openingCr: 5000, turnoverCr: 5000, closingCr: 10000 }),
];

type Mode = "apply" | "reject" | "ignore";

function fakeCtx(regRows: Array<Record<string, unknown>>, mode: Mode = "apply") {
  const paths: string[] = [];
  const chart = [
    { Ref_Key: A6001, Code: "60.01", Description: "Расчеты с поставщиками" },
    { Ref_Key: A6002, Code: "60.02", Description: "Авансы выданные" },
    { Ref_Key: A6201, Code: "62.01", Description: "Расчеты с покупателями" },
    { Ref_Key: A6202, Code: "62.02", Description: "Авансы полученные" },
  ];
  const catalogs: Record<string, Array<{ Ref_Key: string; Description: string }>> = {
    [CP]: [
      { Ref_Key: IP, Description: "ИП Тестовый" },
      { Ref_Key: OTHER, Description: "ООО Другой" },
    ],
    [CONTRACTS]: [
      { Ref_Key: D1, Description: "Договор аренды" },
      { Ref_Key: D2, Description: "Договор услуг" },
    ],
  };
  const client = {
    async getCollection(path: string): Promise<{ value: unknown[] }> {
      paths.push(path);
      const dec = decodeURIComponent(path);
      const top = Number(/\$top=(\d+)/.exec(path)?.[1] ?? 1000);
      const skip = Number(/\$skip=(\d+)/.exec(path)?.[1] ?? 0);
      let pool: unknown[] = [];
      if (path.startsWith(CHART)) {
        const pfx = /startswith\(Code, '([^']*)'\)/.exec(dec)?.[1] ?? "";
        pool = chart.filter((a) => a.Code.startsWith(pfx));
      } else if (path.startsWith(ORGS)) {
        pool = [{ Ref_Key: CLINIC, Description: "ООО Клиника" }];
      } else if (path.startsWith(REG)) {
        const keys = [...dec.matchAll(/Account_Key eq guid'([^']+)'/g)].map((m) => m[1]);
        pool = regRows.filter((r) => keys.includes(String(r["Account_Key"])));
        // Группы вида (ExtDimension1 eq cast(guid'…', 'Catalog_…') or ExtDimension2 eq …)
        const dims = [...dec.matchAll(/ExtDimension\d eq cast\(guid'([^']+)', '([^']+)'\)/g)];
        if (dims.length > 0) {
          if (mode === "reject") {
            throw new ODataError({
              kind: "bad_request",
              message: "Некорректный запрос к OData (400)",
              status: 400,
            });
          }
          if (mode === "apply") {
            const wanted = new Map(dims.map((m) => [m[2]!, m[1]!]));
            pool = (pool as Array<Record<string, unknown>>).filter((r) =>
              [...wanted.entries()].every(([set, ref]) =>
                [1, 2, 3].some((n) => r[`ExtDimension${n}`] === ref && r[`ExtDimension${n}_Type`] === T(set)),
              ),
            );
          }
        }
      } else {
        const set = Object.keys(catalogs).find((s) => path.startsWith(s));
        if (set) {
          const keys = [...dec.matchAll(/Ref_Key eq guid'([^']+)'/g)].map((m) => m[1]!.toLowerCase());
          pool = catalogs[set]!.filter((c) => keys.includes(c.Ref_Key.toLowerCase()));
        }
      }
      return { value: pool.slice(skip, skip + top) };
    },
  };
  const conn = {
    cfg: { name: "buh" },
    client,
    behavior: { analyticsMaxRows: 200_000, pageSize: 100, maxRows: 1000 },
    available: async () => new Set([REG, CHART, ORGS, CP, CONTRACTS]),
  };
  return { ctx: { db: () => conn } as unknown as ServerContext, paths };
}

async function callTool(ctx: ServerContext, args: Record<string, unknown>) {
  const server = new McpServer({ name: "t", version: "0" });
  registerRegisterTools(server, ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "c", version: "0" });
  await client.connect(b);
  try {
    return await client.callTool({ name: "read.accounting.get_account_turnover", arguments: args });
  } finally {
    await client.close();
  }
}

const JAN = { account: "60", from: "2025-01-01", to: "2025-01-31" };
const IP_JAN = { ...JAN, counterpartyRef: IP };
// eslint-disable-next-line @typescript-eslint/no-explicit-any -- разбор произвольного JSON-ответа в тесте
type Out = Record<string, any>;
const out = (r: Awaited<ReturnType<typeof callTool>>): Out => r.structuredContent as Out;
const errText = (r: Awaited<ReturnType<typeof callTool>>): string =>
  (r.content as Array<{ text: string }>)[0]!.text;

/** Ожидаемая ОСВ клиники с ИП по счёту 60 за январь. */
const EXPECTED_IP = {
  openingDebit: 280, // 250 (аванс по D1) + 30 (аванс без договора)
  openingCredit: 1000,
  debitTurnover: 550,
  creditTurnover: 800.1,
  closingDebit: 420,
  closingCredit: 1390.1,
  consistent: true,
};

describe("отбор ОСВ по контрагенту (counterpartyRef)", () => {
  it("остаются только строки контрагента; сальдо развёрнутое: долг и аванс не сворачиваются", async () => {
    const { ctx } = fakeCtx(ROWS);
    const s = out(await callTool(ctx, IP_JAN));
    expect(s).toMatchObject(EXPECTED_IP);
    // свёрнутое сальдо было бы 970.1 Кт и 0 Дт — так быть не должно
    expect(s["closingDebit"]).not.toBe(0);
    expect(s["counterparty"]).toEqual({ ref: IP, name: "ИП Тестовый" });
    expect(s["rowsMatched"]).toBe(4);
    expect(s["accounts"].map((a: Out) => [a.code, a.closingDebit, a.closingCredit])).toEqual([
      ["60.01", 0, 1390.1],
      ["60.02", 420, 0],
    ]);
  });

  it("отбор уходит в $filter 1С через cast по всем позициям субконто", async () => {
    const { ctx, paths } = fakeCtx(ROWS);
    const s = out(await callTool(ctx, IP_JAN));
    expect(s["dimensionFilter"]).toBe("server");
    const call = decodeURIComponent(paths.find((p) => p.startsWith(REG))!);
    for (const n of [1, 2, 3]) {
      expect(call).toContain(`ExtDimension${n} eq cast(guid'${IP}', 'Catalog_Контрагенты')`);
    }
    expect(call).not.toContain("$select"); // иначе 1С свернула бы аналитику
  });

  it("база отвергает $filter по субконто (400): отбор выполняется на стороне сервера, итог тот же", async () => {
    const { ctx, paths } = fakeCtx(ROWS, "reject");
    const s = out(await callTool(ctx, IP_JAN));
    expect(s).toMatchObject(EXPECTED_IP);
    expect(s["dimensionFilter"]).toBe("client");
    expect(paths.filter((p) => p.startsWith(REG))).toHaveLength(2); // с фильтром и без
  });

  it("если 1С вернула чужие строки, они не попадают в итог", async () => {
    const { ctx } = fakeCtx(ROWS, "ignore");
    const s = out(await callTool(ctx, IP_JAN));
    expect(s).toMatchObject(EXPECTED_IP);
    expect(s["scan"].rowsScanned).toBe(5);
    expect(s["rowsMatched"]).toBe(4);
  });

  it("без отбора результат прежний: все контрагенты, без служебных полей отбора", async () => {
    const { ctx } = fakeCtx(ROWS);
    const s = out(await callTool(ctx, JAN));
    expect(s["closingCredit"]).toBe(11390.1);
    expect(s).not.toHaveProperty("dimensionFilter");
    expect(s).not.toHaveProperty("counterparty");
    expect(s).not.toHaveProperty("contracts");
  });

  it("несуществующий контрагент — ошибка, а не ОСВ из нулей", async () => {
    const { ctx, paths } = fakeCtx(ROWS);
    const r = await callTool(ctx, { ...JAN, counterpartyRef: guid(777, "bbbbbbbb") });
    expect(r.isError).toBe(true);
    expect(errText(r)).toMatch(/не найден/);
    expect(paths.some((p) => p.startsWith(REG))).toBe(false);
  });

  it("не GUID в counterpartyRef отсекается валидацией входа", async () => {
    const { ctx } = fakeCtx(ROWS);
    const r = await callTool(ctx, { ...JAN, counterpartyRef: "ИП Тестовый" });
    expect(r.isError).toBe(true);
  });

  it("у контрагента нет расчётов по счёту: нули и пояснение", async () => {
    const { ctx } = fakeCtx(ROWS);
    const s = out(await callTool(ctx, { ...IP_JAN, account: "62" }));
    expect(s).toMatchObject({ openingDebit: 0, closingCredit: 0, consistent: true, rowsMatched: 0 });
    expect(s["note"]).toMatch(/с такой аналитикой/);
  });
});

describe("отбор и разбивка по договорам", () => {
  it("breakdown=contract: договоры, строки без договора, субсчета внутри договора", async () => {
    const { ctx } = fakeCtx(ROWS);
    const s = out(await callTool(ctx, { ...IP_JAN, breakdown: "contract" }));
    const byName = Object.fromEntries(s["contracts"].map((c: Out) => [c.name, c]));
    expect(Object.keys(byName).sort()).toEqual(["(без договора)", "Договор аренды", "Договор услуг"]);
    // D1: долг по одному документу и аванс по другому остаются развёрнутыми
    expect(byName["Договор аренды"]).toMatchObject({
      ref: D1,
      openingDebit: 250,
      openingCredit: 1000,
      closingDebit: 400,
      closingCredit: 1300,
      consistent: true,
      rows: 2,
    });
    expect(byName["Договор аренды"].accounts.map((a: Out) => a.code)).toEqual(["60.01", "60.02"]);
    expect(byName["Договор услуг"]).toMatchObject({
      creditTurnover: 90.1,
      closingCredit: 90.1,
      closingDebit: 0,
    });
    expect(byName["(без договора)"]).toMatchObject({ ref: null, closingDebit: 20, rows: 1 });
  });

  it("сумма по договорам равна итогу ОСВ по каждому показателю", async () => {
    const { ctx } = fakeCtx(ROWS);
    const s = out(await callTool(ctx, { ...IP_JAN, breakdown: "contract" }));
    for (const k of [
      "openingDebit",
      "openingCredit",
      "debitTurnover",
      "creditTurnover",
      "closingDebit",
      "closingCredit",
    ]) {
      const cents = s["contracts"].reduce((acc: number, c: Out) => acc + Math.round(c[k] * 100), 0);
      expect(cents).toBe(Math.round(s[k] * 100));
    }
  });

  it("contractRef оставляет один договор; вместе с counterpartyRef оба условия уходят в 1С", async () => {
    const { ctx, paths } = fakeCtx(ROWS);
    const s = out(await callTool(ctx, { ...IP_JAN, contractRef: D1 }));
    expect(s).toMatchObject({
      openingDebit: 250,
      openingCredit: 1000,
      closingDebit: 400,
      closingCredit: 1300,
    });
    expect(s["contract"]).toEqual({ ref: D1, name: "Договор аренды" });
    const call = decodeURIComponent(paths.find((p) => p.startsWith(REG))!);
    expect(call).toContain(`cast(guid'${IP}', 'Catalog_Контрагенты')`);
    expect(call).toContain(`cast(guid'${D1}', 'Catalog_ДоговорыКонтрагентов')`);
  });

  it("разбивка по договорам без полей типа субконто — ошибка, а не догадка по номеру субконто", async () => {
    const untyped = ROWS.map((r) => row(r["Account_Key"] as string, IP, D1, 1, { closingCr: 1 }, false));
    const { ctx } = fakeCtx(untyped, "ignore");
    const r = await callTool(ctx, { ...IP_JAN, breakdown: "contract" });
    expect(r.isError).toBe(true);
    expect(errText(r)).toMatch(/полей типа субконто/);
  });

  it("отбор по контрагенту работает и без полей типа (сравнение по GUID)", async () => {
    const untyped = [
      row(A6001, IP, D1, 1, { closingCr: 7 }, false),
      row(A6001, OTHER, D1, 2, { closingCr: 9 }, false),
    ];
    const { ctx } = fakeCtx(untyped, "ignore");
    expect(out(await callTool(ctx, IP_JAN))["closingCredit"]).toBe(7);
  });
});

describe("вспомогательные функции субконто", () => {
  it("extDimensionFilter: три позиции через OR, GUID и имя набора экранируются", () => {
    expect(extDimensionFilter({ entitySet: CP, ref: IP })).toBe(
      `(ExtDimension1 eq cast(guid'${IP}', 'Catalog_Контрагенты') or ` +
        `ExtDimension2 eq cast(guid'${IP}', 'Catalog_Контрагенты') or ` +
        `ExtDimension3 eq cast(guid'${IP}', 'Catalog_Контрагенты'))`,
    );
    expect(() => extDimensionFilter({ entitySet: CP, ref: "x' or 1 eq 1" })).toThrow();
  });

  it("тот же GUID в субконто другого типа не считается совпадением", () => {
    const r = { ExtDimension1: IP, ExtDimension1_Type: T(CONTRACTS) };
    expect(rowMatchesDimension(r, { entitySet: CP, ref: IP })).toBe(false);
    expect(
      rowMatchesDimension({ ...r, ExtDimension1_Type: T(CP) }, { entitySet: CP, ref: IP.toUpperCase() }),
    ).toBe(true);
  });

  it("dimensionRefOf находит договор в любой позиции по типу; пустая ссылка — нет договора", () => {
    expect(dimensionRefOf({ ExtDimension3: D2, ExtDimension3_Type: T(CONTRACTS) }, CONTRACTS)).toBe(D2);
    expect(dimensionRefOf(row(A6001, IP, null, 1, {}), CONTRACTS)).toBeUndefined();
    expect(dimensionRefOf({ ExtDimension2: D2 }, CONTRACTS)).toBeUndefined();
  });

  it("aggregateTurnoverByContract: контроль Сн + ОбДт − ОбКт = Ск по каждому договору", () => {
    const accounts = [
      { key: A6001, code: "60.01", description: "" },
      { key: A6002, code: "60.02", description: "" },
    ];
    const broken = [...ROWS.slice(0, 2), row(A6001, IP, D2, 3, { turnoverCr: 90.1, closingCr: 80 })];
    const groups = aggregateTurnoverByContract(broken, accounts as never, CONTRACTS);
    expect(groups.map((g) => [g.contractRef, g.consistent])).toEqual([
      [D1, true],
      [D2, false],
    ]);
  });
});
