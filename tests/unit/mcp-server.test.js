import { afterEach, describe, expect, test } from "vitest";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import {
  addWorksheet,
  createBlankWorkbook,
  getWorkbookCalculation,
} from "../../src/app/03-workbook-model.js";
import {
  buildAssistantSelectionContext,
  createApprovedSelectionCalculator,
  validateAssistantFormula,
  validateAssistantProposal,
  validateProposedCalculation,
} from "../../src/app/19b-assistant-protocol.js";
import { createWorkbookMcpServer } from "../../src/mcp/server.js";

function ledger() {
  const workbook = createBlankWorkbook("Ledger");
  const sheet = workbook.sheets[0];
  Object.assign(sheet.cells, {
    A1: { raw: "Item" },
    B1: { raw: "Amount" },
    A2: { raw: "Rent" },
    B2: { raw: 1200 },
    A3: { raw: "Power" },
    B3: { raw: "85.50" },
    A4: { raw: "Total" },
    B4: { raw: "=B2+B3" },
    C4: { raw: "=1/0" },
    D9: { raw: "outside the selection" },
  });
  return workbook;
}

const open = [];
afterEach(async () => {
  for (const close of open.splice(0)) await close();
});

async function connect(workbook, range = "A1:C4") {
  const queued = [];
  const server = createWorkbookMcpServer({
    getWorkbook: () => workbook,
    getSelection: () => ({ sheetId: workbook.sheets[0].id, range }),
    queueProposal: (proposal) => queued.push(proposal),
  });
  const client = new Client({ name: "unit", version: "1" });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  await client.connect(b);
  open.push(async () => {
    await client.close();
    await server.close();
  });
  const call = async (name, args = {}) => {
    const result = await client.callTool({ name, arguments: args });
    return result.isError
      ? { error: result.content[0].text }
      : result.structuredContent;
  };
  return { client, call, queued };
}

const proposalFor = (workbook, patches, extra = {}) => ({
  id: "p-1",
  workbookId: workbook.id,
  revision: workbook.revision,
  explanation: "Store the amount as a number.",
  citations: ["B3"],
  patches,
  ...extra,
});

describe("MCP server", () => {
  test("offers nine tools", async () => {
    const { client } = await connect(ledger());
    const { tools } = await client.listTools();
    expect(tools.map((tool) => tool.name).sort()).toEqual([
      "workbook_check_formula_compatibility",
      "workbook_evaluate_formula",
      "workbook_get_capabilities",
      "workbook_get_context",
      "workbook_get_formula_contracts",
      "workbook_inspect_formula",
      "workbook_queue_proposal",
      "workbook_read_range",
      "workbook_run_audit_checks",
    ]);
  });

  test("capabilities state the limits and that nothing is written", async () => {
    const { call } = await connect(ledger());
    const capabilities = await call("workbook_get_capabilities");
    expect(capabilities.limits.maxSelectionCells).toBe(500);
    expect(capabilities.authority.writes).toMatch(/never applied/);
    expect(capabilities.customFunctions.execution).toMatch(/cannot install/);
  });

  test("context holds only the approved selection, with calculated values", async () => {
    const workbook = ledger();
    const { call } = await connect(workbook);
    const context = await call("workbook_get_context");
    expect(context).toMatchObject({
      workbookId: workbook.id,
      revision: workbook.revision,
      sheetName: "Sheet1",
      range: "A1:C4",
    });
    expect(context.cells).toHaveLength(12);
    const total = context.cells.find((cell) => cell.address === "B4");
    expect(total.raw).toBe("=B2+B3");
    expect(JSON.stringify(context)).not.toContain("outside the selection");
  });

  test("reads a range inside the selection and refuses one outside it", async () => {
    const { call } = await connect(ledger());
    const inside = await call("workbook_read_range", { range: "A2:B2" });
    expect(inside.cells.map((cell) => cell.raw)).toEqual(["Rent", 1200]);
    const outside = await call("workbook_read_range", { range: "D9" });
    expect(outside.error).toMatch(/outside the approved selection/);
  });

  test("inspects a formula and says whether it could be proposed", async () => {
    const { call } = await connect(ledger());
    const good = await call("workbook_inspect_formula", {
      formula: "=SUM(B2:B3)",
    });
    expect(good).toMatchObject({
      functions: ["SUM"],
      references: ["B2", "B3"],
      proposable: true,
    });
    const outside = await call("workbook_inspect_formula", {
      formula: "=D9*2",
    });
    expect(outside.proposable).toBe(false);
    expect(outside.reason).toMatch(/approved selection/);
    const volatile = await call("workbook_inspect_formula", {
      formula: "=NOW()",
    });
    expect(volatile.volatileFunctions).toEqual(["NOW"]);
    expect(volatile.proposable).toBe(false);
  });

  test("evaluates a formula without storing it", async () => {
    const workbook = ledger();
    const before = JSON.stringify(workbook);
    const { call } = await connect(workbook);
    const result = await call("workbook_evaluate_formula", {
      formula: "=B2*2",
    });
    expect(result.result).toBe(2400);
    expect(JSON.stringify(workbook)).toBe(before);
    const refused = await call("workbook_evaluate_formula", {
      formula: "=D9",
    });
    expect(refused.error).toMatch(/approved selection/);
  });

  test("describes supported functions", async () => {
    const { call } = await connect(ledger());
    const { functions } = await call("workbook_get_formula_contracts", {
      functions: ["pmt", "NOW", "NOT_A_FUNCTION"],
    });
    expect(functions).toEqual([
      {
        name: "PMT",
        supported: true,
        family: "financial",
        volatile: false,
        allowedInProposals: true,
      },
      expect.objectContaining({
        name: "NOW",
        supported: true,
        volatile: true,
        allowedInProposals: false,
      }),
      expect.objectContaining({ name: "NOT_A_FUNCTION", supported: false }),
    ]);
  });

  test("reports formulas that evaluate to an error", async () => {
    const { call } = await connect(ledger());
    const report = await call("workbook_check_formula_compatibility");
    expect(report.formulaCount).toBe(2);
    expect(report.issues).toEqual([
      {
        address: "C4",
        formula: "=1/0",
        problems: ["Evaluates to #DIV/0!."],
      },
    ]);
  });

  test("audit checks cite the cells they are about", async () => {
    const { call } = await connect(ledger());
    const { findings } = await call("workbook_run_audit_checks");
    expect(findings).toContainEqual(
      expect.objectContaining({ check: "error-value", address: "C4" }),
    );
    expect(findings).toContainEqual(
      expect.objectContaining({
        check: "number-stored-as-text",
        address: "B3",
      }),
    );
  });

  test("queues a valid proposal and leaves the workbook unchanged", async () => {
    const workbook = ledger();
    const before = JSON.stringify(workbook);
    const { call, queued } = await connect(workbook);
    const sheetId = workbook.sheets[0].id;
    const result = await call("workbook_queue_proposal", {
      proposal: proposalFor(workbook, [
        { sheetId, address: "B3", cell: { raw: 85.5 } },
      ]),
    });
    expect(result).toMatchObject({ queued: true, applied: false, id: "p-1" });
    expect(queued).toHaveLength(1);
    expect(queued[0].patches[0]).toMatchObject({
      address: "B3",
      cell: { raw: 85.5 },
    });
    expect(JSON.stringify(workbook)).toBe(before);
  });

  test("refuses proposals outside the selection, stale ones and unsafe formulas", async () => {
    const workbook = ledger();
    const { call, queued } = await connect(workbook);
    const sheetId = workbook.sheets[0].id;
    const outside = await call("workbook_queue_proposal", {
      proposal: proposalFor(workbook, [
        { sheetId, address: "D9", cell: { raw: 1 } },
      ]),
    });
    expect(outside.error).toMatch(/inside the approved selection/);
    const stale = await call("workbook_queue_proposal", {
      proposal: proposalFor(
        workbook,
        [{ sheetId, address: "B3", cell: { raw: 1 } }],
        { revision: workbook.revision + 1 },
      ),
    });
    expect(stale.error).toMatch(/stale/);
    const unsafe = await call("workbook_queue_proposal", {
      proposal: proposalFor(workbook, [
        { sheetId, address: "B3", cell: { raw: "=Sheet2!A1" } },
      ]),
    });
    expect(unsafe.error).toBeTruthy();
    expect(queued).toHaveLength(0);
  });

  test("needs the three host functions", () => {
    expect(() => createWorkbookMcpServer({ getWorkbook: () => ({}) })).toThrow(
      /getSelection/,
    );
  });
});

/** A1:B2 of the first sheet is the approved selection. Every other cell here is outside it. */
function payroll() {
  let workbook = createBlankWorkbook("Payroll");
  Object.assign(workbook.sheets[0].cells, {
    A1: { raw: "Item" },
    B1: { raw: "Cost" },
    A2: { raw: 2 },
    // Approved, and its own formula reads a cell outside the selection.
    B2: { raw: "=G8*100" },
    F7: { raw: "CEO salary" },
    G7: { raw: 987654 },
    G8: { raw: 12 },
    A40: { raw: "bank password: hunter2" },
    LOG10: { raw: "a cell named like a function" },
  });
  workbook = addWorksheet(workbook, '2" valves');
  workbook.sheets[1].cells.Z50 = { raw: "on another sheet" };
  return workbook;
}
const payrollApproved = new Set(["A1", "B1", "A2", "B2"]);
const outsideData =
  /CEO salary|987654|hunter2|named like a function|another sheet/;

/**
 * Formulas that got past an earlier check, each with what the calculation engine answers when
 * nothing stops it. The engine reads any number beside ":" as a whole row.
 */
const formulasThatReadOutside = [
  ["=INDEX(7:7,1,7)", 987654],
  ['=TEXTJOIN("|",TRUE,7:7,40:40)', "CEO salary|987654|bank password: hunter2"],
  ["=INDEX((7):(7),1,7)", 987654],
  ["=INDEX(7:SUM(7),1,7)", 987654],
  ["=INDEX(7.0:7.0,1,7)", 987654],
  ["=INDEX(A1:(7),7,7)", 987654],
  ["=INDEX(A1:B2:7,7,7)", 987654],
  ["=INDEX((A2+5):(A2+5),1,7)", 987654],
  ['=XLOOKUP("CEO salary",7:7,7:7)', "CEO salary"],
  // The lexer reads "LOG10 (" as cell LOG10, not as a call of LOG10.
  ["=CONCAT(LOG10 (10:10))", "a cell named like a function"],
  // A double quote in a sheet name must not hide the "!" as part of a string.
  [`='2" valves'!Z50&""`, "on another sheet"],
  // Names bound by LET are values, so beside ":" they are rows too.
  ["=LET(n,7,INDEX(n : n,1,7))", 987654],
  ["=LET(n,7,INDEX((n):(n),1,7))", 987654],
  ["=INDEX(SUM(7):SUM(7),1,7)", 987654],
  ["=G7", 987654],
  ["=F7&$G$7", "CEO salary987654"],
];

/** Further shapes that must be refused, whatever the engine would answer. */
const otherFormulasOutside = [
  '=XLOOKUP("x",7:7,7:7)',
  "=SUM(A:B)",
  "=SUM(1:1)",
  "=SUM(A1:B2:G7)",
  "=SUM(A1:G7)",
  "=SUM(A1:B2,7:7)",
  "=SUM(A1:B2 7:7)",
  "=LET(a,A1,b,B2,SUM(a:b))",
  "=LOG10 (A2)",
  "=SUM (7:7)",
  "=_xlfn.CONCAT(7:7)",
  "=Sheet1!G7",
  "=A1#",
  "=Table1[Cost]",
  "=Budget",
  "=TRUE1",
  "=2A2",
  "=LET(x,1,y)",
  '=INDIRECT("G7")',
  "=OFFSET(A1,6,6)",
  "='Sheet1'!G7",
  "=A2:INDEX(A1:B2,2,2)",
];

/** Ordinary formulas over the approved selection, with their results. */
const formulasInsideSelection = [
  ["=B2*2", 2400],
  ["=SUM(A2:B2)", 1202],
  ["=SUM($A$2:$B$2)*2", 2404],
  ["=SUM(A2 : B2)", 1202],
  ["=sum(b2:a2)", 1202],
  ['=IF(B2>1000,"it\'s: high","low")', "it's: high"],
  ['=TEXT(B2,"0.00")', "1200.00"],
  ['=TEXTJOIN(",",TRUE,A1:B2)', "Item,Cost,2,1200"],
  ['=ROUND(B2/7,2)&"%"', "171.43%"],
  ["=INDEX(A1:B2,2,2)", 1200],
  ["=VLOOKUP(2,A2:B2,2,FALSE)", 1200],
  ["=IF(TRUE,A2,B2)", 2],
  ["=A2:B2 B1:B2", 1200],
  ["=2E+3+B2", 3200],
  ["=SUM({1,2;3,4})+A2", 12],
  ["=LET(rate,0.5,B2*rate)", 600],
  ["=LET(x,A2,y,B2,x+y)", 1202],
  ["=LET(k,1e3,B2/k)", 1.2],
  ["=LAMBDA(v,v*2)(B2)", 2400],
  ["=LET(double,LAMBDA(v,v*2),double(B2))", 2400],
];

describe("assistant formulas stay inside the approved selection", () => {
  test.each(formulasThatReadOutside)(
    "the formula check refuses %s",
    async (formula) => {
      expect(() =>
        validateAssistantFormula(formula, payrollApproved),
      ).toThrow();
      const workbook = payroll();
      const sheetId = workbook.sheets[0].id;
      const { call, queued } = await connect(workbook, "A1:B2");
      const evaluated = await call("workbook_evaluate_formula", { formula });
      expect(evaluated.error).toBeTruthy();
      expect(JSON.stringify(evaluated)).not.toMatch(outsideData);
      const inspected = await call("workbook_inspect_formula", { formula });
      expect(inspected.proposable).toBe(false);
      const proposed = await call("workbook_queue_proposal", {
        proposal: proposalFor(
          workbook,
          [{ sheetId, address: "A2", cell: { raw: formula } }],
          { citations: [] },
        ),
      });
      expect(proposed.error).toBeTruthy();
      expect(queued).toHaveLength(0);
    },
  );

  test.each(otherFormulasOutside)("the formula check refuses %s", (formula) => {
    expect(() => validateAssistantFormula(formula, payrollApproved)).toThrow();
  });

  test("a range needs every cell of its rectangle approved, not only its two ends", () => {
    const corners = new Set(["A1", "B2"]);
    expect(() => validateAssistantFormula("=A1+B2", corners)).not.toThrow();
    expect(() => validateAssistantFormula("=SUM(A1:B2)", corners)).toThrow(
      /approved selection/,
    );
  });

  test.each(formulasInsideSelection)(
    "%s is accepted and calculated",
    async (formula, result) => {
      expect(() =>
        validateAssistantFormula(formula, payrollApproved),
      ).not.toThrow();
      const { call } = await connect(payroll(), "A1:B2");
      const evaluated = await call("workbook_evaluate_formula", { formula });
      expect(evaluated.result).toBe(result);
      const inspected = await call("workbook_inspect_formula", { formula });
      expect(inspected.proposable).toBe(true);
    },
  );

  // The second guard on its own: no formula check runs in this test.
  test.each(formulasThatReadOutside)(
    "the approved-selection calculator alone blocks %s",
    (formula, engineAnswer) => {
      const workbook = payroll();
      const sheetId = workbook.sheets[0].id;
      const unguarded = getWorkbookCalculation(workbook).calculator;
      expect(unguarded.evaluateFormula(sheetId, "A1", formula)).toBe(
        engineAnswer,
      );
      const calculator = createApprovedSelectionCalculator(
        workbook,
        sheetId,
        payrollApproved,
      );
      expect(calculator.evaluateFormula(sheetId, "A1", formula)).toBe(
        "#BLOCKED!",
      );
      expect(calculator.readWasRefused()).toBe(true);
    },
  );

  test("the approved-selection calculator reports a refused read that the formula hides", () => {
    const workbook = payroll();
    const sheetId = workbook.sheets[0].id;
    const calculator = createApprovedSelectionCalculator(
      workbook,
      sheetId,
      payrollApproved,
    );
    expect(
      calculator.evaluateFormula(sheetId, "A1", '=IFERROR(G7,"hidden")'),
    ).toBe("hidden");
    expect(calculator.readWasRefused()).toBe(true);
  });

  test("the approved-selection calculator gives approved cells their calculated values", () => {
    const workbook = payroll();
    const sheetId = workbook.sheets[0].id;
    const calculator = createApprovedSelectionCalculator(
      workbook,
      sheetId,
      payrollApproved,
    );
    // B2 is =G8*100. Its value is approved; G8 itself is not.
    expect(calculator.evaluateFormula(sheetId, "A1", "=B2+A2")).toBe(1202);
    expect(calculator.evaluateFormula(sheetId, "A1", "=FORMULATEXT(B2)")).toBe(
      "=G8*100",
    );
    expect(calculator.readWasRefused()).toBe(false);
    expect(calculator.evaluateFormula(sheetId, "A1", "=G8")).toBe("#BLOCKED!");
    expect(calculator.readWasRefused()).toBe(true);
  });

  test("a defined name called TRUE does not change a formula the check accepts: TRUE stays the constant", async () => {
    // The formula check cannot see the workbook's names, so it has to accept TRUE and FALSE.
    // The engine reads an unqualified TRUE as the constant, never as a name, so nothing outside
    // the selection is read.
    const workbook = payroll();
    workbook.definedNames = [{ name: "TRUE", refersTo: "=Sheet1!$G$7" }];
    const sheetId = workbook.sheets[0].id;
    const formula = "=TRUE+0";
    expect(() =>
      validateAssistantFormula(formula, payrollApproved),
    ).not.toThrow();
    expect(
      getWorkbookCalculation(workbook).calculator.evaluateFormula(
        sheetId,
        "A1",
        formula,
      ),
    ).toBe(1);

    const { call } = await connect(workbook, "A1:B2");
    const evaluated = await call("workbook_evaluate_formula", { formula });
    expect(evaluated.result).toBe(1);
    expect(JSON.stringify(evaluated)).not.toMatch(outsideData);

    // The editor's own check of a proposal accepts it for the same reason.
    const selection = { sheetId, range: "A1:B2" };
    const proposal = proposalFor(
      workbook,
      [{ sheetId, address: "A2", cell: { raw: formula } }],
      { citations: [] },
    );
    expect(() =>
      validateAssistantProposal(
        proposal,
        buildAssistantSelectionContext(workbook, selection),
      ),
    ).not.toThrow();
    expect(
      validateProposedCalculation(workbook, selection, proposal).patches,
    ).toHaveLength(1);
  });

  test("the editor's check of a proposal accepts an ordinary formula and refuses one that reads outside", () => {
    const workbook = payroll();
    const sheetId = workbook.sheets[0].id;
    const selection = { sheetId, range: "A1:B2" };
    const propose = (raw) =>
      validateProposedCalculation(
        workbook,
        selection,
        proposalFor(workbook, [{ sheetId, address: "A2", cell: { raw } }], {
          citations: [],
        }),
      );
    expect(propose("=B2*2").patches[0].cell.raw).toBe("=B2*2");
    expect(propose("=LET(rate,0.5,B2*rate)").patches).toHaveLength(1);
    for (const [formula] of formulasThatReadOutside)
      expect(() => propose(formula)).toThrow();
  });
});

describe("MCP stdio command", () => {
  test("serves a workbook file and appends proposals to the outbox", async () => {
    const workbook = ledger();
    const folder = await mkdtemp(join(tmpdir(), "ledgerglass-mcp-"));
    const workbookPath = join(folder, "workbook.json");
    const outboxPath = join(folder, "proposals.json");
    await writeFile(workbookPath, JSON.stringify(workbook));
    const client = new Client({ name: "unit-stdio", version: "1" });
    await client.connect(
      new StdioClientTransport({
        command: process.execPath,
        args: [
          "scripts/mcp-stdio.mjs",
          "--workbook",
          workbookPath,
          "--sheet",
          workbook.sheets[0].id,
          "--range",
          "A1:C4",
          "--outbox",
          outboxPath,
        ],
      }),
    );
    open.push(() => client.close());
    const context = (
      await client.callTool({ name: "workbook_get_context", arguments: {} })
    ).structuredContent;
    expect(context.cells).toHaveLength(12);
    const queued = await client.callTool({
      name: "workbook_queue_proposal",
      arguments: {
        proposal: proposalFor(workbook, [
          {
            sheetId: workbook.sheets[0].id,
            address: "B3",
            cell: { raw: 85.5 },
          },
        ]),
      },
    });
    expect(queued.structuredContent.queued).toBe(true);
    const outbox = JSON.parse(await readFile(outboxPath, "utf8"));
    expect(outbox).toHaveLength(1);
    expect(outbox[0].proposal.id).toBe("p-1");
    expect(JSON.parse(await readFile(workbookPath, "utf8"))).toEqual(
      JSON.parse(JSON.stringify(workbook)),
    );
  });
});
