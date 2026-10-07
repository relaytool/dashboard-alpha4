(() => {
  "use strict";

  const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
  const CHANGE_LOG_SHEET_NAME = "Inventory Change Log";
  const CHANGE_LOG_HEADERS = [
    "Timestamp",
    "Event ID",
    "Asset",
    "Inventory Row",
    "Delta",
    "Source",
    "Movement",
    "Quantity",
    "Client",
    "User",
    "Transaction Event ID"
  ];

  const api = {
    async get(spreadsheetId, range, accessToken) {
      return request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`, accessToken);
    },
    async batchUpdate(spreadsheetId, body, accessToken) {
      return request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, accessToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body)
      });
    }
  };

  async function request(url, accessToken, options = {}) {
    const response = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        ...(options.headers || {})
      }
    });
    const raw = await response.text();
    let data = {};
    try { data = raw ? JSON.parse(raw) : {}; } catch (_) {}
    if (!response.ok) {
      throw new Error(data?.error?.message || `Google Sheets request failed (${response.status})`);
    }
    return data;
  }

  function quoteSheetName(value) {
    return `'${String(value).replace(/'/g, "''")}'`;
  }

  function columnLetter(value) {
    let n = Number(value) || 0;
    let result = "";
    while (n > 0) {
      const remainder = (n - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      n = Math.floor((n - 1) / 26);
    }
    return result || "A";
  }

  function randomEventId(prefix = "EVT") {
    const time = Date.now().toString(36).toUpperCase();
    const random = Math.random().toString(36).slice(2, 8).toUpperCase();
    return `${prefix}-${time}-${random}`;
  }

  function cellValue(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return { userEnteredValue: { numberValue: value } };
    }
    return { userEnteredValue: { stringValue: String(value ?? "") } };
  }

  function appendCellsRequest(sheetId, rows) {
    return {
      appendCells: {
        sheetId,
        fields: "userEnteredValue",
        rows: rows.map(row => ({ values: row.map(cellValue) }))
      }
    };
  }

  function balanceFormula(assetColumn = 1) {
    const assetCol = columnLetter(assetColumn);
    return `=SUMIF(${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!$C:$C,INDIRECT("${assetCol}"&ROW()),${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!$E:$E)`;
  }

  function formulaCell(formula) {
    return { userEnteredValue: { formulaValue: formula } };
  }

  function updateBalanceCellRequest(sheetId, item) {
    return {
      updateCells: {
        start: {
          sheetId,
          rowIndex: Math.max(0, Number(item.rowNumber) - 1),
          columnIndex: Math.max(0, Number(item.balanceColumn) - 1)
        },
        rows: [{ values: [formulaCell(balanceFormula(Number(item.assetColumn) || 1))] }],
        fields: "userEnteredValue"
      }
    };
  }

  function buildBaselineRow(item, user, timestamp, eventId) {
    return [
      timestamp,
      eventId,
      String(item.asset || "").trim(),
      Number(item.rowNumber) || 0,
      Number(item.balance) || 0,
      "BASELINE",
      "BASELINE",
      Number(item.balance) || 0,
      "",
      user || "System",
      ""
    ];
  }

  function buildChangeRow(item, movement, quantity, client, user, timestamp, transactionEventId) {
    const delta = movement === "RECEIVED" ? Math.abs(Number(quantity) || 0) : -Math.abs(Number(quantity) || 0);
    return [
      timestamp,
      randomEventId("BAL"),
      String(item.asset || "").trim(),
      Number(item.rowNumber) || 0,
      delta,
      "MOVEMENT",
      String(movement || "").trim(),
      Math.abs(Number(quantity) || 0),
      String(client || "").trim(),
      user || "Google user",
      transactionEventId || ""
    ];
  }

  function metadataSheetInfo(metadata, title) {
    return (metadata?.sheets || []).find(sheet => sheet?.properties?.title === title)?.properties || null;
  }

  async function ensureChangeLog({ spreadsheetId, accessToken, inventoryItems = [], user = "System" }) {
    if (!spreadsheetId || !accessToken) throw new Error("Google Sheets connection is not available.");

    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const inventorySheet = metadataSheetInfo(book, CONFIG.INVENTORY_SHEET_NAME);
    if (!inventorySheet?.sheetId) throw new Error(`Sheet "${CONFIG.INVENTORY_SHEET_NAME}" was not found.`);

    let logSheet = metadataSheetInfo(book, CHANGE_LOG_SHEET_NAME);
    if (logSheet?.sheetId) {
      const headerResponse = await api.get(spreadsheetId, `${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!A1:K1`, accessToken);
      const header = headerResponse?.values?.[0] || [];
      const valid = CHANGE_LOG_HEADERS.every((name, index) => String(header[index] ?? "").trim() === name);
      if (!valid) {
        throw new Error(`The "${CHANGE_LOG_SHEET_NAME}" sheet exists but has an unexpected header. Keep its first row as: ${CHANGE_LOG_HEADERS.join(", ")}`);
      }
      return { initialized: true, sheetId: logSheet.sheetId };
    }

    const items = Array.isArray(inventoryItems) ? inventoryItems : [];

    const existingIds = (book.sheets || []).map(sheet => Number(sheet?.properties?.sheetId)).filter(Number.isFinite);
    let newSheetId = Math.max(1000, ...existingIds) + 1;
    if (newSheetId > 2147483000) newSheetId = 1000000000;

    const timestamp = new Date().toISOString();
    const baselineRows = items.map(item => buildBaselineRow(item, user, timestamp, randomEventId("BASE")));
    const requests = [
      {
        addSheet: {
          properties: {
            title: CHANGE_LOG_SHEET_NAME,
            sheetId: newSheetId
          }
        }
      },
      appendCellsRequest(newSheetId, [CHANGE_LOG_HEADERS, ...baselineRows]),
      ...items.map(item => updateBalanceCellRequest(inventorySheet.sheetId, item))
    ];

    try {
      await api.batchUpdate(spreadsheetId, { requests }, accessToken);
    } catch (error) {
      const message = String(error?.message || "");
      if (/already exists|already used|duplicate/i.test(message)) {
        const latest = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
        const latestLog = metadataSheetInfo(latest, CHANGE_LOG_SHEET_NAME);
        if (latestLog?.sheetId) return { initialized: true, sheetId: latestLog.sheetId, concurrent: true };
      }
      throw error;
    }
    return { initialized: false, sheetId: newSheetId, created: true };
  }

  async function appendMovementBatch({
    spreadsheetId,
    accessToken,
    transactionSheetId,
    transactionRows = [],
    changes = [],
    user = "Google user",
    timestamp = new Date().toISOString()
  }) {
    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const logSheet = metadataSheetInfo(book, CHANGE_LOG_SHEET_NAME);
    const resolvedTransactionSheetId = transactionSheetId || metadataSheetInfo(book, CONFIG.TRANSACTIONS_SHEET_NAME)?.sheetId;
    if (!resolvedTransactionSheetId || !logSheet?.sheetId) {
      throw new Error(`The ${CHANGE_LOG_SHEET_NAME} system sheet is not ready. Refresh the page and try again.`);
    }

    const transactionEventId = randomEventId("TXN");
    const transactionHeaders = transactionRows.map(row => {
      const copy = Array.isArray(row) ? row.slice() : [];
      return copy;
    });
    const changeRows = changes.map(change => buildChangeRow(
      change.item,
      change.movement,
      change.quantity,
      change.client,
      user,
      timestamp,
      transactionEventId
    ));

    const requests = [
      appendCellsRequest(resolvedTransactionSheetId, transactionHeaders),
      appendCellsRequest(logSheet.sheetId, changeRows)
    ];

    await api.batchUpdate(spreadsheetId, { requests }, accessToken);
    return { transactionEventId, changeCount: changeRows.length };
  }

  async function appendAsset({ spreadsheetId, accessToken, inventorySheetId, assetName, startingBalance, user = "Google user" }) {
    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const resolvedInventorySheetId = inventorySheetId || metadataSheetInfo(book, CONFIG.INVENTORY_SHEET_NAME)?.sheetId;
    const logSheet = metadataSheetInfo(book, CHANGE_LOG_SHEET_NAME);
    if (!resolvedInventorySheetId || !logSheet?.sheetId) throw new Error(`The ${CHANGE_LOG_SHEET_NAME} system sheet is not ready. Refresh the page and try again.`);

    const timestamp = new Date().toISOString();
    const eventId = randomEventId("BASE");
    const changeRow = [
      timestamp,
      eventId,
      String(assetName || "").trim(),
      "",
      Number(startingBalance) || 0,
      "BASELINE",
      "BASELINE",
      Number(startingBalance) || 0,
      "",
      user,
      ""
    ];

    await api.batchUpdate(spreadsheetId, {
      requests: [
        {
          appendCells: {
            sheetId: resolvedInventorySheetId,
            fields: "userEnteredValue",
            rows: [{ values: [cellValue(assetName), formulaCell(balanceFormula(1))] }]
          }
        },
        appendCellsRequest(logSheet.sheetId, [changeRow])
      ]
    }, accessToken);

    return { eventId };
  }

  async function migrateRename({ spreadsheetId, accessToken, oldName, newName, inventorySheetId, rowNumber, assetColumn = 1 }) {
    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const inventorySheet = metadataSheetInfo(book, CONFIG.INVENTORY_SHEET_NAME);
    const logSheet = metadataSheetInfo(book, CHANGE_LOG_SHEET_NAME);
    const resolvedInventorySheetId = inventorySheetId || inventorySheet?.sheetId;
    if (!resolvedInventorySheetId) throw new Error(`Sheet "${CONFIG.INVENTORY_SHEET_NAME}" was not found.`);

    const requests = [{
      updateCells: {
        start: {
          sheetId: resolvedInventorySheetId,
          rowIndex: Math.max(0, Number(rowNumber) - 1),
          columnIndex: Math.max(0, Number(assetColumn) - 1)
        },
        rows: [{ values: [cellValue(newName)] }],
        fields: "userEnteredValue"
      }
    }];

    if (logSheet?.sheetId) {
      requests.push({
        findReplace: {
          find: String(oldName || "").trim(),
          replacement: String(newName || "").trim(),
          matchCase: false,
          matchEntireCell: true,
          searchByRegex: false,
          includeFormulas: false,
          range: {
            sheetId: logSheet.sheetId,
            startColumnIndex: 2,
            endColumnIndex: 3
          }
        }
      });
    }

    await api.batchUpdate(spreadsheetId, { requests }, accessToken);
    return true;
  }

  window.FM_INVENTORY_LEDGER = Object.freeze({
    CHANGE_LOG_SHEET_NAME,
    CHANGE_LOG_HEADERS,
    ensureChangeLog,
    appendMovementBatch,
    appendAsset,
    migrateRename,
    balanceFormula
  });
})();
