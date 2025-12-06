function findSkuGlobal_(sku) {
  const ss = SpreadsheetApp.getActive();
  const tabs = ["Clubs","Sets","Tassen","Trolley's","Overig","Diensten"];

  for (let t of tabs) {
    const sh = ss.getSheetByName(t);
    if (!sh) continue;

    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === sku) {
        return { sheet: t, row: i + 2 };
      }
    }
  }
  return null;
}

function norm(s) {
  return String(s || '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')  // zero width weg
    .replace(/\s/g, '')                     // spaties weg
    .toUpperCase();
}

function jsonSafe(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function apiReturnGetSafe(receiptNo) {
  Logger.log("SAFE START raw:", receiptNo);

  const no = String(receiptNo || "").trim();
  Logger.log("SAFE cleaned:", no);

  const data = apiGetReceiptForPrint(no);
  Logger.log("SAFE after apiGetReceiptForPrint:", JSON.stringify(data));

  if (!data || !data.head) {
    throw new Error("Bon niet gevonden: " + no);
  }

  const result = {
    head: {
      receipt_no: data.head.receipt_no || "",
      date:       String(data.head.date || ""),
      pay:        data.head.pay || "",
      total:      Number(data.head.total || 0),
      email:      data.head.email || "",
      pdfUrl:     String(data.head.pdfUrl || ""),
      mail:       data.head.mail || "",
      bon80Url:   String(data.head.bon80Url || "")
    },
    lines: (data.items || []).map(it => ({
      sku:      String(it.sku || ""),
      desc:     String(it.desc || ""),
      qty:      Number(it.qty || 0),
      price:    Number(it.price || 0),
      subtotal: Number(it.subtotal || 0)
    }))
  };

  Logger.log("SAFE FINAL RESULT:", JSON.stringify(result));
  return jsonSafe(result);
}

function debugReturnGet() {
  const testNo = "GL-20251118-009";  // <== vul zelf tijdelijk een bestaand bonnummer in
  const res = apiReturnGetSafe(testNo);
  Logger.log("DEBUG RESULT:\n" + JSON.stringify(res, null, 2));
  return res;
}

function apiProcessReturn(receiptNo, payload) {
  Logger.log('[RET] apiProcessReturn START');
  Logger.log('[RET] receiptNo = ' + receiptNo);
  Logger.log('[RET] payload = ' + JSON.stringify(payload));

  if (!payload || !payload.items) {
    throw new Error("Geen retourdata ontvangen (payload.items ontbreekt)");
  }

  // payload kan een object of array zijn → altijd array maken
  let items = payload.items;
  if (!Array.isArray(items)) {
    items = Object.values(items);
  }

  // Alleen items met "selected: true"
  const itemsToReturn = items.filter(it => it.selected === true);

  if (!itemsToReturn.length) {
    throw new Error("Geen artikelen geselecteerd voor retour.");
  }

  const ss = SpreadsheetApp.getActive();
  const retourSheet = ss.getSheetByName("Retouren");
  if (!retourSheet) throw new Error("Tab 'Retouren' ontbreekt.");

  const now = new Date();
  const returnNo = generateReturnNumber_();

  // Voor elk geselecteerd item
  itemsToReturn.forEach(it => {
    const sku = String(it.sku || "").trim();
    const reason = String(it.reason || "");
    const price = Number(it.price || 0);
    const desc = String(it.desc || "");

    if (!sku) return;

    // 1️⃣ Voorraad herstellen
    _revertSoldItem_(sku);

    // 2️⃣ Loggen in retour-sheet
    retourSheet.appendRow([
      receiptNo,     // originele bon
      returnNo,      // retournummer
      now,           // datum/tijd
      sku,           // artikel
      price,         // bedrag terug
      reason,        // reden
      desc           // omschrijving
    ]);
  });

  Logger.log("[RET] Retour succesvol: " + returnNo);

  return {
    ok: true,
    returnNo: returnNo,
    itemsCount: itemsToReturn.length,
    receiptNo: receiptNo
  };
}


/**
 * Herstelt een verkochte SKU in de voorraad.
 * - zet COL.F (verkoopprijs) = backup expected prijs (COL.L)
 * - zet COL.G (verkoopdatum) = ""
 * - zet COL.J (marge) = 0
 */
function _revertSoldItem_(sku) {
  console.log("[RET] _revertSoldItem_ CALLED →", sku);

  const ss = SpreadsheetApp.getActive();
  const sheets = ['Clubs','Sets','Tassen',"Trolley's",'Overig','Diensten'];

  sku = String(sku).trim();
  if (!sku) return;

  for (const name of sheets) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const vals = sh.getDataRange().getValues();

    for (let r = 1; r < vals.length; r++) {
      if (String(vals[r][0]).trim() === sku) {

        const row = r + 1;
        console.log("[RET] SKU FOUND in sheet:", name, "row:", row);

        const expectedBackup = vals[r][11];

        sh.getRange(row, 6).setValue(expectedBackup);
        sh.getRange(row, 7).setValue("");
        sh.getRange(row, 10).setValue(0);

        return;
      }
    }
  }

  throw new Error("SKU niet gevonden in voorraad: " + sku);
}
