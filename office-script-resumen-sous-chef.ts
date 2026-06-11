// Office Script — Resumen Sous Chef
// Pegar en Excel Online: abrir "Evolucion ventas Facturado.xlsx" → pestaña Automatizar → Nuevo script.
// Filtra la hoja "Base" por UDN = "Sous Chef", agrega por mes y cliente, lee TC MEP de "Indices"
// y devuelve un JSON chico que el dashboard consume directo (ver DATA_JSON_URL en index.html).

function main(workbook: ExcelScript.Workbook): string {
  // --- Hoja Base: ventas Sous Chef ---
  const base = workbook.getWorksheet('Base');
  if (!base) throw new Error("No se encontró la hoja 'Base'");
  const used = base.getUsedRange();
  if (!used) throw new Error("La hoja 'Base' está vacía");

  const firstRow = used.getRowIndex();
  const firstCol = used.getColumnIndex();
  const dataRows = used.getRowCount() - 1;

  const headers = (used.getRow(0).getValues()[0] as (string | number | boolean)[])
    .map(v => String(v).trim().toLowerCase());
  const col = (name: string): number => headers.indexOf(name);

  const cUdn = col('udn');
  const cPeriodo = col('periodo fiscal');
  const cFantasia = col('fantasia');
  const cCliente = col('cliente');
  const cUsd = col('neto u$s');
  const cArs = col('monto neto');
  const cEmision = col('emision');
  if (cUdn < 0 || cPeriodo < 0) throw new Error("Faltan las columnas 'UDN' o 'Periodo fiscal' en la hoja Base");

  // Lee solo las columnas necesarias (una llamada por columna, mucho más liviano que todo el rango)
  const readCol = (idx: number): (string | number | boolean)[] => {
    if (idx < 0 || dataRows < 1) return [];
    return base.getRangeByIndexes(firstRow + 1, firstCol + idx, dataRows, 1)
      .getValues().map(r => r[0]);
  };

  const vUdn = readCol(cUdn);
  const vPeriodo = readCol(cPeriodo);
  const vFantasia = readCol(cFantasia);
  const vCliente = readCol(cCliente);
  const vUsd = readCol(cUsd);
  const vArs = readCol(cArs);
  const vEmision = readCol(cEmision);

  const scMonthly: { [k: string]: { montoNeto: number; netoUsd: number } } = {};
  const clientsByMonthSet: { [k: string]: { [c: string]: boolean } } = {};
  const clientFirstMonth: { [c: string]: string } = {};
  const clientMonthData: { [k: string]: { [c: string]: { netoUsd: number; montoNeto: number; lastPurchase: string | null } } } = {};
  const clientLastPurchase: { [c: string]: string } = {};

  for (let i = 0; i < dataRows; i++) {
    if (String(vUdn[i]).trim() !== 'Sous Chef') continue;
    const periodo = toDate(vPeriodo[i]);
    if (!periodo) continue;
    const mKey = periodo.toISOString().slice(0, 7);
    const client = String(vFantasia[i] ?? '').trim() || String(vCliente[i] ?? '').trim() || 'Sin nombre';
    const usd = Number(vUsd[i]) || 0;
    const ars = Number(vArs[i]) || 0;
    const emision = toDate(vEmision[i]);

    if (!scMonthly[mKey]) scMonthly[mKey] = { montoNeto: 0, netoUsd: 0 };
    scMonthly[mKey].montoNeto += ars;
    scMonthly[mKey].netoUsd += usd;

    if (!clientsByMonthSet[mKey]) clientsByMonthSet[mKey] = {};
    clientsByMonthSet[mKey][client] = true;

    if (!clientFirstMonth[client] || mKey < clientFirstMonth[client]) clientFirstMonth[client] = mKey;

    if (!clientMonthData[mKey]) clientMonthData[mKey] = {};
    if (!clientMonthData[mKey][client]) clientMonthData[mKey][client] = { netoUsd: 0, montoNeto: 0, lastPurchase: null };
    const cd = clientMonthData[mKey][client];
    cd.netoUsd += usd;
    cd.montoNeto += ars;
    if (emision) {
      const iso = emision.toISOString();
      if (!cd.lastPurchase || iso > cd.lastPurchase) cd.lastPurchase = iso;
      if (!clientLastPurchase[client] || iso > clientLastPurchase[client]) clientLastPurchase[client] = iso;
    }
  }

  const clientsByMonth: { [k: string]: string[] } = {};
  for (const k in clientsByMonthSet) clientsByMonth[k] = Object.keys(clientsByMonthSet[k]);

  // --- Hoja Indices: TC MEP (col B = periodo, col D = valor) ---
  const mepData: { [k: string]: number } = {};
  const idxSheet = workbook.getWorksheet('Indices');
  if (idxSheet) {
    const u = idxSheet.getUsedRange();
    if (u) {
      const totalRows = u.getRowIndex() + u.getRowCount();
      const vals = idxSheet.getRangeByIndexes(0, 0, totalRows, 4).getValues();
      let hdr = -1;
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][1]).toLowerCase().indexOf('periodo') >= 0) { hdr = i; break; }
      }
      if (hdr >= 0) {
        for (let i = hdr + 1; i < vals.length; i++) {
          const d = toDate(vals[i][1]);
          const val = Number(vals[i][3]);
          if (!d || !val) continue;
          mepData[d.toISOString().slice(0, 7)] = Math.round(val);
        }
      }
    }
  }

  return JSON.stringify({
    generatedAt: new Date().toISOString(),
    scMonthly,
    mepData,
    clientsByMonth,
    clientFirstMonth,
    clientMonthData,
    clientLastPurchase,
  });
}

// Convierte serial de Excel o texto (DD/MM/YYYY, ISO) a Date en UTC; null si no es fecha válida
function toDate(v: string | number | boolean): Date | null {
  if (typeof v === 'number') {
    // Serial de Excel: días desde 1900. 25569 = 1970-01-01. Umbral descarta números que no son fechas.
    if (v < 20000) return null;
    return new Date(Math.round((v - 25569) * 86400000));
  }
  const s = String(v).trim();
  if (!s) return null;
  const dmy = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (dmy) {
    const d = new Date(Date.UTC(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1])));
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
