/* === Centro Técnico GZ – Historial + Estado con Cedis y prefijos ===
   Libro: "Autotanque Centro Técnico GZ"
   - Respuestas: "Respuestas de formulario 1"
   - Historial:  "Historial_Sellos"
   - Estado:     "Estado_Actual_Laja" | "Estado_Actual_Zap" | "Estado_Actual_Col"

   Historial_Sellos (encabezados):
   PdV | Cedis | Ubicación | Sello Nuevo | Sello Removido | Fecha Instalación | Fecha Remoción | Estado | Motivo de revisión

   Estado_Actual_* (encabezados):
   PdV | Caja de Control | Válvula Solenoide | Gaspar
*/

// === CONFIG ===
const SHEET_ID = '1AgIVVj6OzdAh3QH8w2EIVxSIl175CIfOuHQWQyEddGU';
const SH_RESPUESTAS = 'Respuestas de formulario 1';
const SH_HISTORIAL  = 'Historial_Sellos';

// Dónde vive cada Cedis y qué prefijo aplica a sus sellos
const SH_ESTADOS = [
  { name: 'Estado_Actual_Laja', cedis: 'La Laja',  prefix: 'L-' },
  { name: 'Estado_Actual_Zap',  cedis: 'Zapopan',  prefix: 'Z-' },
  { name: 'Estado_Actual_Col',  cedis: 'Colotlán', prefix: 'C-' },
];

// Nombres EXACTOS tal como aparecen en la hoja de respuestas
const FORM = {
  CEDIS: 'Selecciona un Cedis',
  PDV:   {
    L: 'L. Selecciona un punto de venta',
    Z: 'Z. Selecciona un punto de venta',
    C: 'C. Selecciona un punto de venta'
  },
  FECHA : 'Fecha de la visita',
  MOTIVO: 'Motivo de revisión',
  UBIC_FORM: ['Caja Control','Válvula solenoide','Gaspar'], // tal cual en respuestas
};

// Normaliza nombre del form -> nombre de columna en Estado_Actual
const UBIC_NORMAL = {
  'Caja Control'      : 'Caja de Control',
  'Válvula solenoide' : 'Válvula Solenoide',
  'Gaspar'            : 'Gaspar'
};

/* ===================== MENÚ DIAGNÓSTICO (opcional) ===================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Diagnóstico')
    .addItem('Listar hojas', '__diag_listSheets__')
    .addToUi();
}
function __diag_listSheets__() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = ss.getSheets().map(s => s.getName());
  SpreadsheetApp.getUi().alert('Hojas encontradas:\n\n' + names.join('\n'));
}

/* ===================== DISPARADOR PRINCIPAL ===================== */
/** Se ejecuta al enviar el Form (trigger: From spreadsheet → On form submit) */
function onFormSubmit(e){
  // 1) Obtener el Spreadsheet de forma robusta
  const ss = (e && e.source) ? e.source : SpreadsheetApp.openById(SHEET_ID);

  // 2) Determinar la hoja de respuestas y la fila a procesar
  let shResp, rowIndex;
  if (e && e.range && e.range.getSheet && e.range.getSheet().getName() === SH_RESPUESTAS) {
    shResp = e.range.getSheet();
    rowIndex = e.range.getRow();
  } else {
    shResp = mustGetSheet(ss, SH_RESPUESTAS);
    rowIndex = shResp.getLastRow(); // fallback: última fila
  }

  // 3) Leer encabezados y la fila recién agregada
  const headers = shResp.getRange(1, 1, 1, shResp.getLastColumn()).getValues()[0].map(v => String(v||'').trim());
  const rowVals = shResp.getRange(rowIndex, 1, 1, shResp.getLastColumn()).getValues()[0];

  // 4) Indexar columnas reales de esta fila
  const idx = indexarRespuestas(headers);

  // 5) Determinar Cedis seleccionado y metadatos
  const cedisSel = getSafe(rowVals, idx.cedis).toString().trim(); // 'La Laja' | 'Zapopan' | 'Colotlán'
  if (!cedisSel) return;

  const metaEstado = SH_ESTADOS.find(m => m.cedis === cedisSel);
  if (!metaEstado) {
    throw new Error(`Cedis "${cedisSel}" no está mapeado en SH_ESTADOS`);
  }

  // 6) Seleccionar sección (L/Z/C) según Cedis
  const sec = (cedisSel === 'La Laja') ? 'L' : (cedisSel === 'Zapopan' ? 'Z' : 'C');

  // 7) Leer datos base
  const pdv    = getSafe(rowVals, idx.pdv[sec]).toString().trim();
  const fecha  = parseFecha(getSafe(rowVals, idx.fecha[sec]));
  const motivo = getSafe(rowVals, idx.motivo[sec]).toString().trim();
  if (!pdv || !fecha) return;

  // 8) Abrir hojas destino con chequeo estricto
  const shHist  = mustGetSheet(ss, SH_HISTORIAL);
  const idxHist = headerMap(shHist);

  const shEst   = mustGetSheet(ss, metaEstado.name);
  const idxEst  = headerMap(shEst);
  const rowEst  = findOrCreatePdVRow(shEst, idxEst, pdv);

  // 9) Procesar únicamente las 3 ubicaciones definidas
  FORM.UBIC_FORM.forEach(uf => {
    const valor = getSafe(rowVals, idx.ubic[uf]).toString().trim();
    if (!valor) return; // esa ubicación no se tocó

    const ubicNorm   = UBIC_NORMAL[uf];                 // nombre exacto de columna en Estado_Actual
    const colUbicEst = idxEst[ubicNorm] || 0;
    const selloPrev  = colUbicEst ? shEst.getRange(rowEst, colUbicEst).getValue() : '';
    const selloNuevo = `${metaEstado.prefix}${valor}`;

    // 1) Registrar INSTALADO en Historial_Sellos (con Cedis)
    shHist.appendRow(buildHistRow(idxHist, {
      PdV: pdv, Cedis: cedisSel, Ubicacion: ubicNorm,
      SelloNuevo: selloNuevo, FechaInstalacion: fecha,
      Estado: 'Instalado', Motivo: motivo
    }));

    // 2) Registrar REMOVIDO (si había previo)
    if (selloPrev) {
      shHist.appendRow(buildHistRow(idxHist, {
        PdV: pdv, Cedis: cedisSel, Ubicacion: ubicNorm,
        SelloRemovido: selloPrev, FechaRemocion: fecha,
        Estado: 'Removido', Motivo: motivo
      }));
    }

    // 3) Actualizar Estado_Actual_* con el nuevo sello
    if (colUbicEst) shEst.getRange(rowEst, colUbicEst).setValue(selloNuevo);
  });
}

/* ===================== HELPERS ===================== */

// Lanza error claro si no existe la hoja solicitada
function mustGetSheet(ss, name){
  const sh = ss.getSheetByName(name);
  if (!sh) {
    const names = ss.getSheets().map(s => s.getName()).join(' | ');
    throw new Error(`No se encontró la hoja "${name}". Hojas disponibles: ${names}`);
  }
  return sh;
}

// Devuelve valor seguro (por índice 1-based). Si i=0/undefined, retorna ''.
function getSafe(arr, i){ return (i && i>0) ? (arr[i-1] ?? '') : ''; }

// Mapea encabezado -> índice 1-based
function headerMap(sheet){
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  return headers.reduce((m,h,i)=>{ m[String(h).trim()] = i+1; return m; },{});
}

// Busca o crea la fila para un PdV en la hoja de Estado_Actual
function findOrCreatePdVRow(sheet, idx, pdv){
  const colPdV = idx['PdV'] || 1;
  const nRows  = Math.max(sheet.getLastRow()-1, 0);
  if (nRows > 0){
    const data = sheet.getRange(2, colPdV, nRows, 1).getValues();
    for (let i=0;i<data.length;i++){
      if (String(data[i][0]).trim().toUpperCase() === pdv.toUpperCase()) return i+2;
    }
  }
  const rowArr = Array(sheet.getLastColumn()).fill('');
  rowArr[colPdV-1] = pdv;
  sheet.appendRow(rowArr);
  return sheet.getLastRow();
}

// Construye una fila para Historial_Sellos por nombre de columna
function buildHistRow(idx, obj){
  const fila = Array(Object.keys(idx).length).fill('');
  if (obj.PdV)              fila[idx['PdV']               -1] = obj.PdV;
  if (obj.Cedis)            fila[idx['Cedis']             -1] = obj.Cedis;
  if (obj.Ubicacion)        fila[idx['Ubicación']         -1] = obj.Ubicacion;
  if (obj.SelloNuevo)       fila[idx['Sello Nuevo']       -1] = obj.SelloNuevo;
  if (obj.SelloRemovido)    fila[idx['Sello Removido']    -1] = obj.SelloRemovido;
  if (obj.FechaInstalacion) fila[idx['Fecha Instalación'] -1] = obj.FechaInstalacion;
  if (obj.FechaRemocion)    fila[idx['Fecha Remoción']    -1] = obj.FechaRemocion;
  if (obj.Estado)           fila[idx['Estado']            -1] = obj.Estado;
  if (obj.Motivo)           fila[idx['Motivo de revisión']-1] = obj.Motivo;
  return fila;
}

// Convierte texto/Date a Date válido
function parseFecha(v){
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d) ? new Date() : d;
}

/**
 * Escanea la fila de encabezados para asociar columnas por secciones L/Z/C.
 * Regresa índices 1-based para:
 *   - cedis
 *   - pdv: {L,Z,C}
 *   - fecha: {L,Z,C}
 *   - motivo: {L,Z,C}
 *   - ubic: { 'Caja Control': i, 'Válvula solenoide': i, 'Gaspar': i }
 */
function indexarRespuestas(headers){
  const h = headers;
  const find = (text, start=0, end=h.length) => {
    for (let i=start;i<end;i++){ if (h[i] === text) return i+1; }
    return 0;
  };

  const iCEDIS = find(FORM.CEDIS);
  const iPDV_L = find(FORM.PDV.L);
  const iPDV_Z = find(FORM.PDV.Z);
  const iPDV_C = find(FORM.PDV.C);

  // Límites (desde PDV de una sección hasta el siguiente PDV)
  const bounds = [iPDV_L, iPDV_Z, iPDV_C, h.length+1].map(i => i||h.length+1);
  const nextWithin = (label, from, to) => find(label, Math.max(from,1)-1, Math.max(to-1,0));

  const iFecha  = {
    L: nextWithin(FORM.FECHA,  bounds[0], bounds[1]),
    Z: nextWithin(FORM.FECHA,  bounds[1], bounds[2]),
    C: nextWithin(FORM.FECHA,  bounds[2], bounds[3]),
  };
  const iMotivo = {
    L: nextWithin(FORM.MOTIVO, bounds[0], bounds[1]),
    Z: nextWithin(FORM.MOTIVO, bounds[1], bounds[2]),
    C: nextWithin(FORM.MOTIVO, bounds[2], bounds[3]),
  };

  // Las 3 ubicaciones (index global por nombre tal cual en respuestas)
  const iUbic = {};
  FORM.UBIC_FORM.forEach(u => { iUbic[u] = find(u); });

  return {
    cedis: iCEDIS,
    pdv:   { L: iPDV_L, Z: iPDV_Z, C: iPDV_C },
    fecha: iFecha,
    motivo:iMotivo,
    ubic:  iUbic
  };
}
