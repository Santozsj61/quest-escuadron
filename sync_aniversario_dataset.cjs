const fs = require('fs');
const XLSX = require('D:/analista.retail/.gemini/antigravity/scratch/control-turnos-pdv/node_modules/xlsx');

const excelPath = 'D:/analista.retail/.gemini/antigravity/scratch/quest-escuadron/ppto_ventas_aniversario.xlsx';
const wb = XLSX.readFile(excelPath);
const rows = XLSX.utils.sheet_to_json(wb.Sheets['Base2']);

const html = fs.readFileSync('D:/analista.retail/.gemini/antigravity/scratch/quest-escuadron/index.html', 'utf8');
const SEED_PDVS = eval(html.match(/const SEED_PDVS = (\[[\s\S]*?\]);/)[1]);
const SEED_ZONES = eval(html.match(/const SEED_ZONES = (\[[\s\S]*?\]);/)[1]);

function findMatch(rawText) {
  if (!rawText) return null;
  const t = rawText.trim().toLowerCase();
  const m = t.match(/^(fq|q|n)\s*0*(\d+)/i);
  if (m) {
    const prefix = m[1].toLowerCase();
    const num = parseInt(m[2], 10);
    const found = SEED_PDVS.find(p => {
      const pm = p.name.toLowerCase().match(/^(fq|q|n)\s*0*(\d+)/i);
      if (pm && pm[1].toLowerCase() === prefix && parseInt(pm[2], 10) === num) return true;
      if (p.co && parseInt(p.co, 10) === num && (p.name.toLowerCase().startsWith(prefix) || (prefix === 'fq' && p.name.toLowerCase().includes('franquicia')))) return true;
      return false;
    });
    if (found) return found;
  }
  for (const p of SEED_PDVS) {
    const pn = p.name.toLowerCase();
    if (pn === t || t.includes(pn) || pn.includes(t)) return p;
  }
  return null;
}

const timelineDays = [
  { mes: 'septiembre', dia: 23, id: 'corte-01-23sep', date: '2026-09-23', name: 'Corte 1 - Día 23 Sep (Inicio de Campaña)' },
  { mes: 'septiembre', dia: 24, id: 'corte-02-24sep', date: '2026-09-24', name: 'Corte 2 - Acumulado al 24 Sep' },
  { mes: 'septiembre', dia: 25, id: 'corte-03-25sep', date: '2026-09-25', name: 'Corte 3 - Acumulado al 25 Sep' },
  { mes: 'septiembre', dia: 26, id: 'corte-04-26sep', date: '2026-09-26', name: 'Corte 4 - Acumulado al 26 Sep' },
  { mes: 'septiembre', dia: 27, id: 'corte-05-27sep', date: '2026-09-27', name: 'Corte 5 - Acumulado al 27 Sep (Fin de Semana 1)' },
  { mes: 'septiembre', dia: 28, id: 'corte-06-28sep', date: '2026-09-28', name: 'Corte 6 - Acumulado al 28 Sep' },
  { mes: 'septiembre', dia: 29, id: 'corte-07-29sep', date: '2026-09-29', name: 'Corte 7 - Acumulado al 29 Sep' },
  { mes: 'septiembre', dia: 30, id: 'corte-08-30sep', date: '2026-09-30', name: 'Corte 8 - Acumulado al 30 Sep (Cierre Mes Septiembre)' },
  { mes: 'octubre', dia: 1, id: 'corte-09-01oct', date: '2026-10-01', name: 'Corte 9 - Acumulado al 01 Oct' },
  { mes: 'octubre', dia: 2, id: 'corte-10-02oct', date: '2026-10-02', name: 'Corte 10 - Acumulado al 02 Oct' },
  { mes: 'octubre', dia: 3, id: 'corte-11-03oct', date: '2026-10-03', name: 'Corte 11 - Acumulado al 03 Oct' },
  { mes: 'octubre', dia: 4, id: 'corte-12-04oct', date: '2026-10-04', name: 'Corte 12 - Acumulado al 04 Oct (Fin de Semana 2)' },
  { mes: 'octubre', dia: 5, id: 'corte-13-05oct', date: '2026-10-05', name: 'Corte 13 - Acumulado al 05 Oct' },
  { mes: 'octubre', dia: 6, id: 'corte-14-06oct', date: '2026-10-06', name: 'Corte 14 - Acumulado al 06 Oct (Recta Final)' },
  { mes: 'octubre', dia: 7, id: 'corte-final-23sep-07oct', date: '2026-10-07', name: 'Corte Final - Cierre Campaña Aniversario (23 Sep - 07 Oct)' }
];

const allCutsPayload = [];

for (let i = 0; i < timelineDays.length; i++) {
  const targetDay = timelineDays[i];
  
  // Filter rows accumulated from day 0 to i
  const validDays = timelineDays.slice(0, i + 1);
  const cumRows = rows.filter(r => {
    return validDays.some(vd => vd.mes === r['Mes de Fecha'] && Number(r['Día de Fecha']) === vd.dia);
  });

  const pdvPpto = {}, pdvVenta = {};
  cumRows.forEach(r => {
    const m = findMatch(r['Desc. C.O.']);
    if (m) {
      pdvPpto[m.id] = (pdvPpto[m.id] || 0) + (Number(r['Ppto Valor subtotal']) || 0);
      pdvVenta[m.id] = (pdvVenta[m.id] || 0) + (Number(r['Valor subtotal']) || 0);
    }
  });

  const pdvDataMap = {};
  SEED_PDVS.forEach(p => {
    const pVal = pdvPpto[p.id] || 0;
    const vVal = pdvVenta[p.id] || 0;
    const comp = pVal > 0 ? Math.round((vVal / pVal) * 1000) / 10 : 0;
    pdvDataMap[p.id] = comp;
  });

  const zoneDataArr = SEED_ZONES.map(z => {
    const pdvsInZ = SEED_PDVS.filter(p => p.zoneId === z.id);
    let zPpto = 0, zVenta = 0;
    pdvsInZ.forEach(p => {
      zPpto += (pdvPpto[p.id] || 0);
      zVenta += (pdvVenta[p.id] || 0);
    });
    const comp = zPpto > 0 ? Math.round((zVenta / zPpto) * 1000) / 10 : 0;
    return {
      zoneId: z.id,
      compliance: comp
    };
  });

  allCutsPayload.push({
    id: targetDay.id,
    name: targetDay.name,
    date: targetDay.date,
    data: zoneDataArr,
    pdv_data: pdvDataMap
  });
}

console.log('Generated', allCutsPayload.length, 'cuts.');
const finalCut = allCutsPayload[allCutsPayload.length - 1];
console.log('Final Cut:', finalCut.name);
console.log('Zone Compliances in Final Cut:');
finalCut.data.forEach(z => {
  const zInfo = SEED_ZONES.find(x => x.id === z.zoneId);
  console.log(`- ${zInfo.name}: ${z.compliance}%`);
});

fs.writeFileSync('D:/analista.retail/.gemini/antigravity/scratch/quest-escuadron/aniversario_cuts_payload.json', JSON.stringify(allCutsPayload, null, 2));
console.log('Saved aniversario_cuts_payload.json');

// Upload to Supabase
async function uploadToSupabase() {
  console.log('Uploading all cuts to Supabase...');
  const res = await fetch('https://aqgfocnbsjyhcpqfxrsa.supabase.co/rest/v1/quest_cortes', {
    method: 'POST',
    headers: {
      'apikey': 'sb_publishable_jkaQRTZe82IDDPiXTb0jMg_bj19rK0U',
      'Authorization': 'Bearer sb_publishable_jkaQRTZe82IDDPiXTb0jMg_bj19rK0U',
      'Content-Type': 'application/json',
      'Prefer': 'resolution=merge-duplicates'
    },
    body: JSON.stringify(allCutsPayload)
  });

  if (!res.ok) {
    const txt = await res.text();
    console.error('Error uploading to Supabase:', res.status, txt);
  } else {
    console.log('SUCCESS! All 15 official cuts uploaded to Supabase successfully.');
  }
}

uploadToSupabase();
