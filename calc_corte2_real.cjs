const fs = require('fs');
const XLSX = require('xlsx');

// 1. Read SEED_PDVS and SEED_ZONES from index.html
const html = fs.readFileSync('C:/Users/analista.retail/Downloads/quest-escuadron/index.html', 'utf8');
const seedZonesMatch = html.match(/const SEED_ZONES = (\[[\s\S]*?\]);/);
const seedPdvsMatch = html.match(/const SEED_PDVS = (\[[\s\S]*?\]);/);
const SEED_ZONES = eval(seedZonesMatch[1]);
const SEED_PDVS = eval(seedPdvsMatch[1]);

function findPDVMatch(rawText, rawLeader) {
  if (!rawText) return null;
  const textClean = rawText.trim().toLowerCase();
  const leaderClean = (rawLeader || '').trim().toLowerCase();

  const textCodeMatch = textClean.match(/^(fq|q|n)\s*0*(\d+)/i);
  let targetPrefix = textCodeMatch ? textCodeMatch[1].toLowerCase() : null;
  let targetNum = textCodeMatch ? parseInt(textCodeMatch[2], 10) : null;

  if (!targetPrefix) {
    if (textClean.includes('franquicia') || textClean.startsWith('fq')) targetPrefix = 'fq';
    else if (textClean.startsWith('q')) targetPrefix = 'q';
  }

  if (targetNum !== null) {
    const candidates = SEED_PDVS.filter(p => {
      const pCodeMatch = p.name.toLowerCase().match(/^(fq|q|n)\s*0*(\d+)/i);
      const pNum = pCodeMatch ? parseInt(pCodeMatch[2], 10) : (p.co ? parseInt(p.co, 10) : null);
      const pPrefix = pCodeMatch ? pCodeMatch[1].toLowerCase() : (p.name.toLowerCase().startsWith('fq') ? 'fq' : 'q');
      if (pNum === targetNum) {
        if (targetPrefix) return pPrefix === targetPrefix;
        return true;
      }
      return false;
    });

    if (candidates.length === 1) return candidates[0];
    if (candidates.length > 1) {
      if (leaderClean) {
        const byLeader = candidates.find(c => {
          const l = c.leader.toLowerCase();
          return leaderClean.includes(l) || l.includes(leaderClean) || leaderClean.split(' ').some(w => w.length > 3 && l.includes(w));
        });
        if (byLeader) return byLeader;
      }
      return candidates[0];
    }
  }

  for (const p of SEED_PDVS) {
    const pn = p.name.toLowerCase();
    if (textClean && (pn === textClean || textClean.includes(pn) || pn.includes(textClean))) {
      return p;
    }
  }
  return null;
}

// 2. Read full 23-30 Presupuesto from Consursos (14).xlsx
console.log('Reading Consursos (14).xlsx for total 23-30 budget...');
const wb14 = XLSX.readFile('C:/Users/analista.retail/Downloads/Consursos (14).xlsx');
const ws14 = wb14.Sheets[wb14.SheetNames[0]];
const rows14 = XLSX.utils.sheet_to_json(ws14);

const pptoTotalMap = {}; // by pdvId
let pptoRowsMatched = 0;

rows14.forEach(r => {
  if (r['Día de Fecha'] === 'Total') return;
  const desc = String(r['Desc. C.O.'] || '').trim();
  const leader = String(r['Lider'] || '').trim();
  if (!desc || desc.toLowerCase().includes('qst')) return;

  const pptoVal = Number(r[' Ppto Valor subtotal ']) || 0;
  const matched = findPDVMatch(desc, leader);
  if (matched) {
    pptoTotalMap[matched.id] = (pptoTotalMap[matched.id] || 0) + pptoVal;
    pptoRowsMatched++;
  }
});
console.log('Stores with 23-30 budget mapped:', Object.keys(pptoTotalMap).length);

// 3. Read Sales of 23 and 24 from Consursos (16).xlsx (Base sheet)
console.log('Reading Consursos (16).xlsx for sales of 23 and 24...');
const wb16 = XLSX.readFile('C:/Users/analista.retail/Downloads/Consursos (16).xlsx');
const ws16Base = wb16.Sheets['Base'];
const rows16 = XLSX.utils.sheet_to_json(ws16Base, { defval: '' });

const ventaTotalMap = {}; // by pdvId
let ventaMatchedCount = 0;

rows16.forEach(b => {
  const name = String(b['Nombre usuario'] || '').trim();
  const leader = String(b['Supervisor'] || '').trim();
  const venta = Number(b[' Venta '] || b['Venta']) || 0;
  if (!name || name.toLowerCase().includes('qst')) return;

  const matched = findPDVMatch(name, leader);
  if (matched) {
    ventaTotalMap[matched.id] = (ventaTotalMap[matched.id] || 0) + venta;
    ventaMatchedCount++;
  }
});
console.log('Stores with 23-24 sales mapped:', Object.keys(ventaTotalMap).length);

// 4. Calculate PDV compliance
const pdvDataMap = {};
SEED_PDVS.forEach(p => {
  const ppto = pptoTotalMap[p.id] || 0;
  const venta = ventaTotalMap[p.id] || 0;
  const comp = ppto > 0 ? Math.round((venta / ppto) * 1000) / 10 : 0;
  pdvDataMap[p.id] = comp;
});

// 5. Calculate Zone compliance
const zoneDataArr = [];
SEED_ZONES.forEach(z => {
  const pdvsInZone = SEED_PDVS.filter(p => p.zoneId === z.id);
  let zPpto = 0;
  let zVenta = 0;
  pdvsInZone.forEach(p => {
    zPpto += (pptoTotalMap[p.id] || 0);
    zVenta += (ventaTotalMap[p.id] || 0);
  });
  const comp = zPpto > 0 ? Math.round((zVenta / zPpto) * 1000) / 10 : 0;
  zoneDataArr.push({
    zoneId: z.id,
    compliance: comp
  });
  console.log(`Zona ${z.id} (${z.name}): ${comp}% (Venta 23-24: $${Math.round(zVenta).toLocaleString()} / Ppto Total 23-30: $${Math.round(zPpto).toLocaleString()})`);
});

// 6. Save payload and update Supabase
const newCutPayload = {
  id: 'corte-2-dias-23-24',
  name: 'Corte 2 - Días 23 y 24 (Avance 23-30)',
  date: '2026-09-24',
  data: zoneDataArr,
  pdv_data: pdvDataMap
};

console.log('Syncing to Supabase...');
fetch('https://aqgfocnbsjyhcpqfxrsa.supabase.co/rest/v1/quest_cortes?id=eq.corte-2-dias-23-24', {
  method: 'DELETE',
  headers: {
    'apikey': 'sb_publishable_jkaQRTZe82IDDPiXTb0jMg_bj19rK0U',
    'Authorization': 'Bearer sb_publishable_jkaQRTZe82IDDPiXTb0jMg_bj19rK0U'
  }
}).then(() => {
  return fetch('https://aqgfocnbsjyhcpqfxrsa.supabase.co/rest/v1/quest_cortes', {
    method: 'POST',
    headers: {
      'apikey': 'sb_publishable_jkaQRTZe82IDDPiXTb0jMg_bj19rK0U',
      'Authorization': 'Bearer sb_publishable_jkaQRTZe82IDDPiXTb0jMg_bj19rK0U',
      'Content-Type': 'application/json',
      'Prefer': 'resolution=merge-duplicates'
    },
    body: JSON.stringify([newCutPayload])
  });
})
.then(async res => {
  if (!res.ok) console.error('Error status:', res.status, await res.text());
  else console.log('SUCCESS! Corte 2 actualizado en Supabase con Presupuesto Total 23 al 30!');
})
.catch(err => console.error('Error:', err));
