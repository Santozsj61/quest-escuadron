const fs = require('fs');
const XLSX = require('xlsx');

const html = fs.readFileSync('C:/Users/analista.retail/Downloads/quest-escuadron/index.html', 'utf8');
const seedZonesMatch = html.match(/const SEED_ZONES = (\[[\s\S]*?\]);/);
const seedPdvsMatch = html.match(/const SEED_PDVS = (\[[\s\S]*?\]);/);
const SEED_ZONES = eval(seedZonesMatch[1]);
const SEED_PDVS = eval(seedPdvsMatch[1]);

const wb = XLSX.readFile('C:/Users/analista.retail/Downloads/Consursos (16).xlsx');
const ws1 = wb.Sheets['Sheet 1'];
const ws2 = wb.Sheets['Base'];
const r1 = XLSX.utils.sheet_to_json(ws1, { defval: '' });
const r2 = XLSX.utils.sheet_to_json(ws2, { defval: '' });

// Ppto map from Sheet 1
const pptoMap = {};
r1.forEach(x => {
  const name = String(x['Desc. C.O.'] || '').trim().toLowerCase();
  pptoMap[name] = (pptoMap[name] || 0) + (Number(x['Ppto Valor subtotal']) || 0);
});

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

const pdvDataMap = {};
const pdvPptoMap = {};
const pdvVentaMap = {};

r2.forEach(b => {
  const name = String(b['Nombre usuario'] || '').trim();
  const leader = String(b['Supervisor'] || '').trim();
  const venta = Number(b[' Venta '] || b['Venta']) || 0;
  const ppto = pptoMap[name.toLowerCase()] || 0;

  if (name.toLowerCase().includes('qst')) return;

  const matched = findPDVMatch(name, leader);
  if (matched) {
    pdvPptoMap[matched.id] = (pdvPptoMap[matched.id] || 0) + ppto;
    pdvVentaMap[matched.id] = (pdvVentaMap[matched.id] || 0) + venta;
  }
});

SEED_PDVS.forEach(p => {
  const ppto = pdvPptoMap[p.id] || 0;
  const venta = pdvVentaMap[p.id] || 0;
  const comp = ppto > 0 ? Math.round((venta / ppto) * 1000) / 10 : 0;
  pdvDataMap[p.id] = comp;
});

// Zones
const zoneDataArr = [];
SEED_ZONES.forEach(z => {
  const pdvsInZone = SEED_PDVS.filter(p => p.zoneId === z.id);
  let zPpto = 0;
  let zVenta = 0;
  pdvsInZone.forEach(p => {
    zPpto += (pdvPptoMap[p.id] || 0);
    zVenta += (pdvVentaMap[p.id] || 0);
  });
  const comp = zPpto > 0 ? Math.round((zVenta / zPpto) * 1000) / 10 : 0;
  zoneDataArr.push({
    zoneId: z.id,
    compliance: comp
  });
});

const newCutPayload = {
  id: 'corte-2-dias-23-24',
  name: 'Corte 2 - Días 23 y 24 (Avance Acumulado)',
  date: '2026-09-24',
  data: zoneDataArr,
  pdv_data: pdvDataMap
};

console.log('Sending payload to Supabase...');
fetch('https://aqgfocnbsjyhcpqfxrsa.supabase.co/rest/v1/quest_cortes', {
  method: 'POST',
  headers: {
    'apikey': 'sb_publishable_jkaQRTZe82IDDPiXTb0jMg_bj19rK0U',
    'Authorization': 'Bearer sb_publishable_jkaQRTZe82IDDPiXTb0jMg_bj19rK0U',
    'Content-Type': 'application/json',
    'Prefer': 'resolution=merge-duplicates'
  },
  body: JSON.stringify([newCutPayload])
})
.then(async res => {
  if (!res.ok) {
    const txt = await res.text();
    console.error('Error status:', res.status, txt);
  } else {
    console.log('SUCCESS! Corte 2 guardado en Supabase exitosamente.');
  }
})
.catch(err => console.error('Fetch error:', err));
