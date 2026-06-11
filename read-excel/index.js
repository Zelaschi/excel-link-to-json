const functions = require('@google-cloud/functions-framework');
const XLSX = require('xlsx');

functions.http('readExcel', async (req, res) => {
  res.set('Access-Control-Allow-Origin', '*');
  if (req.method === 'OPTIONS') {
    res.set('Access-Control-Allow-Methods', 'POST');
    res.set('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    return res.status(204).send('');
  }
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const apiKey = (req.headers['authorization'] || '').replace('Bearer ', '');
  if (!process.env.API_KEY || apiKey !== process.env.API_KEY) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  const { url, sheet, max_rows } = req.body || {};
  if (!url) return res.status(400).json({ error: 'Falta el campo "url"' });

  try {
    const response = await fetch(url);
    if (!response.ok) {
      return res.status(502).json({ error: `No se pudo descargar el archivo (HTTP ${response.status})` });
    }
    const buffer = Buffer.from(await response.arrayBuffer());
    const workbook = XLSX.read(buffer, { type: 'buffer' });

    const sheetNames = sheet ? [sheet] : workbook.SheetNames;
    const data = {};
    for (const name of sheetNames) {
      const ws = workbook.Sheets[name];
      if (!ws) {
        data[name] = { error: 'Hoja no encontrada' };
        continue;
      }
      let rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
      if (max_rows) rows = rows.slice(0, Number(max_rows));
      data[name] = rows;
    }

    return res.status(200).json({ ok: true, sheets: workbook.SheetNames, data });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'Error procesando el archivo', detail: err.message });
  }
});
