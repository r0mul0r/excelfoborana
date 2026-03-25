'use strict';

const express  = require('express');
const multer   = require('multer');
const XLSX     = require('xlsx');
const path     = require('path');

const app    = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits:  { fileSize: 10 * 1024 * 1024 }, // 10 MB
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls') return cb(null, true);
    cb(new Error('Solo se aceptan archivos .xlsx o .xls'));
  }
});

// ── Utilidades ────────────────────────────────────────────────────────────────

/**
 * Convierte distintos formatos de monto a número.
 * Soporta: '10.000,00' | '10,000.00' | '10000' | 10000 (número)
 */
function parseMonto(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === 'number') return value;

  let text = String(value).trim();

  // Quitar símbolo de moneda, espacios, letras
  text = text.replace(/[^\d.,\-]/g, '');

  if (!text || text === '.' || text === ',' || text === '-') return 0;

  const hasPoint = text.includes('.');
  const hasComma = text.includes(',');

  if (hasPoint && hasComma) {
    const lastDot   = text.lastIndexOf('.');
    const lastComma = text.lastIndexOf(',');

    if (lastComma > lastDot) {
      // Coma es decimal: 10.000,00 → europeo
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      // Punto es decimal: 10,000.00 → anglosajón
      text = text.replace(/,/g, '');
    }
  } else if (hasComma) {
    // Solo coma: decidir si es miles (1,000) o decimal (1,5)
    const parts = text.split(',');
    if (parts.length === 2 && parts[1].length === 3 && parts[0].length > 0) {
      // Coma como miles: 10,000
      text = text.replace(',', '');
    } else {
      // Coma como decimal: 10,5
      text = text.replace(',', '.');
    }
  }
  // Si solo tiene punto, ya está en formato correcto

  const num = parseFloat(text);
  return isNaN(num) ? 0 : num;
}

/**
 * Busca las columnas USUARIO y MONTO RECARGADO de forma flexible.
 */
function findColumns(headers) {
  let usuario = null;
  let monto   = null;

  for (const h of headers) {
    const norm = String(h).trim().toUpperCase();
    if (['USUARIO', 'USER', 'USUARIOS'].includes(norm)) {
      usuario = h;
    } else if (norm.includes('MONTO') || norm.includes('IMPORTE') || norm.includes('AMOUNT')) {
      monto = h;
    }
  }
  return { usuario, monto };
}

/**
 * Formatea un número como string estilo europeo: 10.000,00
 */
function formatMonto(num) {
  return num.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ── Procesamiento Excel ───────────────────────────────────────────────────────

function processExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Leer como array de objetos
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

  if (!rows.length) throw new Error('El archivo está vacío o no tiene datos.');

  const headers = Object.keys(rows[0]);
  const { usuario: colUsuario, monto: colMonto } = findColumns(headers);

  if (!colUsuario) throw new Error('No se encontró la columna USUARIO en el archivo.');
  if (!colMonto)   throw new Error('No se encontró la columna MONTO RECARGADO en el archivo.');

  // Agrupar por usuario
  const grupos = new Map();

  for (const row of rows) {
    const user   = String(row[colUsuario] ?? '').trim();
    const amount = parseMonto(row[colMonto]);

    if (!user || user === 'null' || user === 'undefined') continue;

    if (!grupos.has(user)) {
      grupos.set(user, { count: 0, total: 0 });
    }
    const g = grupos.get(user);
    g.count++;
    g.total += amount;
  }

  if (!grupos.size) throw new Error('No se encontraron datos válidos en el archivo.');

  // Ordenar por monto total descendente
  const sorted = [...grupos.entries()].sort((a, b) => b[1].total - a[1].total);

  return buildOutputExcel(sorted);
}

function buildOutputExcel(data) {
  const wb = XLSX.utils.book_new();

  // Construir array de filas
  const header = ['USUARIO', 'CANTIDAD DE RECARGAS', 'MONTO TOTAL RECARGADO'];

  let totalCount = 0;
  let totalMonto = 0;

  const rows = data.map(([user, { count, total }]) => {
    totalCount += count;
    totalMonto += total;
    return [user, count, formatMonto(total)];
  });

  // Fila de totales
  rows.push(['TOTAL', totalCount, formatMonto(totalMonto)]);

  const wsData = [header, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Ancho de columnas
  ws['!cols'] = [{ wch: 25 }, { wch: 22 }, { wch: 25 }];

  XLSX.utils.book_append_sheet(wb, ws, 'Resumen Recargas');

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// ── Rutas ─────────────────────────────────────────────────────────────────────

app.get('/', (_req, res) => {
  res.sendFile(path.join(__dirname, 'views', 'index.html'));
});

app.post('/procesar', (req, res) => {
  upload.single('archivo')(req, res, (err) => {
    if (err) {
      const status = err.message.includes('Solo se aceptan') ? 400 : 413;
      return res.status(status).json({ error: err.message });
    }

    if (!req.file) {
      return res.status(400).json({ error: 'No se envió ningún archivo.' });
    }

    try {
      const outputBuffer = processExcel(req.file.buffer);

      res.setHeader('Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition',
        'attachment; filename="resumen_recargas.xlsx"');
      res.send(outputBuffer);

    } catch (e) {
      const status = e.message.includes('No se encontró') ? 422 : 500;
      res.status(status).json({ error: e.message });
    }
  });
});

// ── Inicio ────────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});

module.exports = app;
