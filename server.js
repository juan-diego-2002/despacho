const express = require('express');
const multer = require('multer');
const path = require('path');
const { procesarLoteCfdi } = require('./src/cfdiProcessor');
const { crearExcel } = require('./src/excelReport');

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024, files: 3000 }
});

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/procesar', upload.fields([
  { name: 'emitidas', maxCount: 1500 },
  { name: 'recibidas', maxCount: 1500 }
]), async (req, res) => {
  try {
    const filtro = {
      modo: req.body.modo || 'mensual',
      anio: Number(req.body.anio || new Date().getFullYear()),
      mes: Number(req.body.mes || 1),
      trimestre: Number(req.body.trimestre || 1)
    };
    const resultado = await procesarLoteCfdi({
      emitidas: req.files.emitidas || [],
      recibidas: req.files.recibidas || [],
      filtro
    });
    res.json(resultado);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.post('/api/exportar', express.json({ limit: '20mb' }), async (req, res) => {
  try {
    const buffer = await crearExcel(req.body);
    res.setHeader('Content-Type', 'application/vnd.ms-excel; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="cedula-cfdi.xls"');
    res.send(buffer);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Cedulas CFDI disponible en http://localhost:${port}`);
});
