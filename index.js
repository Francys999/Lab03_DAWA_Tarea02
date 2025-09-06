const http = require('http');
const ExcelJS = require('exceljs');

const PORT = 3000;

const server = http.createServer(async (req, res) => {
  try {
    if (req.url === '/reporte' && req.method === 'GET') {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Ventas');

      sheet.columns = [
        { header: 'Producto', key: 'producto', width: 25 },
        { header: 'Cantidad', key: 'cantidad', width: 12 },
        { header: 'Precio',   key: 'precio',   width: 12 },
      ];
      sheet.getRow(1).font = { bold: true };
      sheet.getColumn(3).numFmt = '$#,##0.00';

      const filas = Array.from({ length: 20 }, (_, i) => ({
        producto: `Producto ${i + 1}`,
        cantidad: Math.floor(Math.random() * 10) + 1,
        precio: Number((Math.random() * 100 + 5).toFixed(2)),
      }));
      sheet.addRows(filas);

      res.writeHead(200, {
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="reporte.xlsx"',
        'Cache-Control': 'no-store',
      });

      await workbook.xlsx.write(res);
      res.end();
      return;
    }

    res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Visita /reporte para descargar el Excel');
  } catch (err) {
    
    console.error('Error al generar/enviar el Excel:', err);
    if (!res.headersSent) {
      res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
    }
    res.end('Ocurrió un error al generar el reporte.');
  }
});

server.listen(PORT, () =>
  console.log(`Servidor listo en http://localhost:${PORT}  →  /reporte`)
);
