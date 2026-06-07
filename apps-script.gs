// ═══════════════════════════════════════════════════════════════
//  BODA MARINA & JAVIER – Google Apps Script
//
//  INSTRUCCIONES DE INSTALACIÓN:
//  1. Abre Google Sheets en drive.google.com → crear nueva hoja
//  2. Extensiones → Apps Script → pega este código → Guardar (💾)
//  3. Implementar → Nueva implementación
//       Tipo: Aplicación web
//       Ejecutar como: Yo
//       Quién puede acceder: Cualquier persona
//  4. Copia la URL que aparece
//  5. Pégala en index.html donde pone 'YOUR_APPS_SCRIPT_URL_HERE'
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME = 'Respuestas';

function doGet() {
  return ContentService
    .createTextOutput('OK – El formulario de Marina & Javier está activo.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    // Crear hoja si no existe
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Cabeceras si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      const headers = ['Fecha y hora', 'Asistencia', 'Nombre', 'Opción alimentaria', 'Autobús Alcañiz'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#C9A84C')
           .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const data      = JSON.parse(e.postData.contents);
    const timestamp = data.timestamp
                        ? new Date(data.timestamp).toLocaleString('es-ES')
                        : new Date().toLocaleString('es-ES');

    if (!data.attending) {
      // No asiste — una fila vacía de aviso
      sheet.appendRow([timestamp, 'NO', '', '', '']);
    } else {
      // Una fila por persona
      data.people.forEach(p => {
        sheet.appendRow([
          timestamp,
          'SÍ',
          p.name        || '',
          p.restriction || 'Ninguna',
          p.bus ? 'Sí' : 'No',
        ]);
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Función de prueba — ejecutadla manualmente para verificar que todo funciona
function testPost() {
  const fake = {
    postData: {
      contents: JSON.stringify({
        attending: true,
        timestamp: new Date().toISOString(),
        people: [
          { name: 'Ana García',  restriction: 'Celíaco/a',  bus: true  },
          { name: 'Luis García', restriction: '',            bus: false },
        ],
      }),
    },
  };
  Logger.log(doPost(fake).getContent());
}
