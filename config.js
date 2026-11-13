// Configuración del dashboard. Solo edita este archivo.
window.APP_CONFIG = {
  // Id de aplicación (cliente) del registro en Microsoft Entra (ver README.md)
  clientId: 'PEGAR-AQUI-EL-ID-DE-APLICACION',
  // 'organizations' funciona para cualquier cuenta de trabajo; o pon el Id de tu tenant
  tenant: 'organizations',
  // Biblioteca de SharePoint (drive) donde viven los archivos y el Id de cada archivo
  driveId: 'b!SmRLu76aXE6slyY60dxlZYGy-CKjnExIqBuhfV9qRi87jVvb3AprSpErISiudpJv',
  files: {
    matriz: '01LSQ5723TR5LRXLVXIFBLNVTDW3SAFOUB',        // Matriz de compras internacionales.xlsx (original, solo lectura)
    recibidas: '01LSQ5727FNEJXJKYFW5CZBG27VEIFCR5B',     // Recibidas resumen.csv (flujo diario 10 pm)
    inspecciones: '01LSQ5724N3EP2SOAWKFHKDE5CKRY3UPNZ'   // Inspecciones resumen.csv (flujo diario 10 pm)
  }
};
