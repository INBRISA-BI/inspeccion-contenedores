// Configuración del dashboard. Solo edita este archivo.
window.APP_CONFIG = {
  // Id de aplicación (cliente) del registro en Microsoft Entra (ver README.md)
  clientId: '5f4d98f1-0d8f-48a2-b4fa-d08d14d5ac1f',
  // 'organizations' funciona para cualquier cuenta de trabajo; o pon el Id de tu tenant
  tenant: '7842b47f-7df1-435d-9880-eda8e4177489',
  // Biblioteca de SharePoint (drive) donde viven los archivos y el Id de cada archivo
  driveId: 'b!SmRLu76aXE6slyY60dxlZYGy-CKjnExIqBuhfV9qRi87jVvb3AprSpErISiudpJv',
  files: {
    matriz: '01LSQ5723TR5LRXLVXIFBLNVTDW3SAFOUB',        // Matriz de compras internacionales.xlsx (original, solo lectura)
    recibidas: '01LSQ5727FNEJXJKYFW5CZBG27VEIFCR5B',     // Recibidas resumen.csv (flujo diario 10 pm)
    inspecciones: '01LSQ5724N3EP2SOAWKFHKDE5CKRY3UPNZ',
    historial: '01LSQ5722NTXLFAUEW6NALFKUOT6DPY23T'   // Inspecciones historial.csv (inspecciones antiguas ya depuradas de la lista)   // Inspecciones resumen.csv (flujo diario 10 pm)
  }
};
