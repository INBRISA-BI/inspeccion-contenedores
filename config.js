// Configuración del dashboard. Solo edita este archivo.
window.APP_CONFIG = {
  // Dirección del sitio (única URL registrada en Entra como redirección)
  redirectUri: 'https://inbrisa-bi.github.io/inspeccion-contenedores/',
  fichasUrl: 'https://inbrisa-bi.github.io/fichas-calidad/',
  // Sitio de SharePoint y listas usadas por el registro de producto off spec
  site: { host: 'inbrisa.sharepoint.com', path: '/sites/Produccion' },
  lists: {
    productos: 'a360b1ef-db85-4b6b-86de-d7c65c843782',   // Info de Productos (producto terminado)
    ps: 'fc1da8a8-c62c-4bd5-a5c8-0989f46fc9fc',          // Info de PS (semielaborados)
    offspec: 'a7f18829-2e8d-4c7b-b3cf-1f8b219647f9'      // CALI_ProductoOffSpec (registros)
  },
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
