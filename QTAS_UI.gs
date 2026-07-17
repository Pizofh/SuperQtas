function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('QTAS ERP')
    .addItem('Abrir ERP', 'mostrarAppQTAS')
    .addSeparator()
    .addItem('Aplicar vista operador', 'aplicarVistaOperadorQTAS')
    .addItem('Mostrar todas las hojas', 'mostrarTodasLasHojasQTAS')
    .addItem('Exportar libro TSV', 'exportarLibroTSVQTAS')
    .addSeparator()
    .addItem('Pausar inventario para carga historica', 'pausarInventarioParaCargaHistoricaQTAS')
    .addItem('Reanudar inventario operativo', 'reanudarInventarioOperativoQTAS')
    .addItem('Ver estado de inventario', 'mostrarEstadoSincronizacionInventarioQTAS')
    .addToUi();
}

function crearSalidaAppQTAS_(vistaInicial) {
  const template = HtmlService.createTemplateFromFile('App');
  template.vistaInicial = vistaInicial || 'ventas';

  return template
    .evaluate()
    .setTitle('SuperQTAS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerCorreosAutorizadosWebQTAS_() {
  const raw = valorPropiedadScriptQTAS_('QTAS_ALLOWED_WEB_EMAILS');
  const seen = {};

  return texto_(raw)
    .split(/[\n,;]+/)
    .map(value => texto_(value).toLowerCase())
    .filter(value => {
      if (!value || seen[value]) return false;
      seen[value] = true;
      return true;
    });
}

function estadoAccesoWebQTAS_() {
  const email = texto_(Session.getActiveUser().getEmail()).toLowerCase();
  const allowedEmails = obtenerCorreosAutorizadosWebQTAS_();

  return {
    email: email,
    allowedEmails: allowedEmails,
    configured: allowedEmails.length > 0,
    authorized: Boolean(email) && allowedEmails.indexOf(email) >= 0
  };
}

function crearSalidaAccesoWebRestringidoQTAS_(state) {
  const email = texto_(state && state.email);
  const configured = Boolean(state && state.configured);
  const title = configured ? 'Access restricted' : 'Web access not configured';
  const detail = configured
    ? 'This SuperQTAS web app is limited to authorized accounts.'
    : 'This SuperQTAS web app requires an allowlist of authorized emails.';
  const hint = configured
    ? 'Sign in with an authorized Google account or ask the administrator to add your email.'
    : 'Set QTAS_ALLOWED_WEB_EMAILS in Script Properties and deploy the web app as User accessing the web app.';
  const signedIn = email
    ? `<p><strong>Signed in as:</strong> ${email}</p>`
    : '<p><strong>Signed in as:</strong> unavailable</p>';

  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        body {
          margin: 0;
          min-height: 100vh;
          display: grid;
          place-items: center;
          padding: 24px;
          background: linear-gradient(135deg, #f6e8d7, #f4f0df 52%, #dce8d1);
          color: #231f20;
          font-family: Verdana, Geneva, sans-serif;
        }
        .card {
          width: min(100%, 560px);
          padding: 28px;
          border-radius: 24px;
          background: rgba(255, 250, 242, 0.94);
          border: 1px solid rgba(35, 31, 32, 0.08);
          box-shadow: 0 18px 40px rgba(35, 31, 32, 0.12);
        }
        h1 {
          margin: 0 0 10px;
          font-family: Georgia, "Times New Roman", serif;
          font-size: 32px;
        }
        p {
          margin: 0 0 10px;
          line-height: 1.5;
        }
        .muted {
          color: #5a5153;
        }
      </style>
    </head>
    <body>
      <div class="card">
        <h1>${title}</h1>
        <p>${detail}</p>
        ${signedIn}
        <p class="muted">${hint}</p>
      </div>
    </body>
    </html>
  `)
    .setTitle('SuperQTAS');
}

function guardarCorreosAutorizadosWebQTAS(correos) {
  const allowedEmails = Array.isArray(correos)
    ? correos
    : texto_(correos).split(/[\n,;]+/);
  const normalized = [];
  const seen = {};

  allowedEmails.forEach(value => {
    const email = texto_(value).toLowerCase();
    if (!email || seen[email]) return;
    seen[email] = true;
    normalized.push(email);
  });

  PropertiesService.getScriptProperties().setProperty(
    'QTAS_ALLOWED_WEB_EMAILS',
    normalized.join('\n')
  );

  return verCorreosAutorizadosWebQTAS();
}

function verCorreosAutorizadosWebQTAS() {
  return {
    ok: true,
    allowedEmails: obtenerCorreosAutorizadosWebQTAS_(),
    state: estadoAccesoWebQTAS_()
  };
}

function doGet(e) {
  const access = estadoAccesoWebQTAS_();
  if (!access.authorized) {
    return crearSalidaAccesoWebRestringidoQTAS_(access);
  }

  const vista = e && e.parameter ? e.parameter.vista : 'ventas';
  return crearSalidaAppQTAS_(vista);
}

function mostrarAppQTAS() {
  const html = crearSalidaAppQTAS_('ventas')
    .setWidth(1280)
    .setHeight(920);

  SpreadsheetApp.getUi().showModalDialog(html, 'SuperQTAS');
}

function getCatalogoQTAS(fechaVenta) {
  validarModeloSoloLecturaQTAS_({
    sheetNames: [
      QTAS.sheets.productos,
      QTAS.sheets.precios,
      QTAS.sheets.clientes
    ],
    validarConfig: true
  });

  const ss = SpreadsheetApp.getActive();
  const fechaConsulta = resolverFechaOperacion_(fechaVenta, new Date());
  const productos = leerObjetos_(ss.getSheetByName(QTAS.sheets.productos))
    .filter(row => estaActivo_(row.Activo))
    .map(row => ({
      producto: texto_(row.Producto_Estandar),
      unidad: texto_(row.Unidad_Default)
  }));

  const priceCache = cargarPreciosReferenciaEnMemoria_();
  cargarReglasDistribucionEnMemoria_();

  const productosConPrecio = productos.map(producto => ({
    producto: producto.producto,
    unidad: producto.unidad,
    precio: obtenerPrecioVigenteDesdeCache_(
      priceCache,
      producto.producto,
      producto.unidad,
      fechaConsulta
    )
  }));

  const mediosPago = leerMediosPagoConfiguradosQTAS_()
    .filter(row => row.activo)
    .map(row => row.medioPago);

  const clientes = leerObjetos_(ss.getSheetByName(QTAS.sheets.clientes))
    .filter(row => estaActivo_(row.Activo))
    .map(row => ({
      clienteId: texto_(row.Cliente_ID),
      nombre: texto_(row.Nombre)
    }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre));

  return {
    productos: productosConPrecio,
    mediosPago,
    clientes,
    hoy: fechaInput_(new Date())
  };
}

function getDashboardQTAS() {
  validarModeloSoloLecturaQTAS_({
    sheetNames: [QTAS.sheets.ventas, QTAS.sheets.pagos],
    validarConfig: false
  });
  return dashboardVentasConsistenteQTAS_();
}
