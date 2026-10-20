/* DendroRed — aplicación principal
 * Vistas: mapa público, ficha de sitio, formularios de sitio y cronología,
 * importador, cuenta, mis registros, administración y "acerca de".
 */
(function () {
  'use strict';

  const CFG = window.DENDRORED_CONFIG || {};
  const C = window.DRCoord;
  const I = window.DRImport;

  // ===================================================================
  // Utilidades
  // ===================================================================
  const $ = (sel, raiz = document) => raiz.querySelector(sel);
  const $$ = (sel, raiz = document) => [...raiz.querySelectorAll(sel)];

  // Todo texto que viene de la base pasa por aquí antes de ir al HTML.
  function esc(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function urlSegura(u) {
    if (!u) return null;
    try {
      const x = new URL(String(u).trim());
      return x.protocol === 'http:' || x.protocol === 'https:' ? x.href : null;
    } catch (e) { return null; }
  }
  function enlaceDOI(doi) {
    if (!doi) return null;
    const limpio = String(doi).trim().replace(/^(https?:\/\/)?(dx\.)?doi\.org\//i, '').replace(/^doi:\s*/i, '');
    if (!/^10\.\d{4,9}\/\S+$/.test(limpio)) return null;
    return { texto: limpio, url: 'https://doi.org/' + encodeURI(limpio) };
  }
  const nf = (n, dec) => (n === null || n === undefined || n === '') ? '—'
    : Number(n).toLocaleString('es-CL', dec === undefined ? {} : { minimumFractionDigits: dec, maximumFractionDigits: dec });
  const anio = (n) => (n === null || n === undefined) ? '—' : (n < 0 ? Math.abs(n) + ' a. C.' : String(n));
  const periodo = (a, b) => (a === null || a === undefined) && (b === null || b === undefined) ? '—' : anio(a) + '–' + anio(b);
  const fecha = (t) => t ? new Date(t).toLocaleDateString('es-CL', { day: 'numeric', month: 'short', year: 'numeric' }) : '—';
  const fechaHora = (t) => t ? new Date(t).toLocaleString('es-CL', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';
  const vacioANulo = (v) => (v === undefined || v === null || String(v).trim() === '') ? null : String(v).trim();
  const plural = (n, uno, varios) => n + ' ' + (n === 1 ? uno : varios);

  function aviso(texto, tipo) {
    const d = document.createElement('div');
    d.className = 'aviso-flotante' + (tipo === 'error' ? ' error' : '');
    d.textContent = texto;
    $('#avisos').appendChild(d);
    setTimeout(() => d.remove(), tipo === 'error' ? 7000 : 4000);
  }

  // Traduce errores de la base a mensajes que se entienden.
  function mensajeError(error) {
    if (!error) return 'Ocurrió un error desconocido';
    const m = error.message || String(error);
    const codigo = error.code || '';
    if (codigo === '23505') {
      if (/sitios_codigo_key/.test(m)) return 'Ya existe un sitio con ese código. Elige otro o agrega tu cronología al sitio existente.';
      if (/cronologias_sitio_id_codigo_key/.test(m)) return 'Este sitio ya tiene una cronología con ese código.';
      if (/especies_nombre_cientifico_key/.test(m)) return 'Esa especie ya está en el catálogo.';
      if (/especies_codigo_itrdb_key/.test(m)) return 'Ese código ITRDB ya está asignado a otra especie.';
      if (/laboratorios_nombre_key/.test(m)) return 'Ya existe un laboratorio con ese nombre.';
      return 'Ya existe un registro con esos datos.';
    }
    if (codigo === '23503') return 'No se puede borrar: hay otros registros que dependen de este.';
    if (codigo === '23514') {
      const regla = (m.match(/constraint "([^"]+)"/) || [])[1] || '';
      if (/orcid/.test(regla)) return 'El ORCID no tiene el formato 0000-0000-0000-0000.';
      if (/codigo/.test(regla)) return 'El código solo admite mayúsculas, números, _ y -, entre 2 y 16 caracteres.';
      if (/latitud|longitud/.test(regla)) return 'Las coordenadas están fuera de rango.';
      return 'Algún valor está fuera del rango permitido (' + regla + ').';
    }
    if (codigo === '42501' || /row-level security/.test(m)) return 'No tienes permiso para hacer esto. Si tu cuenta es nueva, espera a que un administrador la apruebe.';
    if (/Invalid login credentials/i.test(m)) return 'Correo o contraseña incorrectos.';
    if (/Email not confirmed/i.test(m)) return 'Aún no confirmas tu correo. Revisa tu bandeja de entrada.';
    if (/User already registered/i.test(m)) return 'Ya hay una cuenta con ese correo. Ingresa o recupera tu contraseña.';
    if (/Password should be at least/i.test(m)) return 'La contraseña debe tener al menos 8 caracteres.';
    if (/Failed to fetch|NetworkError/i.test(m)) return 'No se pudo conectar con la base de datos. Revisa tu conexión.';
    if (/Solo un administrador/.test(m)) return m;
    return m;
  }

  function cargarScript(url) {
    return new Promise((ok, mal) => {
      if ($(`script[src="${url}"]`)) return ok();
      const s = document.createElement('script');
      s.src = url; s.onload = ok; s.onerror = () => mal(new Error('No se pudo cargar ' + url));
      document.head.appendChild(s);
    });
  }
  const URL_XLSX = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';

  function descargar(nombre, contenido, tipo) {
    const blob = contenido instanceof Blob ? contenido : new Blob([contenido], { type: tipo });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  // Ordena tablas haciendo clic en el encabezado.
  function ordenarFilas(filas, campo, asc) {
    return [...filas].sort((a, b) => {
      const x = a[campo], y = b[campo];
      if (x === null || x === undefined || x === '') return 1;
      if (y === null || y === undefined || y === '') return -1;
      const r = typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), 'es', { numeric: true });
      return asc ? r : -r;
    });
  }

  // ===================================================================
  // Textos de los valores fijos
  // ===================================================================
  const ETIQ = {
    parametro: { ancho_total: 'Ancho total', madera_temprana: 'Madera temprana', madera_tardia: 'Madera tardía', densidad_maxima: 'Densidad máxima', anatomia: 'Anatomía de la madera', isotopos: 'Isótopos estables', otro: 'Otro' },
    disponibilidad: { itrdb: 'Publicada en el ITRDB', repositorio: 'En otro repositorio', a_pedido: 'A pedido del responsable', no_disponible: 'No disponible' },
    disponibilidadCorta: { itrdb: 'ITRDB', repositorio: 'Repositorio', a_pedido: 'A pedido', no_disponible: 'No disponible' },
    estado: { en_curso: 'En curso', terminada: 'Terminada', publicada: 'Publicada' },
    rol: { investigador: 'Investigador', curador: 'Curador', administrador: 'Administrador' },
    tabla: { sitios: 'Sitio', cronologias: 'Cronología', cronologia_autores: 'Autores' },
    accion: { crear: 'Creó', modificar: 'Modificó', borrar: 'Borró' }
  };
  const REGIONES = ['Arica y Parinacota', 'Tarapacá', 'Antofagasta', 'Atacama', 'Coquimbo', 'Valparaíso',
    'Metropolitana de Santiago', "Libertador General Bernardo O'Higgins", 'Maule', 'Ñuble', 'Biobío', 'La Araucanía',
    'Los Ríos', 'Los Lagos', 'Aysén del General Carlos Ibáñez del Campo', 'Magallanes y de la Antártica Chilena'];
  const opciones = (obj, actual, vacio) =>
    (vacio !== undefined ? `<option value="">${esc(vacio)}</option>` : '') +
    Object.entries(obj).map(([v, t]) => `<option value="${esc(v)}"${v === actual ? ' selected' : ''}>${esc(t)}</option>`).join('');

  // ===================================================================
  // Estado y permisos
  // ===================================================================
  let sb = null;
  const estado = {
    sesion: null, perfil: null,
    especies: [], laboratorios: [], catalogo: [], perfilesPublicos: [],
    cargado: false,
    filtros: { texto: '', especie: '', region: '', laboratorio: '', desde: '', hasta: '', publica: false },
    orden: { campo: 'codigo', asc: true },
    modo: 'mapa'
  };
  const uid = () => estado.sesion && estado.sesion.user ? estado.sesion.user.id : null;
  const puede = {
    aportar: () => !!(estado.perfil && estado.perfil.aprobado),
    curar: () => !!(estado.perfil && estado.perfil.aprobado && ['curador', 'administrador'].includes(estado.perfil.rol)),
    administrar: () => !!(estado.perfil && estado.perfil.aprobado && estado.perfil.rol === 'administrador'),
    editar: (reg) => !!reg && (puede.curar() || (puede.aportar() && reg.creado_por === uid()))
  };

  // Trae todas las filas de una consulta, de a 1000 (límite de la API).
  async function traerTodo(armarConsulta) {
    const tam = 1000;
    let desde = 0, todas = [];
    for (;;) {
      const { data, error } = await armarConsulta().range(desde, desde + tam - 1);
      if (error) throw error;
      todas = todas.concat(data);
      if (data.length < tam) break;
      desde += tam;
    }
    return todas;
  }

  async function cargarPerfil() {
    estado.perfil = null;
    if (!uid()) return;
    const { data, error } = await sb.from('perfiles')
      .select('id, nombre_completo, orcid, laboratorio_id, rol, aprobado')
      .eq('id', uid()).maybeSingle();
    if (!error) estado.perfil = data;
  }

  async function cargarBase(forzar) {
    if (estado.cargado && !forzar) return;
    const [esp, lab, cat] = await Promise.all([
      traerTodo(() => sb.from('especies').select('*').order('nombre_cientifico')),
      traerTodo(() => sb.from('laboratorios').select('*').order('nombre')),
      traerTodo(() => sb.from('catalogo').select('*').order('codigo'))
    ]);
    estado.especies = esp;
    estado.laboratorios = lab;
    estado.catalogo = cat.map((s) => Object.assign(s, { latitud: Number(s.latitud), longitud: Number(s.longitud) }));
    estado.cargado = true;
  }

  // ===================================================================
  // Navegación
  // ===================================================================
  function pintarNav() {
    const ruta = location.hash || '#mapa';
    const enlace = (href, texto, extra = '') =>
      `<a href="${href}"${ruta === href || (href === '#mapa' && ruta === '') ? ' aria-current="page"' : ''} ${extra}>${texto}</a>`;
    let html = enlace('#mapa', 'Mapa');
    if (puede.aportar()) {
      html += enlace('#sitio/nuevo', 'Nuevo sitio') + enlace('#importar', 'Importar') + enlace('#mis-registros', 'Mis registros');
    }
    if (puede.curar()) html += enlace('#admin', 'Administrar');
    html += enlace('#acerca', 'Acerca de') + enlace('#contacto', 'Contacto');
    if (uid()) {
      const nombre = (estado.perfil && estado.perfil.nombre_completo) || 'Mi cuenta';
      const pendiente = estado.perfil && !estado.perfil.aprobado;
      html += `<a href="#cuenta" class="empuje cuenta-enlace"${ruta === '#cuenta' ? ' aria-current="page"' : ''}>`
        + (pendiente ? '<span class="punto-estado" title="Cuenta pendiente de aprobación"></span>' : '')
        + esc(nombre.split(' ')[0]) + '</a>';
    } else {
      html += `<a href="#cuenta" class="empuje"${ruta === '#cuenta' ? ' aria-current="page"' : ''}>Ingresar</a>`;
    }
    $('#nav').innerHTML = html;
  }

  const RUTAS = [
    [/^#?(mapa)?$/, (m, t) => vistaMapa(t)],
    [/^#sitio\/nuevo$/, (m, t) => exigirAporte() && vistaFormSitio(null, t)],
    [/^#sitio\/(\d+)$/, (m, t) => vistaFicha(+m[1], t)],
    [/^#sitio\/(\d+)\/editar$/, (m, t) => exigirAporte() && vistaFormSitio(+m[1], t)],
    [/^#sitio\/(\d+)\/cronologia$/, (m, t) => exigirAporte() && vistaFormCrono(null, +m[1], t)],
    [/^#cronologia\/(\d+)\/editar$/, (m, t) => exigirAporte() && vistaFormCrono(+m[1], null, t)],
    [/^#importar$/, (m, t) => exigirAporte() && vistaImportar(t)],
    [/^#mis-registros$/, (m, t) => exigirAporte() && vistaMisRegistros(t)],
    [/^#admin$/, (m, t) => vistaAdmin(t)],
    [/^#cuenta$/, (m, t) => vistaCuenta(t)],
    [/^#nueva-clave$/, () => vistaNuevaClave()],
    [/^#acerca$/, () => vistaAcerca()],
    [/^#contacto$/, () => vistaContacto()]
  ];

  // Cada navegación abre un "turno". Una vista que termina de cargar cuando ya
  // se navegó a otra parte no debe dibujar encima de la nueva.
  let turnoVista = 0;
  const vigente = (t) => t === turnoVista;

  async function enrutar() {
    const turno = ++turnoVista;
    let ruta = location.hash;
    // Los enlaces de confirmación de correo traen la sesión en el hash
    if (/access_token=|error_description=/.test(ruta)) {
      const recuperar = /type=recovery/.test(ruta);
      const err = decodeURIComponent((ruta.match(/error_description=([^&]+)/) || [])[1] || '').replace(/\+/g, ' ');
      history.replaceState(null, '', location.pathname + (recuperar ? '#nueva-clave' : '#cuenta'));
      if (err) aviso('El enlace no es válido o ya expiró: ' + err, 'error');
      ruta = location.hash;
    }
    pintarNav();
    window.scrollTo(0, 0);
    for (const [patron, fn] of RUTAS) {
      const m = ruta.match(patron);
      if (m) {
        try { await fn(m, turno); } catch (e) {
          if (e && e.obsoleta) return;
          console.error(e);
          if (vigente(turno)) mostrar(`<div class="hoja"><div class="mensaje error"><p>${esc(mensajeError(e))}</p></div></div>`);
        }
        return;
      }
    }
    location.hash = '#mapa';
  }

  // Muestra una vista de contenido (oculta el mapa principal). Cada vista se dibuja
  // en un contenedor nuevo; si el turno ya pasó, no dibuja nada.
  function mostrar(html, turno = turnoVista) {
    if (!vigente(turno)) { const e = new Error('vista obsoleta'); e.obsoleta = true; throw e; }
    $('#vista-mapa').hidden = true;
    const v = $('#vista');
    v.hidden = false;
    const cont = document.createElement('div');
    cont.innerHTML = html;
    v.replaceChildren(cont);
    return cont;
  }

  function exigirAporte() {
    if (!uid()) {
      sessionStorage.setItem('dr-volver', location.hash);
      mostrar(`<div class="hoja"><h1>Ingresa para aportar</h1>
        <p>Para registrar o editar sitios necesitas una cuenta aprobada. El catálogo se puede consultar sin cuenta.</p>
        <p class="fila-botones"><a class="boton" href="#cuenta">Ingresar o crear cuenta</a><a class="boton secundario" href="#mapa">Volver al mapa</a></p></div>`);
      return false;
    }
    if (!puede.aportar()) {
      mostrar(`<div class="hoja"><h1>Tu cuenta está pendiente</h1>
        <div class="mensaje aviso"><p>Un administrador debe aprobar tu cuenta antes de que puedas registrar sitios. Mientras tanto, completa tu perfil para que sepan quién eres.</p></div>
        <p class="fila-botones"><a class="boton" href="#cuenta">Completar mi perfil</a><a class="boton secundario" href="#mapa">Volver al mapa</a></p></div>`);
      return false;
    }
    return true;
  }

  // ===================================================================
  // Mapa público
  // ===================================================================
  let mapa = null, capaGrupos = null, encuadrado = false;
  const marcadores = new Map();

  function svgAnillos(publica, tam) {
    const centro = publica ? '#7F4A1F' : '#E7C98E';
    return `<svg width="${tam}" height="${tam}" viewBox="0 0 26 26" aria-hidden="true">
      <circle cx="13" cy="13" r="12" fill="#5B3A1E"/>
      <circle cx="13" cy="13" r="10.4" fill="#E7C98E"/>
      <circle cx="13" cy="13" r="8.2" fill="none" stroke="#9A6233" stroke-width="1.1"/>
      <circle cx="13" cy="13" r="5.8" fill="none" stroke="#9A6233" stroke-width="1.1"/>
      <circle cx="13" cy="13" r="3.4" fill="${centro}" stroke="#9A6233" stroke-width="1.1"/>
    </svg>`;
  }
  const iconoSitio = (s) => L.divIcon({ className: 'marcador-anillos', html: svgAnillos(s.hay_serie_publica, 26), iconSize: [26, 26], iconAnchor: [13, 13], popupAnchor: [0, -12] });

  function capasBase() {
    return {
      'Mapa': L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '© colaboradores de OpenStreetMap' }),
      'Relieve': L.tileLayer('https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png', { maxZoom: 17, attribution: '© OpenStreetMap, SRTM · estilo © OpenTopoMap (CC BY-SA)' }),
      'Satélite': L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', { maxZoom: 19, attribution: 'Imágenes © Esri, Maxar, Earthstar Geographics' })
    };
  }

  async function vistaMapa(turno) {
    $('#vista').hidden = true;
    const cont = $('#vista-mapa');
    cont.hidden = false;
    if (!mapa) {
      cont.innerHTML = `
        <div class="explorar">
          <aside class="panel" aria-label="Filtros y lista de sitios">
            <div class="panel-cabeza">
              <h1>Sitios y cronologías</h1>
              <p class="resumen num" id="resumen">Cargando catálogo…</p>
            </div>
            <form class="filtros" id="filtros" role="search">
              <div>
                <label for="f-texto">Buscar</label>
                <input type="search" id="f-texto" placeholder="Código, nombre, comuna o laboratorio">
              </div>
              <div class="par">
                <div><label for="f-especie">Especie</label><select id="f-especie"></select></div>
                <div><label for="f-region">Región</label><select id="f-region"></select></div>
              </div>
              <details class="filtros-mas">
                <summary>Más filtros</summary>
                <div>
                  <div><label for="f-lab">Laboratorio</label><select id="f-lab"></select></div>
                  <div>
                    <label>Que cubra el período</label>
                    <div class="par">
                      <input type="number" id="f-desde" placeholder="Desde" aria-label="Desde el año">
                      <input type="number" id="f-hasta" placeholder="Hasta" aria-label="Hasta el año">
                    </div>
                  </div>
                  <label class="marca-check"><input type="checkbox" id="f-publica"> Solo con serie pública (ITRDB u otro repositorio)</label>
                </div>
              </details>
              <div><button type="button" class="boton-texto chico" id="f-limpiar">Quitar filtros</button></div>
            </form>
            <div class="barra-lista">
              <div class="segmentado" role="group" aria-label="Forma de ver">
                <button type="button" data-modo="mapa" aria-pressed="true">Mapa</button>
                <button type="button" data-modo="tabla" aria-pressed="false">Tabla</button>
              </div>
              <div class="menu-exportar">
                <button type="button" class="boton secundario chico" id="b-exportar" aria-expanded="false" aria-haspopup="true">Exportar</button>
                <ul id="m-exportar" hidden>
                  <li><button type="button" data-exp="xlsx">Excel<span>Hojas de sitios y de cronologías</span></button></li>
                  <li><button type="button" data-exp="csv">CSV<span>Una fila por cronología, para R o Python</span></button></li>
                  <li><button type="button" data-exp="geojson">GeoJSON<span>Puntos para QGIS u otro SIG</span></button></li>
                </ul>
              </div>
            </div>
            <ol class="lista-sitios" id="lista"></ol>
          </aside>
          <div class="mapa-zona">
            <div id="mapa-principal" aria-label="Mapa de sitios"></div>
            <div class="leyenda" id="leyenda">
              <div>${svgAnillos(true, 18)} Con serie pública</div>
              <div>${svgAnillos(false, 18)} Serie a pedido o no disponible</div>
            </div>
            <div class="zona-tabla" id="zona-tabla" hidden></div>
          </div>
        </div>`;

      const capas = capasBase();
      mapa = L.map('mapa-principal', { layers: [capas['Mapa']], zoomControl: true })
        .setView(CFG.centroMapa || [-39.5, -72], CFG.zoomMapa || 5);
      L.control.layers(capas, null, { position: 'topright' }).addTo(mapa);
      L.control.scale({ imperial: false }).addTo(mapa);
      capaGrupos = L.markerClusterGroup({
        showCoverageOnHover: false, maxClusterRadius: 45,
        iconCreateFunction: (g) => {
          const n = g.getChildCount();
          const t = n < 10 ? 34 : n < 100 ? 40 : 46;
          return L.divIcon({ className: 'grupo-anillos', html: `<div>${n}</div>`, iconSize: [t, t] });
        }
      }).addTo(mapa);

      // Filtros
      const f = estado.filtros;
      const enlazar = (id, clave, ev = 'input', leer = (e) => e.target.value) =>
        $('#' + id).addEventListener(ev, (e) => { f[clave] = leer(e); aplicarFiltros(); });
      enlazar('f-texto', 'texto');
      enlazar('f-especie', 'especie', 'change');
      enlazar('f-region', 'region', 'change');
      enlazar('f-lab', 'laboratorio', 'change');
      enlazar('f-desde', 'desde');
      enlazar('f-hasta', 'hasta');
      enlazar('f-publica', 'publica', 'change', (e) => e.target.checked);
      $('#f-limpiar').addEventListener('click', () => {
        Object.assign(f, { texto: '', especie: '', region: '', laboratorio: '', desde: '', hasta: '', publica: false });
        $('#filtros').reset();
        aplicarFiltros();
      });
      $$('.segmentado button', cont).forEach((b) => b.addEventListener('click', () => cambiarModo(b.dataset.modo)));
      const bExp = $('#b-exportar'), mExp = $('#m-exportar');
      bExp.addEventListener('click', () => { mExp.hidden = !mExp.hidden; bExp.setAttribute('aria-expanded', String(!mExp.hidden)); });
      document.addEventListener('click', (e) => { if (!e.target.closest('.menu-exportar')) { mExp.hidden = true; bExp.setAttribute('aria-expanded', 'false'); } });
      $$('[data-exp]', mExp).forEach((b) => b.addEventListener('click', () => { mExp.hidden = true; exportar(b.dataset.exp); }));
    } else {
      setTimeout(() => mapa.invalidateSize(), 0);
    }

    try {
      await cargarBase();
    } catch (e) {
      $('#resumen').textContent = mensajeError(e);
      return;
    }
    if (!vigente(turno)) return;
    pintarOpcionesFiltro();
    pintarMarcadores();
    aplicarFiltros();
    // La primera vez, el mapa se encuadra sobre todos los sitios
    if (!encuadrado && marcadores.size) {
      encuadrado = true;
      const limites = L.latLngBounds([...marcadores.values()].map((m) => m.getLatLng()));
      mapa.fitBounds(limites, { padding: [50, 50], maxZoom: 9 });
    }
  }

  function pintarOpcionesFiltro() {
    const usadas = new Set();
    estado.catalogo.forEach((s) => (s.especies || '').split('; ').filter(Boolean).forEach((e) => usadas.add(e)));
    const especies = estado.especies.filter((e) => usadas.has(e.nombre_cientifico));
    const f = estado.filtros;
    $('#f-especie').innerHTML = '<option value="">Todas</option>' +
      especies.map((e) => `<option value="${esc(e.nombre_cientifico)}"${f.especie === e.nombre_cientifico ? ' selected' : ''}>${esc(e.nombre_cientifico)}${e.nombre_comun ? ' (' + esc(e.nombre_comun.split(',')[0]) + ')' : ''}</option>`).join('');
    const regiones = [...new Set(estado.catalogo.map((s) => s.region).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
    $('#f-region').innerHTML = '<option value="">Todas</option>' + regiones.map((r) => `<option${f.region === r ? ' selected' : ''}>${esc(r)}</option>`).join('');
    const labs = [...new Set(estado.catalogo.flatMap((s) => (s.laboratorios || '').split('; ').filter(Boolean)))].sort((a, b) => a.localeCompare(b, 'es'));
    $('#f-lab').innerHTML = '<option value="">Todos</option>' + labs.map((r) => `<option${f.laboratorio === r ? ' selected' : ''}>${esc(r)}</option>`).join('');
  }

  function pintarMarcadores() {
    capaGrupos.clearLayers();
    marcadores.clear();
    estado.catalogo.forEach((s) => {
      if (!isFinite(s.latitud) || !isFinite(s.longitud)) return;
      const m = L.marker([s.latitud, s.longitud], { icon: iconoSitio(s), title: s.codigo + ' ' + s.nombre, keyboard: true });
      m.bindPopup(() => contenidoPopup(s));
      marcadores.set(s.id, m);
    });
  }

  function contenidoPopup(s) {
    return `<h3>${esc(s.nombre)}</h3>
      <p class="suave">${esc(s.codigo)}${s.region ? ', ' + esc(s.region) : ''}</p>
      <p>${s.n_cronologias ? '<em>' + esc(s.especies) + '</em>' : 'Sin cronologías registradas'}</p>
      ${s.n_cronologias ? `<p class="num">${plural(Number(s.n_cronologias), 'cronología', 'cronologías')}, ${periodo(s.anio_inicio, s.anio_fin)}</p>` : ''}
      <a class="boton chico" href="#sitio/${s.id}">Ver ficha</a>`;
  }

  function filtrados() {
    const f = estado.filtros;
    const texto = I.normalizar(f.texto);
    const desde = f.desde === '' ? null : Number(f.desde);
    const hasta = f.hasta === '' ? null : Number(f.hasta);
    return estado.catalogo.filter((s) => {
      if (texto && !I.normalizar([s.codigo, s.nombre, s.comuna, s.region, s.laboratorios, s.especies, s.registrado_por].join(' ')).includes(texto)) return false;
      if (f.especie && !(s.especies || '').split('; ').includes(f.especie)) return false;
      if (f.region && s.region !== f.region) return false;
      // Coincide tanto con el laboratorio del sitio como con el de cualquiera de sus cronologías.
      if (f.laboratorio && !(s.laboratorios || '').split('; ').includes(f.laboratorio)) return false;
      if (desde !== null && !(s.anio_inicio !== null && s.anio_inicio <= desde)) return false;
      if (hasta !== null && !(s.anio_fin !== null && s.anio_fin >= hasta)) return false;
      if (f.publica && !s.hay_serie_publica) return false;
      return true;
    });
  }

  function aplicarFiltros() {
    const lista = filtrados();
    const nCrono = lista.reduce((a, s) => a + Number(s.n_cronologias || 0), 0);
    const total = estado.catalogo.length;
    $('#resumen').textContent = (lista.length === total ? plural(total, 'sitio', 'sitios') : lista.length + ' de ' + total + ' sitios')
      + ', ' + plural(nCrono, 'cronología', 'cronologías');

    capaGrupos.clearLayers();
    capaGrupos.addLayers(lista.map((s) => marcadores.get(s.id)).filter(Boolean));

    const ol = $('#lista');
    if (!lista.length) {
      ol.innerHTML = total
        ? '<li class="lista-vacia">Ningún sitio coincide con los filtros.</li>'
        : '<li class="lista-vacia">Aún no hay sitios en el catálogo.' + (puede.aportar() ? ' <a href="#sitio/nuevo">Registra el primero</a>.' : '') + '</li>';
    } else {
      ol.innerHTML = lista.map((s) => `<li><button type="button" data-id="${s.id}">
          ${svgAnillos(s.hay_serie_publica, 20)}
          <span><span class="codigo">${esc(s.codigo)}</span> ${esc(s.nombre)}
            <span class="detalle">${s.n_cronologias ? '<em>' + esc(s.especies) + '</em>' : 'Sin cronologías'}${s.region ? ', ' + esc(s.region) : ''}</span></span>
          <span class="periodo num">${s.n_cronologias ? periodo(s.anio_inicio, s.anio_fin) : ''}</span>
        </button></li>`).join('');
      $$('button', ol).forEach((b) => b.addEventListener('click', () => enfocarSitio(+b.dataset.id)));
    }
    if (estado.modo === 'tabla') pintarTabla(lista);
  }

  function enfocarSitio(id) {
    const m = marcadores.get(id);
    if (!m) { location.hash = '#sitio/' + id; return; }
    if (estado.modo === 'tabla') cambiarModo('mapa');
    $$('#lista button').forEach((b) => b.classList.toggle('activo', +b.dataset.id === id));
    capaGrupos.zoomToShowLayer(m, () => m.openPopup());
    if (window.matchMedia('(max-width: 760px)').matches) $('#mapa-principal').scrollIntoView({ behavior: 'smooth' });
  }

  function cambiarModo(modo) {
    estado.modo = modo;
    $$('.segmentado button').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.modo === modo)));
    $('#zona-tabla').hidden = modo !== 'tabla';
    $('.explorar').classList.toggle('modo-tabla', modo === 'tabla');
    if (modo === 'tabla') pintarTabla(filtrados());
    else setTimeout(() => mapa.invalidateSize(), 0);
  }

  const COLUMNAS_TABLA = [
    ['codigo', 'Código'], ['nombre', 'Nombre'], ['region', 'Región'], ['especies', 'Especies'],
    ['n_cronologias', 'Crono.', 'num'], ['anio_inicio', 'Desde', 'num'], ['anio_fin', 'Hasta', 'num'],
    ['latitud', 'Latitud', 'num'], ['longitud', 'Longitud', 'num'], ['altitud_m', 'Altitud', 'num'],
    ['laboratorios', 'Laboratorios'], ['registrado_por', 'Registrado por']
  ];
  function pintarTabla(lista) {
    const o = estado.orden;
    const filas = ordenarFilas(lista.map((s) => Object.assign({}, s, { n_cronologias: Number(s.n_cronologias) })), o.campo, o.asc);
    const z = $('#zona-tabla');
    z.innerHTML = `<div class="tabla-envoltura"><table>
      <thead><tr>${COLUMNAS_TABLA.map(([c, t, cl]) => `<th class="${cl || ''}" aria-sort="${o.campo === c ? (o.asc ? 'ascending' : 'descending') : 'none'}"><button type="button" data-campo="${c}">${t}</button></th>`).join('')}</tr></thead>
      <tbody>${filas.map((s) => `<tr class="clicable" data-id="${s.id}" tabindex="0">
        <td><strong>${esc(s.codigo)}</strong></td><td>${esc(s.nombre)}</td><td>${esc(s.region)}</td><td><em>${esc(s.especies)}</em></td>
        <td class="num">${s.n_cronologias}</td><td class="num">${anio(s.anio_inicio)}</td><td class="num">${anio(s.anio_fin)}</td>
        <td class="num">${nf(s.latitud, 4)}</td><td class="num">${nf(s.longitud, 4)}</td><td class="num">${s.altitud_m === null ? '—' : nf(s.altitud_m)}</td>
        <td>${esc(s.laboratorios)}</td><td>${esc(s.registrado_por)}</td></tr>`).join('')}</tbody></table></div>`;
    $$('th button', z).forEach((b) => b.addEventListener('click', () => {
      o.asc = o.campo === b.dataset.campo ? !o.asc : true;
      o.campo = b.dataset.campo;
      pintarTabla(filtrados());
    }));
    $$('tbody tr', z).forEach((tr) => {
      const ir = () => { location.hash = '#sitio/' + tr.dataset.id; };
      tr.addEventListener('click', ir);
      tr.addEventListener('keydown', (e) => { if (e.key === 'Enter') ir(); });
    });
  }

  // ---------- Exportación ----------
  async function exportar(formato) {
    const sitios = filtrados();
    if (!sitios.length) { aviso('No hay sitios para exportar con estos filtros', 'error'); return; }
    aviso('Preparando la exportación…');
    try {
      const ids = new Set(sitios.map((s) => s.id));
      const cronos = (await traerTodo(() => sb.from('cronologias')
        .select('*, especies(nombre_cientifico, codigo_itrdb), laboratorios(nombre), cronologia_autores(orden, nombre_autor, orcid, es_responsable, es_autor, perfiles(nombre_completo, orcid))')
        .order('sitio_id').order('codigo'))).filter((c) => ids.has(c.sitio_id));
      const detalle = await traerTodo(() => sb.from('sitios').select('id, comuna, localidad, coord_original, datum_original, ambiente, descripcion, precision'));
      const det = new Map(detalle.map((d) => [d.id, d]));
      const hoy = new Date().toISOString().slice(0, 10);

      const filaSitio = (s) => {
        const d = det.get(s.id) || {};
        return {
          sitio_codigo: s.codigo, sitio_nombre: s.nombre, pais: s.pais, region: s.region, comuna: d.comuna || null, localidad: d.localidad || null,
          latitud_wgs84: s.latitud, longitud_wgs84: s.longitud, altitud_m: s.altitud_m, precision_coordenadas: s.precision,
          coord_original: d.coord_original || null, datum_original: d.datum_original || null, ambiente: d.ambiente || null,
          laboratorio: s.laboratorio, registrado_por: s.registrado_por
        };
      };
      // Dos columnas separadas, igual que en la ficha: quien responde por la
      // cronología y quienes la cofecharon. Puede ser la misma persona.
      const nombresDe = (c, filtro) => (c.cronologia_autores || []).sort((a, b) => a.orden - b.orden)
        .filter(filtro).map((a) => (a.perfiles ? a.perfiles.nombre_completo : a.nombre_autor) || '')
        .filter(Boolean).join('; ');
      const autoresTexto = (c) => nombresDe(c, (a) => a.es_autor);
      const responsableTexto = (c) => nombresDe(c, (a) => a.es_responsable);
      const filaCrono = (c) => ({
        cronologia_codigo: c.codigo, especie: c.especies ? c.especies.nombre_cientifico : null, codigo_itrdb: c.especies ? c.especies.codigo_itrdb : null,
        parametro: ETIQ.parametro[c.parametro], estado: ETIQ.estado[c.estado], anio_inicio: c.anio_inicio, anio_fin: c.anio_fin,
        anio_eps_085: c.anio_eps_085, anio_muestreo: c.anio_muestreo, n_arboles: c.n_arboles, n_radios: c.n_radios,
        intercorrelacion: c.intercorrelacion === null ? null : Number(c.intercorrelacion), sensibilidad_media: c.sensibilidad_media === null ? null : Number(c.sensibilidad_media),
        rbar: c.rbar === null ? null : Number(c.rbar), eps: c.eps === null ? null : Number(c.eps),
        estandarizacion: c.estandarizacion, disponibilidad_serie: ETIQ.disponibilidadCorta[c.disponibilidad], enlace_serie: c.enlace_serie,
        doi: c.doi, cita: c.cita,
        laboratorio_cronologia: c.laboratorios ? c.laboratorios.nombre : null,
        responsable: responsableTexto(c), autores: autoresTexto(c)
      });

      if (formato === 'geojson') {
        const fc = {
          type: 'FeatureCollection',
          features: sitios.map((s) => ({
            type: 'Feature',
            geometry: { type: 'Point', coordinates: [s.longitud, s.latitud].concat(s.altitud_m !== null ? [s.altitud_m] : []) },
            properties: Object.assign(filaSitio(s), {
              especies: s.especies, n_cronologias: Number(s.n_cronologias), anio_inicio: s.anio_inicio, anio_fin: s.anio_fin,
              ficha: location.origin + location.pathname + '#sitio/' + s.id
            })
          }))
        };
        descargar('dendrored_sitios_' + hoy + '.geojson', JSON.stringify(fc, null, 1), 'application/geo+json');
        return;
      }

      const porSitio = new Map();
      cronos.forEach((c) => { if (!porSitio.has(c.sitio_id)) porSitio.set(c.sitio_id, []); porSitio.get(c.sitio_id).push(c); });
      const planas = [];
      sitios.forEach((s) => {
        const cs = porSitio.get(s.id) || [];
        if (!cs.length) planas.push(Object.assign(filaSitio(s), filaCrono({})));
        cs.forEach((c) => planas.push(Object.assign(filaSitio(s), filaCrono(c))));
      });
      // Una fila de sitio sin cronología queda con columnas vacías
      planas.forEach((p) => Object.keys(p).forEach((k) => { if (p[k] === undefined) p[k] = null; }));

      if (formato === 'csv') {
        const cols = Object.keys(planas[0]);
        const celda = (v) => {
          if (v === null || v === undefined) return '';
          const t = String(v);
          return /[",\n\r]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
        };
        const csv = '﻿' + [cols.join(',')].concat(planas.map((p) => cols.map((c) => celda(p[c])).join(','))).join('\r\n');
        descargar('dendrored_' + hoy + '.csv', csv, 'text/csv;charset=utf-8');
        return;
      }

      await cargarScript(URL_XLSX);
      const libro = XLSX.utils.book_new();
      const hojaSitios = sitios.map((s) => Object.assign(filaSitio(s), {
        especies: s.especies, n_cronologias: Number(s.n_cronologias), anio_inicio: s.anio_inicio, anio_fin: s.anio_fin
      }));
      XLSX.utils.book_append_sheet(libro, XLSX.utils.json_to_sheet(hojaSitios), 'Sitios');
      const hojaCronos = cronos.map((c) => {
        const s = sitios.find((x) => x.id === c.sitio_id);
        return Object.assign({ sitio_codigo: s.codigo, sitio_nombre: s.nombre }, filaCrono(c));
      });
      XLSX.utils.book_append_sheet(libro, XLSX.utils.json_to_sheet(hojaCronos.length ? hojaCronos : [{ aviso: 'Sin cronologías en la selección' }]), 'Cronologías');
      const fuente = [{ fuente: 'DendroRed, catálogo colaborativo de sitios y cronologías dendrocronológicas', fecha_descarga: hoy, enlace: location.origin + location.pathname }];
      XLSX.utils.book_append_sheet(libro, XLSX.utils.json_to_sheet(fuente), 'Fuente');
      const datos = XLSX.write(libro, { bookType: 'xlsx', type: 'array' });
      descargar('dendrored_' + hoy + '.xlsx', new Blob([datos], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    } catch (e) {
      console.error(e);
      aviso('No se pudo exportar: ' + mensajeError(e), 'error');
    }
  }

  // ===================================================================
  // Ficha de un sitio
  // ===================================================================
  async function vistaFicha(id, turno) {
    const v = mostrar('<div class="hoja-ancha"><p class="suave">Cargando sitio…</p></div>', turno);
    const { data: s, error } = await sb.from('sitios')
      .select('*, laboratorios(nombre, institucion), perfiles(nombre_completo, orcid)')
      .eq('id', id).maybeSingle();
    if (error) throw error;
    if (!s) {
      v.innerHTML = '<div class="hoja"><h1>Sitio no encontrado</h1><p>Puede que lo hayan borrado. <a href="#mapa">Volver al mapa</a></p></div>';
      return;
    }
    const { data: cronos, error: e2 } = await sb.from('cronologias')
      .select('*, especies(nombre_cientifico, nombre_comun, codigo_itrdb), perfiles(nombre_completo), laboratorios(nombre), cronologia_autores(orden, nombre_autor, orcid, es_responsable, es_autor, perfil_id, perfiles(nombre_completo, orcid))')
      .eq('sitio_id', id).order('codigo');
    if (e2) throw e2;
    if (!vigente(turno)) return;

    const lat = Number(s.latitud), lon = Number(s.longitud);
    const datos = [
      ['Ubicación', [s.localidad, s.comuna, s.region, s.pais].filter(Boolean).map(esc).join(', ') || '—'],
      ['Coordenadas (WGS84)', `<span class="num">${lat.toFixed(6)}, ${lon.toFixed(6)}</span><br><span class="suave chico">${C.aGMS(lat, 'lat')} ${C.aGMS(lon, 'lon')}</span>`
        + (s.precision === 'aproximada' ? '<br><span class="etiqueta duramen">Aproximadas, ~1 km</span>' : '')],
      ['Altitud', s.altitud_m === null ? '—' : nf(s.altitud_m) + ' m'],
      ['Coordenadas originales', s.coord_original ? esc(s.coord_original) + ' <span class="suave">(' + esc(s.datum_original) + ')</span>' : '—'],
      ['Ambiente', esc(s.ambiente) || '—'],
      ['Laboratorio', s.laboratorios ? esc(s.laboratorios.nombre) + (s.laboratorios.institucion ? ', ' + esc(s.laboratorios.institucion) : '') : '—'],
      ['Registrado por', s.perfiles ? esc(s.perfiles.nombre_completo) : '—'],
      ['Última modificación', fecha(s.actualizado_en)]
    ];

    v.innerHTML = `<div class="hoja-ancha">
      <a class="volver" href="#mapa">Volver al mapa</a>
      <div class="ficha-cabeza">
        <div>
          <p class="ficha-codigo">${esc(s.codigo)}</p>
          <h1>${esc(s.nombre)}</h1>
          ${s.descripcion ? `<p>${esc(s.descripcion)}</p>` : ''}
          <dl class="datos">${datos.map(([t, d]) => `<dt>${t}</dt><dd>${d}</dd>`).join('')}</dl>
          <p class="fila-botones" style="margin-top:1rem">
            ${puede.aportar() ? `<a class="boton" href="#sitio/${s.id}/cronologia">Agregar cronología</a>` : ''}
            ${puede.editar(s) ? `<a class="boton secundario" href="#sitio/${s.id}/editar">Editar sitio</a><button type="button" class="boton peligro" id="b-borrar-sitio">Borrar sitio</button>` : ''}
          </p>
        </div>
        <div id="mapa-ficha" class="mapa-chico" aria-label="Ubicación del sitio"></div>
      </div>
      <section class="caja">
        <h2>${cronos.length ? plural(cronos.length, 'cronología', 'cronologías') : 'Cronologías'}</h2>
        ${cronos.length ? cronos.map(htmlCronologia).join('') : `<p class="suave">Este sitio aún no tiene cronologías registradas.${puede.aportar() ? ` <a href="#sitio/${s.id}/cronologia">Agrega la tuya</a>.` : ''}</p>`}
      </section>
    </div>`;

    const mf = L.map($('#mapa-ficha', v), { scrollWheelZoom: false }).setView([lat, lon], s.precision === 'aproximada' ? 11 : 12);
    const capas = capasBase();
    capas['Relieve'].addTo(mf);
    L.control.layers(capas).addTo(mf);
    L.marker([lat, lon], { icon: iconoSitio({ hay_serie_publica: cronos.some((c) => ['itrdb', 'repositorio'].includes(c.disponibilidad)) }) }).addTo(mf);
    if (s.precision === 'aproximada') L.circle([lat, lon], { radius: 1000, color: '#7F4A1F', weight: 1, fillOpacity: .08 }).addTo(mf);

    const bBorrar = $('#b-borrar-sitio');
    if (bBorrar) bBorrar.addEventListener('click', async () => {
      const otras = cronos.filter((c) => c.creado_por !== uid()).length;
      const texto = '¿Borrar el sitio ' + s.codigo + '?' + (cronos.length ? ' También se borrarán sus ' + plural(cronos.length, 'cronología', 'cronologías') + (otras ? ', incluidas ' + otras + ' de otras personas' : '') + '.' : '') + ' Esta acción no se puede deshacer.';
      if (!confirm(texto)) return;
      const { error, count } = await sb.from('sitios').delete({ count: 'exact' }).eq('id', s.id);
      if (error) { aviso(mensajeError(error), 'error'); return; }
      if (!count) { aviso('No se borró: no tienes permiso sobre este sitio', 'error'); return; }
      aviso('Sitio ' + s.codigo + ' borrado');
      estado.cargado = false;
      location.hash = '#mapa';
    });
    $$('[data-borrar-crono]').forEach((b) => b.addEventListener('click', async () => {
      const c = cronos.find((x) => x.id === +b.dataset.borrarCrono);
      if (!confirm('¿Borrar la cronología ' + c.codigo + '? Esta acción no se puede deshacer.')) return;
      const { error, count } = await sb.from('cronologias').delete({ count: 'exact' }).eq('id', c.id);
      if (error || !count) { aviso(error ? mensajeError(error) : 'No tienes permiso para borrarla', 'error'); return; }
      aviso('Cronología borrada');
      estado.cargado = false;
      vistaFicha(id, ++turnoVista);
    }));
    $$('[data-contacto]').forEach((b) => b.addEventListener('click', async () => {
      const { data, error } = await sb.rpc('correo_de', { p_perfil: b.dataset.contacto });
      if (error || !data) { aviso('No se pudo obtener el correo', 'error'); return; }
      const a = document.createElement('a');
      a.href = 'mailto:' + data; a.textContent = data;
      b.replaceWith(a);
    }));
  }

  function htmlCronologia(c) {
    const e = c.especies || {};
    // Responsable y autor son cosas distintas: el responsable responde por la
    // cronología, el autor la cofechó. Se muestran en filas separadas, y quien
    // sea las dos cosas aparece en las dos.
    const personas = (c.cronologia_autores || []).sort((a, b) => a.orden - b.orden);
    const comoTexto = (a) => {
      const nombre = a.perfiles ? a.perfiles.nombre_completo : a.nombre_autor;
      const orcid = (a.perfiles && a.perfiles.orcid) || a.orcid;
      return esc(nombre) + (orcid ? ` <a href="https://orcid.org/${esc(orcid)}" target="_blank" rel="noopener" title="ORCID ${esc(orcid)}" class="chico">ORCID</a>` : '');
    };
    const responsables = personas.filter((a) => a.es_responsable).map(comoTexto);
    const autores = personas.filter((a) => a.es_autor).map(comoTexto);
    const est = [
      ['Período', periodo(c.anio_inicio, c.anio_fin)], ['EPS ≥ 0,85 desde', anio(c.anio_eps_085)],
      ['Árboles', nf(c.n_arboles)], ['Radios', nf(c.n_radios)],
      ['Intercorrelación', nf(c.intercorrelacion, 3)], ['Sensibilidad media', nf(c.sensibilidad_media, 3)],
      ['rbar', nf(c.rbar, 3)], ['EPS', nf(c.eps, 3)]
    ].filter(([, v]) => v !== '—');
    const enlace = urlSegura(c.enlace_serie);
    const doi = enlaceDOI(c.doi);
    const responsable = personas.find((a) => a.es_responsable && a.perfil_id);
    const idContacto = responsable ? responsable.perfil_id : c.creado_por;
    let serie = '<span class="etiqueta ' + (['itrdb', 'repositorio'].includes(c.disponibilidad) ? '' : 'gris') + '">' + ETIQ.disponibilidad[c.disponibilidad] + '</span>';
    if (enlace) serie += ` <a href="${esc(enlace)}" target="_blank" rel="noopener">Abrir la serie</a>`;
    if (c.disponibilidad === 'a_pedido' && idContacto) {
      serie += puede.aportar()
        ? ` <button type="button" class="boton-texto" data-contacto="${esc(idContacto)}">Ver correo de contacto</button>`
        : ' <span class="suave chico">(ingresa con tu cuenta para ver el contacto)</span>';
    }
    return `<article class="cronologia">
      <div class="cronologia-cabeza">
        <h3><em>${esc(e.nombre_cientifico)}</em>${e.nombre_comun ? ' <span class="suave">(' + esc(e.nombre_comun.split(',')[0]) + ')</span>' : ''}, ${ETIQ.parametro[c.parametro].toLowerCase()}</h3>
        <span class="chico"><span class="suave">${esc(c.codigo)}</span> <span class="etiqueta gris">${ETIQ.estado[c.estado]}</span></span>
      </div>
      ${est.length ? `<div class="estadisticos">${est.map(([t, v]) => `<div><span>${t}</span><strong>${v}</strong></div>`).join('')}</div>` : ''}
      <dl class="datos">
        ${responsables.length ? `<dt>${responsables.length === 1 ? 'Responsable' : 'Responsables'}</dt><dd>${responsables.join(', ')}</dd>` : ''}
        ${autores.length ? `<dt>${autores.length === 1 ? 'Autor' : 'Autores'}</dt><dd>${autores.join(', ')}</dd>` : ''}
        ${c.laboratorios ? `<dt>Laboratorio</dt><dd>${esc(c.laboratorios.nombre)}</dd>` : ''}
        ${c.estandarizacion ? `<dt>Estandarización</dt><dd>${esc(c.estandarizacion)}</dd>` : ''}
        ${c.anio_muestreo ? `<dt>Muestreo</dt><dd>${c.anio_muestreo}</dd>` : ''}
        <dt>Serie de tiempo</dt><dd>${serie}</dd>
        ${doi ? `<dt>Publicación</dt><dd><a href="${esc(doi.url)}" target="_blank" rel="noopener">${esc(doi.texto)}</a>${c.cita ? '<br>' + esc(c.cita) : ''}</dd>` : c.cita ? `<dt>Publicación</dt><dd>${esc(c.cita)}</dd>` : ''}
        ${c.notas ? `<dt>Notas</dt><dd>${esc(c.notas)}</dd>` : ''}
      </dl>
      ${puede.editar(c) ? `<p class="fila-botones" style="margin:.75rem 0 0"><a class="boton secundario chico" href="#cronologia/${c.id}/editar">Editar</a><button type="button" class="boton peligro chico" data-borrar-crono="${c.id}">Borrar</button></p>` : ''}
    </article>`;
  }

  // ===================================================================
  // Formulario de sitio
  // ===================================================================
  async function vistaFormSitio(id, turno) {
    await cargarBase();
    let s = { codigo: '', nombre: '', pais: 'Chile', region: '', comuna: '', localidad: '', latitud: '', longitud: '', altitud_m: '',
      precision: 'exacta', datum_original: 'WGS84', ambiente: '', descripcion: '', laboratorio_id: estado.perfil.laboratorio_id };
    if (id) {
      const { data, error } = await sb.from('sitios').select('*').eq('id', id).maybeSingle();
      if (error) throw error;
      if (!data) { mostrar('<div class="hoja"><h1>Sitio no encontrado</h1></div>', turno); return; }
      if (!puede.editar(data)) { mostrar('<div class="hoja"><h1>No puedes editar este sitio</h1><p>Solo quien lo registró o un curador puede modificarlo. Sí puedes <a href="#sitio/' + id + '/cronologia">agregarle tu cronología</a>.</p></div>', turno); return; }
      s = data;
    }
    const v = mostrar(`<div class="hoja">
      <a class="volver" href="${id ? '#sitio/' + id : '#mapa'}">${id ? 'Volver a la ficha' : 'Volver al mapa'}</a>
      <div class="encabezado-vista"><div><h1>${id ? 'Editar sitio ' + esc(s.codigo) : 'Nuevo sitio'}</h1>
        <p class="suave">${id ? '' : 'Registra el lugar de muestreo. Después podrás agregarle una o más cronologías.'}</p></div></div>
      <form id="form-sitio" class="caja" novalidate>
        <fieldset>
          <legend>Identificación</legend>
          <div class="rejilla">
            <div><label for="s-codigo" class="requerido">Código</label>
              <input type="text" id="s-codigo" required maxlength="16" value="${esc(s.codigo)}" autocomplete="off" spellcheck="false">
              <span class="ayuda">Corto y único, en mayúsculas. Ej.: PUY01, CONGUI</span></div>
            <div><label for="s-nombre" class="requerido">Nombre</label>
              <input type="text" id="s-nombre" required value="${esc(s.nombre)}"></div>
            <div><label for="s-region">Región o provincia</label>
              <input type="text" id="s-region" list="l-regiones" value="${esc(s.region)}">
              <datalist id="l-regiones">${REGIONES.map((r) => `<option value="${esc(r)}">`).join('')}</datalist></div>
            <div><label for="s-comuna">Comuna</label><input type="text" id="s-comuna" value="${esc(s.comuna)}"></div>
            <div><label for="s-localidad">Localidad o sector</label><input type="text" id="s-localidad" value="${esc(s.localidad)}"></div>
            <div><label for="s-pais">País</label><input type="text" id="s-pais" value="${esc(s.pais)}"></div>
          </div>
        </fieldset>

        <fieldset>
          <legend>Ubicación</legend>
          <div class="coord-bloque">
            <div>
              <div class="fila-botones" style="margin-bottom:.75rem">
                <div class="segmentado" role="group" aria-label="Formato de coordenadas">
                  <button type="button" data-fmt="geo" aria-pressed="true">Grados</button>
                  <button type="button" data-fmt="utm" aria-pressed="false">UTM</button>
                </div>
              </div>
              <div id="c-geo" class="rejilla">
                <div><label for="s-lat" class="requerido">Latitud</label><input type="text" id="s-lat" inputmode="text" value="${s.latitud === '' ? '' : Number(s.latitud)}" placeholder="-38.6567 o 38°39'24&quot;S"></div>
                <div><label for="s-lon" class="requerido">Longitud</label><input type="text" id="s-lon" value="${s.longitud === '' ? '' : Number(s.longitud)}" placeholder="-71.6489 o 71°38'56&quot;O"></div>
              </div>
              <div id="c-utm" class="rejilla rejilla-3" hidden>
                <div><label for="s-zona">Zona</label>
                  <select id="s-zona">${[17, 18, 19, 20].map((z) => `<option value="${z}"${z === 19 ? ' selected' : ''}>${z} S</option>`).join('')}</select>
                  <span class="ayuda">17 S: Juan Fernández</span></div>
                <div><label for="s-este">Este (m)</label><input type="text" id="s-este" inputmode="decimal"></div>
                <div><label for="s-norte">Norte (m)</label><input type="text" id="s-norte" inputmode="decimal"></div>
              </div>
              <div class="rejilla" style="margin-top:.9rem">
                <div><label for="s-datum">Datum de las coordenadas</label>
                  <select id="s-datum">${Object.entries(C.DATUMS).map(([k, d]) => `<option value="${k}"${k === s.datum_original ? ' selected' : ''}>${esc(d.etiqueta)}</option>`).join('')}</select>
                  <span class="ayuda">Si vienen de un GPS o Google Earth, es WGS84</span></div>
                <div><label for="s-alt">Altitud (m s. n. m.)</label><input type="number" id="s-alt" min="-500" max="7000" step="1" value="${esc(s.altitud_m)}"></div>
              </div>
              <div id="c-resultado" class="resultado-coord" aria-live="polite"></div>
              <div id="c-cercanos"></div>
              <label class="marca-check" style="margin-top:.75rem"><input type="checkbox" id="s-aprox"${s.precision === 'aproximada' ? ' checked' : ''}>
                <span>Publicar coordenadas aproximadas
                <span class="ayuda">Se redondean a dos decimales (~1 km). Recomendado para especies amenazadas o sitios vulnerables a la extracción.</span></span></label>
            </div>
            <div>
              <div id="mapa-form" class="mapa-form" aria-label="Mapa: haz clic para ubicar el sitio"></div>
              <span class="ayuda">Haz clic en el mapa o arrastra el marcador para ubicar el sitio.</span>
            </div>
          </div>
        </fieldset>

        <fieldset>
          <legend>Descripción</legend>
          <div class="rejilla">
            <div><label for="s-lab">Laboratorio</label>
              <select id="s-lab"><option value="">Sin laboratorio</option>${estado.laboratorios.map((l) => `<option value="${l.id}"${l.id === s.laboratorio_id ? ' selected' : ''}>${esc(l.nombre)}</option>`).join('')}</select></div>
            <div><label for="s-ambiente">Ambiente</label><input type="text" id="s-ambiente" value="${esc(s.ambiente)}" placeholder="Ej.: bosque de lenga en límite altitudinal, ladera sur"></div>
            <div class="ancho-total"><label for="s-desc">Descripción y acceso</label><textarea id="s-desc">${esc(s.descripcion)}</textarea></div>
          </div>
        </fieldset>
        <div id="s-errores"></div>
        <div class="fila-botones">
          <button type="submit" class="boton">${id ? 'Guardar cambios' : 'Guardar sitio'}</button>
          <a class="boton secundario" href="${id ? '#sitio/' + id : '#mapa'}">Cancelar</a>
        </div>
      </form></div>`, turno);

    // --- Mapa del formulario ---
    const tieneCoord = s.latitud !== '' && s.latitud !== null;
    const mf = L.map($('#mapa-form', v)).setView(tieneCoord ? [Number(s.latitud), Number(s.longitud)] : (CFG.centroMapa || [-39.5, -72]), tieneCoord ? 12 : 5);
    const capas = capasBase();
    capas['Relieve'].addTo(mf);
    L.control.layers(capas).addTo(mf);
    let marcador = null;
    const ponerMarcador = (lat, lon) => {
      if (!marcador) {
        marcador = L.marker([lat, lon], { draggable: true, icon: L.divIcon({ className: 'marcador-anillos', html: svgAnillos(false, 30), iconSize: [30, 30], iconAnchor: [15, 15] }) }).addTo(mf);
        marcador.on('dragend', () => { const p = marcador.getLatLng(); escribirDesdeMapa(p.lat, p.lng); });
      } else marcador.setLatLng([lat, lon]);
    };
    let formato = 'geo';
    const escribirDesdeMapa = (lat, lon) => {
      cambiarFormato('geo');
      $('#s-datum').value = 'WGS84';
      $('#s-lat').value = lat.toFixed(6);
      $('#s-lon').value = lon.toFixed(6);
      actualizar(false);
    };
    mf.on('click', (e) => escribirDesdeMapa(e.latlng.lat, e.latlng.lng));

    function cambiarFormato(f) {
      formato = f;
      $$('[data-fmt]', v).forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.fmt === f)));
      $('#c-geo').hidden = f !== 'geo';
      $('#c-utm').hidden = f !== 'utm';
    }
    $$('[data-fmt]', v).forEach((b) => b.addEventListener('click', () => { cambiarFormato(b.dataset.fmt); actualizar(true); }));

    // --- Lectura de coordenadas en vivo ---
    let wgs = null, temporizador = null;
    function leer() {
      const datum = $('#s-datum').value;
      if (formato === 'utm') {
        if (!$('#s-este').value.trim() || !$('#s-norte').value.trim()) return { vacio: true };
        const r = C.utmAWGS84($('#s-este').value, $('#s-norte').value, $('#s-zona').value, 'S', datum);
        return r.error ? { error: r.error } : Object.assign(r, { original: $('#s-zona').value + 'S ' + $('#s-este').value.trim() + ' ' + $('#s-norte').value.trim(), fmt: 'utm' });
      }
      const tLat = $('#s-lat').value, tLon = $('#s-lon').value;
      if (!tLat.trim() && !tLon.trim()) return { vacio: true };
      const a = C.leerCoordenada(tLat, 'lat'), b = C.leerCoordenada(tLon, 'lon');
      if (a.error || b.error) return { error: [a.error && 'latitud ' + a.error, b.error && 'longitud ' + b.error].filter(Boolean).join('; ') };
      const g = C.geograficasAWGS84(a.valor, b.valor, datum);
      return { lat: g.lat, lon: g.lon, original: tLat.trim() + ' ' + tLon.trim(), fmt: /[°'"NSEWO\s]/i.test(tLat.trim().replace(/^-/, '')) ? 'gms' : 'decimal' };
    }
    function actualizar(centrar) {
      const r = leer();
      const caja = $('#c-resultado');
      caja.classList.toggle('error', !!r.error);
      if (r.vacio) { wgs = null; caja.textContent = 'Escribe las coordenadas o haz clic en el mapa.'; return; }
      if (r.error) { wgs = null; caja.textContent = 'No se pudo leer: ' + r.error; return; }
      wgs = r;
      const aprox = $('#s-aprox').checked;
      const la = aprox ? Math.round(r.lat * 100) / 100 : r.lat, lo = aprox ? Math.round(r.lon * 100) / 100 : r.lon;
      const desfase = $('#s-datum').value !== 'WGS84' ? ' Corrección de datum aplicada.' : '';
      caja.innerHTML = `Se guardará en WGS84: <strong class="num">${la.toFixed(aprox ? 2 : 6)}, ${lo.toFixed(aprox ? 2 : 6)}</strong><br><span class="suave">${C.aGMS(r.lat, 'lat')} ${C.aGMS(r.lon, 'lon')}.${desfase}</span>`
        + (r.lat > 0 || r.lon > 0 ? '<br><strong>Revisa el signo:</strong> en Chile la latitud es S y la longitud O (valores negativos).' : '');
      ponerMarcador(r.lat, r.lon);
      if (centrar) mf.setView([r.lat, r.lon], Math.max(mf.getZoom(), 11));
      clearTimeout(temporizador);
      temporizador = setTimeout(() => revisarCercanos(r.lat, r.lon), 400);
    }
    async function revisarCercanos(lat, lon) {
      const { data } = await sb.rpc('sitios_cercanos', { p_latitud: lat, p_longitud: lon, p_radio_m: 1000 });
      const caja = $('#c-cercanos');
      if (!caja) return;   // la persona ya salió del formulario
      const otros = (data || []).filter((x) => x.id !== id);
      caja.innerHTML = otros.length
        ? `<div class="mensaje aviso" style="margin-top:.75rem"><p>Hay ${otros.length === 1 ? 'un sitio' : otros.length + ' sitios'} a menos de 1 km: ${otros.map((o) => `<a href="#sitio/${o.id}" target="_blank">${esc(o.codigo)}</a> (${Math.round(o.distancia_m)} m)`).join(', ')}.
            Si es el mismo lugar, mejor agrega tu cronología a ese sitio.</p></div>`
        : '';
    }
    ['s-lat', 's-lon', 's-este', 's-norte'].forEach((i) => $('#' + i).addEventListener('change', () => actualizar(true)));
    ['s-lat', 's-lon', 's-este', 's-norte'].forEach((i) => $('#' + i).addEventListener('input', () => actualizar(false)));
    ['s-datum', 's-zona', 's-aprox'].forEach((i) => $('#' + i).addEventListener('change', () => actualizar(true)));
    $('#s-codigo').addEventListener('input', (e) => {
      // Mientras se escribe: mayúsculas, sin tildes y espacios como _ (sin recortar el final)
      const p = e.target.selectionStart;
      e.target.value = e.target.value.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/\s/g, '_').replace(/[^A-Z0-9_-]/g, '');
      e.target.setSelectionRange(p, p);
    });
    actualizar(false);
    setTimeout(() => mf.invalidateSize(), 50);

    // --- Guardar ---
    $('#form-sitio').addEventListener('submit', async (e) => {
      e.preventDefault();
      const errores = [];
      const codigo = I.codigoValido($('#s-codigo').value);
      if (!/^[A-Z0-9_-]{2,16}$/.test(codigo)) errores.push('El código debe tener entre 2 y 16 caracteres (mayúsculas, números, _ o -).');
      if (!$('#s-nombre').value.trim()) errores.push('Falta el nombre del sitio.');
      if (!wgs) errores.push('Faltan las coordenadas o no se pudieron leer.');
      const alt = $('#s-alt').value;
      if (alt !== '' && (Number(alt) < -500 || Number(alt) > 7000)) errores.push('La altitud debe estar entre -500 y 7000 m.');
      $('#s-codigo').setAttribute('aria-invalid', String(errores.some((x) => /código/.test(x))));
      if (errores.length) {
        $('#s-errores').innerHTML = `<div class="mensaje error"><p>${errores.map(esc).join('<br>')}</p></div>`;
        return;
      }
      $('#s-errores').innerHTML = '';
      const aprox = $('#s-aprox').checked;
      const red = (x) => aprox ? Math.round(x * 100) / 100 : Math.round(x * 1e6) / 1e6;
      const fila = {
        codigo, nombre: $('#s-nombre').value.trim(), pais: vacioANulo($('#s-pais').value) || 'Chile',
        region: vacioANulo($('#s-region').value), comuna: vacioANulo($('#s-comuna').value), localidad: vacioANulo($('#s-localidad').value),
        latitud: red(wgs.lat), longitud: red(wgs.lon), altitud_m: alt === '' ? null : Math.round(Number(alt)),
        precision: aprox ? 'aproximada' : 'exacta',
        coord_original: aprox ? null : wgs.original, formato_original: wgs.fmt, datum_original: $('#s-datum').value,
        ambiente: vacioANulo($('#s-ambiente').value), descripcion: vacioANulo($('#s-desc').value),
        laboratorio_id: $('#s-lab').value ? Number($('#s-lab').value) : null
      };
      const boton = $('button[type=submit]', e.target);
      boton.disabled = true;
      const q = id ? sb.from('sitios').update(fila).eq('id', id).select('id') : sb.from('sitios').insert(fila).select('id');
      const { data, error } = await q;
      boton.disabled = false;
      if (error) { $('#s-errores').innerHTML = `<div class="mensaje error"><p>${esc(mensajeError(error))}</p></div>`; return; }
      if (!data || !data.length) { $('#s-errores').innerHTML = '<div class="mensaje error"><p>No se guardó: no tienes permiso sobre este sitio.</p></div>'; return; }
      estado.cargado = false;
      aviso(id ? 'Cambios guardados' : 'Sitio ' + codigo + ' guardado');
      location.hash = id ? '#sitio/' + id : '#sitio/' + data[0].id + '/cronologia';
    });
  }

  // ===================================================================
  // Formulario de cronología
  // ===================================================================
  async function vistaFormCrono(idCrono, idSitio, turno) {
    await cargarBase();
    let c = { codigo: '', especie_id: '', parametro: 'ancho_total', estado: 'terminada', disponibilidad: 'a_pedido', cronologia_autores: [] };
    if (idCrono) {
      const { data, error } = await sb.from('cronologias').select('*, cronologia_autores(*, perfiles(nombre_completo, orcid))').eq('id', idCrono).maybeSingle();
      if (error) throw error;
      if (!data) { mostrar('<div class="hoja"><h1>Cronología no encontrada</h1></div>', turno); return; }
      if (!puede.editar(data)) { mostrar('<div class="hoja"><h1>No puedes editar esta cronología</h1><p>Solo quien la registró o un curador puede modificarla.</p></div>', turno); return; }
      c = data;
      idSitio = c.sitio_id;
    }
    const { data: s, error: es } = await sb.from('sitios').select('id, codigo, nombre').eq('id', idSitio).maybeSingle();
    if (es) throw es;
    if (!s) { mostrar('<div class="hoja"><h1>Sitio no encontrado</h1></div>', turno); return; }
    const { data: perfiles } = await sb.from('perfiles').select('id, nombre_completo, orcid').order('nombre_completo');
    estado.perfilesPublicos = (perfiles || []).filter((p) => p.nombre_completo);

    let autores = (c.cronologia_autores || []).sort((a, b) => a.orden - b.orden).map((a) => ({
      perfil_id: a.perfil_id, nombre: a.perfiles ? a.perfiles.nombre_completo : a.nombre_autor,
      orcid: (a.perfiles && a.perfiles.orcid) || a.orcid || '', es_responsable: a.es_responsable, es_autor: a.es_autor
    }));
    // Quien crea la cronología a mano normalmente la hizo y responde por ella.
    if (!idCrono) autores = [{ perfil_id: uid(), nombre: estado.perfil.nombre_completo || '', orcid: estado.perfil.orcid || '', es_responsable: true, es_autor: true }];

    const val = (k) => c[k] === null || c[k] === undefined ? '' : esc(c[k]);
    const num = (k, t, extra = '') => `<div><label for="c-${k}">${t}</label><input type="number" id="c-${k}" value="${val(k)}" ${extra}></div>`;
    const v = mostrar(`<div class="hoja">
      <a class="volver" href="#sitio/${s.id}">Volver a ${esc(s.codigo)}</a>
      <div class="encabezado-vista"><div><h1>${idCrono ? 'Editar cronología' : 'Nueva cronología'}</h1>
        <p class="suave">Sitio ${esc(s.codigo)}, ${esc(s.nombre)}. Solo la especie es obligatoria; completa lo que tengas.</p></div></div>
      <form id="form-crono" class="caja" novalidate>
        <fieldset><legend>Cronología</legend>
          <div class="rejilla">
            <div><label for="c-especie" class="requerido">Especie</label>
              <select id="c-especie"><option value="">Elige una especie</option>
                ${estado.especies.map((e) => `<option value="${e.id}"${e.id === c.especie_id ? ' selected' : ''}>${esc(e.nombre_cientifico)}${e.nombre_comun ? ' (' + esc(e.nombre_comun.split(',')[0]) + ')' : ''}</option>`).join('')}
                <option value="nueva">Otra especie…</option></select></div>
            <div><label for="c-parametro">Parámetro</label><select id="c-parametro">${opciones(ETIQ.parametro, c.parametro)}</select></div>
            <div id="c-nueva-especie" class="ancho-total rejilla" hidden>
              <div><label for="n-cientifico" class="requerido">Nombre científico</label><input type="text" id="n-cientifico" placeholder="Género especie"></div>
              <div><label for="n-comun">Nombre común</label><input type="text" id="n-comun"></div>
              <div><label for="n-familia">Familia</label><input type="text" id="n-familia"></div>
              <div><label for="n-itrdb">Código ITRDB</label><input type="text" id="n-itrdb" maxlength="4" placeholder="4 letras, si existe"></div>
            </div>
            <div><label for="c-codigo" class="requerido">Código de la cronología</label>
              <input type="text" id="c-codigo" value="${val('codigo')}" maxlength="40" spellcheck="false">
              <span class="ayuda">Se sugiere uno a partir del sitio, la especie y el parámetro</span></div>
            <div><label for="c-estado">Estado</label><select id="c-estado">${opciones(ETIQ.estado, c.estado)}</select></div>
          </div>
        </fieldset>
        <fieldset><legend>Período y muestra</legend>
          <div class="rejilla rejilla-4">
            ${num('anio_inicio', 'Año inicial', 'step="1"')}${num('anio_fin', 'Año final', 'step="1"')}
            ${num('anio_eps_085', 'EPS ≥ 0,85 desde', 'step="1"')}${num('anio_muestreo', 'Año de muestreo', 'step="1" min="1900" max="2100"')}
            ${num('n_arboles', 'N.º de árboles', 'step="1" min="1"')}${num('n_radios', 'N.º de radios', 'step="1" min="1"')}
          </div>
        </fieldset>
        <fieldset><legend>Estadísticos</legend>
          <div class="rejilla rejilla-4">
            ${num('intercorrelacion', 'Intercorrelación', 'step="0.001" min="-1" max="1"')}${num('sensibilidad_media', 'Sensibilidad media', 'step="0.001" min="0"')}
            ${num('rbar', 'rbar', 'step="0.001" min="-1" max="1"')}${num('eps', 'EPS', 'step="0.001" min="0" max="1"')}
            <div class="ancho-total"><label for="c-estandarizacion">Estandarización</label><input type="text" id="c-estandarizacion" value="${val('estandarizacion')}" placeholder="Ej.: spline cúbico de 2/3 de la longitud de la serie"></div>
          </div>
        </fieldset>
        <fieldset><legend>Serie de tiempo y publicación</legend>
          <div class="rejilla">
            <div><label for="c-disponibilidad">Disponibilidad de la serie</label><select id="c-disponibilidad">${opciones(ETIQ.disponibilidad, c.disponibilidad)}</select>
              <span class="ayuda">"A pedido": los usuarios registrados verán tu correo para pedírtela</span></div>
            <div><label for="c-enlace">Enlace a la serie</label><input type="url" id="c-enlace" value="${val('enlace_serie')}" placeholder="https://"></div>
            <div><label for="c-doi">DOI de la publicación</label><input type="text" id="c-doi" value="${val('doi')}" placeholder="10.1016/j.dendro.2026.000000"></div>
            <div><label for="c-cita">Cita</label><input type="text" id="c-cita" value="${val('cita')}"></div>
            <div class="ancho-total"><label for="c-notas">Notas</label><textarea id="c-notas">${val('notas')}</textarea></div>
          </div>
        </fieldset>
        <fieldset><legend>Autoría</legend>
          <div class="rejilla" style="margin-bottom:.9rem"><div><label for="c-lab">Laboratorio</label>
            <select id="c-lab"><option value="">Sin laboratorio</option>${estado.laboratorios.map((l) => `<option value="${l.id}"${l.id === c.laboratorio_id ? ' selected' : ''}>${esc(l.nombre)}</option>`).join('')}</select>
            <span class="ayuda">La institución que desarrolló esta cronología, si es distinta de la que muestreó el sitio</span></div></div>
          <p class="ayuda" style="margin-bottom:.6rem">En el orden en que deben citarse. Si la persona tiene cuenta, elige su nombre de la lista y se vinculará a su perfil.
            <strong>Responsable</strong> es quien responde por la cronología; <strong>autor</strong>, quien la hizo y cofechó las muestras.
            Marca las dos casillas cuando sea la misma persona.</p>
          <div class="tabla-envoltura"><table class="autores-tabla"><thead><tr><th>Nombre</th><th>ORCID (opcional)</th><th>Responsable</th><th>Autor</th><th><span class="saltar">Acciones</span></th></tr></thead><tbody id="autores"></tbody></table></div>
          <datalist id="l-perfiles">${estado.perfilesPublicos.map((p) => `<option value="${esc(p.nombre_completo)}">`).join('')}</datalist>
          <p style="margin-top:.6rem"><button type="button" class="boton secundario chico" id="b-autor">Agregar autor</button></p>
        </fieldset>
        <div id="c-errores"></div>
        <div class="fila-botones"><button type="submit" class="boton">${idCrono ? 'Guardar cambios' : 'Guardar cronología'}</button>
          <a class="boton secundario" href="#sitio/${s.id}">Cancelar</a></div>
      </form></div>`, turno);

    // Código sugerido mientras la persona no lo edite
    let codigoTocado = !!idCrono;
    const sugerirCodigo = () => {
      if (codigoTocado) return;
      const e = estado.especies.find((x) => x.id === Number($('#c-especie').value));
      const esp = e ? (e.codigo_itrdb || I.codigoValido(e.nombre_cientifico).slice(0, 4)).trim()
        : ($('#c-especie').value === 'nueva' ? (I.codigoValido($('#n-itrdb').value) || I.codigoValido($('#n-cientifico').value).slice(0, 4)) : '');
      const suf = { ancho_total: 'AT', madera_temprana: 'MT', madera_tardia: 'MTA', densidad_maxima: 'DM', anatomia: 'AN', isotopos: 'IS', otro: 'OT' }[$('#c-parametro').value];
      $('#c-codigo').value = esp ? s.codigo + '-' + esp + '-' + suf : '';
    };
    $('#c-codigo').addEventListener('input', () => { codigoTocado = true; });
    $('#c-especie').addEventListener('change', () => { $('#c-nueva-especie').hidden = $('#c-especie').value !== 'nueva'; sugerirCodigo(); });
    ['c-parametro', 'n-cientifico', 'n-itrdb'].forEach((i) => $('#' + i).addEventListener('input', sugerirCodigo));
    sugerirCodigo();

    // Tabla de autores
    function pintarAutores() {
      $('#autores').innerHTML = autores.map((a, i) => `<tr>
        <td><input type="text" data-i="${i}" data-k="nombre" list="l-perfiles" value="${esc(a.nombre)}" aria-label="Nombre del autor ${i + 1}"></td>
        <td><input type="text" data-i="${i}" data-k="orcid" value="${esc(a.orcid)}" placeholder="0000-0000-0000-0000" aria-label="ORCID del autor ${i + 1}"${a.perfil_id ? ' readonly title="Viene de su perfil"' : ''}></td>
        <td class="num"><input type="checkbox" data-i="${i}" data-k="es_responsable"${a.es_responsable ? ' checked' : ''} aria-label="Responsable de la cronología (persona ${i + 1})"></td>
        <td class="num"><input type="checkbox" data-i="${i}" data-k="es_autor"${a.es_autor ? ' checked' : ''} aria-label="Autor de la cronología (persona ${i + 1})"></td>
        <td style="white-space:nowrap">
          <button type="button" class="boton secundario chico" data-subir="${i}" ${i === 0 ? 'disabled' : ''} aria-label="Subir">↑</button>
          <button type="button" class="boton peligro chico" data-quitar="${i}" aria-label="Quitar">✕</button></td></tr>`).join('');
      $$('#autores [data-k]').forEach((el) => el.addEventListener('change', () => {
        const a = autores[+el.dataset.i];
        a[el.dataset.k] = el.type === 'checkbox' ? el.checked : el.value;
        // Una persona que no es ni responsable ni autora no tiene por qué estar
        // en la lista: se devuelve la casilla que acaba de desmarcarse.
        if (!a.es_responsable && !a.es_autor) {
          a[el.dataset.k] = true;
          pintarAutores();
          aviso('Cada persona tiene que ser responsable, autora, o las dos cosas', 'error');
          return;
        }
        if (el.dataset.k === 'nombre') {
          const p = estado.perfilesPublicos.find((x) => x.nombre_completo === el.value.trim());
          a.perfil_id = p ? p.id : null;
          if (p) a.orcid = p.orcid || '';
          pintarAutores();
        }
      }));
      $$('[data-quitar]').forEach((b) => b.addEventListener('click', () => { autores.splice(+b.dataset.quitar, 1); pintarAutores(); }));
      $$('[data-subir]').forEach((b) => b.addEventListener('click', () => {
        const i = +b.dataset.subir; [autores[i - 1], autores[i]] = [autores[i], autores[i - 1]]; pintarAutores();
      }));
    }
    $('#b-autor').addEventListener('click', () => { autores.push({ perfil_id: null, nombre: '', orcid: '', es_responsable: false, es_autor: true }); pintarAutores(); });
    pintarAutores();

    $('#form-crono').addEventListener('submit', async (e) => {
      e.preventDefault();
      const errores = [];
      const ent = (k) => { const t = $('#c-' + k).value; return t === '' ? null : Number(t); };
      const fila = {
        sitio_id: s.id, codigo: $('#c-codigo').value.trim().toUpperCase(), parametro: $('#c-parametro').value, estado: $('#c-estado').value,
        anio_inicio: ent('anio_inicio'), anio_fin: ent('anio_fin'), anio_eps_085: ent('anio_eps_085'), anio_muestreo: ent('anio_muestreo'),
        n_arboles: ent('n_arboles'), n_radios: ent('n_radios'),
        intercorrelacion: ent('intercorrelacion'), sensibilidad_media: ent('sensibilidad_media'), rbar: ent('rbar'), eps: ent('eps'),
        estandarizacion: vacioANulo($('#c-estandarizacion').value), disponibilidad: $('#c-disponibilidad').value,
        enlace_serie: vacioANulo($('#c-enlace').value), doi: vacioANulo($('#c-doi').value), cita: vacioANulo($('#c-cita').value), notas: vacioANulo($('#c-notas').value),
        laboratorio_id: $('#c-lab').value ? Number($('#c-lab').value) : null
      };
      const esNueva = $('#c-especie').value === 'nueva';
      if (!$('#c-especie').value) errores.push('Elige la especie.');
      if (esNueva && !$('#n-cientifico').value.trim()) errores.push('Escribe el nombre científico de la especie nueva.');
      if (!fila.codigo) errores.push('Falta el código de la cronología.');
      ['anio_inicio', 'anio_fin', 'anio_eps_085', 'anio_muestreo', 'n_arboles', 'n_radios'].forEach((k) => {
        if (fila[k] !== null && !Number.isInteger(fila[k])) errores.push('"' + $('label[for=c-' + k + ']').textContent + '" debe ser un número entero.');
      });
      if (fila.anio_inicio !== null && fila.anio_fin !== null && fila.anio_fin < fila.anio_inicio) errores.push('El año final es anterior al inicial.');
      if (fila.anio_eps_085 !== null && fila.anio_inicio !== null && fila.anio_fin !== null && (fila.anio_eps_085 < fila.anio_inicio || fila.anio_eps_085 > fila.anio_fin)) errores.push('El año con EPS ≥ 0,85 debe estar dentro del período.');
      if (fila.n_radios !== null && fila.n_arboles !== null && fila.n_radios < fila.n_arboles) errores.push('Hay menos radios que árboles.');
      [['rbar', -1, 1], ['eps', 0, 1], ['intercorrelacion', -1, 1]].forEach(([k, a, b]) => { if (fila[k] !== null && (fila[k] < a || fila[k] > b)) errores.push(k + ' debe estar entre ' + a + ' y ' + b + '.'); });
      if (fila.enlace_serie && !urlSegura(fila.enlace_serie)) errores.push('El enlace a la serie debe comenzar con https:// o http://');
      if (fila.doi && !enlaceDOI(fila.doi)) errores.push('El DOI no tiene un formato válido (10.xxxx/...).');
      const autoresValidos = autores.filter((a) => (a.nombre || '').trim());
      autoresValidos.forEach((a) => { if (a.orcid && !/^\d{4}-\d{4}-\d{4}-\d{3}[\dX]$/.test(a.orcid.trim())) errores.push('ORCID inválido para ' + a.nombre + '.'); });
      if (errores.length) { $('#c-errores').innerHTML = `<div class="mensaje error"><p>${errores.map(esc).join('<br>')}</p></div>`; return; }
      $('#c-errores').innerHTML = '';
      const boton = $('button[type=submit]', e.target);
      boton.disabled = true;
      try {
        if (esNueva) {
          const { data: ne, error: ee } = await sb.from('especies').insert({
            nombre_cientifico: $('#n-cientifico').value.trim(), nombre_comun: vacioANulo($('#n-comun').value),
            familia: vacioANulo($('#n-familia').value), codigo_itrdb: vacioANulo(I.codigoValido($('#n-itrdb').value))
          }).select().single();
          if (ee) throw ee;
          estado.especies.push(ne);
          fila.especie_id = ne.id;
        } else fila.especie_id = Number($('#c-especie').value);

        const q = idCrono ? sb.from('cronologias').update(fila).eq('id', idCrono).select('id') : sb.from('cronologias').insert(fila).select('id');
        const { data, error } = await q;
        if (error) throw error;
        if (!data.length) throw new Error('No se guardó: no tienes permiso sobre esta cronología.');
        const idC = data[0].id;
        // Autores: se reemplaza la lista completa
        const { error: ed } = await sb.from('cronologia_autores').delete().eq('cronologia_id', idC);
        if (ed) throw ed;
        if (autoresValidos.length) {
          const { error: ea } = await sb.from('cronologia_autores').insert(autoresValidos.map((a, i) => ({
            cronologia_id: idC, orden: i + 1, perfil_id: a.perfil_id || null,
            nombre_autor: a.perfil_id ? null : a.nombre.trim(), orcid: a.perfil_id ? null : vacioANulo(a.orcid),
            es_responsable: !!a.es_responsable, es_autor: !!a.es_autor
          })));
          if (ea) throw ea;
        }
        estado.cargado = false;
        aviso(idCrono ? 'Cambios guardados' : 'Cronología ' + fila.codigo + ' guardada');
        location.hash = '#sitio/' + s.id;
      } catch (err) {
        $('#c-errores').innerHTML = `<div class="mensaje error"><p>${esc(mensajeError(err))}</p></div>`;
      } finally { boton.disabled = false; }
    });
  }

  // ===================================================================
  // Importador
  // ===================================================================
  async function vistaImportar(turno) {
    await cargarBase(true);
    const imp = { archivo: null, lectura: null, asignacion: {}, formatoCoord: 'auto', datum: 'WGS84', zonaFija: '19', aproximadas: false, validacion: null };
    const v = mostrar(`<div class="hoja-ancha">
      <div class="encabezado-vista"><div><h1>Importar sitios y cronologías</h1>
        <p class="suave">Sube una tabla o un archivo GPS. Revisarás cada fila antes de guardar.</p></div>
        <a class="boton secundario chico" href="plantilla_importacion.csv" download>Descargar plantilla</a></div>
      <ol class="pasos" id="pasos"><li>Archivo</li><li>Columnas</li><li>Revisión</li><li>Guardado</li></ol>
      <div id="paso"></div></div>`, turno);
    const marcarPaso = (n) => $$('#pasos li').forEach((li, i) => { li.className = i < n ? 'hecho' : i === n ? 'actual' : ''; });

    // ---- Paso 1 ----
    function paso1(error) {
      marcarPaso(0);
      $('#paso').innerHTML = `<div class="caja">
        ${error ? `<div class="mensaje error"><p>${esc(error)}</p></div>` : ''}
        <label class="zona-archivo" id="zona" for="archivo">
          <strong>Arrastra el archivo aquí o haz clic para elegirlo</strong>
          <span class="suave chico">CSV, TXT o TSV (separado por comas, punto y coma o tabulaciones), Excel, GPX, KML o GeoJSON</span>
        </label>
        <input type="file" id="archivo" accept=".csv,.txt,.tsv,.xlsx,.xls,.gpx,.kml,.geojson,.json" hidden>
        <p class="ayuda" style="margin-top:.75rem">Cada fila es un sitio, o un sitio con una cronología. Si un sitio tiene varias cronologías, repite su código en varias filas.
        Las coordenadas pueden venir en grados decimales, grados-minutos-segundos o UTM, y en WGS84, PSAD56 o SAD69. De un GPS se importan los puntos guardados.</p></div>`;
      const zona = $('#zona'), input = $('#archivo');
      input.addEventListener('change', () => input.files[0] && leer(input.files[0]));
      ['dragenter', 'dragover'].forEach((ev) => zona.addEventListener(ev, (e) => { e.preventDefault(); zona.classList.add('encima'); }));
      ['dragleave', 'drop'].forEach((ev) => zona.addEventListener(ev, (e) => { e.preventDefault(); zona.classList.remove('encima'); }));
      zona.addEventListener('drop', (e) => e.dataTransfer.files[0] && leer(e.dataTransfer.files[0]));
    }
    async function leer(archivo) {
      try {
        if (I.formatoDe(archivo.name) === 'xlsx') await cargarScript(URL_XLSX);
        imp.archivo = archivo;
        imp.lectura = await I.leerArchivo(archivo);
        imp.asignacion = I.adivinarAsignacion(imp.lectura.columnas);
        imp.formatoCoord = I.sugerirFormatoCoord(imp.lectura.filas, imp.asignacion);
        paso2();
      } catch (e) { paso1('No se pudo leer el archivo: ' + e.message); }
    }

    // ---- Paso 2 ----
    function paso2() {
      marcarPaso(1);
      const L1 = imp.lectura;
      const opcionesCol = (clave) => '<option value="">No importar</option>' +
        L1.columnas.map((c) => `<option value="${esc(c)}"${imp.asignacion[clave] === c ? ' selected' : ''}>${esc(c)}</option>`).join('');
      const muestra = (clave) => {
        const col = imp.asignacion[clave];
        if (!col) return '';
        return L1.filas.slice(0, 3).map((f) => f[col]).filter((x) => x !== '' && x !== null && x !== undefined).map((x) => esc(String(x))).join(', ');
      };
      const filaCampo = (cp) => `<tr><td>${esc(cp.etiqueta)}${cp.requerido ? ' <span class="requerido"></span>' : ''}</td>
        <td><select data-campo="${cp.clave}" aria-label="Columna para ${esc(cp.etiqueta)}">${opcionesCol(cp.clave)}</select></td>
        <td class="muestra" data-muestra="${cp.clave}">${muestra(cp.clave)}</td></tr>`;
      const camposSitio = I.CAMPOS.filter((cp) => cp.grupo === 'sitio' && (!cp.coord || (imp.formatoCoord === 'utm' ? cp.coord === 'utm' : cp.coord === 'geo')));
      const camposCrono = I.CAMPOS.filter((cp) => cp.grupo === 'crono');
      $('#paso').innerHTML = `<div class="caja">
        <p><strong>${esc(imp.archivo.name)}</strong>: ${plural(L1.filas.length, 'fila', 'filas')} y ${plural(L1.columnas.length, 'columna', 'columnas')}.
          ${L1.hojas && L1.hojas.length > 1 ? `Se leyó la primera hoja (${esc(L1.hojas[0])}).` : ''} Revisa a qué campo corresponde cada columna; las que reconocí ya están asignadas.</p>
        <div class="rejilla rejilla-3" style="margin-bottom:1rem">
          <div><label for="i-fmt">Formato de coordenadas</label><select id="i-fmt">
            <option value="auto"${imp.formatoCoord === 'auto' ? ' selected' : ''}>Grados (decimales o GMS)</option>
            <option value="utm"${imp.formatoCoord === 'utm' ? ' selected' : ''}>UTM</option></select></div>
          <div><label for="i-datum">Datum</label><select id="i-datum">${Object.entries(C.DATUMS).map(([k, d]) => `<option value="${k}"${k === imp.datum ? ' selected' : ''}>${esc(d.etiqueta)}</option>`).join('')}</select></div>
          <div ${imp.formatoCoord === 'utm' ? '' : 'hidden'}><label for="i-zona">Zona UTM si no hay columna</label><select id="i-zona">${[17, 18, 19, 20].map((z) => `<option value="${z}"${String(z) === imp.zonaFija ? ' selected' : ''}>${z} S</option>`).join('')}</select></div>
        </div>
        <div class="dos-columnas">
          <div><h3>Sitio</h3><div class="tabla-envoltura"><table class="asignacion"><thead><tr><th>Campo</th><th>Columna del archivo</th><th>Ejemplo</th></tr></thead><tbody>${camposSitio.map(filaCampo).join('')}</tbody></table></div></div>
          <div><h3>Cronología</h3><p class="ayuda" style="margin-top:-.25rem">Opcional. Se crea una cronología en cada fila que tenga especie.</p>
            <div class="tabla-envoltura"><table class="asignacion"><thead><tr><th>Campo</th><th>Columna del archivo</th><th>Ejemplo</th></tr></thead><tbody>${camposCrono.map(filaCampo).join('')}</tbody></table></div></div>
        </div>
        <div id="i-err2"></div>
        <div class="fila-botones" style="margin-top:1rem"><button type="button" class="boton" id="b-revisar">Revisar filas</button>
          <button type="button" class="boton secundario" id="b-otro">Elegir otro archivo</button></div></div>`;
      $$('[data-campo]').forEach((sel) => sel.addEventListener('change', () => {
        if (sel.value) imp.asignacion[sel.dataset.campo] = sel.value; else delete imp.asignacion[sel.dataset.campo];
        $(`[data-muestra="${sel.dataset.campo}"]`).innerHTML = muestra(sel.dataset.campo);
      }));
      $('#i-fmt').addEventListener('change', (e) => { imp.formatoCoord = e.target.value; paso2(); });
      $('#i-datum').addEventListener('change', (e) => { imp.datum = e.target.value; });
      $('#i-zona').addEventListener('change', (e) => { imp.zonaFija = e.target.value; });
      $('#b-otro').addEventListener('click', () => paso1());
      $('#b-revisar').addEventListener('click', () => {
        const a = imp.asignacion, faltan = [];
        if (!a.codigo && !a.nombre) faltan.push('el código o el nombre del sitio');
        if (imp.formatoCoord === 'utm') { if (!a.este || !a.norte) faltan.push('las columnas de UTM este y norte'); }
        else if (!a.latitud || !a.longitud) faltan.push('las columnas de latitud y longitud');
        if (faltan.length) { $('#i-err2').innerHTML = `<div class="mensaje error" style="margin-top:1rem"><p>Asigna ${faltan.join(' y ')}.</p></div>`; return; }
        paso3();
      });
    }

    // ---- Paso 3 ----
    function paso3() {
      marcarPaso(2);
      const val = I.validar(imp.lectura.filas, imp.asignacion, {
        formatoCoord: imp.formatoCoord === 'utm' ? 'utm' : 'auto', datum: imp.datum, zonaFija: imp.zonaFija,
        especies: estado.especies, sitiosExistentes: estado.catalogo, hemisferioSurOeste: true
      });
      // Una fila con especie desconocida no se puede guardar hasta agregarla
      val.filas.forEach((f) => { if (f.crono && !f.crono.especie_id) f.errores.push('especie "' + f.crono.especie_texto + '" no está en el catálogo'); });
      imp.validacion = val;
      const buenas = val.filas.filter((f) => !f.errores.length);
      const nuevos = new Set(buenas.filter((f) => !f.existente).map((f) => f.sitio.codigo));
      const existentes = new Set(buenas.filter((f) => f.existente).map((f) => f.sitio.codigo));
      const nCronos = buenas.filter((f) => f.crono).length;
      const conError = val.filas.length - buenas.length;
      const desconocidas = [...val.especiesDesconocidas];
      let soloProblemas = conError > 0;

      $('#paso').innerHTML = `
        <div class="cifras">
          <div><strong>${nuevos.size}</strong><span>sitios nuevos</span></div>
          <div><strong>${existentes.size}</strong><span>sitios ya registrados (solo se agregan cronologías)</span></div>
          <div><strong>${nCronos}</strong><span>cronologías</span></div>
          <div><strong>${conError}</strong><span>filas con errores (no se guardan)</span></div>
        </div>
        ${desconocidas.length ? `<div class="mensaje aviso"><p><strong>${desconocidas.length === 1 ? 'Una especie no está' : desconocidas.length + ' especies no están'} en el catálogo:</strong> ${desconocidas.map((d) => '<em>' + esc(d) + '</em>').join(', ')}.</p>
          <p>Si el nombre está bien escrito, agrégala; si es un error, corrígelo en tu archivo o cambia la columna asignada.</p>
          <p><button type="button" class="boton chico" id="b-especies">Agregar ${desconocidas.length === 1 ? 'esta especie' : 'estas especies'} al catálogo</button></p></div>` : ''}
        <div class="caja">
          <div class="fila-botones" style="justify-content:space-between;margin-bottom:.75rem">
            <label class="marca-check"><input type="checkbox" id="i-solo"${soloProblemas ? ' checked' : ''}> Mostrar solo filas con errores o avisos</label>
            <label class="marca-check"><input type="checkbox" id="i-aprox"${imp.aproximadas ? ' checked' : ''}> Publicar todas las coordenadas aproximadas (~1 km)</label>
          </div>
          <div class="tabla-envoltura" style="max-height:60vh"><table><thead><tr><th class="num">Fila</th><th>Estado</th><th>Sitio</th><th class="num">Latitud</th><th class="num">Longitud</th><th>Cronología</th><th>Detalle</th></tr></thead>
            <tbody id="i-filas"></tbody></table></div>
          <div class="fila-botones" style="margin-top:1rem">
            <button type="button" class="boton" id="b-guardar"${buenas.length ? '' : ' disabled'}>Guardar ${plural(nuevos.size, 'sitio', 'sitios')} y ${plural(nCronos, 'cronología', 'cronologías')}</button>
            <button type="button" class="boton secundario" id="b-volver2">Cambiar columnas</button>
          </div>
        </div>`;
      const pintarFilas = () => {
        const filas = soloProblemas ? val.filas.filter((f) => f.errores.length || f.avisos.length) : val.filas;
        $('#i-filas').innerHTML = filas.length ? filas.map((f) => {
          const tipo = f.errores.length ? 'error' : f.avisos.length ? 'aviso' : 'ok';
          return `<tr><td class="num">${f.n}</td>
            <td><span class="estado-fila ${tipo}">${{ ok: 'Lista', aviso: 'Revisar', error: 'Error' }[tipo]}</span></td>
            <td><strong>${esc(f.sitio.codigo)}</strong><br><span class="suave chico">${esc(f.sitio.nombre)}</span></td>
            <td class="num">${f.sitio.latitud === null ? '—' : f.sitio.latitud.toFixed(5)}</td><td class="num">${f.sitio.longitud === null ? '—' : f.sitio.longitud.toFixed(5)}</td>
            <td>${f.crono ? '<em>' + esc(f.crono.especie_texto) + '</em><br><span class="suave chico">' + esc(f.crono.codigo) + ', ' + periodo(f.crono.anio_inicio, f.crono.anio_fin) + '</span>' : '<span class="suave">—</span>'}</td>
            <td><ul class="problemas">${f.errores.map((x) => `<li class="error">${esc(x)}</li>`).join('')}${f.avisos.map((x) => `<li class="aviso">${esc(x)}</li>`).join('')}</ul></td></tr>`;
        }).join('') : '<tr><td colspan="7" class="suave">Todas las filas están listas.</td></tr>';
      };
      pintarFilas();
      $('#i-solo').addEventListener('change', (e) => { soloProblemas = e.target.checked; pintarFilas(); });
      $('#i-aprox').addEventListener('change', (e) => { imp.aproximadas = e.target.checked; });
      $('#b-volver2').addEventListener('click', paso2);
      const bEsp = $('#b-especies');
      if (bEsp) bEsp.addEventListener('click', async () => {
        bEsp.disabled = true;
        const { data, error } = await sb.from('especies').insert(desconocidas.map((n) => ({ nombre_cientifico: n.trim() }))).select();
        if (error) { aviso(mensajeError(error), 'error'); bEsp.disabled = false; return; }
        estado.especies = estado.especies.concat(data);
        aviso(plural(data.length, 'especie agregada', 'especies agregadas'));
        paso3();
      });
      $('#b-guardar').addEventListener('click', () => guardar(buenas));
    }

    // ---- Paso 4 ----
    async function guardar(buenas) {
      marcarPaso(3);
      $('#paso').innerHTML = `<div class="caja"><p id="g-texto">Guardando…</p><div class="barra-progreso"><div id="g-barra"></div></div></div>`;
      const progreso = (t, p) => { $('#g-texto').textContent = t; $('#g-barra').style.width = Math.round(p * 100) + '%'; };
      try {
        // Registro de la importación
        const { data: reg, error: er } = await sb.from('importaciones').insert({
          archivo: imp.archivo.name, formato: imp.lectura.formato === 'csv' && /\.tsv$/i.test(imp.archivo.name) ? 'tsv' : imp.lectura.formato,
          n_filas: imp.lectura.filas.length, asignacion_columnas: Object.assign({ formato_coordenadas: imp.formatoCoord, datum: imp.datum }, imp.asignacion)
        }).select('id').single();
        if (er) throw er;

        // El archivo trae nombres; la base guarda identificadores.
        const norm = (t) => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
        const labPorNombre = new Map();
        estado.laboratorios.forEach((l) => {
          labPorNombre.set(norm(l.nombre), l.id);
          if (l.institucion) labPorNombre.set(norm(l.institucion), l.id);
        });
        const labsDesconocidos = new Set();
        // undefined = la fila no trae laboratorio; null = lo trae pero no se reconoce.
        const idLab = (texto) => {
          if (!texto) return undefined;
          const id = labPorNombre.get(norm(texto));
          if (!id) { labsDesconocidos.add(texto); return null; }
          return id;
        };
        const { data: listaPerfiles } = await sb.from('perfiles').select('id, nombre_completo');
        const perfilPorNombre = new Map((listaPerfiles || [])
          .filter((p) => p.nombre_completo).map((p) => [norm(p.nombre_completo), p.id]));

        // "Pérez, J.; Soto, M." → el responsable primero y los autores después,
        // en el orden del archivo. Separador ";" o "|", nunca la coma, porque
        // los nombres suelen llevarla. Si el nombre coincide con un perfil
        // registrado se vincula a su cuenta; si no, queda como texto.
        //
        // Responder por una cronología y haberla cofechado son cosas
        // distintas, así que no se deduce una de la otra: el responsable solo
        // queda como autor si el archivo lo nombra también entre los autores.
        const autoresDeTextos = (resp, otros) => {
          const partir = (t) => String(t || '').split(/\s*[;|]\s*/).map((x) => x.trim()).filter(Boolean);
          const lista = [];
          partir(resp).slice(0, 1).forEach((n) => lista.push({ nombre: n, es_responsable: true, es_autor: false }));
          partir(otros).forEach((n) => {
            const ya = lista.find((a) => norm(a.nombre) === norm(n));
            if (ya) ya.es_autor = true;   // es responsable y además cofechó
            else lista.push({ nombre: n, es_responsable: false, es_autor: true });
          });
          return lista.map((a, i) => {
            const pid = perfilPorNombre.get(norm(a.nombre));
            return { orden: i + 1, es_responsable: a.es_responsable, es_autor: a.es_autor,
                     perfil_id: pid || null, nombre_autor: pid ? null : a.nombre };
          });
        };

        // Sitios nuevos (uno por código)
        const red = (x) => imp.aproximadas ? Math.round(x * 100) / 100 : x;
        const nuevos = new Map();
        buenas.filter((f) => !f.existente).forEach((f) => {
          if (!nuevos.has(f.sitio.codigo)) {
            const lab = idLab(f.sitio.laboratorio_texto);
            const fila = Object.assign({}, f.sitio, {
              latitud: red(f.sitio.latitud), longitud: red(f.sitio.longitud),
              precision: imp.aproximadas ? 'aproximada' : 'exacta', coord_original: imp.aproximadas ? null : f.sitio.coord_original,
              importacion_id: reg.id,
              // Si el archivo no trae laboratorio, se asume el de quien importa.
              laboratorio_id: lab === undefined ? (estado.perfil.laboratorio_id || null) : lab
            });
            delete fila.laboratorio_texto;
            nuevos.set(f.sitio.codigo, fila);
          }
        });
        const idPorCodigo = new Map(estado.catalogo.map((s) => [s.codigo, s.id]));
        const listaNuevos = [...nuevos.values()];
        for (let i = 0; i < listaNuevos.length; i += 200) {
          progreso('Guardando sitios…', i / Math.max(listaNuevos.length, 1) * .5);
          const { data, error } = await sb.from('sitios').insert(listaNuevos.slice(i, i + 200)).select('id, codigo');
          if (error) throw error;
          data.forEach((d) => idPorCodigo.set(d.codigo, d.id));
        }

        // Cronologías, evitando códigos repetidos dentro de cada sitio
        const conCrono = buenas.filter((f) => f.crono);
        const idsSitios = [...new Set(conCrono.map((f) => idPorCodigo.get(f.sitio.codigo)))];
        const usados = new Set();
        if (idsSitios.length) {
          const previas = await traerTodo(() => sb.from('cronologias').select('sitio_id, codigo').in('sitio_id', idsSitios));
          previas.forEach((p) => usados.add(p.sitio_id + '|' + p.codigo));
        }
        const renombradas = [];
        const autoresPorClave = new Map();   // sitio_id|codigo → filas de cronologia_autores
        const cronos = conCrono.map((f) => {
          const sitio_id = idPorCodigo.get(f.sitio.codigo);
          let codigo = f.crono.codigo, k = 2;
          while (usados.has(sitio_id + '|' + codigo)) codigo = f.crono.codigo + '-' + k++;
          if (codigo !== f.crono.codigo) renombradas.push(f.crono.codigo + ' → ' + codigo);
          usados.add(sitio_id + '|' + codigo);
          autoresPorClave.set(sitio_id + '|' + codigo,
            autoresDeTextos(f.crono.responsable_texto, f.crono.autores_texto));
          const lab = idLab(f.crono.laboratorio_texto);
          const c = Object.assign({}, f.crono, {
            sitio_id, codigo, importacion_id: reg.id,
            laboratorio_id: lab === undefined ? null : lab
          });
          delete c.especie_texto;
          delete c.laboratorio_texto; delete c.responsable_texto; delete c.autores_texto;
          return c;
        });
        let guardadas = 0;
        let sinAutoria = 0;
        for (let i = 0; i < cronos.length; i += 200) {
          progreso('Guardando cronologías…', .5 + i / Math.max(cronos.length, 1) * .5);
          const { data, error } = await sb.from('cronologias').insert(cronos.slice(i, i + 200)).select('id, sitio_id, codigo');
          if (error) throw error;
          guardadas += data.length;
          // Autoría del archivo. Si la fila no trae nada, queda como
          // responsable quien importa, porque alguien tiene que responder por
          // la cronología; pero no como autor, porque nada dice que la haya
          // cofechado. El resumen final avisa para que se corrija.
          const filasAutores = [];
          data.forEach((d) => {
            const lista = autoresPorClave.get(d.sitio_id + '|' + d.codigo) || [];
            if (lista.length) lista.forEach((a) => filasAutores.push(Object.assign({ cronologia_id: d.id }, a)));
            else { sinAutoria++; filasAutores.push({ cronologia_id: d.id, orden: 1, perfil_id: uid(), es_responsable: true, es_autor: false }); }
          });
          if (filasAutores.length) {
            const { error: ea } = await sb.from('cronologia_autores').insert(filasAutores);
            if (ea) throw ea;
          }
        }
        progreso('Listo', 1);
        estado.cargado = false;
        $('#paso').innerHTML = `<div class="caja"><div class="mensaje ok"><p><strong>Importación terminada.</strong> Se guardaron ${plural(listaNuevos.length, 'sitio nuevo', 'sitios nuevos')} y ${plural(guardadas, 'cronología', 'cronologías')}.</p>
          ${renombradas.length ? `<p>Algunos códigos de cronología ya existían y se renombraron: ${renombradas.map(esc).join(', ')}.</p>` : ''}
          ${labsDesconocidos.size ? `<p><strong>Laboratorios no reconocidos</strong>, guardados sin laboratorio: ${[...labsDesconocidos].map(esc).join(', ')}. Agrégalos en Administrar → Laboratorios y corrige los registros.</p>` : ''}
          ${sinAutoria ? `<p><strong>${plural(sinAutoria, 'cronología quedó', 'cronologías quedaron')} contigo como responsable y sin autores</strong>, porque el archivo no traía las columnas de responsable ni de autores. Complétalas desde cada ficha: responder por una cronología no es lo mismo que haberla cofechado.</p>`
            : '<p>La autoría se tomó del archivo. Revisa alguna ficha para confirmar que el responsable y los autores quedaron como esperabas.</p>'}</div>
          <p class="fila-botones"><a class="boton" href="#mapa">Ver en el mapa</a><a class="boton secundario" href="#mis-registros">Ir a mis registros</a><button type="button" class="boton secundario" id="b-otra">Importar otro archivo</button></p></div>`;
        $('#b-otra').addEventListener('click', () => vistaImportar(++turnoVista));
      } catch (e) {
        console.error(e);
        estado.cargado = false;
        $('#paso').innerHTML = `<div class="caja"><div class="mensaje error"><p><strong>La importación se detuvo:</strong> ${esc(mensajeError(e))}</p>
          <p>Lo que alcanzó a guardarse quedó registrado. Revisa <a href="#mis-registros">tus registros</a> antes de volver a intentarlo, para no duplicar sitios.</p></div></div>`;
      }
    }
    paso1();
  }

  // ===================================================================
  // Cuenta
  // ===================================================================
  async function vistaCuenta(turno) {
    if (!uid()) return vistaIngreso(turno);
    await cargarPerfil();
    await cargarBase();
    const p = estado.perfil || {};
    const v = mostrar(`<div class="hoja">
      <div class="encabezado-vista"><div><h1>Mi cuenta</h1><p class="suave">${esc(estado.sesion.user.email)}</p></div>
        <button type="button" class="boton secundario chico" id="b-salir">Cerrar sesión</button></div>
      ${p.aprobado
        ? `<div class="mensaje ok"><p>Cuenta aprobada, con rol de ${ETIQ.rol[p.rol].toLowerCase()}. Ya puedes <a href="#sitio/nuevo">registrar sitios</a> o <a href="#importar">importar una tabla</a>.</p></div>`
        : `<div class="mensaje aviso"><p><strong>Tu cuenta está pendiente de aprobación.</strong> Un administrador la revisará; completa tu nombre y laboratorio para que pueda reconocerte.</p></div>`}
      <form id="form-perfil" class="caja" novalidate>
        <h2>Perfil</h2>
        <p class="ayuda" style="margin-top:-.25rem">Tu nombre, ORCID y laboratorio son públicos y aparecen como autoría de tus registros. Tu correo no se publica.</p>
        <div class="rejilla">
          <div class="ancho-total"><label for="p-nombre" class="requerido">Nombre completo</label><input type="text" id="p-nombre" value="${esc(p.nombre_completo)}" autocomplete="name"></div>
          <div><label for="p-orcid">ORCID</label><input type="text" id="p-orcid" value="${esc(p.orcid)}" placeholder="0000-0000-0000-0000">
            <span class="ayuda">Opcional</span></div>
          <div><label for="p-lab">Laboratorio</label><select id="p-lab"><option value="">Sin laboratorio</option>
            ${estado.laboratorios.map((l) => `<option value="${l.id}"${l.id === p.laboratorio_id ? ' selected' : ''}>${esc(l.nombre)}</option>`).join('')}</select>
            <span class="ayuda">¿No aparece? Pídele a un administrador que lo agregue.</span></div>
        </div>
        <div id="p-err"></div>
        <p class="fila-botones" style="margin-top:1rem"><button type="submit" class="boton">Guardar perfil</button></p>
      </form>
      <form id="form-clave" class="caja" novalidate>
        <h2>Cambiar contraseña</h2>
        <div class="rejilla"><div><label for="k-nueva">Nueva contraseña</label><input type="password" id="k-nueva" minlength="8" autocomplete="new-password"><span class="ayuda">Al menos 8 caracteres</span></div></div>
        <p class="fila-botones" style="margin-top:1rem"><button type="submit" class="boton secundario">Cambiar contraseña</button></p>
      </form></div>`, turno);
    $('#b-salir').addEventListener('click', async () => { await sb.auth.signOut(); aviso('Sesión cerrada'); location.hash = '#mapa'; });
    $('#form-perfil').addEventListener('submit', async (e) => {
      e.preventDefault();
      const nombre = $('#p-nombre').value.trim(), orcid = $('#p-orcid').value.trim().toUpperCase();
      const errs = [];
      if (!nombre) errs.push('Escribe tu nombre completo.');
      if (orcid && !/^\d{4}-\d{4}-\d{4}-\d{3}[\dX]$/.test(orcid)) errs.push('El ORCID debe tener el formato 0000-0000-0000-0000 (o déjalo vacío).');
      if (errs.length) { $('#p-err').innerHTML = `<div class="mensaje error"><p>${errs.join('<br>')}</p></div>`; return; }
      const { error } = await sb.from('perfiles').update({ nombre_completo: nombre, orcid: orcid || null, laboratorio_id: $('#p-lab').value ? Number($('#p-lab').value) : null }).eq('id', uid());
      if (error) { $('#p-err').innerHTML = `<div class="mensaje error"><p>${esc(mensajeError(error))}</p></div>`; return; }
      $('#p-err').innerHTML = '';
      await cargarPerfil();
      estado.cargado = false;
      pintarNav();
      aviso('Perfil guardado');
    });
    $('#form-clave').addEventListener('submit', async (e) => {
      e.preventDefault();
      const clave = $('#k-nueva').value;
      if (clave.length < 8) { aviso('La contraseña debe tener al menos 8 caracteres', 'error'); return; }
      const { error } = await sb.auth.updateUser({ password: clave });
      if (error) { aviso(mensajeError(error), 'error'); return; }
      $('#k-nueva').value = '';
      aviso('Contraseña cambiada');
    });
  }

  function vistaIngreso(turno) {
    const v = mostrar(`<div class="hoja">
      <div class="encabezado-vista"><div><h1>Ingresar</h1>
        <p class="suave">El catálogo es público. La cuenta es para registrar y editar tus propios sitios.</p></div></div>
      <div class="caja">
        <div class="pestanas" role="tablist">
          <button type="button" role="tab" aria-selected="true" data-tab="ingresar">Ya tengo cuenta</button>
          <button type="button" role="tab" aria-selected="false" data-tab="registro">Crear cuenta</button>
          <button type="button" role="tab" aria-selected="false" data-tab="olvido">Olvidé mi contraseña</button>
        </div>
        <form data-panel="ingresar" novalidate>
          <div class="rejilla">
            <div><label for="i-correo">Correo</label><input type="email" id="i-correo" autocomplete="email" required></div>
            <div><label for="i-clave">Contraseña</label><input type="password" id="i-clave" autocomplete="current-password" required></div>
          </div>
          <p class="fila-botones" style="margin-top:1rem"><button type="submit" class="boton">Ingresar</button></p>
        </form>
        <form data-panel="registro" hidden novalidate>
          <div class="rejilla">
            <div class="ancho-total"><label for="r-nombre" class="requerido">Nombre completo</label><input type="text" id="r-nombre" autocomplete="name" required>
              <span class="ayuda">Así aparecerá en tus registros</span></div>
            <div><label for="r-correo" class="requerido">Correo</label><input type="email" id="r-correo" autocomplete="email" required>
              <span class="ayuda">Mejor el institucional: ayuda a que aprueben tu cuenta</span></div>
            <div><label for="r-clave" class="requerido">Contraseña</label><input type="password" id="r-clave" autocomplete="new-password" minlength="8" required>
              <span class="ayuda">Al menos 8 caracteres</span></div>
          </div>
          <p class="ayuda" style="margin-top:.75rem">Después de confirmar tu correo, un administrador aprobará la cuenta. El ORCID y el laboratorio se agregan después en tu perfil.</p>
          <p class="fila-botones" style="margin-top:1rem"><button type="submit" class="boton">Crear cuenta</button></p>
        </form>
        <form data-panel="olvido" hidden novalidate>
          <div class="rejilla"><div><label for="o-correo">Correo de tu cuenta</label><input type="email" id="o-correo" autocomplete="email" required></div></div>
          <p class="fila-botones" style="margin-top:1rem"><button type="submit" class="boton">Enviar enlace para cambiarla</button></p>
        </form>
        <div id="i-msg" style="margin-top:1rem"></div>
      </div></div>`, turno);
    const msg = (html, tipo) => { $('#i-msg').innerHTML = `<div class="mensaje ${tipo}"><p>${html}</p></div>`; };
    $$('[data-tab]', v).forEach((b) => b.addEventListener('click', () => {
      $$('[data-tab]', v).forEach((x) => x.setAttribute('aria-selected', String(x === b)));
      $$('[data-panel]', v).forEach((f) => { f.hidden = f.dataset.panel !== b.dataset.tab; });
      $('#i-msg').innerHTML = '';
    }));
    $('[data-panel=ingresar]').addEventListener('submit', async (e) => {
      e.preventDefault();
      const b = $('button', e.target); b.disabled = true;
      const { error } = await sb.auth.signInWithPassword({ email: $('#i-correo').value.trim(), password: $('#i-clave').value });
      b.disabled = false;
      if (error) { msg(esc(mensajeError(error)), 'error'); return; }
      // onAuthStateChange se encarga de redirigir
    });
    $('[data-panel=registro]').addEventListener('submit', async (e) => {
      e.preventDefault();
      const nombre = $('#r-nombre').value.trim(), correo = $('#r-correo').value.trim(), clave = $('#r-clave').value;
      if (!nombre || !correo || clave.length < 8) { msg('Completa tu nombre, correo y una contraseña de al menos 8 caracteres.', 'error'); return; }
      const b = $('button', e.target); b.disabled = true;
      const { data, error } = await sb.auth.signUp({ email: correo, password: clave, options: { data: { nombre_completo: nombre }, emailRedirectTo: location.origin + location.pathname } });
      b.disabled = false;
      if (error) { msg(esc(mensajeError(error)), 'error'); return; }
      if (!data.session) msg(`<strong>Revisa tu correo.</strong> Te enviamos un enlace a ${esc(correo)} para confirmar la cuenta. Después, un administrador la aprobará.`, 'ok');
    });
    $('[data-panel=olvido]').addEventListener('submit', async (e) => {
      e.preventDefault();
      const { error } = await sb.auth.resetPasswordForEmail($('#o-correo').value.trim(), { redirectTo: location.origin + location.pathname });
      if (error) { msg(esc(mensajeError(error)), 'error'); return; }
      msg('Si hay una cuenta con ese correo, te llegará un enlace para elegir una nueva contraseña.', 'ok');
    });
  }

  function vistaNuevaClave() {
    if (!uid()) { location.hash = '#cuenta'; return; }
    mostrar(`<div class="hoja"><h1>Elige una nueva contraseña</h1>
      <form id="form-nueva" class="caja" novalidate><div class="rejilla"><div><label for="nc-clave">Nueva contraseña</label>
        <input type="password" id="nc-clave" minlength="8" autocomplete="new-password"><span class="ayuda">Al menos 8 caracteres</span></div></div>
        <p class="fila-botones" style="margin-top:1rem"><button type="submit" class="boton">Guardar contraseña</button></p></form></div>`);
    $('#form-nueva').addEventListener('submit', async (e) => {
      e.preventDefault();
      const clave = $('#nc-clave').value;
      if (clave.length < 8) { aviso('La contraseña debe tener al menos 8 caracteres', 'error'); return; }
      const { error } = await sb.auth.updateUser({ password: clave });
      if (error) { aviso(mensajeError(error), 'error'); return; }
      aviso('Contraseña guardada');
      location.hash = '#cuenta';
    });
  }

  // ===================================================================
  // Mis registros
  // ===================================================================
  async function vistaMisRegistros(turno) {
    const v = mostrar('<div class="hoja-ancha"><p class="suave">Cargando…</p></div>', turno);
    const [sitios, cronos] = await Promise.all([
      traerTodo(() => sb.from('sitios').select('id, codigo, nombre, region, actualizado_en, cronologias(count)').eq('creado_por', uid()).order('codigo')),
      traerTodo(() => sb.from('cronologias').select('id, codigo, anio_inicio, anio_fin, disponibilidad, actualizado_en, sitio_id, sitios(codigo, nombre), especies(nombre_cientifico)').eq('creado_por', uid()).order('codigo'))
    ]);
    if (!vigente(turno)) return;
    v.innerHTML = `<div class="hoja-ancha">
      <div class="encabezado-vista"><div><h1>Mis registros</h1><p class="suave">${plural(sitios.length, 'sitio', 'sitios')} y ${plural(cronos.length, 'cronología', 'cronologías')} registrados por ti.</p></div>
        <p class="fila-botones"><a class="boton" href="#sitio/nuevo">Nuevo sitio</a><a class="boton secundario" href="#importar">Importar</a></p></div>
      <section class="caja"><h2>Sitios</h2>
        ${sitios.length ? `<div class="tabla-envoltura"><table><thead><tr><th>Código</th><th>Nombre</th><th>Región</th><th class="num">Cronologías</th><th>Modificado</th><th></th></tr></thead><tbody>
          ${sitios.map((s) => `<tr><td><a href="#sitio/${s.id}"><strong>${esc(s.codigo)}</strong></a></td><td>${esc(s.nombre)}</td><td>${esc(s.region)}</td>
            <td class="num">${s.cronologias && s.cronologias[0] ? s.cronologias[0].count : 0}</td><td>${fecha(s.actualizado_en)}</td>
            <td><a class="boton secundario chico" href="#sitio/${s.id}/editar">Editar</a></td></tr>`).join('')}</tbody></table></div>`
        : '<p class="suave">Aún no registras sitios. <a href="#sitio/nuevo">Registra el primero</a> o <a href="#importar">importa una tabla</a>.</p>'}
      </section>
      <section class="caja"><h2>Cronologías</h2>
        ${cronos.length ? `<div class="tabla-envoltura"><table><thead><tr><th>Código</th><th>Sitio</th><th>Especie</th><th>Período</th><th>Serie</th><th>Modificado</th><th></th></tr></thead><tbody>
          ${cronos.map((c) => `<tr><td><strong>${esc(c.codigo)}</strong></td><td><a href="#sitio/${c.sitio_id}">${esc(c.sitios && c.sitios.codigo)}</a></td>
            <td><em>${esc(c.especies && c.especies.nombre_cientifico)}</em></td><td class="num">${periodo(c.anio_inicio, c.anio_fin)}</td>
            <td>${ETIQ.disponibilidadCorta[c.disponibilidad]}</td><td>${fecha(c.actualizado_en)}</td>
            <td><a class="boton secundario chico" href="#cronologia/${c.id}/editar">Editar</a></td></tr>`).join('')}</tbody></table></div>`
        : '<p class="suave">Aún no registras cronologías.</p>'}
      </section></div>`;
  }

  // ===================================================================
  // Administración
  // ===================================================================
  async function vistaAdmin(turno) {
    if (!puede.curar()) {
      mostrar('<div class="hoja"><h1>Sin acceso</h1><p>Esta sección es para curadores y administradores.</p></div>', turno);
      return;
    }
    const pestanas = [];
    if (puede.administrar()) pestanas.push(['cuentas', 'Cuentas']);
    pestanas.push(['laboratorios', 'Laboratorios'], ['especies', 'Especies'], ['importaciones', 'Importaciones'], ['historial', 'Historial']);
    const v = mostrar(`<div class="hoja-ancha">
      <div class="encabezado-vista"><div><h1>Administrar</h1></div></div>
      <div class="pestanas" role="tablist">${pestanas.map(([k, t], i) => `<button type="button" role="tab" data-tab="${k}" aria-selected="${i === 0}">${t}</button>`).join('')}</div>
      <div id="admin-panel"></div></div>`, turno);
    const ir = (k) => {
      $$('[data-tab]', v).forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === k)));
      ({ cuentas: adminCuentas, laboratorios: adminLaboratorios, especies: adminEspecies,
         importaciones: adminImportaciones, historial: adminHistorial })[k]();
    };
    $$('[data-tab]', v).forEach((b) => b.addEventListener('click', () => ir(b.dataset.tab)));
    ir(pestanas[0][0]);
  }

  async function adminCuentas() {
    const panel = $('#admin-panel');
    const { data, error } = await sb.rpc('administrar_cuentas');
    if (!panel.isConnected) return;
    if (error) { panel.innerHTML = `<div class="mensaje error"><p>${esc(mensajeError(error))}</p></div>`; return; }
    const pendientes = data.filter((c) => !c.aprobado).length;
    panel.innerHTML = `${pendientes ? `<div class="mensaje aviso"><p>${pendientes === 1 ? 'Hay una cuenta pendiente' : 'Hay ' + pendientes + ' cuentas pendientes'} de aprobación.</p></div>` : ''}
      <div class="tabla-envoltura"><table><thead><tr><th>Nombre</th><th>Correo</th><th>Laboratorio</th><th>Creada</th><th>Rol</th><th>Estado</th></tr></thead><tbody>
      ${data.map((c) => {
        const propio = c.id === uid();
        return `<tr><td>${esc(c.nombre_completo) || '<span class="suave">Sin nombre</span>'}${c.orcid ? '<br><span class="suave chico">' + esc(c.orcid) + '</span>' : ''}</td>
          <td><a href="mailto:${esc(c.correo)}">${esc(c.correo)}</a></td><td>${esc(c.laboratorio) || '—'}</td><td>${fecha(c.creado_en)}</td>
          <td><select data-rol="${c.id}" ${propio ? 'disabled title="No puedes cambiar tu propio rol"' : ''} aria-label="Rol de ${esc(c.nombre_completo || c.correo)}">${opciones(ETIQ.rol, c.rol)}</select></td>
          <td>${c.aprobado
            ? `<span class="etiqueta">Aprobada</span> ${propio ? '' : `<button type="button" class="boton-texto chico" data-aprobar="${c.id}" data-valor="false">Quitar</button>`}`
            : `<button type="button" class="boton chico" data-aprobar="${c.id}" data-valor="true">Aprobar</button>`}</td></tr>`;
      }).join('')}</tbody></table></div>`;
    $$('[data-aprobar]', panel).forEach((b) => b.addEventListener('click', async () => {
      const valor = b.dataset.valor === 'true';
      if (!valor && !confirm('¿Quitar la aprobación? La persona ya no podrá registrar ni editar.')) return;
      const { error } = await sb.from('perfiles').update({ aprobado: valor }).eq('id', b.dataset.aprobar);
      if (error) { aviso(mensajeError(error), 'error'); return; }
      aviso(valor ? 'Cuenta aprobada' : 'Aprobación quitada');
      adminCuentas();
    }));
    $$('[data-rol]', panel).forEach((s) => s.addEventListener('change', async () => {
      const { error } = await sb.from('perfiles').update({ rol: s.value }).eq('id', s.dataset.rol);
      if (error) { aviso(mensajeError(error), 'error'); adminCuentas(); return; }
      aviso('Rol cambiado a ' + ETIQ.rol[s.value].toLowerCase());
    }));
  }

  async function adminLaboratorios() {
    const panel = $('#admin-panel');
    const labs = await traerTodo(() => sb.from('laboratorios').select('*').order('nombre'));
    if (!panel.isConnected) return;
    panel.innerHTML = `<form id="f-lab-nuevo" class="caja" novalidate><h2>Agregar laboratorio</h2>
        <div class="rejilla rejilla-3"><div><label for="l-nombre" class="requerido">Nombre</label><input type="text" id="l-nombre"></div>
          <div><label for="l-inst">Institución</label><input type="text" id="l-inst"></div><div><label for="l-ciudad">Ciudad</label><input type="text" id="l-ciudad"></div></div>
        <p class="fila-botones" style="margin-top:1rem"><button type="submit" class="boton">Agregar laboratorio</button></p></form>
      <div class="tabla-envoltura"><table><thead><tr><th>Nombre</th><th>Institución</th><th>Ciudad</th><th></th></tr></thead><tbody>
      ${labs.map((l) => `<tr data-lab="${l.id}"><td><input type="text" data-k="nombre" value="${esc(l.nombre)}" aria-label="Nombre"></td>
        <td><input type="text" data-k="institucion" value="${esc(l.institucion)}" aria-label="Institución"></td><td><input type="text" data-k="ciudad" value="${esc(l.ciudad)}" aria-label="Ciudad"></td>
        <td style="white-space:nowrap"><button type="button" class="boton secundario chico" data-guardar>Guardar</button> <button type="button" class="boton peligro chico" data-borrar>Borrar</button></td></tr>`).join('')
        || '<tr><td colspan="4" class="suave">Aún no hay laboratorios.</td></tr>'}</tbody></table></div>`;
    $('#f-lab-nuevo', panel).addEventListener('submit', async (e) => {
      e.preventDefault();
      const nombre = $('#l-nombre', panel).value.trim();
      if (!nombre) { aviso('Escribe el nombre del laboratorio', 'error'); return; }
      const { error } = await sb.from('laboratorios').insert({ nombre, institucion: vacioANulo($('#l-inst', panel).value), ciudad: vacioANulo($('#l-ciudad', panel).value) });
      if (error) { aviso(mensajeError(error), 'error'); return; }
      estado.cargado = false;
      aviso('Laboratorio agregado');
      adminLaboratorios();
    });
    $$('[data-lab]', panel).forEach((tr) => {
      $('[data-guardar]', tr).addEventListener('click', async () => {
        const fila = {};
        $$('[data-k]', tr).forEach((i) => { fila[i.dataset.k] = vacioANulo(i.value); });
        if (!fila.nombre) { aviso('El nombre no puede quedar vacío', 'error'); return; }
        const { error } = await sb.from('laboratorios').update(fila).eq('id', tr.dataset.lab);
        if (error) { aviso(mensajeError(error), 'error'); return; }
        estado.cargado = false;
        aviso('Laboratorio guardado');
      });
      $('[data-borrar]', tr).addEventListener('click', async () => {
        if (!confirm('¿Borrar este laboratorio? Los sitios y perfiles que lo usan quedarán sin laboratorio.')) return;
        const { error } = await sb.from('laboratorios').delete().eq('id', tr.dataset.lab);
        if (error) { aviso(mensajeError(error), 'error'); return; }
        estado.cargado = false;
        adminLaboratorios();
      });
    });
  }

  async function adminEspecies() {
    const panel = $('#admin-panel');
    const esp = await traerTodo(() => sb.from('especies').select('*').order('nombre_cientifico'));
    if (!panel.isConnected) return;
    panel.innerHTML = `<p class="suave">Corrige nombres y códigos. Una especie solo se puede borrar si ninguna cronología la usa.</p>
      <div class="tabla-envoltura"><table><thead><tr><th>Nombre científico</th><th>Nombre común</th><th>Familia</th><th>Código ITRDB</th><th></th></tr></thead><tbody>
      ${esp.map((e) => `<tr data-esp="${e.id}"><td><input type="text" data-k="nombre_cientifico" value="${esc(e.nombre_cientifico)}" aria-label="Nombre científico"></td>
        <td><input type="text" data-k="nombre_comun" value="${esc(e.nombre_comun)}" aria-label="Nombre común"></td><td><input type="text" data-k="familia" value="${esc(e.familia)}" aria-label="Familia"></td>
        <td><input type="text" data-k="codigo_itrdb" value="${esc(e.codigo_itrdb ? e.codigo_itrdb.trim() : '')}" maxlength="4" aria-label="Código ITRDB" style="max-width:6rem"></td>
        <td style="white-space:nowrap"><button type="button" class="boton secundario chico" data-guardar>Guardar</button> <button type="button" class="boton peligro chico" data-borrar>Borrar</button></td></tr>`).join('')}
      </tbody></table></div>`;
    $$('[data-esp]', panel).forEach((tr) => {
      $('[data-guardar]', tr).addEventListener('click', async () => {
        const fila = {};
        $$('[data-k]', tr).forEach((i) => { fila[i.dataset.k] = vacioANulo(i.value); });
        if (!fila.nombre_cientifico) { aviso('El nombre científico no puede quedar vacío', 'error'); return; }
        if (fila.codigo_itrdb) fila.codigo_itrdb = fila.codigo_itrdb.toUpperCase();
        const { error } = await sb.from('especies').update(fila).eq('id', tr.dataset.esp);
        if (error) { aviso(mensajeError(error), 'error'); return; }
        estado.cargado = false;
        aviso('Especie guardada');
      });
      $('[data-borrar]', tr).addEventListener('click', async () => {
        if (!confirm('¿Borrar esta especie?')) return;
        const { error } = await sb.from('especies').delete().eq('id', tr.dataset.esp);
        if (error) { aviso(mensajeError(error), 'error'); return; }
        estado.cargado = false;
        adminEspecies();
      });
    });
  }

  // Deshacer una importación completa. Borrar 60 sitios de a uno desde su ficha
  // es inviable, y una importación mal armada (columnas corridas, autoría
  // equivocada) se arregla antes borrándola y volviéndola a subir.
  async function adminImportaciones() {
    const panel = $('#admin-panel');
    panel.innerHTML = '<p class="suave">Cargando…</p>';
    let imps, sitios, cronos;
    try {
      [imps, sitios, cronos] = await Promise.all([
        traerTodo(() => sb.from('importaciones').select('*, perfiles(nombre_completo)').order('id', { ascending: false })),
        traerTodo(() => sb.from('sitios').select('id, importacion_id')),
        traerTodo(() => sb.from('cronologias').select('id, sitio_id, importacion_id'))
      ]);
    } catch (e) { panel.innerHTML = `<div class="mensaje error"><p>${esc(mensajeError(e))}</p></div>`; return; }
    if (!panel.isConnected) return;

    // Qué sigue existiendo de cada importación, y qué se llevaría de paso:
    // borrar un sitio arrastra todas sus cronologías, incluso las que otra
    // persona le agregó después.
    const resumen = (id) => {
      const ids = new Set(sitios.filter((s) => s.importacion_id === id).map((s) => s.id));
      return {
        sitios: ids.size,
        cronos: cronos.filter((c) => c.importacion_id === id).length,
        colateral: cronos.filter((c) => c.importacion_id !== id && ids.has(c.sitio_id)).length
      };
    };

    panel.innerHTML = `<p class="suave">Cada importación guarda qué archivo se subió y qué columna fue a qué campo.
        Borrar sus datos elimina los sitios y las cronologías que creó; el registro de la importación queda como constancia.</p>
      <div class="tabla-envoltura"><table><thead><tr><th>Fecha</th><th>Archivo</th><th>Quién</th>
        <th class="num">Filas</th><th class="num">Sitios</th><th class="num">Crono.</th><th></th></tr></thead><tbody>
      ${imps.map((i) => {
        const r = resumen(i.id);
        const vacia = !r.sitios && !r.cronos;
        return `<tr data-imp="${i.id}"><td style="white-space:nowrap">${fechaHora(i.creado_en)}</td>
          <td>${esc(i.archivo)}<br><span class="suave chico">${esc(i.formato)}</span></td>
          <td>${esc(i.perfiles ? i.perfiles.nombre_completo : '') || '<span class="suave">—</span>'}</td>
          <td class="num">${i.n_filas === null ? '—' : nf(i.n_filas)}</td>
          <td class="num">${nf(r.sitios)}</td><td class="num">${nf(r.cronos)}</td>
          <td>${vacia
            ? '<span class="suave chico">Sin datos</span>'
            : `<button type="button" class="boton peligro chico" data-borrar>Borrar datos</button>${
                r.colateral ? `<br><span class="suave chico">+${plural(r.colateral, 'cronología ajena', 'cronologías ajenas')}</span>` : ''}`}</td></tr>`;
      }).join('') || '<tr><td colspan="7" class="suave">Todavía no has importado archivos.</td></tr>'}
      </tbody></table></div>`;

    $$('[data-imp]', panel).forEach((tr) => {
      const b = $('[data-borrar]', tr);
      if (!b) return;
      const id = Number(tr.dataset.imp);
      const imp = imps.find((i) => i.id === id);
      const r = resumen(id);
      b.addEventListener('click', async () => {
        const aviso1 = r.colateral
          ? '\n\nATENCIÓN: ' + plural(r.colateral, 'cronología', 'cronologías') + ' que no viene' + (r.colateral === 1 ? '' : 'n')
            + ' de esta importación cuelga' + (r.colateral === 1 ? '' : 'n') + ' de estos sitios y también se borrará'
            + (r.colateral === 1 ? '' : 'n') + '.'
          : '';
        if (!confirm('¿Borrar los datos de «' + imp.archivo + '»?\n\nSe borrarán '
          + plural(r.cronos, 'cronología', 'cronologías') + ' y ' + plural(r.sitios, 'sitio', 'sitios') + '.'
          + aviso1 + '\n\nEsto no se puede deshacer.')) return;
        b.disabled = true;
        // Las cronologías van primero: algunas pueden estar en sitios que ya
        // existían antes de esta importación y que no hay que borrar.
        const { error: eC } = await sb.from('cronologias').delete().eq('importacion_id', id);
        if (eC) { aviso(mensajeError(eC), 'error'); b.disabled = false; return; }
        const { error: eS } = await sb.from('sitios').delete().eq('importacion_id', id);
        if (eS) { aviso(mensajeError(eS), 'error'); b.disabled = false; return; }
        estado.cargado = false;
        aviso('Datos de la importación borrados');
        adminImportaciones();
      });
    });
  }

  async function adminHistorial() {
    const panel = $('#admin-panel');
    const { data, error } = await sb.from('historial').select('*').order('id', { ascending: false }).limit(150);
    if (error) { panel.innerHTML = `<div class="mensaje error"><p>${esc(mensajeError(error))}</p></div>`; return; }
    const { data: perfiles } = await sb.from('perfiles').select('id, nombre_completo');
    if (!panel.isConnected) return;
    const nombres = new Map((perfiles || []).map((p) => [p.id, p.nombre_completo]));
    const IGNORAR = ['actualizado_en', 'creado_en'];
    const cambios = (h) => {
      if (h.accion !== 'modificar') {
        const d = h.datos_despues || h.datos_antes || {};
        if (h.tabla === 'cronologia_autores') return esc((d.nombre_autor || nombres.get(d.perfil_id) || '') + ' (autor ' + d.orden + ')');
        return esc(d.codigo || '');
      }
      const a = h.datos_antes || {}, d = h.datos_despues || {};
      return Object.keys(d).filter((k) => !IGNORAR.includes(k) && JSON.stringify(a[k]) !== JSON.stringify(d[k]))
        .map((k) => `<strong>${esc(k)}</strong>: ${esc(a[k] === null ? '—' : a[k])} → ${esc(d[k] === null ? '—' : d[k])}`).join('<br>');
    };
    panel.innerHTML = `<p class="suave">Últimos 150 cambios en sitios, cronologías y autores.</p>
      <div class="tabla-envoltura"><table><thead><tr><th>Fecha</th><th>Persona</th><th>Acción</th><th>Registro</th><th>Detalle</th></tr></thead><tbody>
      ${(data || []).map((h) => `<tr><td style="white-space:nowrap">${fechaHora(h.fecha)}</td><td>${esc(nombres.get(h.usuario) || (h.usuario ? 'Usuario borrado' : 'Base de datos'))}</td>
        <td>${ETIQ.accion[h.accion]} ${ETIQ.tabla[h.tabla].toLowerCase()}</td>
        <td>${h.tabla === 'sitios' && h.accion !== 'borrar' ? `<a href="#sitio/${esc(h.registro)}">n.º ${esc(h.registro)}</a>` : 'n.º ' + esc(h.registro)}</td>
        <td class="chico">${cambios(h)}</td></tr>`).join('') || '<tr><td colspan="5" class="suave">Sin cambios registrados.</td></tr>'}</tbody></table></div>`;
  }

  // ===================================================================
  // Contacto
  // ===================================================================
  // El formulario va a Web3Forms, que reenvía el mensaje por correo. Así nadie
  // tiene que publicar su dirección en la página.
  function vistaContacto() {
    const clave = CFG.web3formsClave;
    if (!clave) {
      mostrar(`<div class="hoja"><h1>Contacto</h1>
        <div class="mensaje aviso"><p>Falta configurar el formulario. Escribe la clave de Web3Forms en <strong>js/config.js</strong>.</p></div></div>`);
      return;
    }
    mostrar(`<div class="hoja"><h1>Contacto</h1>
      <p>¿Tienes una consulta sobre el catálogo, quieres sumar tu laboratorio o encontraste un error? Escríbenos y te responderemos por correo.</p>
      <p class="suave">Si lo que quieres es registrar sitios, no hace falta escribir aquí: <a href="#cuenta">crea una cuenta</a> y un administrador la aprobará.</p>
      <form id="form-contacto" class="caja" novalidate>
        <div class="rejilla">
          <div><label for="ct-nombre">Tu nombre</label><input type="text" id="ct-nombre" maxlength="120" autocomplete="name"></div>
          <div><label for="ct-correo">Tu correo</label><input type="email" id="ct-correo" maxlength="160" autocomplete="email"><span class="ayuda">Solo lo usamos para responderte</span></div>
        </div>
        <div style="margin-top:.75rem"><label for="ct-mensaje">Mensaje</label><textarea id="ct-mensaje" rows="7" maxlength="4000"></textarea></div>
        <input type="checkbox" id="ct-trampa" style="display:none" tabindex="-1" autocomplete="off" aria-hidden="true">
        <p class="fila-botones" style="margin-top:1rem"><button type="submit" class="boton" id="ct-enviar">Enviar mensaje</button></p>
      </form></div>`);

    $('#form-contacto').addEventListener('submit', async (e) => {
      e.preventDefault();
      if ($('#ct-trampa').checked) return;   // casilla oculta: si viene marcada, es un robot
      const nombre = vacioANulo($('#ct-nombre').value);
      const correo = vacioANulo($('#ct-correo').value);
      const mensaje = vacioANulo($('#ct-mensaje').value);
      if (!nombre || !correo || !mensaje) { aviso('Completa tu nombre, tu correo y el mensaje', 'error'); return; }
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) { aviso('Ese correo no parece válido', 'error'); return; }
      const boton = $('#ct-enviar');
      boton.disabled = true;
      boton.textContent = 'Enviando…';
      try {
        const r = await fetch('https://api.web3forms.com/submit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
          body: JSON.stringify({
            access_key: clave,
            from_name: 'DendroRed',
            subject: 'DendroRed — consulta de ' + nombre,
            replyto: correo,
            name: nombre, email: correo, message: mensaje
          })
        });
        const res = await r.json().catch(() => ({}));
        if (!r.ok || res.success === false) throw new Error(res.message || 'envío rechazado');
        $('#form-contacto').reset();
        aviso('Mensaje enviado. Te responderemos a tu correo.');
      } catch (err) {
        aviso('No se pudo enviar el mensaje. Revisa tu conexión e inténtalo de nuevo.', 'error');
      } finally {
        boton.disabled = false;
        boton.textContent = 'Enviar mensaje';
      }
    });
  }

  // ===================================================================
  // Acerca de
  // ===================================================================
  function vistaAcerca() {
    const hoy = new Date();
    // Las consultas van por el formulario, para no publicar direcciones de correo.
    // Si además quieres mostrar una dirección, llena correoContacto en config.js.
    const correos = [].concat(CFG.correoContacto || []).filter(Boolean);
    const contacto = `<p>¿Dudas o comentarios? Escríbenos desde la <a href="#contacto">página de contacto</a>.`
      + (correos.length ? ' También puedes escribir a ' + correos.map((c) => `<a href="mailto:${esc(c)}">${esc(c)}</a>`).join(' o ') + '.' : '')
      + '</p>';
    mostrar(`<div class="hoja">
      <h1>Acerca de DendroRed</h1>
      <p>DendroRed es un catálogo colaborativo de los sitios de muestreo y las cronologías dendrocronológicas desarrolladas por los laboratorios de Chile.
        Su objetivo es saber qué se ha muestreado, dónde y por quién, y reconocer el trabajo de cada investigador, aunque la serie de tiempo no esté publicada.</p>
      <div class="caja"><h2>Cómo participar</h2>
        <p>Cualquier persona puede consultar el mapa y exportar los datos. Para registrar sitios hay que crear una cuenta, que un administrador aprueba.
          Cada investigador edita solo sus propios registros, y puede agregar su cronología a un sitio que ya registró otra persona.</p>
        <p>La serie de tiempo no es obligatoria. Si está "a pedido", los usuarios registrados ven el correo del responsable para solicitarla.</p></div>
      <div class="caja"><h2>Cómo citar</h2>
        <p class="cita-bloque">DendroRed (${hoy.getFullYear()}). Catálogo colaborativo de sitios y cronologías dendrocronológicas. ${esc(location.origin + location.pathname)} (consultado el ${hoy.toLocaleDateString('es-CL', { day: 'numeric', month: 'long', year: 'numeric' })}).</p>
        <p>Si usas una cronología en particular, cita también a sus autores y su publicación, que aparecen en la ficha del sitio.</p></div>
      <div class="caja"><h2>Coordenadas</h2>
        <p>Todas las coordenadas se guardan en WGS84 y grados decimales. Las que llegan en PSAD56 o SAD69 se convierten con parámetros generales para Sudamérica, con un error residual de decenas de metros.
          Algunos sitios se publican con coordenadas aproximadas (~1 km) para proteger especies o lugares vulnerables.</p></div>
      ${contacto}</div>`);
  }

  // ===================================================================
  // Arranque
  // ===================================================================
  function sinConfigurar() {
    pintarNavMinima();
    mostrar(`<div class="hoja"><h1>Falta conectar la base de datos</h1>
      <div class="mensaje aviso"><p>Edita el archivo <strong>js/config.js</strong> y escribe la dirección de tu proyecto de Supabase y su clave pública. Los pasos están en el archivo LEEME.md.</p></div></div>`);
  }
  function pintarNavMinima() { $('#nav').innerHTML = ''; }

  async function iniciar() {
    if (!window.supabase || !window.L) {
      mostrar('<div class="hoja"><div class="mensaje error"><p>No se pudieron cargar las bibliotecas de la página. Revisa tu conexión y recarga.</p></div></div>');
      return;
    }
    if (!CFG.supabaseUrl || /TU-PROYECTO/.test(CFG.supabaseUrl) || !CFG.supabaseClave || /TU-CLAVE/.test(CFG.supabaseClave)) { sinConfigurar(); return; }
    sb = window.supabase.createClient(CFG.supabaseUrl, CFG.supabaseClave);
    const { data } = await sb.auth.getSession();
    estado.sesion = data.session;
    await cargarPerfil();
    sb.auth.onAuthStateChange((evento, sesion) => {
      // No se llama a la base dentro de este aviso (recomendación de Supabase): se difiere.
      setTimeout(async () => {
        const antes = uid();
        estado.sesion = sesion;
        if (evento === 'PASSWORD_RECOVERY') { await cargarPerfil(); location.hash = '#nueva-clave'; return; }
        if ((uid() || null) === (antes || null) && evento !== 'USER_UPDATED') return;
        await cargarPerfil();
        estado.cargado = false;
        if (evento === 'SIGNED_IN' && /^#cuenta$|^#?$/.test(location.hash)) {
          const volver = sessionStorage.getItem('dr-volver');
          sessionStorage.removeItem('dr-volver');
          aviso('Sesión iniciada');
          if (volver) { location.hash = volver; return; }
        }
        enrutar();
      }, 0);
    });
    window.addEventListener('hashchange', enrutar);
    enrutar();
  }

  iniciar();
})();
