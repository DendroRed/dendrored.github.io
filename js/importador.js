/* DendroRed — lectura de tablas y archivos GPS, asignación de columnas y validación.
 * Formatos: CSV (coma, punto y coma o tabulaciones), TXT, TSV, Excel (.xlsx/.xls),
 * GPX, KML y GeoJSON. Excel requiere SheetJS (XLSX) cargado desde el CDN.
 */
(function (raiz) {
  'use strict';

  // ---------- Campos que se pueden importar ----------
  // grupo: 'sitio' o 'crono'. sinonimos: encabezados habituales (sin tildes, en minúscula).
  const CAMPOS = [
    { clave: 'codigo', etiqueta: 'Código del sitio', grupo: 'sitio', requerido: true,
      sinonimos: ['codigo', 'cod', 'codigositio', 'sitecode', 'code', 'site', 'id', 'idsitio', 'sitio_id'] },
    { clave: 'nombre', etiqueta: 'Nombre del sitio', grupo: 'sitio',
      sinonimos: ['nombre', 'nombresitio', 'sitio', 'sitename', 'name', 'localidad'] },
    { clave: 'latitud', etiqueta: 'Latitud', grupo: 'sitio', coord: 'geo',
      sinonimos: ['latitud', 'lat', 'latitude', 'y', 'latdd', 'latdec', 'lat_s', 'lats'] },
    { clave: 'longitud', etiqueta: 'Longitud', grupo: 'sitio', coord: 'geo',
      sinonimos: ['longitud', 'lon', 'long', 'lng', 'longitude', 'x', 'londd', 'londec', 'lon_w', 'lonw'] },
    { clave: 'este', etiqueta: 'UTM este', grupo: 'sitio', coord: 'utm',
      sinonimos: ['este', 'easting', 'utme', 'utmeste', 'e', 'x_utm', 'xutm', 'coordx'] },
    { clave: 'norte', etiqueta: 'UTM norte', grupo: 'sitio', coord: 'utm',
      sinonimos: ['norte', 'northing', 'utmn', 'utmnorte', 'n', 'y_utm', 'yutm', 'coordy'] },
    { clave: 'zona', etiqueta: 'UTM zona', grupo: 'sitio', coord: 'utm',
      sinonimos: ['zona', 'zone', 'huso', 'zonautm', 'utmzone'] },
    { clave: 'altitud_m', etiqueta: 'Altitud (m)', grupo: 'sitio',
      sinonimos: ['altitud', 'alt', 'altura', 'elevacion', 'elev', 'elevation', 'ele', 'msnm', 'altitudm'] },
    { clave: 'region', etiqueta: 'Región', grupo: 'sitio', sinonimos: ['region', 'reg'] },
    { clave: 'comuna', etiqueta: 'Comuna', grupo: 'sitio', sinonimos: ['comuna', 'municipio'] },
    { clave: 'localidad', etiqueta: 'Localidad', grupo: 'sitio', sinonimos: ['localidad', 'lugar', 'sector', 'location'] },
    { clave: 'pais', etiqueta: 'País', grupo: 'sitio', sinonimos: ['pais', 'country'] },
    { clave: 'ambiente', etiqueta: 'Ambiente', grupo: 'sitio', sinonimos: ['ambiente', 'habitat', 'tipobosque', 'bosque', 'vegetacion'] },
    { clave: 'descripcion', etiqueta: 'Descripción del sitio', grupo: 'sitio', sinonimos: ['descripcion', 'description', 'comentario', 'obs', 'observaciones'] },
    { clave: 'laboratorio', etiqueta: 'Laboratorio del sitio', grupo: 'sitio',
      sinonimos: ['laboratorio', 'lab', 'laboratory', 'institucion', 'universidad', 'centro', 'laboratoriositio'] },

    { clave: 'especie', etiqueta: 'Especie', grupo: 'crono',
      sinonimos: ['especie', 'species', 'sp', 'spp', 'taxon', 'codigoespecie', 'speciescode'] },
    { clave: 'codigo_crono', etiqueta: 'Código de la cronología', grupo: 'crono',
      sinonimos: ['codigocronologia', 'cronologia', 'chronology', 'chronocode', 'codigocrono', 'crono'] },
    { clave: 'parametro', etiqueta: 'Parámetro', grupo: 'crono', sinonimos: ['parametro', 'parameter', 'variable', 'tipo'] },
    { clave: 'estado', etiqueta: 'Estado', grupo: 'crono', sinonimos: ['estado', 'status'] },
    { clave: 'anio_inicio', etiqueta: 'Año inicial', grupo: 'crono',
      sinonimos: ['anioinicio', 'anoinicio', 'anioinicial', 'anoinicial', 'inicio', 'desde', 'firstyear', 'start', 'startyear', 'primerano', 'from'] },
    { clave: 'anio_fin', etiqueta: 'Año final', grupo: 'crono',
      sinonimos: ['aniofin', 'anofin', 'aniofinal', 'anofinal', 'fin', 'hasta', 'lastyear', 'end', 'endyear', 'ultimoano', 'to'] },
    { clave: 'anio_eps_085', etiqueta: 'Año con EPS ≥ 0,85', grupo: 'crono',
      sinonimos: ['anioeps', 'anoeps', 'anioeps085', 'anoeps085', 'eps085', 'epsyear', 'epsdesde'] },
    { clave: 'n_arboles', etiqueta: 'N.º de árboles', grupo: 'crono',
      sinonimos: ['narboles', 'arboles', 'ntrees', 'trees', 'n_arb', 'narb'] },
    { clave: 'n_radios', etiqueta: 'N.º de radios', grupo: 'crono',
      sinonimos: ['nradios', 'radios', 'nseries', 'series', 'ncores', 'cores', 'nrad'] },
    { clave: 'intercorrelacion', etiqueta: 'Intercorrelación', grupo: 'crono',
      sinonimos: ['intercorrelacion', 'interseriescorrelation', 'intercorr', 'r', 'correlacion', 'seriesintercorrelation'] },
    { clave: 'sensibilidad_media', etiqueta: 'Sensibilidad media', grupo: 'crono',
      sinonimos: ['sensibilidad', 'sensibilidadmedia', 'meansensitivity', 'ms', 'sens'] },
    { clave: 'rbar', etiqueta: 'rbar', grupo: 'crono', sinonimos: ['rbar', 'rbt', 'rbartot'] },
    { clave: 'eps', etiqueta: 'EPS', grupo: 'crono', sinonimos: ['eps', 'expressedpopulationsignal'] },
    { clave: 'estandarizacion', etiqueta: 'Estandarización', grupo: 'crono', sinonimos: ['estandarizacion', 'detrending', 'standardization', 'metodo'] },
    { clave: 'anio_muestreo', etiqueta: 'Año de muestreo', grupo: 'crono', sinonimos: ['aniomuestreo', 'muestreo', 'samplingyear', 'anomuestreo', 'colecta', 'aniocolecta'] },
    { clave: 'disponibilidad', etiqueta: 'Disponibilidad de la serie', grupo: 'crono', sinonimos: ['disponibilidad', 'availability', 'serie', 'datos'] },
    { clave: 'enlace_serie', etiqueta: 'Enlace a la serie', grupo: 'crono', sinonimos: ['enlace', 'link', 'url', 'enlaceserie'] },
    { clave: 'doi', etiqueta: 'DOI', grupo: 'crono', sinonimos: ['doi'] },
    { clave: 'cita', etiqueta: 'Cita', grupo: 'crono', sinonimos: ['cita', 'referencia', 'reference', 'citation', 'publicacion'] },
    { clave: 'notas', etiqueta: 'Notas de la cronología', grupo: 'crono', sinonimos: ['notas', 'notes', 'nota'] },
    { clave: 'laboratorio_crono', etiqueta: 'Laboratorio de la cronología', grupo: 'crono',
      sinonimos: ['laboratoriocronologia', 'laboratoriocrono', 'labcronologia', 'labcrono', 'institucioncronologia'] },
    { clave: 'responsable', etiqueta: 'Responsable de la cronología', grupo: 'crono',
      sinonimos: ['responsable', 'autorprincipal', 'principal', 'contacto', 'encargado', 'pi', 'leadauthor', 'corresponding'] },
    { clave: 'autores', etiqueta: 'Autores (separados por ;)', grupo: 'crono',
      sinonimos: ['autores', 'autor', 'authors', 'author', 'coautores', 'colaboradores', 'equipo'] }
  ];

  // Los sinónimos se comparan ya normalizados (sin tildes, espacios ni guiones bajos)
  CAMPOS.forEach((c) => { c.sinonimos = c.sinonimos.map((x) => normalizarBasico(x)); });
  function normalizarBasico(t) { return String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, ''); }

  const PARAMETROS = {
    ancho_total: ['anchototal', 'ancho', 'rw', 'trw', 'ringwidth', 'totalringwidth', 'at', 'total'],
    madera_temprana: ['maderatemprana', 'ew', 'earlywood', 'mt', 'temprana'],
    madera_tardia: ['maderatardia', 'lw', 'latewood', 'tardia'],
    densidad_maxima: ['densidadmaxima', 'mxd', 'densidad', 'maximumdensity'],
    anatomia: ['anatomia', 'anatomy', 'qwa', 'roxas'],
    isotopos: ['isotopos', 'isotopes', 'd13c', 'd18o', 'isotopo'],
    otro: ['otro', 'other']
  };
  const DISPONIBILIDAD = {
    itrdb: ['itrdb', 'noaa', 'ncei'],
    repositorio: ['repositorio', 'zenodo', 'pangaea', 'suplementario', 'repository', 'publica', 'publico'],
    a_pedido: ['apedido', 'pedido', 'consultar', 'contactar', 'onrequest', 'request', 'si'],
    no_disponible: ['nodisponible', 'no', 'notavailable', 'privada', 'privado']
  };
  const ESTADOS = {
    en_curso: ['encurso', 'curso', 'inprogress', 'enproceso', 'proceso'],
    terminada: ['terminada', 'terminado', 'finalizada', 'lista', 'done', 'finished'],
    publicada: ['publicada', 'publicado', 'published']
  };

  // ---------- Utilidades ----------
  function normalizar(texto) {
    return String(texto || '')
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  function codigoValido(texto) {
    return String(texto || '')
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toUpperCase().trim().replace(/\s+/g, '_').replace(/[^A-Z0-9_-]/g, '');
  }

  function buscarEnDiccionario(texto, dic) {
    const n = normalizar(texto);
    if (!n) return null;
    for (const [valor, opciones] of Object.entries(dic)) {
      if (valor.replace(/_/g, '') === n || opciones.includes(n)) return valor;
    }
    return undefined;
  }

  // ---------- Lectura de archivos ----------
  function detectarSeparador(primeraLinea) {
    const cuenta = (c) => primeraLinea.split(c).length - 1;
    const opciones = [['\t', cuenta('\t')], [';', cuenta(';')], [',', cuenta(',')]];
    opciones.sort((a, b) => b[1] - a[1]);
    return opciones[0][1] > 0 ? opciones[0][0] : ',';
  }

  // Lector de CSV con comillas, saltos de línea entre comillas y separador detectado.
  function leerCSV(texto) {
    texto = texto.replace(/^﻿/, '');
    const primera = texto.split(/\r?\n/).find((l) => l.trim() !== '') || '';
    const sep = detectarSeparador(primera);
    const filas = [];
    let fila = [], campo = '', comillas = false;
    for (let i = 0; i < texto.length; i++) {
      const c = texto[i];
      if (comillas) {
        if (c === '"') {
          if (texto[i + 1] === '"') { campo += '"'; i++; } else { comillas = false; }
        } else campo += c;
      } else if (c === '"' && campo === '') comillas = true;   // comillas solo al inicio del campo (así 30"S queda tal cual)
      else if (c === sep) { fila.push(campo); campo = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && texto[i + 1] === '\n') i++;
        fila.push(campo); filas.push(fila); fila = []; campo = '';
      } else campo += c;
    }
    if (campo !== '' || fila.length) { fila.push(campo); filas.push(fila); }
    return tablaAObjetos(filas.filter((f) => f.some((v) => String(v).trim() !== '')));
  }

  function tablaAObjetos(filas) {
    if (!filas.length) return { columnas: [], filas: [] };
    const columnas = [];
    filas[0].forEach((c, i) => {
      let nombre = String(c ?? '').trim() || 'Columna ' + (i + 1);
      while (columnas.includes(nombre)) nombre += '_';
      columnas.push(nombre);
    });
    const objetos = filas.slice(1).map((f) => {
      const o = {};
      columnas.forEach((c, i) => { o[c] = f[i] === undefined || f[i] === null ? '' : f[i]; });
      return o;
    });
    return { columnas, filas: objetos };
  }

  function leerExcel(buffer) {
    if (!raiz.XLSX) throw new Error('No se pudo cargar el lector de Excel');
    const libro = raiz.XLSX.read(buffer, { type: 'array' });
    const hoja = libro.Sheets[libro.SheetNames[0]];
    const filas = raiz.XLSX.utils.sheet_to_json(hoja, { header: 1, raw: true, defval: '' });
    const res = tablaAObjetos(filas.filter((f) => f.some((v) => String(v).trim() !== '')));
    res.hojas = libro.SheetNames;
    return res;
  }

  function leerGPX(texto) {
    const doc = new DOMParser().parseFromString(texto, 'application/xml');
    if (doc.querySelector('parsererror')) throw new Error('El archivo GPX no es válido');
    const puntos = [...doc.getElementsByTagName('wpt')];
    if (!puntos.length) throw new Error('El GPX no tiene puntos guardados (waypoints). Las rutas y tracks no se importan.');
    const filas = puntos.map((p) => ({
      nombre: texto1(p, 'name'), latitud: p.getAttribute('lat'), longitud: p.getAttribute('lon'),
      altitud: texto1(p, 'ele'), descripcion: texto1(p, 'desc') || texto1(p, 'cmt')
    }));
    return { columnas: ['nombre', 'latitud', 'longitud', 'altitud', 'descripcion'], filas };
  }

  function leerKML(texto) {
    const doc = new DOMParser().parseFromString(texto, 'application/xml');
    if (doc.querySelector('parsererror')) throw new Error('El archivo KML no es válido');
    const filas = [];
    for (const pm of doc.getElementsByTagName('Placemark')) {
      const punto = pm.getElementsByTagName('Point')[0];
      if (!punto) continue;
      const [lon, lat, alt] = texto1(punto, 'coordinates').trim().split(/[\s,]+/);
      filas.push({ nombre: texto1(pm, 'name'), latitud: lat, longitud: lon, altitud: alt || '', descripcion: texto1(pm, 'description') });
    }
    if (!filas.length) throw new Error('El KML no tiene puntos. Las líneas y polígonos no se importan.');
    return { columnas: ['nombre', 'latitud', 'longitud', 'altitud', 'descripcion'], filas };
  }

  function leerGeoJSON(texto) {
    const g = JSON.parse(texto);
    const feats = g.type === 'FeatureCollection' ? g.features : g.type === 'Feature' ? [g] : [];
    const filas = [];
    const cols = new Set(['latitud', 'longitud']);
    for (const f of feats) {
      if (!f.geometry || f.geometry.type !== 'Point') continue;
      const [lon, lat, alt] = f.geometry.coordinates;
      const o = Object.assign({}, f.properties || {});
      Object.keys(o).forEach((k) => cols.add(k));
      o.latitud = lat; o.longitud = lon;
      if (alt !== undefined && o.altitud === undefined) { o.altitud = alt; cols.add('altitud'); }
      filas.push(o);
    }
    if (!filas.length) throw new Error('El GeoJSON no tiene puntos');
    return { columnas: [...cols], filas };
  }

  function texto1(nodo, etiqueta) {
    const n = nodo.getElementsByTagName(etiqueta)[0];
    return n ? n.textContent.trim() : '';
  }

  function formatoDe(nombre) {
    const ext = (nombre.split('.').pop() || '').toLowerCase();
    return { csv: 'csv', txt: 'csv', tsv: 'tsv', xlsx: 'xlsx', xls: 'xlsx', gpx: 'gpx', kml: 'kml', geojson: 'geojson', json: 'geojson' }[ext] || null;
  }

  /** Lee un File del navegador. Devuelve { formato, columnas, filas }. */
  async function leerArchivo(archivo) {
    const formato = formatoDe(archivo.name);
    if (!formato) throw new Error('Formato no reconocido. Usa CSV, TXT, TSV, Excel, GPX, KML o GeoJSON.');
    let res;
    if (formato === 'xlsx') res = leerExcel(await archivo.arrayBuffer());
    else {
      const texto = await archivo.text();
      if (formato === 'csv' || formato === 'tsv') res = leerCSV(texto);
      else if (formato === 'gpx') res = leerGPX(texto);
      else if (formato === 'kml') res = leerKML(texto);
      else res = leerGeoJSON(texto);
    }
    if (!res.filas.length) throw new Error('El archivo no tiene filas con datos');
    res.formato = formato;
    return res;
  }

  // ---------- Asignación de columnas ----------
  function adivinarAsignacion(columnas) {
    const asignacion = {};
    const usadas = new Set();
    // Primero coincidencias exactas, después parciales
    for (const exacta of [true, false]) {
      for (const campo of CAMPOS) {
        if (asignacion[campo.clave]) continue;
        const col = columnas.find((c) => {
          if (usadas.has(c)) return false;
          const n = normalizar(c);
          if (exacta) return campo.sinonimos.includes(n) || n === normalizar(campo.etiqueta);
          return n.length > 2 && campo.sinonimos.some((s) => s.length > 3 && n.startsWith(s));
        });
        if (col) { asignacion[campo.clave] = col; usadas.add(col); }
      }
    }
    return asignacion;
  }

  /** Si la columna de latitud parece UTM (valores grandes), sugiere ese formato. */
  function sugerirFormatoCoord(filas, asignacion) {
    if (asignacion.este && asignacion.norte) return 'utm';
    const col = asignacion.latitud;
    if (!col) return 'auto';
    const muestra = filas.slice(0, 20).map((f) => raiz.DRCoord.aNumero(f[col])).filter((v) => !isNaN(v));
    if (muestra.length && muestra.every((v) => Math.abs(v) > 1000)) return 'utm';
    return 'auto';
  }

  // ---------- Validación ----------
  /**
   * opciones: { formatoCoord: 'auto'|'utm', datum, zonaFija, especies: [...], sitiosExistentes: [...] }
   * Devuelve { sitios: Map(codigo → {...}), filas: [{ n, sitio, crono, errores, avisos }] }
   */
  function validar(filas, asignacion, opciones) {
    const C = raiz.DRCoord;
    const resultado = { filas: [], sitios: new Map(), especiesDesconocidas: new Set() };
    const valor = (fila, clave) => {
      const col = asignacion[clave];
      if (!col) return '';
      const v = fila[col];
      return v === null || v === undefined ? '' : String(v).trim();
    };
    const entero = (t, nombre, errores, min, max) => {
      if (t === '') return null;
      const n = C.aNumero(t);
      if (isNaN(n) || !Number.isInteger(n)) { errores.push(nombre + ' debe ser un número entero'); return null; }
      if ((min !== undefined && n < min) || (max !== undefined && n > max)) { errores.push(nombre + ' fuera de rango'); return null; }
      return n;
    };
    const decimal = (t, nombre, errores, min, max) => {
      if (t === '') return null;
      const n = C.aNumero(t);
      if (isNaN(n)) { errores.push(nombre + ' debe ser numérico'); return null; }
      if ((min !== undefined && n < min) || (max !== undefined && n > max)) { errores.push(nombre + ' fuera de rango (' + min + ' a ' + max + ')'); return null; }
      return Math.round(n * 1000) / 1000;
    };

    filas.forEach((fila, i) => {
      const errores = [], avisos = [];
      // --- Sitio ---
      let codigo = codigoValido(valor(fila, 'codigo'));
      if (!asignacion.codigo && valor(fila, 'nombre')) {
        // Sin columna de código: se arma desde el nombre
        codigo = codigoValido(valor(fila, 'nombre')).slice(0, 16);
        avisos.push('sin columna de código: se usa ' + codigo);
      }
      if (!codigo) errores.push('falta el código del sitio');
      else if (!/^[A-Z0-9_-]{2,16}$/.test(codigo)) errores.push('código "' + codigo + '" inválido: 2 a 16 letras, números, _ o -');
      if (codigo && asignacion.codigo && codigo !== valor(fila, 'codigo')) avisos.push('código normalizado a ' + codigo);

      let lat = null, lon = null, coordOriginal = '', formatoOriginal = '';
      const datum = opciones.datum || 'WGS84';
      if (opciones.formatoCoord === 'utm') {
        const e = valor(fila, 'este'), n = valor(fila, 'norte');
        const zona = valor(fila, 'zona') || opciones.zonaFija;
        if (!e || !n) errores.push('faltan coordenadas UTM');
        else if (!zona) errores.push('falta la zona UTM');
        else {
          const r = C.utmAWGS84(e, n, zona, null, datum);
          if (r.error) errores.push('UTM: ' + r.error);
          else { lat = r.lat; lon = r.lon; coordOriginal = zona + ' ' + e + ' ' + n; formatoOriginal = 'utm'; }
        }
      } else {
        const tLat = valor(fila, 'latitud'), tLon = valor(fila, 'longitud');
        const rLat = C.leerCoordenada(tLat, 'lat'), rLon = C.leerCoordenada(tLon, 'lon');
        if (rLat.error) errores.push('latitud ' + rLat.error);
        if (rLon.error) errores.push('longitud ' + rLon.error);
        if (!rLat.error && !rLon.error) {
          const g = C.geograficasAWGS84(rLat.valor, rLon.valor, datum);
          lat = g.lat; lon = g.lon;
          coordOriginal = tLat + ' ' + tLon;
          formatoOriginal = /[°'"NSEWO\s]/i.test(tLat.replace(/^-/, '')) ? 'gms' : 'decimal';
        }
      }
      if (lat !== null && opciones.hemisferioSurOeste && (lat > 0 || lon > 0)) {
        avisos.push('coordenada positiva: revisa el signo (¿falta S u O?)');
      }

      const sitio = {
        codigo,
        nombre: valor(fila, 'nombre') || codigo,
        pais: valor(fila, 'pais') || 'Chile',
        region: valor(fila, 'region') || null,
        comuna: valor(fila, 'comuna') || null,
        localidad: valor(fila, 'localidad') || null,
        latitud: lat !== null ? Math.round(lat * 1e6) / 1e6 : null,
        longitud: lon !== null ? Math.round(lon * 1e6) / 1e6 : null,
        altitud_m: null,
        coord_original: coordOriginal || null,
        formato_original: formatoOriginal || null,
        datum_original: datum,
        ambiente: valor(fila, 'ambiente') || null,
        descripcion: valor(fila, 'descripcion') || null,
        // Nombre tal como viene en el archivo. La app lo traduce a un id de
        // laboratorio antes de guardar, y lo quita de la fila.
        laboratorio_texto: valor(fila, 'laboratorio') || null
      };
      const alt = valor(fila, 'altitud_m');
      if (alt !== '') {
        const a = C.aNumero(alt);
        if (isNaN(a) || a < -500 || a > 7000) errores.push('altitud fuera de rango');
        else sitio.altitud_m = Math.round(a);
      }

      // --- Cronología (solo si hay especie) ---
      let crono = null;
      const tEspecie = valor(fila, 'especie');
      if (tEspecie) {
        const especie = buscarEspecie(tEspecie, opciones.especies || []);
        if (!especie) resultado.especiesDesconocidas.add(tEspecie);
        crono = {
          especie_texto: tEspecie,
          especie_id: especie ? especie.id : null,
          codigo: '',
          parametro: 'ancho_total',
          estado: 'terminada',
          anio_inicio: entero(valor(fila, 'anio_inicio'), 'año inicial', errores, -10000, 2100),
          anio_fin: entero(valor(fila, 'anio_fin'), 'año final', errores, -10000, 2100),
          anio_eps_085: entero(valor(fila, 'anio_eps_085'), 'año EPS', errores, -10000, 2100),
          anio_muestreo: entero(valor(fila, 'anio_muestreo'), 'año de muestreo', errores, 1900, 2100),
          n_arboles: entero(valor(fila, 'n_arboles'), 'n.º de árboles', errores, 1),
          n_radios: entero(valor(fila, 'n_radios'), 'n.º de radios', errores, 1),
          intercorrelacion: decimal(valor(fila, 'intercorrelacion'), 'intercorrelación', errores, -1, 1),
          sensibilidad_media: decimal(valor(fila, 'sensibilidad_media'), 'sensibilidad media', errores, 0, 9.999),
          rbar: decimal(valor(fila, 'rbar'), 'rbar', errores, -1, 1),
          eps: decimal(valor(fila, 'eps'), 'EPS', errores, 0, 1),
          estandarizacion: valor(fila, 'estandarizacion') || null,
          disponibilidad: 'a_pedido',
          enlace_serie: valor(fila, 'enlace_serie') || null,
          doi: valor(fila, 'doi') || null,
          cita: valor(fila, 'cita') || null,
          notas: valor(fila, 'notas') || null,
          // Igual que en el sitio: textos que la app resuelve antes de guardar.
          laboratorio_texto: valor(fila, 'laboratorio_crono') || null,
          responsable_texto: valor(fila, 'responsable') || null,
          autores_texto: valor(fila, 'autores') || null
        };
        const p = valor(fila, 'parametro');
        if (p) {
          const r = buscarEnDiccionario(p, PARAMETROS);
          if (r) crono.parametro = r; else errores.push('parámetro "' + p + '" no reconocido');
        }
        const d = valor(fila, 'disponibilidad');
        if (d) {
          const r = buscarEnDiccionario(d, DISPONIBILIDAD);
          if (r) crono.disponibilidad = r; else errores.push('disponibilidad "' + d + '" no reconocida');
        }
        const est = valor(fila, 'estado');
        if (est) {
          const r = buscarEnDiccionario(est, ESTADOS);
          if (r) crono.estado = r; else errores.push('estado "' + est + '" no reconocido');
        }
        if (crono.anio_inicio !== null && crono.anio_fin !== null && crono.anio_fin < crono.anio_inicio) {
          errores.push('el año final es anterior al inicial');
        }
        if (crono.anio_eps_085 !== null && crono.anio_inicio !== null && crono.anio_fin !== null &&
            (crono.anio_eps_085 < crono.anio_inicio || crono.anio_eps_085 > crono.anio_fin)) {
          errores.push('el año con EPS ≥ 0,85 está fuera del período');
        }
        if (crono.n_radios !== null && crono.n_arboles !== null && crono.n_radios < crono.n_arboles) {
          errores.push('hay menos radios que árboles');
        }
        const sufijo = { ancho_total: 'AT', madera_temprana: 'MT', madera_tardia: 'MTA', densidad_maxima: 'DM', anatomia: 'AN', isotopos: 'IS', otro: 'OT' }[crono.parametro];
        const codEsp = especie && especie.codigo_itrdb ? especie.codigo_itrdb.trim() : codigoValido(tEspecie).slice(0, 4);
        crono.codigo = codigoValido(valor(fila, 'codigo_crono')) || (codigo + '-' + codEsp + '-' + sufijo);
      } else {
        // Sin especie no hay cronología; sus notas pasan a describir el sitio
        if (valor(fila, 'notas') && !sitio.descripcion) sitio.descripcion = valor(fila, 'notas');
        if (CAMPOS.some((c) => c.grupo === 'crono' && !['especie', 'notas'].includes(c.clave) && valor(fila, c.clave) !== '')) {
          avisos.push('hay datos de cronología pero falta la especie: solo se importa el sitio');
        }
      }

      // --- Mismo código repetido en el archivo ---
      if (codigo && !errores.length) {
        const previo = resultado.sitios.get(codigo);
        if (previo) {
          const d = C.distanciaMetros(previo.sitio.latitud, previo.sitio.longitud, sitio.latitud, sitio.longitud);
          if (d > 100) avisos.push('el código ' + codigo + ' se repite con coordenadas a ' + Math.round(d) + ' m: se usan las de la fila ' + previo.fila);
        } else {
          resultado.sitios.set(codigo, { sitio, fila: i + 2 });
        }
      }

      // --- Sitios ya registrados ---
      let existente = null;
      if (codigo && opciones.sitiosExistentes) {
        existente = opciones.sitiosExistentes.find((s) => s.codigo === codigo) || null;
        if (existente) avisos.push('el sitio ' + codigo + ' ya existe: se agregan solo las cronologías');
        else if (lat !== null) {
          const cerca = opciones.sitiosExistentes.find((s) =>
            C.distanciaMetros(lat, lon, Number(s.latitud), Number(s.longitud)) < 500);
          if (cerca) avisos.push('hay otro sitio a menos de 500 m (' + cerca.codigo + '): revisa si es el mismo');
        }
      }

      resultado.filas.push({ n: i + 2, sitio, crono, existente, errores, avisos, original: fila });
    });
    return resultado;
  }

  function buscarEspecie(texto, especies) {
    const n = normalizar(texto);
    return especies.find((e) =>
      normalizar(e.nombre_cientifico) === n ||
      (e.codigo_itrdb && normalizar(e.codigo_itrdb) === n) ||
      (e.nombre_comun && e.nombre_comun.split(',').some((c) => normalizar(c) === n))
    ) || null;
  }

  const api = { CAMPOS, PARAMETROS, DISPONIBILIDAD, ESTADOS, normalizar, codigoValido, leerCSV, tablaAObjetos,
    leerGPX, leerKML, leerGeoJSON, leerArchivo, formatoDe, adivinarAsignacion, sugerirFormatoCoord, validar, buscarEspecie };
  raiz.DRImport = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
