/* DendroRed — lectura y conversión de coordenadas
 * Acepta grados decimales, grados-minutos-segundos y UTM, en WGS84, PSAD56 o SAD69,
 * y entrega siempre latitud/longitud WGS84 en grados decimales.
 * Requiere proj4 (cargado desde el CDN en index.html).
 */
(function (raiz) {
  'use strict';

  // Parámetros de transformación a WGS84 (valores generales de EPSG para Sudamérica).
  // El error residual es de decenas de metros, muy inferior al desfase de
  // cientos de metros que se corrige.
  const DATUMS = {
    WGS84: { etiqueta: 'WGS84 (GPS actuales)', proj: '+datum=WGS84' },
    PSAD56: { etiqueta: 'PSAD56 (cartografía IGM antigua)', proj: '+ellps=intl +towgs84=-288,175,-376,0,0,0,0' },
    SAD69: { etiqueta: 'SAD69', proj: '+ellps=aust_SA +towgs84=-57,1,-41,0,0,0,0' }
  };

  const WGS84 = '+proj=longlat +datum=WGS84 +no_defs';

  function obtenerProj4() {
    if (raiz.proj4) return raiz.proj4;
    if (typeof require === 'function') return require('proj4');
    throw new Error('No se pudo cargar proj4');
  }

  function defGeograficas(datum) {
    const d = DATUMS[datum];
    if (!d) throw new Error('Datum desconocido: ' + datum);
    return '+proj=longlat ' + d.proj + ' +no_defs';
  }

  function defUTM(zona, hemisferio, datum) {
    const d = DATUMS[datum];
    if (!d) throw new Error('Datum desconocido: ' + datum);
    return '+proj=utm +zone=' + zona + (hemisferio === 'N' ? '' : ' +south') + ' ' + d.proj + ' +units=m +no_defs';
  }

  // Convierte texto con coma decimal o signo menos tipográfico a número.
  function aNumero(texto) {
    if (texto === null || texto === undefined) return NaN;
    if (typeof texto === 'number') return texto;
    let t = String(texto).trim().replace(/[−–—]/g, '-').replace(/\s+/g, '');
    if (t === '') return NaN;
    // "1.234,56" → 1234.56 ; "38,6567" → 38.6567
    if (/,\d+$/.test(t) && t.indexOf('.') !== -1 && t.indexOf('.') < t.indexOf(',')) {
      t = t.replace(/\./g, '').replace(',', '.');
    } else {
      t = t.replace(',', '.');
    }
    if (!/^[+-]?(\d+\.?\d*|\.\d+)$/.test(t)) return NaN;
    return parseFloat(t);
  }

  /**
   * Lee una coordenada geográfica en decimal o en grados-minutos-segundos.
   * eje: 'lat' o 'lon'. Devuelve { valor } o { error }.
   * Ejemplos válidos: -38.6567 · 38,6567 S · 38°39'24"S · 38 39 24 S · 71°38.93'O · -71 38 56
   */
  function leerCoordenada(texto, eje) {
    if (texto === null || texto === undefined || String(texto).trim() === '') {
      return { error: 'vacía' };
    }
    if (typeof texto === 'number') return validarRango(texto, eje);

    let t = String(texto).trim().toUpperCase()
      .replace(/[−–—]/g, '-')
      .replace(/[º˚]/g, '°')
      .replace(/[’′´`]/g, "'")
      .replace(/[”″]|''/g, '"');

    // Hemisferio por letra: N S E W O (oeste en español)
    let signo = 1;
    const letra = t.match(/[NSEWO]/);
    if (letra) {
      if (eje === 'lat' && !/[NS]/.test(letra[0])) return { error: 'la latitud debe usar N o S' };
      if (eje === 'lon' && !/[EWO]/.test(letra[0])) return { error: 'la longitud debe usar E, O o W' };
      if (letra[0] === 'S' || letra[0] === 'W' || letra[0] === 'O') signo = -1;
      t = t.replace(/[NSEWO]/g, ' ');
    }
    if (/^\s*-/.test(t)) {
      signo = -1;
      t = t.replace('-', ' ');
    }

    // Separar en partes numéricas (grados, minutos, segundos)
    const partes = t.split(/[°'"\s:]+/).filter(Boolean);
    if (partes.length === 0 || partes.length > 3) return { error: 'formato no reconocido' };
    const nums = partes.map(aNumero);
    if (nums.some(isNaN)) return { error: 'formato no reconocido' };

    const [g, m = 0, s = 0] = nums;
    if (partes.length >= 2 && (m < 0 || m >= 60)) return { error: 'minutos fuera de rango (0–59)' };
    if (partes.length === 3 && (s < 0 || s >= 60)) return { error: 'segundos fuera de rango (0–59)' };
    if (partes.length >= 2 && !Number.isInteger(g)) return { error: 'los grados deben ser enteros si hay minutos' };

    let valor = Math.abs(g) + m / 60 + s / 3600;
    if (g < 0) signo = -1;
    return validarRango(signo * valor, eje);
  }

  function validarRango(valor, eje) {
    const max = eje === 'lat' ? 90 : 180;
    if (!isFinite(valor) || Math.abs(valor) > max) {
      return { error: (eje === 'lat' ? 'latitud' : 'longitud') + ' fuera de rango (±' + max + ')' };
    }
    return { valor };
  }

  /** Lleva latitud/longitud de un datum a WGS84. */
  function geograficasAWGS84(lat, lon, datum) {
    if (!datum || datum === 'WGS84') return { lat, lon };
    const p = obtenerProj4()(defGeograficas(datum), WGS84, [lon, lat]);
    return { lat: p[1], lon: p[0] };
  }

  /** Convierte UTM (este, norte, zona, hemisferio, datum) a WGS84 geográficas. */
  function utmAWGS84(este, norte, zona, hemisferio, datum) {
    const e = aNumero(este), n = aNumero(norte);
    const z = parseInt(String(zona).replace(/[^\d]/g, ''), 10);
    // Hemisferio: "19S" o "19" se leen como sur (uso habitual en Chile); "19N" como norte.
    const h = hemisferio || (/N\s*$/i.test(String(zona)) ? 'N' : 'S');
    if (isNaN(e) || isNaN(n)) return { error: 'este o norte no numéricos' };
    if (!(z >= 1 && z <= 60)) return { error: 'zona UTM inválida' };
    if (e < 100000 || e > 900000) return { error: 'este fuera de rango (100 000–900 000 m)' };
    if (n < 0 || n > 10000000) return { error: 'norte fuera de rango (0–10 000 000 m)' };
    const p = obtenerProj4()(defUTM(z, h, datum || 'WGS84'), WGS84, [e, n]);
    return { lat: p[1], lon: p[0] };
  }

  /** Distancia en metros entre dos puntos (fórmula del semiverseno). */
  function distanciaMetros(lat1, lon1, lat2, lon2) {
    const R = 6371000, rad = Math.PI / 180;
    const dLat = (lat2 - lat1) * rad, dLon = (lon2 - lon1) * rad;
    const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLon / 2) ** 2;
    return 2 * R * Math.asin(Math.sqrt(a));
  }

  /** Texto en grados-minutos-segundos para mostrar. */
  function aGMS(valor, eje) {
    const hem = eje === 'lat' ? (valor < 0 ? 'S' : 'N') : (valor < 0 ? 'O' : 'E');
    let v = Math.abs(valor);
    let g = Math.floor(v);
    let m = Math.floor((v - g) * 60);
    let s = Math.round(((v - g) * 60 - m) * 60 * 10) / 10;
    if (s >= 60) { s = 0; m += 1; }
    if (m >= 60) { m = 0; g += 1; }
    return g + '°' + String(m).padStart(2, '0') + "'" + s.toFixed(1).padStart(4, '0') + '"' + hem;
  }

  /** Zona UTM sugerida para una longitud (Chile continental: 18–19; Juan Fernández: 17). */
  function zonaUTM(lon) {
    return Math.floor((lon + 180) / 6) + 1;
  }

  const api = { DATUMS, aNumero, leerCoordenada, geograficasAWGS84, utmAWGS84, distanciaMetros, aGMS, zonaUTM };
  raiz.DRCoord = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
