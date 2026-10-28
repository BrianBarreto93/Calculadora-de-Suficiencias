/*
 * Excel sin internet.
 * Escribe y lee archivos .xlsx en el navegador y en Node, sin librerías externas.
 * Un .xlsx es un ZIP con archivos XML: aquí se arma y se abre ese ZIP a mano.
 *  - Escribir: hojas con celdas, estilos, anchos, altos, celdas unidas, validaciones y enlaces.
 *  - Leer: el valor de cada celda (texto, número o sí/no) de cada hoja. No lee estilos.
 */
(function (raiz, fabrica) {
  if (typeof module === 'object' && module.exports) module.exports = fabrica();
  else raiz.SuficienciaExcel = fabrica();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var aBytes = new TextEncoder();
  var aTexto = new TextDecoder('utf-8');
  var XML_CAB = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  var NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var NS_PAQ = 'http://schemas.openxmlformats.org/package/2006/relationships';

  // ------------------------------------------------------------------ texto y referencias
  function esc(t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  // Caracteres de control que el XML no admite
  function limpiar(t) { return String(t).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, ''); }
  function desescapar(t) {
    return t.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|quot|apos|amp);/g, function (_, e) {
      if (e === 'lt') return '<';
      if (e === 'gt') return '>';
      if (e === 'quot') return '"';
      if (e === 'apos') return "'";
      if (e === 'amp') return '&';
      var cp = e.charAt(1) === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return cp > 0 && cp <= 0x10FFFF ? String.fromCodePoint(cp) : '';
    }).replace(/_x([0-9A-Fa-f]{4})_/g, function (_, hex) { return String.fromCharCode(parseInt(hex, 16)); });
  }
  function letra(col) {
    var t = '';
    for (var n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) t = String.fromCharCode(65 + (n - 1) % 26) + t;
    return t;
  }
  function ref(fila, col) { return letra(col) + (fila + 1); }
  function partirRef(r) {
    var m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(String(r));
    if (!m) return null;
    var col = 0;
    for (var i = 0; i < m[1].length; i++) col = col * 26 + m[1].charCodeAt(i) - 64;
    return { fila: Number(m[2]) - 1, col: col - 1 };
  }
  // Recorta en una palabra completa y termina en «…» (así no queda un paréntesis abierto ni media palabra)
  function recortar(texto, max) {
    if (texto.length <= max) return texto;
    var corte = texto.slice(0, max - 1);
    var espacio = corte.lastIndexOf(' ');
    if (espacio >= max / 2) corte = corte.slice(0, espacio);
    return corte.replace(/[\s(«<&,;:.\-–]+$/, '') + '…';
  }
  // Nombre de hoja válido para Excel: sin []:*?/\, sin apóstrofo al inicio o al final, máximo 31 caracteres y sin repetir.
  function nombreHoja(nombre, usados) {
    var base = String(nombre || '').replace(/[[\]:*?/\\\u0000-\u001F]/g, ' ').replace(/\s+/g, ' ').trim().replace(/^'+|'+$/g, '').trim() || 'Hoja';
    var n = 1;
    var final;
    do {
      var sufijo = n === 1 ? '' : ' (' + n + ')';
      final = recortar(base, 31 - sufijo.length) + sufijo;
      n += 1;
    } while (usados && usados.some(function (u) { return u.toLowerCase() === final.toLowerCase(); }));
    if (usados) usados.push(final);
    return final;
  }
  // Fechas de Excel: número de días desde 1899-12-30
  function serieDeFecha(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
    return m ? Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / 86400000 + 25569 : null;
  }
  function fechaDeSerie(n) {
    if (typeof n !== 'number' || !isFinite(n) || n < 1) return '';
    var d = new Date((Math.floor(n) - 25569) * 86400000);
    var dos = function (x) { return (x < 10 ? '0' : '') + x; };
    return d.getUTCFullYear() + '-' + dos(d.getUTCMonth() + 1) + '-' + dos(d.getUTCDate());
  }

  // ------------------------------------------------------------------ ZIP
  var TABLA_CRC = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = TABLA_CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function transformar(bytes, flujo) {
    return new Response(new Blob([bytes]).stream().pipeThrough(flujo)).arrayBuffer().then(function (b) { return new Uint8Array(b); });
  }
  // Comprime con el navegador si puede; si no, el archivo se guarda sin comprimir (igual de válido).
  function comprimir(bytes) {
    try {
      if (typeof CompressionStream !== 'function') return Promise.resolve(null);
      return transformar(bytes, new CompressionStream('deflate-raw')).catch(function () { return null; });
    } catch (e) { return Promise.resolve(null); }
  }
  function descomprimir(bytes) {
    try {
      if (typeof DecompressionStream !== 'function') return Promise.reject(new Error('sin-descompresor'));
      return transformar(bytes, new DecompressionStream('deflate-raw'));
    } catch (e) { return Promise.reject(new Error('sin-descompresor')); }
  }
  function unirBytes(partes) {
    var total = 0;
    partes.forEach(function (p) { total += p.length; });
    var salida = new Uint8Array(total);
    var pos = 0;
    partes.forEach(function (p) { salida.set(p, pos); pos += p.length; });
    return salida;
  }
  function fechaDos(d) {
    return { hora: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1), dia: ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate() };
  }

  function zip(archivos) {
    return Promise.all(archivos.map(function (a) {
      return comprimir(a.datos).then(function (c) {
        var usar = !!c && c.length < a.datos.length;
        return { nombre: aBytes.encode(a.nombre), datos: usar ? c : a.datos, metodo: usar ? 8 : 0, crc: crc32(a.datos), tam: a.datos.length };
      });
    })).then(function (lista) {
      var partes = [];
      var central = [];
      var pos = 0;
      var f = fechaDos(new Date());
      lista.forEach(function (a) {
        var loc = new DataView(new ArrayBuffer(30));
        loc.setUint32(0, 0x04034b50, true);
        loc.setUint16(4, 20, true);
        loc.setUint16(6, 0x0800, true);
        loc.setUint16(8, a.metodo, true);
        loc.setUint16(10, f.hora, true);
        loc.setUint16(12, f.dia, true);
        loc.setUint32(14, a.crc, true);
        loc.setUint32(18, a.datos.length, true);
        loc.setUint32(22, a.tam, true);
        loc.setUint16(26, a.nombre.length, true);
        partes.push(new Uint8Array(loc.buffer), a.nombre, a.datos);
        var cen = new DataView(new ArrayBuffer(46));
        cen.setUint32(0, 0x02014b50, true);
        cen.setUint16(4, 20, true);
        cen.setUint16(6, 20, true);
        cen.setUint16(8, 0x0800, true);
        cen.setUint16(10, a.metodo, true);
        cen.setUint16(12, f.hora, true);
        cen.setUint16(14, f.dia, true);
        cen.setUint32(16, a.crc, true);
        cen.setUint32(20, a.datos.length, true);
        cen.setUint32(24, a.tam, true);
        cen.setUint16(28, a.nombre.length, true);
        cen.setUint32(42, pos, true);
        central.push(new Uint8Array(cen.buffer), a.nombre);
        pos += 30 + a.nombre.length + a.datos.length;
      });
      var tamCentral = 0;
      central.forEach(function (p) { tamCentral += p.length; });
      var fin = new DataView(new ArrayBuffer(22));
      fin.setUint32(0, 0x06054b50, true);
      fin.setUint16(8, lista.length, true);
      fin.setUint16(10, lista.length, true);
      fin.setUint32(12, tamCentral, true);
      fin.setUint32(16, pos, true);
      return unirBytes(partes.concat(central, [new Uint8Array(fin.buffer)]));
    });
  }

  // Devuelve { 'ruta/archivo.xml': Uint8Array } solo con los XML (lo único que se necesita leer).
  function unzip(bytes) {
    var dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    var fin = -1;
    for (var i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { fin = i; break; }
    }
    if (fin < 0) return Promise.reject(new Error('no-es-zip'));
    var total = dv.getUint16(fin + 10, true);
    var p = dv.getUint32(fin + 16, true);
    var entradas = [];
    for (var k = 0; k < total; k++) {
      if (p + 46 > bytes.length || dv.getUint32(p, true) !== 0x02014b50) return Promise.reject(new Error('zip-danado'));
      var metodo = dv.getUint16(p + 10, true);
      var comp = dv.getUint32(p + 20, true);
      var largo = dv.getUint16(p + 28, true);
      var extra = dv.getUint16(p + 30, true);
      var comentario = dv.getUint16(p + 32, true);
      var local = dv.getUint32(p + 42, true);
      var nombre = aTexto.decode(bytes.subarray(p + 46, p + 46 + largo));
      if (/\.(xml|rels)$/i.test(nombre) && local + 30 <= bytes.length) {
        var inicio = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
        entradas.push({ nombre: nombre, metodo: metodo, datos: bytes.subarray(inicio, inicio + comp) });
      }
      p += 46 + largo + extra + comentario;
    }
    var salida = {};
    return Promise.all(entradas.map(function (e) {
      var leer = e.metodo === 0 ? Promise.resolve(e.datos) : (e.metodo === 8 ? descomprimir(e.datos) : Promise.reject(new Error('zip-metodo')));
      return leer.then(function (d) { salida[e.nombre] = d; });
    })).then(function () { return salida; });
  }

  // ------------------------------------------------------------------ escribir: hoja
  function Hoja(nombre) {
    this.nombre = nombre;
    this.filas = [];          // filas[f][c] = { v: valor, e: estilo }
    this.altos = {};
    this.anchos = [];
    this.uniones = [];
    this.validaciones = [];
    this.enlaces = [];
    this.congelar = null;     // { filas, columnas }
    this.color = null;        // color de la pestaña, 'RRGGBB'
    this.horizontal = false;  // impresión horizontal
  }
  Hoja.prototype.poner = function (fila, col, v, e) {
    var f = this.filas[fila] || (this.filas[fila] = []);
    f[col] = { v: v, e: e || null };
    return this;
  };
  // Da estilo a un rango sin cambiar los valores que ya tenga
  Hoja.prototype.pintar = function (fila, col, fila2, col2, e) {
    for (var f = fila; f <= fila2; f++) {
      for (var c = col; c <= col2; c++) {
        var x = (this.filas[f] || [])[c];
        if (x) x.e = e; else this.poner(f, c, null, e);
      }
    }
    return this;
  };
  Hoja.prototype.alto = function (fila, alto) { this.altos[fila] = alto; return this; };
  Hoja.prototype.unir = function (fila, col, fila2, col2) { this.uniones.push(ref(fila, col) + ':' + ref(fila2, col2)); return this; };
  // regla = { tipo: 'lista', valores: [...] } o { tipo: 'decimal' | 'entero', min, max, titulo, mensaje }
  Hoja.prototype.validar = function (fila, col, fila2, col2, regla) {
    var copia = {};
    Object.keys(regla).forEach(function (k) { copia[k] = regla[k]; });
    copia.rango = ref(fila, col) + (fila2 !== fila || col2 !== col ? ':' + ref(fila2, col2) : '');
    this.validaciones.push(copia);
    return this;
  };
  Hoja.prototype.enlazar = function (fila, col, hoja, texto) { this.enlaces.push({ ref: ref(fila, col), hoja: hoja, texto: texto }); return this; };

  // ------------------------------------------------------------------ escribir: estilos
  // e = { negrita, cursiva, tam, color, fondo, borde, bordeAbajo, formato, h, v, ajustar, sangria }
  var FORMATOS_FIJOS = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, 'fecha': 14 };
  function Estilos() {
    this.fuentes = ['<font><sz val="10"/><color rgb="FF262626"/><name val="Arial"/><family val="2"/></font>'];
    this.rellenos = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.bordes = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.formatos = [];
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    this.indice = {};
  }
  function indiceDe(lista, x) {
    var i = lista.indexOf(x);
    if (i < 0) { lista.push(x); i = lista.length - 1; }
    return i;
  }
  function lado(nombre, color, grosor) {
    return color ? '<' + nombre + ' style="' + grosor + '"><color rgb="FF' + color + '"/></' + nombre + '>' : '<' + nombre + '/>';
  }
  Estilos.prototype.id = function (e) {
    if (!e) return 0;
    var clave = JSON.stringify(e);
    if (this.indice[clave] !== undefined) return this.indice[clave];
    var fuente = indiceDe(this.fuentes, '<font>' + (e.negrita ? '<b/>' : '') + (e.cursiva ? '<i/>' : '') + (e.subrayado ? '<u/>' : '') +
      '<sz val="' + (e.tam || 10) + '"/><color rgb="FF' + (e.color || '262626') + '"/><name val="Arial"/><family val="2"/></font>');
    var relleno = e.fondo ? indiceDe(this.rellenos, '<fill><patternFill patternType="solid"><fgColor rgb="FF' + e.fondo + '"/><bgColor indexed="64"/></patternFill></fill>') : 0;
    var borde = 0;
    if (e.borde || e.bordeAbajo) {
      borde = indiceDe(this.bordes, '<border>' + lado('left', e.borde, 'thin') + lado('right', e.borde, 'thin') + lado('top', e.borde, 'thin') +
        lado('bottom', e.bordeAbajo || e.borde, e.bordeAbajo ? 'medium' : 'thin') + '<diagonal/></border>');
    }
    var formato = 0;
    if (e.formato) {
      if (FORMATOS_FIJOS[e.formato] !== undefined) formato = FORMATOS_FIJOS[e.formato];
      else {
        var f = this.formatos.filter(function (x) { return x.codigo === e.formato; })[0];
        if (!f) { f = { id: 164 + this.formatos.length, codigo: e.formato }; this.formatos.push(f); }
        formato = f.id;
      }
    }
    var alineacion = '<alignment' + (e.h ? ' horizontal="' + e.h + '"' : '') + ' vertical="' + (e.v || 'center') + '"' +
      (e.ajustar ? ' wrapText="1"' : '') + (e.sangria ? ' indent="' + e.sangria + '"' : '') + '/>';
    this.xfs.push('<xf numFmtId="' + formato + '" fontId="' + fuente + '" fillId="' + relleno + '" borderId="' + borde + '" xfId="0"' +
      (formato ? ' applyNumberFormat="1"' : '') + (fuente ? ' applyFont="1"' : '') + (relleno ? ' applyFill="1"' : '') +
      (borde ? ' applyBorder="1"' : '') + ' applyAlignment="1">' + alineacion + '</xf>');
    this.indice[clave] = this.xfs.length - 1;
    return this.indice[clave];
  };
  Estilos.prototype.xml = function () {
    return XML_CAB + '<styleSheet xmlns="' + NS_MAIN + '">' +
      (this.formatos.length ? '<numFmts count="' + this.formatos.length + '">' + this.formatos.map(function (f) {
        return '<numFmt numFmtId="' + f.id + '" formatCode="' + esc(f.codigo) + '"/>';
      }).join('') + '</numFmts>' : '') +
      '<fonts count="' + this.fuentes.length + '">' + this.fuentes.join('') + '</fonts>' +
      '<fills count="' + this.rellenos.length + '">' + this.rellenos.join('') + '</fills>' +
      '<borders count="' + this.bordes.length + '">' + this.bordes.join('') + '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + this.xfs.length + '">' + this.xfs.join('') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>';
  };

  // ------------------------------------------------------------------ escribir: XML de cada parte
  function refHoja(nombre) { return "'" + String(nombre).replace(/'/g, "''") + "'"; }
  function xmlValidacion(v) {
    var comun = ' allowBlank="1" showErrorMessage="1"' + (v.titulo ? ' errorTitle="' + esc(v.titulo) + '"' : '') +
      (v.mensaje ? ' error="' + esc(v.mensaje) + '"' : '') + ' sqref="' + v.rango + '"';
    if (v.tipo === 'lista') return '<dataValidation type="list"' + comun + '><formula1>"' + esc(v.valores.join(',')) + '"</formula1></dataValidation>';
    var tipo = v.tipo === 'entero' ? 'whole' : 'decimal';
    var conMax = v.max !== undefined && v.max !== null;
    return '<dataValidation type="' + tipo + '"' + (conMax ? '' : ' operator="greaterThanOrEqual"') + comun + '><formula1>' + v.min + '</formula1>' +
      (conMax ? '<formula2>' + v.max + '</formula2>' : '') + '</dataValidation>';
  }
  function xmlCelda(c, fila, col, estilos) {
    var s = estilos.id(c.e);
    var atr = ' r="' + ref(fila, col) + '"' + (s ? ' s="' + s + '"' : '');
    var v = c.v;
    if (typeof v === 'number' && isFinite(v)) return '<c' + atr + '><v>' + v + '</v></c>';
    if (typeof v === 'boolean') return '<c' + atr + ' t="b"><v>' + (v ? 1 : 0) + '</v></c>';
    if (typeof v === 'string' && v !== '') return '<c' + atr + ' t="inlineStr"><is><t xml:space="preserve">' + esc(limpiar(v)) + '</t></is></c>';
    return s ? '<c' + atr + '/>' : '';
  }
  // El orden de los elementos de la hoja es el que exige Excel.
  function xmlHoja(h, estilos, activa) {
    var ultimaCol = 0;
    h.filas.forEach(function (f) { if (f) ultimaCol = Math.max(ultimaCol, f.length - 1); });
    var x = [XML_CAB, '<worksheet xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '">'];
    x.push('<sheetPr>' + (h.color ? '<tabColor rgb="FF' + h.color + '"/>' : '') + '<pageSetUpPr fitToPage="1"/></sheetPr>');
    x.push('<dimension ref="A1:' + ref(Math.max(0, h.filas.length - 1), ultimaCol) + '"/>');
    var vista = '<sheetView workbookViewId="0" showGridLines="0"' + (activa ? ' tabSelected="1"' : '') + '>';
    if (h.congelar && (h.congelar.filas || h.congelar.columnas)) {
      var fz = h.congelar.filas || 0;
      var cz = h.congelar.columnas || 0;
      var panel = fz && cz ? 'bottomRight' : (fz ? 'bottomLeft' : 'topRight');
      vista += '<pane' + (cz ? ' xSplit="' + cz + '"' : '') + (fz ? ' ySplit="' + fz + '"' : '') + ' topLeftCell="' + ref(fz, cz) +
        '" activePane="' + panel + '" state="frozen"/><selection pane="' + panel + '" activeCell="' + ref(fz, cz) + '" sqref="' + ref(fz, cz) + '"/>';
    }
    x.push('<sheetViews>' + vista + '</sheetView></sheetViews>');
    x.push('<sheetFormatPr defaultRowHeight="13.2"/>');
    var cols = h.anchos.map(function (w, i) { return w ? '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>' : ''; }).join('');
    if (cols) x.push('<cols>' + cols + '</cols>');
    x.push('<sheetData>');
    for (var fi = 0; fi < h.filas.length; fi++) {
      var alto = h.altos[fi];
      var celdas = '';
      (h.filas[fi] || []).forEach(function (c, ci) { if (c) celdas += xmlCelda(c, fi, ci, estilos); });
      if (celdas || alto) x.push('<row r="' + (fi + 1) + '"' + (alto ? ' ht="' + alto + '" customHeight="1"' : '') + '>' + celdas + '</row>');
    }
    x.push('</sheetData>');
    if (h.uniones.length) x.push('<mergeCells count="' + h.uniones.length + '">' + h.uniones.map(function (u) { return '<mergeCell ref="' + u + '"/>'; }).join('') + '</mergeCells>');
    if (h.validaciones.length) x.push('<dataValidations count="' + h.validaciones.length + '">' + h.validaciones.map(xmlValidacion).join('') + '</dataValidations>');
    if (h.enlaces.length) {
      x.push('<hyperlinks>' + h.enlaces.map(function (e) {
        return '<hyperlink ref="' + e.ref + '" location="' + esc(refHoja(e.hoja) + '!A1') + '" display="' + esc(e.texto || e.hoja) + '"/>';
      }).join('') + '</hyperlinks>');
    }
    x.push('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>');
    x.push('<pageSetup orientation="' + (h.horizontal ? 'landscape' : 'portrait') + '" fitToWidth="1" fitToHeight="0"/>');
    x.push('</worksheet>');
    return x.join('');
  }

  // libro = { hojas: [Hoja], activa: índice de la hoja que se abre primero, titulo }
  // Devuelve una promesa con los bytes del .xlsx.
  function crearXlsx(libro) {
    var estilos = new Estilos();
    var hojas = libro.hojas;
    var activa = libro.activa || 0;
    var xmlsHojas = hojas.map(function (h, i) { return xmlHoja(h, estilos, i === activa); });
    var ahora = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    var archivos = [
      ['[Content_Types].xml', XML_CAB + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        hojas.map(function (_, i) { return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'; }).join('') +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '</Types>'],
      ['_rels/.rels', XML_CAB + '<Relationships xmlns="' + NS_PAQ + '">' +
        '<Relationship Id="rId1" Type="' + NS_REL + '/officeDocument" Target="xl/workbook.xml"/>' +
        '<Relationship Id="rId2" Type="' + NS_PAQ + '/metadata/core-properties" Target="docProps/core.xml"/>' +
        '</Relationships>'],
      ['docProps/core.xml', XML_CAB + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
        ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        '<dc:title>' + esc(libro.titulo || 'Suficiencia') + '</dc:title><dc:creator>' + esc(libro.autor || 'Calculadora de suficiencia') + '</dc:creator>' +
        '<dcterms:created xsi:type="dcterms:W3CDTF">' + ahora + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + ahora + '</dcterms:modified>' +
        '</cp:coreProperties>'],
      ['xl/workbook.xml', XML_CAB + '<workbook xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '">' +
        '<bookViews><workbookView activeTab="' + activa + '"/></bookViews><sheets>' +
        hojas.map(function (h, i) { return '<sheet name="' + esc(h.nombre) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join('') +
        '</sheets></workbook>'],
      ['xl/_rels/workbook.xml.rels', XML_CAB + '<Relationships xmlns="' + NS_PAQ + '">' +
        hojas.map(function (_, i) { return '<Relationship Id="rId' + (i + 1) + '" Type="' + NS_REL + '/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>'; }).join('') +
        '<Relationship Id="rId' + (hojas.length + 1) + '" Type="' + NS_REL + '/styles" Target="styles.xml"/></Relationships>'],
      ['xl/styles.xml', estilos.xml()]
    ];
    xmlsHojas.forEach(function (x, i) { archivos.push(['xl/worksheets/sheet' + (i + 1) + '.xml', x]); });
    return zip(archivos.map(function (a) { return { nombre: a[0], datos: aBytes.encode(a[1]) }; }));
  }

  // ------------------------------------------------------------------ leer
  function atributos(t) {
    var a = {};
    var re = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    var m;
    while ((m = re.exec(t))) a[m[1]] = desescapar(m[2] !== undefined ? m[2] : m[3]);
    return a;
  }
  function cadaEtiqueta(xml, etiqueta, fn) {
    var re = new RegExp('<' + etiqueta + '\\b([^>]*)>', 'g');
    var m;
    while ((m = re.exec(xml))) fn(atributos(m[1]));
  }
  // Texto de un <si> o un <is>: une todos los <t>, sin la guía fonética (<rPh>).
  function textoRico(x) {
    x = x.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').replace(/<t\b[^>]*\/>/g, '');
    var salida = '';
    var re = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
    var m;
    while ((m = re.exec(x))) salida += m[1];
    return desescapar(salida);
  }
  function leerHoja(xml, compartidos) {
    var celdas = {};
    var re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    var m;
    while ((m = re.exec(xml))) {
      var a = atributos(m[1]);
      var dentro = m[2] || '';
      if (!a.r) continue;
      var tipo = a.t || 'n';
      var valor;
      if (tipo === 'inlineStr') {
        var is = /<is>([\s\S]*?)<\/is>/.exec(dentro);
        valor = is ? textoRico(is[1]) : '';
      } else {
        var v = /<v>([\s\S]*?)<\/v>/.exec(dentro);
        if (!v || tipo === 'e') continue;
        if (tipo === 's') valor = compartidos[Number(v[1])];
        else if (tipo === 'str') valor = desescapar(v[1]);
        else if (tipo === 'b') valor = v[1] === '1';
        else {
          valor = Number(v[1]);
          if (!isFinite(valor)) continue;
        }
      }
      if (valor === undefined || valor === '') continue;
      celdas[a.r] = { v: valor };
    }
    return celdas;
  }

  // Devuelve una promesa con { nombres: [hojas en orden], hojas: { nombre: { 'A1': { v } } } }
  function leerXlsx(entrada) {
    var bytes = entrada instanceof Uint8Array ? entrada : new Uint8Array(entrada);
    return unzip(bytes).then(function (archivos) {
      var texto = function (n) { return archivos[n] ? aTexto.decode(archivos[n]) : ''; };
      var libro = texto('xl/workbook.xml');
      if (!libro) throw new Error('no-es-xlsx');
      var rels = {};
      cadaEtiqueta(texto('xl/_rels/workbook.xml.rels'), 'Relationship', function (a) { rels[a.Id] = a.Target; });
      var compartidos = [];
      var reSi = /<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g;
      var ss = texto('xl/sharedStrings.xml');
      var m;
      while ((m = reSi.exec(ss))) compartidos.push(m[1] ? textoRico(m[1]) : '');
      var salida = { nombres: [], hojas: {} };
      cadaEtiqueta(libro, 'sheet', function (a) {
        var idRel = a['r:id'] || a[Object.keys(a).filter(function (k) { return /:id$/.test(k); })[0]];
        var destino = rels[idRel];
        if (!destino) return;
        var ruta = destino.charAt(0) === '/' ? destino.slice(1) : 'xl/' + destino.replace(/^\.\//, '');
        salida.nombres.push(a.name);
        salida.hojas[a.name] = leerHoja(texto(ruta), compartidos);
      });
      return salida;
    });
  }

  // Texto de cada XML del archivo (para pruebas y diagnóstico)
  function leerPartes(entrada) {
    return unzip(entrada instanceof Uint8Array ? entrada : new Uint8Array(entrada)).then(function (archivos) {
      var salida = {};
      Object.keys(archivos).forEach(function (n) { salida[n] = aTexto.decode(archivos[n]); });
      return salida;
    });
  }

  return {
    Hoja: Hoja,
    crearXlsx: crearXlsx,
    leerXlsx: leerXlsx,
    leerPartes: leerPartes,
    ref: ref,
    letra: letra,
    partirRef: partirRef,
    nombreHoja: nombreHoja,
    serieDeFecha: serieDeFecha,
    fechaDeSerie: fechaDeSerie,
    TIPO: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  };
});
