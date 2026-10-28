// Prueba de exportar e importar Excel. La ejecuta herramientas/probar_tablero.ps1 (no necesita internet).
// 1. Carga indicadores de prueba (incluido uno con 15 dolores, 15 iniciativas y nombres con símbolos).
// 2. Exporta con el diálogo y lee el archivo con el mismo lector del tablero.
// 3. Lo edita como lo haría alguien en Excel (inserta una fila, escribe 45 en vez de 45 %, una fecha como texto…).
// 4. Lo vuelve a importar y compara con el motor de cálculo.
(function () {
  var errores = [];
  window.addEventListener('error', function (e) { errores.push(e.message + ' @' + e.lineno); });
  window.addEventListener('unhandledrejection', function (e) { errores.push('promesa: ' + (e.reason && e.reason.message)); });
  function $(s, r) { return (r || document).querySelector(s); }
  function $$(s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); }
  var out = [];
  var base64 = '';
  function ok(c, m) { out.push((c ? 'OK    ' : 'FALLA ') + m); }
  function terminar() {
    out.push('errores JS: ' + JSON.stringify(errores));
    var pre = document.createElement('pre');
    pre.id = 'resultado-prueba';
    pre.textContent = out.join('\n');
    document.body.appendChild(pre);
    if (base64) {
      var b = document.createElement('pre');
      b.id = 'xlsx-base64';
      b.textContent = base64;
      document.body.appendChild(b);
    }
  }
  function esperar(cond, listo, intentos) {
    intentos = intentos || 0;
    if (cond() || intentos > 120) { listo(); return; }
    setTimeout(function () { esperar(cond, listo, intentos + 1); }, 100);
  }
  function aviso() { return document.getElementById('aviso').textContent; }
  function subir(archivo) {
    var dt = new DataTransfer();
    dt.items.add(archivo);
    var inp = document.getElementById('archivo');
    inp.files = dt.files;
    inp.dispatchEvent(new Event('change'));
  }
  function aBase64(bytes) {
    var t = '';
    for (var i = 0; i < bytes.length; i += 0x8000) t += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(t);
  }

  // Lo que la página intenta descargar se guarda aquí en vez de bajarse
  var descargas = [];
  var crearUrl = URL.createObjectURL;
  URL.createObjectURL = function (blob) { descargas.push({ blob: blob }); return crearUrl.call(URL, blob); };
  var clic = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function () {
    if (this.download) { descargas[descargas.length - 1].nombre = this.download; return; }
    return clic.call(this);
  };

  // Indicadores de prueba
  var ej = SuficienciaEjemplos();
  var grande = { id: 'grande', nombre: 'Disponibilidad «red» & <móvil> "5G"', sentido: 'sube', unidad: '%', ejecutado: 97.5, meta: 99.2, ideal: null, totalCasos: 15000, dolores: [], iniciativas: [] };
  for (var j = 0; j < 15; j++) grande.dolores.push({ id: 'g' + j, nombre: 'Dolor ' + (j + 1) + (j === 0 ? ' con un nombre bastante largo para ver cómo se ajusta en la celda' : ''), casos: 1000 + j * 100, atribucionManual: null, evidencia: j % 3 ? 'Fuente ' + j : '' });
  for (var i = 0; i < 15; i++) {
    var ef = {};
    ef['g' + i] = 0.5;
    ef['g' + ((i + 3) % 15)] = 0.2;
    grande.iniciativas.push({ id: 'n' + i, nombre: 'Iniciativa ' + (i + 1), efectividad: ef, ejecucion: i === 4 ? null : 0.5 + (i % 5) * 0.1, estado: ['En plan', 'En riesgo', 'Retrasada'][i % 3], responsable: 'Área ' + i, fecha: i % 2 ? '2027-0' + (1 + (i % 9)) + '-15' : '', evidencia: i % 4 ? 'Piloto ' + i : '' });
  }
  var vacio = Motor.nuevoIndicador();
  var datos = { app: 'Suficiencia 2027', version: 2, sel: 'ej-churn', params: Motor.PARAMS_DEFECTO, indicadores: [ej[0], ej[1], grande, vacio] };

  setTimeout(function () {
    try {
      subir(new File([JSON.stringify(datos)], 'prueba.json', { type: 'application/json' }));
    } catch (e) { out.push('EXCEPCION ' + e.message); terminar(); return; }
    esperar(function () { return aviso().indexOf('Abriste la copia') >= 0; }, exportar);
  }, 300);

  function exportar() {
    try {
      ok($$('.tarjeta[data-id]').length === 4, 'se cargaron 4 indicadores de prueba');
      $('[data-accion="exportar-excel"]').click();
      var marcas = $$('#exp-lista input[type="checkbox"]');
      ok($('#dlg-exportar').open && marcas.length === 4, 'el diálogo lista los 4 indicadores');
      ok(!marcas[0].checked && !marcas[1].checked && marcas[2].checked && marcas[3].checked, 'los ejemplos vienen sin marcar porque hay cálculos propios');
      $('[data-accion="exp-ninguno"]').click();
      ok($('#exp-ok').disabled, 'sin indicadores elegidos no se puede exportar');
      $('[data-accion="exp-todos"]').click();
      ok(!$('#exp-ok').disabled && /4 indicadores/.test($('#exp-ok').textContent), 'botón → ' + $('#exp-ok').textContent);
      $('#exp-ok').click();
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); return; }
    esperar(function () { return descargas.some(function (d) { return d.nombre && /\.xlsx$/.test(d.nombre); }); }, revisar);
  }

  var leido = null;
  function revisar() {
    var d = descargas.filter(function (x) { return x.nombre && /\.xlsx$/.test(x.nombre); })[0];
    if (!d) { ok(false, 'no se descargó ningún .xlsx'); terminar(); return; }
    ok(/^Suficiencia_CX_\d{4}-\d{2}-\d{2}\.xlsx$/.test(d.nombre), 'nombre del archivo → ' + d.nombre);
    ok(!$('#dlg-exportar').open, 'el diálogo se cierra al exportar');
    d.blob.arrayBuffer().then(function (buf) {
      var bytes = new Uint8Array(buf);
      base64 = aBase64(bytes);
      return SuficienciaExcel.leerXlsx(bytes);
    }).then(function (libro) {
      leido = libro;
      ok(libro.nombres.join(' | ') === 'Resumen | 01 Churn mensual pospago | 02 Instalaciones en fecha… | 03 Disponibilidad «red»… | 04 Indicador sin nombre | Ajustes | Glosario',
        'hojas (los nombres largos se recortan en una palabra) → ' + libro.nombres.join(' | '));
      var res = libro.hojas.Resumen;
      var filaChurn = Object.keys(res).filter(function (k) { return res[k].v === 'Churn mensual pospago  (ejemplo)'; })[0];
      ok(!!filaChurn, 'el Resumen nombra el ejemplo');
      var churn = libro.hojas['01 Churn mensual pospago'];
      ok(churn.A1.v === 'SUFICIENCIA 2027 · HOJA DE INDICADOR', 'cada hoja de indicador tiene su marca');
      var suf = Object.keys(churn).filter(function (k) { return churn[k].v === 'Suficiencia esperada'; })[0];
      var fila = suf && Number(suf.slice(1));
      ok(suf && Math.abs(churn['C' + fila].v - Motor.calcular(ej[0]).real.suficiencia) < 1e-12, 'suficiencia esperada del Churn = motor');
      var mueve = churn['C' + (fila + 1)];
      ok(churn['B' + (fila + 1)].v === 'Mueve el indicador' && Math.abs(mueve.v - Motor.calcular(ej[0]).real.pp) < 1e-12, 'puntos que mueve el indicador = motor (' + mueve.v.toFixed(4) + ')');
      var textos = JSON.stringify(libro.hojas);
      ok(textos.indexOf('% dolor') >= 0 && textos.indexOf('Impacto del dolor') >= 0 && !/[Pp]revalencia|[Pp]arte del pastel/.test(textos), 'el Excel usa «% dolor» e «Impacto del dolor»');
      ok(/Sin cobertura/.test(textos) && /Reemplazar/.test(textos) && /Cerca/.test(textos), 'trae semáforo, etiquetas y alertas en palabras');
      ok(!!libro.hojas.Glosario && JSON.stringify(libro.hojas.Glosario).indexOf('Mueve el indicador') >= 0, 'el glosario explica «Mueve el indicador»');

      // Primera vuelta: importar el archivo tal cual
      subir(new File([Uint8Array.from(atob(base64), function (c) { return c.charCodeAt(0); })], 'exportado.xlsx'));
      // La página guarda en el navegador un instante después de pintar
      esperar(function () { return aviso().indexOf('Importaste') >= 0; }, function () { setTimeout(vuelta1, 500); });
    }).catch(function (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); });
  }

  function guardados() { return JSON.parse(localStorage.getItem('suficiencia.datos.v1') || '{"indicadores":[]}').indicadores; }
  function comparar(a, b, que) {
    var ra = Motor.calcular(a);
    var rb = Motor.calcular(b);
    ok(Math.abs(ra.real.suficiencia - rb.real.suficiencia) < 1e-12 && Math.abs(ra.plan.suficiencia - rb.plan.suficiencia) < 1e-12,
      que + ': misma suficiencia (' + Math.round(ra.real.suficiencia * 100) + ' % / ' + Math.round(ra.plan.suficiencia * 100) + ' %)');
    var resumen = function (x) {
      return JSON.stringify([x.nombre, x.sentido, x.unidad, x.ejecutado, x.meta, x.ideal, x.totalCasos,
        x.dolores.map(function (d) { return [d.nombre, d.casos, d.atribucionManual, d.evidencia]; }),
        x.iniciativas.map(function (y) { return [y.nombre, y.ejecucion, y.estado, y.responsable, y.fecha, y.evidencia, x.dolores.map(function (d) { return y.efectividad[d.id] === undefined ? null : y.efectividad[d.id]; })]; })]);
    };
    ok(resumen(a) === resumen(b), que + ': mismos datos, dolores e iniciativas');
  }

  function vuelta1() {
    try {
      ok(/Importaste 4 actualizados desde exportado\.xlsx/.test(aviso()), 'aviso → ' + aviso().replace('Deshacer', '').trim());
      var inds = guardados();
      ok(inds.length === 4, 'siguen 4 indicadores (se actualizan por nombre)');
      [ej[0], ej[1], grande].forEach(function (orig) {
        var imp = inds.filter(function (x) { return x.nombre === orig.nombre; })[0];
        if (!imp) { ok(false, 'no volvió ' + orig.nombre); return; }
        comparar(imp, Motor.normalizarIndicador(orig), orig.nombre);
      });
      var g = inds.filter(function (x) { return x.nombre === grande.nombre; })[0];
      ok(g && g.dolores.length === 15 && g.iniciativas.length === 15, 'el indicador grande vuelve con 15 dolores y 15 iniciativas');
      ok(g && g.iniciativas[4].ejecucion === null, 'una ejecución vacía sigue vacía');
      ok(inds.every(function (x) { return !x.ejemplo; }), 'lo importado ya no se marca como ejemplo');
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); return; }
    editarEnExcel();
  }

  // Simula a alguien que editó el archivo en Excel y lo vuelve a importar
  function editarEnExcel() {
    try {
      var hojas = leido.nombres.map(function (nombre) {
        var celdas = leido.hojas[nombre];
        var h = new SuficienciaExcel.Hoja(nombre);
        var pos = {};
        Object.keys(celdas).forEach(function (k) { var p = SuficienciaExcel.partirRef(k); if (p) { pos[k] = p; h.poner(p.fila, p.col, celdas[k].v); } });
        return { h: h, celdas: celdas, pos: pos };
      });
      var ins = hojas.filter(function (x) { return x.h.nombre.indexOf('02 ') === 0; })[0];
      var c = ins.celdas;
      // Primera fila (después de «desde») cuya columna indicada tiene ese texto
      var filaDe = function (columna, texto, desde) {
        var filas = Object.keys(c).filter(function (r) { return ins.pos[r].col === columna && ins.pos[r].fila >= (desde || 0) && (texto instanceof RegExp ? texto.test(c[r].v) : c[r].v === texto); })
          .map(function (r) { return ins.pos[r].fila; });
        return filas.length ? Math.min.apply(null, filas) : -1;
      };
      var fHoy = filaDe(1, 'Hoy');
      ins.h.poner(fHoy, 2, 85);
      var fCab = filaDe(1, 'Iniciativa', filaDe(0, /^PASO 3/));
      // Primera iniciativa: 45 en vez de 45 % en el dolor 2 y la fecha escrita como texto
      ins.h.poner(fCab + 1, 3, 45);
      var colFecha = Object.keys(c).filter(function (r) { return ins.pos[r].fila === fCab && c[r].v === 'Fecha comprometida'; })[0];
      ins.h.poner(fCab + 1, ins.pos[colFecha].col, '15/02/2027');
      // Insertar una fila nueva después de la primera iniciativa (todo lo de abajo baja una fila)
      var nueva = new SuficienciaExcel.Hoja(ins.h.nombre);
      Object.keys(c).forEach(function (r) {
        var p = ins.pos[r];
        var v = (ins.h.filas[p.fila] || [])[p.col];
        nueva.poner(p.fila > fCab + 1 ? p.fila + 1 : p.fila, p.col, v ? v.v : c[r].v);
      });
      nueva.poner(fCab + 2, 1, 'Nueva desde Excel');
      nueva.poner(fCab + 2, 2, '20 %');
      var colEjec = Object.keys(c).filter(function (r) { return ins.pos[r].fila === fCab && c[r].v === 'Ejecución a 2027'; })[0];
      nueva.poner(fCab + 2, ins.pos[colEjec].col, 1);
      ins.h = nueva;
      var libro = { hojas: hojas.map(function (x) { return x.h; }) };
      SuficienciaExcel.crearXlsx(libro).then(function (bytes) {
        subir(new File([bytes], 'editado.xlsx'));
        esperar(function () { return /^(Importaste|No )/.test(aviso()) && /editado\.xlsx/.test(aviso()); }, function () { setTimeout(vuelta2, 500); });
      }).catch(function (e) { out.push('EXCEPCION ' + e.message); terminar(); });
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); }
  }

  function vuelta2() {
    try {
      ok(/Importaste 4 actualizados desde editado\.xlsx/.test(aviso()), 'aviso del archivo editado → ' + aviso().replace('Deshacer', '').trim());
      var imp = guardados().filter(function (x) { return x.nombre === ej[1].nombre; })[0];
      ok(imp && imp.ejecutado === 85, 'el valor de hoy editado en Excel llega → ' + (imp && imp.ejecutado));
      ok(imp && imp.iniciativas.length === 5 && imp.iniciativas[1].nombre === 'Nueva desde Excel', 'la fila insertada en Excel es una iniciativa nueva → ' + (imp && imp.iniciativas.map(function (x) { return x.nombre; }).join(' | ')));
      ok(imp && imp.iniciativas[1].efectividad[imp.dolores[0].id] === 0.2 && imp.iniciativas[1].ejecucion === 1, '«20 %» escrito como texto se lee como 20 %');
      ok(imp && imp.iniciativas[0].efectividad[imp.dolores[1].id] === 0.45, '45 escrito sin % se lee como 45 %');
      ok(imp && imp.iniciativas[0].fecha === '2027-02-15', 'la fecha escrita como texto 15/02/2027 → ' + (imp && imp.iniciativas[0].fecha));
      var esperado = JSON.parse(JSON.stringify(ej[1]));
      esperado.ejecutado = 85;
      esperado.iniciativas[0].efectividad.d2 = 0.45;
      esperado.iniciativas.splice(1, 0, { id: 'x', nombre: 'Nueva desde Excel', efectividad: { d1: 0.2 }, ejecucion: 1, estado: 'En plan' });
      var a = Motor.calcular(imp).real.suficiencia;
      var b = Motor.calcular(Motor.normalizarIndicador(esperado)).real.suficiencia;
      ok(Math.abs(a - b) < 1e-12, 'el resultado recalculado coincide con el motor (' + Math.round(a * 100) + ' %)');
      ok($('#titulo-indicador').textContent === ej[1].nombre || guardados().length === 4, 'la importación no duplica indicadores');
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); }
    terminar();
  }
})();
