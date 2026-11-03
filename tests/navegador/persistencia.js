// Prueba de lo que se guarda en el navegador. La ejecuta herramientas/probar_tablero.ps1 dos veces con el mismo perfil:
//  1.ª vez: crea un indicador Transversal con la parte del problema de un segmento, lo guarda en Mis cálculos, lo sigue editando y elige otro indicador.
//  2.ª vez (como si se volviera a abrir el link): todo debe estar igual. Luego prueba el respaldo .json
//  (descargar, borrar y volver a cargar) y que la página siga funcionando si el navegador no deja guardar.
(function () {
  var errores = [];
  window.addEventListener('error', function (e) { errores.push(e.message + ' @' + e.lineno); });
  function $(s, r) { return (r || document).querySelector(s); }
  function $$(s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); }
  var out = [];
  function ok(c, m) { out.push((c ? 'OK    ' : 'FALLA ') + m); }
  function disparar(el, tipo) { el.dispatchEvent(new Event(tipo, { bubbles: true })); }
  // Como en el navegador: «input» con cada tecla y «change» al salir del campo
  function escribir(el, v) { el.focus(); el.value = v; disparar(el, 'input'); disparar(el, 'change'); el.blur(); }
  function titulo() { return $('#titulo-indicador').textContent; }
  function aviso() { return $('#aviso').textContent; }
  function guardado() { return JSON.parse(localStorage.getItem('suficiencia.espacio.v3') || 'null'); }
  function terminar() {
    out.push('errores JS: ' + JSON.stringify(errores));
    var pre = document.createElement('pre');
    pre.id = 'resultado-prueba';
    pre.textContent = out.join('\n');
    document.body.appendChild(pre);
  }
  // Corre los pasos uno tras otro. Un paso puede devolver una condición: se espera a que se cumpla (lecturas de archivos).
  function pasos(lista) {
    var k = 0;
    (function siguiente() {
      if (k >= lista.length) { terminar(); return; }
      var espera;
      try { espera = lista[k++](); } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); return; }
      var intentos = 0;
      (function revisar() {
        if (espera && !espera() && intentos++ < 100) { setTimeout(revisar, 50); return; }
        setTimeout(siguiente, 120);
      })();
    })();
  }
  function marcar(nombre) { $$('#paso-1 input[data-campo="segmento"]').filter(function (x) { return x.value === nombre; })[0].click(); }
  function peso(nombre) { return $$('#paso-1 input[data-campo="peso"]').filter(function (x) { return x.dataset.segmento === nombre; })[0]; }
  function tarjeta(nombre) { return $$('.tarjeta[data-id]').filter(function (t) { return $('.tarjeta-nombre', t).textContent === nombre; })[0]; }

  // Lo que la página intenta descargar se guarda aquí en vez de bajarse
  var descargas = [];
  var crearUrl = URL.createObjectURL;
  URL.createObjectURL = function (blob) { descargas.push({ blob: blob }); return crearUrl.call(URL, blob); };
  var clic = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function () {
    if (this.download) { descargas[descargas.length - 1].nombre = this.download; return; }
    return clic.call(this);
  };

  var fase = localStorage.getItem('prueba.persistencia');
  out.push('fase ' + (fase === '2' ? 2 : 1));
  if (fase !== '2') {
    setTimeout(function () {
      pasos([
        function () {
          ok(titulo() === 'Churn mensual pospago' && !localStorage.getItem('suficiencia.datos.v1'), 'primera vez: solo el ejemplo');
          $('.tarjeta--nueva').click();
          escribir($('#paso-1 input[data-campo="nombre"]'), 'Persistente');
          $('[data-accion="tipo-segmento"][data-valor="Transversal"]').click();
          marcar('Hogar');
          marcar('Pyme Móvil');
          escribir(peso('Hogar'), '70');
          $('[aria-controls="menu-calculos"]').click();
          $('[data-accion="guardar-calculo"]').click();
          ok($('#dlg-pregunta').open && !$('#preg-campo').hidden, 'Guardar pide un nombre');
          $('#preg-nombre').value = 'Guardado A';
          $$('#preg-botones button')[0].click();
        },
        function () {
          ok(/Guardaste «Guardado A»/.test(aviso()), 'guardado en Mis cálculos');
          // Sigue editando sin guardar y deja elegido el ejemplo
          escribir($('#paso-1 input[data-campo="meta"]'), '5');
          tarjeta('Churn mensual pospago').click();
          localStorage.setItem('prueba.persistencia', '2');
        },
        function () {
          var g = guardado();
          ok(g && g.version === 3 && g.calculos.length === 1 && g.actual.indicadores.length === 2, 'se guardó solo, a cada cambio');
        }
      ]);
    }, 300);
    return;
  }

  // Segunda vez: como si se volviera a abrir el link
  var respaldo = null;
  setTimeout(function () {
    pasos([
      function () {
        ok(/respaldo/.test(aviso()) && !!$('#aviso [data-accion="descargar-respaldo"]'), 'con cálculos guardados y sin respaldo, al abrir recuerda descargarlo → ' + aviso());
        ok(titulo() === 'Churn mensual pospago' && $$('.tarjeta[data-id]').length === 2, 'al volver se ve lo mismo: 2 indicadores y el elegido');
        tarjeta('Persistente').click();
        ok(titulo() === 'Persistente' && $('#paso-1 input[data-campo="meta"]').value === '5', 'el indicador conserva lo último que se escribió');
        ok($('[data-accion="tipo-segmento"][data-valor="Transversal"]').getAttribute('aria-pressed') === 'true' && $$('#paso-1 input[data-campo="segmento"]:checked').length === 2, 'conserva Transversal y sus 2 segmentos');
        ok(peso('Hogar').value === '70' && peso('Pyme Móvil').value === '', 'conserva el peso de Hogar (70 %) y Pyme Móvil sin peso');
        $('[aria-controls="menu-calculos"]').click();
        var menu = $('#menu-calculos');
        ok($$('.calc-item', menu).length === 1 && $('.calc-nombre', menu).value === 'Guardado A' && $('.calc-item.abierto', menu), 'Mis cálculos conserva «Guardado A» como el abierto');
        ok(/cambios sin guardar/.test(menu.textContent), 'y recuerda que hay cambios sin guardar');
        $('[data-accion="descargar-respaldo"]', menu).click();
      },
      function () {
        var d = descargas.filter(function (x) { return x.nombre && /\.json$/.test(x.nombre); })[0];
        ok(d && /^suficiencia-respaldo-\d{4}-\d{2}-\d{2}\.json$/.test(d.nombre), 'descargar respaldo → ' + (d && d.nombre));
        ok(!!localStorage.getItem('suficiencia.respaldo.v1'), 'recuerda cuándo se descargó (el aviso vuelve en 7 días)');
        d.blob.text().then(function (t) { respaldo = t; });
        return function () { return respaldo !== null; };
      },
      function () {
        ok(respaldo !== null, 'se pudo leer el respaldo descargado');
        var r = JSON.parse(respaldo);
        ok(r.version === 3 && r.tipo === 'respaldo' && r.calculos.length === 1 && r.actual.indicadores.length === 2 && r.activo === r.calculos[0].id, 'el respaldo trae todo: cálculos, tablero y cálculo activo');
        // Se borra todo: nuevo cálculo sin guardar y se elimina el guardado
        $('.barra [data-accion="nuevo-calculo"]').click();
        $$('#preg-botones button')[1].click();
      },
      function () {
        $('[aria-controls="menu-calculos"]').click();
        $('[data-accion="eliminar-calculo"]').click();
        $$('#preg-botones button')[0].click();
      },
      function () {
        ok($$('.tarjeta[data-id]').length === 1 && guardado().calculos.length === 0, 'todo borrado antes de cargar el respaldo → ' + $$('.tarjeta[data-id]').length + ' tarjetas, ' + guardado().calculos.length + ' cálculos, aviso: ' + aviso());
        var dt = new DataTransfer();
        dt.items.add(new File([respaldo], 'respaldo.json', { type: 'application/json' }));
        var inp = document.getElementById('archivo');
        inp.files = dt.files;
        disparar(inp, 'change');
        return function () { return /respaldo\.json/.test(aviso()); };
      },
      function () {
        var g = guardado();
        ok(/Cargaste el respaldo respaldo\.json: 1 cálculo guardado/.test(aviso()), 'aviso → ' + aviso().replace('Deshacer', '').trim());
        ok(g.calculos.length === 1 && g.calculos[0].nombre === 'Guardado A' && $$('.tarjeta[data-id]').length === 2 && titulo() === 'Persistente', 'cargar respaldo devuelve los cálculos y el tablero');
        var p = g.actual.indicadores.filter(function (x) { return x.nombre === 'Persistente'; })[0];
        ok(p && JSON.stringify(p.segmentos) === '[{"nombre":"Hogar","peso":0.7,"casos":null},{"nombre":"Pyme Móvil","peso":null,"casos":null}]', 'con sus segmentos y su parte del problema');
        // Un archivo que no es respaldo se rechaza con un mensaje simple
        var dt = new DataTransfer();
        dt.items.add(new File(['{"hola": 1}'], 'otro.json', { type: 'application/json' }));
        var inp = document.getElementById('archivo');
        inp.files = dt.files;
        disparar(inp, 'change');
        return function () { return /otro\.json/.test(aviso()); };
      },
      function () {
        ok(/No se pudo abrir otro\.json/.test(aviso()) && $$('.tarjeta[data-id]').length === 2, 'un .json cualquiera no borra nada → ' + aviso());
        // Si el navegador no deja guardar, se avisa en español y se sigue trabajando
        var setItem = Storage.prototype.setItem;
        Storage.prototype.setItem = function () { throw new Error('QuotaExceededError'); };
        escribir($('#paso-1 input[data-campo="nombre"]'), 'Sin almacenamiento');
        setTimeout(function () { Storage.prototype.setItem = setItem; }, 400);
      },
      function () {},
      function () {},
      function () {},
      function () {
        ok(/No se pudo guardar en este navegador/.test(aviso()) && titulo() === 'Sin almacenamiento', 'sin almacenamiento avisa y deja seguir → ' + aviso());
        localStorage.removeItem('prueba.persistencia');
      }
    ]);
  }, 300);
})();
