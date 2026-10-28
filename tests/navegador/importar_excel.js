// Prueba de importación: carga Plantilla_Suficiencia.xlsx en el tablero y la compara con los ejemplos.
// La ejecuta herramientas/probar_tablero.ps1 (necesita internet para descargar el lector de Excel).
(function () {
  var errores = [];
  window.addEventListener('error', function (e) { errores.push(e.message + ' @' + e.lineno); });
  var out = [];
  function ok(c, m) { out.push((c ? 'OK    ' : 'FALLA ') + m); }
  function terminar() {
    out.push('errores JS: ' + JSON.stringify(errores));
    var pre = document.createElement('pre');
    pre.id = 'resultado-prueba';
    pre.textContent = out.join('\n');
    document.body.appendChild(pre);
  }
  setTimeout(function () {
    // Borrar un ejemplo para comprobar que la importación lo trae de vuelta como nuevo
    document.querySelector('[data-accion="eliminar"]').click();
    var xhr = new XMLHttpRequest();
    xhr.open('GET', new URL('Plantilla_Suficiencia.xlsx', document.baseURI).href);
    xhr.responseType = 'blob';
    xhr.onload = function () {
      var dt = new DataTransfer();
      dt.items.add(new File([xhr.response], 'Plantilla_Suficiencia.xlsx'));
      var inp = document.getElementById('archivo');
      inp.files = dt.files;
      inp.dispatchEvent(new Event('change'));
      var intentos = 0;
      (function esperar() {
        intentos++;
        var aviso = document.getElementById('aviso').textContent;
        if (aviso.indexOf('Importaste') < 0 && intentos < 80) { setTimeout(esperar, 250); return; }
        try {
          ok(aviso.indexOf('Importaste') >= 0, 'aviso → ' + aviso.replace('Deshacer', '').trim());
          var datos = JSON.parse(localStorage.getItem('suficiencia.datos.v1') || '{"indicadores":[]}');
          var ej = SuficienciaEjemplos();
          ok(datos.indicadores.length === 2, 'quedan 2 indicadores (1 actualizado + 1 nuevo) → ' + datos.indicadores.length);
          ej.forEach(function (e) {
            var imp = datos.indicadores.filter(function (x) { return x.nombre === e.nombre; })[0];
            ok(!!imp, 'importado: ' + e.nombre);
            if (!imp) return;
            var rI = Motor.calcular(imp, datos.params);
            var rE = Motor.calcular(e);
            ok(Math.abs(rI.real.suficiencia - rE.real.suficiencia) < 1e-12 && Math.abs(rI.plan.suficiencia - rE.plan.suficiencia) < 1e-12,
              e.nombre + ': misma suficiencia que el ejemplo (' + Math.round(rI.real.suficiencia * 100) + ' % / ' + Math.round(rI.plan.suficiencia * 100) + ' %)');
            ok(imp.iniciativas.map(function (x) { return x.fecha; }).join() === e.iniciativas.map(function (x) { return x.fecha; }).join(), e.nombre + ': fechas → ' + imp.iniciativas.map(function (x) { return x.fecha; }).join(' '));
            ok(imp.iniciativas.map(function (x) { return x.estado + '|' + x.responsable + '|' + x.evidencia; }).join() === e.iniciativas.map(function (x) { return x.estado + '|' + x.responsable + '|' + x.evidencia; }).join(), e.nombre + ': estado, responsable y evidencia');
            ok(imp.dolores.map(function (d) { return d.nombre + '|' + d.casos + '|' + d.evidencia; }).join() === e.dolores.map(function (d) { return d.nombre + '|' + d.casos + '|' + d.evidencia; }).join(), e.nombre + ': dolores');
          });
          ok(JSON.stringify(datos.params) === JSON.stringify(Motor.PARAMS_DEFECTO), 'parámetros leídos de la hoja Parametros');
        } catch (err) { out.push('EXCEPCION ' + err.message); }
        terminar();
      })();
    };
    xhr.onerror = function () { out.push('No se pudo leer el xlsx local'); terminar(); };
    xhr.send();
  }, 300);
})();
