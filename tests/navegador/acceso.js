// Prueba del acceso con usuario y contraseña y de «Cerrar sesión».
// La ejecuta herramientas/probar_tablero.ps1 SIN la sesión abierta.
(function () {
  var out = [];
  function ok(c, m) { out.push((c ? 'OK    ' : 'FALLA ') + m); }
  function $(s) { return document.querySelector(s); }
  function tapado() { return document.documentElement.classList.contains('acceso-bloqueado'); }
  function visible(el) { return getComputedStyle(el).visibility === 'visible'; }
  var errores = [];
  window.addEventListener('error', function (e) { errores.push(e.message + ' @' + e.lineno); });
  function entrar(u, c) {
    $('#acceso-usuario').value = u;
    $('#acceso-clave').value = c;
    $('#acceso-form').dispatchEvent(new Event('submit', { cancelable: true }));
  }
  function terminar() {
    out.push('errores JS: ' + JSON.stringify(errores));
    var pre = document.createElement('pre');
    pre.id = 'resultado-prueba';
    pre.textContent = out.join('\n');
    document.body.appendChild(pre);
  }
  setTimeout(function () {
    try {
      ok(tapado() && getComputedStyle($('#acceso')).display !== 'none', 'sin sesión se ve el acceso');
      ok(!visible($('#escenario')), 'el tablero queda oculto detrás');
      ok(document.activeElement === $('#acceso-usuario'), 'el cursor queda en Usuario');
      entrar('Analista1', 'otra');
      ok(tapado() && !$('#acceso-error').hidden, 'contraseña mala: sigue tapado y avisa');
      ok($('#acceso-clave').value === '', 'y borra la contraseña');
      entrar('analyst', 'secret123');
      ok(tapado(), 'el usuario y la contraseña anteriores ya no sirven');
      entrar('Analista1', 'secret123');
      ok(tapado(), 'la contraseña distingue mayúsculas');
      entrar('  analista1 ', 'Secret123');
      ok(!tapado() && getComputedStyle($('#acceso')).display === 'none', 'usuario y contraseña buenos: se quita el acceso (usuario sin espacios ni mayúsculas)');
      ok(sessionStorage.getItem('suficiencia.acceso.v1') !== null, 'la sesión queda abierta en la pestaña');
      var html = document.documentElement.outerHTML;
      ok(html.indexOf('Secret123') < 0 && html.indexOf('secret123') < 0, 'la contraseña no está escrita en la página');
    } catch (e) { ok(false, 'error: ' + e.message); }
    // El navegador termina de volver a pintar el tablero en el cuadro siguiente
    setTimeout(function () {
      try {
        ok(visible($('#escenario')), 'y se ve el tablero');
        var salir = $('#cerrar-sesion');
        ok(!!salir && visible(salir), 'el botón «Cerrar sesión» está en la barra');
        var antes = localStorage.getItem('suficiencia.espacio.v3');
        salir.click();
        ok(tapado() && sessionStorage.getItem('suficiencia.acceso.v1') === null, 'Cerrar sesión: vuelve a tapar y borra la sesión');
        ok($('#acceso-usuario').value === '' && $('#acceso-clave').value === '' && $('#acceso-error').hidden, 'el formulario queda limpio');
        ok(document.activeElement === $('#acceso-usuario'), 'y el cursor en Usuario');
        ok(localStorage.getItem('suficiencia.espacio.v3') === antes, 'los cálculos guardados no se tocan');
        entrar('Analista1', 'Secret123');
        ok(!tapado(), 'se puede volver a entrar');
      } catch (e) { ok(false, 'error: ' + e.message); }
      setTimeout(function () { ok(visible($('#escenario')), 'y se ve el tablero otra vez'); terminar(); }, 200);
    }, 200);
  }, 800);
})();
