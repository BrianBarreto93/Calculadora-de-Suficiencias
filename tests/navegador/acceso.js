// Prueba del acceso con usuario y contraseña. La ejecuta herramientas/probar_tablero.ps1 SIN la sesión abierta.
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
  setTimeout(function () {
    try {
      ok(tapado() && getComputedStyle($('#acceso')).display !== 'none', 'sin sesión se ve el acceso');
      ok(!visible($('#escenario')), 'el tablero queda oculto detrás');
      ok(document.activeElement === $('#acceso-usuario'), 'el cursor queda en Usuario');
      entrar('analyst', 'otra');
      ok(tapado() && !$('#acceso-error').hidden, 'contraseña mala: sigue tapado y avisa');
      ok($('#acceso-clave').value === '', 'y borra la contraseña');
      entrar('otro', 'secret123');
      ok(tapado(), 'usuario malo: sigue tapado');
      entrar('analyst', 'Secret123');
      ok(tapado(), 'la contraseña distingue mayúsculas');
      entrar('  Analyst ', 'secret123');
      ok(!tapado() && getComputedStyle($('#acceso')).display === 'none', 'usuario y contraseña buenos: se quita el acceso (usuario sin espacios ni mayúsculas)');
      ok(sessionStorage.getItem('suficiencia.acceso.v1') !== null, 'la sesión queda abierta en la pestaña');
      ok(document.documentElement.outerHTML.indexOf('secret123') < 0, 'la contraseña no está escrita en la página');
    } catch (e) { ok(false, 'error: ' + e.message); }
    // El navegador termina de volver a pintar el tablero en el cuadro siguiente
    setTimeout(function () {
      ok(visible($('#escenario')), 'y se ve el tablero');
      out.push('errores JS: ' + JSON.stringify(errores));
      var pre = document.createElement('pre');
      pre.id = 'resultado-prueba';
      pre.textContent = out.join('\n');
      document.body.appendChild(pre);
    }, 200);
  }, 800);
})();
