/*
 * Acceso al tablero: usuario y contraseña antes de ver la calculadora.
 * Se carga en el <head>, antes que todo, para tapar la página desde el primer pintado.
 * Es una puerta de cortesía en el navegador (la página es estática): no protege los archivos del servidor.
 * La clave no está escrita aquí: solo su huella SHA-256 de «usuario:contraseña».
 * La sesión dura lo que la pestaña (sessionStorage): al cerrar la pestaña o el navegador se vuelve a pedir.
 */
(function () {
  'use strict';
  var HUELLA = '98191e343ffd286b102bd8c60f8752f1fbd00a7a3fa1bf14983adbc912021ece';
  var CLAVE_SESION = 'suficiencia.acceso.v1';
  var raiz = document.documentElement;

  function dentro() {
    try { return window.sessionStorage.getItem(CLAVE_SESION) === HUELLA; } catch (e) { return false; }
  }
  if (dentro()) return;
  raiz.classList.add('acceso-bloqueado');

  // SHA-256 en JS puro (crypto.subtle no existe al abrir el archivo desde el disco en algunos navegadores)
  function sha256(texto) {
    var bytes = unescape(encodeURIComponent(texto)).split('').map(function (c) { return c.charCodeAt(0); });
    var K = [], H = [], primos = [], n = 2;
    while (primos.length < 64) {
      if (primos.every(function (p) { return n % p; })) primos.push(n);
      n += 1;
    }
    var frac = function (x) { return ((x - Math.floor(x)) * 4294967296) | 0; };
    primos.forEach(function (p, i) {
      if (i < 8) H[i] = frac(Math.pow(p, 1 / 2));
      K[i] = frac(Math.pow(p, 1 / 3));
    });
    var largo = bytes.length * 8;
    bytes.push(0x80);
    while (bytes.length % 64 !== 56) bytes.push(0);
    for (var s = 56; s >= 0; s -= 8) bytes.push(s >= 32 ? 0 : (largo >>> s) & 255);
    var rot = function (x, r) { return (x >>> r) | (x << (32 - r)); };
    for (var b = 0; b < bytes.length; b += 64) {
      var W = [];
      for (var t = 0; t < 64; t++) {
        if (t < 16) W[t] = (bytes[b + 4 * t] << 24) | (bytes[b + 4 * t + 1] << 16) | (bytes[b + 4 * t + 2] << 8) | bytes[b + 4 * t + 3];
        else {
          var s0 = rot(W[t - 15], 7) ^ rot(W[t - 15], 18) ^ (W[t - 15] >>> 3);
          var s1 = rot(W[t - 2], 17) ^ rot(W[t - 2], 19) ^ (W[t - 2] >>> 10);
          W[t] = (W[t - 16] + s0 + W[t - 7] + s1) | 0;
        }
      }
      var a = H.slice();
      for (t = 0; t < 64; t++) {
        var S1 = rot(a[4], 6) ^ rot(a[4], 11) ^ rot(a[4], 25);
        var ch = (a[4] & a[5]) ^ (~a[4] & a[6]);
        var t1 = (a[7] + S1 + ch + K[t] + W[t]) | 0;
        var S0 = rot(a[0], 2) ^ rot(a[0], 13) ^ rot(a[0], 22);
        var maj = (a[0] & a[1]) ^ (a[0] & a[2]) ^ (a[1] & a[2]);
        a = [(t1 + S0 + maj) | 0, a[0], a[1], a[2], (a[3] + t1) | 0, a[4], a[5], a[6]];
      }
      for (t = 0; t < 8; t++) H[t] = (H[t] + a[t]) | 0;
    }
    return H.map(function (x) { return ('00000000' + (x >>> 0).toString(16)).slice(-8); }).join('');
  }

  document.addEventListener('DOMContentLoaded', function () {
    var form = document.getElementById('acceso-form');
    var usuario = document.getElementById('acceso-usuario');
    var clave = document.getElementById('acceso-clave');
    var error = document.getElementById('acceso-error');
    if (!form) return;
    usuario.focus();
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var huella = sha256(usuario.value.trim().toLowerCase() + ':' + clave.value);
      if (huella !== HUELLA) {
        error.hidden = false;
        clave.value = '';
        clave.setAttribute('aria-invalid', 'true');
        clave.focus();
        return;
      }
      try { window.sessionStorage.setItem(CLAVE_SESION, HUELLA); } catch (err) { /* sin almacenamiento: se pide en cada recarga */ }
      raiz.classList.remove('acceso-bloqueado');
      form.reset();
      // Lo que se midió con la página tapada (flechas del riel, anchos) se vuelve a medir
      window.dispatchEvent(new Event('resize'));
    });
    clave.addEventListener('input', function () { error.hidden = true; clave.removeAttribute('aria-invalid'); });
  });
})();
