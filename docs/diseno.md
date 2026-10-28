# Diseño del tablero

## Idea central: teatro y taller

Una sola vista con dos zonas y una regla que se explica en una frase: **lo oscuro es el resultado (solo se lee) y lo blanco es el taller (ahí se escribe).**

| Zona | Referencia | Qué contiene |
|---|---|---|
| Barra superior | Apple (vidrio esmerilado) | Marca, acciones y un resumen fijo (anillo, %, estado) que aparece al bajar por la página |
| Riel "Tus indicadores" | Netflix (filas de títulos) | Una tarjeta por indicador con su pastel en miniatura, su suficiencia y su estado |
| Cartel | Netflix (título destacado) | Nombre del indicador, frase en palabras simples, Plan / proyectado / faltante |
| El pastel | Apple (anillos de actividad) | Firma visual del producto (ver abajo) |
| Lo que debes vigilar · Quién aporta más | — | Alertas ordenadas y ranking de iniciativas con "si falla" |
| Taller (pasos 1, 2 y 3) | Notion (bloques editables) | Formularios y tablas que se editan en el mismo lugar; el resultado cambia en vivo |

## El pastel

- Cada porción es un dolor, y su ángulo es su parte del pastel (atribución).
- Lo blanco de cada porción es lo que resuelven las iniciativas con la ejecución esperada. El **área** pintada es proporcional a lo resuelto, no el radio.
- La línea punteada marca lo que se resolvería si todo se cumple (Plan). La distancia entre lo blanco y la línea es lo que cuestan los retrasos.
- El anillo exterior es la suficiencia realista. Toma el color del semáforo y da una segunda vuelta si pasa de 100 %.
- Al pasar el mouse por un dolor se iluminan las iniciativas que lo atacan, y al revés.

## Colores (Claro)

| Token | Valor | Uso |
|---|---|---|
| `--rojo` | `#DA291C` | Acento sobre fondo claro: botones, foco, mapa de calor de la matriz |
| `--rojo-vivo` | `#EF3829` | Acento sobre fondo oscuro: marca, tarjeta elegida, estado "No alcanza" |
| `--noche` | `#141414` | Fondo del escenario |
| `--papel` / `--hoja` | `#F2F2F2` / `#FFFFFF` | Fondo y paneles del taller |
| `--ok-vivo` / `--alerta` | `#22C55E` / `#FAB219` | Estados "Alcanza" y "Cerca" |

Los estados siempre llevan **icono + texto**, nunca solo color.

## Tipografía

- **Archivo** condensada (títulos de indicadores y marca): el toque de "cartel" de Netflix.
- **Geist** (interfaz y números): limpia, al estilo Apple.
- **Geist Mono** (códigos como D1 o I3, unidades y etiquetas pequeñas).

Si no hay internet, se usan las fuentes del sistema y todo sigue funcionando.

## Arquitectura

| Archivo | Responsabilidad |
|---|---|
| `assets/js/engine.js` | Motor de cálculo puro. Replica la hoja CÁLCULOS del Excel; funciona en el navegador y en Node |
| `assets/js/ejemplos.js` | Los dos indicadores de ejemplo (iguales a los del Excel) |
| `assets/js/app.js` | Interfaz: pinta, edita, guarda en el navegador e importa Excel o JSON |
| `assets/css/app.css` | Estilos con tokens; diseño responsivo con *container queries* |
| `tests/paridad_excel.test.js` | Comprueba que el motor da lo mismo que el Excel (hasta 12 decimales) |
| `tests/navegador/*.js` | Pruebas de interacción e importación que corre `herramientas/probar_tablero.ps1` |
| `herramientas/*.ps1` | Generar y verificar la plantilla de Excel, y probar el tablero |

Reglas que conviene mantener:

- Toda la matemática va en `engine.js`. Si cambia una fórmula, debe cambiar también en el Excel (`herramientas/generar_plantilla.ps1`) y en la prueba de paridad.
- Los textos que escribe el usuario se insertan siempre con `textContent`, nunca como HTML.
- Los datos se guardan en el navegador (`localStorage`). Para compartirlos se usa **Guardar una copia** (.json) o el Excel.
