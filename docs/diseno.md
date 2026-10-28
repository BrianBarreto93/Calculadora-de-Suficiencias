# Diseño del tablero

## Idea central: teatro y taller

Una sola vista con dos zonas y una regla que se explica en una frase: **lo oscuro es el resultado (solo se lee) y lo blanco es el taller (ahí se escribe).**

| Zona | Referencia | Qué contiene |
|---|---|---|
| Barra superior | Apple (vidrio esmerilado) | Marca, **Cómo funciona**, **Importar Excel**, **Exportar Excel**, **Nuevo cálculo** y un resumen fijo (anillo, %, estado) que aparece al bajar por la página |
| Riel "Tus indicadores" | Netflix (filas de títulos) | Una tarjeta por indicador con su pastel en miniatura, su suficiencia y su estado. El ejemplo lleva la marca "Ejemplo" |
| Cartel | Netflix (título destacado) | Nombre del indicador, frase en palabras simples y 4 tarjetas: mueve el indicador / así quedarías / te faltaría / si todo se cumple |
| El pastel | Apple (anillos de actividad) | Firma visual del producto (ver abajo) |
| Lo que debes vigilar · Quién aporta más | — | Alertas ordenadas y ranking de iniciativas con "si falla" |
| Taller (pasos 1, 2 y 3) | Notion (bloques editables) | Formularios y tablas que se editan en el mismo lugar; el resultado cambia en vivo |

## Nuevo cálculo y ejemplo

- **Nuevo cálculo** crea un indicador vacío para definir todo desde cero (nombre, qué es mejor, hoy, meta, casos, dolores e iniciativas) y lleva al paso 1.
- Al abrir por primera vez hay **un solo ejemplo lleno** (Churn mensual pospago) marcado como "Ejemplo". Un cálculo vacío ofrece verlo, y el menú **Más › Ver el ejemplo lleno** lo vuelve a traer si se borró, sin tocar los cálculos propios.
- Cuando le pones tu propio nombre al ejemplo, deja de ser ejemplo: pasa a ser tu cálculo.

## El pastel

- Cada porción es un dolor, y su ángulo es su **impacto del dolor**.
- Lo blanco de cada porción es lo que resuelven las iniciativas con la ejecución esperada. El **área** pintada es proporcional a lo resuelto, no el radio.
- La línea punteada marca lo que se resolvería si todo se cumple (Plan). La distancia entre lo blanco y la línea es lo que cuestan los retrasos.
- El anillo exterior es la suficiencia realista. Toma el color del semáforo y da una segunda vuelta si pasa de 100 %.
- Al pasar el mouse por un dolor, su porción se asoma hacia afuera y se iluminan las iniciativas que lo atacan, y al revés.
- Con más de 6 dolores la leyenda muestra los primeros 6 y un botón para ver todos.

## Animaciones

Mismo estilo, más movimiento:

- **Al entrar a un indicador**, el pastel se despliega en el sentido del reloj, luego se llena lo resuelto, el anillo se dibuja y el número del centro cuenta hasta la suficiencia. Si alcanza la meta, el anillo da un destello.
- **Al editar**, el pastel pasa suavemente del valor anterior al nuevo (no se redibuja de golpe).
- **Tarjetas del riel**: entran escalonadas, su pastel se despliega y al pasar el mouse giran un poco.
- **Cómo funciona**: los cuatro pasteles de ejemplo se despliegan uno tras otro.
- Las barras de la leyenda y del ranking crecen una tras otra al entrar.

Reglas: con "reducir movimiento" activado en el sistema no hay animaciones; cada animación tiene un seguro que la termina a tiempo aunque el navegador no pinte (pestaña en segundo plano), y antes de imprimir todas se terminan.

## Excel: exportar e importar

`assets/js/excel.js` escribe y lee .xlsx sin librerías ni internet (un .xlsx es un ZIP con XML). El archivo exportado tiene:

| Hoja | Qué tiene |
|---|---|
| Resumen | Cómo leer el archivo, el comparativo de todos los indicadores con su semáforo y un enlace a cada hoja |
| Una por indicador | Resultado (suficiencia, mueve el indicador, así quedarías, te faltaría), lo que debes vigilar, el ranking y los pasos 1, 2 y 3 |
| Ajustes | Semáforo y umbrales de las alertas |
| Glosario | Las palabras, las etiquetas y cómo volver a abrir el archivo en la calculadora |

- Celdas **azules** = datos (se pueden editar, con listas y validaciones); **grises** = resultados.
- Los dolores van en columnas y son las mismas columnas en el paso 2 y en la matriz del paso 3.
- El importador busca cada dato **por su título**, no por su celda: el archivo sigue funcionando si alguien inserta filas o columnas. Acepta 45 o 45 % y fechas escritas como 31/03/2027.
- La primera celda de cada hoja de indicador (`SUFICIENCIA 2027 · HOJA DE INDICADOR`) la identifica al importar. Si no aparece, se intenta leer como la plantilla.

## Límites

- Tablero y Excel exportado: hasta **15 dolores y 15 iniciativas** por indicador. Con muchos dolores, la tabla del paso 3 se desliza de lado y el número y el nombre de la iniciativa se quedan fijos.
- Plantilla de Excel: 6 dolores y 8 iniciativas por hoja (su diseño es de posiciones fijas).

## Colores (Claro)

| Token | Valor | Uso |
|---|---|---|
| `--rojo` | `#DA291C` | Acento sobre fondo claro: botones, foco, mapa de calor de la matriz |
| `--rojo-vivo` | `#EF3829` | Acento sobre fondo oscuro: marca, tarjeta elegida, estado "No alcanza" |
| `--noche` | `#141414` | Fondo del escenario |
| `--papel` / `--hoja` | `#F2F2F2` / `#FFFFFF` | Fondo y paneles del taller |
| `--ok-vivo` / `--alerta` | `#22C55E` / `#FAB219` | Estados "Alcanza" y "Cerca" |

Los estados siempre llevan **icono + texto**, nunca solo color. En el Excel los iconos son símbolos que Excel sí dibuja (✔ ⚠ ✖ ⚑ ○ ◷ ⇄).

## Tipografía

- **Archivo** condensada (títulos de indicadores y marca): el toque de "cartel" de Netflix.
- **Geist** (interfaz y números): limpia, al estilo Apple.
- **Geist Mono** (códigos como D1 o I3, unidades y etiquetas pequeñas).

Si no hay internet, se usan las fuentes del sistema y todo sigue funcionando.

## Arquitectura

| Archivo | Responsabilidad |
|---|---|
| `assets/js/engine.js` | Motor de cálculo puro. Replica la hoja CÁLCULOS del Excel; funciona en el navegador y en Node |
| `assets/js/ejemplos.js` | Los dos indicadores de ejemplo (iguales a los del Excel); el tablero muestra el primero |
| `assets/js/excel.js` | Escribe y lee .xlsx (ZIP + XML) en el navegador y en Node |
| `assets/js/app.js` | Interfaz: pinta, anima, edita, guarda en el navegador, arma el Excel exportado e importa Excel o JSON |
| `assets/css/app.css` | Estilos con tokens; diseño responsivo con *container queries* |
| `tests/paridad_excel.test.js` | Comprueba que el motor da lo mismo que el Excel (hasta 12 decimales) |
| `tests/excel.test.js` | Comprueba que el .xlsx está bien formado, se relee igual y que se lee la plantilla real |
| `tests/navegador/*.js` | Interacción, importar la plantilla, exportar e importar Excel, y animaciones; las corre `herramientas/probar_tablero.ps1` |
| `herramientas/*.ps1` | Generar y verificar la plantilla de Excel, y probar el tablero |

Reglas que conviene mantener:

- Toda la matemática va en `engine.js`. Si cambia una fórmula, debe cambiar también en el Excel (`herramientas/generar_plantilla.ps1`) y en la prueba de paridad.
- Si cambia un título del Excel exportado (`hojaIndicador` en `app.js`), revisa que el importador (`leerHojaIndicador`) lo siga encontrando.
- Los textos que escribe el usuario se insertan siempre con `textContent`, nunca como HTML.
- Los datos se guardan en el navegador (`localStorage`). Para compartirlos se usa **Exportar Excel** o **Guardar una copia** (.json).
