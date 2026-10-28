# Calculadora de Suficiencia 2027

Calcula si las iniciativas de cada indicador alcanzan para cumplir la meta 2027 y qué pasa si alguna se retrasa o falla.

- **`index.html`**: el tablero web. Ábrelo con doble clic o publícalo con GitHub Pages. No necesita instalar nada.
- **`Plantilla_Suficiencia.xlsx`**: la plantilla de Excel. Usa las mismas fórmulas que el tablero, y el tablero puede importarla.

## Descargar el proyecto en otro computador

- **Sin instalar nada:** en GitHub, botón verde **Code › Download ZIP**. Descomprime el archivo y abre `index.html`.
- **Con Git:** `git clone https://github.com/BrianBarreto93/Calculadora-de-Suficiencias.git`

Qué necesita cada parte:

| Parte | Requisito |
|---|---|
| Tablero (`index.html`) | Un navegador moderno: Edge, Chrome, Firefox o Safari |
| Botón *Importar Excel* | Internet la primera vez que se usa, porque el lector de Excel se descarga de cdnjs |
| Fuentes del tablero | Internet (Google Fonts). Sin internet se usan las fuentes del sistema |
| Plantilla de Excel | Microsoft Excel 2016 o posterior |
| Herramientas `.ps1` | Windows con Excel (y Edge o Chrome para `probar_tablero.ps1`) |
| Prueba del motor | Node.js 18 o posterior (opcional) |

## Cómo se usa

1. Elige un indicador en **Tus indicadores** o crea uno nuevo.
2. En el **taller** (la parte blanca), llena los 3 pasos:
   1. Hoy, meta 2027 y si es mejor que el indicador suba o baje.
   2. Los dolores y sus casos.
   3. Las iniciativas: qué % de cada dolor resuelven y cuánto estará ejecutado en 2027.
3. Lee el resultado arriba (la parte oscura): el pastel, el anillo de suficiencia, lo que debes vigilar y quién aporta más.

Si ya trabajas en Excel, usa **Importar Excel** y elige la plantilla llena.

Los datos se guardan solo en tu navegador. Para compartirlos, usa **Más › Guardar una copia** (.json) y la otra persona la abre con **Importar**.

## Cómo se calcula

- **Problema total (el pastel)** = distancia entre hoy y el valor ideal.
- **Brecha** = distancia entre hoy y la meta 2027.
- Cada dolor tiene una **parte del pastel** (sus casos ÷ la suma de casos de todos los dolores).
- En cada dolor, cada iniciativa deja `1 − efectividad × ejecución`. Lo que queda del dolor es el producto de esos valores: la segunda iniciativa solo resuelve lo que dejó la primera.
- **Puntos recuperados** = problema × parte del pastel × lo resuelto, sumado en todos los dolores.
- **Suficiencia** = puntos recuperados ÷ brecha.

## Estructura

```
index.html                    Tablero web
Plantilla_Suficiencia.xlsx    Plantilla de Excel
assets/js/engine.js           Motor de cálculo (las mismas fórmulas del Excel)
assets/js/ejemplos.js         Indicadores de ejemplo
assets/js/app.js              Interfaz del tablero
assets/css/app.css            Estilos
tests/paridad_excel.test.js   Prueba: motor web = Excel
tests/navegador/              Pruebas del tablero en el navegador
herramientas/                 Scripts para generar y verificar el Excel y probar el tablero
docs/prompts.md               Pedidos y prompt final del proyecto
docs/diseno.md                Decisiones de diseño y arquitectura
```

## Verificar

```
node --test tests/paridad_excel.test.js
powershell -ExecutionPolicy Bypass -File herramientas\probar_tablero.ps1
powershell -ExecutionPolicy Bypass -File herramientas\verificar_plantilla.ps1
```

Para volver a crear la plantilla desde cero (por ejemplo, después de cambiar una fórmula):

```
powershell -ExecutionPolicy Bypass -File herramientas\generar_plantilla.ps1 -Salida Plantilla_nueva.xlsx
```

## Publicar con GitHub Pages

En el repositorio: **Settings › Pages › Deploy from a branch › `main` / root**. El tablero queda en `https://brianbarreto93.github.io/Calculadora-de-Suficiencias/`.
