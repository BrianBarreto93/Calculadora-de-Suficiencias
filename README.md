# Calculadora de Suficiencia 2027

Calcula si las iniciativas de cada indicador alcanzan para cumplir la meta 2027 y qué pasa si alguna se retrasa o falla.

- **`index.html`**: el tablero web. Ábrelo con doble clic o publícalo con GitHub Pages. No necesita instalar nada.
- **`Plantilla_Suficiencia.xlsx`**: la plantilla de Excel. Usa las mismas fórmulas que el tablero, y el tablero puede importarla.

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

El motor de cálculo está en `assets/js/engine.js` y replica la hoja *CÁLCULOS* del Excel. Para comprobar que ambos dan lo mismo:

```
node --test tests/paridad_excel.test.js
```

## Publicar con GitHub Pages

En el repositorio: **Settings › Pages › Deploy from a branch › `main` / root**. El tablero queda en `https://<usuario>.github.io/Calculadora-de-Suficiencias/`.
