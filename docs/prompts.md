# Prompts del proyecto

Este archivo guarda los pedidos que dieron origen al proyecto y el prompt final que se usó para construir la plantilla de Excel.

## Fase 1 · Plantilla de Excel

### Pedido original

> Actua como experto en suficiencia de indicadores de gestión para telcos, experto en enseñanza para personas con déficit de atención y lento aprendizaje, para que diseñes una plantilla en Excel que posteriormente será un html, en el cual calcules la suficiencia de los indicadores que propongan los usuarios que deseen consultar la plantilla de suficiencia en Excel. La plantilla deberá contener las siguientes hojas:
>
> 1. Una hoja donde el usuario colocará los indicadores, el ejecutado de esos indicadores, la meta proyectada a 2027, la descripción de los dolores de esos indicadores, porcentaje de los dolores de esos indicadores, describirá las iniciativas que le pegan a esos dolores, una breve descripción de la causa raíz y un porcentaje estimado de cuánto porcentaje de solución tendrá la iniciativa en el dolor. (OJO puede darse el caso de que una iniciativa le resuelva más de un dolor, por eso es importante tener en cuenta cómo los porcentajes al final suman o restan en la suficiencia del indicador.)
> 2. Una hoja del indicador con sus dolores y % del dolor, junto con la data que respalde ese % del dolor. En base a la data de este indicador se deben estimar la suficiencia de cada dolor e iniciativa del indicador que se esté analizando, siempre respaldados con datos. Al final estas hojas de indicadores alimentarán una hoja resumen donde se mostrará las iniciativas y su suficiencia (cuántos pp impactaría en el indicador) y una tabla con el comparativo para entender cuánto mejora si se logra a cabalidad con las iniciativas y pesos propuestos.
>
> Quiero que esta plantilla sea fácil de entender para todo público sin experiencia en suficiencia. Limita a usar fórmulas que sean claras y fáciles de seguir. Al final este proyecto tiene que ser escalable para en una segunda fase poder crear un html intuitivo que funcione para múltiples indicadores y que toda persona fácilmente pueda usarla sin muchos pasos. Usa ejemplos de prueba.

Precisiones que se acordaron después:

- Los clientes se quejan por más de un dolor, así que los % de los dolores no suman 100 %.
- Los dolores salen del análisis de datos de cada indicador. Una iniciativa puede impactar varios dolores, en la proporción que el equipo defina, y esa proporción debe poder modificarse.
- Una iniciativa puede retrasarse y no resolver el dolor principal, o no aportar nada y tener que cambiarse.
- Debe ser lo más preciso posible y, a la vez, fácil de explicar, porque la usará mucha gente.

### Prompt final (versión que corresponde a la plantilla construida)

```
# ROL
Experto en suficiencia de indicadores para telcos y en diseño instruccional para personas con déficit
de atención y aprendizaje lento: frases cortas, un concepto por vez, cero jerga sin explicar.

# OBJETIVO
Construir una plantilla de Excel (.xlsx) que calcule la SUFICIENCIA de indicadores: qué parte de la
brecha hacia la meta 2027 cubren las iniciativas propuestas. Debe ser PRECISA pero EXPLICABLE EN
60 SEGUNDOS a alguien sin experiencia. La usará mucha gente y luego migrará a un HTML (fase 2):
datos en tablas planas, sin celdas combinadas, sin macros, sin lógica oculta.

# MODELO (metáfora oficial: "el pastel")
El pastel es TODO el problema de hoy: la distancia entre el valor actual y el valor ideal.
Cada DOLOR es una porción. Cada INICIATIVA se come un pedazo de las porciones que ataca, y solo
cuenta lo que realmente se ejecuta. La BRECHA es lo que hay que comerse para llegar a la meta.
SUFICIENCIA = lo que se comen las iniciativas ÷ la brecha. Puede pasar de 100 %.

Fórmulas (cada una en su propia columna, con encabezado en lenguaje simple):
1. Valor ideal: el que escriba el usuario; si lo deja vacío, 100 si sube es mejor y 0 si baja es mejor.
2. Problema total (pastel) = |ideal − ejecutado|.
3. Brecha = meta − ejecutado si sube es mejor, ejecutado − meta si baja es mejor (mínimo 0).
4. % Prevalencia = casos con el dolor ÷ casos analizados. La suma puede pasar de 100 % (un caso puede
   tener varios dolores).
5. % Atribución (parte del pastel) = prevalencia ÷ suma de prevalencias (suma 100 %). Editable; aviso
   si el usuario la sobrescribe y no suma 100 %.
6. Matriz iniciativa × dolor (editable): % del dolor que la iniciativa resuelve si se hace completa.
   Vacío = 0. Una iniciativa puede tocar varios dolores con porcentajes distintos.
7. % Ejecución esperada a 2027 por iniciativa (editable, 0–100 %), con estado (En plan / En riesgo /
   Retrasada), responsable, fecha comprometida y evidencia.
8. Lo que deja una iniciativa en un dolor = 1 − MIN(tope, efectividad × ejecución). Tope = 99 %.
9. Lo que queda del dolor = PRODUCTO de lo que deja cada iniciativa (la segunda solo come lo que dejó
   la primera). Dolor resuelto = 1 − lo que queda.
10. Puntos recuperados por dolor = problema total × atribución × dolor resuelto. Total = suma.
11. Suficiencia = puntos recuperados ÷ brecha. Semáforo: ≥100 % verde, ≥70 % amarillo, <70 % rojo
    (umbrales editables).

# ESCENARIOS
- Plan: ejecución 100 % en todas las iniciativas.
- Realista: con las ejecuciones ingresadas.

# ANÁLISIS DE INICIATIVAS (escenario realista)
- Suficiencia si la iniciativa falla: se deshace su multiplicación (lo que queda ÷ lo que deja ella).
- Aporte = puntos con todas − puntos sin ella; también como % de la brecha.
- Etiqueta: REEMPLAZAR (aporte < 5 % de la brecha o ejecución 0 %); EN RIESGO (estado Retrasada o
  En riesgo); APORTA (el resto).
- Revisión: "Sin respaldo" si falta la evidencia; "Falta % de ejecución" si está vacía.
- Dolor: SIN COBERTURA si su atribución > 20 % y queda resuelto < 30 %; Cubierto si resuelto ≥ 30 %;
  Peso bajo en otro caso.
- Ranking de iniciativas por aporte (a igual aporte, primero la que está más arriba) y una frase
  automática: "Aporta X pp. Si falla, la suficiencia baja de A % a B %".

# HOJAS
1. Guia: el pastel, los 4 pasos (Brecha → Dolores → Iniciativas → Resultado), colores, ejemplo vivo,
   glosario, etiquetas, cómo agregar un indicador y reglas para no romper la plantilla.
2. Resumen: comparativo de hasta 10 indicadores (se escribe el nombre de la hoja), gráfico de
   suficiencia Plan vs Realista con línea de 100 %, ranking del indicador elegido.
3. Una hoja por indicador (IND01_Churn, IND02_Instalaciones) y IND_PLANTILLA en blanco para copiar:
   Paso 1 datos y brecha · Paso 2 dolores · Paso 3 matriz de iniciativas · Paso 4 resultado ·
   CÁLCULOS visibles paso a paso. Máximo 6 dolores y 8 iniciativas.
4. Parametros: semáforo, umbrales de alertas, tope y listas desplegables.

# RESTRICCIONES
- Solo funciones simples: SUMA, SI, SI.ERROR, PRODUCTO, SUMAPRODUCTO, INDICE, COINCIDIR, BUSCAR,
  INDIRECTO. Sin LAMBDA/LET, sin matrices dinámicas, sin macros.
- Validaciones visibles con mensaje claro. Celdas azules = las llena el usuario; grises = fórmula.
- Agregar un indicador = copiar IND_PLANTILLA, llenar los pasos y escribir el nombre en el Resumen.

# DATOS DE EJEMPLO (marcados "EJEMPLO – reemplazar")
- Churn mensual pospago (baja es mejor): 2,8 % → 2,0 %, 12.000 bajas analizadas, 5 dolores,
  5 iniciativas; una retrasada al 40 % que deja sin cubrir el dolor principal y una sin respaldo
  con aporte ≈ 0.
- Instalaciones en fecha FTTH (sube es mejor): 82 % → 92 %, 9.000 casos, 4 dolores, 4 iniciativas.

# CRITERIOS DE ACEPTACIÓN
- Churn: 130 % plan / 85 % realista (amarillo). Instalaciones: 126 % / 122 % (verde).
- Ningún dolor supera 100 % resuelto; la atribución suma 100 %.
- Cambiar cualquier % de la matriz o de ejecución recalcula todo.
- Cero errores de fórmula.
```

Cambios frente a la primera versión del prompt:

- El pastel pasó a ser el **problema total**, no la brecha. Con el pastel igual a la brecha, la suficiencia nunca podía llegar a 100 %.
- La hoja *Entrada* se integró en cada hoja de indicador, para que la gente no salte entre hojas.

## Fase 2 · Tablero web

### Pedido

> Procede con la fase 2, tomando en cuenta las recomendaciones. Tienes total libertad para crear algo diferencial; usa la skill /frontend-design, pero garantiza total y fácil entendimiento de la calculadora. En lo posible, una sola vista con todo, como si fuera un dashboard. Usa referencias creativas y de diseño como las usadas por Netflix, Apple y Notion, priorizando el diseño responsivo. Usa la colorimetría de Claro Colombia.

Las recomendaciones que venían de la fase 1:

1. Un archivo de datos por indicador: indicador, dolores y matriz de iniciativas.
2. Reutilizar exactamente las fórmulas de la sección CÁLCULOS.
3. Deslizadores para el % de ejecución, para ver en vivo el efecto de un retraso.
4. Los pasos guiados, uno por paso.
5. Importar el Excel al HTML para no perder lo que la gente ya llenó.

El resultado y las decisiones de diseño están en [diseno.md](diseno.md).
