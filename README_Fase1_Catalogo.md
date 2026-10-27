# AUNADOS · Cotizador quirúrgico — Fase 1: catálogo desde Excel

Prototipo: `aunados-presupuestos.html` (un solo archivo, datos locales; **no** hay base de datos real).

## 1. Cambios implementados
- **Modal «Agregar / Editar ítem» único** (mismo componente para ambos) con búsqueda predictiva en el campo Descripción (por descripción o código, ≥ 2 caracteres, debounce, teclado ↑ ↓ Enter Esc, ARIA combobox/listbox).
- **Selector de tarifario** (financiador · producto · grupo empresa) y resultados con código, descripción y monto (o rango de montos si no hay tarifario elegido).
- **Autocompletado** de código, descripción y monto; los tres siguen **editables**. Si se cambia de concepto y el usuario editó código o monto, se pregunta «Reemplazar / Conservar mis valores».
- Estados: inicial, cargando catálogo, resultados, sin coincidencias («No encontramos coincidencias»), selección, edición manual, validaciones, error de lectura del catálogo, guardado.
- Editar un ítem actualiza el mismo registro (no duplica). El origen queda como `Catálogo`, `Catálogo (editado)` o `Manual`.
- Código en Servicios Clínicos Quirúrgicos: acepta 5 dígitos o 3 letras + 5 dígitos (p. ej. `PRQ02745`).
- Se agregaron las dudas a «Pendientes de definición».

## 2. Mapeo pestañas del Excel → cotizador
| Pestaña | Contenido | Uso |
|---|---|---|
| TABLA DATOS GENERALES (570 504 filas) | Código, descripción, tarifa; clave `Código+Sede+Financiador+Producto+Grupo` | **Catálogo** (8 700 códigos × 200 tarifarios) |
| Sala de Recup | 120 filas, todas idénticas a la hoja general | Ya incluida en el catálogo |
| PROCE. CX + RF + Recargos | Ejemplo RIMAC EPS vs. grupo Telefónica, % AI, personal | Referencia de reglas (no cargada) |
| Reglas de Facturación / Hoja 3 | Recargos 20/25/50 %, honorarios del personal | Referencia de reglas |
| PROC-SIN TARIFA / PROC- TARIFA 0.01 | Casos sin tarifa | Criterio «sin tarifa» |
| SERV. CLINICOS, PNQ - EX-PQ | Notas de regla | Referencia |
| EQUIPOS | Vacía | — |

Categoría de interfaz → familias de código (**propuesta por validar**, el Excel no la declara):
Quirúrgicos = PRQ, QEN · Procedimientos = PRD, END, TRA, BIO, QBI, QTR, CAR · Servicios Clínica = FAC, ADM, ENF, REG · Exámenes Auxiliares = SAN, ORI, LIQ, MIC, PAR, APA, RXS, RXC, ECO, TOM, RMN, MAM, DMO, PET, MNU, PAQ · Farmacia HOSP / SOP y Materiales: sin catálogo (manual). Las demás familias aparecen con «Buscar en todo el tarifario».

## 3. Reglas que sí se implementaron
- La tarifa es la del par código + tarifario elegido; nunca se elige automáticamente entre varios.
- Montos 0 y 0.01 = «sin tarifa»: no se autocompleta el monto.
- Subtotales y totales con las fórmulas existentes (100/50/25 y recargos).
- Recargos Emergencia 20 %, Complejidad 25 %, Laparoscopía 50 %: porcentajes y suma confirmados por la hoja «Reglas de Facturación» (p. ej. 20 % + 50 % → ×1,7).

## 4. Pendientes de validación
1. Relación categoría ↔ familias de código (arriba).
2. Sede/financiador/producto/grupo: el Excel solo trae **Clínica Delgado** y concatena financiador+producto sin separador (la separación mostrada es inferida). Los financiadores demo del prototipo son ficticios y no cruzan.
3. Montos: 215 filas con 0, 4 075 con 0.01 y **16 110 con S/ 1.00** (12 961 PRQ): ¿marcadores o tarifas reales? Tratamiento sin tarifa (Pricing / código similar con Auditoría Presupuestal).
4. 226 pares código+tarifario con montos distintos (duplicados).
5. Recargos: «Reglas» aplica sobre el importe ya distribuido; «Hoja 3» calcula Complejidad/Emergencia sobre la tarifa base; fila 4 de «Reglas» usa `G×K` (49,04) en lugar de `G×(1+K)` (294,26), probable error. Precedencia Complejidad > Emergencia no está en el Excel.
6. Sala de operaciones 0.01 → 30 % de procedimientos y horas; honorarios del personal (anestesiólogo 30 %, instrumentista 6 %, ayudantes 25/15/15 %, perfusionista 15 %); si el honorario del cirujano está incluido.
7. ¿Hay conceptos que obligatoriamente deban salir del catálogo? (hipótesis: no).
8. Fase 1 excluye comentarios, historial, asignación, cancelación, adjuntos e interpretación de 2.ª orden: **no se retiraron** del prototipo porque se pidieron antes; confirmar si se ocultan.

## 5. Pruebas realizadas (Chromium automatizado, sin errores de consola)
Cargar catálogo (8 700 códigos, 200 tarifarios) · estado inicial · búsqueda por descripción («colecis») · búsqueda por código («prq02745») · tarifario elegido y no elegido · selección con teclado · autocompletado de código/descripción/monto · edición manual del monto · cambio de concepto con aviso y «Conservar» · guardar (nuevo, sin duplicar al editar) · sin coincidencias · categoría sin catálogo (Farmacia) · Exámenes Auxiliares («hemograma») · viewport móvil 390 px sin desborde.

## 6. Ejecutar y actualizar el catálogo
- Abrir `aunados-presupuestos.html` en un navegador actual (usa `DecompressionStream`; si no existe se muestra el error y se puede registrar manualmente).
- Regenerar datos: `python3 build_catalog.py <archivo.xlsx> <salida.b64>` (requiere `openpyxl`; no modifica el Excel). El `.b64` se incrusta en el HTML como `CAT_B64`. `split_ctx.py` contiene la separación financiador/producto.
