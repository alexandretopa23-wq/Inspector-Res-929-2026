# Revisión de la asignación de ítems por área (campo de intervención vs. `ID_A_AREA`)

## Resolución (2026-09-08, confirmada por el equipo)

Aplicado en `Codigo.gs` → `ID_A_AREA` (el dashboard y los informes por área leen
ese mismo mapa desde el servidor, no hace falta tocar nada más):

| CHK | Antes | Ahora | Motivo (equipo) |
|---|---|---|---|
| CHK-013 | Eléctrico | **T&H** | El sensor de inmersión es gestión operativa de la línea T&H |
| CHK-141 | T&H | **Infraestructura** | Control de acceso a un recinto; = CHK-105 |

**Sin cambio, confirmado:**
- **Cap. 27 (CHK-223–230)** → sigue **T&H**: la línea T&H se encarga de la operación de las piscinas y de la gestión de hallazgos.
- **Cap. 25 completo (CHK-210–215)** → sigue **T&H**: mismo motivo (operación de piscinas).
- **CHK-059 / CHK-060** → siguen **T&H**: son las rejillas del vaso / canaletas perimetrales.
- **CHK-054 / CHK-055** → siguen **Infraestructura**: rejillas de piso (desagüe), no tienen que ver con el sistema hidráulico ni la recirculación.
- **CHK-136 / CHK-137** → siguen **Eléctrico**: hacen parte de su gestión inmediata (con apoyo de Infraestructura).

Reparto actualizado: **T&H 122 · Infraestructura 70 · Gestión de sede 31 · Eléctrico 24 · Compartido 1.**

---

## Análisis original


**Fecha:** 2026-09-08
**Fuente de "lo que está escrito":** columna *Enfoque técnico* de la hoja de checklist
(gid 322693342) — idéntica al catálogo `items-data` de `index.html`.
**Fuente de "lo que hoy tenemos hecho":** `ID_A_AREA` en `Codigo.gs` (el dashboard
no tiene mapa propio: `_areasDashboard()` → `_areaDe()` lee ese mismo objeto, así
que corregir `Codigo.gs` corrige también el dashboard y los informes por área).

> **Nota sobre la hoja:** el conector de Drive solo devolvió una porción de la
> pestaña de datos (111 filas, todas de CUR / JACUZZI2) y la columna *Responsable
> cierre* llegó vacía en esa porción. El cruce fila-por-fila de `respCierre` real
> sigue siendo el de `Novedades_Revision_Inspecciones.md` §3 (27 discrepancias,
> casi todas error del inspector). Lo de abajo es la revisión del **mapa maestro**
> ítem→área contra el enfoque declarado.

---

## Tier 1 — contradicción clara con el campo de intervención

### Capítulo 27 completo — CHK-223 a CHK-230 → hoy **Térmica e Hidráulica**, debe ser **Gestión de sede**

| CHK | Ítem | Enfoque declarado |
|---|---|---|
| CHK-223 | Matriz de hallazgos por piscina | Gestión |
| CHK-224 | Clasificación de riesgo: alto, medio, bajo | Gestión |
| CHK-225 | Acciones inmediatas definidas | Gestión |
| CHK-226 | Responsable asignado | Gestión |
| CHK-227 | Fecha compromiso | Gestión |
| CHK-228 | Evidencia fotográfica antes/después | Gestión |
| CHK-229 | Validación de cierre técnico | Gestión |
| CHK-230 | **Seguimiento periódico por Facility Management** | Gestión |

Los 8 ítems tienen enfoque "Gestión", el capítulo se llama "Gestión preventiva y
cierre de hallazgos" y CHK-230 nombra explícitamente a Facility Management. No hay
contenido hidráulico ni térmico. Es el grueso de lo mal asignado: 8 de ~10 ítems
"fuera de sitio" están aquí.

---

## Tier 2 — inconsistencia con ítems gemelos casi idénticos

| CHK | Ítem | Hoy | Gemelo ya asignado distinto | Propuesta |
|---|---|---|---|---|
| CHK-059 | Rejillas completas, sin fisuras, deformaciones ni pérdida de rigidez (canaletas) | T&H | CHK-054 "Rejillas de piso estables, completas y sin fracturas" → **Infraestructura**; CHK-062 (mismo cap. 6) → Infraestructura | → **Infraestructura** |
| CHK-060 | Rejillas correctamente apoyadas, niveladas y aseguradas (canaletas) | T&H | CHK-055 "Tapas de cárcamos niveladas, resistentes y aseguradas" → **Infraestructura** | → **Infraestructura** |
| CHK-141 | Acceso restringido (depósito de químicos) | T&H | CHK-105 "Acceso restringido" (cuarto de equipos), texto idéntico → **Infraestructura** | → **Infraestructura** |

En el cap. 6, hoy quedan 058/059/060/063/064/065 en T&H y 061/062 en
Infraestructura: la integridad física de la rejilla (059/060) va con el mismo
oficio que 061/062 y que el cap. 5, no con la función hidráulica de la canaleta
(058/063/064).

---

## Tier 3 — el enfoque es genuinamente mixto; conviene decisión del equipo

| CHK | Ítem | Hoy | Comentario |
|---|---|---|---|
| CHK-136 | Torres o postes fuera de zonas de tránsito | Eléctrico | Es ubicación/tránsito, no trabajo eléctrico → Infraestructura o Gestión de sede |
| CHK-137 | Obstáculos señalizados (cap. 15) | Eléctrico | Señalización → Gestión de sede / Infraestructura |
| CHK-134 | Ausencia de zonas oscuras o puntos ciegos | Eléctrico | Diseño lumínico; defendible en Eléctrico, pero es resultado, no intervención eléctrica |
| CHK-211 | Turnos y responsables definidos (operario) | T&H | Dotación de personal → Gestión de sede |
| CHK-214 | Capacitación en químicos, equipos y emergencias | T&H | Formación de personal → Gestión de sede |
| CHK-013 | Certificados sensores ópticos o de inmersión | Eléctrico | Su gemelo de seguridad CHK-012 (rejillas de fondo) está en T&H; decidir si detección de inmersión es instrumentación (Eléctrico) o sistema antiatrapamiento (T&H) |

El resto del cap. 15 (luminarias, iluminación de emergencia, niveles de lux,
mantenimiento de luminarias) sí es Eléctrico y está bien; CHK-212/213 (rutinas y
bitácora de operación) del cap. 25 sí son T&H.

---

## Lo que se revisó y quedó BIEN (no tocar)

- Cap. 8/9/10/11 (recirculación, filtros, boquillas, desnatadores) → T&H. Correcto.
- Cap. 13 (instalaciones eléctricas) → Eléctrico. Correcto.
- Cap. 14 (gas y calentamiento) → T&H. Correcto (no hay área "Gas"; T&H opera el calentamiento).
- Cap. 7 (drenajes / antiatrapamiento) → T&H, con CHK-070 (botón) tallado a Eléctrico. Correcto.
- Reasignaciones de `Novedades` §1 (CHK-030, 034, 036, 039, 111, 173, 175–177, 202, 009). Coherentes.
- Cap. 1: planos/bitácoras repartidos por sistema (eléctrico→Eléctrico, gas/bombeo/filtros/calentamiento→T&H, plan de seguridad y registros→Gestión de sede). Coherente.
