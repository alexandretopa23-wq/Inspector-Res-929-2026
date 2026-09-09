# Novedades para la revisión de inspecciones

**Última actualización:** 2026-09-01
**Uso:** documento de referencia cuando se pida *"revisar las inspecciones"*. Contiene
(1) el criterio acordado de a qué área le toca cerrar cada ítem dudoso,
(2) cómo traducir los `respCierre` que no son una de las 4 áreas,
(3) los errores recurrentes que hay que buscar en cada revisión.

Las 4 áreas de cierre son: **Eléctrico · Infraestructura · Térmica e Hidráulica (T&H) · Gestión de sede**.
(`Compartido (los 3)` = solo CHK-003, el plano consolidado).

---

## 1. Reasignaciones acordadas (aplicadas en `Codigo.gs` → `ID_A_AREA`, rev. 2026-09-01)

10 ítems se movieron para que el reparto del informe y del dashboard coincida con
quién cierra de verdad:

| CHK | Ítem | Antes | **Ahora** | Motivo |
|---|---|---|---|---|
| CHK-030 | Reglamento de uso visible para usuarios | Infraestructura | **Gestión de sede** | Es un documento de la sede, no obra |
| CHK-034 | Aforo máximo visible y controlado | Infraestructura | **Gestión de sede** | Documento / control operativo de sede |
| CHK-036 | Revestimiento impermeable y de fácil limpieza (vaso) | Infraestructura | **T&H** | Lo cierra T&H como operador de la piscina |
| CHK-039 | Vértices redondeados en media caña | Infraestructura | **T&H** | Si no hay media caña, la norma pide un **documento de protocolo de limpieza especial de vértices**; le corresponde a T&H como encargado directo de la operación |
| CHK-075 | Tuberías en buen estado, sin fugas ni corrosión | T&H | **T&H** (sin cambio) | Confirmado T&H |
| CHK-110 | Drenajes de piso con rejilla o protector (cuarto de equipos) | T&H | **T&H** (sin cambio) | Confirmado T&H |
| CHK-111 | Extintor disponible y vigente (cuarto de equipos) | T&H | **Infraestructura** | Es dotación, no operación hidráulica |
| CHK-173 | Señalización de acceso incluyente | Infraestructura | **Gestión de sede** | Señalización = sede |
| CHK-175 | Área de primeros auxilios demarcada | Gestión de sede | **Infraestructura** | Demarcar exige intervención física |
| CHK-176 | Área señalizada y accesible (primeros auxilios) | Gestión de sede | **Infraestructura** | Íd. |
| CHK-177 | Punto de agua para consumo humano (primeros auxilios) | Gestión de sede | **Infraestructura** | Requiere punto hidráulico nuevo |
| CHK-189 | Camilla o elemento de traslado disponible | Gestión de sede | **Gestión de sede** (sin cambio) | Es dotación de sede |
| CHK-202 | Ventilación 4–6 vol/hora en sanitarios | Infraestructura | **T&H** | Ventilación mecánica |
| CHK-204 | Aforo máximo determinado por estanque | Gestión de sede | **Gestión de sede** (sin cambio) | Cálculo / documento |
| CHK-207 | Aforo visible para usuarios | Gestión de sede | **Gestión de sede** (sin cambio) | Documento |

**Reparto resultante del checklist (248 ítems):** T&H 122 · Infraestructura 69 · Gestión de sede 31 · Eléctrico 25 · Compartido 1.

### Ajuste 2026-09-08 (cruce del mapa contra el "Enfoque técnico" de la hoja)

Ver `Revision_Asignacion_Items_por_Area.md`. Dos movimientos, confirmados por el equipo:

| CHK | Ítem | Antes | **Ahora** | Motivo |
|---|---|---|---|---|
| CHK-013 | Certificados sensores ópticos o de inmersión | Eléctrico | **T&H** | El sensor de inmersión es gestión operativa de la línea T&H, no instrumentación eléctrica |
| CHK-141 | Acceso restringido (depósito de químicos) | T&H | **Infraestructura** | Control de acceso a un recinto; mismo oficio que CHK-105 (cuarto de equipos) |

Confirmados **sin cambio** en la misma revisión: cap. 27 (CHK-223–230) y cap. 25
(CHK-210–215) siguen en **T&H** — T&H opera las piscinas y hace la gestión de
hallazgos; CHK-059/060 siguen en **T&H** (rejillas de las canaletas perimetrales
del vaso) y CHK-054/055 en **Infraestructura** (rejillas de piso / desagüe, ajenas
al circuito hidráulico); CHK-136/137 siguen en **Eléctrico**.

**Reparto actualizado:** T&H 122 · Infraestructura 70 · Gestión de sede 31 · Eléctrico 24 · Compartido 1.

---

## 2. Traducción de `respCierre` que no son una de las 4 áreas

Cuando en la inspección aparezca un nombre de contratista/tercero en `respCierre`,
para efectos del reparto por área se interpreta así:

| `respCierre` escrito en la inspección | Se cuenta como |
|---|---|
| Casa limpia | **T&H** |
| Aseos la perfección | **T&H** |
| Piscinas | **T&H** |
| Línea *(y "Línea T&H" mal cortado)* | **T&H** |
| Mantenimiento | **La que diga `ID_A_AREA` para ese CHK** (sigue el patrón de las piscinas más actualizadas + las correcciones de la sección 1) |
| AIS | **PENDIENTE de definir** (aparece en Jardín Girardot CHK-001) |

También normalizar errores de digitación al valor exacto del área:
`Linea T&H`, `L nea T&H`, `Lonea T&H`, `Linea T&zh` → **Térmica e Hidráulica** ·
`I fraestructura`, `Knfraestructura`, `Infraestructuras` → **Infraestructura** ·
`Gestion de sede`, `Gestuon de sede` → **Gestión de sede** ·
`Eléctrico`/`Electrico` → **Eléctrico**.

---

## 3. Errores recurrentes a buscar en cada revisión

Tras aplicar las secciones 1 y 2, de ~237 asignaciones con `respCierre` concuerdan
**210** y quedan **27** discrepancias. Casi todas son error del inspector, no del motor:

| CHK | Área correcta (motor) | Lo asignaron a | Veces | Diagnóstico |
|---|---|---|---|---|
| CHK-110 | T&H | Infraestructura | 6 | Error del inspector — es T&H |
| CHK-207 | Gestión de sede | Infraestructura | 5 | Error del inspector — el aforo lo cierra Gestión |
| CHK-201 | Infraestructura | T&H ("Casa limpia") | 4 | **Sin definir**: limpieza/desinfección documentada — ¿la deja Infraestructura o T&H? |
| CHK-204 | Gestión de sede | Infraestructura | 3 | Error del inspector |
| CHK-034 | Gestión de sede | Infraestructura (2) / T&H (1) | 3 | Colas del cambio de sección 1 — reasignar a Gestión |
| CHK-175 | Infraestructura | Gestión de sede | 2 | Colas del cambio de sección 1 — reasignar a Infraestructura |
| CHK-189 | Gestión de sede | Infraestructura | 2 | Error del inspector — es dotación de sede |
| CHK-075 | T&H | Infraestructura | 1 | Error del inspector |
| CHK-111 | Infraestructura | T&H | 1 | Cola del cambio de sección 1 |

**Además, `respCierre` vacío:** muchos ítems en No cumple / Pendiente no traen
responsable de cierre. En la revisión hay que listarlos y exigir que se asignen.

---

## 4. Pendientes de decisión

- **AIS** (contratista en Jardín Girardot, CHK-001): definir a qué área equivale.
- **CHK-201** (limpieza y desinfección documentada de unidades sanitarias):
  el motor lo tiene en Infraestructura; varias sedes lo cierran con "Casa limpia" (T&H).
  Falta acordar el criterio.
- Cerrar la lista de valores permitidos de `respCierre` en la app (selector con las
  4 áreas fijas) para que no vuelvan a entrar nombres de contratista ni typos.
