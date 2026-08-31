# Anexo B — Pérdida real del filtro (borrador de redacción)

Texto que el informe genera dinámicamente (Codigo.gs, sección **B.5 Verificación
cruzada y contraste con el diseño**) cuando la ficha trae las dos lecturas del
manómetro del filtro: entrada (`presionManometro`), salida (`presionSalidaFiltro`)
y, si están a distinta altura, la diferencia de cota (`desnivelManometrosFiltro`).

---

## El punto clave: un manómetro mide presión estática LOCAL, no energía

La energía total (carga) **siempre** baja en el sentido del flujo a través de un
elemento pasivo — eso no se discute. Pero el manómetro no mide energía total;
mide la presión estática en el punto donde está conectado. Y esa presión incluye
la columna de agua que tiene encima.

Ecuación de energía entre la toma de entrada (1) y la de salida (2), con el mismo
diámetro de tubería en ambas (v₁ = v₂):

```
p₁/ρg + z₁ = p₂/ρg + z₂ + h_L
```

Despejando lo que marcan los dos manómetros:

```
p₁ − p₂ = ρg·(z₂ − z₁) + ρg·h_L
        = ρg·h_L − ρg·Δz          con Δz = z₁ − z₂  (positivo si la entrada está más alta)
```

**El ΔP que se lee en los manómetros = pérdida real del lecho − diferencia de cota.**

Por lo tanto la pérdida real del lecho se recupera sumando de vuelta la cota:

```
h_L [m c.a.] = (P_entrada − P_salida)·0.703  +  Δz [m]
```

- Δz = 0.4 m  →  el manómetro más bajo lee de más ≈ 0.57 PSI (0.40 m c.a.)
- Δz = 1.0 m  →  ≈ 1.42 PSI (1.00 m c.a.)

### De dónde sale Δz en la herramienta (campo `filtroValvula`, CHK-083)

| Situación | Δz aplicado |
|---|---|
| `desnivelManometrosFiltro` medido en sitio | el valor medido (gana siempre) |
| **Batería de válvulas / colector** — bocas de entrada y salida a distinta altura | **0.4 m** (entrada arriba; valor típico) |
| **Válvula selectora multipuerto** — entrada y salida en la misma válvula | **0** (misma cota) |
| No se declaró el tipo de válvula | 0.4 m, marcado como *asumido* — el informe pide declararlo |

El desnivel típico de 0.4 m es el que reportaste: la entrada del filtro casi
siempre queda por encima de la salida, salvo con válvula selectora, donde entrada
y salida son dos bocas del mismo cuerpo a la misma altura.

### Consecuencia práctica

| Situación | ΔP bruto en los manómetros | Pérdida real del lecho |
|---|---|---|
| Filtro sucio (h_L = 10 PSI), entrada 1 m más alta | ≈ 8.6 PSI | 10 PSI → claramente detectable |
| Filtro limpio recién retrolavado (h_L = 1.5 PSI), entrada 1 m más alta | **≈ 0.1 PSI o incluso negativo** | 1.5 PSI |

O sea: **sí, con la entrada 50 cm – 1 m por encima de la salida, un filtro limpio
puede dar ΔP bruto ≤ 0 sin que haya nada mal.** No es que la presión de salida
"supere" a la de entrada en energía — es que el manómetro de salida, al estar más
abajo, tiene más columna de agua encima.

La herramienta ahora pide `desnivelManometrosFiltro` y corrige: solo si la
pérdida **después de corregir** sale ≤ 0 la lectura se declara no física y se
descarta.

---

## B.5 — Filas que genera el informe (con ΔP válido)

| Concepto | Contenido |
|---|---|
| **ΔP bruto del filtro (entrada − salida)** | `X.X PSI (Y.YY m c.a.)` |
| **Corrección por diferencia de cota entre los dos manómetros** *(solo si Δz ≠ 0)* | `±Z.ZZ m c.a.` El manómetro lee presión estática local, no energía: el que quede más abajo gana ≈ 1.4 PSI por metro que no es pérdida del filtro. Pérdida real del lecho = `X.X PSI ± Δz = W.WW m c.a.` |
| **Cómo entra al cálculo** | La pérdida real del lecho se escala **linealmente** con el caudal — `h_L(Q) = h_L_ref · Q / Q_ref`, banda ×0.8–1.3 — y ese término se suma a la curva del sistema. Lineal y no cuadrático porque a la velocidad de filtración de una piscina el lecho de arena está en régimen viscoso (Ergun; ver [Auditoria_perdida_filtro.md](Auditoria_perdida_filtro.md)) y un lecho colmatado se comporta como filtración de torta, ambos ≈ lineales. *(Si h_L ≥ 6 m c.a.:)* La pérdida es alta: el lecho probablemente requiere retrolavado. |
| **Relación con el cruce del manómetro** | No hay doble conteo: el cruce del manómetro solo acumula la pérdida hasta la ENTRADA del filtro; este término cubre solo el salto a través del lecho. |

Si la pérdida neta sale ≤ 0, en lugar de las tres últimas filas el informe pone:

> **Resultado.** La pérdida neta del lecho sale en `W.WW m c.a.`, es decir cero o
> negativa. En un filtro con flujo hacia adelante el lecho solo disipa energía,
> nunca la añade, de modo que ese valor no es físico. La causa está en la
> instrumentación (manómetros descalibrados, tomas de entrada y salida
> intercambiadas) o en que no había flujo por el lecho al momento de leer. Una
> válvula de bypass parcialmente cerrada aguas abajo no produce este efecto: sube
> ambas lecturas por igual y reduce el caudal. La pérdida del filtro no se
> incorpora a la curva del sistema.

---

## Respuesta a la pregunta: ¿una válvula de bypass cerrada parcialmente aguas abajo sube la presión de salida por encima de la de entrada?

**No.** Estrangular una válvula aguas abajo:

1. **Sube las DOS lecturas por igual**, no solo la de salida — la resistencia
   añadida presuriza todo el tramo aguas arriba de la válvula, incluida la
   entrada del filtro.
2. **Baja el caudal**, y como la pérdida del lecho escala con Q², el ΔP real del
   filtro incluso **disminuye** al cerrar la válvula.

Para tener p_salida > p_entrada **a igual cota** habría que entregarle energía al
agua entre los dos puntos (una bomba intercalada). Una válvula, un lecho de arena
o un bypass no lo hacen.

## ¿Y por qué a veces se lee p_salida ≥ p_entrada?

| Causa | Detalle |
|---|---|
| **Diferencia de cota entre los dos manómetros** ← la que preguntaste | Si el manómetro de salida está más abajo que el de entrada, su lectura lleva ≈ 1.4 PSI por metro de columna de agua que no es del filtro. Con un filtro limpio (h_L de 1–2 PSI) y 1 m de desnivel, el ΔP bruto puede salir cero o ligeramente negativo. **Legítimo, no es un error** — por eso hay que capturar `desnivelManometrosFiltro` y corregir. |
| **Manómetros descalibrados** | Dos Bourdon independientes, ±5–10 % de fondo de escala. Si el ΔP real son 2–3 PSI y un manómetro está corrido 3–4 PSI, la lectura se invierte. |
| **Tomas intercambiadas** | Entrada y salida cambiadas, o un manómetro está en realidad sobre otra línea. |
| **Sin flujo por el lecho al leer** | Bomba apagada, multipuerto en "recirculación" o "cerrado", válvula aguas arriba cerrada. |
| **Flujo invertido (retrolavado)** | El agua va al revés; el "ΔP del filtro" no tiene sentido en ese modo. |

## Qué hace la herramienta

- `deltaPFiltroMedidoV2` calcula `h_L = (P_in − P_out)·0.703 + Δz`.
- `kFiltroMedidoV2` devuelve **0** solo cuando `h_L ≤ 0` **después** de la
  corrección de cota. En ese caso el término del filtro no entra a la curva del
  sistema y el Anexo D.3 deja escrito por qué se descartó.
- Si de verdad se sospecha que una válvula está estrangulando el circuito, eso lo
  captura el **cruce del manómetro** (caudal despejado muy por debajo del
  modelado → se adopta como caudal de operación). Vía independiente del ΔP del
  filtro.
