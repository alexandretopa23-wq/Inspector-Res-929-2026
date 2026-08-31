# Auditoría de conceptos — pérdida del filtro en el motor hidráulico v2

Revisión de cada afirmación introducida en esta sesión, contra física, hidráulica
y realidad de campo. Herramientas: ecuación de Ergun, Bernoulli con pérdidas,
`fluids` 1.3.1, Crane TP-410, skill de componentes hidráulicos.

Agua a 30 °C (vaso climatizado): ρ = 995.7 kg/m³, μ = 7.98×10⁻⁴ Pa·s.

---

## 1. Conversión PSI → m c.a. (`·0.703`)

| | valor |
|---|---|
| 1 PSI con ρ=1000, g=9.80665 | 0.7031 m c.a. |
| 1 PSI con ρ=995.7 (30 °C) | 0.7061 m c.a. |

**VÁLIDO.** El `0.703` del código es la convención estándar de "metro de columna
de agua" (ρ=1000). Usar la densidad real del agua caliente subiría el resultado
0.4 % — despreciable frente a la incertidumbre del resto del cálculo.

## 2. Corrección hidrostática entre los dos manómetros (`+Δz`)

Bernoulli entre la toma de entrada (1) y la de salida (2), **mismo diámetro de
tubería en ambas** (v₁ = v₂):

```
p₁/ρg + z₁ + v₁²/2g  =  p₂/ρg + z₂ + v₂²/2g + h_L
⟹  p₁ − p₂  =  ρg·h_L  −  ρg·(z₁ − z₂)
⟹  h_L [m]  =  (p₁ − p₂)/ρg  +  Δz          con Δz = z₁ − z₂
```

**VÁLIDO — es exactamente lo que hace `deltaPFiltroMedidoV2`.** El manómetro mide
presión estática local; el que quede más abajo carga una columna de agua extra
de ρg·Δz que no es del filtro (1 m ⇒ 1.42 PSI).

- **Caveat menor:** la cabeza de velocidad `v²/2g` solo se cancela si los tubos
  de entrada y salida del filtro tienen el mismo diámetro. Si difieren (2" vs
  2½"), queda un término de 0.1–0.3 PSI a 2 m/s. En la práctica de piscinas
  entrada y salida son del mismo diámetro → se ignora con razón.

## 3. "P_salida ≥ P_entrada no es físico" y "una válvula de bypass aguas abajo no invierte el signo"

**VÁLIDO, con la corrección de esta sesión.**

- A **igual cota**, un lecho (elemento pasivo) solo disipa energía → p₁ > p₂
  siempre en flujo hacia adelante.
- Estrangular una válvula **aguas abajo**: sube la presión en todo el tramo
  presurizado aguas arriba de ella (las dos tomas del filtro por igual) y **baja
  el caudal**. La caída a través del lecho es h_L(Q), monótona creciente en Q, así
  que al bajar Q el ΔP del filtro *disminuye*. Nunca se vuelve negativo por esta
  causa.
- Con los manómetros a **distinta cota**, sí puede leerse ΔP bruto ≤ 0 con un
  filtro limpio: es el término ρg·Δz, no una inversión de energía. Por eso el
  modelo corrige por Δz **antes** de decidir si la lectura es no física.

## 4. Cómo entra la pérdida del filtro a la curva del sistema — CAMBIO POR LA AUDITORÍA

**La versión inicial (`K·Q²`) estaba mal planteada.** Ecuación de Ergun para el
lecho de arena (arena #20, d_p ≈ 0.5 mm, ε ≈ 0.40, lecho 0.6 m):

| Tasa de filtración | U | Re_partícula | ΔP lecho limpio | término inercial (∝Q²) |
|---|---|---|---|---|
| 20 m³/h/m² | 5.6 mm/s | 3.5 | 1.0 m (1.4 PSI) | 6 % |
| 30 m³/h/m² | 8.3 mm/s | 5.2 | 1.5 m (2.2 PSI) | 9 % |
| 37 m³/h/m² | 10.3 mm/s | 6.4 | 1.9 m (2.7 PSI) | 11 % |
| 50 m³/h/m² | 13.9 mm/s | 8.7 | 2.7 m (3.8 PSI) | 14 % |

El lecho está en **régimen viscoso** (Re_p ≈ 3–9): exponente efectivo de ΔP vs Q
≈ **1.1**, no 2. Un lecho **colmatado** —el caso para el que existe este
término— se comporta como **filtración de torta**, ΔP ∝ μ·U·R_torta, también
**lineal**.

Error de anclar en un punto con exponente equivocado (Ergun vs `K·Q²`):

| Q / Q_ref | ΔP real (Ergun) | K·Q² | error |
|---|---|---|---|
| 0.7 | 1.83 PSI | 1.33 PSI | **−28 %** |
| 1.0 | 2.71 | 2.71 | 0 % |
| 1.3 | 3.64 | 4.58 | **+26 %** |

**Corrección aplicada:** el modelo ahora escala la pérdida **linealmente**:

```
h_L(Q) = h_L_ref · (Q / Q_ref)   × (banda 0.8–1.3)
```

Verificación del modelo lineal:

| Escenario | error a Q/Q_ref = 0.7 … 1.3 |
|---|---|
| **Lecho colmatado** (clean + torta) | **< ±1 %** |
| Filtro limpio con internos ∝Q² pesando 50 % | +20 % … −14 % (término chico, <0.6 m absoluto; lo cubre la banda) |

Las toberas y el underdrain internos sí van con Q², pero (a) no se pueden separar
de la lectura única y (b) su peso relativo cae cuando el filtro está sucio, que
es cuando el término importa.

## 5. Piso de sospecha para filtro de arena (`PISO_HL_ARENA = 0.7 m`)

**VÁLIDO y conservador.** Ergun da 1.0 m para el lecho limpio aun a la tasa más
baja del rango normal (20 m³/h/m²). Un h_L corregido por debajo de 0.7 m en un
filtro de arena solo se explica por caudal real muy bajo o instrumentación
dudosa; el modelo lo marca pero **no lo rechaza** (deja entrar el término y
advierte que no se concluya nada sobre el lecho). Correcto: a caudal
genuinamente bajo el lecho sí puede perder <0.7 m.

## 6. "No se cuenta doble con el cruce del manómetro"

**VÁLIDO por construcción.** El término del filtro solo vive en `curvaSistemaV2`
(lazo completo). El cruce del manómetro usa `hfHastaManometroV2`, que acumula
tramos hasta el marcado `t.manometro` (entrada del filtro) — aguas arriba del
lecho. Cero solapamiento. Si el inspector no marca el tramo del manómetro, el
cruce cae a "solo succión" (subestima), pero sigue sin haber doble conteo.

## 7. Δz por tipo de válvula (`filtroValvula`)

**Realidad de campo confirmada por el usuario:** entrada y salida del filtro casi
siempre a cotas distintas (~0.4 m, entrada arriba), salvo con **válvula
selectora multipuerto**, donde entrada y salida son dos bocas del mismo cuerpo →
Δz = 0. El modelo aplica: medido > selectora (0) > batería (0.4) > sin declarar
(0.4, marcado como asumido).

---

## Veredicto

| Afirmación | Estado |
|---|---|
| Conversión 0.703 | ✅ válido |
| Corrección `h_L = ΔP_bruto + Δz` | ✅ válido (derivación exacta) |
| P_sal ≥ P_ent no físico a igual cota | ✅ válido |
| Bypass aguas abajo no invierte el signo | ✅ válido |
| Pérdida del filtro ∝ Q² | ❌ **inválido** → corregido a lineal |
| Pérdida del filtro ∝ Q (lineal) + banda | ✅ válido (Ergun, error <1 % en lecho sucio) |
| Piso 0.7 m para arena | ✅ válido y conservador |
| Sin doble conteo | ✅ válido por estructura |
| Δz típico 0.4 m / 0 con selectora | ✅ válido (Ergun + campo) |

Único cambio de fondo: **K·Q² → escalado lineal con banda ±30 %.**
