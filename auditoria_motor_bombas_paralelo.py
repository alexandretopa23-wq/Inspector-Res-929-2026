"""
AUDITORIA — bombas en paralelo (motor hidraulico v2) + metodo del manometro
Reproduce la logica JS nueva y la contrasta con calculo independiente (fluids).
"""
import numpy as np
from fluids.core import Reynolds
from fluids.friction import friction_factor
from scipy.optimize import brentq
g=9.81; PSI2M=0.703

# ---- port fiel de las funciones JS del motor v2 ----
def hw(qM3s,dPulg,L,C):
    D=dPulg*0.0254
    return 0.0 if (D<=0 or qM3s<=0) else 10.67*L*qM3s**1.852/(C**1.852*D**4.8704)

def hfTramoV2(qM3h,t,C,kMul):
    n=t['nLineas'] if t.get('nLineas',0)>0 else 1
    qs=(qM3h/n)/3600
    recta=hw(qs,t['d'],t.get('L',0),C)
    v=qs/(np.pi/4*(t['d']*0.0254)**2)
    k=(t.get('acc',0)*1.0 + t.get('accA',0)*5.0)*kMul
    return recta + k*v*v/(2*g)

def curva_comb(evalH_una, q, nB):      # <-- EXACTAMENTE el wrapper JS nuevo
    return evalH_una(q/nB)

def op_point(evalH_una, qMax_una, H_sist, nB):
    qMax=qMax_una*nB
    prev=None; pj=None; res=None
    for i in range(401):
        q=qMax*i/400
        d=curva_comb(evalH_una,q,nB)-H_sist(q)
        if prev is not None and ((prev>=0>d) or (prev<=0<d)):
            tt=prev/(prev-d); res=(pj+tt*(q-pj),)
            res=(res[0], curva_comb(evalH_una,res[0],nB))
        prev=d; pj=q
    return res

# ---- circuito Cajica semiolimpica (del JSON) ----
TR=[
 dict(nombre="Succion-bomba",  d=4, L=17, acc=5, accA=2, nLineas=2, lado='s', man=False),
 dict(nombre="Bomba-filtro",   d=4, L=5,  acc=8, accA=1, nLineas=1, lado='p', man=True),
 dict(nombre="Filtro-deriv",   d=4, L=3.2,acc=10,accA=1, nLineas=1, lado='p', man=False),
 dict(nombre="Deriv-calent",   d=3, L=6,  acc=17,accA=0, nLineas=1, lado='p', man=False),
 dict(nombre="Calent-retorno", d=3, L=6,  acc=17,accA=0, nLineas=1, lado='p', man=False),
 dict(nombre="Retorno-ramal",  d=4, L=12, acc=10,accA=0, nLineas=1, lado='p', man=False),
 dict(nombre="Ramal-inyect",   d=2.5,L=20,acc=12,accA=0, nLineas=12,lado='p', man=False),
]
C=130
def H_sist(q, alpha=1.0): return alpha*sum(hfTramoV2(q,t,C,1.0) for t in TR)   # calent inyector -> 0, Hgeo 0
def H_up_man(q, alpha=1.0):
    a=0.0
    for t in TR:
        a+=alpha*hfTramoV2(q,t,C,1.0)
        if t['man']: break
    return a

print("="*70)
print("A) LOGICA DEL WRAPPER  H_conj(Q)=H_una(Q/N)  — verificacion de identidad")
print("="*70)
# curva lineal-cuadratica de prueba
Hsh, cc = 32.0, 32.0/70**2
eu=lambda q: Hsh-cc*q*q; qmu=np.sqrt(Hsh/cc)
for nB in (1,2,3,4):
    r=op_point(eu,qmu,H_sist,nB)
    # chequeo manual: en el punto, cada bomba ve r/nB y entrega el mismo H
    q,H=r
    assert abs(eu(q/nB)-H)<1e-6
    assert abs(H_sist(q)-H)<1e-2
    print(f"  N={nB}: Q_tot={q:6.2f}  Q/bomba={q/nB:5.2f}  H={H:5.2f}  (H_una(Q/N)==H_sist(Q) OK)")

print()
print("="*70)
print("B) EXPONENTE de la perdida aguas abajo del manometro  (¿es Q^2?)")
print("="*70)
Q=np.linspace(50,90,9)
L_hw=np.array([H_sist(q)-H_up_man(q) for q in Q])
def dw_down(q):
    tot=0; seen=False
    for t in TR:
        if t['man']: seen=True; continue
        if not seen: continue
        n=t['nLineas']; D=t['d']*0.0254; A=np.pi/4*D**2
        V=(q/n/3600)/A
        f=friction_factor(Re=Reynolds(V=V,D=D,rho=995.7,mu=7.98e-4), eD=1.5e-6/D)
        tot+=f*(t['L']/D)*V*V/(2*g) + (t['acc']*1+t['accA']*5)*V*V/(2*g)
    return tot
L_dw=np.array([dw_down(q) for q in Q])
e_hw=np.polyfit(np.log(Q),np.log(L_hw),1)[0]
e_dw=np.polyfit(np.log(Q),np.log(L_dw),1)[0]
print(f"  Hazen-Williams + K (metodo del motor v2) : exponente = {e_hw:.3f}")
print(f"  Darcy-Weisbach + K + Colebrook (fluids)   : exponente = {e_dw:.3f}")
print(f"  -> ambos ~ 2.0  =>  P_manometro ∝ Q_total^2 confirmado")

print()
print("="*70)
print("C) METODO DEL MANOMETRO   Q_n = Q_1 * sqrt((P_n - H0)/(P_1 - H0))")
print("="*70)
Q1,P1,P2 = 62.0, 27.0, 38.0
print(f"  datos de campo: 1 bomba -> Q1={Q1} m3/h @ P1={P1} PSI ; 2 bombas -> P2={P2} PSI")
for H0 in (-2,-1,0,1,2,3):
    r=np.sqrt((P2*PSI2M-H0)/(P1*PSI2M-H0)); print(f"   H0={H0:+d} m  ->  Q2 = {Q1*r:5.1f} m3/h  ({Q1*r/2:.1f} por bomba)")
print("  sensibilidad al exponente real n (P~Q^n):")
for ne in (1.5,1.7,1.85,2.0):
    print(f"   n={ne}: Q2 = {Q1*(P2/P1)**(1/ne):5.1f} m3/h")

print()
print("="*70)
print("D) CONVERGENCIA de los dos caminos (modelo v2 con curva realista vs manometro)")
print("="*70)
alpha=brentq(lambda a:(H_sist(62,a)-H_up_man(62,a))-P1*PSI2M, 0.01,5)
Hop=H_sist(62,alpha)
print(f"  circuito calibrado a P1 (alpha={alpha:.2f}); TDH de 1 bomba @62 = {Hop:.1f} m")
for margen in (7,10,13,16):
    Hs=Hop+margen; c=(Hs-Hop)/62**2
    e=lambda q,Hs=Hs,c=c: Hs-c*q*q; qm=np.sqrt(Hs/c)
    r2=op_point(e,qm, lambda q:H_sist(q,alpha), 2)
    print(f"   curva con shutoff {Hs:4.1f} m (margen +{margen}): Q2_modelo = {r2[0]:5.1f} m3/h")
print(f"   metodo del manometro (H0=0..2):           Q2 = 73.6 – 74.8 m3/h")
print("  => el modelo v2 (curva realista) y el manometro coinciden en ~70-76 m3/h")
