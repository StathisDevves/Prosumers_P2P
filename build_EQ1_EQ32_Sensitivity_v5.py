"""Deterministic alpha sensitivity with a fixed residential omega profile; v5.
Based on build_EQ1_EQ32_AlphaSensitivity_v4.py. Only the preference specification,
its checks, documentation and output identifiers change; corrected v4 dispatch
and economic evaluation definitions are retained.
36 cases: alpha = 0.20 / 0.25 / 0.30 x A/B/C x 10/15/20/30 kWh.
Hours 1-6: 0.55; 7-10: 1.00; 11-16: 0.80; 17-22: 1.35;
Hour 23: 0.65; Hour 24: 0.55 EUR/kWh. No omega recalibration is used.
The EQ indicators retain their documented scope and economic limitations.
"""
from __future__ import annotations
import argparse, calendar, hashlib, json, math, platform, sys, time, zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
import xml.etree.ElementTree as ET
import numpy as np
import scipy
from scipy.optimize import linprog

NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def colname(k:int)->str:
    s=''
    while k:
        k,r=divmod(k-1,26);s=chr(65+r)+s
    return s

def read_xlsx_rows(path:Path, sheet:str|None=None):
    """Read stored numeric inputs using standard-library OOXML, no Excel engine.
    Formula inputs must have cached values; uncached formulas are rejected.
    """
    with zipfile.ZipFile(path) as z:
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        rel=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        targets={r.attrib['Id']:r.attrib['Target'] for r in rel}
        sheets=wb.find('s:sheets',NS)
        sel=next((x for x in sheets if sheet is None or x.attrib['name']==sheet),None)
        if sel is None: raise ValueError(f'Sheet {sheet!r} not found in {path.name}')
        target=targets[sel.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]
        part=target.lstrip('/') if target.startswith('/') else 'xl/'+target
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            shared=[''.join(n.itertext()) for n in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        out=[]
        for row in ET.fromstring(z.read(part)).findall('.//s:sheetData/s:row',NS):
            r=int(row.attrib['r']);d={}
            for c in row:
                addr=c.attrib['r']; letters=''.join(x for x in addr if x.isalpha());k=0
                for l in letters:k=k*26+ord(l)-64
                typ=c.attrib.get('t'); v=c.find('s:v',NS)
                if typ=='inlineStr': value=''.join(c.find('s:is',NS).itertext())
                elif v is None:
                    if c.find('s:f',NS) is not None: raise ValueError(f'Uncached input formula {path.name}/{sheet}!{addr}')
                    value=None
                elif typ=='s': value=shared[int(v.text)]
                elif typ in ('str','e'): value=v.text
                else:
                    value=float(v.text) if v.text is not None else None
                    if value is not None and value.is_integer():value=int(value)
                d[k]=value
            while len(out)<r:out.append([])
            out[r-1]=[d.get(k) for k in range(1,max(d,default=0)+1)]
        return out

@dataclass(frozen=True)
class Settings:
    pv_area:float=50.0
    pv_efficiency:float=0.18
    eta_ch:float=0.92
    eta_dis:float=0.92
    p_ch:float=3.0
    p_dis:float=3.0
    soc0:float=0.90
    soc_min:float=0.10
    soc_max:float=1.00
    sigma:float=0.00002
    psi:float=0.03
    zeta:float=0.20
    tau_base:float=0.05
    theta:float=0.022
    rho:float=1.0
    b_i:float=0.065
    a_i:float=0.025
    a1_min:float=0.10
    a1_max:float=0.50
    low:float=1/3
    high:float=2/3
    adaptive_low:float=0.30
    adaptive_high:float=0.70
    reserve_base:float=0.10
    reserve_evening:float=0.30
    arb_margin:float=0.01
    grid_charge_limit:float=3.0
    sell_limit:float=3.0
    peak_penalty:float=0.04
    self_reward:float=0.02
    export_factor:float=0.80
    numeric_tol:float=1e-9
    qa_tol:float=1e-6
    solver_tol:float=1e-8
    mip_rel_gap:float=1e-8
    horizon:int=24


def load_inputs(input_file:Path, core_file:Path):
    rows=read_xlsx_rows(input_file,'Solar_8760_2025')
    records=[(r+[None]*10)[:10] for r in rows[1:]]
    if len(records)!=8760: raise ValueError(f'Expected 8760 rows, got {len(records)}')
    a=np.array([[r[i] for i in (0,1,3,4,5,6,8,9)] for r in records],float)
    if not np.isfinite(a).all():raise ValueError('Missing or non-finite hourly input')
    if np.max(np.abs(np.diff(a[:,0])*24-1))>1e-6:raise ValueError('Timestamps are not continuous hourly records')
    if not np.array_equal(a[:,4],np.tile(np.arange(1,25),365)):raise ValueError('Hour labels must repeat 1..24')
    if np.any(a[:,5:7]<0):raise ValueError('Negative irradiation or base load')
    vals={str(r[0]).strip().lower():r[1] for r in read_xlsx_rows(core_file) if len(r)>1 and r[0] is not None}
    cfg=Settings(pv_area=float(vals.get('pv area',50)),eta_ch=float(vals.get('battery efficiency',.92)),eta_dis=float(vals.get('battery efficiency',.92)))
    prices=a[:,7]/1000
    inp={'rows':records,'monthly_input':read_xlsx_rows(input_file,'Monthly_Profile'),
         'ts':a[:,0],'date':a[:,1],'month':a[:,2].astype(int),'day':a[:,3].astype(int),'hour':a[:,4].astype(int),
         'irr':a[:,5],'load':a[:,6],'price':prices,'pv':a[:,5]*cfg.pv_area*cfg.pv_efficiency,
         'daily_mean':np.repeat(prices.reshape(365,24).mean(1),24),
         'daily_min':np.repeat(prices.reshape(365,24).min(1),24),
         'daily_max':np.repeat(prices.reshape(365,24).max(1),24),
         'input_file':str(input_file),'core_file':str(core_file),
         'input_sha256':hashlib.sha256(input_file.read_bytes()).hexdigest(),
         'core_sha256':hashlib.sha256(core_file.read_bytes()).hexdigest()}
    return inp,cfg


FIXED_OMEGA_BY_HOUR = (
    0.55, 0.55, 0.55, 0.55, 0.55, 0.55,
    1.00, 1.00, 1.00, 1.00,
    0.80, 0.80, 0.80, 0.80, 0.80, 0.80,
    1.35, 1.35, 1.35, 1.35, 1.35, 1.35,
    0.65, 0.55,
)

def omega_profile(inp, alpha, legacy=False):
    """Return the same prescribed hourly preferences for every alpha.
    alpha/legacy are kept in the signature for compatibility with the v4 API;
    neither alters omega in this FixedOmega version. Hour labels are 1..24.
    """
    hours = np.asarray(inp['hour'])
    if np.any(hours != hours.astype(int)) or np.any((hours < 1) | (hours > 24)):
        raise ValueError("FixedOmega requires integer hour labels from 1 to 24.")
    return np.asarray(FIXED_OMEGA_BY_HOUR, dtype=float)[hours.astype(int)-1].copy()


def dispatch_ab(inp,cfg,alpha,battery,strategy,legacy=False,legacy_omega=False):
    n=len(inp['load']);p=inp['price'];D=inp['load'];pv=inp['pv'];hrs=inp['hour']
    om=omega_profile(inp,alpha,legacy_omega)
    direct=np.maximum(0,np.minimum.reduce([D,pv,(om-p)/alpha]))
    out={k:np.zeros(n) for k in ('pv_ch','grid_ch','dis_load','dis_sell','soc_beg','soc_end','loss','low','high','reserve','simultaneous_candidate')}
    emin=battery*cfg.soc_min;emax=battery*cfg.soc_max;soc=battery*cfg.soc0
    for t in range(n):
        if legacy and t%24==0:soc=battery*cfg.soc0
        out['soc_beg'][t]=soc
        decay=soc*cfg.sigma if legacy else max(0,soc-emin)*cfg.sigma
        available=soc if legacy else soc-decay
        pv_surplus=max(0,pv[t]-direct[t]);residual=max(0,D[t]-direct[t])
        if strategy=='A':
            lo=inp['daily_min'][t]+cfg.low*(inp['daily_max'][t]-inp['daily_min'][t])
            hi=inp['daily_min'][t]+cfg.high*(inp['daily_max'][t]-inp['daily_min'][t])
            res_frac=cfg.soc_min
            ch=min(cfg.p_ch,pv_surplus,max(0,(emax-available)/cfg.eta_ch)) if pv_surplus>0 and p[t]<=lo else 0.
            dis=min(cfg.p_dis,residual,max(0,(available-emin)*cfg.eta_dis)) if residual>0 and p[t]>=hi else 0.
        else:
            future=p[t:min(t+24,n)];mn=float(future.min());mx=float(future.max())
            lo=mn+cfg.adaptive_low*(mx-mn);hi=mn+cfg.adaptive_high*(mx-mn)
            res_frac=cfg.reserve_evening if 17<=hrs[t]<=22 else cfg.reserve_base
            arb=mx*cfg.eta_ch*cfg.eta_dis > p[t]+cfg.arb_margin+cfg.psi
            if pv_surplus>0:
                potential=pv_surplus if p[t]<=lo or arb else 0.
            else:
                potential=cfg.grid_charge_limit if p[t]<=lo and arb else 0.
            ch=min(cfg.p_ch,potential,max(0,(emax-available)/cfg.eta_ch))
            dis=min(cfg.p_dis,residual+cfg.sell_limit,max(0,(available-battery*res_frac)*cfg.eta_dis)) if p[t]>=hi else 0.
        if ch>cfg.numeric_tol and dis>cfg.numeric_tol:
            out['simultaneous_candidate'][t]=1
            if not legacy:ch=0. # explicit high-price discharge priority; no simultaneous operation
        pv_ch=min(ch,pv_surplus);grid_ch=max(0,ch-pv_ch)
        dis_load=min(dis,residual);dis_sell=max(0,dis-dis_load)
        end=soc-decay+cfg.eta_ch*ch-dis/cfg.eta_dis
        if legacy:end=max(emin,min(emax,end))
        if not legacy and not(emin-1e-7<=end<=emax+1e-7):raise AssertionError(f'SOC bounds {strategy} t={t}: {end}')
        for k,v in [('pv_ch',pv_ch),('grid_ch',grid_ch),('dis_load',dis_load),('dis_sell',dis_sell),('soc_end',end),('loss',decay),('low',lo),('high',hi),('reserve',res_frac)]:out[k][t]=v
        soc=end
    out['direct']=direct;out['solve_seconds']=np.zeros(n);out['solver_mode']=np.zeros(n)
    return out


def solve_storage_window(D,pv,p,hours,cfg,battery,soc0,legacy=False):
    n=len(p);direct=np.minimum(D,pv);resid=np.maximum(D-direct,0);surplus=np.maximum(pv-direct,0)
    sell=np.minimum(p,cfg.export_factor*p+cfg.theta)
    high=p>=np.quantile(p,.7);N=5*n
    c=np.zeros(N)
    c[:n]=cfg.psi-cfg.self_reward+(0 if legacy else sell)
    c[n:2*n]=p+cfg.psi+(0 if legacy else cfg.peak_penalty*high)
    c[2*n:3*n]=-p+cfg.psi-cfg.peak_penalty*high
    c[3*n:4*n]=-sell+cfg.psi
    lower=np.zeros(N);upper=np.zeros(N)
    upper[:n]=np.minimum(cfg.p_ch,surplus)
    upper[n:2*n]=min(cfg.p_ch,cfg.grid_charge_limit)
    upper[2*n:3*n]=np.minimum(cfg.p_dis,resid)
    upper[3*n:4*n]=min(cfg.p_dis,cfg.sell_limit)
    reserve=np.where((hours>=17)&(hours<=22),cfg.reserve_evening,cfg.reserve_base)
    lower[4*n:]=battery*reserve;upper[4*n:]=battery*cfg.soc_max
    eq=np.zeros((n+1,N));rhs=np.zeros(n+1)
    for t in range(n):
        eq[t,4*n+t]=1;eq[t,t]=-cfg.eta_ch;eq[t,n+t]=-cfg.eta_ch
        eq[t,2*n+t]=1/cfg.eta_dis;eq[t,3*n+t]=1/cfg.eta_dis
        if t: eq[t,4*n+t-1]=-(1-cfg.sigma)
        rhs[t]=(soc0*(1-cfg.sigma) if t==0 else 0)+(0 if legacy else cfg.sigma*battery*cfg.soc_min)
    eq[n,5*n-1]=1
    rhs[n]=soc0 if legacy else battery*cfg.soc0
    ub=np.zeros((2*n,N));ubrhs=np.r_[np.full(n,cfg.p_ch),np.full(n,cfg.p_dis)]
    for t in range(n):
        ub[t,t]=ub[t,n+t]=1;ub[n+t,2*n+t]=ub[n+t,3*n+t]=1
    options={'primal_feasibility_tolerance':cfg.solver_tol,'dual_feasibility_tolerance':cfg.solver_tol}
    res=linprog(c,A_ub=ub,b_ub=ubrhs,A_eq=eq,b_eq=rhs,bounds=list(zip(lower,upper)),method='highs',options=options)
    if not res.success:raise RuntimeError(f'LP error: {res.message}; soc={soc0}, n={n}')
    mode=0
    totalch=res.x[:n]+res.x[n:2*n];totaldis=res.x[2*n:3*n]+res.x[3*n:4*n]
    if not legacy and np.any((totalch>cfg.qa_tol)&(totaldis>cfg.qa_tol)):
        # Exact binary-mode formulation, used only when the LP relaxation cycles.
        # LP solutions without overlap are already feasible and optimal for this MIP.
        mode=1; cm=np.r_[c,np.zeros(n)];eqm=np.c_[eq,np.zeros((n+1,n))]
        ubm=np.c_[ub,np.zeros((2*n,n))]
        rhs_m=np.r_[np.zeros(n),np.full(n,cfg.p_dis)]
        for t in range(n):ubm[t,N+t]=-cfg.p_ch;ubm[n+t,N+t]=cfg.p_dis
        res=linprog(cm,A_ub=ubm,b_ub=rhs_m,A_eq=eqm,b_eq=rhs,
                    bounds=list(zip(np.r_[lower,np.zeros(n)],np.r_[upper,np.ones(n)])),
                    integrality=np.r_[np.zeros(N,dtype=int),np.ones(n,dtype=int)],method='highs',
                    options={'mip_rel_gap':cfg.mip_rel_gap,'time_limit':60})
        if not res.success:raise RuntimeError(f'MIP error: {res.message}')
    return res.x[:N].reshape(5,n),mode


def dispatch_c(inp,cfg,battery,legacy=False,progress=False):
    n=len(inp['load']);out={k:np.zeros(n) for k in ('pv_ch','grid_ch','dis_load','dis_sell','soc_beg','soc_end','loss','low','high','reserve','simultaneous_candidate','solve_seconds','solver_mode')}
    out['direct']=np.minimum(inp['load'],inp['pv']);soc=battery*cfg.soc0
    starts=range(0,n,24) if legacy else range(n)
    for t in starts:
        end=min(t+cfg.horizon,n);start=time.perf_counter()
        sol,mode=solve_storage_window(inp['load'][t:end],inp['pv'][t:end],inp['price'][t:end],inp['hour'][t:end],cfg,battery,soc,legacy)
        elapsed=time.perf_counter()-start
        apply=end-t if legacy else 1
        for j in range(apply):
            k=t+j;out['soc_beg'][k]=soc
            for block,name in enumerate(('pv_ch','grid_ch','dis_load','dis_sell','soc_end')):out[name][k]=float(sol[block,j])
            out['loss'][k]=cfg.sigma*(soc if legacy else max(0,soc-battery*cfg.soc_min))
            soc=float(sol[4,j]);out['solve_seconds'][k]=elapsed/apply;out['solver_mode'][k]=mode
            out['reserve'][k]=cfg.reserve_evening if 17<=inp['hour'][k]<=22 else cfg.reserve_base
            ch=out['pv_ch'][k]+out['grid_ch'][k];dis=out['dis_load'][k]+out['dis_sell'][k]
            out['simultaneous_candidate'][k]=float(ch>cfg.qa_tol and dis>cfg.qa_tol)
        if progress and (t%2000==0 or t==n-1):print(f'C {battery:g} kWh: {t+1}/{n} windows',flush=True)
    out['low']=inp['daily_min']+cfg.low*(inp['daily_max']-inp['daily_min'])
    out['high']=inp['daily_min']+cfg.high*(inp['daily_max']-inp['daily_min'])
    return out


HEADERS = [
'Timestamp','Date','Month','Month_Number','Day','Hour',
'Solar availability (kW/m2)','PV generation (kWh)','Base load Demand (kWh)','Spot price (EUR/MWh)',
'omega_t (EUR/kWh)','Spot price (EUR/kWh)','Daily minimum spot (EUR/kWh)','Daily maximum spot (EUR/kWh)',
'Charge threshold display (EUR/kWh)','Discharge threshold display (EUR/kWh)','Daily maximum PV (kWh)','Decision',
'E_self EQ1 (kWh)','E_charge (kWh)','E_discharge (kWh)','SOC_begin (kWh)','SOC_end EQ9/12 (kWh)',
'E_grid_to_load EQ2-EQ4 (kWh)','Elasticity EQ6 proxy (-)','U_cons EQ7 (EUR)','E_actual EQ22 (kWh)','DeltaE (kWh)',
'D_shift (EUR)','E_bat net EQ11 (kWh)','C_deg EQ10 (EUR)','E_sell (kWh)','E_buy total (kWh)',
'W_spot (EUR)','Utility EQ8/29 (EUR)','Daily Utility EQ8/29 (EUR)','SDR capped proxy (-)',
'Charge-weighted reference (EUR/kWh)','Price_P2P EQ13 (EUR/kWh)','E_inj EQ16 (kWh)',
'U_prosumer EQ14 reported (EUR)','a1 EQ18 (-)','Load ratio (-)','gamma EQ18 (EUR/kWh)','tau EQ17 (EUR/kWh)',
'Green benefit (EUR)','SW_total EQ15 reported (EUR)','U_buyer EQ19 (EUR)','A EQ21 (EUR)',
'Price_eq EQ20 diagnostic (EUR/kWh)','CS EQ23 (EUR)','PS EQ24 (EUR)','LC EQ26 (EUR)','QC EQ27 (EUR)',
'SW_P2P EQ25 reported (EUR)','CS_pros EQ28 reported (EUR)','EQ30 not calculated','Daily price peak (EUR/kWh)',
'GS_grid EQ31 proxy (EUR)','SW_grid EQ32 reported (EUR)',
'P2P revenues (EUR)','PS_to_grid arithmetic AU-BG (EUR)','Reference daily price (EUR/kWh)',
'PV_to_battery (kWh)','Grid_to_battery (kWh)','Battery_to_load (kWh)','Battery_to_export (kWh)',
'Usable-energy self-discharge (kWh)','SOC balance residual (kWh)','Site balance residual (kWh)',
'Grid import cost (EUR)','U_P2P net of grid cost (EUR)','Delta price spot-P2P (EUR/kWh)',
'Delta utility AI-AO reported (EUR)','Delta utility AI-BT net (EUR)','SOC_fraction (-)',
'Battery throughput (kWh)','Alpha (EUR/kWh2)','Fixed omega profile residual (EUR/kWh)',
'P2P denominator diagnostic','Active P2P (0/1)','Elasticity defined (0/1)',
'Simultaneous charge-discharge (0/1)','SOC bound violation (kWh)','Terminal inventory change (kWh)',
'Solver seconds for window (s)','Mixed-integer fallback (0/1)','Export-price proxy C (EUR/kWh)',
'No-storage import (kWh)','No-storage export (kWh)','No-storage operating cost (EUR)',
'Operating cost incl degradation (EUR)','Battery operating saving vs no storage (EUR)',
'Battery loss (kWh)','Buyer marginal utility at injection (EUR/kWh)'
]


def evaluate(inp,cfg,alpha,battery,strategy,dispatch,legacy=False,legacy_omega=False):
    n=len(inp['load']);p=inp['price'];D=inp['load'];pv=inp['pv'];o=omega_profile(inp,alpha,legacy_omega)
    s=dispatch['direct'];pc=dispatch['pv_ch'];gc=dispatch['grid_ch'];dl=dispatch['dis_load'];de=dispatch['dis_sell']
    ch=pc+gc;dis=dl+de;sb=dispatch['soc_beg'];se=dispatch['soc_end']
    grid=np.maximum(0,D-s-dl);buy=grid+gc;sell=np.maximum(0,pv-s-pc+de)
    actual=s+dl+(buy if legacy else grid);dE=D-actual
    consumption=o*actual-alpha/2*actual**2;shift=cfg.zeta/2*dE**2
    deg=cfg.psi*(np.abs(ch-dis) if legacy else ch+dis)
    spotcash=p*(sell-buy);uconv=consumption-shift+spotcash-deg
    inj=sell.copy();sdr=np.minimum(1,np.divide(inj,D,out=np.zeros(n),where=D>0));sdr[inj<=cfg.numeric_tol]=0
    wsum=(ch*p).reshape(-1,24).sum(1);csum=ch.reshape(-1,24).sum(1)
    cp=np.repeat(np.divide(wsum,csum,out=np.zeros_like(wsum),where=csum>cfg.numeric_tol),24);cp[inj<=cfg.numeric_tol]=0
    denom=(p-cp)*sdr+cp
    pp=p.copy();valid=(sdr!=0)&(np.abs(cp)>cfg.numeric_tol)&(np.abs(denom)>cfg.numeric_tol)
    pp[valid]=p[valid]*cp[valid]/denom[valid]
    upros=consumption-shift+inj*pp-deg
    dmin=np.repeat(D.reshape(-1,24).min(1),24);dmax=np.repeat(D.reshape(-1,24).max(1),24)
    lratio=np.divide(D,dmax,out=np.zeros(n),where=dmax!=0)
    a1=cfg.a1_min+(cfg.a1_max-cfg.a1_min)*np.divide(D-dmin,dmax-dmin,out=np.zeros(n),where=(dmax-dmin)!=0)
    gam=cfg.tau_base*a1;tau=cfg.tau_base+gam*lratio**2
    green=cfg.theta*cfg.rho*inj;sw=consumption-shift-deg+green-tau*inj
    ub=o*inj-alpha/2*inj**2-pp*inj+green
    a=o*inj-alpha/2*inj**2+green-consumption+shift+deg
    eq=np.full(n,np.nan);active=inj>cfg.numeric_tol;eq[active]=a[active]/(2*inj[active])
    cs=ub.copy();ps=inj*(pp-cfg.b_i-cfg.psi-cfg.zeta*np.abs(dE))-cfg.a_i/2*inj**2
    lc=(o+cfg.theta*cfg.rho-cfg.b_i-cfg.psi-cfg.zeta*np.abs(dE))*inj;qc=(alpha+cfg.a_i)/2*inj**2
    p2psw=lc-qc;csgrid=grid*(o-alpha*pv-p)-alpha/2*grid**2
    gs=np.maximum(0,(inp['daily_max']-p)*dl)
    elasticity=np.divide(-p/alpha,grid,out=np.zeros(n),where=grid>cfg.numeric_tol)
    site=pv+buy+dis-actual-ch-sell
    socres=se-(sb-dispatch['loss']+cfg.eta_ch*ch-dis/cfg.eta_dis)
    continuity=sb[1:]-se[:-1]
    bounds=np.maximum.reduce([np.zeros(n),battery*cfg.soc_min-se,se-battery*cfg.soc_max])
    noimp=np.maximum(D-pv,0);noexp=np.maximum(pv-D,0);nocost=p*(noimp-noexp)
    netcost=p*(buy-sell)+deg
    cols={
        'K':o,'L':p,'M':inp['daily_min'],'N':inp['daily_max'],
        'O':inp['daily_min']+cfg.low*(inp['daily_max']-inp['daily_min']),
        'P':inp['daily_min']+cfg.high*(inp['daily_max']-inp['daily_min']),
        'Q':np.repeat(pv.reshape(-1,24).max(1),24),'S':s,'T':ch,'U':dis,'V':sb,'W':se,'X':grid,'Y':elasticity,'Z':consumption,
        'AA':actual,'AB':dE,'AC':shift,'AD':ch-dis,'AE':deg,'AF':sell,'AG':buy,'AH':spotcash,'AI':uconv,
        'AJ':np.repeat(uconv.reshape(-1,24).sum(1),24),'AK':sdr,'AL':cp,'AM':pp,'AN':inj,'AO':upros,
        'AP':a1,'AQ':lratio,'AR':gam,'AS':tau,'AT':green,'AU':sw,'AV':ub,'AW':a,'AX':eq,'AY':cs,'AZ':ps,
        'BA':lc,'BB':qc,'BC':p2psw,'BD':csgrid,'BF':inp['daily_max'],'BG':gs,'BH':csgrid+gs,
        'BI':inj*pp,'BJ':sw-gs,'BK':inp['daily_mean'],'BL':pc,'BM':gc,'BN':dl,'BO':de,'BP':dispatch['loss'],
        'BQ':socres,'BR':site,'BS':p*buy,'BT':upros-p*buy,'BU':p-pp,'BV':uconv-upros,'BW':uconv-(upros-p*buy),
        'BX':se/battery,'BY':ch+dis,'BZ':np.full(n,alpha),'CA':o-np.asarray(FIXED_OMEGA_BY_HOUR,dtype=float)[inp['hour'].astype(int)-1],
        'CB':denom,'CC':active.astype(float),'CD':(grid>cfg.numeric_tol).astype(float),
        'CE':((ch>cfg.qa_tol)&(dis>cfg.qa_tol)).astype(float),'CF':bounds,'CG':np.full(n,se[-1]-battery*cfg.soc0),
        'CH':dispatch['solve_seconds'],'CI':dispatch['solver_mode'],'CJ':np.minimum(p,cfg.export_factor*p+cfg.theta),
        'CK':noimp,'CL':noexp,'CM':nocost,'CN':netcost,'CO':nocost-netcost,
        'CP':(1-cfg.eta_ch)*ch+(1/cfg.eta_dis-1)*dis+dispatch['loss'],'CQ':o+cfg.theta*cfg.rho-alpha*inj
    }
    qa={'omega_profile_max_abs':float(np.max(np.abs(cols['CA']))),
        'soc_balance_max_abs':float(np.max(np.abs(socres))),
        'site_balance_max_abs':float(np.max(np.abs(site))),
        'soc_continuity_max_abs':float(np.max(np.abs(continuity))),
        'soc_bound_max_violation':float(np.max(bounds)),
        'charge_power_violation':float(max(0,np.max(ch)-cfg.p_ch)),
        'discharge_power_violation':float(max(0,np.max(dis)-cfg.p_dis)),
        'simultaneous_hours':int(cols['CE'].sum()),
        'calculation_nonfinite':int(sum(np.count_nonzero(~np.isfinite(v)) for k,v in cols.items() if k!='AX')),
        'p2p_singular_fallback_hours':int(np.sum((sdr!=0)&(np.abs(cp)>cfg.numeric_tol)&(np.abs(denom)<=cfg.numeric_tol))),
        'negative_p2p_prices':int(np.sum(pp<0)),
        'negative_buyer_utility_active':int(np.sum((ub<0)&active)),
        'negative_buyer_marginal_active':int(np.sum((cols['CQ']<0)&active)),
        'active_p2p_hours':int(active.sum()),
        'grid_near_zero_hours':int((~(grid>cfg.numeric_tol)).sum()),
        'solver_windows':int(np.count_nonzero(dispatch['solve_seconds'])),
        'mip_fallback_windows':int(np.sum(dispatch['solver_mode'])),
        'solver_seconds_total':float(np.sum(dispatch['solve_seconds'])),
        'end_soc_minus_start_kwh':float(se[-1]-battery*cfg.soc0),
        'cs_ps_identity_max_abs':float(np.max(np.abs(cs+ps-p2psw))),
        'equal_utility_identity_max_abs':float(np.max(np.abs((ub-upros)[active]-2*inj[active]*(eq[active]-pp[active])))) if np.any(active) else 0.}
    numeric_checks=['omega_profile_max_abs','soc_balance_max_abs','site_balance_max_abs','soc_continuity_max_abs','soc_bound_max_violation','charge_power_violation','discharge_power_violation','cs_ps_identity_max_abs','equal_utility_identity_max_abs']
    qa['technical_pass']=not legacy and all(qa[k]<=cfg.qa_tol for k in numeric_checks) and qa['simultaneous_hours']==0 and qa['calculation_nonfinite']==0
    return cols,qa


def compute_case(inp,cfg,alpha,battery,strategy,cached_c=None,legacy=False,legacy_omega=False):
    start=time.perf_counter()
    if strategy=='C':dispatch=cached_c if cached_c is not None else dispatch_c(inp,cfg,battery,legacy)
    else:dispatch=dispatch_ab(inp,cfg,alpha,battery,strategy,legacy,legacy_omega)
    cols,qa=evaluate(inp,cfg,alpha,battery,strategy,dispatch,legacy,legacy_omega)
    qa['evaluation_seconds']=time.perf_counter()-start
    return cols,qa


"""Shared workbook layout and formula definitions (no spreadsheet I/O)."""
import calendar
import numpy as np

SUM_COLS=['H','I','S','T','U','X','AF','AG','AI','AO','AU','AV','AY','AZ','BC','BD','BG','BH','BI','BJ','BT','CN','CO','CP']
PARAM_MAP={'alpha':3,'zeta':4,'psi':5,'battery':6,'p_ch':7,'p_dis':8,'eta_ch':9,'eta_dis':10,'soc0':11,'soc_min':12,'soc_max':13,
           'tau_base':15,'theta':16,'rho':17,'b_i':18,'a_i':19,'a1_min':20,'a1_max':21,'low':22,'high':23,
           'grid_charge_limit':32,'sell_limit':33,'peak_penalty':34,'self_reward':35,'reserve_base':37,'reserve_evening':38,
           'adaptive_low':39,'adaptive_high':40,'arb_margin':41,'sigma':42,'numeric_tol':43,'pv_efficiency':44,'pv_area':45,
           'solver_tol':46,'mip_rel_gap':47,'qa_tol':48,'export_factor':49,'horizon':50}

NOTES = [('Purpose', '36 deterministic cases: alpha 0.20/0.25/0.30 x A/B/C x 10/15/20/30 kWh. Same supplied 2025 inputs.'), ('Fixed preference profile', 'omega_t is prescribed by hour and is independent of alpha and prices. Hours 1-6 and 24: 0.55; 7-10: 1.00; 11-16: 0.80; 17-22: 1.35; 23: 0.65 EUR/kWh. The price-plus-load calibration equation is not used.'), ('Reference data', 'The supplied 2025 price/PV/base-load inputs are unchanged. Alpha=0.25 is the baseline of this fixed-profile corrected-v4 experiment, not a reproduction guarantee for the pre-review model.'), ('Version and comparison', 'v5 FixedOmega retains corrected v4 dispatch and economic formulas but replaces the recalibrated omega with the agreed fixed table. Keep these outputs separate from both v4 recalibrated and pre-review outputs.'), ('Excel interpretation', 'Dispatch sheet holds Python controller/solver snapshots. Hourly evaluation and aggregations use Excel formulas. RERUN PYTHON after changing inputs or scenario parameters; Excel alone does not re-optimize dispatch.'), ('Strategy A', 'Calendar-day min/max thresholds (shares 1/3 and 2/3); PV-only charging; discharge to load. Continuous SOC, no daily reset. Discharge takes precedence if charge and discharge rules overlap.'), ('Strategy B', '24-hour look-ahead price min/max thresholds (shares 0.30/0.70); PV/grid charging; battery export; 30% evening discharge reserve, 10% otherwise. Continuous SOC and no simultaneous charge/discharge.'), ('Strategy C', 'Re-optimizes at every hour over the next min(24, remaining) records; applies only the first decision. Cost/proxy-incentive objective, not maximization of the reported utility/welfare indicators.'), ('C objective corrections', 'PV charging includes forgone proxy export revenue. Peak proxy penalty applies to grid charging too. Discharge/charge complementarity is enforced; LP solutions are accepted only when complementarity holds, otherwise binary-mode HiGHS MIP is solved.'), ('C alpha interpretation', 'Base load is fixed and direct PV consumption equals min(load, PV). Neither alpha nor omega enters C dispatch. Dispatch is reused across alpha cases; only ex-post utility/welfare valuation changes.'), ('Forecast assumption', 'B/C use the supplied look-ahead trajectory as perfect information within the window. There is no forecast-error model. A also uses full calendar-day price extrema.'), ('Self-discharge convention', 'Hourly self-discharge acts on usable energy above the minimum non-dispatchable reserve. This explicitly avoids silently creating energy by clamping an under-minimum SOC.'), ('Physical accounting', 'Grid-to-load and grid-to-battery are separated. E_actual excludes grid charging. Base household demand remains served; this is not a newly implemented endogenous load-curtailment model.'), ('Degradation', 'Cost = psi x (charge + discharge), AC-side throughput. No capacity fade is simulated. Both charging and discharging incur cost; the original absolute-net formula is not used.'), ('Initial and terminal energy', 'All start at 90%. C has 90% end-of-window target and 90% year-end SOC. A/B have no forced year-end target. Final-minus-initial inventory is reported; policy comparisons are not pure controller-intelligence tests.'), ('Reported EQ14 and net utility', 'AO retains the original EQ14 expression, which omits grid-purchase cost. BT provides AO minus grid cost. Do not label AO a complete net P2P profit.'), ('Reported welfare formulas', 'AU, AY/AZ/BC, BD/BG/BH retain the source algebra as diagnostic indicators. Economic definitions, proxy costs and common utility calibration require review before journal interpretation. Technical QA is NOT economic/field validation.'), ('PS_to_grid arithmetic', 'BJ = AU - BG is preserved as the requested arithmetic indicator. AU contains P2P injection/premium/network terms; this subtraction alone does not prove a pure pre-P2P welfare decomposition.'), ('P2P transaction scope', 'All available exports are used as a P2P injection proxy in ex-post evaluation; grid-sale and P2P-sale amounts are alternative valuations, not additive simultaneous sales.'), ('P2P participation limitation', 'No buyer demand cap, outside option or individual-rationality constraint is enforced. Negative buyer utility/marginal utility is retained and counted, not clipped or hidden.'), ('P2P price numerical rule', 'Source hyperbolic SDR formula retained, with SDR capped at one. Zero/near-zero denominator falls back to spot price and is counted; negative prices are retained. It is not a validated double-auction clearing rule.'), ('Utility-indifference price', 'AX equals AW/(2*injection) when injection exceeds tolerance, otherwise blank. AX is diagnostic, not the dispatch or implemented transaction price. No feasible-price bounds are silently imposed.'), ('Elasticity limitation', 'Y is the original -price/(alpha x grid-to-load) proxy, reported as zero at zero/near-zero import. It is NOT a newly estimated causal elasticity of optimized net imports. CD marks where this proxy is defined.'), ('Numerical verification', 'Fixed omega profile, balance, SOC continuity/bounds, power, complementarity and two algebraic utility identities are tested. Technical PASS is not economic or field validation.'), ('Source input year', 'Input timestamps form a continuous 2025 calendar. Prices are used as supplied by the user; their historical provenance is not independently re-downloaded here.'), ('Downstream studies', 'This run varies only alpha at fixed omega. No new price year, independent omega variation, degradation sensitivity, Fourier analysis or field validation is introduced.'), ('Solver documentation', 'https://docs.scipy.org/doc/scipy/reference/optimize.linprog-highs.html'), ('Source code', 'Parent: build_EQ1_EQ32_AlphaSensitivity_v4.py. FixedOmega changes are documented in CHANGELOG_FixedOmega.md. Original input and code files are not overwritten.'), ('Hourly convention', 'Use the existing dataset labels 1..24. Hour 24 remains 0.55 EUR/kWh; no shift to a different clock convention is made.')]

def parameters(alpha,bat,st,cfg,inp):
    rows=[[None]*6 for _ in range(55)]
    rows[0]=[f'Alpha {alpha:.2f} | Strategy {st} | Battery {bat:g} kWh',None,None,None,None,None]
    rows[1]=['Parameter','Symbol','Value','Unit','Scope','Notes']
    names={'alpha':('Utility curvature','EUR/kWh2'),'battery':('Battery capacity','kWh'),
      'zeta':('Discomfort coefficient','EUR/kWh2'),'psi':('Degradation coefficient','EUR/kWh'),
      'p_ch':('Maximum charging power','kW'),'p_dis':('Maximum discharging power','kW'),
      'eta_ch':('Charging efficiency','fraction'),'eta_dis':('Discharging efficiency','fraction'),
      'soc0':('Initial/target SOC','fraction'),'soc_min':('Minimum SOC','fraction'),'soc_max':('Maximum SOC','fraction'),
      'pv_efficiency':('PV conversion efficiency','fraction'),'pv_area':('PV area','m2'),'numeric_tol':('Zero guard','kWh / price divisor'),
      'solver_tol':('LP feasibility tolerance','absolute'),'mip_rel_gap':('MIP relative objective gap','fraction'),
      'qa_tol':('Numerical QA tolerance','absolute'),'horizon':('C look-ahead horizon','hours')}
    for k,r in PARAM_MAP.items():
        val=alpha if k=='alpha' else bat if k=='battery' else getattr(cfg,k)
        label,unit=names.get(k,(k.replace('_',' '),'see source code'))
        rows[r-1]=[label,k,val,unit,'All cases unless strategy-specific','Fixed within this run; rerun Python after changing']
    rows[13]=['SDR denominator','SDR demand proxy','Base load Demand','kWh','P2P ex-post','Not a separate measured buyer demand']
    rows[24]=['Model version',None,'v5_FixedOmega_v4_physical',None,None,None]
    rows[25]=['Input dataset',None,inp['input_file'].split('/')[-1],None,None,None]
    rows[26]=['Core assumptions',None,inp['core_file'].split('/')[-1],None,None,None]
    rows[28]=['Dispatch strategy',None,st,None,None,None]
    rows[30]=['Strategy-specific settings',None,None,None,None,None]
    rows[35]=['Terminal constraint',None,'C: 90% each rolling window; A/B: no terminal target',None,None,None]
    rows[51]=['Input SHA256',None,inp['input_sha256'],None,None,None]
    rows[52]=['Core SHA256',None,inp['core_sha256'],None,None,None]
    rows[53]=['Calendar',None,'2025 | 365 days | 8760 one-hour slots',None,None,None]
    rows[54]=['Review status',None,'Technical/algebraic QA; economic limitations in Model_Notes',None,None,None]
    return rows


def dispatch_matrix(inp,vals,qa,dispatch):
    return np.column_stack([dispatch[k] for k in ['direct','pv_ch','grid_ch','dis_load','dis_sell','soc_beg','soc_end','loss','solve_seconds','solver_mode','low','high','reserve','simultaneous_candidate']]).tolist()
DISPATCH_HEADERS=['Direct PV self-use (kWh)','PV charge (kWh)','Grid charge (kWh)','Discharge to load (kWh)','Discharge to export (kWh)',
'SOC start (kWh)','SOC end (kWh)','Self-discharge (kWh)','Solver seconds (s)','MIP fallback (0/1)',
'Lower threshold (EUR/kWh)','Upper threshold (EUR/kWh)','Reserve fraction (-)','Overlapping rule candidates (0/1)']


def hourly_formula_row(r:int,n:int=8760):
    ds=2+((r-2)//24)*24;de=ds+23;end=n+1;p='Parameters!$C$';h='Solar_8760_2025'
    f={
    'K':f"=INDEX('Omega_Profile'!$B$5:$B$28,F{r})",'L':f'=J{r}/1000','M':f'=MIN($L${ds}:$L${de})','N':f'=MAX($L${ds}:$L${de})',
    'O':f'=Dispatch!K{r}','P':f'=Dispatch!L{r}','Q':f'=MAX($H${ds}:$H${de})',
    'R':f'=IF(BM{r}>{p}43,"Grid charge",IF(BO{r}>{p}43,"Battery sell",IF(U{r}>{p}43,"Discharge",IF(T{r}>{p}43,"Charge",IF(AF{r}>{p}43,"Sell",IF(S{r}>{p}43,"Self consume","Buy from grid"))))))',
    'S':f'=Dispatch!A{r}','T':f'=BL{r}+BM{r}','U':f'=BN{r}+BO{r}','V':f'=Dispatch!F{r}','W':f'=Dispatch!G{r}',
    'X':f'=MAX(0,I{r}-S{r}-BN{r})','Y':f'=IF(X{r}<={p}43,0,-L{r}/({p}3*X{r}))',
    'Z':f'=K{r}*AA{r}-{p}3/2*AA{r}^2','AA':f'=S{r}+BN{r}+X{r}','AB':f'=I{r}-AA{r}',
    'AC':f'={p}4/2*AB{r}^2','AD':f'=T{r}-U{r}','AE':f'={p}5*BY{r}',
    'AF':f'=MAX(0,H{r}-S{r}-BL{r}+BO{r})','AG':f'=X{r}+BM{r}','AH':f'=L{r}*(AF{r}-AG{r})',
    'AI':f'=Z{r}-AC{r}+AH{r}-AE{r}','AJ':f'=SUM($AI${ds}:$AI${de})',
    'AK':f'=IF(OR(I{r}<=0,AN{r}<={p}43),0,MIN(1,AN{r}/I{r}))',
    'AL':f'=IF(OR(AN{r}<={p}43,SUM($T${ds}:$T${de})<={p}43),0,SUMPRODUCT($T${ds}:$T${de},$L${ds}:$L${de})/SUM($T${ds}:$T${de}))',
    'AM':f'=IF(OR(AK{r}=0,ABS(AL{r})<={p}43,ABS(CB{r})<={p}43),L{r},L{r}*AL{r}/CB{r})',
    'AN':f'=AF{r}','AO':f'=Z{r}-AC{r}+AN{r}*AM{r}-AE{r}',
    'AP':f'={p}20+({p}21-{p}20)*IF(MAX($I${ds}:$I${de})=MIN($I${ds}:$I${de}),0,(I{r}-MIN($I${ds}:$I${de}))/(MAX($I${ds}:$I${de})-MIN($I${ds}:$I${de})))',
    'AQ':f'=IF(MAX($I${ds}:$I${de})=0,0,I{r}/MAX($I${ds}:$I${de}))','AR':f'={p}15*AP{r}',
    'AS':f'={p}15+AR{r}*AQ{r}^2','AT':f'={p}16*{p}17*AN{r}','AU':f'=Z{r}-AC{r}-AE{r}+AT{r}-AS{r}*AN{r}',
    'AV':f'=K{r}*AN{r}-{p}3/2*AN{r}^2-AM{r}*AN{r}+AT{r}',
    'AW':f'=K{r}*AN{r}-{p}3/2*AN{r}^2+AT{r}-Z{r}+AC{r}+AE{r}',
    'AX':f'=IF(AN{r}<={p}43,"",AW{r}/(2*AN{r}))','AY':f'=AV{r}',
    'AZ':f'=AN{r}*(AM{r}-{p}18-{p}5-{p}4*ABS(AB{r}))-{p}19/2*AN{r}^2',
    'BA':f'=(K{r}+{p}16*{p}17-{p}18-{p}5-{p}4*ABS(AB{r}))*AN{r}',
    'BB':f'=({p}3+{p}19)/2*AN{r}^2','BC':f'=BA{r}-BB{r}',
    'BD':f'=X{r}*(K{r}-{p}3*H{r}-L{r})-{p}3/2*X{r}^2','BE':'=""',
    'BF':f'=N{r}','BG':f'=MAX(0,(BF{r}-L{r})*BN{r})','BH':f'=BD{r}+BG{r}',
    'BI':f'=AN{r}*AM{r}','BJ':f'=AU{r}-BG{r}','BK':f'=AVERAGE($L${ds}:$L${de})',
    'BL':f'=Dispatch!B{r}','BM':f'=Dispatch!C{r}','BN':f'=Dispatch!D{r}','BO':f'=Dispatch!E{r}',
    'BP':f'={p}42*MAX(0,V{r}-{p}6*{p}12)',
    'BQ':f'=W{r}-(V{r}-BP{r}+{p}9*T{r}-U{r}/{p}10)',
    'BR':f'=H{r}+AG{r}+U{r}-AA{r}-T{r}-AF{r}',
    'BS':f'=L{r}*AG{r}','BT':f'=AO{r}-BS{r}','BU':f'=L{r}-AM{r}',
    'BV':f'=AI{r}-AO{r}','BW':f'=AI{r}-BT{r}','BX':f'=W{r}/{p}6','BY':f'=T{r}+U{r}',
    'BZ':f'={p}3','CA':f"=K{r}-INDEX('Omega_Profile'!$B$5:$B$28,F{r})",
    'CB':f'=(L{r}-AL{r})*AK{r}+AL{r}','CC':f'=IF(AN{r}>{p}43,1,0)',
    'CD':f'=IF(X{r}>{p}43,1,0)','CE':f'=IF(AND(T{r}>{p}48,U{r}>{p}48),1,0)',
    'CF':f'=MAX(0,{p}6*{p}12-W{r},W{r}-{p}6*{p}13)',
    'CG':f'=$W${end}-{p}6*{p}11','CH':f'=Dispatch!I{r}','CI':f'=Dispatch!J{r}',
    'CJ':f'=MIN(L{r},{p}49*L{r}+{p}16)','CK':f'=MAX(0,I{r}-H{r})',
    'CL':f'=MAX(0,H{r}-I{r})','CM':f'=L{r}*(CK{r}-CL{r})',
    'CN':f'=L{r}*(AG{r}-AF{r})+AE{r}','CO':f'=CM{r}-CN{r}',
    'CP':f'=(1-{p}9)*T{r}+(1/{p}10-1)*U{r}+BP{r}',
    'CQ':f'=K{r}+{p}16*{p}17-{p}3*AN{r}'
    }
    return [f[colname(k)] for k in range(11,96)]


def aggregation_header(period):
    names=[HEADERS[next(i for i in range(len(HEADERS)) if colname(i+1)==c)] for c in SUM_COLS]
    return [period]+names+['Mean spot (EUR/kWh)','Mean P2P (all hours, EUR/kWh)','Mean omega (EUR/kWh)','Median active Price_eq (EUR/kWh)','Closing SOC (kWh)','Mean elasticity proxy including zeros (-)']


def aggregation_rows(inp,kind):
    rows=[];sh="'Solar_8760_2025'!"
    groups=[np.arange(d*24,(d+1)*24) for d in range(365)] if kind=='Daily' else [np.flatnonzero(inp['month']==m) for m in range(1,13)]
    for i,ix in enumerate(groups):
        a=int(ix[0])+2;b=int(ix[-1])+2
        label=float(inp['date'][ix[0]]) if kind=='Daily' else calendar.month_name[i+1]
        formulas=[f'=SUM({sh}{c}{a}:{c}{b})' for c in SUM_COLS]
        formulas += [f'=AVERAGE({sh}{c}{a}:{c}{b})' for c in ('L','AM','K')]
        formulas += [f'=IF(COUNT({sh}AX{a}:AX{b})=0,"",MEDIAN({sh}AX{a}:AX{b}))',f'={sh}W{b}',f'=AVERAGE({sh}Y{a}:Y{b})']
        rows.append([label]+formulas)
    return rows


def annual_rows(inp):
    s="'Solar_8760_2025'!";out=[]
    lookup={colname(i+1):v for i,v in enumerate(HEADERS)}
    for c in SUM_COLS:out.append([lookup[c],f'=SUM({s}{c}2:{c}8761)','kWh' if '(kWh)' in lookup[c] else 'EUR'])
    out += [
    ['Minimum SOC',f'=MIN({s}W2:W8761)','kWh'],['Maximum SOC',f'=MAX({s}W2:W8761)','kWh'],
    ['Closing minus opening SOC',f'={s}W8761-Parameters!C6*Parameters!C11','kWh'],
    ['Discharge / usable capacity proxy',f'=SUM({s}U2:U8761)/(Parameters!C6*(Parameters!C13-Parameters!C12))','cycles proxy'],
    ['Active P2P hours',f'=SUM({s}CC2:CC8761)','hours'],['Undefined elasticity proxy hours',f'=8760-SUM({s}CD2:CD8761)','hours'],
    ['Mean omega',f'=AVERAGE({s}K2:K8761)','EUR/kWh'],['Minimum omega',f'=MIN({s}K2:K8761)','EUR/kWh'],['Maximum omega',f'=MAX({s}K2:K8761)','EUR/kWh'],
    ['Mean spot price',f'=AVERAGE({s}L2:L8761)','EUR/kWh'],['Solver time (unique C dispatch)',f'=SUM({s}CH2:CH8761)','seconds'],
    ['Simultaneous charging/discharging',f'=SUM({s}CE2:CE8761)','hours'],
    ['Important interpretation','Reported EQ welfare metrics retain source algebra. Read Model_Notes.',''],
    ['Operational objective','C minimizes cost plus proxy incentives; it does not maximize reported utility.',''],
    ['Scope','Alpha sensitivity at fixed omega; no new price year or degradation sensitivity.','']]
    return out

import types
layout = types.SimpleNamespace(**{k: globals()[k] for k in ['SUM_COLS','PARAM_MAP','NOTES','parameters','dispatch_matrix','DISPATCH_HEADERS','hourly_formula_row','aggregation_header','aggregation_rows','annual_rows']})

"""Memory-bounded serialization of the styled workbook template.

Uses standard-library ZIP/XML only. All formula caches are supplied by the
independently evaluated numerical model. The layout/template was authored with
artifact_tool; streaming avoids materialising ~1 million styled cells at once.
"""
import base64,calendar,csv,hashlib,io,json,math,re,time,zipfile
from pathlib import Path
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET
import numpy as np

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
C='http://schemas.openxmlformats.org/drawingml/2006/chart'


def cell(addr,value,style=0,formula=None):
    sty=f' s="{style}"' if style else ''
    if formula is not None:
        f='<x:f>'+escape(formula.lstrip('='))+'</x:f>'
        if value is None or (isinstance(value,(float,np.floating)) and not math.isfinite(value)) or value=='':
            return f'<x:c r="{addr}"{sty} t="str">{f}<x:v></x:v></x:c>'
        if isinstance(value,str):return f'<x:c r="{addr}"{sty} t="str">{f}<x:v>{escape(value)}</x:v></x:c>'
        return f'<x:c r="{addr}"{sty}>{f}<x:v>{float(value):.17g}</x:v></x:c>'
    if value is None:return f'<x:c r="{addr}"{sty}/>' if sty else ''
    if isinstance(value,(bool,np.bool_)):return f'<x:c r="{addr}"{sty} t="b"><x:v>{int(value)}</x:v></x:c>'
    if isinstance(value,(int,float,np.number)):
        if not math.isfinite(float(value)):return f'<x:c r="{addr}"{sty}/>'
        return f'<x:c r="{addr}"{sty}><x:v>{float(value):.17g}</x:v></x:c>'
    return f'<x:c r="{addr}"{sty} t="inlineStr"><x:is><x:t xml:space="preserve">{escape(str(value))}</x:t></x:is></x:c>'


def template_styles(xml):
    root=ET.fromstring(xml); styles={c.attrib['r']:int(c.attrib.get('s','0')) for c in root.findall('.//s:c',NS)}
    rows={int(r.attrib['r']):r.attrib for r in root.findall('.//s:row',NS)}
    return styles,rows


def rowtext(r,cells,height=None):
    h=f' ht="{height}" customHeight="1"' if height is not None else ''
    return f'<x:row r="{r}"{h}>'+''.join(cells)+'</x:row>'


def wrap_sheet(template,rows,rows_count,cols_count,freeze=None,autofilter=False):
    txt=template.decode('utf-8')
    before,tail=txt.split('<x:sheetData>',1);_,after=tail.split('</x:sheetData>',1)
    # Order: dimension, sheetViews, sheetFormatPr, cols, sheetData, autoFilter, ...
    before=re.sub(r'<x:dimension[^>]*/>','',before)
    before=re.sub(r'<x:sheetViews>.*?</x:sheetViews>','',before,flags=re.S)
    insert=f'<x:dimension ref="A1:{colname(cols_count)}{rows_count}"/>'
    if freeze:
        x,y=freeze;top=f'{colname(x+1)}{y+1}'
        insert+=f'<x:sheetViews><x:sheetView workbookViewId="0"><x:pane xSplit="{x}" ySplit="{y}" topLeftCell="{top}" activePane="bottomRight" state="frozen"/></x:sheetView></x:sheetViews>'
    pos=before.index('>',before.index('<x:worksheet'))+1
    before=before[:pos]+insert+before[pos:]
    after=re.sub(r'<x:autoFilter[^>]*/>','',after)
    if autofilter:after=f'<x:autoFilter ref="A1:{colname(cols_count)}{rows_count}"/>'+after
    yield (before+'<x:sheetData>').encode()
    for row in rows:yield row.encode('utf-8')
    yield ('</x:sheetData>'+after).encode('utf-8')


def hourly_rows(inp,cfg,vals,styles):
    yield rowtext(1,[cell(f'{colname(k+1)}1',v,styles.get(f'{colname(k+1)}1',0)) for k,v in enumerate(HEADERS)],48)
    nums=[vals.get(colname(k)) for k in range(11,96)]
    for i,raw in enumerate(inp['rows']):
        r=i+2;ff=layout.hourly_formula_row(r);cells=[]
        for j,v in enumerate(raw,1):
            c=colname(j);f=f'=G{r}*Parameters!$C$44*Parameters!$C$45' if j==8 else None
            if j==8:v=float(inp['pv'][i])
            cells.append(cell(f'{c}{r}',v,styles.get(f'{c}2',0),f))
        for j,f in enumerate(ff,11):
            c=colname(j)
            if c=='R':
                if vals['BM'][i]>cfg.numeric_tol:v='Grid charge'
                elif vals['BO'][i]>cfg.numeric_tol:v='Battery sell'
                elif vals['U'][i]>cfg.numeric_tol:v='Discharge'
                elif vals['T'][i]>cfg.numeric_tol:v='Charge'
                elif vals['AF'][i]>cfg.numeric_tol:v='Sell'
                elif vals['S'][i]>cfg.numeric_tol:v='Self consume'
                else:v='Buy from grid'
            elif c=='BE':v=''
            else:v=float(vals[c][i])
            cells.append(cell(f'{c}{r}',v,styles.get(f'{c}2',0),f))
        yield rowtext(r,cells)


def dmatrix(vals,dispatch):
    return np.column_stack([dispatch[k] for k in ['direct','pv_ch','grid_ch','dis_load','dis_sell','soc_beg','soc_end','loss','solve_seconds','solver_mode','low','high','reserve','simultaneous_candidate']])


def raw_matrix_rows(matrix,styles,heights=None,defaultheight=None,repeat_style_row=2):
    for i,row in enumerate(matrix,1):
        cells=[]
        for j,v in enumerate(row,1):
            c=colname(j);s=styles.get(f'{c}{i}',styles.get(f'{c}{repeat_style_row}',0))
            cells.append(cell(f'{c}{i}',v,s))
        height=heights.get(i,{}).get('ht',defaultheight) if heights else defaultheight
        yield rowtext(i,cells,height)


def aggregate_values(inp,vals,kind):
    groups=[np.arange(d*24,(d+1)*24) for d in range(365)] if kind=='Daily' else [np.flatnonzero(inp['month']==m) for m in range(1,13)]
    out=[]
    for j,ix in enumerate(groups):
        name=float(inp['date'][ix[0]]) if kind=='Daily' else calendar.month_name[j+1]
        row=[name]+[float((inp['pv'] if c=='H' else inp['load'] if c=='I' else vals[c])[ix].sum()) for c in layout.SUM_COLS]
        row += [float(vals[c][ix].mean()) for c in ('L','AM','K')]
        pp=vals['AX'][ix];row += [float(np.nanmedian(pp)) if np.any(np.isfinite(pp)) else '',float(vals['W'][ix[-1]]),float(vals['Y'][ix].mean())]
        out.append(row)
    return out


def aggr_rows(inp,vals,styles,kind):
    headers=layout.aggregation_header(kind)
    yield rowtext(1,[cell(f'{colname(j+1)}1',v,styles.get(f'{colname(j+1)}1',0)) for j,v in enumerate(headers)],48)
    mat=aggregate_values(inp,vals,kind);forms=layout.aggregation_rows(inp,kind)
    for i,(row,fr) in enumerate(zip(mat,forms),2):
        yield rowtext(i,[cell(f'{colname(j+1)}{i}',v,styles.get(f'{colname(j+1)}2',0),fr[j] if j else None) for j,v in enumerate(row)],20)


def annual_values(inp,cfg,alpha,bat,vals):
    out=[float((inp['pv'] if c=='H' else inp['load'] if c=='I' else vals[c]).sum()) for c in layout.SUM_COLS]
    out += [float(vals['W'].min()),float(vals['W'].max()),float(vals['W'][-1]-bat*cfg.soc0),float(vals['U'].sum()/(bat*(cfg.soc_max-cfg.soc_min))),
            int(vals['CC'].sum()),int(len(vals['CD'])-vals['CD'].sum()),float(vals['K'].mean()),float(vals['K'].min()),float(vals['K'].max()),
            float(vals['L'].mean()),float(vals['CH'].sum()),int(vals['CE'].sum())]
    return out


def patch_summary(template,inp,cfg,alpha,bat,st,vals):
    root=ET.fromstring(template);styles,heights=template_styles(template)
    sd=root.find('s:sheetData',NS)
    # Preserve the artifact-authored layout and use cached formula values for the KPI table.
    cellmap={c.attrib['r']:c for c in sd.findall('s:row/s:c',NS)}
    replacements={'A1':cell('A1',f'FIXED OMEGA | 2025 | {st} | {bat:g} kWh | alpha {alpha:.2f}',styles.get('A1',0))}
    ar=layout.annual_rows(inp);values=annual_values(inp,cfg,alpha,bat,vals)
    for j,(label,f,unit) in enumerate(ar,7):
        if f.startswith('='):replacements[f'B{j}']=cell(f'B{j}',values[j-7],styles.get(f'B{j}',0),f)
    text=template.decode()
    for addr,xml in replacements.items():
        pattern=rf'<x:c\b(?=[^>]*\br="{addr}")[^>]*(?:/>|>.*?</x:c>)'
        text,n=re.subn(pattern,lambda _:xml,text,count=1,flags=re.S)
        if n!=1:raise ValueError(f'Summary target not unique: {addr}')
    return text.encode()


def patch_chart(content,monthly):
    ET.register_namespace('c',C)
    root=ET.fromstring(content)
    for ref in root.findall(f'.//{{{C}}}numRef'):
        f=ref.find(f'{{{C}}}f')
        if f is None or 'Monthly_Summary' not in (f.text or ''):continue
        m=re.search(r'\$([A-Z]+)\$2',f.text)
        if not m:continue
        col=0
        for cc in m.group(1):col=col*26+ord(cc)-64
        nc=ref.find(f'{{{C}}}numCache')
        if nc is not None:ref.remove(nc)
        nc=ET.SubElement(ref,f'{{{C}}}numCache');ET.SubElement(nc,f'{{{C}}}formatCode').text='0.000'
        ET.SubElement(nc,f'{{{C}}}ptCount',{'val':'12'})
        for i,row in enumerate(monthly):
            pt=ET.SubElement(nc,f'{{{C}}}pt',{'idx':str(i)});ET.SubElement(pt,f'{{{C}}}v').text=f'{float(row[col-1]):.17g}'
    return ET.tostring(root,encoding='utf-8',xml_declaration=True)


def write_xlsx(path,template_bytes,inp,cfg,alpha,bat,st,vals,qa,dispatch):
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as z:
        parts={n:z.read(n) for n in z.namelist()}
    monthly=aggregate_values(inp,vals,'Monthly')
    dynamic={}
    for k in (2,4,5,6,7,8):
        key=f'xl/worksheets/sheet{k}.xml';sty,heights=template_styles(parts[key])
        if k==2:
            rows=raw_matrix_rows(layout.parameters(alpha,bat,st,cfg,inp),sty,heights);n=55;c=6;freeze=(0,2);flt=False
        elif k==4:
            rows=hourly_rows(inp,cfg,vals,sty);n=8761;c=len(HEADERS);freeze=(6,1);flt=True
        elif k==5:
            def dr():
                yield rowtext(1,[cell(f'{colname(j+1)}1',x,sty0.get(f'{colname(j+1)}1',0)) for j,x in enumerate(layout.DISPATCH_HEADERS)],48)
                for i,row in enumerate(dmatrix(vals,dispatch),2):yield rowtext(i,[cell(f'{colname(j+1)}{i}',float(x),sty0.get(f'{colname(j+1)}2',0)) for j,x in enumerate(row)])
            # Bind local styles before the lazy generator is consumed.
            sty0=sty.copy()
            rows=dr();n=8761;c=14;freeze=(0,1);flt=True
        elif k in (6,7):
            kind='Daily' if k==6 else 'Monthly';rows=aggr_rows(inp,vals,sty,kind);n=366 if k==6 else 13;c=len(layout.aggregation_header(kind));freeze=(1,1);flt=True
        else:
            mat=[['Check or warning','Value','Interpretation','Scope']]
            for check,v in qa.items():
                explanation='Warning/context (not a feasibility failure)' if any(x in check for x in ('negative','fallback','near_zero','inventory','end_soc')) else 'Technical/algebraic check; not economic or field validation'
                mat.append([check,v,explanation,'Computed from Python arrays'])
            rows=raw_matrix_rows(mat,sty,heights,defaultheight=35);n=len(mat);c=4;freeze=(0,1);flt=True
        dynamic[key]=(rows,n,c,freeze,flt)
    # Consume generators inside this function while all case variables are fixed.
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED,compresslevel=5) as out:
        for name,b in parts.items():
            if name in dynamic:
                rows,n,c,freeze,flt=dynamic[name]
                with out.open(name,'w',force_zip64=True) as target:
                    for chunk in wrap_sheet(b,rows,n,c,freeze,flt):target.write(chunk)
            elif name=='xl/worksheets/sheet1.xml':out.writestr(name,patch_summary(b,inp,cfg,alpha,bat,st,vals))
            elif name.endswith('.xml') and '/charts/chart' in name:out.writestr(name,patch_chart(b,monthly))
            elif name=='xl/workbook.xml':
                text=b.decode();text=re.sub(r'<x:calcPr[^>]*/>','',text)
                text=text.replace('</x:workbook>','<x:calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></x:workbook>')
                out.writestr(name,text.encode())
            else:out.writestr(name,b)
    return monthly

TEMPLATE_BASE64 = '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'

def resolve_input(requested,kind):
    path=Path(requested)
    if path.exists():return path
    patterns=['Input Data*.xlsx'] if kind=='data' else ['Input Core assumptions*.xlsx']
    matches=sorted({p for pat in patterns for p in Path.cwd().glob(pat)})
    if len(matches)!=1:raise FileNotFoundError(f'Specify --{ "input-data" if kind=="data" else "core-assumptions"}: found {len(matches)} matching files.')
    return matches[0]


def write_registers(out_dir,records,inp,cfg):
    """Create small, machine-readable run registers, not another output workbook."""
    rows=[];monthly=[]
    for rec in records:
        vals=rec['cols'];row={'case':rec['key'],'alpha':rec['alpha'],'strategy':rec['strategy'],'battery_kWh':rec['battery']}
        for c in layout.SUM_COLS:row['annual_'+c]=float((inp['pv'] if c=='H' else inp['load'] if c=='I' else vals[c]).sum())
        row.update({'omega_min':float(vals['K'].min()),'omega_mean':float(vals['K'].mean()),'omega_max':float(vals['K'].max())})
        row.update(rec['qa']);rows.append(row)
        for m,rowv in enumerate(aggregate_values(inp,vals,'Monthly'),1):
            monthly.append({'case':rec['key'],'alpha':rec['alpha'],'strategy':rec['strategy'],'battery_kWh':rec['battery'],
                            **dict(zip(layout.aggregation_header('Month'),rowv))})
    for name,data in [('Run_Register.csv',rows),('Monthly_Register.csv',monthly)]:
        with open(Path(out_dir)/name,'w',encoding='utf-8-sig',newline='') as f:
            w=csv.DictWriter(f,fieldnames=list(data[0]));w.writeheader();w.writerows(data)
    provenance={'model_version':'v5_FixedOmega_v4_physical','python':platform.python_version(),'numpy':np.__version__,
      'scipy':scipy.__version__,'input_sha256':inp['input_sha256'],'core_sha256':inp['core_sha256'],
      'settings':asdict(cfg),'calibration':'fixed residential hour-band omega; identical for all alpha values; no recalibration',
      'fixed_omega_by_hour_1_to_24':list(FIXED_OMEGA_BY_HOUR),
      'parent_code_sha256':'f47dc845aa2ffa585ee199f5989784b30fd3647662c3c00fa76d8f57e4b59df7',
      'cases':[{k:v for k,v in x.items() if k not in ('cols','dispatch')} for x in records],
      'note':'C dispatch is computed once per capacity and reused across alpha, because alpha and omega do not enter its cost-only optimization. Technical QA is not economic or field validation.'}
    (Path(out_dir)/'Execution_Metadata.json').write_text(json.dumps(provenance,indent=2),encoding='utf-8')
    return rows


def package_cases(out_dir):
    root=Path(out_dir);zips=[]
    for folder in sorted(root.glob('FixedOmega_Alpha_*')):
        if not folder.is_dir():continue
        zp=root.parent/f'{folder.name}_12_Excel_Outputs.zip'
        with zipfile.ZipFile(zp,'w',zipfile.ZIP_DEFLATED,compresslevel=5) as z:
            for path in sorted(folder.glob('*.xlsx')):z.write(path,arcname=f'{folder.name}/{path.name}')
            for name in ['README.md','Run_Register.csv','Monthly_Register.csv','Execution_Metadata.json','CHANGELOG_FixedOmega.md','Omega_Profile.csv','Verification_Report.json']:
                f=root/name
                if f.exists():z.write(f,arcname=name)
        zips.append(zp)
    return root,zips


def write_readme(out_dir):
    text = "# FixedOmega - 36 calculated Excel outputs\n\n"
    text += "Sensitivity: alpha = 0.20 / 0.25 / 0.30 EUR/kWh^2. The same fixed 24-hour omega profile is used in all cases; no price/load recalibration.\n\n"
    text += "Hours 1-6: 0.55; 7-10: 1.00; 11-16: 0.80; 17-22: 1.35; 23: 0.65; 24: 0.55 EUR/kWh.\n\n"
    text += "Three strategies and four capacities (10, 15, 20, 30 kWh) give 36 cases over the same 8,760 supplied 2025 hours.\n\n"
    text += "## Workbook map\nSummary; Parameters; Model_Notes; Solar_8760_2025; Dispatch; Daily_Summary; Monthly_Summary; QA; Monthly_Profile; Omega_Profile.\n\n"
    text += "Hourly evaluation and daily/monthly/annual aggregations contain Excel formulas and precomputed numeric caches. Dispatch is a Python-result snapshot. Editing alpha or omega in Excel alone does not re-optimize physical dispatch; rerun Python.\n\n"
    for title, paragraph in layout.NOTES:
        text += f"### {title}\n{paragraph}\n\n"
    text += "## Reproduction\nRun `python build_EQ1_EQ32_FixedOmega_Sensitivity_v5.py --input-data \"Input Data.xlsx\" --core-assumptions \"Input Core assumptions.xlsx\"`. Only the two input Excel files are required. Runtime dependencies: NumPy and SciPy; standard-library ZIP/XML handles the embedded workbook template. No Google Drive connection is required.\n\n"
    text += "## Interpretation\nDo not mix these FixedOmega results with the previous recalibrated-omega sensitivity charts. Fixed omega does not reverse the v4 physical corrections. Negative economic indicators are retained. The controller objective and economic limitations are unchanged.\n"
    (Path(out_dir)/"README.md").write_text(text, encoding="utf-8")


def run_pipeline(input_file,core_file,out_dir='FixedOmega_36_Outputs',alphas=(.20,.25,.30),strategies=('A','B','C'),batteries=(10,15,20,30),make_zips=True):
    inp,cfg=load_inputs(Path(input_file),Path(core_file));root=Path(out_dir);root.mkdir(parents=True,exist_ok=True)
    if any(not np.isfinite(a) or a<=0 for a in alphas):raise ValueError('Alpha values must be finite and positive.')
    if any(b<=0 for b in batteries):raise ValueError('Battery capacities must be positive.')
    cache={}
    if 'C' in strategies:
        for bat in batteries:
            print(f'Computing hourly rolling C dispatch, {bat:g} kWh...',flush=True)
            cache[bat]=dispatch_c(inp,cfg,bat,progress=True)
    recs=[];template=base64.b64decode(TEMPLATE_BASE64)
    for alpha in alphas:
        for strategy in strategies:
            if strategy not in ('A','B','C'):raise ValueError('Strategies must be A, B or C.')
            for bat in batteries:
                d=cache[bat] if strategy=='C' else dispatch_ab(inp,cfg,alpha,bat,strategy)
                vals,qa=evaluate(inp,cfg,alpha,bat,strategy,d)
                if not qa['technical_pass']:raise RuntimeError(f'Numerical QA failure alpha={alpha}, strategy={strategy}, battery={bat}: {qa}')
                key=f'FixedOmega_Alpha_{alpha:.2f}_Strategy_{strategy}_Battery_{bat:g}kWh'
                folder=root/f'FixedOmega_Alpha_{alpha:.2f}';folder.mkdir(exist_ok=True)
                write_xlsx(folder/(key+'.xlsx'),template,inp,cfg,alpha,bat,strategy,vals,qa,d)
                recs.append({'key':key,'alpha':alpha,'strategy':strategy,'battery':bat,'cols':vals,'qa':qa})
                print('Saved '+key+'.xlsx',flush=True)
    write_readme(root);write_registers(root,recs,inp,cfg)
    with open(root/'Omega_Profile.csv','w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['Hour label','omega_t EUR/kWh'])
        w.writerows(enumerate(FIXED_OMEGA_BY_HOUR,1))
    (root/'CHANGELOG_FixedOmega.md').write_text(
        "# FixedOmega v5 changes from v4\n\n"
        "Only the omega rule, its Excel formulas/profile checks, metadata and output identifiers change. "
        "The corrected v4 controllers, SOC accounting, grid-to-load/grid-to-battery separation, "
        "cost coefficients, economic formulas and tolerances are retained. "
        "Hour 24 = 0.55 EUR/kWh. Recalibration from daily prices and demand is disabled.\n\n"
        "Column K uses the fixed Omega_Profile table; column CA checks the fixed profile. "
        "Column BK remains the daily mean price for reference, not a source of omega. "
        "All cases have hourly and daily/monthly/annual values; negative outputs are retained.\n",
        encoding='utf-8')
    result=package_cases(root) if make_zips else (root,[])
    print(f'Complete: {len(recs)} workbooks. Read Model_Notes and the QA warnings before journal interpretation.',flush=True)
    return result


def main():
    parser=argparse.ArgumentParser(description='36 FixedOmega EQ1-EQ32 output workbooks; corrected v4 dispatch, fixed-profile v5.')
    parser.add_argument('--input-data',default='Input Data.xlsx')
    parser.add_argument('--core-assumptions',default='Input Core assumptions.xlsx')
    parser.add_argument('--output-dir',default='FixedOmega_36_Outputs')
    parser.add_argument('--alphas',type=float,nargs='+',default=[.20,.25,.30])
    parser.add_argument('--strategies',nargs='+',choices=['A','B','C'],default=['A','B','C'])
    parser.add_argument('--battery-sizes',type=float,nargs='+',default=[10,15,20,30])
    args,_=parser.parse_known_args()
    return run_pipeline(resolve_input(args.input_data,'data'),resolve_input(args.core_assumptions,'core'),args.output_dir,args.alphas,args.strategies,args.battery_sizes)

if __name__=='__main__':
    main()
