#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera el dashboard de control de facturas a partir del Excel.

Salida: index.html + datos.js, en la carpeta indicada con --web (por
defecto la carpeta actual). Se publica en GitHub Pages, igual que el
dashboard del WMS, para que el servidor de las pantallas lo alcance.

Uso:
    python generar_dashboard_html.py
    python generar_dashboard_html.py "archivo.xlsx" --web .

Requiere: pandas, openpyxl   ->   pip install pandas openpyxl
"""
import sys, os, json, csv, datetime, calendar
import pandas as pd

# ----------------------------------------------------------------------
# Regla del modelo (reconstruida desde Hoja3 y validada contra sus totales)
#   Estado "Cerrada" o "Cerrada para recepcion"
#   Y Fecha de cierre dentro del mes
#
# CIERRE DE MES = TOTAL FACTURADO + PROVISION DEL MES - REVIERTE PROVISION
# El PRESUPUESTO se compara contra el CIERRE DE MES (no contra el total facturado).
# ----------------------------------------------------------------------
ESTADOS = ['Cerrada', 'Cerrada para recepción']
GRUPOS = {
    'PRIMARIA':   ['PRIMARIA'],
    'SECUNDARIA': ['SECUNDARIA'],
    'OTROS':      ['OTROS', 'MATERIA PRIMA', 'BANDEJAS'],
}
HOJA_PARAM = {'PRIMARIA': 'PRIMARIA', 'SECUNDARIA': 'SECUNDARIA', 'OTROS': None}

# Palabras clave para deducir el CANAL desde la Descripcion, en el MISMO
# orden de prioridad que la formula de la columna CANAL del Excel.
# Se usa como respaldo: si la celda trae texto, ese manda.
CANAL_KEYWORDS = ['MATERIA PRIMA', 'SECUNDARIA', 'PRIMARIA', 'BANDEJAS', 'OTROS']


def normalizar_canal(v):
    """Quita espacios y pasa a mayusculas; devuelve None si queda vacio."""
    if v is None:
        return None
    t = str(v).strip().upper()
    return t if t and t != 'NAN' else None


def derivar_canal(descripcion):
    """Replica la formula de la columna CANAL sobre la Descripcion."""
    t = str(descripcion or '').upper()
    for kw in CANAL_KEYWORDS:
        if kw in t:
            return kw
    return None

_args = [a for a in sys.argv[1:] if not a.startswith('--')]
XLSX = _args[0] if len(_args) > 0 else 'control de ingresos facturas.xlsx'
HTML = _args[1] if len(_args) > 1 else 'dashboard_facturas.html'
WEB = None
if '--web' in sys.argv:
    i = sys.argv.index('--web')
    WEB = sys.argv[i + 1] if len(sys.argv) > i + 1 else 'dashboard-web'

MESES_ES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
            'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']


def leer(path):
    ex = pd.read_excel(path, sheet_name='Exported')
    ex = ex[['CANAL', 'Orden', 'Descripción', 'Estado', 'Proveedor',
             'Ordenado', 'Fecha de cierre']].copy()
    # El CANAL puede venir como texto o como formula. Si es formula, pandas
    # lee el valor que Excel dejo guardado, y ese valor NO existe cuando el
    # libro fue escrito por una herramienta que no ejecuta formulas.
    # Por eso: se normaliza lo que haya, y lo que quede vacio se deduce de
    # la Descripcion con la misma regla que usa la formula.
    ex['CANAL'] = ex['CANAL'].map(normalizar_canal)
    faltan = ex['CANAL'].isna()
    if faltan.any():
        ex.loc[faltan, 'CANAL'] = ex.loc[faltan, 'Descripción'].map(derivar_canal)
        recuperadas = int(faltan.sum() - ex['CANAL'].isna().sum())
        print(f'  CANAL: {int(faltan.sum())} filas sin valor guardado, '
              f'{recuperadas} deducidas desde la Descripcion.')
    ex['Fecha de cierre'] = pd.to_datetime(ex['Fecha de cierre'], errors='coerce')
    ex = ex[ex['Estado'].isin(ESTADOS) & ex['Fecha de cierre'].notna()]
    params = {}
    for h in ('PRIMARIA', 'SECUNDARIA'):
        p = pd.read_excel(path, sheet_name=h)
        p['FECHA'] = pd.to_datetime(p['FECHA'], errors='coerce')
        params[h] = {
            d.strftime('%Y-%m'): {
                'presupuesto': None if pd.isna(r['PRESUPUESTO']) else float(r['PRESUPUESTO']),
                'provision':   None if pd.isna(r['PROVISION']) else float(r['PROVISION']),
                'reversa':     None if pd.isna(r['REVERSA PROVISION']) else float(r['REVERSA PROVISION']),
            }
            for d, (_, r) in zip(p['FECHA'], p.iterrows()) if not pd.isna(d)
        }
    return ex, params


def construir(ex, params):
    ex = ex.copy()
    ex['mes'] = ex['Fecha de cierre'].dt.strftime('%Y-%m')
    ex['dia'] = ex['Fecha de cierre'].dt.day

    canal_a_grupo = {c: g for g, cs in GRUPOS.items() for c in cs}
    ex['grupo'] = ex['CANAL'].map(canal_a_grupo)

    meses = sorted(set(ex['mes']) | set(params['PRIMARIA']) | set(params['SECUNDARIA']))
    out = {}
    for m in meses:
        sub = ex[ex['mes'] == m]
        anio, mm = int(m[:4]), int(m[5:])
        ndias = calendar.monthrange(anio, mm)[1]
        filas, series = [], {}
        for g in GRUPOS:
            s = sub[sub['grupo'] == g]
            total = float(s['Ordenado'].sum())
            hoja = HOJA_PARAM[g]
            pr = params[hoja].get(m, {}) if hoja else {}
            presup, prov, rev = pr.get('presupuesto'), pr.get('provision'), pr.get('reversa')
            cierre = total - (rev or 0) + (prov or 0)
            filas.append({
                'cuenta': g, 'cierre': cierre, 'presupuesto': presup,
                'reversa': rev, 'provision': prov, 'total': total,
                'n': int(len(s)),
                'avance': (cierre / presup) if presup else None,
                'saldo': (presup - cierre) if presup else None,
            })
            if g in ('PRIMARIA', 'SECUNDARIA'):
                # ingreso de facturas: monto facturado acumulado dia a dia.
                # Sin ajuste de provision; termina en TOTAL FACTURAS MES.
                por_dia = s.groupby('dia')['Ordenado'].sum()
                acum, running = [], 0.0
                for d in range(1, ndias + 1):
                    running += float(por_dia.get(d, 0.0))
                    acum.append(running)
                series[g] = acum

            det = (s.groupby('Proveedor')['Ordenado']
                     .agg(['count', 'sum']).reset_index()
                     .sort_values('sum', ascending=False))
            filas[-1]['detalle'] = [
                {'prov': r['Proveedor'], 'n': int(r['count']), 'monto': float(r['sum'])}
                for _, r in det.iterrows()]

        sin_canal = sub[sub['grupo'].isna()]
        out[m] = {
            'etiqueta': f'{MESES_ES[mm-1]} {anio}',
            'ndias': ndias,
            'filas': filas,
            'series': series,
            'presupuestos': {g: (params[HOJA_PARAM[g]].get(m, {}) or {}).get('presupuesto')
                             for g in ('PRIMARIA', 'SECUNDARIA')},
            'sin_canal': {'n': int(len(sin_canal)),
                          'monto': float(sin_canal['Ordenado'].sum())},
        }
    return out


TPL = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="refresh" content="600">
<title>Control de facturas</title>
<style>
  :root{
    color-scheme: light;
    --plane:#f9f9f7; --surface:#fcfcfb;
    --ink:#0b0b0b; --ink-2:#52514e; --muted:#898781;
    --grid:#e1e0d9; --axis:#c3c2b7; --ring:rgba(11,11,11,.10);
    --s1:#2a78d6; --s2:#eb6834;
    --good:#0ca30c; --warn:#fab219; --crit:#d03b3b;
  }
  @media (prefers-color-scheme: dark){
    :root:not([data-theme="light"]){
      color-scheme: dark;
      --plane:#0d0d0d; --surface:#1a1a19;
      --ink:#fff; --ink-2:#c3c2b7; --muted:#898781;
      --grid:#2c2c2a; --axis:#383835; --ring:rgba(255,255,255,.10);
      --s1:#3987e5; --s2:#d95926;
    }
  }
  :root[data-theme="dark"]{
    color-scheme: dark;
    --plane:#0d0d0d; --surface:#1a1a19;
    --ink:#fff; --ink-2:#c3c2b7; --muted:#898781;
    --grid:#2c2c2a; --axis:#383835; --ring:rgba(255,255,255,.10);
    --s1:#3987e5; --s2:#d95926;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--plane);color:var(--ink);
       font:14px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;}
  .wrap{max-width:1180px;margin:0 auto;padding-block:28px 56px;padding-left:20px;padding-right:20px}
  /* pantalla de monitoreo: aprovecha el ancho y sube el tamano de letra */
  @media (min-width:1600px){
    body{font-size:16px}
    .wrap{max-width:1560px;padding-block:36px 64px;padding-left:34px;padding-right:34px}
    h1{font-size:28px}
    .tile .val{font-size:36px}
    .tile .lbl{font-size:13px}
    .tile .meta{font-size:14px}
    h2{font-size:16px}
    table{font-size:14px}
    th,td{padding:11px 13px}
  }
  @media (min-width:2200px){
    body{font-size:19px}
    .wrap{max-width:2000px}
    h1{font-size:34px}
    .tile .val{font-size:46px}
    table{font-size:16.5px}
    th,td{padding:14px 16px}
  }
  header{display:flex;flex-wrap:wrap;gap:16px;align-items:flex-end;justify-content:space-between;margin-bottom:6px}
  h1{font-size:22px;margin:0;letter-spacing:-.01em}
  .sub{color:var(--muted);font-size:12.5px;margin:2px 0 0}
  select{font:inherit;padding:7px 12px;border-radius:8px;border:1px solid var(--axis);
         background:var(--surface);color:var(--ink);min-width:190px}
  label.sel{display:flex;gap:9px;align-items:center;font-size:12.5px;color:var(--ink-2)}
  .card{background:var(--surface);border:1px solid var(--ring);border-radius:12px;padding:18px 20px}
  .tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px;margin:22px 0}
  .tile .lbl{font-size:11.5px;letter-spacing:.06em;text-transform:uppercase;color:var(--muted)}
  .tile .val{font-size:27px;font-weight:600;margin:7px 0 2px;letter-spacing:-.02em}
  .tile .meta{font-size:12.5px;color:var(--ink-2)}
  .meter{height:6px;border-radius:3px;background:var(--grid);margin-top:13px;overflow:hidden}
  .meter i{display:block;height:100%;border-radius:3px}
  h2{font-size:14px;letter-spacing:.04em;text-transform:uppercase;color:var(--ink-2);
     margin:30px 0 12px;font-weight:600}
  .scroll{overflow-x:auto}
  table{border-collapse:collapse;width:100%;font-size:12.5px}
  th,td{padding:9px 10px;text-align:right;border-bottom:1px solid var(--grid);white-space:nowrap}
  th:first-child,td:first-child{text-align:left;white-space:normal}
  thead th{font-size:11px;letter-spacing:.05em;color:var(--muted);
           font-weight:600;border-bottom:1px solid var(--axis)}
  tbody tr:last-child td{border-bottom:none}
  tfoot td{font-weight:600;border-top:1px solid var(--axis);border-bottom:none}
  td.num,th.num{font-variant-numeric:tabular-nums}
  .legend{display:flex;gap:20px;flex-wrap:wrap;font-size:12.5px;color:var(--ink-2);margin-bottom:8px}
  .legend b{display:inline-block;width:22px;height:0;border-top-width:2.5px;border-top-style:solid;
            vertical-align:middle;margin-right:7px}
  .cols{display:grid;grid-template-columns:repeat(auto-fit,minmax(310px,1fr));gap:14px}
  .note{font-size:12px;color:var(--muted);margin-top:10px}
  .flag{display:inline-flex;align-items:center;gap:7px;font-size:12.5px;padding:7px 12px;
        border-radius:8px;border:1px solid var(--ring);margin-top:14px}
  .flag.ok{color:var(--good)} .flag.bad{color:var(--crit)}
  svg{display:block;width:100%;height:auto;overflow:visible}
  .tip{position:fixed;pointer-events:none;opacity:0;transition:opacity .1s;
       background:var(--surface);border:1px solid var(--axis);border-radius:8px;
       padding:9px 12px;font-size:12.5px;box-shadow:0 4px 16px rgba(0,0,0,.14);z-index:9}
  .tip .r{display:flex;gap:10px;justify-content:space-between;font-variant-numeric:tabular-nums}
  .tip .r span:first-child{color:var(--ink-2)}
  @media (max-width:560px){ header{align-items:stretch} .wrap{padding-block:20px 40px} }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <h1>Control de facturas</h1>
      <p class="sub">Estado &laquo;Cerrada&raquo; o &laquo;Cerrada para recepci&oacute;n&raquo; con fecha de cierre dentro del mes &middot; generado <span id="sello">__STAMP__</span></p>
      __AVISO__
    </div>
    <label class="sel">Mes <select id="mes"></select></label>
  </header>

  <div class="tiles" id="tiles"></div>

  <h2>Resumen por cuenta</h2>
  <div class="card scroll"><table id="resumen"></table></div>

  <h2>Evoluci&oacute;n de ingreso de facturas</h2>
  <div class="card">
    <div class="legend" id="leyenda"></div>
    <div id="chart"></div>
    <p class="note">Monto facturado acumulado d&iacute;a a d&iacute;a dentro del mes. Sin ajuste de provisi&oacute;n: cada l&iacute;nea termina en el TOTAL FACTURAS MES de su cuenta.</p>
  </div>

  <h2>Detalle por proveedor</h2>
  <div class="cols" id="detalle"></div>

  <div id="control"></div>
</div>
<div class="tip" id="tip"></div>

__DATOS_INLINE__
<script>
const fmt  = n => n==null ? "–" : "$" + Math.round(n).toLocaleString("es-CL");
const fmtK = n => n==null ? "–" : "$" + Math.round(n/1e6).toLocaleString("es-CL") + "M";
const fmtP = n => n==null ? "–" : (n*100).toFixed(1).replace(".",",") + "%";
const COL  = {PRIMARIA:"var(--s1)", SECUNDARIA:"var(--s2)", OTROS:"var(--muted)"};

const selMes = document.getElementById("mes");
Object.keys(DATOS).sort().reverse().forEach(m=>{
  const o=document.createElement("option"); o.value=m; o.textContent=DATOS[m].etiqueta; selMes.appendChild(o);
});
selMes.addEventListener("change", ()=>render(selMes.value));

function render(m){
  const d = DATOS[m];

  // ---- tiles ----
  document.getElementById("tiles").innerHTML = d.filas.map(f=>{
    const pct = f.avance;
    const col = pct==null ? "var(--muted)" : (pct>1 ? "var(--crit)" : pct>0.9 ? "var(--warn)" : COL[f.cuenta]);
    const ancho = pct==null ? 0 : Math.min(pct,1)*100;
    return `<div class="card tile">
      <div class="lbl">${f.cuenta} · cierre de mes</div>
      <div class="val">${fmt(f.cierre)}</div>
      <div class="meta">${fmt(f.total)} facturado en ${f.n} ${f.n===1?"orden":"órdenes"}${
        f.presupuesto==null ? " · sin presupuesto" : ` · ${fmtP(pct)} del presupuesto`}</div>
      <div class="meter"><i style="width:${ancho}%;background:${col}"></i></div>
    </div>`;
  }).join("");

  // ---- resumen ----
  const cols = [["Cierre de mes","cierre"],["Presupuesto","presupuesto"],
                ["Revierte prov.","reversa"],["Provisión mes","provision"],
                ["Total facturas","total"],["N°","n"],
                ["% avance","avance"],["Saldo ppto.","saldo"]];
  const suma = k => d.filas.reduce((a,f)=>a+(f[k]||0),0);
  document.getElementById("resumen").innerHTML =
    `<thead><tr><th>Cuenta</th>${cols.map(c=>`<th class="num">${c[0]}</th>`).join("")}</tr></thead>
     <tbody>${d.filas.map(f=>`<tr><td><strong>${f.cuenta}</strong></td>${
        cols.map(c=>`<td class="num">${
          c[1]==="n" ? f.n : c[1]==="avance" ? fmtP(f.avance) : fmt(f[c[1]])}</td>`).join("")
      }</tr>`).join("")}</tbody>
     <tfoot><tr><td>TOTAL</td>${cols.map(c=>`<td class="num">${
        c[1]==="n" ? suma("n")
        : c[1]==="avance" ? (suma("presupuesto") ? fmtP(suma("cierre")/suma("presupuesto")) : "–")
        : fmt(suma(c[1]))}</td>`).join("")}</tr></tfoot>`;

  // ---- detalle ----
  document.getElementById("detalle").innerHTML = d.filas.map(f=>
    `<div class="card"><h2 style="margin:0 0 10px">${f.cuenta}</h2><div class="scroll"><table>
      <thead><tr><th>Proveedor</th><th class="num">N°</th><th class="num">Monto</th></tr></thead>
      <tbody>${f.detalle.length ? f.detalle.map(p=>
        `<tr><td>${p.prov}</td><td class="num">${p.n}</td><td class="num">${fmt(p.monto)}</td></tr>`).join("")
        : `<tr><td colspan="3" style="color:var(--muted)">Sin facturas este mes</td></tr>`}</tbody>
      <tfoot><tr><td>Total general</td><td class="num">${f.n}</td><td class="num">${fmt(f.total)}</td></tr></tfoot>
    </table></div></div>`).join("");

  // ---- control de integridad ----
  const sc = d.sin_canal;
  document.getElementById("control").innerHTML = sc.n
    ? `<div class="flag bad">▲ ${sc.n} facturas del mes (${fmt(sc.monto)}) no tienen CANAL asignado en Exported y quedan fuera de las tres cuentas.</div>`
    : `<div class="flag ok">✓ Todas las facturas del mes tienen CANAL asignado.</div>`;

  dibujar(d);
}

// ---------------- gráfico ----------------
let ESTADO = null;
function dibujar(d){
  const W=1100, H=380, ML=92, MR=22, MT=16, MB=40;
  const iw=W-ML-MR, ih=H-MT-MB, N=d.ndias;
  const sx = i => ML + (N<2?iw/2:iw*i/(N-1));
  const vals=[];
  ["PRIMARIA","SECUNDARIA"].forEach(g=>{ (d.series[g]||[]).forEach(v=>vals.push(v)); });
  const max = Math.max(1, ...vals), min = Math.min(0, ...vals);
  const top = max*1.08, bot = min*1.08;
  const sy = v => MT + ih - ih*(v-bot)/(top-bot);

  // eje Y: 5 marcas
  let ticks="";
  for(let k=0;k<=4;k++){
    const v=bot+(top-bot)*k/4, y=sy(v);
    ticks += `<line x1="${ML}" x2="${W-MR}" y1="${y}" y2="${y}" stroke="var(--grid)" stroke-width="1"/>
      <text x="${ML-10}" y="${y+4}" text-anchor="end" font-size="11.5" fill="var(--muted)"
        style="font-variant-numeric:tabular-nums">${fmtK(v)}</text>`;
  }
  // eje X
  let xlab="";
  for(let i=0;i<N;i++){
    const d1=i+1;
    if(d1===1||d1%5===0||d1===N)
      xlab += `<text x="${sx(i)}" y="${H-MB+20}" text-anchor="middle" font-size="11.5" fill="var(--muted)">${d1}</text>`;
  }

  const hoy=new Date(), mesActual = d.etiqueta === DATOS[selMes.value].etiqueta;
  const claveMes = selMes.value;
  const esMesEnCurso = claveMes === `${hoy.getFullYear()}-${String(hoy.getMonth()+1).padStart(2,"0")}`;
  const ultimo = esMesEnCurso ? Math.min(hoy.getDate(), N) : N;

  let paths="", puntos="";
  ["PRIMARIA","SECUNDARIA"].forEach(g=>{
    const col = g==="PRIMARIA" ? "var(--s1)" : "var(--s2)";
    const serie=(d.series[g]||[]).slice(0,ultimo);
    if(serie.length){
      paths += `<path d="${serie.map((v,i)=>`${i?"L":"M"}${sx(i)},${sy(v)}`).join(" ")}"
        fill="none" stroke="${col}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
      const li=serie.length-1;
      puntos += `<circle cx="${sx(li)}" cy="${sy(serie[li])}" r="4.5" fill="${col}"
        stroke="var(--surface)" stroke-width="2"/>`;
    }
  });

  document.getElementById("leyenda").innerHTML =
    `<span><b style="border-top-color:var(--s1)"></b>PRIMARIA</span>
     <span><b style="border-top-color:var(--s2)"></b>SECUNDARIA</span>`;

  document.getElementById("chart").innerHTML =
    `<svg id="svg" viewBox="0 0 ${W} ${H}" role="img"
       aria-label="Monto facturado acumulado dia a dia dentro del mes, por canal">
      ${ticks}
      <line x1="${ML}" x2="${W-MR}" y1="${sy(0)}" y2="${sy(0)}" stroke="var(--axis)" stroke-width="1"/>
      ${xlab}${paths}${puntos}
      <line id="cross" x1="0" x2="0" y1="${MT}" y2="${MT+ih}" stroke="var(--axis)" stroke-width="1" opacity="0"/>
      <rect id="hit" x="${ML}" y="${MT}" width="${iw}" height="${ih}" fill="transparent"/>
    </svg>`;

  ESTADO={d,N,ML,iw,ultimo,sx};
  const svg=document.getElementById("svg"), tip=document.getElementById("tip"),
        cross=document.getElementById("cross");
  svg.addEventListener("pointermove", e=>{
    const r=svg.getBoundingClientRect(), x=(e.clientX-r.left)/r.width*W;
    let i=Math.round((x-ML)/(iw/Math.max(1,N-1)));
    i=Math.max(0,Math.min(N-1,i));
    cross.setAttribute("x1",sx(i)); cross.setAttribute("x2",sx(i)); cross.setAttribute("opacity",".8");
    const fila=["PRIMARIA","SECUNDARIA"].map(g=>{
      const s=d.series[g]||[];
      const v = i<ultimo ? s[i] : null;
      return `<div class="r"><span>${g}</span><span style="color:${g==="PRIMARIA"?"var(--s1)":"var(--s2)"}">${fmt(v)}</span></div>`;
    }).join("");
    tip.innerHTML=`<div class="r"><span>Día</span><span><strong>${i+1}</strong></span></div>${fila}`;
    tip.style.opacity=1;
    tip.style.left=Math.min(e.clientX+14, innerWidth-190)+"px";
    tip.style.top=(e.clientY+14)+"px";
  });
  svg.addEventListener("pointerleave", ()=>{ tip.style.opacity=0; cross.setAttribute("opacity",0); });
}

const HOY = new Date();
const CLAVE_HOY = `${HOY.getFullYear()}-${String(HOY.getMonth()+1).padStart(2,"0")}`;
const CON_DATOS = Object.keys(DATOS).sort().filter(k=>DATOS[k].filas.some(f=>f.n>0));
selMes.value = DATOS[CLAVE_HOY] ? CLAVE_HOY : (CON_DATOS.pop() || Object.keys(DATOS).sort().pop());
render(selMes.value);
</script>
</body>
</html>
"""


def main():
    if not os.path.exists(XLSX):
        sys.exit(f'No encuentro el archivo: {XLSX}')
    ex, params = leer(XLSX)
    datos = construir(ex, params)
    stamp = datetime.datetime.now().strftime('%d-%m-%Y %H:%M')
    destino = WEB or '.'
    os.makedirs(destino, exist_ok=True)

    with open(os.path.join(destino, 'datos.js'), 'w', encoding='utf-8') as f:
        f.write('// Generado automaticamente desde el Excel. No editar a mano.\n')
        f.write('const DATOS = ' + json.dumps(datos, ensure_ascii=False) + ';\n')

    idx = (TPL
           .replace('__DATOS_INLINE__', '<script src="datos.js"></script>')
           .replace('__STAMP__', stamp)
           .replace('__AVISO__', ''))
    with open(os.path.join(destino, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(idx)

    print(f'Listo: {destino}/index.html + datos.js  ({len(datos)} meses)')


if __name__ == '__main__':
    main()
