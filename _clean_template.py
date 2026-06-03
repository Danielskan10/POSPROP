"""Elimina código zombie y funciones legacy del template."""
import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('_dashboard_tpl.html','r',encoding='utf-8') as f:
    content = f.read()

# ── 1. Eliminar bloque zombie (cuerpo de initGeneral sin header)
# Desde "// ═══ TAB ⑥  DETALLE" hasta el "}" que lo cierra (L2063-2203)
ZOMBIE_START = '''// ════════════════════════════════════════════════════════════
// TAB ⑥  DETALLE
// ════════════════════════════════════════════════════════════
  const rows=FD(), ev=FE();'''

ZOMBIE_END = '''  document.getElementById('portCards').innerHTML = buildPortCards(rows, ttl);
}'''

# ── 2. Eliminar _legacyCausaciones
LEGACY_CAUS_START = '''// (initCausaciones e initRentaFija reemplazados por initRendimiento e initInstrumentos)
// ════════════════════════════════════════════════════════════
// TAB 4: CAUSACIONES (versión legada — solo para compatibilidad de alias)
// ════════════════════════════════════════════════════════════
function _legacyCausaciones(){'''

LEGACY_CAUS_END = '''  document.getElementById('tCausBody').innerHTML=
    [...rows].sort((a,b)=>Math.abs(b.caus_mer||0)-Math.abs(a.caus_mer||0)).slice(0,60).map(r=>
      `<tr><td class="main" title="${r.esp}">${r.esp.length>28?r.esp.slice(0,28)+'…':r.esp}</td>
      <td>${r.port||'—'}</td><td>${bdg(r.tipo)}</td>
      <td class="${clr(r.caus_mer)}">${fC(r.caus_mer)}</td>
      <td class="${clr(r.caus_tir)}">${fC(r.caus_tir)}</td>
      <td class="${clr((r.caus_mer||0)-(r.caus_tir||0))}">${fC((r.caus_mer||0)-(r.caus_tir||0))}</td>
      <td class="${clr(r.caus_mon)}">${fC(r.caus_mon)}</td>
      <td class="${clr(r.caus_tasa)}">${fC(r.caus_tasa)}</td>
      <td class="r">${fC(r.adeudados)}</td>
      <td class="r">${fC(r.valor)}</td></tr>`
    ).join('');
}'''

# ── 3. Eliminar _legacyRentaFija
LEGACY_RF_START = '''// ════════════════════════════════════════════════════════════
// TAB 5: RENTA FIJA (versión legada)
// ════════════════════════════════════════════════════════════
function _legacyRentaFija(){'''

LEGACY_RF_END = '''    <td class="${clr(r.caus_tir)}">${fC(r.caus_tir)}</td>
    <td class="r">${fC(r.adeudados)}</td></tr>`
  ).join('');
}'''

# ── 4. Eliminar initAnalisis vacío
INIT_ANALISIS = '''// initAnalisis ya no existe — las funcionalidades se distribuyeron en los nuevos tabs
function initAnalisis(){}'''

def remove_block(text, start, end):
    idx_s = text.find(start)
    if idx_s == -1:
        print(f'  WARN: bloque inicio no encontrado: {start[:50]!r}')
        return text
    idx_e = text.find(end, idx_s)
    if idx_e == -1:
        print(f'  WARN: bloque fin no encontrado: {end[:50]!r}')
        return text
    idx_e += len(end)
    removed = text[idx_s:idx_e]
    print(f'  OK: eliminado {len(removed)} chars ({removed[:60]!r}...)')
    return text[:idx_s] + text[idx_e:]

print('Eliminando bloques de código zombie y legacy...')
content = remove_block(content, ZOMBIE_START, ZOMBIE_END)
content = remove_block(content, LEGACY_CAUS_START, LEGACY_CAUS_END)
content = remove_block(content, LEGACY_RF_START, LEGACY_RF_END)
content = content.replace(INIT_ANALISIS, '// initAnalisis: eliminado — ver initRendimiento e initInstrumentos')
print(f'  initAnalisis: eliminado')

with open('_dashboard_tpl.html','w',encoding='utf-8') as f:
    f.write(content)

print(f'Template guardado: {len(content)//1024} KB')
