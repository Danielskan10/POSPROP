"""
Genera index.html + app.js separando el JS del HTML.
El JS externo no tiene problemas con el parser HTML ni con SES.
"""
import json, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')

ROOT = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(ROOT,'_dashboard_tpl.html'),'r',encoding='utf-8') as f:
    tpl = f.read()

with open(os.path.join(ROOT,'data.json'),'r',encoding='utf-8') as f:
    data = json.load(f)

data_str = json.dumps(data, ensure_ascii=False, separators=(',',':'))

# Sustituir marcadores en el template
tpl2 = (tpl
    .replace("__DATA_JSON__",   data_str)
    .replace("__DATA_URL__",    "./data.json")
    .replace("__ORG__",         "Skandia Colombia")
    .replace("__MODO_REMOTO__", "true"))

# Extraer los bloques <script> inline y separarlos al archivo app.js
# dejando en index.html solo <script src="app.js"></script>
inline_scripts = []
def collect(m):
    inline_scripts.append(m.group(1))
    return ''   # eliminar del HTML

html_clean = re.sub(r'<script(?! src)>(.*?)</script>',
                    collect, tpl2, flags=re.DOTALL)

# Juntar todos los scripts inline en app.js
app_js = '\n'.join(inline_scripts)

# Insertar <script src="app.js"></script> justo antes de </body>
html_final = html_clean.replace('</body>', '<script src="app.js"></script>\n</body>')

# Escribir
html_out = os.path.join(ROOT,'index.html')
js_out   = os.path.join(ROOT,'app.js')

with open(html_out,'w',encoding='utf-8') as f: f.write(html_final)
with open(js_out,  'w',encoding='utf-8') as f: f.write(app_js)

kb_html = os.path.getsize(html_out)//1024
kb_js   = os.path.getsize(js_out)//1024
print(f"index.html : {kb_html} KB")
print(f"app.js     : {kb_js} KB")

# Validar sintaxis del app.js con node
import subprocess
r = subprocess.run(['node','--check',js_out], capture_output=True, text=True)
if r.returncode == 0:
    print("app.js     : sintaxis OK")
else:
    err = r.stderr.strip()
    print(f"app.js     : ERROR SINTAXIS\n{err}")
    # Mostrar contexto
    m = re.search(r':(\d+)', err)
    if m:
        ln = int(m.group(1))
        lines = app_js.split('\n')
        for i in range(max(0,ln-5), min(len(lines),ln+3)):
            marker='>>>' if i+1==ln else '   '
            print(f"{marker} L{i+1}: {lines[i][:90]}")
