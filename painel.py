# -*- coding: utf-8 -*-
"""
Painel web LOCAL para comandar o lancar_atendimentos.py.
Rode:  python painel.py     (abre no navegador em http://localhost:8760)

Nao precisa instalar nada (usa so a biblioteca padrao do Python).
Voce preenche os campos do atendimento (unidade, projeto/acao, tema, subtema,
textos) + os parametros da rodada (planilha, porta CDP, datas/horarios) e clica
em TESTAR (preenche 1 sem finalizar) ou RODAR (lanca de verdade).
O login no ASA continua sendo feito por voce, na janela do Chrome.
"""
import os, sys, json, subprocess, threading, webbrowser
from pathlib import Path
from datetime import date, timedelta
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

PASTA   = Path(__file__).resolve().parent
SCRIPT  = PASTA / "lancar_atendimentos.py"
PERFIL_JSON = PASTA / "perfil_custom.json"
RUN_JSON    = PASTA / "painel_run.json"
TEXTOS_JSON = PASTA / "textos_variados.json"
LOG     = PASTA / "run_painel.log"
PORTA   = 8760

# valores padrao (perfil UAR) usados quando ainda nao ha perfil_custom.json
PADRAO = {
    "atendente": ["YAN", "YAN PEREIRA FERREIRA"],
    "unidade":   ["UAR", "Unidade de Atendimento e Relacionamento"],
    "projeto":   ["DESCENTRALIZADO", "DESCENTRALIZADO E PARCEIROS"],
    "canal":     ["", "Agência fixa"],
    "tipo":      ["", "Orientação"],
    "solucao":   ["", ""],
    "tema":      ["Finan", "Finanças"],
    "subtema":   ["Financeira", "Gestão Financeira"],
    "orientacao": "Visite o Sebrae Aleixo para mais informações sobre cursos, portfólio Sebrae e manter seu CNPJ regularizado",
    "descricao":  "Cliente recebeu orientação sobre leis e normas do MEI e foi encaminhada para consultoria",
}
PADRAO_RUN = {
    "planilha": "", "cdp": "9224", "de": "", "ate": "", "datas": "", "slots": "",
    "duracao": "45", "folga": "1", "fim": "18:59", "limite": "0",
    "refazer": False, "embaralhar": True,
}

_proc = None   # subprocesso da rodada em andamento

# ------------------------------------------------------------------- HTML ----
HTML = r"""<!doctype html><html lang=pt-br><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>Painel ASA</title><style>
*{box-sizing:border-box} body{font-family:system-ui,Segoe UI,Arial;margin:0;background:#0f172a;color:#e2e8f0}
header{background:#1e293b;padding:14px 20px;font-size:20px;font-weight:700;border-bottom:2px solid #334155}
.wrap{max-width:1050px;margin:0 auto;padding:18px}
.card{background:#1e293b;border:1px solid #334155;border-radius:12px;padding:16px 18px;margin-bottom:16px}
h2{margin:0 0 12px;font-size:15px;color:#93c5fd;text-transform:uppercase;letter-spacing:.5px}
.grid{display:grid;grid-template-columns:180px 1fr 1fr;gap:8px 12px;align-items:center}
.grid label{font-size:13px;color:#cbd5e1}
.grid .head{font-size:11px;color:#64748b;text-transform:uppercase}
input,textarea{width:100%;background:#0f172a;border:1px solid #475569;color:#e2e8f0;border-radius:7px;padding:8px 10px;font-size:13px;font-family:inherit}
input:focus,textarea:focus{outline:2px solid #3b82f6;border-color:#3b82f6}
textarea{min-height:60px;resize:vertical}
.run{display:grid;grid-template-columns:repeat(3,1fr);gap:10px 14px}
.run .full{grid-column:1/-1}
.fld{display:flex;flex-direction:column;gap:4px}
.fld label{font-size:12px;color:#94a3b8}
.chk{display:flex;align-items:center;gap:8px;font-size:13px}
.chk input{width:auto}
.btns{display:flex;gap:10px;flex-wrap:wrap;margin-top:6px}
button{border:0;border-radius:8px;padding:11px 20px;font-size:14px;font-weight:600;cursor:pointer}
.b-save{background:#475569;color:#fff} .b-test{background:#d97706;color:#fff}
.b-run{background:#16a34a;color:#fff} .b-stop{background:#dc2626;color:#fff}
button:disabled{opacity:.5;cursor:not-allowed}
#status{font-size:13px;margin-left:auto;align-self:center}
pre{background:#020617;border:1px solid #334155;border-radius:10px;padding:12px;height:320px;overflow:auto;font-size:12px;line-height:1.45;white-space:pre-wrap;margin:0}
.hint{font-size:12px;color:#64748b;margin:4px 0 0}
small{color:#64748b}
.tx{border:1px solid #334155;border-radius:9px;padding:10px 12px;margin-bottom:8px;background:#0f172a}
.tx .srv{width:100%;font-weight:600;margin-bottom:6px}
.tx textarea{margin-top:6px}
.txrm{background:#7f1d1d;color:#fff;border:0;border-radius:6px;padding:4px 10px;font-size:12px;cursor:pointer;float:right}
</style></head><body>
<header>🗂️ Painel ASA — comandar lançamentos</header>
<div class=wrap>

<div class=card>
  <h2>Classificação do atendimento</h2>
  <div class=grid>
    <span></span><span class=head>Buscar (filtro do menu)</span><span class=head>Opção (texto exato)</span>
    <label>Atendente</label><input id=atendente_b><input id=atendente_o>
    <label>Unidade organizacional</label><input id=unidade_b><input id=unidade_o>
    <label>Projeto / Ação</label><input id=projeto_b><input id=projeto_o>
    <label>Canal</label><input id=canal_b><input id=canal_o>
    <label>Tipo de atendimento</label><input id=tipo_b><input id=tipo_o>
    <label>Solução <small>(vazio se não tiver)</small></label><input id=solucao_b><input id=solucao_o>
    <label>Tema</label><input id=tema_b><input id=tema_o>
    <label>Subtema</label><input id=subtema_b><input id=subtema_o>
  </div>
  <p class=hint>“Buscar” é o texto digitado para filtrar o menu (deixe vazio em menus que não têm busca). “Opção” é o item que será selecionado.</p>
</div>

<div class=card>
  <h2>Texto fixo <small>(fallback — usado só quando não há textos variados)</small></h2>
  <div class=fld><label>Orientação ao cliente</label><textarea id=orientacao></textarea></div>
  <div class=fld style=margin-top:10px><label>Relato / Descrição</label><textarea id=descricao></textarea></div>
</div>

<div class=card>
  <h2>Parâmetros da rodada</h2>
  <div class=run>
    <div class="fld full"><label>Planilha (.xlsx) — caminho completo</label><input id=planilha placeholder="C:\Users\...\clientes.xlsx"></div>
    <div class=fld><label>Porta CDP (janela logada)</label><input id=cdp></div>
    <div class=fld><label>Duração (min)</label><input id=duracao oninput=calcCap()></div>
    <div class=fld><label>Folga (min)</label><input id=folga oninput=calcCap()></div>
    <div class=fld><label>Último término (HH:MM)</label><input id=fim oninput=calcCap()></div>
    <div class=fld><label>Limite (0 = todos)</label><input id=limite></div>
    <div class=fld><label class=chk><input type=checkbox id=refazer> Refazer já lançados</label></div>
    <div class=fld><label class=chk><input type=checkbox id=embaralhar> Embaralhar textos a cada rodada</label></div>
    <div class=fld><label>De (1ª data) 📅</label><input type=date id=de oninput=calcCap()></div>
    <div class=fld><label>Até (última data) 📅</label><input type=date id=ate oninput=calcCap()></div>
    <div class="fld full"><span id=capacidade class=hint>Preencha De/Até para ver a capacidade.</span></div>
    <div class="fld full"><label>Datas manuais (ASA_DATAS) — opcional; use só se NÃO preencher De/Até acima. Ex: 27/07/2026,28/07/2026 ou 24/07/2026@11:10</label><input id=datas></div>
    <div class="fld full"><label>Horários exatos (ASA_SLOTS) — ex: 22/07/2026 09:01,22/07/2026 09:47 (um por atendimento; se preenchido, ignora Datas)</label><input id=slots></div>
  </div>
</div>

<div class=card>
  <h2>Textos variados dos atendimentos <small>(rotacionam — 1 serviço por cliente)</small></h2>
  <div id=textos></div>
  <div class=btns style=margin-top:10px>
    <button class=b-save onclick=addTexto()>➕ Adicionar serviço</button>
    <button class=b-test onclick=randomizar()>🔀 Randomizar ordem</button>
    <button class=b-run onclick=salvarTextos()>💾 Salvar textos</button>
    <span id=statusTx></span>
  </div>
  <p class=hint>Estes textos substituem o texto fixo. Deixe a lista vazia (e salve) para voltar ao texto fixo acima. “Randomizar” embaralha a ordem aqui; a rodada também pode embaralhar sozinha (opção nos parâmetros).</p>
</div>

<div class=card>
  <div class=btns>
    <button class=b-save onclick=salvar()>💾 Salvar perfil</button>
    <button class=b-test onclick="rodar(true)">🧪 Testar (1, não finaliza)</button>
    <button class=b-run onclick="rodar(false)">▶️ Rodar (finaliza)</button>
    <button class=b-stop onclick=parar()>⏹️ Parar</button>
    <span id=status>—</span>
  </div>
</div>

<div class=card><h2>Log da execução</h2><pre id=log>Aguardando…</pre></div>
</div>

<script>
const $=id=>document.getElementById(id);
const COMBOS=["atendente","unidade","projeto","canal","tipo","solucao","tema","subtema"];
const RUN=["planilha","cdp","de","ate","datas","slots","duracao","folga","fim","limite"];
function slotsPorDia(){
  const dur=+($("duracao").value||45), fol=+($("folga").value||0);
  const fp=($("fim").value||"18:59").split(":"); const lim=(+fp[0])*60+(+fp[1]||0);
  let t=7*60, n=0;
  while(t+dur<=lim){ n++; t+=dur+fol; }
  return n;
}
function diasUteis(a,b){
  if(!a||!b) return 0;
  let d=new Date(a+"T00:00"), e=new Date(b+"T00:00"), n=0;
  if(e<d) return 0;
  while(d<=e){ const w=d.getDay(); if(w!==0&&w!==6) n++; d.setDate(d.getDate()+1); }
  return n;
}
function calcCap(){
  const du=diasUteis($("de").value,$("ate").value), sd=slotsPorDia();
  $("capacidade").textContent = du
    ? `Capacidade: ${du} dias úteis × ${sd}/dia = ${du*sd} atendimentos (fins de semana são pulados).`
    : "Preencha De/Até para ver a capacidade.";
}
function coletar(){
  const perfil={};
  COMBOS.forEach(k=>perfil[k]=[$(k+"_b").value,$(k+"_o").value]);
  perfil.orientacao=$("orientacao").value; perfil.descricao=$("descricao").value;
  const run={}; RUN.forEach(k=>run[k]=$(k).value);
  run.refazer=$("refazer").checked; run.embaralhar=$("embaralhar").checked;
  return {perfil,run};
}
function aplicar(d){
  const p=d.perfil||{}, r=d.run||{};
  COMBOS.forEach(k=>{const v=p[k]||["",""]; $(k+"_b").value=v[0]||""; $(k+"_o").value=v[1]||"";});
  $("orientacao").value=p.orientacao||""; $("descricao").value=p.descricao||"";
  RUN.forEach(k=>$(k).value=r[k]!==undefined?r[k]:"");
  $("refazer").checked=!!r.refazer; $("embaralhar").checked=r.embaralhar!==false;
}
let TEXTOS=[];
function renderTextos(){
  const box=$("textos"); box.innerHTML="";
  TEXTOS.forEach((t,i)=>{
    const d=document.createElement("div"); d.className="tx";
    d.innerHTML=`<button class=txrm data-i=${i}>✖ remover</button>`+
      `<input class=srv data-k=servico data-i=${i} placeholder="Serviço (ex.: Emissão de DAS)">`+
      `<textarea data-k=orientacao data-i=${i} placeholder="Orientação ao cliente"></textarea>`+
      `<textarea data-k=descricao data-i=${i} placeholder="Relato / Descrição"></textarea>`;
    box.appendChild(d);
    d.querySelector('[data-k=servico]').value=t.servico||"";
    d.querySelector('[data-k=orientacao]').value=t.orientacao||"";
    d.querySelector('[data-k=descricao]').value=t.descricao||"";
  });
  box.querySelectorAll("input,textarea").forEach(el=>el.addEventListener("input",e=>{
    TEXTOS[+e.target.dataset.i][e.target.dataset.k]=e.target.value;
  }));
  box.querySelectorAll(".txrm").forEach(b=>b.addEventListener("click",e=>{
    TEXTOS.splice(+e.target.dataset.i,1); renderTextos();
  }));
}
function addTexto(){ TEXTOS.push({servico:"",orientacao:"",descricao:""}); renderTextos(); }
function randomizar(){ for(let i=TEXTOS.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[TEXTOS[i],TEXTOS[j]]=[TEXTOS[j],TEXTOS[i]];} renderTextos(); $("statusTx").textContent="🔀 ordem embaralhada (salve para gravar)"; }
async function salvarTextos(){
  const lista=TEXTOS.filter(t=>(t.orientacao||"").trim()&&(t.descricao||"").trim());
  const r=await fetch("/textos",{method:"POST",body:JSON.stringify(lista)});
  $("statusTx").textContent=(await r.json()).msg;
}
async function salvar(){
  const r=await fetch("/salvar",{method:"POST",body:JSON.stringify(coletar())});
  $("status").textContent=(await r.json()).msg;
}
async function rodar(teste){
  await salvar();
  const r=await fetch("/rodar?teste="+(teste?1:0),{method:"POST"});
  $("status").textContent=(await r.json()).msg;
}
async function parar(){ $("status").textContent=(await (await fetch("/parar",{method:"POST"})).json()).msg; }
async function puxarLog(){
  try{const r=await fetch("/log"); const d=await r.json();
    const pre=$("log"); const perto=pre.scrollTop+pre.clientHeight>=pre.scrollHeight-30;
    pre.textContent=d.log||"(vazio)"; if(perto)pre.scrollTop=pre.scrollHeight;
    $("status").textContent=d.rodando?"⏳ rodando…":"● parado";
  }catch(e){}
}
(async()=>{
  aplicar(await (await fetch("/config")).json());
  calcCap();
  try{ TEXTOS=await (await fetch("/textos")).json(); }catch(e){ TEXTOS=[]; }
  renderTextos();
  setInterval(puxarLog,1500); puxarLog();
})();
</script></body></html>"""

# ---------------------------------------------------------------- SERVER -----
def carregar_config():
    perfil = json.loads(PERFIL_JSON.read_text(encoding="utf-8")) if PERFIL_JSON.exists() else dict(PADRAO)
    run    = json.loads(RUN_JSON.read_text(encoding="utf-8")) if RUN_JSON.exists() else dict(PADRAO_RUN)
    return {"perfil": perfil, "run": run}

def salvar_config(dados):
    PERFIL_JSON.write_text(json.dumps(dados.get("perfil", {}), ensure_ascii=False, indent=2), encoding="utf-8")
    RUN_JSON.write_text(json.dumps(dados.get("run", {}), ensure_ascii=False, indent=2), encoding="utf-8")

def carregar_textos():
    if TEXTOS_JSON.exists():
        try: return json.loads(TEXTOS_JSON.read_text(encoding="utf-8"))
        except Exception: return []
    return []

def salvar_textos_pool(lista):
    limpa = [{"servico": t.get("servico", ""), "orientacao": t.get("orientacao", ""),
              "descricao": t.get("descricao", "")}
             for t in lista if (t.get("orientacao") or "").strip() and (t.get("descricao") or "").strip()]
    TEXTOS_JSON.write_text(json.dumps(limpa, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(limpa)

def dias_uteis_br(de_iso, ate_iso):
    """Expande um intervalo (YYYY-MM-DD) em datas DD/MM/AAAA de dias úteis (seg-sex)."""
    try:
        d = date.fromisoformat(de_iso); e = date.fromisoformat(ate_iso)
    except ValueError:
        return []
    if e < d:
        return []
    out = []
    while d <= e:
        if d.weekday() < 5:            # 0=seg ... 4=sex; pula sáb/dom
            out.append(d.strftime("%d/%m/%Y"))
        d += timedelta(days=1)
    return out


def iniciar_run(teste):
    global _proc
    if _proc and _proc.poll() is None:
        return "Já há uma rodada em andamento."
    run = json.loads(RUN_JSON.read_text(encoding="utf-8")) if RUN_JSON.exists() else dict(PADRAO_RUN)
    if not run.get("planilha") or not Path(run["planilha"]).exists():
        return "Planilha não encontrada — confira o caminho."
    # Intervalo De/Até tem prioridade: gera as datas de dias úteis automaticamente.
    de, ate = run.get("de", "").strip(), run.get("ate", "").strip()
    datas_intervalo = dias_uteis_br(de, ate) if (de and ate) else []
    env = dict(os.environ)
    env["ASA_PERFIL"]  = "CUSTOM"
    env["ASA_PLANILHA"] = run["planilha"]
    env["ASA_CDP"]      = run.get("cdp", "")
    env["ASA_DURACAO"]  = run.get("duracao", "45")
    env["ASA_FOLGA"]    = run.get("folga", "1")
    env["ASA_FIM"]      = run.get("fim", "18:59")
    env["ASA_LIMITE"]   = run.get("limite", "0")
    env["ASA_DATAS"]    = ",".join(datas_intervalo) if datas_intervalo else run.get("datas", "")
    env["ASA_SLOTS"]    = run.get("slots", "")
    env["ASA_REFAZER"]  = "1" if run.get("refazer") else "0"
    env["ASA_EMBARALHAR"] = "1" if run.get("embaralhar", True) else "0"
    env["ASA_MODO_TESTE"] = "1" if teste else "0"
    env["PYTHONIOENCODING"] = "utf-8"
    LOG.write_text("", encoding="utf-8")
    f = open(LOG, "w", encoding="utf-8")
    _proc = subprocess.Popen([sys.executable, "-u", str(SCRIPT)], env=env,
                             stdout=f, stderr=subprocess.STDOUT, cwd=str(PASTA))
    msg = ("TESTE" if teste else "RODADA") + " iniciada. Acompanhe o log abaixo."
    if datas_intervalo:
        msg += f" [{len(datas_intervalo)} dias úteis de {datas_intervalo[0]} a {datas_intervalo[-1]}]"
    return msg

def parar_run():
    global _proc
    if _proc and _proc.poll() is None:
        _proc.terminate()
        return "Rodada interrompida."
    return "Nada rodando."

class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def _send(self, code, body, ctype="application/json"):
        b = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype + ("; charset=utf-8" if "json" in ctype or "html" in ctype else ""))
        self.send_header("Content-Length", str(len(b)))
        self.end_headers(); self.wfile.write(b)
    def _json(self, obj): self._send(200, json.dumps(obj, ensure_ascii=False))
    def do_GET(self):
        p = self.path.split("?")[0]
        if p == "/": self._send(200, HTML, "text/html")
        elif p == "/config": self._json(carregar_config())
        elif p == "/textos": self._json(carregar_textos())
        elif p == "/log":
            txt = LOG.read_text(encoding="utf-8", errors="replace") if LOG.exists() else ""
            rodando = bool(_proc and _proc.poll() is None)
            self._json({"log": txt[-16000:], "rodando": rodando})
        else: self._send(404, "{}")
    def do_POST(self):
        p = self.path.split("?")[0]
        n = int(self.headers.get("Content-Length", 0) or 0)
        raw = self.rfile.read(n).decode("utf-8") if n else "{}"
        if p == "/salvar":
            try: salvar_config(json.loads(raw)); self._json({"msg": "✔ Perfil salvo."})
            except Exception as e: self._json({"msg": "Erro ao salvar: %s" % e})
        elif p == "/rodar":
            teste = self.path.endswith("teste=1")
            self._json({"msg": iniciar_run(teste)})
        elif p == "/parar":
            self._json({"msg": parar_run()})
        elif p == "/textos":
            try:
                n = salvar_textos_pool(json.loads(raw))
                self._json({"msg": "✔ %d textos salvos." % n})
            except Exception as e:
                self._json({"msg": "Erro ao salvar textos: %s" % e})
        else: self._send(404, "{}")

if __name__ == "__main__":
    url = f"http://localhost:{PORTA}"
    print("Painel ASA em", url, "\n(Ctrl+C para encerrar)")
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    ThreadingHTTPServer(("127.0.0.1", PORTA), H).serve_forever()
