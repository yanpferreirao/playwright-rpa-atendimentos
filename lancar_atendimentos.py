# -*- coding: utf-8 -*-
"""
Lancamento automatico de atendimentos no ASA Externo (Sebrae-AM).

Fluxo do wizard (5 etapas + finalizacao):
  CPF -> Buscar
  Etapa 1: Dados da Pessoa Fisica  -> Salvar e Prosseguir (auto-preenchido)
  Etapa 2: Empreendimento          -> vincular CNPJ + Salvar e Prosseguir
  Etapa 3: Dados da Empresa        -> Salvar e Prosseguir (auto)
  Etapa 4: Contatos & Enderecos    -> Salvar e Prosseguir (auto)
  Etapa 5: Atendimento             -> preenche campos fixos + data/hora
                                      -> Validar Horarios -> FINALIZAR

Seguranca:
  - MODO_TESTE=True: processa so o 1o cliente e PARA antes de finalizar
    (nao cria registro). Tira print de cada passo em 'debug/'.
  - Guarda progresso em 'progresso_<PERFIL>.json' (por perfil) -> nao repete
    clientes ja lancados; retoma de onde parou; pula quem tem tel/numero invalido.
  - Em caso de erro num cliente: print + print em 'debug/' + pula para o proximo.

============================ VARIAVEIS DE AMBIENTE ============================
  ASA_PERFIL    perfil de projeto: UEI (AGLAIR) ou UAR (YAN). Ver dict PERFIS.
  ASA_PLANILHA  caminho do .xlsx dos clientes.
  ASA_MODO_TESTE 1=preenche 1 e NAO finaliza (teste) | 0=finaliza de verdade.
  ASA_LIMITE    quantos clientes processar (0=todos).
  ASA_OFFSET    slots ja usados antes destes (desloca a agenda p/ nao conflitar).
  ASA_CDP       porta de depuracao de um Chrome ja aberto (ex.: 9223) p/ ANEXAR
                (senao, lanca o perfil dedicado 'perfil-chrome' sozinho).
  ASA_HEADLESS  1|0 (so no modo lanca-proprio, sem ASA_CDP).
  --- agenda ---
  ASA_DURACAO   minutos por atendimento (UEI=60, UAR=45).
  ASA_FOLGA     minutos de folga entre atendimentos.
  ASA_FIM       ultimo TERMINO permitido "HH:MM" (o ASA barra fora de 07:00-18:59).
  ASA_DATAS     datas especificas "DD/MM/AAAA,..." ou "DD/MM/AAAA@HH:MM" (inicio
                proprio no dia). Vazio = dias uteis a partir de DATA_INICIO.
  ASA_CEP_PADRAO CEP usado quando a planilha nao traz (padrao 69630-000).

Exemplo (lote UAR, anexando ao Chrome na porta 9223):
  set ASA_PERFIL=UAR & set ASA_CDP=9223 & set ASA_DURACAO=45 & set ASA_FOLGA=1
  set ASA_FIM=18:59 & set ASA_DATAS=27/07/2026,28/07/2026,29/07/2026
  set ASA_PLANILHA=C:\\...\\clientes.xlsx & set ASA_MODO_TESTE=0
  python lancar_atendimentos.py

NOTAS IMPORTANTES (aprendidas na pratica):
  - Se o "Validar Horarios" der "Conflito de Data e Hora" FALSO (data sem
    lancamento), a INSTANCIA do Chrome corrompeu: abra uma NOVA (outro
    user-data-dir + outra porta) e anexe nela. Conflito REAL = o atendente ja
    tem atendimento naquele dia/hora.
  - Planilha sem coluna CNAE/DATA ABERTURA: clientes cujo CNPJ precisa cadastro
    travam (atividade/data vazias) -> sao pulados; complete com reservas.
  - Telefone com DDD invalido (comeca com 0) e' rejeitado e o cliente e' pulado.
"""

from __future__ import annotations
import os, sys, json, math, time, random, traceback
from pathlib import Path
from datetime import date, datetime, timedelta

import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ------------------------------------------------------------------ CONFIG ---
PASTA   = Path(__file__).resolve().parent
PERFIL  = str(PASTA / "perfil-chrome")
DEBUG   = PASTA / "debug";  DEBUG.mkdir(exist_ok=True)
URL_APP = "https://asa-externo.am.sebrae.com.br/"
URL_NOVO = "https://asa-externo.am.sebrae.com.br/atendimento"

def _envbool(nome, padrao):
    v = os.environ.get(nome)
    return padrao if v is None else v.strip() in ("1", "true", "True", "sim")

MODO_TESTE = _envbool("ASA_MODO_TESTE", True)          # True = nao finaliza
HEADLESS   = _envbool("ASA_HEADLESS", False)
PLANILHA   = os.environ.get("ASA_PLANILHA", str(PASTA / "clientes.xlsx"))
LIMITE     = int(os.environ.get("ASA_LIMITE", "0"))    # 0 = todos
# Slots ja usados antes destes clientes (desloca a agenda). Ex.: 7 ja lancados
# -> os novos comecam no 8o horario. Assim nao conflita com o que ja existe.
OFFSET     = int(os.environ.get("ASA_OFFSET", "0"))
# ASA_REFAZER=1 -> reprocessa clientes mesmo que ja estejam em 'feitos' (usar p/
# remarcar/relançar em nova data; nao remove o registro antigo — isso e manual).
REFAZER    = os.environ.get("ASA_REFAZER", "0") == "1"
# Se setado (ex.: ASA_CDP=9222), ANEXA a um Chrome ja aberto com essa porta de
# depuracao (o SEU perfil), em vez de lancar um perfil isolado. Assim os
# lancamentos caem no seu app e usam sua sessao/sincronizacao.
CDP_PORT   = os.environ.get("ASA_CDP", "").strip()

# ---- PERFIS de projeto (valores fixos por projeto/consultor) ----------------
# Cada campo e (termo_de_busca, texto_da_opcao). termo="" = nao digita, so rola.
# solucao=None quando o tipo do atendimento nao tem o campo Solucao (ex.: Orientacao).
# Selecione o perfil ativo com a variavel ASA_PERFIL (ex.: ASA_PERFIL=UEI).
PERFIS = {
    "UEI": {
        "atendente":  ("AGLAIR", "AGLAIR ARAUJO LIMA"),
        "unidade":    ("UEI", "Unidade de Empreendedorismo do Interior"),
        "projeto":    ("TABATINGA", "TABATINGA"),
        "canal":      ("", "Espaço Sebrae"),
        "tipo":       ("", "Regularização"),
        "solucao":    ("", "certid"),
        "tema":       ("Finan", "Finanças"),
        "subtema":    ("Financeira", "Gestão Financeira"),
        "orientacao": "Cliente deve visitar o escritório do Sebrae em Tabatinga para manter seu CNPJ regular",
        "descricao":  "Cliente recebeu consultoria de gestão financeira referente a seu CNPJ para regularização",
    },
    "UAR": {
        "atendente":  ("YAN", "YAN PEREIRA FERREIRA"),
        "unidade":    ("UAR", "Unidade de Atendimento e Relacionamento"),
        "projeto":    ("DESCENTRALIZADO", "DESCENTRALIZADO E PARCEIROS"),
        "canal":      ("", "Agência fixa"),
        "tipo":       ("", "Orientação"),
        "solucao":    None,   # Orientacao nao tem campo Solucao
        "tema":       ("Finan", "Finanças"),
        "subtema":    ("Financeira", "Gestão Financeira"),
        "orientacao": "Visite o Sebrae Aleixo para mais informações sobre cursos, portfólio Sebrae e manter seu CNPJ regularizado",
        "descricao":  "Cliente recebeu orientação sobre leis e normas do MEI e foi encaminhada para consultoria",
    },
}
# Perfil CUSTOM vindo do painel web (perfil_custom.json). Mesmos campos do dict acima.
# Cada combo pode vir como [buscar, opcao] (lista) e vira tupla; solucao "" ou null = None.
_CUSTOM_JSON = PASTA / "perfil_custom.json"
if _CUSTOM_JSON.exists():
    try:
        _d = json.loads(_CUSTOM_JSON.read_text(encoding="utf-8"))
        for _k in ("atendente", "unidade", "projeto", "canal", "tipo", "tema", "subtema"):
            if isinstance(_d.get(_k), list):
                _d[_k] = tuple(_d[_k])
        _sol = _d.get("solucao")
        _d["solucao"] = tuple(_sol) if isinstance(_sol, list) and _sol and _sol[1] else None
        PERFIS["CUSTOM"] = _d
    except Exception as _e:
        print("aviso perfil_custom.json:", _e)

PERFIL = os.environ.get("ASA_PERFIL", "UAR").upper()   # perfil ativo (padrao: UAR)
if PERFIL not in PERFIS:
    raise SystemExit(f"Perfil '{PERFIL}' invalido. Use um de: {list(PERFIS)}")
CFG = PERFIS[PERFIL]
# CEP do municipio usado quando a planilha nao trouxer o CEP (Benjamin Constant = 69630-000)
CEP_PADRAO = os.environ.get("ASA_CEP_PADRAO", "69630-000")

# ---- Textos VARIADOS (rotaciona 1 servico por cliente) ----------------------
# Se existir textos_variados.json (lista de {servico, orientacao, descricao}), o
# script ALTERNA entre eles por cliente (rotacao sequencial, cobrindo todos os
# servicos de forma equilibrada). Senao, usa o texto fixo do perfil (CFG).
_TXT_JSON = PASTA / "textos_variados.json"
TEXTOS_VARIADOS = []
if _TXT_JSON.exists():
    try:
        TEXTOS_VARIADOS = [t for t in json.loads(_TXT_JSON.read_text(encoding="utf-8"))
                           if t.get("orientacao") and t.get("descricao")]
        if os.environ.get("ASA_EMBARALHAR", "1") == "1":
            random.shuffle(TEXTOS_VARIADOS)   # ordem NOVA a cada rodada (varia entre execucoes)
            _ordem = "embaralhada nesta rodada"
        else:
            _ordem = "na ordem do arquivo"
        if TEXTOS_VARIADOS:
            print(f"Textos variados: {len(TEXTOS_VARIADOS)} servicos ({_ordem}).")
    except Exception as _e:
        print("aviso textos_variados.json:", _e)

def escolher_texto(idx):
    """(orientacao, descricao) variando por cliente. Rotacao sequencial pelo pool."""
    pool = CFG.get("textos") or TEXTOS_VARIADOS
    if pool:
        t = pool[(idx - 1) % len(pool)]
        return (t.get("orientacao") or CFG["orientacao"],
                t.get("descricao")  or CFG["descricao"])
    return CFG["orientacao"], CFG["descricao"]

# ---- Agenda (configuravel por env) ------------------------------------------
def _hhmm(s, padrao):
    try:
        h, m = str(s).split(":"); return (int(h), int(m))
    except Exception:
        return padrao
DATA_INICIO = date(2026, 7, 23)
HORA_INICIO = (7, 0)
FOLGA_MIN   = int(os.environ.get("ASA_FOLGA", "5"))       # minutos de folga entre atendimentos
DUR_MIN     = int(os.environ.get("ASA_DURACAO", "60"))    # minutos por atendimento (UAR=45)
HORA_FIM    = _hhmm(os.environ.get("ASA_FIM", "18:00"), (18, 0))  # ultimo TERMINO permitido
# Datas especificas "DD/MM/AAAA,DD/MM/AAAA,...". Se vazio, usa dias uteis a partir de DATA_INICIO.
DATAS_ENV   = os.environ.get("ASA_DATAS", "").strip()
# Horarios EXPLICITOS, um por atendimento (na ordem): "DD/MM/AAAA HH:MM,DD/MM/AAAA HH:MM,..."
# Cada slot vira (data, inicio, fim=inicio+DUR_MIN). Use p/ preencher buracos no meio do dia.
SLOTS_ENV   = os.environ.get("ASA_SLOTS", "").strip()
# O calendario do sistema DESABILITA sabados e domingos, entao so dias uteis.
PULAR_FIM_DE_SEMANA = True

# ----------------------------------------------------------------- AGENDA ----
def slots_do_dia(inicio=None):
    slots = []
    t = datetime(2000, 1, 1, *(inicio or HORA_INICIO))
    limite = datetime(2000, 1, 1, *HORA_FIM)  # atendimento nao pode terminar depois de HORA_FIM
    while True:
        fim = t + timedelta(minutes=DUR_MIN)
        if fim > limite:
            break
        slots.append((t.strftime("%H:%M"), fim.strftime("%H:%M")))
        t = fim + timedelta(minutes=FOLGA_MIN)
    return slots

def gerar_agenda(qtd):
    agenda = []
    if SLOTS_ENV:                             # horarios explicitos, um por atendimento
        for entrada in [x.strip() for x in SLOTS_ENV.split(",") if x.strip()]:
            ds, hhmm = entrada.split()        # "DD/MM/AAAA HH:MM"
            h, m = map(int, hhmm.split(":"))
            fim = (datetime(2000, 1, 1, h, m) + timedelta(minutes=DUR_MIN)).strftime("%H:%M")
            agenda.append((ds, f"{h:02d}:{m:02d}", fim))
            if len(agenda) >= qtd:
                return agenda
        return agenda
    if DATAS_ENV:                             # datas especificas, na ordem dada
        # cada entrada: "DD/MM/AAAA" ou "DD/MM/AAAA@HH:MM" (inicio proprio naquele dia)
        for entrada in [x.strip() for x in DATAS_ENV.split(",") if x.strip()]:
            if "@" in entrada:
                ds, hhmm = entrada.split("@", 1)
                ds, ini_dia = ds.strip(), _hhmm(hhmm.strip(), HORA_INICIO)
            else:
                ds, ini_dia = entrada, HORA_INICIO
            for (ini, fim) in slots_do_dia(ini_dia):
                agenda.append((ds, ini, fim))
                if len(agenda) >= qtd:
                    return agenda
        return agenda                         # pode ficar < qtd se as datas nao comportarem
    slots = slots_do_dia()
    d = DATA_INICIO                           # senao: dias uteis consecutivos
    while len(agenda) < qtd:
        if not (PULAR_FIM_DE_SEMANA and d.weekday() >= 5):
            for (ini, fim) in slots:
                agenda.append((d.strftime("%d/%m/%Y"), ini, fim))
                if len(agenda) >= qtd:
                    break
        d += timedelta(days=1)
    return agenda

# -------------------------------------------------------------- PLANILHA -----
def _digitos(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return "".join(ch for ch in s if ch.isdigit())

def ler_planilha(caminho):
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.active
    linhas = list(ws.iter_rows(values_only=True))
    if not linhas:
        raise SystemExit("Planilha vazia.")
    cab = [str(c).strip().upper() if c is not None else "" for c in linhas[0]]

    def achar(*termos):
        for idx, nome in enumerate(cab):
            if any(t in nome for t in termos):
                return idx
        return None

    col_cpf   = achar("CPF")
    col_cnpj  = achar("CNPJ")
    col_nome  = achar("RESPONSAVEL", "RESPONSÁVEL", "NOME EMPRES")
    col_tel   = achar("TELEFONE", "FONE", "CELULAR")
    col_num   = achar("NUMERO", "NÚMERO", "NUMER")
    col_rua   = achar("ENDEREÇO", "ENDERECO", "LOGRADOURO", "RUA")
    col_bairro= achar("BAIRRO")
    col_cep   = achar("CEP")
    col_cid   = achar("MUNICÍPIO", "MUNICIPIO", "CIDADE")
    col_porte = achar("PORTE")
    col_dtab  = achar("DATA ABERTURA", "DATA DE ABERTURA", "DATA DE CRIA")  # evita 'TEMPO ABERTURA'
    col_ativ  = achar("CNAE PRINCIPAL", "ATIVIDADE ECON", "ATIVIDADE")
    col_cnae  = achar("CNAE")   # codigo (usado p/ buscar a atividade economica)
    if col_cpf is None:
        raise SystemExit(f"Nao achei coluna de CPF. Cabecalho: {cab}")

    def val(row, idx):
        if idx is None or row[idx] is None:
            return ""
        v = row[idx]
        if isinstance(v, float) and v != v:      # NaN
            return ""
        s = str(v).strip()
        return "" if s.lower() in ("nan", "none") else s

    clientes = []
    for row in linhas[1:]:
        cpf = _digitos(row[col_cpf]).zfill(11) if row[col_cpf] not in (None, "") else ""
        if not cpf or cpf == "0" * 11:
            continue
        cnpj = _digitos(row[col_cnpj]).zfill(14) if (col_cnpj is not None and row[col_cnpj]) else ""
        clientes.append({
            "cpf": cpf, "cnpj": cnpj,
            "nome": val(row, col_nome),
            "telefone": formatar_tel(val(row, col_tel)),
            "numero": val(row, col_num),
            "rua": val(row, col_rua),
            "bairro": val(row, col_bairro),
            "cep": val(row, col_cep) or CEP_PADRAO,
            "cidade": val(row, col_cid),
            "porte": val(row, col_porte),
            "data_abertura": val(row, col_dtab),
            "atividade": val(row, col_ativ),
            "cnae": _digitos(val(row, col_cnae)),
        })
    return clientes

def formatar_tel(v):
    """Digitos do telefone, garantindo 11 (celular). Se vier com 10 (DDD+8),
    insere o 9 do celular depois do DDD (ex.: (92) 8765-4321 -> (92) 98765-4321).
    Retorna '' se o DDD for invalido (comeca com 0), p/ o cliente ser pulado."""
    d = _digitos(v)
    if len(d) == 10:
        d = d[:2] + "9" + d[2:]
    if len(d) == 11 and d[0] == "0":   # DDD invalido (ex.: 09)
        return ""
    return d

# -------------------------------------------------------------- PROGRESSO ----
ARQ_PROG = PASTA / f"progresso_{PERFIL}.json"   # progresso separado por perfil
PAUSE_FLAG = PASTA / "pause.flag"   # painel cria p/ PAUSAR; remove p/ CONTINUAR
def carregar_progresso():
    if ARQ_PROG.exists():
        return json.loads(ARQ_PROG.read_text(encoding="utf-8"))
    return {"feitos": [], "erros": {}}
def salvar_progresso(p):
    ARQ_PROG.write_text(json.dumps(p, ensure_ascii=False, indent=2), encoding="utf-8")

# --------------------------------------------------- HELPERS PLAYWRIGHT -------
PATCH_STANDALONE = """
(() => {
  const orig = window.matchMedia ? window.matchMedia.bind(window) : null;
  window.matchMedia = (q) => {
    if (/display-mode:\\s*(standalone|minimal-ui|fullscreen)/.test(q)) {
      return { matches:true, media:q, onchange:null, addListener(){}, removeListener(){},
               addEventListener(){}, removeEventListener(){}, dispatchEvent(){return false;} };
    }
    return orig ? orig(q) : { matches:false, media:q, addListener(){}, removeListener(){},
             addEventListener(){}, removeEventListener(){}, dispatchEvent(){return false;} };
  };
  try { Object.defineProperty(window.navigator,'standalone',{get:()=>true}); } catch(e){}
})();
"""

def shot(page, nome):
    try:
        page.screenshot(path=str(DEBUG / f"{nome}.png"))
    except Exception:
        pass

def clicar_botao(page, texto, timeout=15000):
    """Clica um botao pelo texto (case-insensitive, contains)."""
    page.get_by_role("button", name=texto, exact=False).first.click(timeout=timeout)

def loc_input(page, id_):
    """Retorna o <input> real seja com id no proprio input ou num wrapper div#id."""
    return page.locator(f"input#{id_}").or_(page.locator(f"#{id_} input")).first

class SessaoExpirada(Exception):
    pass

def checar_sessao(page):
    """Levanta SessaoExpirada se a tela de login aparecer (sessao caiu)."""
    try:
        if (page.locator("input[type=password]").count() > 0 or
                page.get_by_role("button", name="Entrar", exact=True).count() > 0):
            raise SessaoExpirada()
    except SessaoExpirada:
        raise
    except Exception:
        pass

def descartar_retomar(page):
    """Se aparecer o dialogo 'retomar atendimento?', comeca do zero."""
    try:
        btn = page.get_by_role("button", name="começar do zero", exact=False)
        if btn.count() > 0:
            btn.first.click(timeout=4000)
            page.wait_for_timeout(800)
    except Exception:
        pass

def _tentar_combo(page, input_id, termo, alvo, timeout):
    """Uma tentativa: abre o select, opcionalmente digita, rola a(s) lista(s) e clica.
    Procura a opcao em TODOS os dropdowns visiveis (evita pegar um dropdown vazio
    remanescente de outro campo)."""
    try:
        page.keyboard.press("Escape")               # fecha dropdown remanescente
    except Exception:
        pass
    page.wait_for_timeout(60)
    sel = page.locator(f".ant-select:has(input#{input_id})")
    # centraliza o campo (senao pode ficar atras da barra fixa e o clique nao abre)
    try:
        sel.evaluate("el => el.scrollIntoView({block:'center'})")
    except Exception:
        pass
    page.wait_for_timeout(100)
    sel.locator(".ant-select-selector").first.click(force=True, timeout=timeout)
    page.wait_for_selector(".ant-select-dropdown:not(.ant-select-dropdown-hidden)", timeout=6000)
    page.wait_for_timeout(120)
    if termo:                                        # best-effort para filtrar listas grandes
        try:
            inp = page.locator(f"input#{input_id}")
            inp.fill("")
            inp.press_sequentially(termo, delay=15)
            page.wait_for_timeout(300)
        except Exception:
            pass
    vis = ".ant-select-dropdown:not(.ant-select-dropdown-hidden)"
    opt_all = page.locator(f"{vis} .ant-select-item-option")
    for _ in range(25):
        opt = opt_all.filter(has_text=alvo).first
        if opt.count() > 0:
            opt.scroll_into_view_if_needed()
            opt.click(timeout=timeout)
            page.wait_for_timeout(100)
            return True
        # rola todas as listas virtuais visiveis numa unica chamada (menos round-trips)
        movido = page.evaluate("""(vis) => {
            let moved = false;
            document.querySelectorAll(vis + ' .rc-virtual-list-holder').forEach(h => {
                const antes = h.scrollTop;
                h.scrollTop = h.scrollTop + h.clientHeight;
                if (h.scrollTop !== antes) moved = true;
            });
            return moved;
        }""", vis)
        if not movido:
            break                                    # nenhuma lista rolou -> fim
        page.wait_for_timeout(70)
    return False

def selecionar_combo(page, input_id, termo, opcao=None, timeout=20000):
    """
    Seleciona valor num Ant Design Select (id do input interno).
    Lida com listas virtualizadas (rola) e com selects sem busca por digitacao
    (tenta com termo; se falhar, reabre e procura sem digitar). Pula se ja ok.
    """
    alvo = opcao if opcao is not None else termo
    sel = page.locator(f".ant-select:has(input#{input_id})")
    item = sel.locator(".ant-select-selection-item")
    if item.count() and alvo:
        cur = (item.first.get_attribute("title") or item.first.inner_text() or "").lower()
        if alvo.lower()[:12] in cur:
            return                                   # ja esta certo
    sel.scroll_into_view_if_needed(timeout=timeout)
    tentativas = [termo, ""] if termo else [""]      # 2a tentativa sempre sem digitar
    # re-tenta em rodadas: opcoes de alguns selects carregam de forma assincrona
    for rodada in range(3):
        for t in tentativas:
            try:
                if _tentar_combo(page, input_id, t, alvo, timeout):
                    return
            except Exception:
                pass
            try:
                page.keyboard.press("Escape")         # fecha dropdown antes de re-tentar
            except Exception:
                pass
            page.wait_for_timeout(200)
        page.wait_for_timeout(300)                     # espera itens async carregarem
    raise RuntimeError(f"opcao '{alvo}' nao encontrada no combo #{input_id}")

def set_texto(page, id_, valor, enter=False):
    """Preenche um input de texto (limpando antes). Para campos mascarados."""
    inp = page.locator(f"input#{id_}")
    inp.scroll_into_view_if_needed()
    inp.click()
    page.keyboard.press("Control+a")
    page.keyboard.press("Delete")
    inp.press_sequentially(valor, delay=30)
    if enter:
        page.keyboard.press("Enter")
    page.wait_for_timeout(200)

def preencher_data(page, id_, dmy, timeout=15000):
    """Seleciona a data num Ant DatePicker (input readonly) pelo calendario.
    dmy no formato DD/MM/AAAA. As celulas tem title ISO (AAAA-MM-DD)."""
    d, m, y = [int(x) for x in dmy.split("/")]
    iso = f"{y:04d}-{m:02d}-{d:02d}"
    alvo_ym = f"{y:04d}-{m:02d}"
    inp = page.locator(f"input#{id_}")
    inp.scroll_into_view_if_needed()
    inp.click(force=True)
    page.wait_for_selector(".ant-picker-dropdown:not(.ant-picker-dropdown-hidden)", timeout=6000)
    page.wait_for_timeout(300)
    for _ in range(36):
        cell = page.locator(f".ant-picker-cell-in-view[title='{iso}']")
        if cell.count() > 0:
            if "ant-picker-cell-disabled" in (cell.first.get_attribute("class") or ""):
                raise RuntimeError(f"data {dmy} esta DESABILITADA no calendario "
                                   f"(fim de semana ou bloqueada pelo sistema).")
            cell.first.click()
            page.wait_for_timeout(300)
            if inp.input_value().strip() == dmy:
                return
            raise RuntimeError(f"cliquei em {dmy} mas o campo ficou "
                               f"'{inp.input_value()}'.")
        # navega ate o mes/ano alvo usando o mes exibido (title da 1a celula in-view)
        prim = page.locator(".ant-picker-cell-in-view").first
        atual_ym = (prim.get_attribute("title") or "")[:7]
        if atual_ym and atual_ym > alvo_ym:
            page.locator(".ant-picker-header-prev-btn").first.click()
        else:
            page.locator(".ant-picker-header-next-btn").first.click()
        page.wait_for_timeout(300)
    raise RuntimeError(f"nao consegui navegar ate a data {dmy}")

def na_tela_final(page):
    """Detecta a tela de Atendimento (tem o botao Validar Horarios)."""
    try:
        return page.get_by_role("button", name="Validar Horários", exact=False).count() > 0
    except Exception:
        return False

def marcar_checkbox(page, id_):
    """Marca um checkbox (Ant) pelo id do input interno, se ainda nao estiver marcado."""
    try:
        cb = page.locator(f"input#{id_}")
        if cb.count() == 0 or cb.is_checked():
            return
        box = page.locator(f".ant-checkbox-wrapper:has(input#{id_}) .ant-checkbox").first
        if box.count() > 0:
            box.click()
        else:
            cb.check(force=True)
        page.wait_for_timeout(400)
    except Exception as e:
        print(f"    aviso checkbox {id_}:", e)

def inferir_genero(nome):
    """Heuristica simples p/ nomes BR: 1o nome terminando em 'a' -> Feminino."""
    primeiro = (nome or "").strip().split()
    if primeiro and primeiro[0][-1:].lower() == "a":
        return "Feminino"
    return "Masculino"

def tratar_contatos(page, cli):
    """Na etapa Contatos & Enderecos: corrige telefone invalido e numero vazio
    com dados da planilha, e marca 'Mesmo endereco da pessoa fisica'."""
    # telefone: precisa ter 11 digitos; corrige se estiver invalido/vazio
    try:
        tel = page.locator("input#multi_Telefone_0_numero")
        if tel.count() > 0 and cli.get("telefone"):
            if len(_digitos(tel.input_value() or "")) != 11:
                tel.click()
                page.keyboard.press("Control+a"); page.keyboard.press("Delete")
                tel.press_sequentially(cli["telefone"], delay=25)
                page.wait_for_timeout(400)
    except Exception as e:
        print("    aviso telefone:", e)
    # campos de endereco obrigatorios (bairro, rua, cep, numero, cidade):
    # preenche da planilha SO os que estiverem vazios (nao sobrescreve auto-fill)
    for id_, chave in [
        ("multi_Endereco_0_cep",          "cep"),
        ("multi_Endereco_0_descEndereco", "rua"),
        ("multi_Endereco_0_numero",       "numero"),
        ("multi_Endereco_0_descBairro",   "bairro"),
        ("multi_Endereco_0_descCid",      "cidade"),
    ]:
        try:
            campo = page.locator(f"input#{id_}")
            if (campo.count() > 0 and cli.get(chave)
                    and not (campo.input_value() or "").strip()):
                # CEP: digita so os digitos (mascara) e espera a busca automatica
                valor = _digitos(cli[chave]) if chave == "cep" else str(cli[chave])
                campo.click()
                campo.press_sequentially(valor, delay=25)
                page.wait_for_timeout(1500 if chave == "cep" else 200)
        except Exception as e:
            print(f"    aviso {chave}:", e)
    # marca 'Mesmo endereco da pessoa fisica' (copia o endereco p/ o empreendimento)
    cb = page.locator("input#multi_Endereco_1_mesmoEnderecoPF")
    if cb.count() > 0:
        try:
            if not cb.is_checked():
                page.get_by_text("Mesmo endereço da pessoa física", exact=False).first.click()
                page.wait_for_timeout(1000)
        except Exception as e:
            print("    aviso 'mesmo endereco':", e)

def cadastrar_cnpj(page, cli):
    """Cadastra um novo CNPJ (Empresa) quando o CPF nao tem vinculo no ASA.
    Fluxo: 2 cliques em '+ Cadastrar novo CNPJ' (PF -> Empresa -> abre o form),
    depois preenche os campos com os dados da planilha."""
    page.get_by_role("button", name="Cadastrar novo CNPJ", exact=False).first.click()
    page.wait_for_timeout(1500)                        # PF -> aba Empresa
    page.get_by_role("button", name="Cadastrar novo CNPJ", exact=False).first.click()
    page.wait_for_timeout(2500)                        # Empresa -> abre o formulario
    page.wait_for_selector("input#multi_Pfpj_0_cnpj", timeout=12000)

    porte = (cli.get("porte") or "").upper()
    if "EPP" in porte or "PEQUENO" in porte:
        porte_opt, qtd = "Empresa de Pequeno Porte", "6"
    elif porte == "ME" or "MICRO" in porte:
        porte_opt, qtd = "Micro Empresa", "3"
    else:                                              # MEI (padrao)
        porte_opt, qtd = "MEI", "0"
    natureza = ("Sociedade Empresária Limitada"
                if "LTDA" in (cli.get("nome") or "").upper()
                else "Empresário (Individual)")

    set_texto(page, "multi_Pfpj_0_cnpj", _digitos(cli["cnpj"]))
    set_texto(page, "multi_Pfpj_0_razaoSocial", cli.get("nome") or "")
    set_texto(page, "multi_Pfpj_0_dataCriacaoRelatorio", cli.get("data_abertura") or "")
    selecionar_combo(page, "multi_Pfpj_0_descricaoStatusReceita", "", "Validado")
    selecionar_combo(page, "multi_Pfpj_0_descPorte", "", porte_opt)
    set_texto(page, "multi_Pfpj_0_quantidadeFuncionarios", qtd)
    selecionar_combo(page, "multi_Pfpj_0_descNaturezaJuridica", "", natureza)
    # Atividade economica (pesquisavel): busca pelo codigo CNAE; fallback = vestuario
    ativ_ok = False
    for termo in [cli.get("cnae") or "", (cli.get("atividade") or "")[:18]]:
        if not termo:
            continue
        try:
            selecionar_combo(page, "multi_Pfpj_0_atividade", termo, termo)
            ativ_ok = True
            break
        except Exception:
            pass
    if not ativ_ok:
        selecionar_combo(page, "multi_Pfpj_0_atividade",
                         "vestuário e acessórios", "vestuário e acessórios")
    selecionar_combo(page, "multi_Pfpj_0_tipoVinculo", "", "PROPRIETÁRIO OU SÓCIO")

# --------------------------------------------------- FLUXO DE UM CLIENTE ------
def processar_cliente(page, cli, agenda_item, idx):
    dt, h_ini, h_fim = agenda_item
    tag = f"{idx:03d}_{cli['cpf']}"

    # 0) abre wizard novo
    page.goto(URL_NOVO, wait_until="domcontentloaded")
    page.wait_for_timeout(1500)
    checar_sessao(page)              # para com aviso se a sessao expirou
    shot(page, f"{tag}_0landing")
    descartar_retomar(page)          # fecha dialogo 'retomar?' se aparecer

    # 1) CPF -> Buscar
    campo_cpf = loc_input(page, "multi_cpf")
    campo_cpf.wait_for(timeout=20000)
    campo_cpf.click()
    campo_cpf.press_sequentially(cli["cpf"], delay=30)   # aciona a mascara
    shot(page, f"{tag}_1cpf")
    clicar_botao(page, "Buscar")
    page.wait_for_timeout(3000)

    # 2) Dados da Pessoa Fisica (auto) -> genero + aceite LGPD -> Salvar e Prosseguir
    page.wait_for_selector("#multi_nome", timeout=20000)
    # genero: cliente autorizou "PREFIRO NAO INFORMAR" para todos (evita registrar errado)
    try:
        if page.locator(".ant-select:has(input#multi_sexo) .ant-select-selection-item").count() == 0:
            selecionar_combo(page, "multi_sexo", "", "informar")
    except Exception as e:
        print("    aviso genero:", e)
    marcar_checkbox(page, "multi_lgpd")   # 'Aceito a Politica de Privacidade' (obrigatorio)
    shot(page, f"{tag}_2pf")
    clicar_botao(page, "Salvar e Prosseguir")
    page.wait_for_timeout(2000)

    # 3) Empreendimento -> vincular CNPJ (se houver) -> Salvar e Prosseguir
    shot(page, f"{tag}_3empreend")
    if cli["cnpj"]:
        try:
            # abre o combo de CNPJ e escolhe a opcao do cliente (contem o numero)
            selecionar_combo(page, "rc_select_1", termo="", opcao=cli["cnpj"])
        except Exception as e:
            print("    CNPJ nao encontrado no vinculo:", e)
        page.wait_for_timeout(2500)   # ao vincular, o app carrega os dados da empresa
        # 'Tipo de vinculo' e obrigatorio e NAO vem preenchido. MEI => PROPRIETARIO OU SOCIO
        if page.locator("input#multi_Pfpj_0_tipoVinculo").count() > 0:
            try:
                selecionar_combo(page, "multi_Pfpj_0_tipoVinculo",
                                 "", "PROPRIETÁRIO OU SÓCIO")
            except Exception as e:
                print("    aviso tipo de vinculo:", e)
        else:
            # CNPJ nao esta cadastrado no ASA -> cadastra novo com dados da planilha
            print(f"    CNPJ {cli['cnpj']} nao vinculado -> cadastrando novo CNPJ...")
            cadastrar_cnpj(page, cli)
            page.wait_for_timeout(1000)
    shot(page, f"{tag}_3empreend_pos")
    clicar_botao(page, "Salvar e Prosseguir")
    page.wait_for_timeout(2500)

    # 4) + 5) telas intermediarias (Contatos & Enderecos): tratar e prosseguir
    for i in range(6):
        if na_tela_final(page):
            break
        tratar_contatos(page, cli)            # corrige tel/numero + 'mesmo endereco'
        shot(page, f"{tag}_prosseguir{i}")
        try:
            clicar_botao(page, "Salvar e Prosseguir", timeout=10000)
        except Exception as e:
            print("    prosseguir travou:", e)
            break
        page.wait_for_timeout(2000)

    if not na_tela_final(page):
        raise RuntimeError("Nao cheguei na tela de Atendimento (Validar Horarios).")

    # 6) Tela de Atendimento: campos do PERFIL ativo (CFG) + textos + data/hora
    shot(page, f"{tag}_6atend_antes")
    selecionar_combo(page, "multi_atendente",        *CFG["atendente"])
    selecionar_combo(page, "multi_unidade",          *CFG["unidade"])
    selecionar_combo(page, "multi_projetoAcao",      *CFG["projeto"])
    selecionar_combo(page, "multi_canalAtendimento", *CFG["canal"])
    selecionar_combo(page, "multi_tipoAtendimento",  *CFG["tipo"])
    if CFG.get("solucao"):        # alguns tipos (ex.: Orientacao) nao tem Solucao
        selecionar_combo(page, "multi_SolucaoFocoId", *CFG["solucao"])
    try:
        selecionar_combo(page, "multi_tema",    *CFG["tema"])
        selecionar_combo(page, "multi_subtema", *CFG["subtema"])
    except Exception as e:
        print("    aviso tema/subtema:", e)

    _ori, _desc = escolher_texto(idx)     # texto variado por cliente (rotaciona)
    page.locator("#multi_OrientacaoCliente").fill(_ori)
    page.locator("#multi_descricao").fill(_desc)

    preencher_data(page, "multi_dataAtendimento", dt)          # Ant DatePicker (calendario)
    set_texto(page, "multi_horaInicio", h_ini)
    set_texto(page, "multi_horaFim", h_fim)
    shot(page, f"{tag}_6atend_preenchido")

    clicar_botao(page, "Validar Horários")
    page.wait_for_timeout(2500)
    shot(page, f"{tag}_7validado")

    if MODO_TESTE:
        print(f"   [TESTE] cliente {cli['cpf']} preenchido em {dt} {h_ini}-{h_fim}. "
              f"NAO finalizei (modo teste).")
        return "teste"

    clicar_botao(page, "FINALIZAR ATENDIMENTO")
    page.wait_for_timeout(3000)
    shot(page, f"{tag}_8finalizado")
    return "ok"

# ----------------------------------------------------------------- MAIN ------
def main():
    if not Path(PLANILHA).exists():
        raise SystemExit(f"Planilha nao encontrada: {PLANILHA}\n"
                         f"Defina ASA_PLANILHA ou coloque 'clientes.xlsx' na pasta.")
    clientes = ler_planilha(PLANILHA)
    # pula clientes sem telefone valido (11 dig.) ou sem numero (o ASA nao preenche esses)
    validos, pulados = [], []
    for c in clientes:
        if len(c["telefone"]) != 11 or not str(c["numero"]).strip():
            pulados.append(c)
        else:
            validos.append(c)
    if pulados:
        print(f"PULADOS {len(pulados)} sem telefone/numero: "
              + ", ".join(c["cpf"] for c in pulados))
    clientes = validos
    if LIMITE:
        clientes = clientes[:LIMITE]
    if MODO_TESTE:
        clientes = clientes[:1]
    # desloca a agenda pelos slots ja usados (OFFSET) para nao conflitar horarios
    agenda = gerar_agenda(OFFSET + len(clientes))[OFFSET:]
    if len(agenda) < len(clientes):
        print(f"*** ATENCAO: a agenda comporta {len(agenda)} horarios, mas ha "
              f"{len(clientes)} clientes. {len(clientes)-len(agenda)} ficarao SEM "
              f"horario (adicione mais datas em ASA_DATAS).")
        clientes = clientes[:len(agenda)]

    print(f"Planilha : {PLANILHA}")
    print(f"Perfil   : {PERFIL} (atendente {CFG['atendente'][1]}, unidade {CFG['unidade'][1]})")
    print(f"Clientes : {len(clientes)} | MODO_TESTE={MODO_TESTE} | HEADLESS={HEADLESS}")
    print(f"Agenda   : {agenda[0][0]} {agenda[0][1]}  ...  {agenda[-1][0]} {agenda[-1][2]}")

    prog = carregar_progresso()
    with sync_playwright() as p:
        if CDP_PORT:
            print(f"Anexando ao Chrome via CDP porta {CDP_PORT} (seu perfil)...")
            # usa 127.0.0.1 (IPv4) em vez de localhost: no Windows 'localhost'
            # resolve p/ IPv6 ::1 e o Chrome escuta so em IPv4 -> ECONNREFUSED ::1
            br = p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
            ctx = br.contexts[0]
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
        else:
            ctx = p.chromium.launch_persistent_context(
                user_data_dir=PERFIL, channel="chrome", headless=HEADLESS,
                args=["--app=" + URL_APP, "--window-size=1400,1000"], no_viewport=True)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
        ctx.add_init_script(PATCH_STANDALONE)

        for idx, (cli, ag) in enumerate(zip(clientes, agenda), 1):
            if PAUSE_FLAG.exists():                 # PAUSA entre clientes (nao perde o atual)
                print("[PAUSA] rodada pausada — clique em Continuar no painel para retomar.")
                while PAUSE_FLAG.exists():
                    time.sleep(1.5)
                print("[RETOMANDO] seguindo os lancamentos.")
            if cli["cpf"] in prog["feitos"] and not REFAZER:
                print(f"[{idx}] {cli['cpf']} ja feito, pulando.")
                continue
            print(f"[{idx}/{len(clientes)}] CPF {cli['cpf']} -> {ag[0]} {ag[1]}-{ag[2]}")
            try:
                r = processar_cliente(page, cli, ag, idx)
                if r == "ok":
                    if cli["cpf"] not in prog["feitos"]:
                        prog["feitos"].append(cli["cpf"])
                    prog["erros"].pop(cli["cpf"], None)
                    salvar_progresso(prog)
            except SessaoExpirada:
                print("\n*** SESSAO EXPIROU ***  Faca login na janela do Chrome e "
                      "rode o script de novo — ele RETOMA de onde parou "
                      f"({len(prog['feitos'])} ja feitos).")
                break
            except Exception as e:
                print(f"    ERRO: {e}")
                shot(page, f"ERRO_{idx:03d}_{cli['cpf']}")
                prog["erros"][cli["cpf"]] = str(e)
                salvar_progresso(prog)
                traceback.print_exc()

        print("\nFim. Feitos:", len(prog["feitos"]), "| Erros:", len(prog["erros"]))
        if CDP_PORT:
            print("(modo CDP: deixei o seu navegador aberto)")
        else:
            if not HEADLESS and not MODO_TESTE:
                input("ENTER para fechar o navegador... ")
            ctx.close()


if __name__ == "__main__":
    main()
