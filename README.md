# 🤖 Playwright RPA — Automação de Atendimentos

![Python](https://img.shields.io/badge/Python-3.14-3776AB?logo=python&logoColor=white)
![Playwright](https://img.shields.io/badge/Playwright-RPA-2EAD33?logo=playwright&logoColor=white)
![Vanilla JS](https://img.shields.io/badge/UI-HTML%2FCSS%2FJS%20puro-F7DF1E?logo=javascript&logoColor=black)

Robô (RPA) que **lança atendimentos em massa** em um sistema web corporativo (o
ASA Externo, PWA do Sebrae-AM) a partir de planilhas — com um **painel web local**
para configurar e disparar as rodadas. Transforma um trabalho manual e repetitivo
de **horas em minutos**, mantendo o login e os dados sensíveis sob controle do usuário.

> Projeto real, usado em produção para dar conta de metas de atendimento.
> Os dados de clientes (CPF/CNPJ) e as sessões de login **nunca** entram no repositório.

---

## 💡 O problema

Registrar centenas de atendimentos à mão em um sistema web é lento e sujeito a erro:
cada lançamento exige percorrer um **wizard de 5 etapas** (buscar cliente, vincular/
cadastrar CNPJ, contatos, endereço, classificação e horário) e o sistema é uma **PWA
offline-first** com menus virtualizados e calendário travado.

## ⚙️ A solução

- **Robô em Python + Playwright** que dirige o próprio app do Chrome de ponta a ponta,
  preenchendo e finalizando cada atendimento a partir da planilha.
- **Painel web local** (sem dependências além do Python) para escolher unidade, projeto,
  tema, textos e os parâmetros da rodada — e acompanhar o **log ao vivo**.
- **Textos variados** que rotacionam por cliente (cobrindo todos os serviços), para os
  registros não ficarem repetitivos.

## ✨ Destaques técnicos

- **Anexa via CDP ao Chrome já logado** — o login é sempre humano; o robô **nunca**
  manipula senhas.
- Contorna o **gate de instalação da PWA** (`matchMedia: display-mode standalone`)
  injetando patch com `add_init_script`.
- Lida com **menus Ant Design virtualizados** (rolagem até a opção) e **DatePicker
  readonly** (navegação por calendário).
- **Config-driven**: perfis de projeto (`PERFIS`/perfil `CUSTOM`) + variáveis de ambiente;
  agenda por **datas** ou **horários exatos** (`ASA_SLOTS`), duração, folga e término.
- **Agenda inteligente**: só dias úteis, respeita janelas livres e o horário comercial.
- **Idempotência e retomada**: progresso persistido (não repete), detecção de **sessão
  expirada** com aviso, prints de debug em cada erro.
- **Painel web em `http.server`** puro (stdlib), com editor de textos e botão *randomizar*.
- **Privacidade por padrão**: `.gitignore` blinda planilhas, perfis do Chrome, progresso
  e logs — nada de dado pessoal versionado.

## 🖥️ Painel de controle

Uma página local (`http://localhost:8760`) monta a configuração e dispara o script,
com log em tempo real:

<!-- screenshot do painel -->
![Painel](docs/painel.png)

## 🧱 Stack

`Python 3` · `Playwright` · `openpyxl` · `HTML/CSS/JS` (vanilla) · `http.server` (stdlib)

## 🚀 Como rodar

```bash
python -m pip install -r requirements.txt
python -m playwright install chrome
```

1. Abra o sistema no Chrome com porta de depuração e **faça login**.
2. Rode o painel (duplo clique em `PAINEL.bat`) **ou** direto:

```bash
ASA_PERFIL=UAR ASA_CDP=9224 ASA_PLANILHA=clientes.xlsx python lancar_atendimentos.py
```

Guia operacional completo em [`PROCEDIMENTOS.md`](PROCEDIMENTOS.md).

## 📁 Estrutura

```
lancar_atendimentos.py   # robô principal (Playwright)
painel.py / PAINEL.bat   # painel web local de controle
textos_variados.json     # pool de textos de orientação (rotacionam)
PROCEDIMENTOS.md         # guia operacional
requirements.txt
```

## 🔒 Privacidade & segurança

Planilhas de clientes, perfis do Chrome (sessões de login), progresso e logs ficam
**fora do repositório** (ver `.gitignore`). O robô nunca digita nem armazena senhas —
a autenticação é feita manualmente pelo operador na janela do navegador.
