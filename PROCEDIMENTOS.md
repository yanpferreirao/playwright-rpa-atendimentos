# Procedimentos — Automação ASA Externo (Sebrae-AM)

Guia operacional para lançar atendimentos automaticamente no **ASA Externo**
(`https://asa-externo.am.sebrae.com.br/`).

Pasta do projeto: `C:\Users\yanma\asa-automacao`
Script principal: `lancar_atendimentos.py`

---

## 0. MODO AUTÔNOMO (clique duplo) — jeito recomendado

Dois atalhos na pasta do projeto:

- **`1-TESTAR (preenche 1, nao finaliza).bat`** — preenche 1 cliente e PARA antes de
  finalizar (nada é criado). Use para conferir se está tudo certo.
- **`2-RODAR TUDO (finaliza de verdade).bat`** — processa TODOS os clientes da
  planilha e finaliza de verdade.

Passo a passo (vale para os dois):
1. Dê **duplo clique** no `.bat`.
2. **Arraste a planilha (.xlsx)** para a janela preta e tecle ENTER.
3. Abre a janela do ASA. Se pedir, **faça o login** (usuário/senha).
4. Quando estiver vendo a **lista de atendimentos**, volte à janela preta e tecle ENTER.
5. O script roda sozinho. Ele **pula quem já foi feito** e, se a sessão cair,
   **para com aviso** — é só relogar e rodar de novo (retoma de onde parou).

Cada lançamento leva ~40 s. Pode deixar rodando. Ao final, confira a lista no app.
Depois, faça o seu processo normal de sincronização / "foco" para gerar os protocolos.

O que o script preenche/corrige sozinho:
- vínculo de CNPJ existente, ou **cadastro de CNPJ novo** quando o CPF não tem
  vínculo (2 cliques em "+ Cadastrar novo CNPJ": aba Pessoa Física → Empresa →
  abre o form; preenche CNPJ, Razão Social, Data de abertura, Situação=Validado,
  Porte, Qtd funcionários [MEI=0/ME=3/EPP=6], Natureza [Empresário Individual, ou
  Sociedade Empresária Limitada se "LTDA" na razão social], Atividade pelo CNAE,
  Tipo de vínculo = Proprietário ou Sócio);
- tipo de vínculo (PROPRIETÁRIO OU SÓCIO), gênero (PREFIRO NÃO INFORMAR), aceite LGPD;
- telefone inválido (→ 11 dígitos, insere o 9 do celular), campos de endereço
  obrigatórios vazios (CEP/rua/número/bairro/cidade da planilha), CEP faltante
  usa o do município (69630-000, env `ASA_CEP_PADRAO`), "mesmo endereço";
- data pelo calendário e a classificação fixa da UEI.
Trata valores vazios/`nan` da planilha (não cola "nan" em campo nenhum).

---

## 1. Pré-requisitos (já instalados)

- Python 3.14 + bibliotecas `playwright` e `openpyxl`
  (reinstalar se preciso: `python -m pip install playwright openpyxl`)
- Google Chrome instalado
- Planilha de clientes (.xlsx) com colunas contendo **CPF** e **CNPJ**
  (o script detecta as colunas pelo cabeçalho e corrige CPF/CNPJ sem zero à esquerda)

---

## 2. Como a agenda é montada (regra de horários)

- Cada atendimento dura **1 hora**, com **5 min de folga** entre um e outro.
- Começa às **07:00**; **10 atendimentos por dia** (último 16:45–17:45).
  Horários: 07:00 · 08:05 · 09:10 · 10:15 · 11:20 · 12:25 · 13:30 · 14:35 · 15:40 · 16:45.
- **Só dias úteis** (o calendário do sistema BLOQUEIA sábado e domingo).
- Início em **23/07/2026**; 80 clientes → 23/07 a 03/08/2026 (todas datas passadas).
- Para mudar: editar no topo de `lancar_atendimentos.py` as variáveis
  `DATA_INICIO`, `HORA_INICIO`, `FOLGA_MIN`, `DUR_MIN`, e `slots_do_dia()`.

## 3. Campos fixos gravados em TODOS os atendimentos

Atendente = AGLAIR ARAUJO LIMA · Unidade = UEI · Projeto = AM2026 CIDADE
EMPREENDEDORA (TABATINGA) · Canal = Espaço Sebrae · Tipo = Regularização ·
Solução = Emissão de certidões · Tema = Finanças / Gestão Financeira ·
Tipo de vínculo = PROPRIETÁRIO OU SÓCIO. Textos de orientação/relato padrão
(editáveis em `FIXOS`, `ORIENTACAO`, `DESCRICAO` no início do script).

---

## 4. Como RODAR

O login (Keycloak/AMEI) é sempre feito por **você** — o script nunca digita senha.

### Modo A — perfil dedicado (o que já usamos)
1. Fechar processos Chrome que travem o perfil:
   ```
   Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" | Where-Object { $_.CommandLine -like "*perfil-chrome*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
   ```
2. Rodar (no PowerShell da pasta do projeto):
   ```
   $env:ASA_PLANILHA = "C:\caminho\da\planilha.xlsx"
   $env:ASA_MODO_TESTE = "1"   # 1 = preenche mas NÃO finaliza (seguro p/ testar)
   $env:ASA_HEADLESS   = "0"   # 0 = janela visível
   python lancar_atendimentos.py
   ```
3. Na 1ª vez, fazer login na janela do Chrome quando pedir (a sessão fica salva
   em `perfil-chrome`, mas **expira com o tempo** — se pedir login de novo, logar).
4. Conferir os prints em `debug\` e, se estiver tudo certo, rodar **pra valer**:
   ```
   $env:ASA_MODO_TESTE = "0"     # 0 = FINALIZA de verdade (cria registro)
   $env:ASA_LIMITE     = "1"     # começar com 1 cliente; depois remover p/ todos
   python lancar_atendimentos.py
   ```

### Modo B — anexar a um Chrome já aberto (CDP)
> Só funciona em perfil **NÃO-padrão** (o Chrome bloqueia depuração no perfil Default).
1. Abrir o app com porta de depuração:
   ```
   & "C:\Program Files\Google\Chrome\Application\chrome.exe" --user-data-dir="C:\Users\yanma\asa-automacao\perfil-chrome" --app="https://asa-externo.am.sebrae.com.br/" --remote-debugging-port=9222
   ```
2. Rodar apontando para a porta:
   ```
   $env:ASA_CDP = "9222"
   python lancar_atendimentos.py
   ```

### Variáveis de ambiente
`ASA_PLANILHA` (caminho) · `ASA_MODO_TESTE` (1/0) · `ASA_HEADLESS` (1/0) ·
`ASA_LIMITE` (nº de clientes; 0 = todos) · `ASA_CDP` (porta CDP).

---

## 5. Progresso e reexecução

- O script salva `progresso.json`: clientes já finalizados NÃO se repetem.
- Para recomeçar do zero, apagar `progresso.json`.
- Em erro, ele tira print `debug\ERRO_*.png`, registra e pula para o próximo.

## 6. Gerar o relatório oficial (PDF) para conferência

No app: botão **"..."** (ao lado de Novo Atendimento) → **Relatórios** → ajustar
período → **Exportar PDF**. O relatório só lista atendimentos que **geraram
protocolo** (ou seja, que sincronizaram).

## 7. IMPORTANTE — sincronização / protocolo

- Finalizar o atendimento no app deixa ele como **"Enviado"**, mas ele só vira
  **oficial** quando **sincroniza e gera protocolo**.
- A sincronização tenta ser automática; se não gerar protocolo, o procedimento
  é ir ao **portal do colaborador** `https://asa.am.sebrae.com.br/#/colaborador/home-atendimento`
  e clicar em **"foco"**.
- **Pendência atual:** um registro de teste finalizou como "Enviado" mas não
  gerou protocolo (relatório veio vazio) — em verificação com o responsável do
  Sebrae para saber se é problema temporário do próprio ASA.

## 8. Limitações conhecidas

- Não dá para automatizar o **seu perfil padrão** do Chrome (bloqueio de
  segurança do Chrome v136+). Por isso usa-se um perfil dedicado.
- A sessão do perfil dedicado **expira** — pode ser preciso relogar antes de um
  lote grande.

---

## 9. Perfis, agenda configurável e casos especiais (atualização ago/2026)

**Perfis (`ASA_PERFIL`)** — todos os valores fixos por projeto ficam no dict `PERFIS`
no topo do `lancar_atendimentos.py`:
- `UEI` — AGLAIR, Regularização, Solução "Emissão de certidões" (projeto Tabatinga).
- `UAR` — YAN, Unidade de Atendimento e Relacionamento, projeto Atendimento
  Descentralizado, tipo **Orientação (sem campo Solução)**, canal Agência fixa.

**Agenda configurável por variáveis de ambiente:**
- `ASA_DURACAO` minutos por atendimento (UEI=60, UAR=45)
- `ASA_FOLGA` minutos de folga entre um e outro
- `ASA_FIM` último término permitido "HH:MM" (o ASA **barra** fora de 07:00–18:59)
- `ASA_DATAS` datas específicas: `13/07/2026,14/07/2026,...` — e para um dia começar
  em outro horário use `24/07/2026@11:10` (ex.: dia com atendimento existente de manhã)
- `ASA_OFFSET` pula N slots já usados (para não conflitar horários)

**Anexar a um Chrome já aberto (`ASA_CDP`):** abra o app com
`chrome.exe --user-data-dir=<pasta> --app=<url> --remote-debugging-port=9223` e rode
com `set ASA_CDP=9223`. Assim o script dirige a janela onde VOCÊ logou.

**⚠️ Conflito de horário FALSO:** se o "Validar Horários" acusar
"Conflito de Data e Hora" numa data que está livre, a **instância do Chrome
corrompeu**. Abra uma **instância nova e limpa** (outro `--user-data-dir` e outra
`--remote-debugging-port`), logue nela e use a nova porta em `ASA_CDP`. (Conflito
REAL = o atendente já tem atendimento naquele dia/hora.)

**Lacunas de dados da planilha:**
- Telefone com DDD inválido (começa com 0) → cliente é **pulado** automaticamente.
- CNPJ que precisa cadastro mas planilha sem CNAE/DATA ABERTURA → **pulado**
  (não dá para preencher atividade/data); complete com clientes reserva ou
  adicione as colunas CNAE e DATA ABERTURA na planilha.
