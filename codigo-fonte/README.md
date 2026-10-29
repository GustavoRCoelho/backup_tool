# Ferramenta de Backup em Massa de Caixas (IMAP -> .pst)

Baixa, via IMAP, as caixas de correio de uma planilha de contas e consolida cada
uma em um arquivo `.pst` local, usando o **Redemption (RDO)** para criar o PST e
importar as mensagens preservando data, remetente e anexos.

> A troca de senha de cada conta (para uma senha comum) e a exclusao da caixa no
> servidor **sao feitas manualmente pelo administrador** e **nao fazem parte
> desta ferramenta**. A ferramenta so entra em acao depois que a senha ja foi
> trocada.

---

## 0. Estrutura das pastas

```
backup_tool/
  LEIA-ME.txt          -> ponto de partida (explica as duas pastas)
  programa/            -> PARA USAR (distribua esta pasta)
      backup_gui.exe   -> interface grafica (duplo clique)
      backup_tool.exe  -> versao linha de comando (CLI)
      config.ini       -> configuracao em uso (edite aqui)
      exemplo.csv      -> planilha modelo
  codigo-fonte/        -> CODIGO e DOCUMENTACAO
      *.py             -> codigo-fonte Python
      README.md        -> este arquivo
      requirements.txt
      build.ps1        -> gera os .exe em ..\programa (apos alterar o codigo)
      config.ini / exemplo.csv -> modelos
```

Para **usar**: vá na pasta `programa\`. Para **alterar/compilar**: `codigo-fonte\`.

---

## 1. Requisitos da maquina de destino

- **Windows** com **Outlook 2010 instalado** (fornece o subsistema MAPI usado
  pelo Redemption e permite abrir o `.pst` depois).
- **Redemption registrado** (COM in-process). Nesta maquina os arquivos estao em
  `C:\Users\admin\Downloads\Redemption\`.
- **Arquitetura: 64 bits** em todos os componentes. **(Correcao em relacao ao
  prompt original, que supunha 32 bits.)** Como o seu Outlook e **64 bits**, use:
  - **Python 64 bits**,
  - **`Redemption64.dll`** (e nao a `Redemption.dll` de 32 bits),
  - PyInstaller gerando um `.exe` **64 bits**.
  Redemption e in-process: a arquitetura de Python/.exe tem de bater com a do
  Outlook. Misturar 32 e 64 bits resulta em erro de carga do COM.

### Registrar o Redemption (uma vez, como Administrador)

```powershell
regsvr32.exe "C:\Users\admin\Downloads\Redemption\Redemption64.dll"
```

(Para desinstalar: `regsvr32.exe /u "...\Redemption64.dll"`.)

---

## 2. Instalacao (ambiente de desenvolvimento/execucao com Python)

```powershell
# No diretorio backup_tool, com Python 64 bits:
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> Observacao: nesta maquina o `python` ainda nao esta instalado de fato (o
> `python.exe` em `WindowsApps` e so o atalho da Microsoft Store). Instale o
> **Python 3.x 64 bits** (ha um instalador `python-3.15.2-amd64.exe` em
> `Downloads`, que e a versao amd64/64 bits) antes de rodar.

---

## 3. Configuracao (`config.ini`)

Fica **externo ao `.exe`**, para mudar parametros sem recompilar.

| Secao      | Chave      | Descricao |
|------------|------------|-----------|
| `[imap]`   | `host`     | Servidor IMAP (padrao: `imap.skymail.net.br`). |
| `[imap]`   | `porta`    | Porta IMAP SSL (padrao: `993`). |
| `[imap]`   | `timeout`  | Timeout de socket em segundos. |
| `[imap]`   | `lote`     | De quantas em quantas mensagens o estado e salvo (checkpoint de retomada). |
| `[saida]`  | `base`     | Pasta base de saida. **Nao use a Area de Trabalho** (OneDrive/outros usuarios). |
| `[pastas]` | `ignorar`  | Pastas a ignorar no download, separadas por `;` (ex.: Gmail "Todos os e-mails"). |

---

## 4. Formato da planilha

`.csv` ou `.xlsx` com as colunas **`nome`** e **`email`** (cabecalho na 1a linha).
Veja `exemplo.csv`. A validacao: remove espacos, descarta e-mails duplicados e
linhas sem e-mail, saneia caracteres invalidos para nome de pasta no Windows e
trata homonimos (dois "Joao Silva" viram pastas distintas usando `Nome (email)`).

---

## 5. Como rodar

A senha comum (ja trocada por voce) e pedida no inicio, via `getpass` — nunca em
arquivo, nunca na planilha, nunca no log.

```powershell
# 1) Dry-run: so conecta, lista pastas (com acentos) e conta, sem baixar nada:
python main.py exemplo.csv --dry-run

# 2) Processar so uma conta:
python main.py exemplo.csv --conta joao.silva@suaempresa.com.br

# 3) Reprocessar uma conta ja concluida:
python main.py exemplo.csv --conta joao.silva@suaempresa.com.br --forcar

# 4) Lote completo (todas as contas da planilha):
python main.py exemplo.csv

# config.ini alternativo:
python main.py exemplo.csv --config D:\outro\config.ini
```

Quando empacotado como `.exe`, troque `python main.py` por `backup_tool.exe`.

### Saida gerada

```
<base>\
  logs\backup_AAAAMMDD_HHMMSS.log
  <Nome Completo>\
    <Nome Completo>.pst        <- final, apos verificacao OK
    _incompleto.pst            <- enquanto nao verificado / em caso de FALHA
    state.json                 <- status + UIDs baixados (retomada)
    meta.json                  <- flags/INTERNALDATE/Message-ID por UID
    _temp\                      <- .eml temporarios (apagados apos sucesso)
```

---

## 5b. Interface grafica (`backup_gui.exe`)

Para quem prefere o visual, use o `backup_gui.exe` (uma janela do Windows):

1. Dê **duplo clique** em `backup_gui.exe` (fica ao lado do `backup_tool.exe`).
2. **Procurar...** e escolha a planilha (`.csv`/`.xlsx`).
3. Digite a **senha comum** (campo mascarado; vai pelo stdin, nunca por arquivo/log).
4. Opcional: "Só esta conta", "Dry-run", "Forçar reprocesso".
5. Clique **Iniciar backup**. Acompanhe:
   - **barra "Conta"** (progresso da caixa atual, com a fase: Baixando / Montando PST / Verificando),
   - **barra "Geral"** (progresso somando todas as contas),
   - o **log** em tempo real.

Como funciona por dentro (e por que é seguro): a GUI **não** executa o MAPI — ela
lança o `backup_tool.exe` como processo separado (que roda o Redemption na própria
thread principal) e só lê o progresso. Assim a janela não congela nem trava o MAPI.
O botão **Parar** só fica ativo durante o **download** (interromper durante a
montagem/verificação do PST poderia travar o MAPI); nessa fase ele fica desabilitado.
`backup_gui.exe` precisa do `backup_tool.exe` e do `config.ini` na mesma pasta.

## 6. Empacotamento (PyInstaller, 64 bits)

**Jeito mais facil:** no PowerShell, dentro de `codigo-fonte\`, rode `.\build.ps1`
— ele gera os dois `.exe` direto em `..\programa\`.

Manualmente (equivalente), rodando a partir de `codigo-fonte\`:

```powershell
# CLI (console)
pyinstaller --onefile --name backup_tool --distpath ..\programa `
  --hidden-import win32timezone --hidden-import openpyxl --exclude-module tkinter `
  main.py

# Interface grafica (janela, sem console)
pyinstaller --onefile --windowed --name backup_gui --distpath ..\programa gui.py
```

- Rode com o **Python 64 bits** (gera `.exe` 64 bits).
- Distribua **junto** o `config.ini` e o `exemplo.csv` (ficam fora do `.exe`).
- O `.exe` **nao** embarca o Redemption nem o Outlook: a maquina de destino
  precisa ter **Outlook 2010** e o **Redemption64.dll registrado** (ver secao 1).

### Sobre o cache do `win32com` (gen_py) dentro do `.exe`

**Nao se aplica aqui.** O acesso ao Redemption e feito por *late binding*
(`win32com.client.Dispatch("Redemption.RDOSession")`), sem `EnsureDispatch`/
`makepy`, entao **nao ha cache gen_py** para falhar dentro do `--onefile`.
Se um dia voce precisar do Object Model do Outlook com *early binding*
(`EnsureDispatch`), ai sim resolva com uma das opcoes:
(a) redirecionar `win32com.__gen_path__` para uma pasta gravavel ao lado do
`.exe` **antes de qualquer chamada COM**, ou (b) gerar o wrapper com `makepy`
antes do build e inclui-lo nos dados do PyInstaller.

---

## 7. Troubleshooting

- **MAPI travado (Dispatch/`LogonPstStore` fica "pendurado" sem erro)**: acontece
  quando um processo que usava MAPI foi **encerrado a forca** (kill) no meio de
  uma operacao. O subsistema MAPI fica em estado ruim e as chamadas seguintes
  travam. **Solucao: reiniciar o Windows.** Por isso, evite matar a ferramenta no
  meio de um import; prefira `Ctrl+C` (tratado), e deixe o pipeline fechar a
  sessao sozinho.
- **`.pst` travado pelo Outlook apos fechar** (comportamento conhecido): a
  renomeacao final (`_incompleto.pst` -> `<Nome>.pst`) ja tenta varias vezes
  com espera (`renomear_com_retentativa`). Se ainda assim travar, feche o Outlook
  e rode de novo com `--forcar` naquela conta.
- **Erro ao carregar o COM / "class not registered"**: Redemption nao registrado
  ou arquitetura trocada (32 vs 64). Registre o `Redemption64.dll` e use Python
  64 bits.
- **Falha no logon MAPI**: o `LogonPstStore` cria um profile temporario proprio,
  mas o subsistema MAPI do Outlook precisa existir — confirme que o Outlook 2010
  esta instalado e ja foi aberto ao menos uma vez.
- **Licenca do Redemption**: a versao de desenvolvedor e gratuita; a
  **distribuicao** do `.exe` final para outras maquinas tem regras proprias —
  verifique antes de distribuir.

---

## 8. Seguranca e LGPD

- Senha **nunca** em arquivo nem em log — so `getpass` no inicio. Mantenha a
  janela em que a senha comum vale para varias contas **a mais curta possivel**.
- Download com `BODY.PEEK[]`: **nao** marca as mensagens como lidas no servidor.
- `.eml` temporarios ficam em pasta restrita e sao **apagados apos o sucesso**
  da conta. **Atencao a uma decisao de projeto** (ver item 9): em caso de FALHA
  ou interrupcao, os `.eml` sao **mantidos** para permitir a retomada — limpe-os
  manualmente ou rode `--forcar` (que limpa antes de recomecar) se nao for
  retomar.
- Log enxuto: conta, horario, etapa, contagens por pasta e Message-IDs faltantes
  — **nunca** assunto ou conteudo.
- O Outlook 2010 esta sem suporte desde 2020: a maquina deveria ficar isolada e
  dedicada a esse uso.

---

## 9. Desvios em relacao ao prompt (sinalizados, conforme pedido)

1. **Arquitetura 64 bits** (nao 32), seguindo sua correcao: Outlook 64 bits
   exige Python/`.exe`/Redemption 64 bits.
2. **`LogonPstStore`** (em vez de `AddStoreEx` do Object Model): e o equivalente
   direto no Redemption, cria o `.pst` como Unicode (`Format=2`, sem limite de
   2 GB) e **cria um profile MAPI temporario proprio**, dispensando um profile do
   Outlook pre-configurado — melhor para um `.exe` standalone.
3. **Download mensagem a mensagem por UID** (e nao um `FETCH` em lote unico): o
   parsing de literais do `imaplib` em respostas de FETCH multiplas e fragil, e o
   1-a-1 da a granularidade exata de retomada. O `lote` do config vira intervalo
   de checkpoint do `state.json`.
4. **Limpeza dos `.eml` so no sucesso.** O prompt pedia limpeza via `finally`
   sempre; isso conflitaria com a retomada (apagaria os `.eml` de um download
   interrompido). Resolvido a favor da retomada: apaga so quando a conta fica
   `CONCLUIDO`; em falha/interrupcao mantem (com a nota de LGPD do item 8).
5. **`util.py` adicionado** (alem dos modulos do prompt) para helpers de
   sanitizacao/caminho compartilhados entre modulos.
6. **Reconstrucao limpa do PST** se houver `_incompleto.pst` de uma tentativa
   anterior: como os `.eml` ja estao completos no disco, o PST e refeito do zero
   para nao duplicar mensagens (a retomada fina e no download).
7. **Sessao MAPI unica por processo** (`RedemptionSession`): um unico `Logon`
   (via `LogonPstStore` na 1a conta) e um unico `Logoff`; cada conta apenas monta
   (`AddPSTStore`) e desmonta (`store.Remove()`) o seu `.pst`. Isso foi necessario
   porque reinicializar o MAPI a cada conta (CoInitialize/Logon/Logoff/
   CoUninitialize em loop) **trava o subsistema MAPI** ao processar varias contas
   — exatamente o cenario de "backup em massa". Descoberto e corrigido durante os
   testes automatizados.

### Testes automatizados (`test_backup_tool.py`)

Rode: `<python64> test_backup_tool.py` (sem rede; precisa do Redemption registrado).

- **Parte A — logica pura (26 checagens, todas passando):** modified UTF-7,
  parsing de LIST/FLAGS/INTERNALDATE/Message-ID, sanitizacao/caminhos, validacao
  de planilha com homonimos/duplicados, `state_manager` (round-trip), `verifier`
  (sucesso e divergencia) e **retomada do download** (pula UIDs ja baixados).
- **Parte B — integracao com Redemption real (sessao unica):** download simulado
  (sem rede) -> cria o `.pst` -> verifica por pasta/Message-ID -> renomeia ->
  limpa; alem de dry-run, PULADA e `--forcar`. Roda sob **watchdog**: se o MAPI
  estiver travado, avisa e encerra em vez de ficar pendurado.

Tambem foi validado, em smoke test isolado, que o import preserva **data
(`ReceivedTime`), lido/nao-lido (`UnRead`) e remetente** (mensagem recebida, nao
rascunho).

**Falta testar com conta real** (precisa das suas credenciais): a conexao IMAP
ao vivo contra uma caixa do Skymail e o fluxo completo de ponta a ponta — o
Roteiro de teste da secao 10.

### Observacao sobre as pastas do `.pst`

O `LogonPstStore` cria automaticamente as pastas padrao do Outlook ("Caixa de
Entrada", "Itens Enviados", etc.), **vazias**. As pastas do IMAP sao recriadas
ao lado delas, com os **nomes originais do servidor** (ex.: `INBOX`, `Sent`).
A verificacao so considera as pastas que tiveram mensagens baixadas, entao as
pastas padrao vazias nao atrapalham — apenas aparecem no `.pst`.

---

## 10. Roteiro de teste (validacao manual)

1. `--dry-run` numa conta de teste: confirmar que conecta, lista pastas
   corretamente (com acentos) e conta mensagens sem baixar nada.
2. Rodar completo com **uma unica conta de teste** (nao uma conta real de
   ex-funcionario) e planilha com 1 linha.
3. Confirmar: IMAP baixa os `.eml` -> Redemption cria o `.pst` -> PST renomeado
   de `_incompleto` para o nome final -> verificacao por pasta e Message-ID sem
   faltantes -> log limpo, sem erros.
4. Abrir o `.pst` no Outlook e conferir visualmente datas, remetente e anexos
   (nao devem aparecer como rascunho).
5. Interromper no meio de um download (Ctrl+C) e rodar de novo: confirmar que
   retoma pelos UIDs do `state.json`, sem recomecar do zero.
6. Rodar a mesma conta ja concluida **sem** `--forcar` (deve pular) e **com**
   `--forcar` (deve reprocessar).
7. So depois disso, testar em lote com 2-3 contas de teste.
