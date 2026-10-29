# Ferramenta de Backup em Massa de Caixas de Correio (IMAP → PST)

Ferramenta **standalone para Windows** que faz o backup das caixas de correio de
uma lista de contas, baixando tudo via **IMAP** e consolidando cada caixa em um
arquivo **`.pst`** local — pronto para abrir no Outlook. Pensada para o cenário de
**desligamento de funcionários**, onde o administrador de TI precisa arquivar a
caixa de e-mail antes de removê-la do servidor.

Tem **interface gráfica** (janela com barras de progresso) e também **linha de
comando** (para automação). Testada e validada contra um servidor **Skymail** real
(backup de ~7.800 mensagens, `.pst` de ~1,1 GB, íntegro).

> A troca da senha de cada conta e a exclusão da caixa no servidor são feitas
> **manualmente pelo administrador** — **não** fazem parte desta ferramenta. Ela
> só entra em ação depois que a senha já foi definida.

---

## O que faz

- Lê uma **planilha** (`.csv` ou `.xlsx`) com as colunas `nome, email`.
- Conecta em cada conta por **IMAP (SSL, porta 993)** usando uma senha comum
  (digitada na hora, nunca armazenada).
- **Baixa todas as pastas** preservando acentos, datas, remetente e anexos.
- Monta um **`.pst` por conta** (`<base>\<Nome>\<Nome>.pst`).
- **Verifica a integridade** (contagem por pasta + comparação por `Message-ID`)
  antes de dar o backup como concluído.
- É **retomável**: se cair ou for interrompido, continua de onde parou.

---

## Como funciona (arquitetura)

Pipeline por conta, nesta ordem:

1. **Login IMAP** (`IMAP4_SSL`, porta 993).
2. **Listagem de pastas** decodificando *modified UTF-7* (acentos) e lendo o
   **delimitador de hierarquia do próprio servidor** (não assume `/` nem `.`).
3. **Download por UID** com `BODY.PEEK[]` (não marca como lida no servidor),
   salvando cada mensagem como `.eml` (nomeado pelo UID) e guardando *flags*,
   `INTERNALDATE` e `Message-ID`. Reconecta sozinho e persiste o progresso.
4. **Montagem do `.pst`** via **Redemption (RDO)**: `LogonPstStore` cria o PST em
   formato **Unicode** (sem o limite de 2 GB do formato antigo) e cada `.eml` é
   importado com `RDOMail.Import(..., olRFC822)`, **preservando data de
   recebimento e status lido/não-lido**. As mensagens entram como **recebidas**
   (não como rascunho).
5. **Verificação**: contagem por pasta e diferença de `Message-ID` entre o que foi
   baixado e o que está no `.pst`. Qualquer falta é registrada no log.
6. Se passar: renomeia `_incompleto.pst` → `<Nome>.pst`. Se falhar: **mantém
   `_incompleto.pst`** e marca a conta como `FALHA` (nunca entrega um backup
   incompleto como se estivesse completo).
7. Limpa os `.eml` temporários (apenas no sucesso) e segue para a próxima conta.

### Decisões técnicas importantes

- **Sessão MAPI única por processo** — um único `Logon`/`Logoff`; cada conta
  apenas monta (`AddPSTStore`) e desmonta (`store.Remove`) o seu `.pst`.
  Reinicializar o MAPI a cada conta trava o subsistema; essa arquitetura evita
  isso ao processar várias caixas em lote.
- **Verificação por `Message-ID`** — garante que o `.pst` tem exatamente as
  mensagens baixadas; é a trava que impede um backup furado silencioso.
- **Retomada (resume)** — o estado (`state.json`) guarda os UIDs já baixados por
  pasta; reexecutar continua sem recomeçar do zero.
- **Auto-cura** — se um `.eml` já baixado sumir do disco (ex.: removido por
  antivírus entre o download e a montagem), a ferramenta o **re-baixa**
  automaticamente.
- **Interface sem travar** — a janela gráfica **não executa o MAPI**: ela lança o
  `backup_tool.exe` como processo separado (que roda o Redemption na própria
  thread principal) e apenas lê o progresso. Assim a janela nunca congela.

---

## O que usa

| Camada | Tecnologia |
|---|---|
| Linguagem | **Python 3.x (64 bits)** |
| Download | **`imaplib`** (IMAP RFC 3501, biblioteca padrão) |
| Criação do `.pst` | **Redemption (RDO)** via **pywin32** (COM) |
| Planilhas `.xlsx` | **openpyxl** (CSV usa a biblioteca padrão) |
| Interface gráfica | **Tkinter** (biblioteca padrão) |
| Empacotamento | **PyInstaller** (gera os `.exe` 64 bits) |

**Requisitos da máquina de destino:** Windows com **Outlook 2010 (64 bits)**
instalado (fornece o subsistema MAPI) e o **Redemption** registrado
(`Redemption64.dll`). Arquitetura **64 bits** em tudo (Python, `.exe`, Redemption)
para casar com o Outlook.

---

## Estrutura do repositório

```
backup_tool/
├─ LEIA-ME.txt          # ponto de partida
├─ programa/            # PARA USAR (pronto para rodar)
│    backup_gui.exe       interface gráfica (duplo clique)
│    backup_tool.exe      versão linha de comando
│    config.ini           configuração (servidor, pasta de saída)
│    exemplo.csv          planilha modelo
└─ codigo-fonte/        # CÓDIGO e DOCUMENTAÇÃO
     *.py                 código-fonte Python
     README.md            instruções detalhadas (uso, build, troubleshooting)
     requirements.txt
     build.ps1            regera os .exe em ..\programa
```

> Documentação detalhada (passo a passo, flags da CLI, empacotamento,
> troubleshooting e roteiro de teste): **[`codigo-fonte/README.md`](codigo-fonte/README.md)**.

---

## Como usar (resumo)

**Interface gráfica:** duplo clique em `programa\backup_gui.exe` → escolha a
planilha, digite a senha comum, clique em **Iniciar backup** e acompanhe as barras
(por conta e geral).

**Linha de comando:**
```powershell
# Teste sem baixar nada (só conecta, lista pastas e conta):
programa\backup_tool.exe planilha.csv --dry-run

# Backup completo:
programa\backup_tool.exe planilha.csv
```
Flags: `--dry-run`, `--conta email@...`, `--forcar` (reprocessa uma conta já
concluída). A senha é pedida na hora (nunca vai para arquivo nem para o log).

---

## Segurança e LGPD

São caixas com dados pessoais de terceiros — a ferramenta foi feita com isso em mente:

- **Senha** nunca em arquivo nem em log — digitada na execução (via `getpass` na
  CLI ou campo mascarado na GUI).
- Download com `BODY.PEEK[]`: **não** marca mensagens como lidas no servidor.
- **Log enxuto**: contas, etapas, contagens e `Message-ID` faltantes — **nunca**
  assunto ou conteúdo das mensagens.
- `.eml` temporários em pasta restrita, apagados após o sucesso.
- Pasta de saída **configurável** (não usar a Área de Trabalho / OneDrive).
- Recomenda-se **excluir a pasta de backup no antivírus** (ele pode remover
  e-mails de phishing/malware do disco e quebrar a integridade do `.pst`).

---

## Status

Validado de ponta a ponta contra um servidor **Skymail** real: download completo,
montagem do `.pst`, verificação por pasta/`Message-ID`, retomada e auto-cura — tudo
funcionando. Inclui uma suíte de testes automatizados
(`codigo-fonte/test_backup_tool.py`) cobrindo a lógica pura e a integração com o
Redemption.

## Build (recompilar os `.exe`)

No PowerShell, dentro de `codigo-fonte/`:
```powershell
pip install -r requirements.txt
.\build.ps1   # gera backup_tool.exe e backup_gui.exe em ..\programa
```
