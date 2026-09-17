# Automação Financeira — Autorizações de Pagamento

Automatiza o fechamento mensal das contas de T.I.: recolhe a fatura, preenche a
autorização de pagamento, monta o PDF, prepara o e-mail para o financeiro e
marca a conta no checklist do time.

Roda **localmente**, em Python, com painel web. Sem serviço em nuvem, sem banco
de dados externo, sem assinatura: só software livre e o Office que a empresa já
licencia. Custo recorrente: R$ 0,00.

## O problema

Todo mês, para cada conta ativa — telefonia, links, licenças de software —,
alguém precisa:

1. achar a fatura, no e-mail ou no portal do fornecedor;
2. criar a pasta do mês na rede;
3. copiar a autorização do mês anterior e reescrever valor, vencimento e nº do documento;
4. exportar a aba em PDF;
5. juntar autorização + demonstrativo + boleto + nota fiscal num arquivo só, nessa ordem;
6. mandar por e-mail ao financeiro, com cópia para a equipe;
7. marcar a conta na planilha de controle.

São sete passos manuais por conta. O registro tem 54 contas cadastradas e cerca
de 26 ficam ativas num mês típico. O trabalho é repetitivo, mas não é mecânico:
cada operadora nomeia arquivo de um jeito, o valor está em lugar diferente em
cada boleto, e errar significa pagar a conta errada.

## Como funciona

Um pipeline de oito etapas, uma por conta, com o julgamento humano no meio:

```
coleta → pasta → autorização → PDF da autorização → PDF único
       → publicação → e-mail → checklist
```

Cada etapa só abre quando a anterior fecha, e refazer uma reabre o que dependia
dela: corrigir a planilha na etapa 3 invalida os PDFs das etapas 4 e 5, para
que o arquivo velho não seja publicado depois da correção.

**Publicação** e **checklist** escrevem em arquivo de outra pessoa e só rodam
com **confirmação explícita**. Antes de gravar, a tela mostra a lista completa
do que será feito, arquivo por arquivo, com aviso quando há sobrescrita.

### O que o painel faz por você

| Etapa | O que acontece |
|---|---|
| **Coletar** | Varre o Outlook por remetente/assunto, ou recebe arquivos por arrastar-e-soltar. Classifica cada PDF como boleto, nota fiscal ou demonstrativo, com nota de confiança e o motivo da decisão. Duplicata é detectada por conteúdo, não por nome. |
| **Autorização** | Copia o modelo do mês anterior e reescreve **só** os campos de valor, preservando todas as fórmulas (`VLOOKUP`, `TODAY`, `SUM`). Valor, vencimento e nº do documento vêm dos documentos lidos — o boleto tem preferência —, e todos são editáveis. |
| **PDF** | Exporta a aba pelo próprio Excel, respeitando a área de impressão, e junta os documentos na ordem configurada. O arquivo aparece na tela para conferência. |
| **Publicação** | Cria a pasta do mês e copia os arquivos para a rede. Cada conta arquiva do seu jeito, e o registro tem campo para isso — nome de pasta próprio, mês deslocado, raiz fora da árvore padrão. |
| **E-mail** | Monta o rascunho no Outlook com o PDF anexado. Destinatário, cópia, assunto e texto são editáveis na hora. Contas irmãs do mesmo fornecedor viram uma mensagem só, com um anexo por fatura. |
| **Checklist** | Marca a caixa da conta na aba do mês da planilha de controle. A planilha tem formatação condicional: marcada a caixa, a linha fica verde sozinha. |

### O mês de relance

A tela inicial abre com o fechamento: quanto já foi autorizado e quanto ainda
falta, com a contagem de contas de cada lado. Dois avisos viajam junto do
total, senão ele parece mais exato do que é — quanto da soma é previsto em vez
de boleto lido, e quantas contas não entraram em soma nenhuma por não ter valor
em lugar algum.

A ferramenta acompanha a **autorização**, não o extrato do banco: "pago" aqui
quer dizer "autorização despachada ao financeiro".

O painel acompanha o tema do sistema e tem chave para claro/escuro no topo.

### Quem entra

O painel pede login. São duas coisas separadas, em arquivos separados:

- **Usuários** (`config/usuarios.yaml`) — quem entra. Senha guardada em
  `scrypt`, com sal por usuário; o arquivo tem o resumo, nunca a senha, e não
  há como voltar atrás a partir dele. Esquecendo, um administrador redefine.
- **Grupos** (`config/permissoes.yaml`) — até onde cada um vai. Vêm três
  prontos: *Administrador* (tudo), *Operador* (o mês, a entrada manual, os
  cadastros) e *Consulta* (só leitura).

A tela que o grupo não enxerga **some do menu**, em vez de aparecer e devolver
"sem permissão" — oferecer um link que não funciona faz a pessoa achar que o
painel quebrou. Mas a trava é no servidor: esconder o link não protege nada, já
que o endereço continua digitável. O administrador não é editável; se desse
para tirar telas dele, um clique errado trancaria todo mundo para fora do
próprio cadastro de usuários.

Tela nova nasce fechada — um recurso que nenhum grupo lista aparece só para o
administrador, que decide liberá-lo.

## Travas de segurança

O projeto lida com pagamento. As garantias são estruturais, não intenções:

1. **Modo simulação** — nada é gravado fora das pastas do painel. As etapas
   sensíveis mostram o que fariam e param aí. É o padrão de quem instala.

2. **Planejar antes de gravar** — toda escrita externa é precedida da lista
   completa do que será feito, com aviso de sobrescrita.

3. **Backup e auditoria** — cópia do arquivo antes de sobrescrever, e registro
   de cada operação com hash SHA-256 em banco local.

4. **Exclusão proibida em produção** — a pasta de autorizações, a planilha de
   controle e qualquer raiz declarada no registro são intocáveis para exclusão,
   **em qualquer configuração**. Desligar a simulação libera gravação; jamais
   exclusão.

   A ordem das perguntas importa: pergunta-se *primeiro* se o alvo é produção
   (proibido, sem exceção) e só depois se é pasta do painel (liberado).
   Invertendo, um caminho de trabalho digitado errado — apontando para dentro
   da produção — abriria a porta. Há um script que prova isso em 24
   conferências, inclusive a de que nenhum arquivo sumiu desde a última
   execução e a de que a limpeza da bancada recusa um plano forjado contra
   a produção:

   ```bash
   python scripts/provar_nao_apaga.py
   ```

5. **O e-mail fica em Rascunhos** — o padrão é montar a mensagem e parar. Quem
   clica em enviar é você.

   Existe `email.enviar_automaticamente` para quem quiser o despacho direto,
   com três ressalvas honestas: a simulação sempre vence; há trava contra
   segunda via da mesma autorização; e o despacho **só funciona com o Outlook
   clássico** — o COM não conversa com o Outlook novo, e nesse arranjo `Send()`
   apenas empurra a mensagem para a Caixa de Saída, onde ela fica parada. O
   painel confere a Caixa de Saída depois de enviar e termina em erro se a
   mensagem continuar lá, em vez de dizer "enviado" para algo que não saiu.

6. **A limpeza só apaga o que está provado no destino** — a bancada
   (`trabalho/`, `entrada/`, `backups/`) é recolhida sozinha, mas um arquivo
   só sai quando o **mesmo conteúdo** é encontrado na pasta de destino, por
   SHA-256, byte a byte. Casar por nome não valeria: a cópia publicada pode
   ter ganhado sufixo no caminho. O que não passa nessa prova fica, e aparece
   listado com o motivo em *Configuração › Manutenção*.

## Instalação

Requer **Python 3.11+**, **Windows** e **Microsoft Office** instalados (Excel e
Outlook são acionados via COM).

```bash
git clone https://github.com/LGiusepe/Atualizacao_Financeira_FolhaDeRosto.git
cd Atualizacao_Financeira_FolhaDeRosto
python -m pip install -r requirements.txt
```

Copie os arquivos de exemplo e preencha com a sua realidade:

```bash
copy config\settings.exemplo.yaml config\settings.yaml
copy config\pagantes.exemplo.yaml config\pagantes.yaml
copy config\beneficiarios.exemplo.yaml config\beneficiarios.yaml
```

O `settings.yaml` é o principal: caminhos da rede, destinatários do e-mail, as
travas e quem será o primeiro administrador. Comece com
`seguranca.simulacao: true`.

Falta o registro das contas (`config/fornecedores.yaml`), gerado a partir da
estrutura de pastas que você já tem:

```bash
python scripts/gerar_registro_fornecedores.py
python scripts/enriquecer_registro.py
copy config\fornecedores.enriquecido.yaml config\fornecedores.yaml
```

`usuarios.yaml` e `permissoes.yaml` não precisam ser criados à mão: o primeiro
administrador nasce na primeira subida, e os grupos valem os padrões enquanto o
arquivo não existir.

A logomarca também não vem no repositório — `python scripts/extrair_logo.py`
gera claro, escuro e favicon a partir do próprio modelo de autorização. Sem
eles o cabeçalho mostra só o nome em texto.

## Uso

```bash
python -m painel
```

Ou duplo clique em `painel.bat`.

O painel prefere a **porta 80**, para o endereço ficar sem `:porta` no fim.
Ocupada — no Windows o `http.sys` costuma reservá-la para o IIS —, ele cai
sozinho na 8000 e escreve em `logs/porta.txt` onde subiu.

Endereço: `http://127.0.0.1`, sempre. Para o nome amigável
`http://organizacao.financeira.local`, rode uma vez, **como administrador**:

```bash
python scripts/configurar_endereco.py --aplicar
```

Ele acrescenta uma linha ao arquivo `hosts`. Vale só nesta máquina; o painel
continua inalcançável de fora, porque escuta em `127.0.0.1` e não na rede.

**Primeiro acesso:** o usuário é o e-mail configurado em
`acesso.primeiro_administrador`, e a senha do ano corrente aparece no console
na primeira subida. Ela é previsível de propósito — serve para entrar uma vez.
Nenhuma outra tela abre antes de você escolher uma senha sua.

Para encerrar: **Ctrl+C** na janela, ou `parar-painel.bat` se ela já foi
fechada e o processo ficou preso.

## Publicar no GitHub

O repositório leva **código e exemplos**; dado real da empresa fica só nesta
máquina. A separação tem três camadas, e nenhuma substitui a outra.

**1. O `.gitignore`** barra o que tem conteúdo real:

| Fora do git | Por quê |
|---|---|
| `config/*.yaml` (menos `*.exemplo.yaml`) | CNPJ, dados bancários, caminhos da rede, contratos |
| `config/usuarios.yaml`, `dados/sessao.chave` | cadastro de pessoas e a chave que assina a sessão |
| `dados/` | faturas, boletos, notas, banco de estado, cofre de senhas |
| `logs/`, `docs/`, `BACKLOG.md` | notas internas: citam fornecedor, valor e nome de pessoa |
| `painel/static/logo*.png`, `favicon.ico` | arte da empresa |
| `.claude/` | configuração local do assistente |

**2. O conferidor**, que lê o que o git *de fato* levaria (`git ls-files`, já
com o `.gitignore` aplicado) e procura CNPJ, e-mail, telefone, conta bancária,
caminho de máquina, nome da empresa e nome de pessoa:

```bash
python scripts/conferir_publicacao.py
```

Sai com código 1 se achar algo. Ele distingue exemplo de dado real — CNPJ cujo
dígito verificador não fecha é número inventado, `@exemplo.com.br` é domínio de
teste —, porque um conferidor que acusa os próprios marcadores acostuma a
pessoa a ignorar o resultado.

No fim ele lista os nomes de fornecedor citados em comentário. Revelam com quem
a empresa contrata, mas não são dado pessoal nem bancário: em repositório
privado, é decisão sua.

**3. A prova da trava de exclusão**, que não é sobre vazamento, mas roda no
mesmo momento — é quando se olha para o projeto inteiro.

O ritual, na ordem:

```bash
python -m compileall -q automacao painel scripts
python scripts/provar_nao_apaga.py
python scripts/conferir_publicacao.py
git status
git add -A
git commit -m "mensagem"
git push
```

Não há hook de pre-commit instalado, de propósito: hook que falha no meio de um
commit tende a ser contornado com `--no-verify`, e aí a conferência vira
formalidade. Aqui ela é um passo consciente.

## Arquitetura

As três primeiras pastas de `automacao/` são o próprio fluxo, na ordem em que
ele acontece:

```
automacao/
  nucleo/       modelos, config, estado — o que todo o resto usa
  coleta/       outlook, manual, classificador — de onde a fatura vem
  documentos/   autorizacao, exportar_pdf, montar_pdf — o que se produz
  entrega/      publicador, email_outlook, planilha_contas — o mundo de fora
  acesso/       usuarios, permissoes, cofre, senhas_pdf — credenciais e papéis
  manutencao/   limpeza — o único módulo que apaga arquivo neste projeto
  orquestrador.py

painel/     app.py (FastAPI) + templates Jinja2
scripts/    geração do registro, diagnóstico, prova de não-exclusão,
            conferência de publicação, endereço amigável, manual em PDF,
            início junto com o Windows
config/     settings.yaml (caminhos, e-mail, travas, primeiro administrador)
            fornecedores.yaml (o registro das contas)
            pagantes.yaml / beneficiarios.yaml (quem paga, quem recebe)
            usuarios.yaml / permissoes.yaml (quem entra, e até onde)
dados/      estado.db e cofre.dat — ficam locais de propósito: banco SQLite
            em pasta sincronizada corrompe, e credencial não se sincroniza
```

O `orquestrador` fica na raiz de `automacao/` porque não pertence a nenhuma
etapa: ele é quem as chama, na ordem, e decide o que já pode rodar. As setas de
dependência apontam para dentro — `nucleo` não importa nada dos outros pacotes;
os outros importam dele. Cada pacote tem um `__init__.py` explicando o que cabe
ali, para um módulo novo não nascer no lugar errado.

A bancada de trabalho (`trabalho/`, `entrada/`, `backups/`) é configurável e
pode morar em pasta sincronizada, para não acumular no disco da máquina — e
`manutencao/` a recolhe sozinha, para ela não acumular também na nuvem.

Cerca de 17.000 linhas de Python. Os módulos pesados (Excel, Outlook, PDF) são
importados dentro das funções, não no topo: é o que deixa o painel abrir numa
máquina sem Office, para conferir configuração.

### Decisões que valem o comentário

**Peculiaridade de fornecedor vai no registro, não no código.** Cada conta
batizou as pastas do seu jeito, e o YAML tem campo para cada caso:
`formato_mes`, `padrao_nome_pasta` e `padrao_nome_arquivo` para quem nomeia
diferente, `deslocamento_pasta` para quem arquiva o mês seguinte,
`raiz_onedrive` para quem mora fora da árvore padrão, `apelido` para o nome de
tela (que não move pasta nenhuma). Um `if` por fornecedor no código seria a
primeira de cinquenta e quatro exceções.

**O valor previsto do mês é o valor pago no mês anterior.** Um número fixo no
cadastro envelhece sem ninguém perceber — ninguém revisa 54 contas uma a uma —,
enquanto o último valor pago acompanha reajuste, linha a mais e licença a menos
sozinho. A ordem do palpite é: valor lido do boleto, depois o último pago,
depois o do cadastro (que fica para conta nova, sem mês nenhum atrás dela). Só
entra no histórico valor gravado de verdade: se um palpite virasse histórico no
mês seguinte, ele se propagaria adiante como se fosse fato. A tela mostra de
que mês o número veio, porque "previsto" sem procedência é o tipo de coisa que
se confere uma vez e nunca mais.

**A planilha é preenchida sem `data_only`.** Abrir com openpyxl preservando as
fórmulas é o que mantém os `VLOOKUP` de CNPJ e dados bancários vivos no arquivo
gerado — ele vira o modelo do mês seguinte. Só as células de valor são
reescritas; célula com fórmula é recusada, salvo nos campos explicitamente
marcados como valor.

**Quando o `VLOOKUP` não tem resposta, o painel avisa.** A tabela do modelo
guarda razões sociais antigas, e empresa renomeada saía como `#N/D`. O valor do
cadastro é escrito no lugar da fórmula — mas com aviso na tela, porque trocar
cálculo por texto fixo precisa ficar à vista. Antes de escrever, a ordem das
colunas é validada: sem isso, escreveríamos telefone no campo do CNPJ, que é
pior que o erro visível.

**O openpyxl perde coisas ao salvar, e algumas importam.** Ele descarta células
vazias que só carregavam formatação — elas herdam o estilo da coluna e o PDF
sai com tarjas cinza — e reescreve o XML dos desenhos. Pior: ao descartar um
desenho que não sabe reescrever, ele **renumera** os outros, e o nome do
arquivo deixa de significar a mesma coisa nos dois pacotes; casar por nome
rendeu uma logomarca de seis polegadas cobrindo o formulário. O módulo
`documentos/autorizacao.py` devolve essas partes por cirurgia no zip, casando
por aba.

**A escrita na planilha compartilhada vai por Excel COM**, não openpyxl: o
arquivo tem partes `customXml` do SharePoint que o openpyxl descarta, e
perdê-las desconfigura o documento para o time inteiro.

**Toda chamada COM é posicional.** O pywin32 em *late binding* descarta
argumento nomeado em silêncio — `Copy(After=x)` vira `Copy()`, e a aba vai
parar em outro arquivo. E `Quit()` apenas *pede* o encerramento: enquanto
sobrar uma referência COM, o processo fica vivo segurando a planilha do time. O
painel solta as referências, coleta o lixo e, se ainda assim o Excel que ele
abriu sobreviver, encerra pelo PID — só o que ele mesmo criou, nunca um Excel
do usuário, que pode ter trabalho não salvo.

**Duplicata é detectada por conteúdo**, não por nome de arquivo: o mesmo boleto
chegando por dois caminhos rendia um PDF final com páginas repetidas.

## Estado do projeto

Em uso real desde agosto/2026, com gravação e publicação ligadas. O envio
automático de e-mail continua desligado pelo motivo descrito nas travas.

Quem instalar do zero começa em `simulacao: true` — é o caminho certo para
conferir os caminhos da rede antes de a automação tocar em qualquer coisa.

## Licença

MIT.
