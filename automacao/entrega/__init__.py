"""
Onde a automação toca o mundo de fora — e onde ela pede licença.

`publicador` copia para o OneDrive, `email_outlook` monta a mensagem para o
financeiro, `planilha_contas` marca a conta no checklist do time. São arquivos
de outras pessoas: publicação e checklist só rodam com confirmação explícita, e
nenhum dos três consegue apagar nada em produção (ver `Ambiente` no núcleo).

Toda chamada COM daqui é posicional. O pywin32 em *late binding* descarta
argumento nomeado em silêncio, e `Copy(After=x)` vira `Copy()` — que manda a
aba para outro arquivo.
"""
