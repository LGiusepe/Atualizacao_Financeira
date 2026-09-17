"""
O que todo o resto usa: contratos, configuração e memória.

`modelos` é a fonte da verdade — `Competencia`, `Conta`, `Documento`, `Etapa`.
`config` lê os YAML e monta o `Ambiente`, que carrega as travas de gravação e
exclusão. `estado` é o SQLite: o que já foi feito, por conta e competência.

Este pacote não importa nada dos outros. É a direção em que as setas apontam:
`coleta`, `documentos`, `entrega` e `acesso` dependem do núcleo; o núcleo não
depende de ninguém. Um import daqui para lá seria o começo de um ciclo.
"""
