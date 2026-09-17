"""
Como a fatura entra: Outlook, arrastar-e-soltar, e o palpite do que é cada PDF.

`outlook` varre a caixa por remetente e assunto; `manual` recebe o arquivo
quando a fatura só existe no portal do fornecedor. Os dois terminam no
`classificador`, que decide se o PDF é boleto, nota fiscal ou demonstrativo —
com nota de confiança e o motivo, porque palpite sem justificativa não dá para
conferir.

Duplicata é detectada por conteúdo, não por nome: o mesmo boleto chegando pelos
dois caminhos rendia um PDF final com páginas repetidas.
"""
