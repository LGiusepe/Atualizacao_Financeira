"""
O que a automação produz: a autorização preenchida e os PDFs.

`autorizacao` copia o modelo do mês anterior e reescreve só os campos de valor,
preservando as fórmulas — é o arquivo mais delicado do projeto, porque o
openpyxl perde partes do xlsx ao salvar e elas precisam ser devolvidas por
cirurgia no zip. `exportar_pdf` manda o Excel exportar a aba respeitando a área
de impressão. `montar_pdf` junta autorização, demonstrativo, boleto e nota na
ordem configurada.

Nada aqui escreve fora da bancada. Publicar é com `entrega`.
"""
