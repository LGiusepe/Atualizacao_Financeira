"""
Manutenção da bancada — o que sai de "Painel Temp" sem que nada se perca.

É o único pacote deste projeto que **apaga** arquivo, e por isso vale a regra
mais apertada: nenhum caminho é excluído sem passar por
`Ambiente.exigir_permissao_para_apagar`, que recusa categoricamente qualquer
coisa em "Autorizações de pagamento" — em qualquer configuração, com ou sem
`seguranca.simulacao`.

O critério do que cabe aqui: cuidar do *espaço de trabalho*, nunca do arquivo
do processo. Publicar, marcar checklist, mandar e-mail é `entrega/`; o que a
automação produz é `documentos/`. Aqui só entra o que **recolhe** o que já
cumpriu o seu papel.

`limpeza.py` é o módulo. Ele trabalha em duas fases, como o publicador:
`planejar()` diz exatamente o que sairia e não toca em nada; `executar()` só
apaga com `confirmado=True`.
"""
