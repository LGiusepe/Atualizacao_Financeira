"""
Credenciais e permissões — as duas coisas que o projeto guarda sob chave.

`usuarios` diz *se* a pessoa entra no painel (senha em scrypt, com sal por
usuário); `permissoes` diz *até onde* ela vai (grupos e telas). São arquivos
separados de propósito: um cadastra gente, o outro cadastra papéis.

`cofre` guarda as senhas de portal e de PDF cifradas pelo DPAPI do Windows, e
`senhas_pdf` as usa para abrir boleto protegido na hora de montar o PDF. O
cofre fica em `dados/`, que não sincroniza: credencial não sobe para a nuvem.
"""
