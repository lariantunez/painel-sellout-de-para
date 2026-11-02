# Painel Atlas Mongo

Projeto criado a partir do painel Vitta para atender o cliente Atlas.

## Situação atual

- O arquivo bruto `Dados Brutos Junho.csv` já está na pasta do projeto.
- A estrutura do CSV bate com o esperado pelo painel Bluesoft: `Data`, `Loja`, `GTIN/PLU`, `Produto`, `Venda (Qtd)`, `Venda (R$)` e `Estoque Diario`.
- O De/Para de categorias e produtos será reaproveitado do projeto anterior.
- Como o Atlas tem GTIN cadastrado diferente, o projeto agora tem um De/Para intermediário: `GTIN_ATLAS` ou `COD_INTERNO_ATLAS` -> `CODBARRAS` correto.
- O De/Para de lojas também precisa ser criado/atualizado para este cliente.

## Próximos passos

1. Criar um banco Mongo separado, usando `DB_NAME=atlas_dashboard`.
2. Copiar `.env.example` para `.env` e preencher `MONGODB_URI`, usuário admin e credenciais de e-mail se for usar reset de senha.
3. Importar o mesmo `categorias_depara` usado no painel anterior.
4. Criar e importar `gtin_atlas_depara` com as colunas `GTIN_ATLAS`, `COD_INTERNO_ATLAS` e `CODBARRAS`.
5. Criar e importar `lojas_depara` com as colunas `Cod_Loja` e `Nome_Fantasia`.
6. Importar `Dados Brutos Junho.csv` pela aba Importar do painel.
7. Subir o projeto para Git e depois criar o deploy no Koyeb.

## De/Para GTIN Atlas

Use `GTIN_ATLAS` quando a equivalência puder ser feita pelo GTIN errado que veio do Bluesoft do Atlas.
Use `COD_INTERNO_ATLAS` quando o código interno do produto no Atlas for a referência mais confiável.
O campo `CODBARRAS` deve conter o código correto que existe no De/Para compartilhado de categorias/produtos.
