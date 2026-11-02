# Manual do Programador - Painel Atlas

Este manual explica a estrutura tecnica do projeto em linguagem simples.

A ideia e ajudar uma pessoa desenvolvedora a entender o projeto sem precisar descobrir tudo do zero.

## 1. Visao Geral

O projeto e um painel web com:

- frontend em HTML, CSS e JavaScript puro;
- backend em Node.js com Express;
- banco de dados MongoDB;
- importacao de arquivos CSV e XLSX;
- autenticacao de usuarios;
- dashboard de vendas, valor e estoque.

Estrutura principal:

```text
frontend/index.html
backend/src/server.js
backend/.env
docs/
```

O arquivo `frontend/index.html` contem praticamente toda a interface do painel.

O arquivo `backend/src/server.js` contem praticamente toda a API, regras de importacao, consultas ao Mongo e autenticacao.

## 2. Como Frontend, Backend E Mongo Conversam

O fluxo geral e:

```text
Navegador
↓
frontend/index.html
↓ chamadas HTTP
backend/src/server.js
↓ consultas e gravacoes
MongoDB
```

O frontend nunca acessa o Mongo diretamente.

Ele sempre chama o backend.

Exemplo:

```text
Frontend chama /api/dashboard/agregados
↓
Backend consulta dados_brutos
↓
Backend calcula totais
↓
Frontend recebe JSON
↓
Frontend renderiza graficos e tabelas
```

## 3. Principais Colecoes Do MongoDB

### dados_brutos

Colecao principal do painel.

Guarda os dados importados do CSV de vendas/estoque.

Cada linha do arquivo vira um documento.

Campos comuns:

```text
Ano
Mes ou Mês
Data
Loja
GTIN/PLU
Produto
Venda (Qtd)
Venda (R$)
Estoque Diario
```

Campos preparados:

```text
_import_id
_qtd_num
_valor_num
_gtin
_data_iso
_cat
_fam
```

Esses campos preparados existem para acelerar as consultas.

### categorias_depara

Guarda o de/para usado pelo dashboard para categoria, familia e nome do produto.

Campos esperados:

```text
CODBARRAS
CATEGORIA
FAMILIA
NOME PRODUTO
```

Essa colecao ajuda a descobrir categoria, familia e nome do produto a partir do GTIN/PLU.

No projeto Atlas, essa colecao pode ser alimentada automaticamente a partir do arquivo unico `gtin_atlas_depara`.

### gtin_atlas_depara

Guarda o de/para especifico do Atlas para corrigir o produto antes das consultas.

Campos esperados:

```text
GTIN_ATLAS
COD_INTERNO_ATLAS
PRODUTO_ATLAS
CODBARRAS_AURORA
CATEGORIA_AURORA
NOME PRODUTO_AURORA
FAMILIA_AURORA
COD SFA_AURORA
```

Essa colecao liga o codigo vindo do Bluesoft Atlas ao codigo correto da Aurora.

### lojas_depara

Guarda o de/para de lojas.

Campos comuns:

```text
Cod_Loja
Nome_Fantasia
```

Essa colecao transforma codigo da loja em nome da loja.

### logs_importacao

Guarda o historico das importacoes.

Campos importantes:

```text
importId
tipo
arquivo
usuario
total
data
```

O `importId` conecta o log aos registros importados.

### usuarios_admin

Guarda administradores do painel.

Um administrador pode ter:

```text
usuario
senha
email
usuarioPai
criadoEm
```

Quando `usuarioPai: true`, esse usuario tem autonomia total.

### usuarios_importacao

Guarda usuarios que acessam a parte de importacao.

### tokens_reset

Guarda tokens temporarios de redefinicao de senha por e-mail.

### templates_importacao

Guarda modelos de arquivos de importacao.

## 4. Backend

Arquivo principal:

```text
backend/src/server.js
```

Responsabilidades principais:

- iniciar servidor Express;
- conectar ao MongoDB;
- servir o frontend;
- autenticar usuarios;
- receber arquivos;
- processar CSV/XLSX;
- salvar dados no Mongo;
- gerar dashboards;
- limpar cache;
- gerenciar usuarios;
- desfazer importacoes.

## 5. Rotas Importantes Do Backend

### Dashboard

```text
GET /api/dashboard/kpis
GET /api/dashboard/agregados
GET /api/dashboard/estoque
GET /api/dashboard/resumo
GET /api/dashboard/categorias
GET /api/dashboard/familias
```

Essas rotas alimentam as abas do painel.

### Importacao

```text
POST /api/importar/dados-brutos
POST /api/importar/categorias-depara
POST /api/importar/gtin-atlas-depara
POST /api/importar/lojas-depara
```

Essas rotas recebem os arquivos enviados pelo frontend.

### Gestao

```text
GET    /api/admin/usuarios
POST   /api/admin/usuarios
PUT    /api/admin/usuarios/:id/usuario
PUT    /api/admin/usuarios/:id/email
PUT    /api/admin/usuarios/:id/senha
DELETE /api/admin/usuarios/:id
```

Para administradores:

```text
GET    /api/admin/admins
POST   /api/admin/admins
PUT    /api/admin/admins/:id/usuario
PUT    /api/admin/admins/:id/email
PUT    /api/admin/admins/:id/senha
DELETE /api/admin/admins/:id
```

### Historico De Importacoes

```text
GET    /api/admin/logs-importacao
DELETE /api/admin/logs-importacao
DELETE /api/admin/logs-importacao/:id
```

## 6. Frontend

Arquivo principal:

```text
frontend/index.html
```

Ele contem:

- HTML das abas;
- CSS visual do painel;
- JavaScript de filtros;
- JavaScript de graficos;
- JavaScript de importacao;
- JavaScript de gestao de usuarios.

As abas principais sao:

```text
tab-vendas
tab-valor
tab-dia
tab-estoque
tab-se
tab-importar
tab-gerir
```

O frontend chama o backend usando `fetch`.

Exemplo:

```js
fetchJSON(`${API_BASE}/api/dashboard/agregados?...`)
```

## 7. Importacao De Dados Brutos

Quando o usuario importa dados brutos:

1. O frontend pega o arquivo.
2. O arquivo pode ser enviado em lotes.
3. Cada lote chega no backend.
4. O backend le cada linha.
5. O backend limpa e normaliza os dados.
6. O backend cria campos preparados.
7. O backend salva no Mongo.
8. No ultimo lote, cria um log em `logs_importacao`.
9. O cache do dashboard e limpo.
10. O backend aquece algumas consultas principais.

## 8. Campos Preparados

Os campos preparados sao campos extras gravados em `dados_brutos`.

Eles evitam que o painel precise recalcular tudo toda vez.

### _qtd_num

Quantidade vendida convertida para numero.

Usado para somar vendas em quantidade.

### _valor_num

Valor vendido convertido para numero.

Usado para somar venda em reais.

### _gtin

Codigo do produto padronizado.

Usado para cruzar com `gtin_atlas_depara` e `categorias_depara`.

### _data_iso

Data em formato padrao:

```text
YYYY-MM-DD
```

Facilita filtros por data.

### _cat

Categoria do produto.

Vem do De/Para Produtos Atlas/Aurora ou do de/para de categorias sincronizado.

### _fam

Familia do produto.

Vem do De/Para Produtos Atlas/Aurora ou do de/para de categorias sincronizado.

### _import_id

Identificador da importacao.

Permite desfazer uma importacao especifica.

## 9. Cache

O backend usa cache em memoria para acelerar respostas.

Funcoes principais:

```js
cacheGet()
cacheSet()
cacheClear()
```

As rotas de dashboard consultam o cache antes de recalcular.

Quando dados sao importados ou apagados, o cache deve ser limpo.

Isso evita que o painel mostre informacoes antigas.

## 10. Pre-Aquecimento De Cache

O backend tambem pode aquecer cache.

Isso significa chamar algumas rotas automaticamente, antes de o usuario precisar delas.

Rotas aquecidas:

```text
/api/dashboard/kpis
/api/dashboard/agregados
/api/dashboard/estoque
```

Isso reduz a chance de o usuario abrir uma aba e esperar muito.

## 11. De/Para Durante A Importacao

Na importacao de dados brutos, o backend carrega `gtin_atlas_depara` e `categorias_depara` em memoria como mapas.

Exemplo:

```text
GTIN_ATLAS ou COD_INTERNO_ATLAS -> CODBARRAS_AURORA -> categoria/familia
```

Assim, enquanto le cada linha do CSV, ele ja grava:

```text
_cat
_fam
```

Isso e mais rapido do que fazer essa busca toda vez que o painel abre.

## 12. Desfazer Importacao

Quando o usuario clica em Desfazer:

1. O backend busca o log em `logs_importacao`.
2. Se for `dados_brutos`, apaga registros com aquele `_import_id`.
3. Se for `categorias_depara`, apaga a colecao de categorias.
4. Se for `gtin_atlas_depara`, apaga a colecao de produtos Atlas/Aurora e reaplica os aliases.
5. Se for `lojas_depara`, apaga a colecao de lojas.
6. Apaga o log.
7. Limpa o cache.

Esse ponto e importante.

Sem limpar cache, o Mongo pode estar correto, mas o painel pode continuar mostrando numero antigo.

## 13. Usuario Pai

O sistema garante que existe um admin marcado como:

```js
usuarioPai: true
```

Esse usuario tem poder total.

Regras importantes:

- somente a senha do usuario pai libera limpezas criticas;
- usuario pai nao pode ser excluido;
- senha do usuario pai nao pode ser alterada pelo painel;
- e-mail do usuario pai nao pode ser alterado pelo painel;
- senha do usuario pai deve ser redefinida por e-mail.

Isso protege o acesso principal do sistema.

## 14. Performance

As consultas podem ficar lentas quando:

- ha muitos registros em `dados_brutos`;
- `_cat` e `_fam` nao existem;
- o backend precisa fazer lookup em tempo real;
- muitas chamadas sao feitas ao mesmo tempo;
- cache foi limpo e ainda nao foi aquecido.

Para melhorar performance, o projeto usa:

- campos preparados;
- indices no Mongo;
- cache em memoria;
- pre-aquecimento de rotas;
- menor quantidade de chamadas automaticas no frontend.

## 15. Pontos De Atencao Para Manutencao

Ao alterar importacao, confira:

- se `_qtd_num` continua sendo criado;
- se `_valor_num` continua correto;
- se `_gtin` esta padronizado;
- se `_data_iso` esta correto;
- se `_cat` e `_fam` estao sendo gravados;
- se `cacheClear()` e chamado no final.

Ao alterar dashboard, confira:

- se esta usando campos preparados;
- se nao esta fazendo lookup desnecessario;
- se filtros continuam funcionando;
- se a rota esta usando cache.

Ao alterar usuarios, confira:

- regras do usuario pai;
- login;
- reset por e-mail;
- permissoes de limpeza.

## 16. Como Rodar Localmente

Entre na pasta do backend:

```text
cd backend
```

Inicie o servidor:

```text
node src/server.js
```

Depois abra no navegador:

```text
http://localhost:3000
```

## 17. Arquivo .env

O backend depende de variaveis no `.env`.

Exemplos:

```text
MONGODB_URI
DB_NAME
ADMIN_USER
ADMIN_PASSWORD
ADMIN_EMAIL
EMAIL_USER
EMAIL_PASS
```

Nao exponha senhas em codigo ou documentacao publica.

## 18. Checklist Depois De Alterar O Projeto

Antes de considerar pronto:

- rodar `node --check backend/src/server.js`;
- reiniciar o backend;
- testar importacao;
- testar desfazer importacao;
- testar as abas principais;
- verificar console do navegador;
- confirmar se nao aparecem erros 504;
- verificar se o Mongo recebeu os dados esperados.

## 19. Resumo Para Programador

O painel e simples na ideia:

```text
Frontend mostra.
Backend processa.
Mongo guarda.
```

O segredo do desempenho esta nos campos preparados.

Sempre que dados brutos entram, eles precisam sair da importacao ja prontos para consulta.

Se o painel precisar descobrir categoria, data ou valor durante a exibicao, ele fica lento.

Por isso, a importacao e a parte mais importante do projeto.
