# Manual do Usuario - Painel Atlas

Este manual explica, em linguagem simples, como o painel funciona e como usar as principais telas.

## 1. O Que E Este Projeto

Este projeto e um painel para acompanhar vendas, valor vendido, estoque e importacoes do Atlas.

Ele transforma arquivos importados em graficos, tabelas e indicadores.

Em vez de abrir planilhas grandes manualmente, o usuario importa os arquivos no painel e o sistema organiza as informacoes.

O painel mostra, por exemplo:

- quantidade vendida;
- valor vendido;
- venda por loja;
- venda por dia;
- estoque por loja;
- estoque por produto;
- categorias e familias;
- historico de importacoes;
- usuarios com acesso ao sistema.

## 2. Como O Painel Funciona Em Termos Simples

Pense no painel como uma cozinha.

O arquivo importado e o ingrediente bruto.

O sistema pega esse ingrediente, limpa, organiza e guarda em um banco de dados.

Depois, quando voce abre uma aba do painel, ele pega os dados ja organizados e monta os graficos.

O fluxo e assim:

```text
Arquivo importado
↓
Sistema le e organiza
↓
Dados ficam salvos no banco
↓
Painel consulta esses dados
↓
Graficos e tabelas aparecem na tela
```

## 3. Principais Abas Do Painel

### Vendas Sell Out

Mostra a quantidade vendida.

Aqui voce ve:

- total de itens vendidos;
- lojas com venda;
- maior filial;
- grafico por filial;
- detalhamento por loja;
- detalhamento por categoria;
- detalhamento por familia.

Essa aba trabalha principalmente com a coluna:

```text
Venda (Qtd)
```

### Venda Valor

Mostra o dinheiro vendido.

Aqui voce ve:

- total vendido em reais;
- lojas com venda;
- maior filial em valor;
- valor vendido por filial;
- valor por categoria;
- valor por familia.

Essa aba trabalha principalmente com a coluna:

```text
Venda (R$)
```

### Venda Por Dia

Mostra como as vendas aconteceram ao longo dos dias.

Aqui voce consegue acompanhar:

- total vendido no periodo;
- dias com venda;
- media diaria;
- melhor dia;
- grafico de venda por dia;
- tabela diaria.

Essa aba usa a coluna:

```text
Data
```

### Estoque

Mostra as informacoes de estoque.

Aqui voce ve:

- estoque total;
- lojas com estoque;
- maior estoque por filial;
- estoque por loja;
- estoque por produto;
- estoque por categoria.

Essa aba usa principalmente a coluna:

```text
Estoque Diario
```

### Sell Out x Estoque

Compara venda e estoque.

Ela ajuda a enxergar situacoes como:

- loja vendeu muito, mas tem pouco estoque;
- loja tem estoque alto e venda baixa;
- risco de ruptura;
- diferenca entre venda e estoque.

### Importar

E a tela usada para enviar arquivos ao sistema.

Normalmente existem tres tipos de arquivos:

- Dados Brutos;
- De/Para Produtos Atlas/Aurora;
- De/Para Lojas.

Cada um tem uma funcao diferente.

## 4. O Que Sao Dados Brutos

Dados brutos sao as linhas principais da planilha de vendas e estoque.

Cada linha do arquivo representa uma informacao de venda ou estoque.

Exemplo simples:

```text
Ano: 2026
Mes: Mar
Data: 01/03/2026
Loja: 23
GTIN/PLU: 789123
Produto: Produto X
Venda (Qtd): 10
Venda (R$): 199,90
Estoque Diario: 50
```

Quando esse arquivo e importado, cada linha vira um registro dentro do banco de dados.

## 5. O Que Sao Os Arquivos De/Para

### De/Para Lojas

Serve para transformar codigo de loja em nome de loja.

Exemplo:

```text
23 -> VITTA LJ23 - MAIRIPORA
```

Se esse arquivo nao estiver correto, o painel pode mostrar codigo sem nome, loja sem identificacao ou informacoes incompletas.

### De/Para Produtos Atlas/Aurora

Serve para ligar o produto cadastrado no Atlas ao produto correto da Aurora.

Ele pode usar o GTIN do Atlas, o codigo interno do Atlas e o codigo de barras correto da Aurora.
Tambem alimenta nome, categoria e familia usados no painel.

Exemplo:

```text
GTIN_ATLAS / COD_INTERNO_ATLAS -> CODBARRAS_AURORA -> CATEGORIA_AURORA -> FAMILIA_AURORA
```

Se esse arquivo nao estiver correto, o painel pode mostrar produtos sem categoria, sem familia ou associados ao item errado.

## 6. Como Importar Arquivos

1. Entre na aba Importar.
2. Escolha o tipo de arquivo.
3. Selecione o arquivo correto.
4. Clique em importar.
5. Aguarde a mensagem de conclusao.
6. Clique em atualizar painel ou recarregue a pagina.

Depois de importar dados brutos, o sistema pode demorar um pouco para preparar todos os dados.

Isso e normal quando o arquivo e grande.

## 7. O Que Acontece Durante A Importacao

Durante a importacao, o sistema nao apenas salva a planilha.

Ele tambem prepara informacoes para deixar o painel mais rapido.

Ele cria, por exemplo:

- quantidade vendida em formato numerico;
- valor vendido em formato numerico;
- data padronizada;
- codigo do produto padronizado;
- categoria do produto;
- familia do produto;
- identificador da importacao.

Essas informacoes preparadas ajudam o painel a carregar mais rapido depois.

## 8. Historico De Importacoes

Na area Gerenciar Usuarios existe a aba Importacoes.

Ela mostra o historico de arquivos importados.

Ali voce pode ver:

- data e hora;
- usuario que importou;
- nome do arquivo;
- tipo da importacao;
- quantidade de registros;
- botao Desfazer.

## 9. O Que Faz O Botao Desfazer

O botao Desfazer reverte uma importacao.

Se for uma importacao de dados brutos, ele apaga do banco os registros daquele arquivo.

Se for uma importacao de de/para, ele remove os dados daquele de/para.

Depois disso, o painel limpa o cache para nao mostrar numeros antigos.

## 10. O Que E Cache

Cache e uma memoria temporaria.

O painel usa cache para carregar mais rapido.

Exemplo:

```text
Primeira vez: sistema calcula tudo
Segunda vez: sistema usa resposta pronta
```

Quando uma importacao e feita ou desfeita, o cache precisa ser limpo.

Se o cache nao for limpo, o painel pode mostrar numeros antigos por alguns minutos.

## 11. Usuarios Do Sistema

Existem usuarios de importacao e administradores.

### Usuarios de Importacao

Sao usuarios que podem acessar a parte de importacao.

### Administradores

Sao usuarios que podem gerenciar:

- usuarios;
- administradores;
- modelos de importacao;
- historico de importacoes.

## 12. Usuario Pai

O sistema tem um usuario especial chamado usuario pai.

Ele e o usuario com autonomia total.

Somente o usuario pai pode liberar algumas acoes criticas, como limpeza de historico ou limpeza geral.

O usuario pai nao pode ter senha alterada diretamente pelo painel.

A troca de senha dele deve ser feita apenas por redefinicao via e-mail.

Isso protege o sistema, porque evita que outro administrador altere a senha do usuario principal.

## 13. Cuidados Importantes

Antes de importar:

- confira se o arquivo esta no modelo correto;
- confira se as colunas estao certas;
- confira se o de/para de lojas esta atualizado;
- confira se o De/Para Produtos Atlas/Aurora esta atualizado.

Depois de importar:

- aguarde a conclusao;
- atualize o painel;
- confira os totais principais;
- se algo estiver estranho, confira o historico de importacao.

## 14. Problemas Comuns

### O painel parece lento apos importar

Pode acontecer quando o arquivo e grande.

O sistema precisa salvar os dados, preparar campos e aquecer o cache.

Aguarde alguns instantes e atualize a pagina.

### O painel mostra dados antigos

Pode ser cache.

Tente atualizar com:

```text
Ctrl + F5
```

### Uma loja aparece sem nome

Provavelmente o codigo da loja nao esta no De/Para Lojas.

### Produto aparece sem categoria

Provavelmente o GTIN/PLU ou codigo interno do Atlas nao foi encontrado no De/Para Produtos Atlas/Aurora.

### Desfiz uma importacao e os numeros continuam aparecendo

Isso geralmente e cache.

O sistema ja foi ajustado para limpar cache ao desfazer importacoes, mas se a tela estiver aberta, use Ctrl + F5.

## 15. Resumo

O painel existe para transformar arquivos de vendas e estoque em informacoes visuais.

O usuario importa os arquivos.

O sistema organiza e salva no banco.

O painel mostra os resultados em abas, graficos e tabelas.

O ponto mais importante e manter os arquivos de dados brutos e de/para sempre corretos.
