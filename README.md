# Painel Atlas - Dashboard MongoDB

**[Ler o estudo de caso](https://lariantunez.github.io/painel-sellout-de-para/)** · **[Abrir a demo do painel](https://lariantunez.github.io/painel-sellout-de-para/painel/)**

![Painel Atlas com dados fictícios](assets/painel-importar.jpg)

> **Versão de portfólio.** O painel original roda em produção, com dados reais do cliente. Esta versão usa o mesmo código, com dados e nomes fictícios.

Dashboard web criado para automatizar a leitura comercial do cliente Atlas, substituindo análises manuais feitas em planilhas dinâmicas por um painel com indicadores filtráveis e importação controlada de dados.

O projeto reaproveita a base técnica do painel Vitta, com adaptações para as particularidades de cadastro, GTIN e De/Para do cliente Atlas.

## Objetivo

Entregar para a diretoria comercial uma visão rápida e confiável de sell out, valor vendido, estoque e desempenho por filial, categoria, família e produto.

## Principais recursos

- Importação de dados brutos em CSV/XLSX.
- De/Para de lojas.
- De/Para de categorias e produtos.
- De/Para intermediário de GTIN Atlas para compatibilizar códigos do cliente.
- Dashboard com KPIs comerciais e gráficos/tabelas analíticas.
- Filtros dinâmicos por período, loja, categoria, família e produto.
- Rotas de agregação no MongoDB com cache em memória.
- Gestão de usuários e administradores.
- Logs de importação e possibilidade de desfazer importações.
- Reset de senha por e-mail.
- Modo demo/readonly para apresentação.

## Tecnologias

- Node.js
- Express
- MongoDB Atlas
- HTML, CSS e JavaScript puro
- Multer
- XLSX
- csv-parser
- Nodemailer
- Docker

## Estrutura

```text
backend/src/server.js      API, importacoes, autenticacao e agregacoes
frontend/index.html        Dashboard e area administrativa
frontend/reset-senha.html  Redefinicao de senha
docs/                      Manuais tecnico e de usuario
README_ATLAS.md            Notas originais do projeto Atlas
Dockerfile                 Deploy conteinerizado
```

## Variáveis de ambiente

Use `.env.example` como base para criar o `.env` local. O `.env` não deve ser commitado.

```env
PORT=3000
DB_NAME=atlas_dashboard
MONGODB_URI=mongodb+srv://USUARIO:SENHA@HOST/?appName=Cluster0
ADMIN_USER=usuario-admin
ADMIN_PASSWORD=senha-admin
ADMIN_EMAIL=email@dominio.com
EMAIL_USER=
EMAIL_PASS=
READ_ONLY=false
```

## Como executar

```bash
npm install
npm start
```

## Demo estática

Este repositório também possui uma demo do painel real em `docs/index.html`, criada para GitHub Pages e portfólio. Ela usa dados fictícios e uma API simulada no navegador, sem depender de MongoDB, backend ou variáveis de ambiente.

## Cuidados de portfólio

Este repositório não deve conter bases reais, credenciais, uploads ou planilhas operacionais. Arquivos sensíveis ficam ignorados por `.gitignore`.

## Status

Projeto funcional com backend, frontend, importação de bases, tratamento de GTIN e deploy preparado via Docker.

---

**🇺🇸 English:** a sell-out dashboard delivered in two days by reusing the architecture of another project. The real challenge was an intermediate GTIN mapping that translates the retailer's internal product codes into the manufacturer's codes, with background remapping and undoable imports. Stack: Node.js, Express, MongoDB Atlas, plain JavaScript, Docker. Portfolio version with fictional data; the original runs in production.
