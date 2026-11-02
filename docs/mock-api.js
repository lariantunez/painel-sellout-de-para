(function () {
  try {
    sessionStorage.setItem("adm_token", "portfolio-demo-token");
    sessionStorage.setItem("gu_token", "portfolio-demo-token");
  } catch (_) {}
  const repo = (location.pathname || "").toLowerCase();
  const isAtlas = repo.includes("de-para");
  const isBella = repo.includes("contagem");
  const nomePainel = isAtlas ? "Atlas" : isBella ? "Bella" : "Vitta";
  const dbName = isAtlas ? "atlas_dashboard_demo" : isBella ? "bella_dashboard_demo" : "vitta_dashboard_demo";
  const fator = isAtlas ? 0.78 : isBella ? 0.92 : 1;

  const lojas = [
    { Cod_Loja: "1", Nome_Fantasia: "Centro" },
    { Cod_Loja: "2", Nome_Fantasia: "Shopping Norte" },
    { Cod_Loja: "3", Nome_Fantasia: "Shopping Sul" },
    { Cod_Loja: "4", Nome_Fantasia: "Outlet" },
    { Cod_Loja: "5", Nome_Fantasia: "E-commerce" },
    { Cod_Loja: "6", Nome_Fantasia: "Shopping Leste" }
  ];

  const categorias = [
    { CODBARRAS: "7891000000011", PRODUTO: "Shampoo Hidratante", CATEGORIA: isAtlas ? "Cabelo" : "Cabelos", FAMILIA: "Tratamento Capilar" },
    { CODBARRAS: "7891000000028", PRODUTO: "Mascara Nutritiva", CATEGORIA: "Tratamento", FAMILIA: "Tratamento Capilar" },
    { CODBARRAS: "7891000000035", PRODUTO: isAtlas ? "Sabonete Liquido" : "Coloracao 6.0", CATEGORIA: isAtlas ? "Higiene" : "Coloracao", FAMILIA: isAtlas ? "Higiene Pessoal" : "Coloracao" },
    { CODBARRAS: "7891000000042", PRODUTO: "Finalizador Termico", CATEGORIA: isAtlas ? "Perfumaria" : "Finalizadores", FAMILIA: isAtlas ? "Perfumaria" : "Finalizacao" },
    { CODBARRAS: "7891000000059", PRODUTO: isBella ? "Esmalte Gel" : "Kit Manutencao", CATEGORIA: isBella ? "Unhas" : "Acessorios", FAMILIA: isBella ? "Manicure" : "Acessorios" },
    { CODBARRAS: "7891000000066", PRODUTO: isBella ? "Secador Profissional" : "Leave-in Reparador", CATEGORIA: isBella ? "Equipamentos" : "Tratamento", FAMILIA: isBella ? "Equipamentos" : "Finalizacao" }
  ];

  const mesesInfo = [
    { mes: "Jun", num: 6, dias: [3, 10, 17, 24] },
    { mes: "Jul", num: 7, dias: [3, 10, 17, 24] },
    { mes: "Ago", num: 8, dias: [3, 10, 17, 21] }
  ];
  const anos = ["2026"];
  const meses = mesesInfo.map(m => m.mes);
  const catByGtin = Object.fromEntries(categorias.map(c => [c.CODBARRAS, c]));

  function iso(mesNum, dia) {
    return `2026-${String(mesNum).padStart(2, "0")}-${String(dia).padStart(2, "0")}`;
  }

  const dadosBrutos = [];
  mesesInfo.forEach((mesInfo, mIdx) => {
    mesInfo.dias.forEach((dia, dIdx) => {
      lojas.forEach((loja, lIdx) => {
        categorias.forEach((cat, cIdx) => {
          const base = 9 + mIdx * 4 + dIdx * 2 + lIdx * 3 + cIdx;
          const qtd = Math.max(1, Math.round(base * fator));
          const preco = 42 + cIdx * 8 + lIdx * 2;
          const valor = qtd * preco;
          dadosBrutos.push({
            Data: iso(mesInfo.num, dia),
            Ano: "2026",
            "Mês": mesInfo.mes,
            "Mês": mesInfo.mes,
            Mes: String(mesInfo.num),
            Loja: loja.Cod_Loja,
            "GTIN/PLU": cat.CODBARRAS,
            Ean: cat.CODBARRAS,
            Produto: cat.PRODUTO,
            "Venda (Qtd)": qtd,
            "Venda (R$)": valor,
            "Faturamento (Unid)": qtd,
            "Faturamento (R$)": valor,
            "Estoque Diario": Math.round(qtd * (2.4 + (cIdx % 3) * 0.45) + lIdx * 7),
            "Quantidade em Estoque": Math.round(qtd * (2.4 + (cIdx % 3) * 0.45) + lIdx * 7)
          });
        });
      });
    });
  });

  function values(params, key) {
    return params.getAll(key).map(String).filter(Boolean);
  }

  function inList(list, actual) {
    return !list.length || list.includes(String(actual ?? ""));
  }

  function catOf(row) {
    return catByGtin[String(row["GTIN/PLU"])] || {};
  }

  function qtyOf(row) {
    return Number(row["Venda (Qtd)"] || row["Faturamento (Unid)"] || 0);
  }

  function valorOf(row) {
    return Number(row["Venda (R$)"] || row["Faturamento (R$)"] || 0);
  }

  function estoqueOf(row) {
    return Number(row["Estoque Diario"] || row["Quantidade em Estoque"] || 0);
  }

  function lojaNome(cod) {
    return lojas.find(l => l.Cod_Loja === String(cod))?.Nome_Fantasia || `Filial ${cod}`;
  }

  function filterRows(params, rows = dadosBrutos) {
    const anosSel = values(params, "ano");
    const mesesSel = values(params, "mes");
    const lojasSel = values(params, "loja");
    const catsSel = values(params, "cat");
    const famsSel = values(params, "familia");
    const produtosSel = values(params, "produto");
    const produtoGtinsSel = values(params, "produto_gtin");
    const ativoLoja = params.get("ativo_loja");
    const ativoCat = params.get("ativo_cat");
    const ativoFam = params.get("ativo_familia");
    const di = params.get("di");
    const df = params.get("df");

    return rows.filter(row => {
      const cat = catOf(row);
      const data = String(row.Data || "");
      if (!inList(anosSel, row.Ano)) return false;
      if (!inList(mesesSel, row["Mês"] || row["Mês"] || row.Mes)) return false;
      if (!inList(lojasSel, row.Loja)) return false;
      if (!inList(catsSel, cat.CATEGORIA)) return false;
      if (!inList(famsSel, cat.FAMILIA)) return false;
      if (!inList(produtosSel, cat.PRODUTO)) return false;
      if (!inList(produtoGtinsSel, cat.CODBARRAS)) return false;
      if (ativoLoja && row.Loja !== ativoLoja) return false;
      if (ativoCat && cat.CATEGORIA !== ativoCat) return false;
      if (ativoFam && cat.FAMILIA !== ativoFam) return false;
      if (di && data < di) return false;
      if (df && data > df) return false;
      return true;
    });
  }

  function group(rows, keyFn, qtyFn = qtyOf, valorFn = valorOf) {
    const map = new Map();
    rows.forEach(row => {
      const key = keyFn(row);
      if (!key) return;
      if (!map.has(key)) map.set(key, { key, qty: 0, valor: 0 });
      const item = map.get(key);
      item.qty += qtyFn(row);
      item.valor += valorFn(row);
    });
    return [...map.values()].sort((a, b) => b.qty - a.qty);
  }

  function agregado(params = new URLSearchParams()) {
    const rows = filterRows(params);
    const porLoja = group(rows, r => r.Loja).map(r => ({ loja: r.key, qty: r.qty, valor: r.valor }));
    const porDia = group(rows, r => r.Data).sort((a, b) => a.key.localeCompare(b.key)).map(r => ({ data: r.key, qty: r.qty, valor: r.valor }));
    const porCat = group(rows, r => catOf(r).CATEGORIA || "Sem mapeamento").map(r => ({ cat: r.key, qty: r.qty, valor: r.valor }));
    const porFam = group(rows, r => catOf(r).FAMILIA || "Sem mapeamento").map(r => ({ fam: r.key, qty: r.qty, valor: r.valor }));
    const porProduto = group(rows, r => catOf(r).PRODUTO || r.Produto || "SEM PRODUTO").map(r => ({ produto: r.key, nome: r.key, qty: r.qty, valor: r.valor }));
    const porCatDiaMap = group(rows, r => `${catOf(r).CATEGORIA || "Sem mapeamento"}|${r.Data}`);
    const porFamDiaMap = group(rows, r => `${catOf(r).FAMILIA || "Sem mapeamento"}|${r.Data}`);
    return {
      por_loja: porLoja,
      por_dia: porDia,
      por_cat: porCat,
      por_fam: porFam,
      por_produto: porProduto,
      por_cat_dia: porCatDiaMap.map(r => { const [cat, data] = r.key.split("|"); return { cat, data, qty: r.qty, valor: r.valor }; }),
      por_fam_dia: porFamDiaMap.map(r => { const [fam, data] = r.key.split("|"); return { fam, data, qty: r.qty, valor: r.valor }; }),
      total: rows.reduce((s, r) => s + qtyOf(r), 0),
      total_valor: rows.reduce((s, r) => s + valorOf(r), 0),
      total_lojas: new Set(rows.map(r => r.Loja)).size
    };
  }

  function estoque(params = new URLSearchParams()) {
    const rows = filterRows(params);
    const porLoja = group(rows, r => r.Loja, estoqueOf, () => 0).map(r => ({ loja: r.key, nome: lojaNome(r.key), qty: r.qty }));
    const porDia = group(rows, r => r.Data, estoqueOf, () => 0).sort((a, b) => a.key.localeCompare(b.key)).map(r => ({ data: r.key, qty: r.qty }));
    const porLojaDia = group(rows, r => `${r.Loja}|${r.Data}`, estoqueOf, () => 0).map(r => { const [loja, data] = r.key.split("|"); return { loja, nome: lojaNome(loja), data, qty: r.qty }; });
    const porProduto = group(rows, r => catOf(r).PRODUTO || r.Produto || "SEM PRODUTO", estoqueOf, () => 0).map(r => ({ nome: r.key, produto: r.key, qty: r.qty }));
    const porProdutoDia = group(rows, r => `${catOf(r).PRODUTO || r.Produto}|${r.Data}`, estoqueOf, () => 0).map(r => { const [nome, data] = r.key.split("|"); return { nome, produto: nome, data, qty: r.qty }; });
    const porCat = group(rows, r => catOf(r).CATEGORIA || "SEM CATEGORIA", estoqueOf, () => 0).map(r => ({ cat: r.key, qty: r.qty }));
    const porCatDia = group(rows, r => `${catOf(r).CATEGORIA || "SEM CATEGORIA"}|${r.Data}`, estoqueOf, () => 0).map(r => { const [cat, data] = r.key.split("|"); return { cat, data, qty: r.qty }; });
    const porFam = group(rows, r => catOf(r).FAMILIA || "SEM FAMÍLIA", estoqueOf, () => 0).map(r => ({ fam: r.key, qty: r.qty }));
    const porFamDia = group(rows, r => `${catOf(r).FAMILIA || "SEM FAMÍLIA"}|${r.Data}`, estoqueOf, () => 0).map(r => { const [fam, data] = r.key.split("|"); return { fam, data, qty: r.qty }; });
    return {
      total: rows.reduce((s, r) => s + estoqueOf(r), 0),
      total_lojas: new Set(rows.map(r => r.Loja)).size,
      data: rows.map(r => r.Data).sort().at(-1) || "",
      por_loja: porLoja,
      por_dia: porDia,
      por_loja_dia: porLojaDia,
      por_produto: porProduto,
      por_produto_dia: porProdutoDia,
      por_cat: porCat,
      por_cat_dia: porCatDia,
      por_fam: porFam,
      por_fam_dia: porFamDia
    };
  }

  function responseFor(url, init) {
    const baseUrl = location.origin && location.origin !== "null" ? location.origin : "https://portfolio.local";
    const u = new URL(String(url), baseUrl);
    const path = u.pathname;
    const params = u.searchParams;
    if (!path.startsWith("/api/")) return null;

    if (path === "/api/config") return { readOnly: false, ambiente: "portfolio", demo: true, cliente: nomePainel };
    if (path === "/api/health") return { ok: true, db: dbName, mongo: "mocked" };
    if (path === "/api/lojas-depara") return lojas;
    if (path === "/api/categorias-depara") return categorias;
    if (path === "/api/dados-brutos" || path === "/api/dados-tratados") return filterRows(params);
    if (path === "/api/dashboard/filtros") return { anos, meses, lojas: lojas.map(l => l.Cod_Loja) };
    if (path === "/api/dashboard/categorias") return agregado(params).por_cat.map(r => ({ _id: r.cat, total: r.qty, valor: r.valor }));
    if (path === "/api/dashboard/familias") return agregado(params).por_fam.map(r => ({ _id: r.fam, total: r.qty, valor: r.valor }));
    if (path === "/api/dashboard/resumo" || path === "/api/dashboard/kpis") {
      const ag = agregado(params);
      return { total_vendido: ag.total, total_lojas: ag.total_lojas, total_valor: ag.total_valor, cliente: nomePainel };
    }
    if (path === "/api/dashboard/agregados") return agregado(params);
    if (path === "/api/dashboard/estoque" || path === "/api/dashboard/estoque-resumo") return estoque(params);
    if (path === "/api/login" || path === "/api/admin/login") return { token: "portfolio-demo-token" };
    if (path === "/api/logout" || path === "/api/admin/logout") return { ok: true };
    if (path === "/api/solicitar-reset" || path === "/api/admin/solicitar-reset") return { ok: true, mensagem: "Demo: e-mail ficticio enviado." };
    if (path === "/api/admin/usuarios") return [{ _id: "u1", usuario: "analista.demo", email: "demo@portfolio.local" }];
    if (path === "/api/admin/admins") return [{ _id: "a1", usuario: "larissa.demo", email: "demo@portfolio.local", usuarioPai: true }];
    if (path === "/api/admin/templates") return [
      { filename: "modelo_dados_brutos.csv", tipo: "dados_brutos", total: dadosBrutos.length },
      { filename: "modelo_categorias_depara.xlsx", tipo: "categorias_depara", total: categorias.length },
      { filename: "modelo_lojas_depara.csv", tipo: "lojas_depara", total: lojas.length }
    ];
    if (path === "/api/admin/logs-importacao") return [
      { _id: "log1", tipo: "dados_brutos", arquivo: `dados_${nomePainel.toLowerCase()}_demo.csv`, usuario: "analista.demo", total: dadosBrutos.length, data: "2026-08-21T12:00:00.000Z" },
      { _id: "log2", tipo: "categorias_depara", arquivo: "categorias_demo.xlsx", usuario: "analista.demo", total: categorias.length, data: "2026-08-20T12:00:00.000Z" }
    ];
    if (path.startsWith("/api/importar/")) return { ok: true, inserido: dadosBrutos.length, ultimo: true, mensagem: "Demo estatica: importacao simulada com dados ficticios." };
    if (path.startsWith("/api/templates/")) return "Arquivo,Descricao\nmodelo_demo,Conteudo ficticio para portfolio\n";
    if (path.startsWith("/api/admin/") || path === "/api/cache/limpar" || path === "/api/importar/dados-brutos/cancelar") return { ok: true };
    return { ok: true, demo: true };
  }

  const realFetch = window.fetch.bind(window);
  window.fetch = function (url, init = {}) {
    const body = responseFor(url, init);
    if (body === null) return realFetch(url, init);
    const isText = typeof body === "string";
    return Promise.resolve(new Response(isText ? body : JSON.stringify(body), {
      status: 200,
      headers: { "Content-Type": isText ? "text/csv; charset=utf-8" : "application/json; charset=utf-8" }
    }));
  };

  class MockXHR {
    constructor() {
      this.headers = {};
      this.upload = {};
      this.readyState = 0;
      this.status = 0;
      this.responseText = "";
    }
    open(method, url) {
      this.method = method;
      this.url = url;
      this.readyState = 1;
    }
    setRequestHeader(key, value) {
      this.headers[key] = value;
    }
    getResponseHeader(name) {
      return name && name.toLowerCase() === "content-type" ? "application/json; charset=utf-8" : null;
    }
    send() {
      setTimeout(() => {
        if (typeof this.upload.onprogress === "function") this.upload.onprogress({ lengthComputable: true, loaded: 100, total: 100 });
        const body = responseFor(this.url, { method: this.method }) || { ok: true };
        this.status = 200;
        this.readyState = 4;
        this.responseText = typeof body === "string" ? body : JSON.stringify(body);
        if (typeof this.onload === "function") this.onload();
        if (typeof this.onreadystatechange === "function") this.onreadystatechange();
      }, 0);
    }
    abort() {}
  }
  window.XMLHttpRequest = MockXHR;

  if (!window.Chart) {
    class PortfolioChartFallback {
      constructor(canvas, config) {
        this.canvas = canvas;
        this.config = config || {};
        const wrap = canvas && canvas.parentElement;
        if (wrap) {
          const state = wrap.querySelector('.state-msg');
          if (state) state.style.display = 'none';
        }
      }
      destroy() {}
      update() {}
      resize() {}
    }
    window.Chart = PortfolioChartFallback;
  }
  window.__PORTFOLIO_DEMO__ = { cliente: nomePainel, linhas: dadosBrutos.length };
})();