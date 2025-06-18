function extrairNoticiasMaisVistas() {
  const planilha = SpreadsheetApp.openById("1gFuiGV5-aR_VVtl853-LbqcfG5RLQprnpvarGDt8wf0");
  const hoje = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");

  const palavrasChave = [
    "imóvel","imobiliário","mercado","financiamento","aluguel",
    "lançamento","venda","compra","crédito","investimento","juros","preço"
  ];

  // Configuração de fetch com User-Agent para evitar bloqueios 403
  const fetchOptions = {
    followRedirects: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/90.0.4430.93 Safari/537.36"
    }
  };

  const fontes = [
    // ─── Portais & veículos ──────────────────────────────────────────────
    { nome:"Imobi Report", url:"https://imobireport.com.br", feedUrl:"https://imobireport.com.br/feed", regexMaisVistas:null },
    { nome:"UOL Economia - Imóveis", url:"https://economia.uol.com.br/imoveis/", 
      regexMaisVistas:/<li[^>]*mostRead[^>]*>\s*<a[^>]+href="([^"]+)"[^>]*>(.*?)<\/a>/gi },
    { nome:"Estadão Imóveis", url:"https://imoveis.estadao.com.br/", 
      regexMaisVistas:/id="mais-lidas"[\s\S]*?<a[^>]+href="([^"]+)"[^>]*>(.*?)<\/a>/gi },
    { nome:"Mercado Imobiliário", url:"https://mercadoimobiliario.com.br/", feedUrl:"https://mercadoimobiliario.com.br/blog/feed", regexMaisVistas:null },
    { nome:"Blog da Lopes", url:"https://www.lopes.com.br/blog/", feedUrl:"https://www.lopes.com.br/blog/feed", regexMaisVistas:null },
    { nome:"Blog da Nivu", url:"https://nivu.com.br/blog/", feedUrl:"https://nivu.com.br/blog/feed", regexMaisVistas:null },

    // ─── Blogs de portais imobiliários e plataformas ──────────────────────
    { nome:"ZAP Imóveis Blog", url:"https://www.zapimoveis.com.br/blog/", feedUrl:"https://www.zapimoveis.com.br/blog/feed", regexMaisVistas:null },
    { nome:"VivaReal Blog", url:"https://www.vivareal.com.br/blog/", feedUrl:"https://www.vivareal.com.br/blog/feed", regexMaisVistas:null },
    { nome:"Imovelweb Blog", url:"https://www.imovelweb.com.br/noticias/", feedUrl:"https://www.imovelweb.com.br/noticias/feed", regexMaisVistas:null },
    { nome:"Blog do QuintoAndar", url:"https://imprensa.quintoandar.com.br/blog/", feedUrl:"https://imprensa.quintoandar.com.br/blog/feed", regexMaisVistas:null },

    // ─── Portais de análises e indicadores econômicos ─────────────────────
    { nome:"Exame - Mercado Imobiliário", url:"https://exame.com/mercado-imobiliario/", 
      regexMaisVistas:/widget--trending[\s\S]*?<a[^>]+href="([^"]+?)"[^>]*>(.*?)<\/a>/gi },
    { nome:"InfoMoney - Imóveis", url:"https://www.infomoney.com.br/tudo-sobre/imoveis/", 
      regexMaisVistas:/popular-news__item[\s\S]*?<a[^>]+href="([^"]+?)"[^>]*>(.*?)<\/a>/gi },
    { nome:"CNN Brasil - Mercado Imobiliário", url:"https://www.cnnbrasil.com.br/tudo-sobre/mercado-imobiliario/", feedUrl:"https://www.cnnbrasil.com.br/feed", regexMaisVistas:null },
    { nome:"ABRAINC", url:"https://www.abrainc.org.br/", feedUrl:"https://www.abrainc.org.br/feed", regexMaisVistas:null },

    // ─── Blogs de ferramentas e soluções para o mercado imobiliário ───────
    { nome:"Blog do CV CRM", url:"https://cvcrm.com.br/blog/", feedUrl:"https://cvcrm.com.br/blog/feed", regexMaisVistas:null },
    { nome:"Blog da Jetimob", url:"https://www.jetimob.com/blog/", feedUrl:"https://www.jetimob.com/blog/feed", regexMaisVistas:null },
    { nome:"Blog da Superlógica", url:"https://blog.superlogica.com/imobiliarias/", feedUrl:"https://blog.superlogica.com/imobiliarias/feed", regexMaisVistas:null }
  ];

  let totalSitesComFalha = 0;

  fontes.forEach(fonte => {
    const noticias = [];

    // 1) Tenta bloco “mais lidas” com regex se definido
    if (fonte.regexMaisVistas) {
      try {
        const html = UrlFetchApp.fetch(fonte.url, fetchOptions).getContentText("UTF-8");
        let m;
        while ((m = fonte.regexMaisVistas.exec(html)) && noticias.length < 5) {
          const titulo = limparTexto(m[2]);
          const link   = normalizarLink(fonte.url,m[1]);
          if (titulo && link) noticias.push([hoje,titulo,link]);
        }
        Logger.log(`${fonte.nome}: ${noticias.length} via bloco mais‑lidas`);
      } catch(e) {
        Logger.log(`Erro bloco mais‑lidas (${fonte.nome}): ${e}`);
      }
    }

    // 2) Tenta RSS/Feed se ainda faltar notícias e feed disponível
    if (noticias.length < 5 && fonte.feedUrl) {
      try {
        noticias.push(...pegarPorFeed(fonte.feedUrl, hoje, 5 - noticias.length));
        Logger.log(`${fonte.nome}: ${noticias.length} acumuladas após RSS`);
      } catch(e) {
        Logger.log(`Erro RSS (${fonte.nome}): ${e}`);
      }
    }

    // 3) Fallback por palavras‑chave no HTML para garantir notícias
    if (noticias.length < 5) {
      try {
        const html = UrlFetchApp.fetch(fonte.url, fetchOptions).getContentText("UTF-8");
        const regexLinks = /<a[^>]+href="([^"]+?)"[^>]*>(.*?)<\/a>/gi;
        let m;
        while ((m = regexLinks.exec(html)) && noticias.length < 5) {
          const texto  = limparTexto(m[2]);
          const ok     = palavrasChave.some(p => texto.toLowerCase().includes(p));
          if (!ok) continue;
          const link = normalizarLink(fonte.url,m[1]);
          if (texto && link) noticias.push([hoje,texto,link]);
        }
        Logger.log(`${fonte.nome}: ${noticias.length} acumuladas após fallback`);
      } catch(e) {
        Logger.log(`Erro fallback (${fonte.nome}): ${e}`);
      }
    }

    // 4) Grava na planilha numa aba por site
    if (noticias.length) {
      let aba = planilha.getSheetByName(fonte.nome);
      if (!aba) aba = planilha.insertSheet(fonte.nome); else aba.clear();
      aba.appendRow(["Data","Título","Link"]);
      noticias.forEach(r => aba.appendRow(r));
    } else {
      totalSitesComFalha++;
      Logger.log(`⚠️  ${fonte.nome}: nenhum resultado`);
    }
  });

  if (totalSitesComFalha)
    Logger.log(`Fim: ${totalSitesComFalha} sites retornaram zero notícias.`);
}

/* ---------------- UTILITÁRIAS ---------------- */
function limparTexto(str){
  return str.replace(/<[^>]+>/g,"").replace(/\s+/g," ").trim();
}

function normalizarLink(base,href){
  if(!href) return "";
  if(/^https?:\/\//i.test(href)) return href;
  if(/^\/\//.test(href))         return "https:"+href;
  if(/^\//.test(href))           return base.replace(/^(https?:\/\/[^\/]+).*/,"$1")+href;
  return "";
}

function pegarPorFeed(feedUrl, dataHoje, maxItens){
  const itens=[];
  const xml = UrlFetchApp.fetch(feedUrl,{followRedirects:true, headers: {"User-Agent": "Mozilla/5.0"}}).getContentText("UTF-8");
  const doc = XmlService.parse(xml);
  const root = doc.getRootElement();
  const canal = root.getName()==="rss" ? root.getChild("channel") : root;
  const entradas = canal.getChildren("item").length ? canal.getChildren("item") : canal.getChildren("entry");
  for(let i=0; i<entradas.length && itens.length<maxItens; i++){
    const item = entradas[i];
    const titulo = item.getChildText("title");
    const link   = item.getChildText("link") || (item.getChild("link") ? item.getChild("link").getAttribute("href").getValue() : "");
    if(titulo && link) itens.push([dataHoje,titulo,link]);
  }
  return itens;
}
