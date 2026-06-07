const GSC_SITE_URL    = 'sc-domain:helenacrm.com';
const GA4_PROPERTY_ID = '505588648';
const DIAS            = 30;
const SHEET_ID        = '1OAwbDeI03w898KpRPOui7IesjXWAwWlEQGPH9LjBOYk';
const LEADS_SHEET_GID = 1698310909;

/* ─────────────────────────────────────────────
   CORS + roteamento principal
───────────────────────────────────────────── */
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  // ── Proteção por token ──────────────────────────
  const props = PropertiesService.getScriptProperties();
  const TOKEN_ESPERADO = props.getProperty('API_TOKEN');
  const tokenRecebido  = e && e.parameter && e.parameter.token ? e.parameter.token : '';
  if (!TOKEN_ESPERADO || tokenRecebido !== TOKEN_ESPERADO) {
    output.setContent(JSON.stringify({ erro: 'Acesso negado. Token inválido.' }));
    return output;
  }
  // ───────────────────────────────────────────────

  try {
    const dias = e && e.parameter && e.parameter.dias ? parseInt(e.parameter.dias) : DIAS;
    // Período personalizado: aceita de/ate no formato YYYY-MM-DD
    const deStr  = e && e.parameter && e.parameter.de  ? e.parameter.de  : null;
    const ateStr = e && e.parameter && e.parameter.ate ? e.parameter.ate : null;
    const dados = {
      gsc:        buscarGSC(dias, deStr, ateStr),
      ga4:        buscarGA4(dias, deStr, ateStr),
      semrush:    buscarSEMrush(),
      leads:      buscarLeads(),
      atualizado: new Date().toISOString(),
    };
    output.setContent(JSON.stringify(dados));
  } catch (err) {
    output.setContent(JSON.stringify({ erro: err.message }));
  }
  return output;
}

/* ─────────────────────────────────────────────
   LIMPAR CACHE (rode manualmente após deploy)
───────────────────────────────────────────── */
function limparCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['leads_meta','leads_porMes','leads_porCanalMes','leads_porCanal','leads_porConteudo','leads_recentes','dados_7','dados_30','dados_90','semrush_12h']);
  Logger.log('Cache limpo com sucesso.');
}

/* ════════════════════════════════════════════════════════════════
   LEADS — planilha Helena CRM  (classificação v2 — jun/2026)
   Regras: base válida (exclui CS-, LEAD FALSO, VAGA(S), HELENA TALKS)
           + dedup contato único por mês (chave = Nome do contato)
           + funil prioridade Ganho > SQL > MQL > Lead
           + 3 grupos de origem: Pago / Ads · Orgânico identificado · Sem origem
═══════════════════════════════════════════════════════════════════ */
function buscarLeads() {
  const cache = CacheService.getScriptCache();
  const c0 = cache.get('leads_meta');
  const c1 = cache.get('leads_porMes');
  const c2 = cache.get('leads_porCanalMes');
  const c3 = cache.get('leads_porCanal');
  const c4 = cache.get('leads_porConteudo');
  const c5 = cache.get('leads_recentes');
  if (c0 && c1 && c2 && c3 && c4 && c5) {
    return {
      ...JSON.parse(c0),
      porMes:      JSON.parse(c1),
      porCanalMes: JSON.parse(c2),
      porCanal:    JSON.parse(c3),
      porConteudo: JSON.parse(c4),
      recentes:    JSON.parse(c5),
    };
  }

  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheets().find(s => s.getSheetId() === LEADS_SHEET_GID) || ss.getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { total:0, cf:0, wl:0, porMes:[], porCanalMes:[], porCanal:[], porConteudo:[], recentes:[] };
    }

    const header = data[0].map(h => String(h).toLowerCase().trim());
    const col = name => header.indexOf(name);
    const iData      = col('data de criação do card');
    const iNome      = col('nome do contato');
    const iCanal     = col('origem');
    const iEtiquetas = col('etiquetas');
    const iStatus    = col('etapa no crm');
    const iFunil     = col('funil/painel do crm');
    const iAnuncio   = col('anúncio');
    const iCampanha  = col('campanha');

    const STAGE_RANK = { lead:1, mql:2, sql:3, ganho:4 };
    const GRUPO_RANK = { sem_origem:1, nao_classificada:2, organico:3, pago:4 };
    const labelGrupo = g => g === 'sem_origem'       ? 'Sem origem de conversão'
                          : g === 'pago'             ? 'Pago / Ads'
                          : g === 'organico'         ? 'Orgânico identificado'
                          :                            'Origem rastreada não classificada';

    // ── classificação de origem em 4 baldes (Pago / Orgânico / Sem origem / Não classificada) ──
    const PAGAS_ORIGEM    = ['METAADS','FACEBOOK','INSTAGRAM','IG','META','AD','GOOGLEADS'];
    const ORGANICA_ORIGEM = ['WIDGET','LINKTREE','CHATGPT.COM','SITE'];
    const ORGANICA_CONTEUDO = [
      'HELENACRM | PLATAFORMA DE CRM PARA WHATSAPP','HELENACRM / PLATAFORMA',
      'CENTRAL DE ATENDIMENTO E CRM WHITE LABEL PARA WHATSAPP','CRM WHITE LABEL PARA WHATSAPP',
      'TESTE GRÁTIS - HELENACRM','QUERO CONHECER A PARCERIA WHITE LABEL DA HELENACRM',
      'PLANOS','DEMONSTRAÇÃO HELENA','MANUAL DA API OFICIAL','CRM INTEGRADO COM WHATSAPP',
    ];
    const classificarOrigem = (origemR, campanhaR, anuncioR) => {
      const o=(origemR||'').toUpperCase().trim();
      const c=(campanhaR||'').toUpperCase().trim();
      const a=(anuncioR||'').toUpperCase().trim();
      if (!o && !c && !a) return 'sem_origem';                       // nada de origem/campanha/anúncio
      const campanhaPaga =
        /(^|[^A-Z])WL2?($|[^A-Z])/.test(c) || /WL\s*#/.test(c) ||    // WL, WL2, WL #
        c.indexOf('SITE_SM_IG')>=0 || c.indexOf('SITE_INT')>=0 ||
        c.indexOf('IG_SD')>=0 || c.indexOf('IG_CAP')>=0 || /AG\s*#/.test(c) ||
        /^\d{6,}$/.test(c) || /^\d{6,}$/.test(a);                    // IDs numéricos de campanha/anúncio
      if (PAGAS_ORIGEM.indexOf(o)>=0 || campanhaPaga) return 'pago';
      const ehConteudo = ORGANICA_CONTEUDO.some(p => c.indexOf(p)>=0 || a.indexOf(p)>=0 || o.indexOf(p)>=0);
      if (ORGANICA_ORIGEM.indexOf(o)>=0 || ehConteudo) return 'organico';  // GOOGLE só é orgânico se casar conteúdo (acima)
      return 'nao_classificada';                                     // tem rastro mas não encaixou → revisão
    };

    /* ── 1ª passada: monta registros válidos (já com exclusões) ── */
    const registros = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawData = row[iData];
      if (!rawData) continue;
      const d = rawData instanceof Date ? rawData : new Date(rawData);
      if (isNaN(d.getTime())) continue;

      const nome      = iNome      >= 0 ? String(row[iNome]      || '').trim() : '';
      const origem    = iCanal     >= 0 ? String(row[iCanal]     || '').trim() : '';
      const etq       = iEtiquetas >= 0 ? String(row[iEtiquetas] || '').toLowerCase().trim() : '';
      const statusRaw = iStatus    >= 0 ? String(row[iStatus]    || '').trim() : '';
      const etapa     = statusRaw.toLowerCase();
      const funil     = iFunil     >= 0 ? String(row[iFunil]     || '').toLowerCase().trim() : '';
      const anuncio   = iAnuncio   >= 0 ? String(row[iAnuncio]   || '').trim() : '';
      const campanha  = iCampanha  >= 0 ? String(row[iCampanha]  || '').trim() : '';

      // ── exclusões da base válida ──
      if (etq.includes('lead falso'))   continue;  // LEAD FALSO
      if (/\bcs\s*-/.test(etq))         continue;  // CS - <responsável> (já é cliente)
      if (etq.includes('helena talks')) continue;  // HELENA TALKS / SP
      // vaga: olhar etiqueta + origem + campanha + anúncio (VAGA / VAGAS / VAGAS_BH)
      const blobVaga = (etq + ' ' + origem + ' ' + campanha + ' ' + anuncio).toLowerCase();
      if (/\bvaga/.test(blobVaga)) continue;

      // ── tipo de cliente: TODO contato válido conta.
      //    WL só com 4'LEAD P WHITE LABEL e SEM etiqueta de Cliente Final; o resto é Cliente Final. ──
      const temWlTag = etq.includes("4'lead p white label");
      const temCfTag = etq.includes('cliente final');
      const tipo = (temWlTag && !temCfTag) ? 'white_label' : 'cliente_final';

      // etapa do funil — prioridade Ganho > SQL > MQL > Lead
      // MQL = SÓ pela etiqueta 7'MQL. Estar no "Funil Parceiros (Qualificação)" NÃO basta,
      // porque esse funil também contém os Lead Frio (que devem ficar como Lead).
      const isGanho = etapa.includes('ganho');
      const isSql   = etq.includes("8'sql") || funil.includes('(sql)') || funil.includes('comercial');
      const isMql   = etq.includes("7'mql");
      const stage   = isGanho ? 'ganho' : isSql ? 'sql' : isMql ? 'mql' : 'lead';  // Lead = 6'Lead Frio ou sem tag de qualificação
      const perdido = etapa.includes('perdido') && !isGanho;

      // origem — 4 baldes (Pago / Orgânico / Sem origem / Não classificada)
      const grupo = classificarOrigem(origem, campanha, anuncio);

      const mesKey = String(d.getFullYear()) + String(d.getMonth() + 1).padStart(2, '0');
      const mes    = d.toLocaleDateString('pt-BR', { month: 'short', year: '2-digit' })
                       .replace('.', '').replace(' de ', '/');

      registros.push({
        nomeKey: nome.toLowerCase().replace(/\s+/g, ' ').trim(),
        nome, tipo, stage, perdido, grupo,
        conteudo: campanha || anuncio || '',
        anuncio, campanha, statusRaw, d, mesKey, mes,
      });
    }

    /* ── 2ª passada: dedup por (mês + nome), mantendo etapa mais avançada ── */
    const dedup = {};
    let anon = 0;
    registros.forEach(r => {
      const key = r.nomeKey ? (r.mesKey + '|' + r.nomeKey) : (r.mesKey + '|__anon' + (anon++));
      const ex = dedup[key];
      if (!ex) { dedup[key] = r; return; }
      const hi = STAGE_RANK[r.stage] > STAGE_RANK[ex.stage] ? r : ex;  // etapa mais avançada
      const gw = GRUPO_RANK[r.grupo] >= GRUPO_RANK[ex.grupo] ? r : ex;  // origem mais informativa
      dedup[key] = {
        nomeKey: ex.nomeKey, nome: hi.nome, tipo: hi.tipo, stage: hi.stage,
        perdido: (ex.perdido || r.perdido) && hi.stage !== 'ganho',
        grupo: gw.grupo, conteudo: gw.conteudo, anuncio: gw.anuncio, campanha: gw.campanha,
        statusRaw: hi.statusRaw, d: (ex.d > r.d ? ex.d : r.d), mesKey: ex.mesKey, mes: ex.mes,
      };
    });
    const unicos = Object.values(dedup);

    /* ── 3ª passada: agregações ── */
    const novo = extra => Object.assign({
      total:0, cf:0, wl:0, lead:0, mql:0, sql:0, ganho:0, perdido:0,
      pago:0, organico:0, sem_origem:0, nao_classificada:0,
      cf_lead:0, cf_mql:0, cf_sql:0, cf_ganho:0,
      wl_lead:0, wl_mql:0, wl_sql:0, wl_ganho:0,
    }, extra);
    const acumula = (o, r) => {
      o.total++; if ('count' in o) o.count++;
      o[r.grupo]++; o[r.stage]++;
      if (r.perdido) o.perdido++;
      if (r.tipo === 'white_label') { o.wl++; o['wl_' + r.stage]++; }
      else                          { o.cf++; o['cf_' + r.stage]++; }
    };

    let total = 0, cf = 0, wl = 0;
    const mesMap = {}, canalMap = {}, canalMesMap = {}, conteudoMap = {};
    unicos.forEach(r => {
      total++; if (r.tipo === 'white_label') wl++; else cf++;
      const gl = labelGrupo(r.grupo);

      if (!mesMap[r.mesKey]) mesMap[r.mesKey] = novo({ mes: r.mes, mesKey: r.mesKey });
      acumula(mesMap[r.mesKey], r);

      if (!canalMap[r.grupo]) canalMap[r.grupo] = novo({ canal: gl, tipoOrigem: r.grupo });
      acumula(canalMap[r.grupo], r);

      const cmKey = r.mesKey + '|' + r.grupo;
      if (!canalMesMap[cmKey]) canalMesMap[cmKey] = novo({ mesKey: r.mesKey, canal: gl, tipoOrigem: r.grupo, count: 0 });
      acumula(canalMesMap[cmKey], r);

      const ctg = r.conteudo || gl;
      const ccKey = r.mesKey + '|' + ctg;
      if (!conteudoMap[ccKey]) conteudoMap[ccKey] = novo({ mesKey: r.mesKey, conteudo: ctg, canal: gl, tipoOrigem: r.grupo });
      acumula(conteudoMap[ccKey], r);
    });

    const recentes = unicos.slice().sort((a, b) => b.d - a.d).slice(0, 50).map(r => ({
      data: r.d.toLocaleDateString('pt-BR'), nome: r.nome, canal: labelGrupo(r.grupo),
      tipo: r.tipo, status: r.statusRaw || 'lead', stage: r.stage,
      isMql: r.stage === 'mql', isSql: r.stage === 'sql', isGanho: r.stage === 'ganho', isPerdido: r.perdido,
      conteudo: r.conteudo, anuncio: r.anuncio, campanha: r.campanha, pago: r.grupo === 'pago',
    }));

    const result = {
      total, cf, wl,
      porMes:      Object.values(mesMap).sort((a, b) => a.mesKey > b.mesKey ? 1 : -1),
      porCanalMes: Object.values(canalMesMap),
      porCanal:    Object.values(canalMap).sort((a, b) => b.total - a.total),
      porConteudo: Object.values(conteudoMap).sort((a, b) => b.total - a.total).slice(0, 300),
      recentes,
    };

    const ttl = 3600;
    cache.put('leads_meta',        JSON.stringify({ total: result.total, cf: result.cf, wl: result.wl }), ttl);
    cache.put('leads_porMes',      JSON.stringify(result.porMes),      ttl);
    cache.put('leads_porCanalMes', JSON.stringify(result.porCanalMes), ttl);
    cache.put('leads_porCanal',    JSON.stringify(result.porCanal),    ttl);
    cache.put('leads_porConteudo', JSON.stringify(result.porConteudo), ttl);
    cache.put('leads_recentes',    JSON.stringify(result.recentes),    ttl);
    return result;

  } catch (err) {
    Logger.log('Erro buscarLeads: ' + err.message);
    return { erro: err.message, total:0, cf:0, wl:0, porMes:[], porCanalMes:[], porCanal:[], porConteudo:[], recentes:[] };
  }
}

/* ─────────────────────────────────────────────
   GSC — Google Search Console (via REST)
───────────────────────────────────────────── */
function buscarGSC(dias, deStr, ateStr) {
  try {
    const fmt  = d => d.toISOString().slice(0,10);
    let ini, hoje;
    if (deStr && ateStr) {
      ini  = new Date(deStr + 'T00:00:00Z');
      hoje = new Date(ateStr + 'T00:00:00Z');
    } else {
      hoje = new Date();
      ini  = new Date(hoje); ini.setDate(ini.getDate() - dias);
    }
    const token = ScriptApp.getOAuthToken();
    const siteEncoded = encodeURIComponent(GSC_SITE_URL);
    const url = `https://www.googleapis.com/webmasters/v3/sites/${siteEncoded}/searchAnalytics/query`;
    const opts = body => ({
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });

    const base = { startDate: fmt(ini), endDate: fmt(hoje) };

    // query sem dimensão = totais reais do site (sem limite de rowLimit)
    const respTot  = UrlFetchApp.fetch(url, opts({ ...base }));
    const respKw   = UrlFetchApp.fetch(url, opts({ ...base, dimensions:['query'], rowLimit:500 }));
    const respTime = UrlFetchApp.fetch(url, opts({ ...base, dimensions:['date'],  rowLimit:90  }));
    const respPags = UrlFetchApp.fetch(url, opts({ ...base, dimensions:['page'],  rowLimit:50  }));

    const tot  = JSON.parse(respTot.getContentText());
    const kw   = JSON.parse(respKw.getContentText());
    const time = JSON.parse(respTime.getContentText());
    const pags = JSON.parse(respPags.getContentText());

    // totais reais do site (sem rowLimit, todos os dados)
    const totRow = tot.rows && tot.rows[0];
    const totais = totRow ? {
      cliques:    totRow.clicks,
      impressoes: totRow.impressions,
      ctr:        +(totRow.ctr * 100).toFixed(1),
      posicao:    +totRow.position.toFixed(1),
    } : null;

    const keywords = (kw.rows || []).map(r => ({
      query:      r.keys[0],
      cliques:    r.clicks,
      impressoes: r.impressions,
      ctr:        +(r.ctr * 100).toFixed(1),
      posicao:    +r.position.toFixed(1),
    }));
    const historico = (time.rows || []).map(r => ({
      data:       r.keys[0],
      cliques:    r.clicks,
      impressoes: r.impressions,
      ctr:        +(r.ctr * 100).toFixed(1),
    }));
    const paginas = (pags.rows || []).map(r => ({
      url:        r.keys[0],
      cliques:    r.clicks,
      impressoes: r.impressions,
      ctr:        +(r.ctr * 100).toFixed(1),
      posicao:    +r.position.toFixed(1),
    }));

    return { totais, keywords, historico, paginas };

  } catch (err) {
    Logger.log('Erro GSC: ' + err.message);
    return { erro: err.message, keywords:[], historico:[], paginas:[] };
  }
}

/* ─────────────────────────────────────────────
   GA4 — Google Analytics 4
───────────────────────────────────────────── */
function buscarGA4(dias, deStr, ateStr) {
  dias = dias || DIAS;
  try {
    const fmt  = d => d.toISOString().slice(0,10);
    let ini, hoje;
    if (deStr && ateStr) {
      ini  = new Date(deStr + 'T00:00:00Z');
      hoje = new Date(ateStr + 'T00:00:00Z');
    } else {
      hoje = new Date();
      ini  = new Date(hoje); ini.setDate(ini.getDate() - dias);
    }

    // totais reais do site (sem dimensão)
    const bodyTot = {
      dateRanges: [{ startDate: fmt(ini), endDate: fmt(hoje) }],
      metrics: [
        { name: 'sessions' },
        { name: 'activeUsers' },
        { name: 'newUsers' },
        { name: 'bounceRate' },
        { name: 'engagementRate' },
      ],
    };
    const respTot = AnalyticsData.Properties.runReport(bodyTot, 'properties/' + GA4_PROPERTY_ID);
    const totRow  = respTot.rows && respTot.rows[0];
    const totais  = totRow ? {
      sessoes:       parseInt(totRow.metricValues[0].value),
      usuariosAtivos:parseInt(totRow.metricValues[1].value),
      novosUsuarios: parseInt(totRow.metricValues[2].value),
      bounce:        +(parseFloat(totRow.metricValues[3].value) * 100).toFixed(1),
      engajamento:   +(parseFloat(totRow.metricValues[4].value) * 100).toFixed(1),
    } : null;

    const body = {
      dateRanges: [{ startDate: fmt(ini), endDate: fmt(hoje) }],
      dimensions: [{ name: 'sessionDefaultChannelGrouping' }],
      metrics: [
        { name: 'sessions' },
        { name: 'newUsers' },
        { name: 'bounceRate' },
        { name: 'engagementRate' },
      ],
    };
    const resp   = AnalyticsData.Properties.runReport(body, 'properties/' + GA4_PROPERTY_ID);
    const canais = (resp.rows || []).map(r => ({
      canal:         r.dimensionValues[0].value,
      sessoes:       parseInt(r.metricValues[0].value),
      novosUsuarios: parseInt(r.metricValues[1].value),
      bounce:        +(parseFloat(r.metricValues[2].value) * 100).toFixed(1),
      engajamento:   +(parseFloat(r.metricValues[3].value) * 100).toFixed(1),
    }));

    const bodyDev = {
      dateRanges: [{ startDate: fmt(ini), endDate: fmt(hoje) }],
      dimensions: [{ name: 'deviceCategory' }],
      metrics:    [{ name: 'sessions' }],
    };
    const respDev      = AnalyticsData.Properties.runReport(bodyDev, 'properties/' + GA4_PROPERTY_ID);
    const dispositivos = (respDev.rows || []).map(r => ({
      tipo:    r.dimensionValues[0].value,
      sessoes: parseInt(r.metricValues[0].value),
    }));

    const bodyOri = {
      dateRanges: [{ startDate: fmt(ini), endDate: fmt(hoje) }],
      dimensions: [
        { name: 'sessionDefaultChannelGrouping' },
        { name: 'sessionSource' },
        { name: 'sessionMedium' },
      ],
      metrics: [
        { name: 'sessions' },
        { name: 'newUsers' },
        { name: 'engagementRate' },
      ],
      limit:    50,
      orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    };
    const respOri = AnalyticsData.Properties.runReport(bodyOri, 'properties/' + GA4_PROPERTY_ID);
    const origens = (respOri.rows || []).map(r => ({
      canal:       r.dimensionValues[0].value,
      origem:      r.dimensionValues[1].value + ' / ' + r.dimensionValues[2].value,
      sessoes:     parseInt(r.metricValues[0].value),
      usuarios:    parseInt(r.metricValues[1].value),
      engajamento: +(parseFloat(r.metricValues[2].value) * 100).toFixed(1),
    }));

    return { totais, canais, dispositivos, origens };

  } catch (err) {
    Logger.log('Erro GA4: ' + err.message);
    return { erro: err.message, canais:[], dispositivos:[], origens:[] };
  }
}

/* ─────────────────────────────────────────────
   SEMrush — cache 12h
───────────────────────────────────────────── */
function buscarSEMrush() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get('semrush_12h');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  try {
    const SEMRUSH_KEY = PropertiesService.getScriptProperties().getProperty('SEMRUSH_KEY');
    if (!SEMRUSH_KEY) return { erro: 'SEMRUSH_KEY não configurada' };

    const domain = 'helenacrm.com';
    const db     = 'br';

    const overviewUrl   = `https://api.semrush.com/?type=domain_ranks&key=${SEMRUSH_KEY}&export_columns=Dn,Rk,Or,Ot,Oc,Ad,At,Ac&domain=${domain}&database=${db}`;
    const overviewResp  = UrlFetchApp.fetch(overviewUrl, { muteHttpExceptions: true });
    const overviewLines = overviewResp.getContentText().trim().split('\n');

    let dominio = { keywords: 0, traffic: 0 };
    if (overviewLines.length >= 2) {
      const vals = overviewLines[1].split(';');
      dominio = { keywords: parseInt(vals[2])||0, traffic: parseInt(vals[3])||0 };
    }

    const kwUrl   = `https://api.semrush.com/?type=domain_organic&key=${SEMRUSH_KEY}&export_columns=Ph,Po,Nq,Cp,Ur&domain=${domain}&database=${db}&display_limit=200&display_sort=nq_desc`;
    const kwResp  = UrlFetchApp.fetch(kwUrl, { muteHttpExceptions: true });
    const kwLines = kwResp.getContentText().trim().split('\n').slice(1);

    const keywords = kwLines.map(l => {
      const p = l.split(';');
      if (p.length < 5) return null;
      return { kw: p[0], posicao: parseInt(p[1])||99, volume: parseInt(p[2])||0, cpc: parseFloat(p[3])||0, url: p[4] };
    }).filter(Boolean);

    const concUrl   = `https://api.semrush.com/?type=domain_organic_organic&key=${SEMRUSH_KEY}&export_columns=Dn,Or,Ot,Np&domain=${domain}&database=${db}&display_limit=10`;
    const concResp  = UrlFetchApp.fetch(concUrl, { muteHttpExceptions: true });
    const concLines = concResp.getContentText().trim().split('\n').slice(1);

    const diretos = ['botconversa.com.br','bolten.io','blackconversa.com.br','growtalks.com.br','kommo.com'];
    const concorrentes = concLines.map(l => {
      const p = l.split(';');
      if (p.length < 3) return null;
      const dom = p[0]; if (dom === domain) return null;
      return {
        domain:   dom,
        keywords: parseInt(p[1])||0,
        traffic:  parseInt(p[2])||0,
        tipo:     diretos.includes(dom) ? 'direto' : 'conteudo',
      };
    }).filter(Boolean);

    const pageMap = {};
    keywords.forEach(k => {
      const url = k.url || '';
      if (!pageMap[url]) pageMap[url] = { url, keywords:[], totalKws:0, melhorPos:999, totalVol:0, defender:[], atacar:[], capturar:[] };
      const pg = pageMap[url];
      pg.keywords.push(k); pg.totalKws++; pg.totalVol += k.volume;
      if (k.posicao < pg.melhorPos) pg.melhorPos = k.posicao;
      if (k.posicao <= 3)       pg.defender.push(k);
      else if (k.posicao <= 20) pg.atacar.push({...k, vol:k.volume, pos:k.posicao});
      else                       pg.capturar.push({...k, vol:k.volume, pos:k.posicao});
    });
    const paginas = Object.values(pageMap).sort((a,b) => b.totalVol - a.totalVol).slice(0, 20);

    const result = { dominio, keywords, concorrentes, paginas, atualizado: new Date().toISOString() };
    cache.put('semrush_12h', JSON.stringify(result), 43200);
    return result;

  } catch (err) {
    Logger.log('Erro SEMrush: ' + err.message);
    return { erro: err.message, dominio:{ keywords:0, traffic:0 }, keywords:[], concorrentes:[], paginas:[] };
  }
}
