const GSC_SITE_URL    = 'sc-domain:helenacrm.com';
const GA4_PROPERTY_ID = '505588648';
const DIAS            = 30;
const SHEET_ID        = '1Ord5aH3rbLFtWmx6BuBoXtM47WfHhi3BB-5gE_EReRo';
const LEADS_SHEET_GID = 1698310909;

/* ─────────────────────────────────────────────
   CORS + roteamento principal
───────────────────────────────────────────── */
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  // === IA de Estratégia — ?estrategia=1 (roda nova) ou ?estrategia=cache (lê a salva) ===
  if (e && e.parameter && e.parameter.estrategia) {
    const _tk = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
    if (!_tk || (e.parameter.token || '') !== _tk) {
      output.setContent(JSON.stringify({ erro: 'Token inválido.' })); return output;
    }
    try {
      const _r = (e.parameter.estrategia === 'cache')
        ? lerEstrategiaSalva()
        : gerarEstrategia(e.parameter.dias ? parseInt(e.parameter.dias) : 90);
      output.setContent(JSON.stringify(_r));
    } catch (err) {
      output.setContent(JSON.stringify({ erro: String(err) }));
    }
    return output;
  }

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
      leads:      buscarLeads(deStr, ateStr),
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
   LEADS — planilha Helena CRM  (classificação v3 — jul/2026)
   Data de ENTRADA = Data de Criação da Conversa (fallback: Data de Criação do Card).
   Volume Gerado no Mês = contatos únicos/mês SEM CS e SEM Vagas (mantém Lead Falso, Helena Talks, Eventos).
   Leads Válidos = Volume − Lead Falso  (base de Tipo / Jornada / Origem).
   Dedup: nome normalizado (minúsculo, sem acento) + mês da Conversa.
   Funil (etapa única): prioridade Ganho / Perdido > SQL > MQL > Lead.
   Origem (5 grupos): Pago / Ads · Novos Canais (Helena Talks/Eventos) · Orgânico · Sem origem · Não classificada.
═══════════════════════════════════════════════════════════════════ */
function buscarLeads(deStr, ateStr) {
  // recorte por DATA exata (Data de Criação da Conversa) — day-level
  const ini = deStr  ? new Date(deStr  + 'T00:00:00') : null;
  const fim = ateStr ? new Date(ateStr + 'T23:59:59') : null;
  const sfx = '_' + (deStr || 'ini') + '_' + (ateStr || 'fim');   // cache por período (expira sozinho)

  const cache = CacheService.getScriptCache();
  const c0 = cache.get('leads_meta'        + sfx);
  const c1 = cache.get('leads_porMes'      + sfx);
  const c2 = cache.get('leads_porCanalMes' + sfx);
  const c3 = cache.get('leads_porCanal'    + sfx);
  const c4 = cache.get('leads_porConteudo' + sfx);
  const c5 = cache.get('leads_recentes'    + sfx);
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
      return { total:0, cf:0, wl:0, porMes:[], porCanalMes:[], porCanal:[], porConteudo:[], recentes:[], conversoesPorMes:[] };
    }

    const header = data[0].map(h => String(h).toLowerCase().trim());
    const col = name => header.indexOf(name);
    const iDataConversa = col('data de criação da conversa');  // data principal de ENTRADA do contato
    const iDataCard     = col('data de criação do card');       // fallback se não houver a da conversa
    const iNome      = col('nome do contato');
    const iCanal     = col('origem');
    const iEtiquetas = col('etiquetas');
    const iStatus    = col('etapa no crm');
    const iFunil     = col('funil/painel do crm');
    const iAnuncio   = col('anúncio');
    const iCampanha  = col('campanha');
    const iMov       = col('data de atualização do card');  // última movimentação (p/ conversões por fluxo)

    const STAGE_RANK = { lead:1, mql:2, sql:3, perdido:4, ganho:5 };
    const GRUPO_RANK = { sem_origem:1, nao_classificada:2, organico:3, pago:4, novos_canais:5 };
    const labelGrupo = g => g === 'sem_origem'       ? 'Sem origem de conversão'
                          : g === 'pago'             ? 'Pago / Ads'
                          : g === 'organico'         ? 'Orgânico identificado'
                          : g === 'novos_canais'     ? 'Novos Canais'
                          :                            'Origem rastreada não classificada';

    // ── classificação de origem em 4 baldes (Pago / Orgânico / Sem origem / Não classificada) ──
    const PAGAS_ORIGEM    = ['METAADS','FACEBOOK','INSTAGRAM','IG','META','AD','GOOGLEADS'];
    const ORGANICA_ORIGEM = ['WIDGET','LINKTREE','CHATGPT.COM','SITE'];
    const NOVOS_CANAIS    = ['helena talks','evento'];   // Novos Canais (válidos): Helena Talks / Eventos / similares
    const ORGANICA_CONTEUDO = [
      'HELENACRM | PLATAFORMA DE CRM PARA WHATSAPP','HELENACRM / PLATAFORMA',
      'CENTRAL DE ATENDIMENTO E CRM WHITE LABEL PARA WHATSAPP','CRM WHITE LABEL PARA WHATSAPP',
      'TESTE GRÁTIS - HELENACRM','QUERO CONHECER A PARCERIA WHITE LABEL DA HELENACRM',
      'PLANOS','DEMONSTRAÇÃO HELENA','MANUAL DA API OFICIAL','CRM INTEGRADO COM WHATSAPP',
      // adicionados após revisão do Arthur (jun/2026) — eram "não classificada", confirmados orgânicos:
      'QUERO RECEBER MEU LINK DO TESTE GRÁTIS','AGENTES DE IA PARA WHATSAPP','EZEN',
    ];
    const classificarOrigem = (origemR, campanhaR, anuncioR, etqR) => {
      // Novos Canais (Helena Talks / Eventos) — em origem, campanha, anúncio OU etiqueta; ANTES de "sem origem"
      const blobNC = ((origemR||'')+' '+(campanhaR||'')+' '+(anuncioR||'')+' '+(etqR||'')).toLowerCase();
      if (NOVOS_CANAIS.some(t => blobNC.indexOf(t) >= 0)) return 'novos_canais';
      const o=(origemR||'').toUpperCase().trim();
      const c=(campanhaR||'').toUpperCase().trim();
      const a=(anuncioR||'').toUpperCase().trim();
      if (!o && !c && !a) return 'sem_origem';                       // nada de origem/campanha/anúncio
      const ca = c + ' ' + a;                                        // padrões de Ads em campanha + anúncio
      const campanhaPaga =
        /(^|[^A-Z])WL2?($|[^A-Z])/.test(ca) || /WL\s*#/.test(ca) || /WL-/.test(ca) || /AG\s*#/.test(ca) ||
        ca.indexOf('SITE_SM')>=0 || ca.indexOf('SITE_INT')>=0 || ca.indexOf('LP_SM')>=0 ||
        ca.indexOf('WL_IG')>=0 || ca.indexOf('IG_SD')>=0 || ca.indexOf('IG_CAP')>=0 || ca.indexOf('IG_BR')>=0 ||
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
      // Data de ENTRADA = Data de Criação da Conversa (principal); fallback p/ Data de Criação do Card
      const rawData = (iDataConversa >= 0 && row[iDataConversa]) ? row[iDataConversa]
                    : (iDataCard     >= 0 ? row[iDataCard] : null);
      if (!rawData) continue;
      const d = rawData instanceof Date ? rawData : new Date(rawData);
      if (isNaN(d.getTime())) continue;
      // NÃO filtra por data aqui — cada lente (criação/movimentação) filtra na sua própria data.
      let dMov = iMov >= 0 ? row[iMov] : null;        // última movimentação do card
      dMov = dMov instanceof Date ? dMov : (dMov ? new Date(dMov) : d);
      if (isNaN(dMov.getTime())) dMov = d;

      const nome      = iNome      >= 0 ? String(row[iNome]      || '').trim() : '';
      const origem    = iCanal     >= 0 ? String(row[iCanal]     || '').trim() : '';
      const etq       = iEtiquetas >= 0 ? String(row[iEtiquetas] || '').toLowerCase().trim() : '';
      const statusRaw = iStatus    >= 0 ? String(row[iStatus]    || '').trim() : '';
      const etapa     = statusRaw.toLowerCase();
      const funil     = iFunil     >= 0 ? String(row[iFunil]     || '').toLowerCase().trim() : '';
      const anuncio   = iAnuncio   >= 0 ? String(row[iAnuncio]   || '').trim() : '';
      const campanha  = iCampanha  >= 0 ? String(row[iCampanha]  || '').trim() : '';

      // ── exclusões da BASE (Volume Gerado): SÓ CS e Vagas ──
      //    Lead Falso NÃO sai aqui — sai na camada de Leads Válidos (flag abaixo).
      //    Helena Talks / Eventos PERMANECEM (são "Novos Canais", contatos válidos).
      if (/\bcs\s*-/.test(etq)) continue;  // CS - <responsável> (já é cliente)
      // vaga: olhar etiqueta + origem + campanha + anúncio (VAGA / VAGAS / VAGAS_BH)
      const blobVaga = (etq + ' ' + origem + ' ' + campanha + ' ' + anuncio).toLowerCase();
      if (/\bvaga/.test(blobVaga)) continue;
      const leadFalso = etq.includes('lead falso');  // marca: sai só de Leads Válidos

      // ── tipo de cliente: TODO contato válido conta.
      //    WL só com 4'LEAD P WHITE LABEL e SEM etiqueta de Cliente Final; o resto é Cliente Final. ──
      const temWlTag = etq.includes("4'lead p white label");
      const temCfTag = etq.includes('cliente final');
      const tipo = (temWlTag && !temCfTag) ? 'white_label' : 'cliente_final';

      // etapa do funil — ÚNICA e final, prioridade Ganho / Perdido > SQL > MQL > Lead
      // Ganho e Perdido vêm do campo "Etapa no CRM" (final do funil, excludentes entre si).
      // Se o contato foi Perdido, conta como Perdido mesmo que tenha chegado a SQL/MQL antes.
      // MQL = SÓ pela etiqueta 7'MQL. Estar no "Funil Parceiros (Qualificação)" NÃO basta,
      // porque esse funil também contém os Lead Frio (que devem ficar como Lead).
      const isGanho   = etapa.includes('ganho');
      const isPerdido = etapa.includes('perdido') && !isGanho;   // "Etapa no CRM" = 7. Perdido
      const isSql     = etq.includes("8'sql") || funil.includes('(sql)') || funil.includes('comercial');
      const isMql     = etq.includes("7'mql");
      const stage     = isGanho ? 'ganho' : isPerdido ? 'perdido' : isSql ? 'sql' : isMql ? 'mql' : 'lead';
      const perdido   = stage === 'perdido';

      // origem — 5 baldes (Pago / Novos Canais / Orgânico / Sem origem / Não classificada)
      const grupo = classificarOrigem(origem, campanha, anuncio, etq);

      const mkOf  = dd => String(dd.getFullYear()) + String(dd.getMonth() + 1).padStart(2, '0');
      const lblOf = dd => dd.toLocaleDateString('pt-BR', { month: 'short', year: '2-digit' }).replace('.', '').replace(' de ', '/');

      registros.push({
        nomeKey: nome.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim(),
        nome, tipo, stage, perdido, leadFalso, grupo,
        conteudo: campanha || anuncio || '',
        anuncio, campanha, statusRaw,
        d, mesKey: mkOf(d), mes: lblOf(d),                  // criação (Data de Criação da Conversa)
        dMov, mesKeyMov: mkOf(dMov), mesMov: lblOf(dMov),   // movimentação
      });
    }

    /* ── dedup reutilizável: agrupa por (mês + nome), mantém a etapa mais avançada ── */
    const dedupPor = (regs, campoMes) => {
      const map = {}; let anon = 0;
      regs.forEach(r => {
        const mk = r[campoMes];
        const key = r.nomeKey ? (mk + '|' + r.nomeKey) : (mk + '|__anon' + (anon++));
        const ex = map[key];
        if (!ex) { map[key] = r; return; }
        const hi = STAGE_RANK[r.stage] > STAGE_RANK[ex.stage] ? r : ex;  // etapa mais avançada
        const gw = GRUPO_RANK[r.grupo] >= GRUPO_RANK[ex.grupo] ? r : ex;  // origem mais informativa
        map[key] = {
          nomeKey: ex.nomeKey, nome: hi.nome, tipo: hi.tipo, stage: hi.stage,
          perdido: hi.stage === 'perdido',
          leadFalso: (ex.leadFalso && r.leadFalso),   // só é Lead Falso se TODOS os cards do contato forem
          grupo: gw.grupo, conteudo: gw.conteudo, anuncio: gw.anuncio, campanha: gw.campanha,
          statusRaw: hi.statusRaw,
          d: (ex.d > r.d ? ex.d : r.d), mesKey: ex.mesKey, mes: ex.mes,
          dMov: (ex.dMov > r.dMov ? ex.dMov : r.dMov), mesKeyMov: ex.mesKeyMov, mesMov: ex.mesMov,
        };
      });
      return Object.values(map);
    };

    // LENTE 1 — CRIAÇÃO (volume + CF/WL + funil): filtra por Data de Criação do Card, dedup por mês de criação
    const regCriacao = registros.filter(r => (!ini || r.d >= ini) && (!fim || r.d <= fim));
    const unicos = dedupPor(regCriacao, 'mesKey');

    /* ── 3ª passada: agregações ── */
    const novo = extra => Object.assign({
      total:0, cf:0, wl:0, volume:0, lead_falso:0, lead:0, mql:0, sql:0, ganho:0, perdido:0,
      pago:0, organico:0, sem_origem:0, nao_classificada:0, novos_canais:0,
      cf_lead:0, cf_mql:0, cf_sql:0, cf_ganho:0, cf_perdido:0,
      wl_lead:0, wl_mql:0, wl_sql:0, wl_ganho:0, wl_perdido:0,
    }, extra);
    const acumula = (o, r) => {
      o.total++; if ('count' in o) o.count++;
      o[r.grupo]++; o[r.stage]++;   // stage agora inclui 'perdido' (etapa final excludente)
      if (r.tipo === 'white_label') { o.wl++; o['wl_' + r.stage]++; }
      else                          { o.cf++; o['cf_' + r.stage]++; }
    };

    let total = 0, cf = 0, wl = 0, volume = 0, leadFalsoTot = 0;
    const mesMap = {}, canalMap = {}, canalMesMap = {}, conteudoMap = {};
    unicos.forEach(r => {
      const gl = labelGrupo(r.grupo);
      if (!mesMap[r.mesKey]) mesMap[r.mesKey] = novo({ mes: r.mes, mesKey: r.mesKey });

      // Volume Gerado no Mês: TODOS os contatos únicos (sem CS/Vagas) — inclui Lead Falso, Helena Talks, Eventos
      volume++; mesMap[r.mesKey].volume++;
      if (r.leadFalso) { leadFalsoTot++; mesMap[r.mesKey].lead_falso++; return; }  // Lead Falso não entra em Leads Válidos

      // Leads Válidos (Volume − Lead Falso): base de Tipo / Jornada / Origem
      total++; if (r.tipo === 'white_label') wl++; else cf++;
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

    const recentes = unicos.filter(r => !r.leadFalso).sort((a, b) => b.d - a.d).slice(0, 50).map(r => ({
      data: r.d.toLocaleDateString('pt-BR'), nome: r.nome, canal: labelGrupo(r.grupo),
      tipo: r.tipo, status: r.statusRaw || 'lead', stage: r.stage,
      isMql: r.stage === 'mql', isSql: r.stage === 'sql', isGanho: r.stage === 'ganho', isPerdido: r.perdido,
      conteudo: r.conteudo, anuncio: r.anuncio, campanha: r.campanha, pago: r.grupo === 'pago',
    }));

    /* ── LENTE 2 — MOVIMENTAÇÃO: conversões POR FLUXO (quando viraram MQL/SQL/Ganho) ──
       Conta todo contato em MQL/SQL/Ganho cuja ÚLTIMA MOVIMENTAÇÃO (Data de Atualização do Card)
       caiu no período — independente de quando foi criado. Dedup por (mês de movimentação + nome). ── */
    const regConv = registros.filter(r =>
      !r.leadFalso && (r.stage === 'mql' || r.stage === 'sql' || r.stage === 'ganho') &&
      (!ini || r.dMov >= ini) && (!fim || r.dMov <= fim));
    const convMap = {};
    dedupPor(regConv, 'mesKeyMov').forEach(r => {
      if (!convMap[r.mesKeyMov]) convMap[r.mesKeyMov] = {
        mesKey: r.mesKeyMov, mes: r.mesMov, total:0, mql:0, sql:0, ganho:0, cf:0, wl:0,
        cf_mql:0, cf_sql:0, cf_ganho:0, wl_mql:0, wl_sql:0, wl_ganho:0,
      };
      const o = convMap[r.mesKeyMov];
      o.total++; o[r.stage]++;
      if (r.tipo === 'white_label') { o.wl++; o['wl_' + r.stage]++; }
      else                          { o.cf++; o['cf_' + r.stage]++; }
    });
    const conversoesPorMes = Object.values(convMap).sort((a, b) => a.mesKey > b.mesKey ? 1 : -1);

    const result = {
      total, cf, wl, volume, leadFalso: leadFalsoTot,
      porMes:      Object.values(mesMap).sort((a, b) => a.mesKey > b.mesKey ? 1 : -1),
      porCanalMes: Object.values(canalMesMap),
      porCanal:    Object.values(canalMap).sort((a, b) => b.total - a.total),
      porConteudo: Object.values(conteudoMap).sort((a, b) => b.total - a.total).slice(0, 300),
      recentes,
      conversoesPorMes,
    };

    const ttl = 1800;  // 30 min; chaves por período expiram sozinhas (não precisa limparCache p/ leads)
    const guardar = (k, v) => { try { cache.put(k + sfx, JSON.stringify(v), ttl); } catch (e) {} }; // best-effort (ignora blocos > 100KB)
    guardar('leads_meta',        { total: result.total, cf: result.cf, wl: result.wl, volume: result.volume, leadFalso: result.leadFalso, conversoesPorMes: result.conversoesPorMes });
    guardar('leads_porMes',      result.porMes);
    guardar('leads_porCanalMes', result.porCanalMes);
    guardar('leads_porCanal',    result.porCanal);
    guardar('leads_porConteudo', result.porConteudo);
    guardar('leads_recentes',    result.recentes);
    return result;

  } catch (err) {
    Logger.log('Erro buscarLeads: ' + err.message);
    return { erro: err.message, total:0, cf:0, wl:0, porMes:[], porCanalMes:[], porCanal:[], porConteudo:[], recentes:[], conversoesPorMes:[] };
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

/* ─────────────────────────────────────────────
   IA DE ESTRATÉGIA (Claude / Anthropic)
   Analisa GSC + GA4 + SEMrush e devolve o relatório em Markdown.
   Requer a propriedade de script ANTHROPIC_KEY.
───────────────────────────────────────────── */
function buscarEstrategiaIA(dados) {
  const KEY = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  if (!KEY) return '⚠️ Configure a propriedade de script ANTHROPIC_KEY com a chave da Anthropic.';

  const MODELO = 'claude-opus-4-8'; // p/ gastar ~40% menos, troque por 'claude-sonnet-5'

  const gsc = dados.gsc || {}, ga4 = dados.ga4 || {}, sm = dados.semrush || {};
  const kws = gsc.keywords || [];
  const porCli = kws.slice().sort(function(a,b){ return (b.cliques||0)-(a.cliques||0); }).slice(0,30);
  const porImp = kws.slice().sort(function(a,b){ return (b.impressoes||0)-(a.impressoes||0); }).slice(0,20);
  const mini = function(k){ return { q:k.query, cli:k.cliques, imp:k.impressoes, ctr:k.ctr, pos:k.posicao }; };

  // compara com o período igual anterior (páginas/queries que subiram/caíram + o pico) — função lá embaixo
  var dias = (gsc.historico || []).length || 90;
  var comparativo = buscarGSCComparativo(gsc, dias);

  const resumo = {
    gsc: {
      totais: gsc.totais,
      historicoDiario: gsc.historico,
      topPorCliques: porCli.map(mini),
      topPorImpressoes: porImp.map(mini),
      paginas: (gsc.paginas || []).slice(0,25),
    },
    comparativoPeriodoAnterior: comparativo,
    ga4: { totais: ga4.totais, canais: ga4.canais, origens: (ga4.origens || []).slice(0,15) },
    semrush: { dominio: sm.dominio, keywords: (sm.keywords || []).slice(0,30), concorrentes: sm.concorrentes },
  };

  const prompt =
    'Você é o Head de Growth e SEO da HelenaCRM — SaaS brasileiro de CRM e atendimento no WhatsApp, com produto White Label para agências/parceiros revenderem com a própria marca. Concorrentes no orgânico: z-api.io, uzapi, w-api, chatsac.\n\n' +
    'Seu papel é ser o BRAÇO DIREITO DE PERFORMANCE ORGÂNICA da equipe: analítico, direto, prático e estratégico. Traga PROATIVAMENTE oportunidades, alertas e insights relevantes mesmo que não tenham sido pedidos — desde que ajudem a performance orgânica da Helena.\n\n' +
    'Analise os dados REAIS abaixo (Google Search Console, Google Analytics 4 e SEMrush) e escreva um relatório executivo, em português do Brasil, em Markdown, conciso e escaneável.\n\n' +
    'Você TEM os dados do período atual E a comparação com o período igual anterior (campo comparativoPeriodoAnterior): queriesEmAlta/queriesEmQueda (variação de cliques vs período anterior), paginasPerdendoPos/paginasGanhandoPos (variação de posição — dPos positivo = MELHOROU, negativo = PIOROU) e pico (dia de maior tráfego + queries desse dia + média diária). USE esses dados.\n\n' +
    'Estrutura (use tabelas quando ajudar; cite SEMPRE os números reais):\n' +
    '## 📍 Diagnóstico rápido (2-3 frases)\n' +
    '## 📊 Números principais (KPIs mais importantes com valor E variação vs período anterior: cliques, impressões, CTR, posição média, sessões orgânicas)\n' +
    '## 🔎 Picos & anomalias (EXPLIQUE o pico de tráfego: qual dia, quais queries dispararam nesse dia vs a média, e HIPÓTESES do porquê; aponte qualquer busca que cresceu fora do normal)\n' +
    '## 📉 Páginas perdendo/ganhando posição (liste as que CAÍRAM e as que SUBIRAM de posição, com números; para cada queda, possíveis causas e o que fazer p/ recuperar)\n' +
    '## 🚨 Alertas & oscilações (quedas anômalas; problema técnico como a mesma página em www E sem-www dividindo autoridade; qualquer coisa estranha)\n' +
    '## 💰 Oportunidades priorizadas (keywords/páginas com muita impressão e pouco clique; posições de borda 5-15 com volume que dá p/ empurrar pro topo; cite números)\n' +
    '## 🔍 Otimizações por página (escolha as páginas mais importantes e diga, POR PÁGINA, o que otimizar: título, meta description, conteúdo, canonical, links internos; esforço Rápido/Médio/Alto)\n' +
    '## ✍️ Lacunas de conteúdo — pautas de blog (sugira posts NOVOS a criar; p/ cada um: tema/título, keyword-alvo com volume e por que vale — baseie em keywords de volume mal posicionadas/ausentes e no que os concorrentes cobrem)\n' +
    '## 🏆 Concorrência\n' +
    '## ✅ Plano de ação priorizado (itens do maior impacto ao menor, marcando esforço: Rápido / Médio / Alto)\n\n' +
    'Regras: baseie TUDO nos números fornecidos (cite cliques, impressões, CTR em %, posição). NÃO invente. CTR está em %; posição é média (menor = melhor, 1 = topo). Priorize por impacto. Seja específico e acionável, sem generalidades. Responda DIRETO com o relatório, sem raciocínio à parte.\n\n' +
    'DADOS:\n' + JSON.stringify(resumo);

  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { 'x-api-key': KEY, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: MODELO,
      max_tokens: 6000,
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  const code = resp.getResponseCode();
  const body = JSON.parse(resp.getContentText());
  if (code !== 200) return '⚠️ Erro da IA (' + code + '): ' + (body && body.error ? body.error.message : resp.getContentText().slice(0,300));
  return (body.content || []).filter(function(b){ return b.type === 'text'; }).map(function(b){ return b.text; }).join('\n');
}

/* Gera a análise, GUARDA (com data) e devolve. Guarda em pedaços p/ caber no limite de 9KB por propriedade. */
function gerarEstrategia(dias) {
  var dados = { gsc: buscarGSC(dias, null, null), ga4: buscarGA4(dias, null, null), semrush: buscarSEMrush() };
  var txt = buscarEstrategiaIA(dados);
  var quando = new Date().toISOString();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('IA_DATA', quando);
  ['IA_TXT_0','IA_TXT_1','IA_TXT_2','IA_TXT_3'].forEach(function(k){ props.deleteProperty(k); });
  var partes = String(txt).match(/[\s\S]{1,8000}/g) || [''];
  var n = Math.min(partes.length, 4);
  for (var i = 0; i < n; i++) props.setProperty('IA_TXT_' + i, partes[i]);
  props.setProperty('IA_PARTES', String(n));
  return { estrategia: txt, atualizado: quando, auto: false };
}

/* Lê a análise salva (sem chamar a IA) — usado quando o painel abre. */
function lerEstrategiaSalva() {
  var props = PropertiesService.getScriptProperties();
  var n = parseInt(props.getProperty('IA_PARTES') || '0', 10);
  if (!n) return { estrategia: '', atualizado: null, auto: true };
  var txt = '';
  for (var i = 0; i < n; i++) txt += (props.getProperty('IA_TXT_' + i) || '');
  return { estrategia: txt, atualizado: props.getProperty('IA_DATA'), auto: true };
}

/* Alvo do gatilho semanal (toda segunda). */
function rodarEstrategiaSemanal() {
  gerarEstrategia(90);
}

/* Rode UMA vez (menu Executar) p/ ligar o gatilho de toda segunda ~7h. */
function instalarGatilhoSemanal() {
  var gs = ScriptApp.getProjectTriggers();
  for (var i = 0; i < gs.length; i++) {
    if (gs[i].getHandlerFunction() === 'rodarEstrategiaSemanal') ScriptApp.deleteTrigger(gs[i]);
  }
  ScriptApp.newTrigger('rodarEstrategiaSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  return 'Gatilho semanal instalado (toda segunda ~7h).';
}

/* Compara o período atual com o ANTERIOR (mesmo tamanho) p/ achar:
   - queries em alta/queda (variação de cliques)
   - páginas perdendo/ganhando posição
   - o dia de PICO e as queries desse dia (p/ explicar o pico)
   Recebe o gsc atual (de buscarGSC) p/ não refazer as buscas do período atual. */
function buscarGSCComparativo(gscAtual, dias) {
  try {
    var token = ScriptApp.getOAuthToken();
    var siteEnc = encodeURIComponent(GSC_SITE_URL);
    var url = 'https://www.googleapis.com/webmasters/v3/sites/' + siteEnc + '/searchAnalytics/query';
    var fmt = function(d){ return d.toISOString().slice(0,10); };
    var opts = function(body){ return { method:'post', contentType:'application/json', headers:{Authorization:'Bearer '+token}, payload:JSON.stringify(body), muteHttpExceptions:true }; };
    var rows = function(ini, fim, dim, lim){
      var r = UrlFetchApp.fetch(url, opts({ startDate:fmt(ini), endDate:fmt(fim), dimensions:[dim], rowLimit:lim }));
      return (JSON.parse(r.getContentText()).rows || []);
    };
    var hoje = new Date();
    var iniAtual = new Date(hoje); iniAtual.setDate(iniAtual.getDate() - dias);
    var fimAnt = new Date(iniAtual); fimAnt.setDate(fimAnt.getDate() - 1);
    var iniAnt = new Date(fimAnt); iniAnt.setDate(iniAnt.getDate() - dias);

    var mapear = function(rs){ var m={}; rs.forEach(function(r){ m[r.keys[0]]={cli:r.clicks, imp:r.impressions, pos:r.position}; }); return m; };
    var qAnt = mapear(rows(iniAnt, fimAnt, 'query', 500));
    var pAnt = mapear(rows(iniAnt, fimAnt, 'page', 100));
    var top = function(arr, cmp, n){ return arr.slice().sort(cmp).slice(0, n); };

    var curKw = gscAtual.keywords || [];
    var qMov = curKw.map(function(k){ var a=qAnt[k.query];
      return { q:k.query, cli:k.cliques, imp:k.impressoes, pos:k.posicao, dCli:a?(k.cliques-a.cli):k.cliques, dPos:a?+(a.pos-k.posicao).toFixed(1):null, novo:!a }; });
    var queriesEmAlta  = top(qMov.filter(function(x){return x.dCli>0;}), function(a,b){return b.dCli-a.dCli;}, 12);
    var queriesEmQueda = top(qMov.filter(function(x){return x.dCli<0;}), function(a,b){return a.dCli-b.dCli;}, 10);

    var curPg = gscAtual.paginas || [];
    var pMov = curPg.map(function(p){ var a=pAnt[p.url];
      return { url:p.url, cli:p.cliques, imp:p.impressoes, pos:p.posicao, dPos:a?+(a.pos-p.posicao).toFixed(1):null, dCli:a?(p.cliques-a.cli):p.cliques, nova:!a }; });
    var paginasPerdendoPos  = top(pMov.filter(function(x){return x.dPos!==null && x.dPos<-0.4;}), function(a,b){return a.dPos-b.dPos;}, 10);
    var paginasGanhandoPos  = top(pMov.filter(function(x){return x.dPos!==null && x.dPos>0.4;}), function(a,b){return b.dPos-a.dPos;}, 10);

    // pico = dia de mais cliques no histórico atual → busca as queries desse dia
    var hist = gscAtual.historico || [], pico = null;
    hist.forEach(function(h){ if(!pico || h.cliques>pico.cliques) pico=h; });
    var picoQueries = [];
    if (pico) {
      var pr = UrlFetchApp.fetch(url, opts({ startDate:pico.data, endDate:pico.data, dimensions:['query'], rowLimit:15 }));
      picoQueries = (JSON.parse(pr.getContentText()).rows || []).map(function(r){ return { q:r.keys[0], cli:r.clicks, imp:r.impressions }; });
    }
    var mediaDia = hist.length ? Math.round(hist.reduce(function(s,h){return s+h.cliques;},0)/hist.length) : 0;

    return {
      periodoAtual: fmt(iniAtual)+' a '+fmt(hoje), periodoAnterior: fmt(iniAnt)+' a '+fmt(fimAnt),
      queriesEmAlta: queriesEmAlta, queriesEmQueda: queriesEmQueda,
      paginasPerdendoPos: paginasPerdendoPos, paginasGanhandoPos: paginasGanhandoPos,
      pico: pico ? { data: pico.data, cliques: pico.cliques, mediaDiaria: mediaDia, topQueries: picoQueries } : null,
    };
  } catch (err) { return { erro: String(err) }; }
}
