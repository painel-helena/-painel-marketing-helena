/* ════════════════════════════════════════════════════════════════
   buscarLeads() — VERSÃO 2 (classificação de leads, jun/2026)
   ----------------------------------------------------------------
   Substitui a função buscarLeads() antiga no Apps Script do painel.
   Regras aplicadas (definidas pela Helena):
     1) Base válida  : só leads WL/CF; exclui CS-, LEAD FALSO, VAGA(S), HELENA TALKS
     2) Dedup        : contato único por mês (chave = Nome do contato),
                       mantendo a etapa mais avançada
     3) Funil        : prioridade Ganho > SQL > MQL > Lead
                       - Ganho  = coluna "Etapa no CRM" contém "ganho"  (6. Ganho)
                       - Perdido= "Etapa no CRM" contém "perdido"        (7. Perdido)
                       - SQL    = etiqueta 8'SQL  OU Funil "Comercial (SQL)"
                       - MQL    = etiqueta 7'MQL  OU Funil "Parceiros (Qualificação)"
                       - Lead   = o resto (inclui 6'Lead Frio)
     4) Origem (3 grupos): Pago / Ads · Orgânico identificado · Sem origem
                       - Sem origem = Origem + Campanha + Anúncio os 3 vazios
   Campos antigos do JSON foram mantidos (compatibilidade do painel) e
   novos foram adicionados: ganho, perdido, sem_origem, cf_ganho, wl_ganho.
   Após colar + salvar: rode limparCache() uma vez e republique a implantação.
   ════════════════════════════════════════════════════════════════ */
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
    const GRUPO_RANK = { sem_origem:1, organico:2, pago:3 };
    const labelGrupo = g => g === 'sem_origem' ? 'Sem origem de conversão'
                          : g === 'pago'       ? 'Pago / Ads'
                          :                      'Orgânico identificado';

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

      // tipo: precisa de etiqueta de lead reconhecida (WL ou CF)
      const isWl = etq.includes("4'lead p white label");
      const isCf = etq.includes('lead cliente final');
      if (!isWl && !isCf) continue;

      // exclusões da base válida
      if (etq.includes('lead falso'))   continue;  // LEAD FALSO
      if (/\bcs\s*-/.test(etq))         continue;  // CS - <responsável>
      if (etq.includes('helena talks')) continue;  // HELENA TALKS / SP
      if (/\bvagas?\b/.test(etq))       continue;  // VAGA / VAGAS

      const tipo = isWl ? 'white_label' : 'cliente_final';

      // etapa do funil — prioridade Ganho > SQL > MQL > Lead
      const isGanho = etapa.includes('ganho');
      const isSql   = etq.includes("8'sql") || funil.includes('(sql)')     || funil.includes('comercial');
      const isMql   = etq.includes("7'mql") || funil.includes('qualifica') || funil.includes('parceiros');
      const stage   = isGanho ? 'ganho' : isSql ? 'sql' : isMql ? 'mql' : 'lead';
      const perdido = etapa.includes('perdido') && !isGanho;

      // origem — 3 grupos
      const c = campanha;
      const isGoogle    = /google/i.test(origem);
      const isInstagram = /instagram|ig\b/i.test(origem);
      const isOrgOrigin = /linktree|chatgpt|weidget|widget/i.test(origem);
      const isMetaCampaign = /^\d{8,}$/.test(c) || /_SM_|_IG_|_LP_|_INT\./i.test(c) || /^IA\/WL/i.test(c);
      const isGooglePaid   = anuncio.length > 0 || /^WL\d*$/i.test(c) || /QUERO CONHECER A PARCERIA WHITE LABEL/i.test(c);
      const isInstagramOrganic = isInstagram && (c.length === 0 || /^PLANOS$/i.test(c));
      const isAd = !isOrgOrigin && !isInstagramOrganic && (isMetaCampaign || (isGoogle && isGooglePaid));

      const temOrigem = origem !== '' || campanha !== '' || anuncio !== '';
      const grupo = !temOrigem ? 'sem_origem' : (isAd ? 'pago' : 'organico');

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
      pago:0, organico:0, sem_origem:0,
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
