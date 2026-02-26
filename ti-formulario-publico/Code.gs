const FORM_APP_TITLE = 'Portal de Chamados TI';
const TARGET_BASE_URL = '';
const PUBLIC_TOKEN = '';
const ENABLE_LOCAL_FALLBACK = true;
const API_BLOCK_CACHE_KEY = 'TI_API_BLOCKED';
const API_BLOCK_CACHE_TTL = 300; // 5 min

const PLANILHA_ID = '';
const FOLDER_PDFS_ID = '';
const TZ = 'America/Sao_Paulo';
const TI_SHEET_NAME = 'Chamados';
const TI_CHAT_SHEET = 'Chamados_Chat';

function doGet() {
  const tpl = HtmlService.createTemplateFromFile('index');
  tpl.appTitle = FORM_APP_TITLE;
  return tpl.evaluate()
    .setTitle(FORM_APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

function getBootstrap() {
  const bootCached = getCache_('FORM_BOOTSTRAP');
  if (bootCached && bootCached.ok) return bootCached;

  const meta = callTargetApi_('ti_public_meta', {});
  if (meta && meta.ok) {
    putCache_('FORM_BOOTSTRAP', meta, 120);
    return meta;
  }

  if (ENABLE_LOCAL_FALLBACK) {
    const local = localBootstrap_();
    putCache_('FORM_BOOTSTRAP', local, 120);
    return local;
  }

  return {
    ok: false,
    msg: (meta && (meta.msg || meta.error)) || 'Falha ao carregar configurações do formulário.'
  };
}

function abrirChamadoPublico(formOrPayload) {
  var input = formOrPayload || {};

  // ── Extract file blob(s) from form submission or legacy payload ──
  var fileBlobs = [];
  var rawAnexos = input.anexos;
  if (rawAnexos) {
    // Single Blob from form submission
    if (typeof rawAnexos.getBytes === 'function') {
      fileBlobs.push(rawAnexos);
    }
    // Array of Blobs (some GAS versions)
    else if (Array.isArray(rawAnexos)) {
      for (var bi = 0; bi < rawAnexos.length; bi++) {
        if (rawAnexos[bi] && typeof rawAnexos[bi].getBytes === 'function') {
          fileBlobs.push(rawAnexos[bi]);
        }
      }
    }
  }
  // Legacy fallback: dataUrl-based files from JSON payload
  if (!fileBlobs.length) {
    var legacyFiles = normalizeFiles_(input.files || []);
    for (var li = 0; li < legacyFiles.length; li++) {
      var lf = legacyFiles[li];
      if (lf && lf.dataUrl) {
        var lb = dataUrlToBlob_(lf.dataUrl, lf.mimeType || 'application/octet-stream');
        if (lb) {
          try { lb.setName(lf.name || 'arquivo'); } catch(_) {}
          fileBlobs.push(lb);
        }
      }
    }
  }

  // ── Resolve setor (handle "OUTROS") ──
  var setor = String(input._setorFinal || input.setor || '').trim();

  var clean = {
    nome: sanitize_(input.nome, 140),
    telefone: sanitize_(input.telefone, 60),
    setor: sanitize_(setor, 140),
    categoria: sanitize_(input.categoria, 120),
    prioridade: sanitize_(input.prioridade, 40) || 'NORMAL',
    descricao: sanitize_(input.descricao, 3000),
    fileBlobs: fileBlobs
  };

  if (!clean.nome || !clean.telefone || !clean.setor || !clean.categoria || !clean.descricao) {
    return { ok: false, msg: 'Preencha todos os campos obrigatórios.' };
  }

  // Try API (text data only – blobs can't be serialized to JSON)
  var apiPayload = {
    nome: clean.nome,
    telefone: clean.telefone,
    setor: clean.setor,
    categoria: clean.categoria,
    prioridade: clean.prioridade,
    descricao: clean.descricao,
    files: []
  };
  var res = callTargetApi_('ti_public_create', apiPayload);

  if (res && res.ok) {
    // Chamado created via API — save files locally to Drive
    if (fileBlobs.length && res.protocolo) {
      try {
        var anexos = saveAnexos_(res.protocolo, fileBlobs);
        res.anexo = anexos.links.length > 0;
        res.anexos = anexos.links.length;
        res.anexoLinks = anexos.links;
        res.anexoNomes = anexos.names;
      } catch(e) { Logger.log('Erro salvando anexos (API path): ' + e); }
    }
    return res;
  }

  if (ENABLE_LOCAL_FALLBACK) {
    return localAbrirChamado_(clean);
  }

  return res || { ok:false, msg:'Falha ao abrir chamado.' };
}


function listarChatPublico(input) {
  input = input || {};
  const protocolo = sanitize_(input.protocolo || input.id, 40);
  if (!protocolo) return { ok: true, items: [] };

  const payload = {
    protocolo,
    since: sanitize_(input.since, 80),
    limit: parseInt(input.limit || 200, 10) || 200,
    leitor: 'SOLICITANTE'
  };

  const res = callTargetApi_('ti_public_chat_list', payload);
  if (res && res.ok) return res;

  if (ENABLE_LOCAL_FALLBACK) {
    return localListarChat_(payload);
  }

  return res || { ok:true, items: [] };
}

function enviarChatPublico(input) {
  input = input || {};

  const protocolo = sanitize_(input.protocolo || input.id, 40);
  const autorNome = sanitize_(input.autorNome || input.nome, 120);
  const mensagem = sanitize_(input.mensagem || input.msg || input.texto, 2000);

  if (!protocolo || !mensagem) {
    return { ok: false, msg: 'Protocolo e mensagem são obrigatórios.' };
  }

  const payload = {
    protocolo,
    autorTipo: 'SOLICITANTE',
    autorNome: autorNome || 'Solicitante',
    mensagem,
    canal: 'FORM_PUBLICO'
  };

  const res = callTargetApi_('ti_public_chat_send', payload);
  if (res && res.ok) return res;

  if (ENABLE_LOCAL_FALLBACK) {
    return localEnviarChat_(payload);
  }

  return res || { ok:false, msg:'Falha ao enviar mensagem.' };
}

function callTargetApi_(api, payload) {
  payload = payload || {};
  if (isApiBlockedCached_()) {
    return { ok: false, msg: 'API principal indisponível (cache local).' };
  }

  const base = String(TARGET_BASE_URL || '').trim();

  if (!base || base.indexOf('COLE_A_URL') >= 0) {
    return { ok: false, msg: 'Configure TARGET_BASE_URL no Code.gs do formulário público.' };
  }

  const finalPayload = Object.assign({}, payload);
  if (PUBLIC_TOKEN) finalPayload.token = PUBLIC_TOKEN;

  const url = base + (base.includes('?') ? '&' : '?') + 'api=' + encodeURIComponent(api);

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      payload: JSON.stringify(finalPayload),
      followRedirects: true,
      muteHttpExceptions: true
    });

    const code = Number(res.getResponseCode() || 0);
    const text = res.getContentText() || '{}';
    if (code >= 300) {
      markApiBlocked_();
      return { ok: false, msg: 'API principal indisponível (HTTP ' + code + ').' };
    }

    if (/^\s*<!doctype|^\s*<html/i.test(text)) {
      markApiBlocked_();
      return { ok: false, msg: 'API principal retornou HTML (provável bloqueio de acesso no deploy).' };
    }

    const data = JSON.parse(text);
    if (!data || data.ok === false) {
      return data || { ok: false, msg: 'Falha na API de destino.' };
    }
    return data;
  } catch (e) {
    markApiBlocked_();
    return { ok: false, msg: 'Erro ao comunicar com API TI: ' + e.message };
  }
}

function sanitize_(value, maxLen) {
  const lim = Math.max(1, parseInt(maxLen || 300, 10));
  return String(value || '').replace(/\s+/g, ' ').trim().slice(0, lim);
}

function normalizeFiles_(files) {
  if (!Array.isArray(files)) return [];

  const out = [];
  const limit = Math.min(files.length, 10);
  for (let i = 0; i < limit; i++) {
    const f = files[i] || {};
    const dataUrl = String(f.dataUrl || '').trim();
    if (!dataUrl) continue;
    out.push({
      name: sanitize_(f.name || ('arquivo_' + (i + 1)), 180),
      mimeType: sanitize_(f.mimeType || '', 120),
      dataUrl: dataUrl
    });
  }
  return out;
}

function localBootstrap_() {
  return {
    ok: true,
    tokenRequired: false,
    prioridades: ['NORMAL', 'ALTA', 'URGENTE'],
    categorias: ['Hardware', 'Software', 'Rede', 'Impressora', 'Sistema', 'Acesso', 'Outro'],
    setores: listarSetoresLocais_(),
    statusInicial: 'ABERTO',
    chat: { enabled: true, maxMensagem: 2000 },
    source: 'fallback_local'
  };
}

function localAbrirChamado_(payload) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const sh = ensureSheetWithHeaders_(ss, TI_SHEET_NAME, [
    'Carimbo','Protocolo','Nome','Email','Telefone','Setor/Local',
    'Categoria','Prioridade','Descrição','Status','Responsável',
    'Atualizado em','Obs','Anexo (Link)','Anexo (Nome)','Anexo (Pasta ID)',
    'TI Última Leitura','Solicitante Última Leitura'
  ]);

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] || [];
  const i = idxMap_(headers);

  const now = new Date();
  const protocolo = nextProtocoloLocal_(sh, i['protocolo']);

  const anexos = saveAnexos_(protocolo, payload.fileBlobs || payload.files || []);

  const row = new Array(sh.getLastColumn()).fill('');
  row[i['carimbo']] = now;
  row[i['protocolo']] = protocolo;
  row[i['nome']] = sanitize_(payload.nome, 140);
  row[i['email']] = 'naoinformado@semfas.local';
  row[i['telefone']] = sanitize_(payload.telefone, 60);
  row[i['setor/local']] = sanitize_(payload.setor, 140);
  row[i['categoria']] = sanitize_(payload.categoria, 120);
  row[i['prioridade']] = sanitize_(payload.prioridade, 40) || 'NORMAL';
  row[i['descrição']] = sanitize_(payload.descricao, 3000);
  row[i['status']] = 'ABERTO';
  row[i['responsável']] = '';
  row[i['atualizado em']] = now;
  row[i['obs']] = anexos.links.length ? ('Aberto pelo portal • ' + anexos.links.length + ' anexo(s)') : 'Aberto pelo portal';
  row[i['anexo (link)']] = anexos.links.join('\n');
  row[i['anexo (nome)']] = anexos.names.join('\n');
  if (i['anexo (pasta id)'] >= 0) row[i['anexo (pasta id)']] = anexos.folderId || '';

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { sh.appendRow(row); }
  finally { lock.releaseLock(); }

  return {
    ok: true,
    protocolo,
    anexo: anexos.links.length > 0,
    anexos: anexos.links.length,
    anexoLinks: anexos.links,
    anexoNomes: anexos.names,
    source: 'fallback_local'
  };
}

function localListarChat_(payload) {
  const protocolo = sanitize_(payload.protocolo || payload.id, 40);
  if (!protocolo) return { ok:true, items: [] };

  const since = sanitize_(payload.since, 80);
  const limit = Math.min(Math.max(parseInt(payload.limit || 200, 10) || 200, 1), 1000);
  const leitor = sanitize_(payload.leitor, 20).toUpperCase();

  const ss = SpreadsheetApp.openById(PLANILHA_ID);

  // Ler/gravar timestamps de leitura na planilha Chamados (não mais ScriptProperties)
  const shCham = ensureSheetWithHeaders_(ss, TI_SHEET_NAME, [
    'Carimbo','Protocolo','Nome','Email','Telefone','Setor/Local',
    'Categoria','Prioridade','Descrição','Status','Responsável',
    'Atualizado em','Obs','Anexo (Link)','Anexo (Nome)','Anexo (Pasta ID)',
    'TI Última Leitura','Solicitante Última Leitura'
  ]);
  const hdrsCham = shCham.getRange(1,1,1,shCham.getLastColumn()).getValues()[0] || [];
  const idxCham = idxMap_(hdrsCham);
  const iReadTI  = idxCham['ti última leitura'];
  const iReadSol = idxCham['solicitante última leitura'];
  const iProtCham = idxCham['protocolo'];

  let chamadoRow = -1;
  if (iProtCham >= 0 && shCham.getLastRow() >= 2) {
    const protVals = shCham.getRange(2, iProtCham + 1, shCham.getLastRow() - 1, 1).getValues();
    for (let k = 0; k < protVals.length; k++) {
      if (String(protVals[k][0] || '').trim() === protocolo) { chamadoRow = k + 2; break; }
    }
  }

  if (chamadoRow > 0 && (leitor === 'TI' || leitor === 'SOLICITANTE')) {
    try {
      const col = leitor === 'TI' ? iReadTI : iReadSol;
      if (col >= 0) shCham.getRange(chamadoRow, col + 1).setValue(new Date());
    } catch(_){}
  }

  let readByTI = null, readBySolic = null;
  if (chamadoRow > 0) {
    try {
      if (iReadTI >= 0) {
        const v = shCham.getRange(chamadoRow, iReadTI + 1).getValue();
        if (v instanceof Date) readByTI = v; else if (v) readByTI = new Date(v);
      }
      if (iReadSol >= 0) {
        const v = shCham.getRange(chamadoRow, iReadSol + 1).getValue();
        if (v instanceof Date) readBySolic = v; else if (v) readBySolic = new Date(v);
      }
    } catch(_){}
  }
  const chat = ensureSheetWithHeaders_(ss, TI_CHAT_SHEET, ['Carimbo','Protocolo','Autor Tipo','Autor Nome','Mensagem','Canal']);
  const last = chat.getLastRow();
  if (last < 2) return { ok:true, items: [] };

  const totalDataRows = last - 1;
  const scanRows = since ? Math.min(totalDataRows, 180) : Math.min(totalDataRows, 500);
  const start = last - scanRows + 1;
  const vals = chat.getRange(start,1,scanRows,chat.getLastColumn()).getValues();

  const out = [];
  for (let r=0; r<vals.length; r++) {
    const row = vals[r];
    if (String(row[1] || '').trim() !== protocolo) continue;
    const quandoDate = row[0] instanceof Date ? row[0] : new Date(row[0]);
    const quandoIso = whenIsValid_(quandoDate)
      ? Utilities.formatDate(quandoDate, TZ, "yyyy-MM-dd'T'HH:mm:ss")
      : String(row[0] || '');
    if (since && quandoIso && quandoIso <= since) continue;

    const tipoUp = String(row[2] || '').toUpperCase();
    let lido = false;
    if (tipoUp === 'SOLICITANTE' && readByTI && quandoDate && quandoDate <= readByTI) lido = true;
    if (tipoUp === 'TI' && readBySolic && quandoDate && quandoDate <= readBySolic) lido = true;
    if (tipoUp === 'SISTEMA') lido = true;

    out.push({
      quando: quandoIso,
      protocolo,
      autorTipo: String(row[2] || 'SOLICITANTE'),
      autorNome: String(row[3] || 'Solicitante'),
      mensagem: String(row[4] || ''),
      canal: String(row[5] || 'FORM_PUBLICO'),
      lido
    });
  }

  out.sort((a,b)=> new Date(a.quando).getTime() - new Date(b.quando).getTime());
  return { ok:true, items: out.slice(-limit), source:'fallback_local' };
}

function localEnviarChat_(payload) {
  const protocolo = sanitize_(payload.protocolo || payload.id, 40);
  const mensagem = sanitize_(payload.mensagem || payload.msg || payload.texto, 2000);
  const autorNome = sanitize_(payload.autorNome || payload.nome, 120) || 'Solicitante';
  if (!protocolo || !mensagem) return { ok:false, msg:'Protocolo e mensagem são obrigatórios.' };

  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const chat = ensureSheetWithHeaders_(ss, TI_CHAT_SHEET, ['Carimbo','Protocolo','Autor Tipo','Autor Nome','Mensagem','Canal']);
  const now = new Date();
  const whenIso = Utilities.formatDate(now, TZ, "yyyy-MM-dd'T'HH:mm:ss");

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    chat.appendRow([now, protocolo, 'SOLICITANTE', autorNome, mensagem, 'FORM_PUBLICO']);
  } finally {
    lock.releaseLock();
  }

  const item = {
    quando: whenIso,
    protocolo,
    autorTipo: 'SOLICITANTE',
    autorNome,
    mensagem,
    canal: 'FORM_PUBLICO'
  };

  return { ok:true, item, source:'fallback_local' };
}

function listarSetoresLocais_() {
  const cached = getCache_('TI_SETOR_LIST');
  if (Array.isArray(cached) && cached.length) return cached;

  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const sh = ss.getSheetByName('Login');
  if (!sh) return ['SEDE'];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return ['SEDE'];

  const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0] || '').trim()).filter(Boolean);
  const map = new Map();
  vals.forEach(v => map.set(v.toLowerCase(), v));
  const setores = Array.from(map.values()).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
  const finalSetores = setores.length ? setores : ['SEDE'];
  putCache_('TI_SETOR_LIST', finalSetores, 300);
  return finalSetores;
}

function saveAnexos_(protocolo, files) {
  const out = { links: [], names: [], folderId: '' };
  if (!Array.isArray(files) || !files.length) return out;

  const root = DriveApp.getFolderById(FOLDER_PDFS_ID);
  const folderName = sanitize_(protocolo || 'SEM_PROTOCOLO', 80).replace(/[\\\/:*?"<>|]+/g, '-');
  const folderIt = root.getFoldersByName(folderName);
  const folder = folderIt.hasNext() ? folderIt.next() : root.createFolder(folderName);
  out.folderId = String(folder.getId() || '').trim();

  const limit = Math.min(files.length, 10);
  for (let i = 0; i < limit; i++) {
    const f = files[i];
    if (!f) continue;
    try {
      let blob = null;
      let name = '';

      // ── Native Blob (from form submission) ──
      if (typeof f.getBytes === 'function') {
        blob = f;
        try { name = f.getName() || ''; } catch(_) {}
      }
      // ── Legacy: {dataUrl, name, mimeType} object ──
      else if (f.dataUrl) {
        const dataUrl = String(f.dataUrl).trim();
        if (!dataUrl) continue;
        blob = dataUrlToBlob_(dataUrl, String(f.mimeType || 'application/octet-stream'));
        name = f.name || '';
      }

      if (!blob) continue;
      if (!name) name = 'arquivo_' + (i + 1);
      name = sanitize_(name, 180).replace(/[^a-zA-Z0-9._\- ]/g, '_');

      const file = folder.createFile(blob).setName(name);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_) {}
      out.links.push('https://drive.google.com/file/d/' + file.getId() + '/view');
      out.names.push(file.getName());
    } catch (e) {
      Logger.log('saveAnexos_ erro arquivo ' + i + ': ' + (e.message || e));
    }
  }
  return out;
}

function dataUrlToBlob_(dataUrl, mimeType) {
  if (!dataUrl) return null;
  const m = String(dataUrl).match(/^data:([^;]+);base64,(.*)$/);
  if (!m) return null;
  const type = m[1] || mimeType || 'application/octet-stream';
  const bytes = Utilities.base64Decode(m[2]);
  return Utilities.newBlob(bytes, type);
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}

function idxMap_(headers) {
  const map = {};
  headers.forEach((h, i) => map[String(h || '').toLowerCase().trim()] = i);
  return map;
}

function nextProtocoloLocal_(sh, idxProtocolo) {
  const year = new Date().getFullYear();
  if (idxProtocolo < 0 || sh.getLastRow() < 2) return 'TI-' + year + '-000001';

  const vals = sh.getRange(2, idxProtocolo + 1, sh.getLastRow() - 1, 1).getValues();
  let max = 0;
  const re = new RegExp('^TI-' + year + '-(\\d+)$');
  vals.forEach(r => {
    const s = String(r[0] || '').trim();
    const m = s.match(re);
    if (m) {
      const n = parseInt(m[1], 10) || 0;
      if (n > max) max = n;
    }
  });
  return 'TI-' + year + '-' + String(max + 1).padStart(6, '0');
}

function whenIsValid_(d) {
  return d instanceof Date && !isNaN(d.getTime());
}

function getCache_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const s = cache.get(String(key || ''));
    return s ? JSON.parse(s) : null;
  } catch (_) {
    return null;
  }
}

function putCache_(key, value, ttl) {
  try {
    const cache = CacheService.getScriptCache();
    cache.put(String(key || ''), JSON.stringify(value), Number(ttl || 60));
  } catch (_) {}
}

function markApiBlocked_() {
  putCache_(API_BLOCK_CACHE_KEY, { blocked: true, at: new Date().toISOString() }, API_BLOCK_CACHE_TTL);
}

function isApiBlockedCached_() {
  const flag = getCache_(API_BLOCK_CACHE_KEY);
  return !!(flag && flag.blocked);
}
