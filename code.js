
const PLANILHA_ID    = '1_jNjcdd27uAqJ-TpfMFvPIAYVAmhZKK6Piyc_DOFjhY';
const ID_TEMPLATE    = '1H9ptzXaKu0I9Ngg48e3lhMVTUzKQa863fkgTU4xdEjE';
const FOLDER_PDFS_ID = '13fA080P3J-9qxWNgcRUvCVOt8sTwYDcY';
const TZ             = 'America/Sao_Paulo';

// RMA - pastas de templates e saida (configure com IDs do Drive)
const RMA_TEMPLATES_FOLDER_ID = '1g3vu98thBzPsY_GoYqBbuxP2tcbZC-5m';
const RMA_OUTPUT_FOLDER_ID = '1g3vu98thBzPsY_GoYqBbuxP2tcbZC-5m';

const RESPOSTAS_SHEET_NAME = 'Respostas';
const CPF_HEADER           = 'CPF';

// Cabeçalhos compartilhados
const BAIXAS_SHEET_NAME = 'Baixas';
const ENTREGUE_HEADER   = 'Entregue';
const ENTREGUE_EM_HDR   = 'Entregue em';
const ENTREGUE_POR_HDR  = 'Entregue por';
const ENTREGUE_UNID_HDR = 'Unidade de Entrega';
const ENTREGUE_OBS_HDR  = 'Obs Entrega';
const PROTOCOLO_HDR     = 'Protocolo';

// E-mail
const OUTBOX_SHEET_NAME = 'FilaEmails';
const OUTBOX_MAX_TRIES  = 5;
const COPIA_EMAIL       = ''; // opcional

/************************
 * ROUTER / VIEWS ✅ (COM TEMPLATE + baseUrl)
 ************************/
function __getBaseUrl_() {
  try {
    const url = ScriptApp.getService().getUrl();
    return url || '';
  } catch (e) {
    return '';
  }
}


function __tpl_(file, extra) {
  const t = HtmlService.createTemplateFromFile(file);
  t.baseUrl = __getBaseUrl_();
  if (extra && typeof extra === 'object') {
    Object.keys(extra).forEach(k => (t[k] = extra[k]));
  }
  return t;
}

function __page_(file, extra) {
  return __tpl_(file, extra).evaluate()
    .setTitle('Sistema SEMFAS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function __content_(file, extra) {
  return __tpl_(file, extra).evaluate().getContent();
}

function doGet(e) {
  try {
    ensureOutboxSheet_();
    ensureQueueTrigger_();
    ensureTiSheets_(); // ✅ garante Chamados / Historico do TI
  } catch (err) {
    Logger.log('[BOOT] ' + (err && err.message ? err.message : err));
  }

  // ✅ API MODE (?api=...)
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    if (p && p.api) return apiDispatch_(e);
  } catch (apiErr) {
    return __json_({ ok:false, error:'Falha no modo API', details:String(apiErr && apiErr.message ? apiErr.message : apiErr) });
  }

  // ✅ VIEW MODE (?view=...)
  const viewRaw = (e && e.parameter && e.parameter.view) ? String(e.parameter.view) : 'login';
  const view = (viewRaw || '').trim().toLowerCase();

  const map = {
    login:'login',
    hub:'hub',
    hub_beneficio:'hub_beneficio',
    hub_beneficio_eventual:'hub_beneficio_eventual',
    hub_beneficio_rma:'hub_beneficio_rma',
    rma:'rma',
    vigilancia:'vigilancia',
    formulario:'formulario',
    baixa:'baixa',
    central:'central',
    admin:'admin',
    analista:'analista',
    ti:'ti',

    // ✅ caso seu TI use rotas separadas:
    ti_tv:'ti',
    ti_relatorios:'ti',
    ti_maquinas:'ti'
  };

  const file = map[view] || 'login';

  try {
    return __page_(file, { view });
  } catch (err) {
    const html = `
      <html><body style="font-family:Arial;padding:24px">
        <h2>View não encontrada: <code>${file}.html</code></h2>
        <p>Voltando ao login…</p>
        <pre style="background:#f6f8fa;border:1px solid #e5e7eb;border-radius:8px;padding:12px;white-space:pre-wrap">${err.message}</pre>
        <script>setTimeout(()=>location.search='?view=login', 1000)</script>
      </body></html>`;
    return HtmlService.createHtmlOutput(html).setTitle('Sistema SEMFAS');
  }
}
function doPost(e){
  const p = (e && e.parameter) ? e.parameter : {};
  if (p && p.api) return apiDispatch_(e);
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

function __json_(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function apiDispatch_(e){
  const p = (e && e.parameter) ? e.parameter : {};
  const api = String(p.api || '').trim().toLowerCase();

  // body pode vir via POST JSON ou via ?payload={}
  let body = {};
  try {
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (p.payload) {
      body = JSON.parse(p.payload);
    }
  } catch (_){ body = {}; }

  // merge leve (querystring também serve como input)
  const input = Object.assign({}, p, body);

  try {
    switch(api){
      // ===== TI =====
      case 'ti_ping':
        return __json_(ti_ping(input));

      case 'ti_boot':
      case 'ti_init':
        return __json_(ti_boot(input));

      case 'ti_list':
      case 'ti_listar':
      case 'ti_chamados':
        return __json_(ti_listarChamados(input));

      case 'ti_get':
      case 'ti_abrir':
      case 'ti_detalhe':
        return __json_(ti_obterChamado(input));

      case 'ti_drive_anexos':
      case 'ti_anexos_drive':
      case 'ti_attachments':
        return __json_(ti_listarAnexosDrive(input));

      case 'ti_create':
      case 'ti_novo':
        return __json_(ti_criarChamado(input));

      case 'ti_update':
      case 'ti_atualizar':
        return __json_(ti_atualizarChamado(input));

      case 'ti_hist':
      case 'ti_historico':
        return __json_(ti_historicoChamado(input));

      case 'ti_report':
      case 'ti_relatorios':
        return __json_(ti_relatorios(input));

      case 'ti_tecnicos':
        return __json_(ti_listarTecnicos());

      case 'ti_set_tecnico':
        return __json_(ti_setTecnicoAtual(input));

      case 'ti_public_meta':
      case 'ti_meta_publico':
        return __json_(ti_public_meta(input));

      case 'ti_public_create':
      case 'ti_criar_publico':
      case 'ti_open_public':
        return __json_(ti_public_create(input));

      case 'ti_chat_send':
      case 'ti_chat_post':
        return __json_(ti_chat_enviar(input));

      case 'ti_chat_list':
      case 'ti_chat_messages':
        return __json_(ti_chat_listar(input));

      case 'ti_public_chat_send':
      case 'ti_public_chat_post':
        return __json_(ti_public_chat_send(input));

      case 'ti_public_chat_list':
      case 'ti_public_chat_messages':
        return __json_(ti_public_chat_list(input));

      // ===== MÁQUINAS =====
      case 'maq_list':
      case 'maq_listar':
        return __json_(maq_listar(input));

      case 'maq_get':
      case 'maq_obter':
        return __json_(maq_obter(input));

      case 'maq_create':
      case 'maq_criar':
        return __json_(maq_criar(input));

      case 'maq_update':
      case 'maq_atualizar':
        return __json_(maq_atualizar(input));

      case 'maq_delete':
      case 'maq_excluir':
        return __json_(maq_excluir(input));

      case 'maq_move':
      case 'maq_mover':
        return __json_(maq_moverLocal(input));

      case 'maq_historico':
        return __json_(maq_historico(input));

      case 'maq_stats':
      case 'maq_estatisticas':
        return __json_(maq_estatisticasPorLocal());

      // ===== fallback =====
      default:
        return __json_({ ok:false, error:'API não reconhecida', api: api });
    }
  } catch (err) {
    return __json_({ ok:false, error:String(err && err.message ? err.message : err), api });
  }
}


function renderView(view){
  const v = (view||'').toString().trim().toLowerCase();
  const map = {
    login: 'login',
    hub: 'hub',
    hub_beneficio: 'hub_beneficio',
    hub_beneficio_eventual: 'hub_beneficio_eventual',
    hub_beneficio_rma: 'hub_beneficio_rma',
    rma: 'rma',
    vigilancia: 'vigilancia',
    formulario: 'formulario',
    baixa: 'baixa',
    central: 'central',
    admin: 'admin',
    analista: 'analista',
    ti: 'ti',
    ti_tv: 'ti',
    ti_relatorios: 'ti'
  };
  const file = map[v] || 'login';
  return __content_(file, { view: v });
}

function include(file){
  return __content_(file, { view: '' });
}

// getters (mantém compatibilidade)
function getFormularioHtml(){ return __content_('formulario', { view:'formulario' }); }
function getAdminHtml()     { return __content_('admin',      { view:'admin' }); }
function getCentralHtml()   { return __content_('central',    { view:'central' }); }
function getAnalistaHtml()  { return __content_('analista',   { view:'analista' }); }
function getHubHtml()       { return __content_('hub',        { view:'hub' }); }
function getHubBeneficioHtml(){ return __content_('hub_beneficio', { view:'hub_beneficio' }); }
function getHubBeneficioEventualHtml(){ return __content_('hub_beneficio_eventual', { view:'hub_beneficio_eventual' }); }
function getHubBeneficioRmaHtml(){ return __content_('hub_beneficio_rma', { view:'hub_beneficio_rma' }); }
function getBaixaHtml()     { return __content_('baixa',      { view:'baixa' }); }
function getRmaHtml()       { return __content_('rma',        { view:'rma' }); }
function getVigilanciaHtml(){ return __content_('vigilancia', { view:'vigilancia' }); }

// ✅ TI: dois nomes (pra não quebrar o front)
function gettiHtml()        { return __content_('ti',         { view:'ti' }); }
function getTiHtml()        { return __content_('ti',         { view:'ti' }); }
function getTiTvHtml()      { return __content_('ti',         { view:'ti_tv' }); }
function getTiRelatoriosHtml(){ return __content_('ti',       { view:'ti_relatorios' }); }

/************************
 * LOGIN
 ************************/
function canonicalSectorKey_(s){
  return (s || '').toString()
    .replace(/\u00A0/g,' ')
    .replace(/[\u200B-\u200D\uFEFF]/g,'')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g,'');
}

function cleanSectorLabel_(s){
  return (s || '').toString()
    .replace(/\u00A0/g,' ')
    .replace(/[\u200B-\u200D\uFEFF]/g,'')
    .replace(/\s+/g,' ')
    .trim();
}

function listarSetores() {
  const ss  = SpreadsheetApp.openById(PLANILHA_ID);
  const sh  = ss.getSheetByName('Login');
  if (!sh) throw new Error('Aba "Login" não encontrada.');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) throw new Error('Aba "Login" sem dados.');

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  let colIdx = findColIdx_(headers, 'Setor','Unidade','Setor/Unidade','Setor (Unidade)');
  if (colIdx < 0) {
    const row2 = sh.getRange(2,1,1,lastCol).getValues()[0];
    colIdx = row2.findIndex(v => String(v||'').trim() !== '');
    if (colIdx < 0) colIdx = 0;
  }

  const valores = sh.getRange(2, colIdx+1, lastRow-1, 1).getValues().map(r=>r[0]);
  const mapa = new Map();
  valores.forEach(v => {
    const label = cleanSectorLabel_(v);
    if (!label) return;
    mapa.set(canonicalSectorKey_(label), label);
  });

  const setores = Array.from(mapa.values()).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
  if (!setores.length) throw new Error('Nenhum setor encontrado na aba "Login".');
  return setores;
}

function isHubSector_(label){
  const s = cleanSectorLabel_(label).toUpperCase();
  return s.includes('CRAS') || s.includes('CREAS') || s.includes('CRAM') ||
         s.includes('SEDE') || s.includes('CENTROPOP');
}
function isTISector_(label){
  const k = canonicalSectorKey_(label);
  return k === 'ti' || k.includes('tecnologiainformacao') || k.includes('tecnologiadainformacao');
}
function isBeneficioEventualSector_(label){
  const k = canonicalSectorKey_(label);
  return k === 'beneficioeventual' || k.includes('beneficioeventual');
}

function getDirectViewForRole_(role){
  const r = (role || '').toString().trim().toLowerCase();
  if (r === 'admin')    return 'admin';
  if (r === 'central')  return 'central';
  if (r === 'analista') return 'analista';
  if (r === 'ti')       return 'ti';
  if (r === 'beneficio') return 'hub_beneficio';
  return 'formulario';
}

function verificarLogin(setor, usuario, senha) {
  const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Login');
  if (!aba) return { ok:false };

  const dados = aba.getDataRange().getValues();
  const setorKey = canonicalSectorKey_(setor);
  const userKey  = canonicalSectorKey_(usuario);
  const pass     = (senha || '').toString().trim();

  for (let i=1; i<dados.length; i++) {
    const set  = cleanSectorLabel_(dados[i][0] || '');
    const usr  = (dados[i][1] || '').toString();
    const pwd  = (dados[i][2] || '').toString();
    const role = (dados[i][4] || 'usuario').toString();
    if (canonicalSectorKey_(set) === setorKey &&
        canonicalSectorKey_(usr) === userKey &&
        pwd.trim() === pass) {
      return { ok:true, role:(role.trim() || 'usuario'), setorLabel:set };
    }
  }
  return { ok:false };
}

function autenticarERetornarTela(setor, usuario, senha){
  const res = verificarLogin(setor, usuario, senha);
  if (!res || !res.ok) return { ok:false };

  try {
    PropertiesService.getUserProperties()
      .setProperties({ semfas_setor:setor, semfas_usuario:usuario }, true);
  } catch (_){}

  const setorLabel = res.setorLabel || setor;

  // ✅ PRIORIDADE: TI sempre vai para a view "ti"
  let view = '';
  if (isTISector_(setorLabel)) {
    view = 'ti';
  } else if (isBeneficioEventualSector_(setorLabel)) {
    view = 'hub_beneficio';
  } else if (isHubSector_(setorLabel)) {
    view = 'hub';
  } else {
    view = getDirectViewForRole_(res.role);
  }

  return {
    ok: true,
    view,
    html: renderView(view)
  };
}

// ✅ helper pro TI.html mostrar topo (usuário/setor)
function getUserContext(){
  try{
    const up = PropertiesService.getUserProperties().getProperties();
    return {
      setor: up.semfas_setor || '',
      usuario: up.semfas_usuario || ''
    };
  }catch(_){
    return { setor:'', usuario:'' };
  }
}

/************************
 * HELPERs
 ************************/
function normText_(s){
  return String(s||'')
    .normalize('NFD').replace(/[\u00B7\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ')
    .trim()
    .toLowerCase();
}

function withRetry(fn, tentativas=5, baseMs=250){
  for (let i=0;i<tentativas;i++){
    try { return fn(); }
    catch(e){
      if (i===tentativas-1) throw e;
      Utilities.sleep(baseMs * Math.pow(2,i));
    }
  }
}

function parseAnyDate(s){
  if (!s) return null;
  if (s instanceof Date) return s;

  if (typeof s === 'number'){
    if (s > 10*365*24*3600*1000) return new Date(s); // epoch ms
    return new Date(Math.round((s - 25569) * 86400 * 1000)); // serial Sheets
  }

  if (typeof s !== 'string') return null;

  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);

  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1]);

  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function coerceSheetDate(v){
  if (!v && v!==0) return null;
  if (v instanceof Date) return startOfDay(v);
  if (typeof v === 'number')
    return startOfDay(new Date(Math.round((v - 25569) * 86400 * 1000)));
  if (typeof v === 'string'){
    const d = parseAnyDate(v);
    return d ? startOfDay(d) : null;
  }
  return null;
}

function startOfDay(d){
  const x=new Date(d);
  x.setHours(0,0,0,0);
  return x;
}

function endOfDay(d){
  const x=new Date(d);
  x.setHours(23,59,59,999);
  return x;
}

function firstDayNextMonth_(d){
  const x=new Date(d.getFullYear(), d.getMonth()+1, 1);
  x.setHours(0,0,0,0);
  return x;
}

function toISODate_(d){
  const y = d.getFullYear();
  const m = ('0'+(d.getMonth()+1)).slice(-2);
  const da= ('0'+d.getDate()).slice(-2);
  return `${y}-${m}-${da}`;
}

/** CPF sempre tratado como string */
function normalizeCPF(cpf){
  return String(cpf == null ? '' : cpf).replace(/\D/g,'');
}

function formatCPF(cpf){
  const num = normalizeCPF(cpf);
  return num.length===11
    ? num.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/,'$1.$2.$3-$4')
    : cpf;
}

function formatDateBR(d){
  if (!(d instanceof Date)) return '';
  return Utilities.formatDate(d, TZ, 'dd/MM/yyyy');
}

function parseISODateSafe(str){
  if (!str) return new Date();
  if (str instanceof Date) return str;
  if (typeof str === 'number') return new Date(str);

  let m = String(str).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1]);

  m = String(str).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);

  const d = new Date(str);
  return isNaN(d) ? new Date() : d;
}

/** ✔︎ Caixinhas de PARECER no PDF */
const BOX_CHECKED   = '☑';
const BOX_UNCHECKED = '☐';

function normalizeParecerOpcao_(v){
  const s = (v||'').toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase().trim();
  if (!s) return '';
  if (/(favoravel|aprovado|deferido|sim|favor)/.test(s)) return 'FAVORAVEL';
  if (/(desfavoravel|reprovado|indeferido|nao|não|contra)/.test(s)) return 'DESFAVORAVEL';
  return '';
}

/************************
 * DRIVE / PDFs
 ************************/
function safeFolderName_(s){
  const label = (s || 'Sem Setor').toString().trim();
  return label
    .replace(/[\\\/<>:"|?*\u0000-\u001F]+/g,'-')
    .replace(/\s+/g,' ')
    .trim();
}

function ensureSubfolder_(parent, name){
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getPdfDestinationFolder_(setorLabel, dateObj){
  const tz   = TZ;
  const root = DriveApp.getFolderById(FOLDER_PDFS_ID);
  const setor= ensureSubfolder_(root, safeFolderName_(setorLabel));
  const year = ensureSubfolder_(setor, Utilities.formatDate(dateObj,tz,'yyyy'));
  const month= ensureSubfolder_(year,  Utilities.formatDate(dateObj,tz,'MM'));
  const day  = ensureSubfolder_(month, Utilities.formatDate(dateObj,tz,'dd'));
  return day;
}

/**
 * Pasta específica do caso:
 *  FOLDER_PDFS_ID / Setor / Ano / Mês / Dia / "[PROTO] - NOME"
 */
function getCaseFolder_(setorLabel, dateObj, protocolo, nomeBeneficiario){
  const baseDay = getPdfDestinationFolder_(setorLabel, dateObj);
  const proto = (protocolo || '').toString().trim();
  const nome  = safeFolderName_(nomeBeneficiario || 'Beneficiário');
  let folderName = proto ? `${proto} - ${nome}` : nome;
  if (folderName.length > 120) folderName = folderName.substring(0,120);
  return ensureSubfolder_(baseDay, folderName);
}

function setPublicSharing_(file){
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e){
    try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch(_){}
  }
  return file;
}

function buildDriveLinks_(id){
  return {
    view:`https://drive.google.com/file/d/${id}/view`,
    preview:`https://drive.google.com/file/d/${id}/preview`,
    download:`https://drive.google.com/uc?export=download&id=${id}`
  };
}

function ti__getOrCreateProtocoloFolder_(protocolo){
  const root = DriveApp.getFolderById(FOLDER_PDFS_ID);
  const name = safeFolderName_(String(protocolo || '').trim() || 'SEM_PROTOCOLO');
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function ti__listDriveFilesByFolderId_(folderId, maxItems){
  const id = String(folderId || '').trim();
  if (!id) return { links: [], nomes: [] };

  const limit = Math.min(Math.max(parseInt(maxItems || 20, 10) || 20, 1), 40);
  const links = [];
  const nomes = [];

  try {
    const folder = DriveApp.getFolderById(id);
    const it = folder.getFiles();
    let scanned = 0;
    while (it.hasNext() && links.length < limit && scanned < 300) {
      scanned++;
      const file = it.next();
      try {
        if (!file) continue;
        const fileId = file.getId();
        const nome = String(file.getName() || ('Anexo ' + (links.length + 1))).trim();
        if (!fileId) continue;
        try { setPublicSharing_(file); } catch(_){ }
        const lnk = buildDriveLinks_(fileId);
        links.push(lnk.view || file.getUrl());
        nomes.push(nome || ('Anexo ' + links.length));
      } catch(_){ }
    }
  } catch(_){ }

  return { links, nomes };
}

/** ***********************
 *  ASSINATURA DIGITAL
 **************************/
function dataUrlToBlob_(dataUrl, defaultMime, fileName) {
  if (!dataUrl) return null;
  const str = String(dataUrl);
  let mime = defaultMime || MimeType.PNG;
  let b64  = str;

  const m = str.match(/^data:([\w\/\-\+\.]+);base64,(.+)$/);
  if (m) { mime = m[1]; b64 = m[2]; }

  const bytes = Utilities.base64Decode(b64);
  return Utilities.newBlob(bytes, mime, fileName || 'arquivo');
}

function insertSignatureImageAtPlaceholder_(body, placeholderKey, blob) {
  if (!body || !blob || !placeholderKey) return;

  const pattern = '\\{\\{?\\s*' + placeholderKey + '\\s*\\}?\\}';
  const result = body.findText(pattern);
  if (!result) return;

  const el   = result.getElement();
  const text = el.asText();
  const start = result.getStartOffset();
  const end   = result.getEndOffsetInclusive();

  text.deleteText(start, end);

  let parent = text.getParent();
  while (parent && parent.getType() !== DocumentApp.ElementType.PARAGRAPH && parent.getParent) {
    parent = parent.getParent();
  }
  if (!parent || parent.getType() !== DocumentApp.ElementType.PARAGRAPH) return;

  const para = parent.asParagraph();
  const img = para.insertInlineImage(para.getNumChildren(), blob);
  try { if (img.getWidth() > 140) img.setWidth(140); } catch (_){}
}

/************************
 * DESCOBERTA AUTOMÁTICA
 ************************/
function getSheet_(name){
  return SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(name);
}

function autoDetectDataSheet_(){
  const ss = SpreadsheetApp.openById(PLANILHA_ID);

  const pref = ss.getSheetByName(RESPOSTAS_SHEET_NAME);
  if (pref && pref.getLastRow() >= 2) return pref;

  const sheets = ss.getSheets();
  let best = null, bestScore = -1;

  sheets.forEach(sh => {
    const lr = sh.getLastRow();
    const lc = sh.getLastColumn();
    if (lr < 2 || lc < 3) return;

    const headers = sh.getRange(1,1,1,lc).getValues()[0];

    const iCPF   = findColIdx_(headers, CPF_HEADER,'cpf','cpf do beneficiário','cpf do beneficiario','documento (cpf)','cpf beneficiário');
    const iNome  = findColIdx_(headers, 'nome','nome do beneficiário','nome do beneficiario','beneficiário','beneficiario');
    const iUnid  = findColIdx_(headers, 'via de entrada','unidade','unidade / setor','setor','cras','creas','cram');
    const iBenef = findColIdx_(headers, 'benefício','beneficio','demanda apresentada','tipo de benefício','tipo de beneficio');
    const iData  = findColIdx_(headers, 'data da solicitação','data','timestamp','solicitado em','data de cadastro');

    const sampleRows = Math.min(400, lr-1);
    const rng = sh.getRange(2,1,sampleRows,lc).getValues();

    let cpfOk = 0;
    if (iCPF >= 0){
      for (let r=0;r<rng.length;r++){
        const n = String(rng[r][iCPF]||'').replace(/\D/g,'');
        if (n.length === 11) cpfOk++;
      }
    }

    let dataOk = 0;
    if (iData >= 0){
      for (let r=0;r<rng.length;r++){
        const v = rng[r][iData];
        if (v instanceof Date || typeof v === 'number'){ dataOk++; continue; }
        const s = String(v||'').trim();
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(s) || /^\d{4}-\d{2}-\d{2}$/.test(s)) dataOk++;
      }
    }

    let unSet = new Set();
    if (iUnid >= 0){
      for (let r=0;r<rng.length;r++){
        const s = String(rng[r][iUnid]||'').trim();
        if (s) unSet.add(s);
      }
    }

    let headerScore = 0;
    [iCPF,iNome,iUnid,iBenef,iData].forEach(i => { if (i>=0) headerScore+=15; });

    const bulk  = Math.min(lr-1, 5000);
    const score = headerScore + cpfOk*2 + dataOk*1.5 + Math.min(unSet.size,50) + bulk/5;

    if (score > bestScore){ best = sh; bestScore = score; }
  });

  if (!best) throw new Error('Nenhuma aba de dados compatível encontrada.');
  return best;
}

function findColIdx_(headers, ...cands){
  const H = (headers||[]).map(h=>String(h||'').trim().toLowerCase());
  const flat = cands.flat();

  for (const cand of flat){
    const i = H.indexOf(String(cand||'').trim().toLowerCase());
    if (i >= 0) return i;
  }

  const patterns = flat.map(c =>
    new RegExp(String(c).replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), 'i')
  );
  for (let i=0;i<(headers||[]).length;i++){
    const h = String(headers[i]||'');
    if (patterns.some(rx => rx.test(h))) return i;
  }
  return -1;
}

function ensureColumn_(sheet, header){
  const range   = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn()));
  const headers = range.getValues()[0] || [];
  const idx     = findColIdx_(headers, header);
  if (idx === -1){
    sheet.insertColumnAfter(headers.length || 1);
    sheet.getRange(1, headers.length+1).setValue(header);
    return headers.length;
  }
  return idx;
}

function guessCpfColumn_(sh, startRow, lastRow, lastCol){
  const rows = Math.min(300, lastRow - startRow + 1);
  if (rows <= 0) return -1;

  let bestIdx = -1, bestScore = 0;
  for (let c=1;c<=lastCol;c++){
    const vals = sh.getRange(startRow, c, rows, 1).getValues();
    let score=0;
    for (let i=0;i<vals.length;i++){
      const n = String(vals[i][0]||'').replace(/\D/g,'');
      if (n.length === 11) score++;
    }
    if (score > bestScore){ bestScore = score; bestIdx = c-1; }
  }
  return bestScore >= 3 ? bestIdx : -1;
}

function guessDateColumn_(sh, startRow, lastRow, lastCol){
  const rows = Math.min(300, lastRow - startRow + 1);
  if (rows <= 0) return -1;

  let bestIdx = -1, bestScore = 0;
  for (let c=1;c<=lastCol;c++){
    const vals = sh.getRange(startRow, c, rows, 1).getValues();
    let score=0;
    for (let i=0;i<vals.length;i++){
      const v = vals[i][0];
      if (v instanceof Date || typeof v === 'number'){ score++; continue; }
      const s = String(v||'').trim();
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(s) || /^\d{4}-\d{2}-\d{2}$/.test(s)) score++;
    }
    if (score > bestScore){ bestScore = score; bestIdx = c-1; }
  }
  return bestScore >= 3 ? bestIdx : -1;
}

function ensureDocumentosColumn_(sheet) {
  return ensureColumn_(sheet, 'Documentos');
}

function getRespostasIndexMap_(){
  const sh = autoDetectDataSheet_();

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0] || [];

  const idx = {
    unidade:   findColIdx_(headers,'via de entrada','unidade','unidade / setor','setor','cras','creas','cram','unidade/setor','setor/unidade','unidade/setor (cras/creas)'),
    beneficio: findColIdx_(headers,'benefício','beneficio','benefício social','demanda apresentada','tipo de benefício','tipo de beneficio','beneficio social'),
    data:      findColIdx_(headers,'data da solicitação','data','timestamp','data de cadastro','solicitado em','solicitação'),
    nome:      findColIdx_(headers,'nome do beneficiário','beneficiário','beneficiario','nome','nome do beneficiario','nome beneficiário'),
    cpf:       findColIdx_(headers,CPF_HEADER,'cpf','cpf do beneficiário','cpf do beneficiario','documento (cpf)','cpf beneficiário'),
    status:    findColIdx_(headers,'status','situação','situacao'),
    pdf:       findColIdx_(headers,'pdf','link do pdf','link','arquivo','drive','url pdf','arquivo pdf'),
    docs:      findColIdx_(headers,'documentos','docs','anexos','arquivos anexos','documentos/anexos'),
    entregueFlag: -1
  };

  if (idx.cpf  < 0) idx.cpf  = guessCpfColumn_(sh, 2, Math.max(2,lastRow), lastCol);
  if (idx.data < 0) idx.data = guessDateColumn_(sh, 2, Math.max(2,lastRow), lastCol);

  if (idx.status < 0) idx.status = ensureColumn_(sh, 'Status');
  idx.entregueFlag = ensureColumn_(sh, ENTREGUE_HEADER);
  idx.docs = idx.docs >= 0 ? idx.docs : ensureDocumentosColumn_(sh);

  // fallback (se planilha for “estranha”)
  if (idx.unidade  < 0) idx.unidade   = 1;
  if (idx.beneficio< 0) idx.beneficio = 2;
  if (idx.data     < 0) idx.data      = 3;
  if (idx.nome     < 0) idx.nome      = 4;
  if (idx.cpf      < 0) idx.cpf       = 12;

  return { sh, headers, idx };
}

/************************
 * AUX — ENTREGUE robusto
 ************************/
function isRowEntregue_(row, idx){
  const flag = String(row[idx.entregueFlag]||'').trim().toUpperCase();
  const st   = idx.status>=0 ? String(row[idx.status]||'').trim().toUpperCase() : '';
  return (
    flag === 'ENTREGUE' || flag === 'SIM' || flag === 'TRUE' || flag === 'VERDADEIRO' ||
    st === 'ENTREGUE' || st.includes('ENTREG')
  );
}

/************************
 * CPF: existência
 ************************/
function checkCpfExistence(cpf) {
  cpf = normalizeCPF(cpf);
  if (!cpf) return { exists:false };

  const { sh, idx } = getRespostasIndexMap_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { exists:false };

  const values = sh.getRange(2, idx.cpf+1, lastRow-1, 1).getValues();
  const exists = values.some(r => normalizeCPF(r[0]) === cpf);
  return { exists };
}

/************************
 * BENEFÍCIOS ILIMITADOS
 ************************/
function normalizeBenefit_(s){
  return (s || '').toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[()\-]/g,' ')
    .replace(/[^a-zA-Z0-9 ]/g,'')
    .trim()
    .toLowerCase()
    .replace(/\s+/g,'');
}

const BENEFICIOS_ILIMITADOS = new Set([
  normalizeBenefit_('AUXÍLIO POR MORTE'),
  normalizeBenefit_('MAIS ACONCHEGO (NATALIDADE)'),
  normalizeBenefit_('PASSAGEM INTERESTADUAL (VIAGEM)'),
  normalizeBenefit_('Outro')
]);

/************************
 * 🔹 ANEXOS – helpers
 ************************/
function salvarAnexosNoDrive_(anexos, pastaDestino, protocolo) {
  if (!anexos || !anexos.length || !pastaDestino) return '';

  const linhas = [];
  anexos.forEach(function (item, idx) {
    try {
      if (!item) return;

      const nomeOriginal = (item.nome || item.filename || ('arquivo_' + (idx+1))).toString();
      const rotulo       = (item.nomeDocumento || '').toString().trim();
      const mime         = (item.mimeType || item.mimetype || MimeType.PDF);

      let base64 = item.base64 || item.conteudoBase64 || item.data || item.dataUrl || '';
      if (!base64) return;

      const m = String(base64).match(/^data:([\w\/\-\+\.]+);base64,(.+)$/);
      let mimeUse = mime;
      if (m) { mimeUse = m[1] || mimeUse; base64 = m[2] || ''; }

      const bytes = Utilities.base64Decode(base64);
      const blob  = Utilities.newBlob(bytes, mimeUse, nomeOriginal);

      const nomeFinal = (protocolo ? protocolo + ' - ' : '') + (rotulo ? (rotulo + ' - ') : '') + nomeOriginal;

      const file = withRetry(function(){
        return pastaDestino.createFile(blob).setName(nomeFinal);
      }, 3, 300);

      setPublicSharing_(file);
      const links = buildDriveLinks_(file.getId());

      const etiqueta = rotulo || nomeOriginal;
      linhas.push(etiqueta + ': ' + links.preview);

    } catch (e) {
      Logger.log('[ANEXO] Erro ao salvar anexo: ' + e.message);
    }
  });

  return linhas.join('\n');
}

/************************
 * SALVAR + PDF + E-MAIL
 ************************/
function salvarFormulario(dados) {
  // Datas & formatos
  const dataSolicitacao   = parseISODateSafe(dados.data_solicitacao);
  const dataSolicitacaoBR = formatDateBR(dataSolicitacao);
  dados.data_solicitacao  = dataSolicitacaoBR;

  if (dados.nasc) {
    const nascDate = parseISODateSafe(dados.nasc);
    dados.nasc = formatDateBR(nascDate);
  }
  dados.cpf = formatCPF(dados.cpf);

  // Se não vier "cadastrador" do front, tenta o usuário logado
  if (!dados.cadastrador) {
    try {
      const up = PropertiesService.getUserProperties().getProperties();
      dados.cadastrador = up.semfas_usuario || Session.getActiveUser().getEmail() || '';
    } catch (_){
      dados.cadastrador = '';
    }
  }

  // 🔹 ASSINATURA DIGITAL (base64 vindo do formulário)
  const assinaturaB64 =
    dados.assinatura_base64 ||
    dados.assinaturaDigitalBase64 ||
    dados.assinatura_digital_base64 ||
    dados.assinaturaDigital ||
    dados.assinatura_img ||
    dados.assinaturaImagem ||
    dados.assinatura_digital ||
    '';

  let assinaturaBlob = null;
  if (assinaturaB64) {
    try { assinaturaBlob = dataUrlToBlob_(assinaturaB64, MimeType.PNG); }
    catch (e) { Logger.log('[ASSINATURA] Erro ao converter base64: ' + e.message); }
  }

  // 🔹 anexos
  const anexos = Array.isArray(dados.anexos) ? dados.anexos : [];

  const shPref = getSheet_(RESPOSTAS_SHEET_NAME);
  const sh = (shPref && shPref.getLastRow() >= 1) ? shPref : autoDetectDataSheet_();

  let newRow, protocoloGerado = '';
  const lock1 = LockService.getScriptLock(); lock1.waitLock(10000);

  try {
    const registros = sh.getDataRange().getValues();

    // Duplicidade (CPF + Benefício + mês/ano)
    const cpfNovo       = normalizeCPF(dados.cpf);
    const beneficioNovo = (dados.demanda || '').toString().trim();
    const benKeyNovo    = normalizeBenefit_(beneficioNovo);
    const mesNovo       = dataSolicitacao.getMonth();
    const anoNovo       = dataSolicitacao.getFullYear();
    const isIlimitado   = BENEFICIOS_ILIMITADOS.has(benKeyNovo);

    let lastMatchDate = null;
    for (let i=1; i<registros.length; i++) {
      const cpfExistente       = normalizeCPF((registros[i][12] || '').toString().trim());
      const beneficioExistente = (registros[i][2]  || '').toString().trim();
      const benKeyExistente    = normalizeBenefit_(beneficioExistente);
      const dataExistente      = registros[i][3] ? parseAnyDate(registros[i][3]) : null;

      if (cpfExistente === cpfNovo &&
          benKeyExistente === benKeyNovo &&
          dataExistente &&
          dataExistente.getMonth() === mesNovo &&
          dataExistente.getFullYear() === anoNovo) {
        if (!lastMatchDate || dataExistente > lastMatchDate) lastMatchDate = dataExistente;
      }
    }

    if (!isIlimitado && lastMatchDate) {
      const liberadoEm = firstDayNextMonth_(lastMatchDate);
      return {
        sucesso:false,
        mensagem:'Já existe um cadastro para este CPF com este benefício neste mês.',
        ultima_data:  formatDateBR(lastMatchDate),
        proxima_data: formatDateBR(liberadoEm)
      };
    }

    // Linha (estrutura atual)
    const linha = [
      new Date(), dados.via_entrada, dados.demanda, dataSolicitacaoBR,
      dados.nomeb, dados.nasc, dados.endereco, dados.bairro, dados.referencia,
      dados.telefone, dados.rg, dados.ssp, dados.cpf, (dados.nis || ''),
      dados.cras_ref, dados.tem_beneficio, dados.qual_beneficio, dados.membros,
      dados.renda, dados.nomes, dados.telefones, dados.enderecos, dados.rgs,
      dados.cpfs, dados.vinculo, dados.situacao,
      'PENDENTE', '', dados.cadastrador
    ];

    withRetry(()=>sh.appendRow(linha), 5, 250);
    newRow = sh.getLastRow();

    // Protocolo
    protocoloGerado = ensureProtocoloForRow_(sh, newRow);

    // PARECER TÉCNICO
    const iPar   = ensureColumn_(sh, 'PARECER TÉCNICO');
    const iNomT  = ensureColumn_(sh, 'NOME TÉCNICO');
    const iMun   = ensureColumn_(sh, 'MUNICÍPIO PARECER');
    const iData  = ensureColumn_(sh, 'DATA PARECER');
    const iAssB  = ensureColumn_(sh, 'ASSINATURA BENEFICIÁRIO');
    const iAssT  = ensureColumn_(sh, 'ASSINATURA/CARIMBO');

    const dataParecerBR = dados.data_parecer ? formatDateBR(parseISODateSafe(dados.data_parecer)) : '';

    const assinaturaSheetValue = assinaturaB64 || dados.assinatura_carimbo || '';

    sh.getRange(newRow, iPar+1 ).setValue(dados.parecer_tecnico || dados.parecer || '');
    sh.getRange(newRow, iNomT+1).setValue(dados.nome_tecnico || '');
    sh.getRange(newRow, iMun+1 ).setValue(dados.municipio_parecer || 'Nossa Senhora do Socorro');
    sh.getRange(newRow, iData+1).setValue(dataParecerBR);
    sh.getRange(newRow, iAssB+1).setValue(dados.assinatura_beneficiario || '');
    sh.getRange(newRow, iAssT+1).setValue(assinaturaSheetValue);

  } finally {
    lock1.releaseLock();
  }

  // Pasta do caso (Setor/Ano/Mês/Dia/[PROTO - NOME])
  const setorLabel = dados.via_entrada || 'Sem Setor';
  const dataRef    = parseISODateSafe(dados.data_solicitacao || new Date());
  const pastaCaso  = getCaseFolder_(setorLabel, dataRef, protocoloGerado, dados.nomeb);

  // PDF
  const pdf = gerarFichaPDF(Object.assign({}, dados, {
    protocolo: protocoloGerado,
    __assinaturaBlob__: assinaturaBlob
  }), pastaCaso);

  // Atualiza PDF + Documentos
  const lock2 = LockService.getScriptLock(); lock2.waitLock(10000);
  try {
    const iPdf  = ensureColumn_(sh, 'PDF');
    const iDocs = ensureDocumentosColumn_(sh);

    withRetry(()=>sh.getRange(newRow, iPdf+1).setValue(pdf.downloadUrl), 5, 250);

    if (anexos && anexos.length) {
      const textoDocs = salvarAnexosNoDrive_(anexos, pastaCaso, protocoloGerado);
      if (textoDocs) {
        const anterior = sh.getRange(newRow, iDocs+1).getValue() || '';
        const novo     = anterior ? (anterior + '\n' + textoDocs) : textoDocs;
        withRetry(()=>sh.getRange(newRow, iDocs+1).setValue(novo), 5, 250);
      }
    }

  } finally {
    lock2.releaseLock();
  }

  // E-mail ao setor (se configurado)
  enviarEmailSetor(dados.via_entrada, pdf, dados.nomeb);

  return {
    sucesso:true,
    mensagem:'Cadastro realizado, PDF e anexos salvos em pasta organizada, e e-mail enviado ao setor!',
    link: pdf.downloadUrl,
    protocolo: protocoloGerado,
    concluirTexto:'Concluído'
  };
}

/************************
 * GERAÇÃO DO PDF (template + assinatura)
 ************************/
function gerarFichaPDF(dados, destFolderOpt) {
  const dSolic = coerceSheetDate(dados.data_solicitacao || dados.data || dados.dataSolicitacao);
  const dNasc  = coerceSheetDate(dados.nasc || dados.data_nascimento || dados.dataNascimento);

  const assinaturaBlob = dados.__assinaturaBlob__ || null;

  const parecerOp = normalizeParecerOpcao_(dados.parecer_opcao || dados.parecer_tecnico || dados.parecer);

  const payload = Object.assign({}, dados, {
    data_solicitacao: dSolic ? formatDateBR(dSolic) : (dados.data_solicitacao || ''),
    nasc:             dNasc  ? formatDateBR(dNasc)  : (dados.nasc || ''),
    cpf:              formatCPF(dados.cpf || ''),
    cadastrador:      dados.cadastrador || '',

    parecer:                 dados.parecer_tecnico || dados.parecer || '',
    nome_tecnico:            dados.nome_tecnico || '',
    municipio_parecer:       dados.municipio_parecer || 'Nossa Senhora do Socorro',
    data_parecer:            dados.data_parecer ? formatDateBR(parseISODateSafe(dados.data_parecer)) : '',
    assinatura_beneficiario: dados.assinatura_beneficiario || '',
    assinatura_carimbo:      dados.assinatura_carimbo || '',
    protocolo:               dados.protocolo || '',

    box_favoravel:    (parecerOp === 'FAVORAVEL')    ? BOX_CHECKED   : BOX_UNCHECKED,
    box_desfavoravel: (parecerOp === 'DESFAVORAVEL') ? BOX_CHECKED   : BOX_UNCHECKED,
    box_indef:        (!parecerOp)                   ? BOX_CHECKED   : BOX_UNCHECKED
  });

  delete payload.__assinaturaBlob__;

  const modelo = DriveApp.getFileById(ID_TEMPLATE);
  const copia  = withRetry(
    () => modelo.makeCopy('Ficha - ' + (payload.nomeb || 'Beneficiário')),
    5, 250
  );
  const doc   = DocumentApp.openById(copia.getId());
  const corpo = doc.getBody();

  try {
    corpo.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);
  } catch (_){}

  function replacePlaceholder_(key, value) {
    const val = (value == null ? '' : String(value)).replace(/\r\n/g,'\n');
    const pattern = '\\{\\{\\s*' + key + '\\s*\\}\\}|\\{\\s*' + key + '\\s*\\}';
    try { corpo.replaceText(pattern, val); }
    catch (e) { Logger.log('[PDF] replaceText ' + key + ': ' + e.message); }
  }

  Object.keys(payload).forEach(k => {
    const isAssinaturaKey =
      k === 'assinatura_carimbo' ||
      k === 'assinatura_img' ||
      k === 'assinatura_digital' ||
      k === 'assinaturaImagem';

    if (assinaturaBlob && isAssinaturaKey) return;
    replacePlaceholder_(k, payload[k]);
  });

  if (assinaturaBlob) {
    try { insertSignatureImageAtPlaceholder_(corpo, 'assinatura_carimbo', assinaturaBlob); }
    catch (e) { Logger.log('[ASSINATURA] Erro ao inserir imagem: ' + e.message); }
  }

  try { corpo.replaceText('\\{\\s*([^{}]+)\\s*\\}', '$1'); } catch (e) { Logger.log('[PDF] limpar chaves simples: ' + e.message); }
  try { corpo.replaceText('\\{\\{\\s*[^}]+\\s*\\}\\}', ''); } catch (e) { Logger.log('[PDF] limpar placeholders {{}}: ' + e.message); }

  doc.saveAndClose();
  Utilities.sleep(300);

  const nomeArquivo = `Ficha - ${(payload.nomeb || 'Beneficiário')} - ${(payload.cpf || '').toString()}.pdf`;

  const pdfBlob = withRetry(
    () => DriveApp.getFileById(copia.getId()).getAs(MimeType.PDF).copyBlob().setName(nomeArquivo),
    5, 250
  );

  const setor   = safeFolderName_(payload.via_entrada || 'Sem Setor');
  const dataRef = dSolic || new Date();

  const pastaDestino = (destFolderOpt && destFolderOpt.createFile)
    ? destFolderOpt
    : getPdfDestinationFolder_(setor, dataRef);

  const pdfFile = withRetry(() => pastaDestino.createFile(pdfBlob), 5, 250);
  setPublicSharing_(pdfFile);

  const fileId = pdfFile.getId();
  const links  = buildDriveLinks_(fileId);

  try { DriveApp.getFileById(copia.getId()).setTrashed(true); } catch (_){}

  return { blob: pdfBlob, file: pdfFile, fileId, previewUrl: links.preview, downloadUrl: links.download };
}

/************************
 * GERAR PDF A PARTIR DE UMA LINHA EXISTENTE
 ************************/
function gerarPdfDeRegistro(linha) {
  const shPref = getSheet_(RESPOSTAS_SHEET_NAME);
  const sh = (shPref && shPref.getLastRow() >= 1) ? shPref : autoDetectDataSheet_();

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] || [];
  const dados   = sh.getRange(linha,1,1,sh.getLastColumn()).getValues()[0];

  const dSolic = coerceSheetDate(dados[3]);
  const dNasc  = coerceSheetDate(dados[5]);

  const iProt = findColIdx_(headers, PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo','numero do protocolo','id','id protocolo');
  const protocoloLido = iProt >= 0 ? (dados[iProt] || '') : '';

  const iPar   = findColIdx_(headers, 'PARECER TÉCNICO');
  const iNomT  = findColIdx_(headers, 'NOME TÉCNICO');
  const iMun   = findColIdx_(headers, 'MUNICÍPIO PARECER');
  const iData  = findColIdx_(headers, 'DATA PARECER');
  const iAssB  = findColIdx_(headers, 'ASSINATURA BENEFICIÁRIO','ASSINATURA BENEFICIARIO');
  const iAssT  = findColIdx_(headers, 'ASSINATURA/CARIMBO','ASSINATURA TÉCNICO','ASSINATURA TECNICO','ASSINATURA TEC','ASSINATURA TEC.','ASSINATURA');

  let assinaturaBlob = null;
  if (iAssT >= 0) {
    const sigCell = dados[iAssT];
    if (sigCell) {
      const s = String(sigCell).trim();
      try {
        if (/^data:image\//i.test(s) || (/^[A-Za-z0-9+/=]+$/.test(s) && s.length > 100)) {
          assinaturaBlob = dataUrlToBlob_(s, MimeType.PNG);
        }
      } catch (e) {
        Logger.log('[ASSINATURA] erro ao recriar blob da planilha: ' + e.message);
      }
    }
  }

  const registro = {
    via_entrada: dados[1],
    demanda: dados[2],
    data_solicitacao: dSolic ? formatDateBR(dSolic) : '',
    nomeb: dados[4],
    nasc: dNasc ? formatDateBR(dNasc) : '',
    endereco: dados[6],
    bairro: dados[7],
    referencia: dados[8],
    telefone: dados[9],
    rg: dados[10],
    ssp: dados[11],
    cpf: formatCPF(dados[12] || ''),
    nis: dados[13],
    cras_ref: dados[14],
    tem_beneficio: dados[15],
    qual_beneficio: dados[16],
    membros: dados[17],
    renda: dados[18],
    nomes: dados[19],
    telefones: dados[20],
    enderecos: dados[21],
    rgs: dados[22],
    cpfs: dados[23],
    vinculo: dados[24],
    situacao: dados[25],
    cadastrador: dados[29] || '',
    protocolo: protocoloLido || '',
    parecer: iPar>=0 ? (dados[iPar]||'') : '',
    nome_tecnico: iNomT>=0 ? (dados[iNomT]||'') : '',
    municipio_parecer: iMun>=0 ? (dados[iMun]||'') : '',
    data_parecer: iData>=0 ? (dados[iData]||'') : '',
    assinatura_beneficiario: iAssB>=0 ? (dados[iAssB]||'') : '',
    assinatura_carimbo:      iAssT>=0 ? (dados[iAssT]||'') : ''
  };

  const setorLabel = registro.via_entrada || 'Sem Setor';
  const dataRefReg = dSolic || new Date();
  const pastaCaso  = getCaseFolder_(setorLabel, dataRefReg, protocoloLido, registro.nomeb);

  const pdf = gerarFichaPDF(Object.assign({}, registro, { __assinaturaBlob__: assinaturaBlob }), pastaCaso);

  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const iPdf = ensureColumn_(sh, 'PDF');
    sh.getRange(linha, iPdf+1).setValue(pdf.downloadUrl);
  } finally {
    lock.releaseLock();
  }
  return { link: pdf.downloadUrl };
}
/************************
 * CENTRAL (consulta)
 ************************/
function buscarBeneficios(filtroCPF = "", dataInicio = "", dataFim = "", beneficio = "", unidade = "") {
  const { sh, headers, idx } = getRespostasIndexMap_();
  const dados = sh.getDataRange().getValues();

  const listaUnidades   = new Set();
  const listaBeneficios = new Set();
  for (let i=1; i<dados.length; i++){
    const r = dados[i];
    if (r[idx.unidade])   listaUnidades.add(String(r[idx.unidade]).trim());
    if (r[idx.beneficio]) listaBeneficios.add(String(r[idx.beneficio]).trim());
  }
  unidade   = resolveFiltroOuVazio_(Array.from(listaUnidades),   unidade);
  beneficio = resolveFiltroOuVazio_(Array.from(listaBeneficios), beneficio);

  const registros = [];
  const cpfFiltroNum = normalizeCPF(filtroCPF);
  const dIni = parseAnyDate(dataInicio);
  const dFim = parseAnyDate(dataFim);
  const benFiltro = (beneficio || '').toString().trim();
  const uniFiltro = (unidade   || '').toString().trim();

  const iProt = findColIdx_(headers, PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo','numero do protocolo','id','id protocolo');

  for (let i=1; i<dados.length; i++) {
    const row = dados[i];

    const cpfStr   = (row[idx.cpf] || '').toString().trim();
    const cpfNum   = normalizeCPF(cpfStr);
    const status   = isRowEntregue_(row, idx) ? 'ENTREGUE' : (String(row[idx.status] || 'PENDENTE').toUpperCase());
    const data     = coerceSheetDate(row[idx.data]);
    const demanda  = (row[idx.beneficio] || '').toString().trim();
    const unidadeV = (row[idx.unidade] || '').toString().trim();
    const nome     = (row[idx.nome] || '').toString().trim();
    const protocolo= iProt>=0 ? String(row[iProt]||'').trim() : '';

    let okData = true;
    if (dIni && data && data < startOfDay(dIni)) okData = false;
    if (dFim && data && data > endOfDay(dFim)) okData = false;

    if (benFiltro && !matchesFilter_(demanda, benFiltro)) continue;
    if (uniFiltro && !matchesFilter_(unidadeV, uniFiltro)) continue;

    if ((!cpfFiltroNum || cpfNum === cpfFiltroNum) && okData) {
      registros.push({
        linha: i+1,
        data: data ? formatDateBR(data) : '',
        unidade: unidadeV,
        demanda,
        nome,
        cpf: formatCPF(cpfStr),
        protocolo,
        status,
        linkPdf: (idx.pdf>=0 ? (row[idx.pdf] || '') : ''),
        documentos: (idx.docs>=0 ? (row[idx.docs] || '') : '')
      });
    }
  }
  return { registros };
}

function matchesFilter_(value, filterTxt){
  const f = normText_(filterTxt);
  if (!f) return true;
  return normText_(value).includes(f);
}

function resolveFiltroOuVazio_(lista, valor){
  const v = normText_(valor||'');
  if (!v) return '';
  for (const item of lista){
    const n = normText_(item);
    if (n.includes(v) || v.includes(n)) return valor;
  }
  return '';
}

/************************
 * BAIXA — Opções / Lista / Busca / Entrega
 ************************/
function ensureBaixasSheet_(){
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName(BAIXAS_SHEET_NAME);
  if (!sh){
    sh = ss.insertSheet(BAIXAS_SHEET_NAME);
    sh.appendRow([
      'Carimbo','Protocolo','CPF','Nome','Benefício','Status Antes','Status Depois',
      'Entregue em','Unidade','Entregue por','Observação','RowRef'
    ]);
  }
  return sh;
}

function sortList_(arr, order){
  const o = String(order||'data_desc').toLowerCase();
  const coll = (a,b)=>String(a||'').localeCompare(String(b||''),'pt-BR',{sensitivity:'base'});
  const getT = v => { const d = coerceSheetDate(v); return d ? d.getTime() : 0; };
  const key = (r, k)=>{
    switch(k){
      case 'data':   return getT(r.solicitadoEm || r.data || r.dataBR || r.dataISO);
      case 'nome':   return normText_(r.nome);
      case 'benef':  return normText_(r.beneficio);
      case 'uni':    return normText_(r.unidade);
      case 'status': return normText_(r.status);
      case 'prot':   return normText_(r.protocolo);
      case 'cpf':    return (String(r.cpf||'').replace(/\D/g,'')) || '';
      default:       return '';
    }
  };
  if (o === 'data_asc')  return arr.sort((a,b)=> key(a,'data') - key(b,'data'));
  if (o === 'data_desc') return arr.sort((a,b)=> key(b,'data') - key(a,'data'));
  if (o === 'nome_asc')  return arr.sort((a,b)=> coll(key(a,'nome'),  key(b,'nome')));
  if (o === 'nome_desc') return arr.sort((a,b)=> coll(key(b,'nome'),  key(a,'nome')));
  if (o === 'benef_asc') return arr.sort((a,b)=> coll(key(a,'benef'), key(b,'benef')));
  if (o === 'benef_desc')return arr.sort((a,b)=> coll(key(b,'benef'), key(a,'benef')));
  if (o === 'uni_asc')   return arr.sort((a,b)=> coll(key(a,'uni'),   key(b,'uni')));
  if (o === 'uni_desc')  return arr.sort((a,b)=> coll(key(b,'uni'),   key(a,'uni')));
  if (o === 'prot_asc')  return arr.sort((a,b)=> coll(key(a,'prot'),  key(b,'prot')));
  if (o === 'prot_desc') return arr.sort((a,b)=> coll(key(b,'prot'),  key(a,'prot')));
  if (o === 'cpf_asc')   return arr.sort((a,b)=> coll(key(a,'cpf'),   key(b,'cpf')));
  if (o === 'cpf_desc')  return arr.sort((a,b)=> coll(key(b,'cpf'),   key(a,'cpf')));
  return arr;
}

function listarUnidadesDashboard(){
  const { sh, idx } = getRespostasIndexMap_();
  const dados = sh.getDataRange().getValues().slice(1);
  const set = new Set();
  dados.forEach(l => {
    const u = (l[idx.unidade] || '').toString().trim();
    if (u) set.add(u);
  });
  return Array.from(set).sort();
}

function listarBeneficiosDashboard(){
  const { sh, idx } = getRespostasIndexMap_();
  const dados = sh.getDataRange().getValues().slice(1);
  const set = new Set();
  dados.forEach(l => {
    const b = (l[idx.beneficio] || '').toString().trim();
    if (b) set.add(b);
  });
  return Array.from(set).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
}

function getOpcoesFiltros(){
  return { unidades: listarUnidadesDashboard(), beneficios: listarBeneficiosDashboard() };
}
function baixa_listarOpcoes(){ return getOpcoesFiltros(); }

function baixa_list(statusFilter='todos', unidadeFiltro='', beneficioFiltro='', order='data_desc', pdfOnly=false, hasProto=false, dataInicio='', dataFim=''){
  const { sh, headers, idx } = getRespostasIndexMap_();
  const values = sh.getRange(2,1, Math.max(0, sh.getLastRow()-1), sh.getLastColumn()).getValues();

  const setUnidades = new Set(), setBenef = new Set();
  for (let r=0;r<values.length;r++){
    const row = values[r];
    if (row[idx.unidade])   setUnidades.add(String(row[idx.unidade]).trim());
    if (row[idx.beneficio]) setBenef.add(String(row[idx.beneficio]).trim());
  }
  unidadeFiltro   = resolveFiltroOuVazio_(Array.from(setUnidades), unidadeFiltro);
  beneficioFiltro = resolveFiltroOuVazio_(Array.from(setBenef),    beneficioFiltro);

  const wantPend = String(statusFilter).toLowerCase() === 'pendentes';
  const wantEnt  = String(statusFilter).toLowerCase() === 'entregues';
  const iProt = findColIdx_(headers, PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo','numero do protocolo','id','id protocolo');

  const dIni = parseAnyDate(dataInicio);
  const dFim = parseAnyDate(dataFim);

  const out = [];
  for (let r=0; r<values.length; r++){
    const row = values[r];

    const unidade = (row[idx.unidade] || '').toString().trim();
    const benef   = (row[idx.beneficio] || '').toString().trim();
    if (unidadeFiltro && !matchesFilter_(unidade, unidadeFiltro)) continue;
    if (beneficioFiltro && !matchesFilter_(benef, beneficioFiltro)) continue;

    const isEnt   = isRowEntregue_(row, idx);
    if (wantPend && isEnt) continue;
    if (wantEnt  && !isEnt) continue;

    const stCell  = idx.status>=0 ? String(row[idx.status]||'').toUpperCase() : '';
    const status  = isEnt ? 'ENTREGUE' : (stCell || 'PENDENTE');
    const dataObj = coerceSheetDate(row[idx.data]);

    if (dIni && dataObj && dataObj < startOfDay(dIni)) continue;
    if (dFim && dataObj && dataObj > endOfDay(dFim))   continue;

    const linkPdf = idx.pdf>=0 ? (row[idx.pdf] || '') : '';
    if (pdfOnly && !linkPdf) continue;

    const protStr = iProt>=0 ? String(row[iProt]||'') : '';
    if (hasProto && !protStr) continue;

    out.push({
      rowRef:      r+2,
      protocolo:   protStr || '',
      nome:        idx.nome>=0 ? String(row[idx.nome]||'').trim() : '',
      cpf:         idx.cpf>=0 ? formatCPF(String(row[idx.cpf]||'')) : '',
      beneficio:   benef,
      unidade,
      solicitadoEm: dataObj ? formatDateBR(dataObj) : '',
      status,
      pdf:         linkPdf,
      entregue:    isEnt,
      documentos:  idx.docs>=0 ? (row[idx.docs] || '') : ''
    });
  }

  sortList_(out, order);
  return { registros: out.slice(0,2000) };
}

function baixa_search(q, statusFilter='todos', unidadeFiltro='', beneficioFiltro='', order='data_desc', pdfOnly=false, hasProto=false, dataInicio='', dataFim=''){
  q = String(q || '').trim();
  if (!q) return { registros: [] };

  const { sh, headers, idx } = getRespostasIndexMap_();
  const values = sh.getRange(2,1, Math.max(0, sh.getLastRow()-1), sh.getLastColumn()).getValues();

  const setUnidades = new Set(), setBenef = new Set();
  for (let r=0;r<values.length;r++){
    const row = values[r];
    if (row[idx.unidade])   setUnidades.add(String(row[idx.unidade]).trim());
    if (row[idx.beneficio]) setBenef.add(String(row[idx.beneficio]).trim());
  }
  unidadeFiltro   = resolveFiltroOuVazio_(Array.from(setUnidades), unidadeFiltro);
  beneficioFiltro = resolveFiltroOuVazio_(Array.from(setBenef),    beneficioFiltro);

  const wantPend = String(statusFilter).toLowerCase() === 'pendentes';
  const wantEnt  = String(statusFilter).toLowerCase() === 'entregues';
  const iProt = findColIdx_(headers, PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo','numero do protocolo','id','id protocolo');

  const qLower = q.toLowerCase();
  const qNum   = q.replace(/\D/g,'');
  const qNorm  = normText_(q);
  const dIni   = parseAnyDate(dataInicio);
  const dFim   = parseAnyDate(dataFim);

  const out = [];
  for (let r=0; r<values.length; r++){
    const row = values[r];

    const unidade = (row[idx.unidade] || '').toString().trim();
    const benef   = (row[idx.beneficio] || '').toString().trim();
    if (unidadeFiltro && !matchesFilter_(unidade, unidadeFiltro)) continue;
    if (beneficioFiltro && !matchesFilter_(benef, beneficioFiltro)) continue;

    const cpfStr   = idx.cpf>=0 ? String(row[idx.cpf]||'') : '';
    const cpfNum   = cpfStr.replace(/\D/g,'');
    const nomeStr  = idx.nome>=0 ? String(row[idx.nome]||'') : '';
    const nomeNorm = normText_(nomeStr);
    const protStr  = iProt>=0 ? String(row[iProt]||'') : '';

    const match =
      (qNum && (cpfNum.includes(qNum) || protStr.replace(/\D/g,'').includes(qNum))) ||
      (qNorm && nomeNorm.includes(qNorm)) ||
      (iProt>=0 && protStr.toLowerCase().includes(qLower));
    if (!match) continue;

    const isEnt   = isRowEntregue_(row, idx);
    if (wantPend && isEnt) continue;
    if (wantEnt  && !isEnt) continue;

    const stCell  = idx.status>=0 ? String(row[idx.status]||'').toUpperCase() : '';
    const status  = isEnt ? 'ENTREGUE' : (stCell || 'PENDENTE');
    const dataObj = coerceSheetDate(row[idx.data]);

    if (dIni && dataObj && dataObj < startOfDay(dIni)) continue;
    if (dFim && dataObj && dataObj > endOfDay(dFim))   continue;

    const linkPdf = idx.pdf>=0 ? (row[idx.pdf] || '') : '';
    if (pdfOnly && !linkPdf) continue;
    if (hasProto && !protStr) continue;

    out.push({
      rowRef:      r+2,
      protocolo:   protStr,
      nome:        nomeStr,
      cpf:         formatCPF(cpfStr),
      beneficio:   benef,
      unidade,
      solicitadoEm: dataObj ? formatDateBR(dataObj) : '',
      status,
      pdf:         linkPdf,
      entregue:    isEnt,
      documentos:  idx.docs>=0 ? (row[idx.docs] || '') : ''
    });
  }

  sortList_(out, order);
  return { registros: out };
}

function baixa_marcarEntrega(payload){
  const { rowRef, dataEntrega, setor, usuario, obs } = payload || {};
  if (!rowRef) throw new Error('RowRef não informado');

  const { sh, headers, idx } = getRespostasIndexMap_();

  const iEnt    = idx.entregueFlag;
  const iEntEm  = ensureColumn_(sh, ENTREGUE_EM_HDR);
  const iEntPor = ensureColumn_(sh, ENTREGUE_POR_HDR);
  const iEntUni = ensureColumn_(sh, ENTREGUE_UNID_HDR);
  const iEntObs = ensureColumn_(sh, ENTREGUE_OBS_HDR);

  let iStatus = idx.status;
  if (iStatus < 0) iStatus = ensureColumn_(sh, 'Status');

  const rowVals = sh.getRange(rowRef, 1, 1, sh.getLastColumn()).getValues()[0];
  const antes = isRowEntregue_(rowVals, idx)
    ? 'ENTREGUE'
    : String(rowVals[iStatus] || 'PENDENTE').toUpperCase();

  const agora = new Date();
  let dt = agora;

  if (dataEntrega) {
    const d = parseISODateSafe(dataEntrega);
    if (d instanceof Date && !isNaN(d.getTime())) {
      d.setHours(agora.getHours(), agora.getMinutes(), agora.getSeconds(), 0);
      dt = d;
    }
  }

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    sh.getRange(rowRef, iEnt+1).setValue('ENTREGUE');
    sh.getRange(rowRef, iEntEm+1).setValue(dt);
    sh.getRange(rowRef, iEntPor+1).setValue(usuario || '');
    sh.getRange(rowRef, iEntUni+1).setValue(setor   || '');
    sh.getRange(rowRef, iEntObs+1).setValue(obs     || '');
    sh.getRange(rowRef, iStatus+1).setValue('ENTREGUE');
  } finally {
    lock.releaseLock();
  }

  const log     = ensureBaixasSheet_();
  const idxCPF  = idx.cpf;
  const idxNome = idx.nome;
  const idxBen  = idx.beneficio;
  const iProt   = findColIdx_(headers, PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo','numero do protocolo','id','id protocolo');

  log.appendRow([
    new Date(),
    iProt   >=0 ? rowVals[iProt]   : '',
    idxCPF  >=0 ? rowVals[idxCPF]  : '',
    idxNome >=0 ? rowVals[idxNome] : '',
    idxBen  >=0 ? rowVals[idxBen]  : '',
    antes, 'ENTREGUE', dt, setor||'', usuario||'', obs||'', rowRef
  ]);

  return true;
}

/************************
 * DASHBOARD
 ************************/
function getDashboardCompleto(inicio, fim, unidade = "") {
  const { sh, idx } = getRespostasIndexMap_();
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return payloadDashboardVazio_();

  const dados = rows.slice(1);
  const dIni = parseAnyDate(inicio);
  const dFim = parseAnyDate(fim);

  const statusTotais = { SOLICITADO:0, APROVADO:0, PENDENTE:0, RECUSADO:0, ENTREGUE:0 };
  const tipos = {};
  const setores = {};
  const meses = {};
  const byMonthStatus = {};
  const dow = [0,0,0,0,0,0,0];
  const registros = [];

  dados.forEach(linha => {
    const data = coerceSheetDate(linha[idx.data]);
    if (!data) return;
    if (dIni && data < startOfDay(dIni)) return;
    if (dFim && data > endOfDay(dFim)) return;

    const setor  = (linha[idx.unidade] || 'Indefinido').toString().trim();
    if (unidade && !matchesFilter_(setor, unidade)) return;

    const isEnt   = isRowEntregue_(linha, idx);
    const status = isEnt ? 'ENTREGUE' : (String(linha[idx.status] || 'PENDENTE').toUpperCase());
    const tipo   = (linha[idx.beneficio] || 'Indefinido').toString().trim();
    const cpfStr = (linha[idx.cpf] || '').toString().trim();
    const nome   = (linha[idx.nome] || '').toString().trim();

    statusTotais.SOLICITADO++;
    if (statusTotais.hasOwnProperty(status)) statusTotais[status]++;

    tipos[tipo] = (tipos[tipo] || 0) + 1;

    if (!setores[setor]) setores[setor] = { SOLICITADO:0, APROVADO:0, PENDENTE:0, RECUSADO:0, ENTREGUE:0 };
    setores[setor].SOLICITADO++;
    if (setores[setor].hasOwnProperty(status)) setores[setor][status]++;

    const chave = (('0'+(data.getMonth()+1)).slice(-2)) + '/' + data.getFullYear();
    meses[chave] = (meses[chave] || 0) + 1;

    if (!byMonthStatus[chave]) byMonthStatus[chave] = { SOLICITADO:0, APROVADO:0, PENDENTE:0, RECUSADO:0, ENTREGUE:0 };
    byMonthStatus[chave][status] = (byMonthStatus[chave][status] || 0) + 1;

    dow[data.getDay()]++;

    registros.push({
      nome,
      cpf: formatCPF(cpfStr),
      unidade:setor,
      beneficio:tipo,
      demanda:tipo,
      status,
      dataISO: toISODate_(data),
      dataBR: formatDateBR(data)
    });
  });

  const mesesLabels = Object.keys(meses).sort((a,b)=>{
    const [ma,aa] = a.split('/').map(Number);
    const [mb,ab] = b.split('/').map(Number);
    return new Date(aa,ma-1,1) - new Date(ab,mb-1,1);
  });
  const mesesData = mesesLabels.map(k=>meses[k]||0);

  const series = { SOLICITADO:[], APROVADO:[], PENDENTE:[], RECUSADO:[], ENTREGUE:[] };
  mesesLabels.forEach(lbl=>{
    const pack = byMonthStatus[lbl] || {};
    series.SOLICITADO.push((pack.SOLICITADO||0));
    series.APROVADO  .push((pack.APROVADO  ||0));
    series.PENDENTE  .push((pack.PENDENTE  ||0));
    series.RECUSADO  .push((pack.RECUSADO  ||0));
    series.ENTREGUE  .push((pack.ENTREGUE  ||0));
  });

  return {
    status: statusTotais,
    tipos,
    setores,
    meses, mesesLabels, mesesData,
    byMonthStatusLabels: mesesLabels,
    byMonthStatusSeries: series,
    dowLabels: ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'],
    dowData: dow,
    registros
  };
}

function payloadDashboardVazio_(){
  return {
    status:{ SOLICITADO:0, APROVADO:0, PENDENTE:0, RECUSADO:0, ENTREGUE:0 },
    tipos:{}, setores:{}, meses:{},
    mesesLabels:[], mesesData:[],
    byMonthStatusLabels:[],
    byMonthStatusSeries:{ SOLICITADO:[], APROVADO:[], PENDENTE:[], RECUSADO:[], ENTREGUE:[] },
    dowLabels:['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'],
    dowData:[0,0,0,0,0,0,0],
    registros:[]
  };
}

function buscarRegistrosIndefinidos(){
  const { sh, idx } = getRespostasIndexMap_();
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return { registros: [] };

  const dados = rows.slice(1);
  const indefinidos = [];

  dados.forEach((linha, i) => {
    const unidade = (linha[idx.unidade] || '').toString().trim();
    if (!unidade || unidade === '') {
      const data = coerceSheetDate(linha[idx.data]);
      indefinidos.push({
        linha: i + 2,
        nome: idx.nome >= 0 ? (linha[idx.nome] || '').toString().trim() : '',
        cpf: idx.cpf >= 0 ? formatCPF(linha[idx.cpf]) : '',
        beneficio: idx.beneficio >= 0 ? (linha[idx.beneficio] || '').toString().trim() : '',
        data: data ? formatDateBR(data) : '',
        viaEntrada: idx.unidade >= 0 ? (linha[idx.unidade] || '').toString().trim() : '',
        status: idx.status >= 0 ? (linha[idx.status] || 'PENDENTE').toString() : 'PENDENTE'
      });
    }
  });

  return { registros: indefinidos };
}

/************************
 * E-MAIL (fila/retry)
 ************************/
function enviarEmailSetor(setor, pdf, nomeBeneficiario) {
  try { ensureOutboxSheet_(); ensureQueueTrigger_(); } catch (_){}

  const toSetor = obterEmailDoSetor_(setor);
  if (!toSetor) { Logger.log('[EMAIL] Setor sem e-mail: ' + setor); return; }

  const assunto = 'Nova Ficha de Benefício - ' + (nomeBeneficiario || '');
  const corpo   = 'Segue em anexo a ficha preenchida do beneficiário: ' + (nomeBeneficiario || '');

  let fileId = null;
  try {
    if (pdf?.file?.getId) fileId = pdf.file.getId();
    else if (pdf?.fileId) fileId = pdf.fileId;
  } catch(_){}

  if (!fileId && pdf && pdf.blob){
    try {
      const f = DriveApp.createFile(pdf.blob);
      fileId = f.getId();
      setPublicSharing_(f);
    } catch(e){
      Logger.log('[EMAIL] blob->file: '+e.message);
    }
  }
  if (!fileId) { Logger.log('[EMAIL] sem fileId'); return; }

  const okSetor = tentarEnviarAgora_(toSetor, assunto, corpo, fileId, pdf && pdf.blob);
  if (!okSetor) enqueueEmail_(toSetor, assunto, corpo, fileId, setor, nomeBeneficiario);

  if (COPIA_EMAIL && validarEmail_(COPIA_EMAIL)) {
    const okCopia = tentarEnviarAgora_(COPIA_EMAIL, assunto + ' (cópia)', corpo, fileId, null);
    if (!okCopia) enqueueEmail_(COPIA_EMAIL, assunto + ' (cópia)', corpo, fileId, setor, nomeBeneficiario);
  }
}

function validarEmail_(e){
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(e||'').trim());
}

function obterEmailDoSetor_(setor){
  try{
    const planilha = SpreadsheetApp.openById(PLANILHA_ID);
    const aba = planilha.getSheetByName('Login');
    if (!aba) return '';
    const dados = aba.getDataRange().getValues();
    const wantedKey = canonicalSectorKey_(setor);
    for (let i=1;i<dados.length;i++){
      const setLabel = cleanSectorLabel_(dados[i][0] || '');
      const key = canonicalSectorKey_(setLabel);
      if (key === wantedKey){
        const em = (dados[i][3] || '').toString().trim();
        if (validarEmail_(em)) return em;
      }
    }
  }catch(e){
    Logger.log('[EMAIL] obter email setor: ' + e.message);
  }
  return '';
}

function tentarEnviarAgora_(to, subject, body, fileId, blobOpt){
  if (!validarEmail_(to)) { Logger.log('[EMAIL] inválido: ' + to); return false; }
  const quota = MailApp.getRemainingDailyQuota();
  if (quota <= 0) { Logger.log('[EMAIL] Sem quota diária'); return false; }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e){ Logger.log('[EMAIL] lock falhou: ' + e.message); }

  try {
    for (let tent=1; tent<=2; tent++){
      try{
        const attachments = [];
        if (blobOpt) attachments.push(blobOpt);
        else attachments.push(DriveApp.getFileById(fileId).getAs(MimeType.PDF));

        if (tent>1) Utilities.sleep(600 * tent);

        try{
          MailApp.sendEmail({ to, subject, body, attachments, name:'SEMFAS Sistema' });
          Logger.log('[EMAIL] MailApp OK -> ' + to);
          return true;
        }catch(e1){
          Logger.log('[EMAIL] MailApp falhou: ' + e1.message);
          GmailApp.sendEmail(to, subject, body, { attachments, name:'SEMFAS Sistema' });
          Logger.log('[EMAIL] GmailApp OK -> ' + to);
          return true;
        }
      }catch(e){
        Logger.log('[EMAIL] tentativa ' + tent + ': ' + e.message);
        if (tent===2) return false;
      }
    }
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
  return false;
}

function ensureOutboxSheet_(){
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName(OUTBOX_SHEET_NAME);
  if (!sh){
    sh = ss.insertSheet(OUTBOX_SHEET_NAME);
    sh.getRange(1,1,1,10).setValues([[
      'Timestamp','To','Assunto','Corpo','FileId','Tentativas','Status','UltimoErro','Setor','Beneficiario'
    ]]);
  }
  return sh;
}

function enqueueEmail_(to, subject, body, fileId, setor, beneficiario){
  const sh = ensureOutboxSheet_();
  sh.appendRow([ new Date(), to, subject, body, fileId, 0, 'PENDING', '', setor||'', beneficiario||'' ]);
  ensureQueueTrigger_();
  Logger.log('[EMAIL] enfileirado -> ' + to);
}

function processEmailQueue(){
  const sh = ensureOutboxSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rng = sh.getRange(2,1,lastRow-1,10).getValues();
  const out = [];
  const quota = MailApp.getRemainingDailyQuota();
  if (quota <= 0) { Logger.log('[EMAIL] Sem quota diária'); return; }

  for (let i=0;i<rng.length;i++){
    let [ts, to, subject, body, fileId, tries, status, lastErr, setor, benef] = rng[i];
    if (status === 'SENT') { out.push(rng[i]); continue; }
    if (tries >= OUTBOX_MAX_TRIES) { out.push([ts,to,subject,body,fileId,tries,'ERROR',lastErr,setor,benef]); continue; }

    let ok=false, errMsg='';
    try{
      const file = DriveApp.getFileById(fileId);
      const attachments = [ file.getAs(MimeType.PDF) ];
      try { MailApp.sendEmail({ to, subject, body, attachments, name:'SEMFAS Sistema' }); ok=true; }
      catch(e1){ GmailApp.sendEmail(to, subject, body, { attachments, name:'SEMFAS Sistema' }); ok=true; }
    }catch(e){
      ok=false; errMsg = e.message || String(e);
    }

    if (ok) out.push([ts,to,subject,body,fileId,tries+1,'SENT','',setor,benef]);
    else    out.push([ts,to,subject,body,fileId,tries+1,'PENDING',errMsg,setor,benef]);
  }
  sh.getRange(2,1,out.length,10).setValues(out);
}

function ensureQueueTrigger_(){
  const fn = 'processEmailQueue';
  const triggers = ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===fn);
  if (triggers.length===0) ScriptApp.newTrigger(fn).timeBased().everyMinutes(5).create();
}

/************************
 * DIAGNÓSTICO
 ************************/
function baixa_ping(){
  const { sh, idx } = getRespostasIndexMap_();
  return { sheet: sh.getName(), rows: sh.getLastRow() - 1, cols: sh.getLastColumn(), idx };
}

/** =========================
 *  PROTOCOLO AUTOMÁTICO
 *  =========================*/
function nextProtocolo_(){
  const year = new Date().getFullYear();
  const key = 'PROTO_SEQ_' + year;
  const sp = PropertiesService.getScriptProperties();
  let n = parseInt(sp.getProperty(key) || '0', 10);
  n++;
  sp.setProperty(key, String(n));
  return `BEV-${year}-${String(n).padStart(6,'0')}`;
}

function getSequenceNumber(key){
  const sp = PropertiesService.getScriptProperties();
  const k = String(key || 'SEQ').trim();
  let n = parseInt(sp.getProperty(k) || '0', 10);
  if (isNaN(n)) n = 0;
  n++;
  sp.setProperty(k, String(n));
  return n;
}

function ensureProtocoloForRow_(sh, rowIndex){
  const iProt = ensureColumn_(sh, PROTOCOLO_HDR);
  const val = sh.getRange(rowIndex, iProt+1).getValue();
  if (!String(val||'').trim()){
    const proto = nextProtocolo_();
    sh.getRange(rowIndex, iProt+1).setValue(proto);
    return proto;
  }
  return val;
}

function backfillProtocolos(){
  const { sh } = getRespostasIndexMap_();
  const iProt = ensureColumn_(sh, PROTOCOLO_HDR);
  const last = sh.getLastRow();
  if (last < 2) return { preenchidos: 0 };
  const rng = sh.getRange(2, iProt+1, last-1, 1);
  const vals = rng.getValues();
  let count = 0;
  for (let i=0;i<vals.length;i++){
    if (!String(vals[i][0]||'').trim()){
      vals[i][0] = nextProtocolo_();
      count++;
    }
  }
  rng.setValues(vals);
  return { preenchidos: count };
}

function garantirProtocolosPreenchidos(){ return backfillProtocolos(); }

/************************
 * DETALHES (Drawer Central)
 ************************/
function getRegistroCompleto(linha){
  if (!linha) throw new Error('Linha não informada.');
  const { sh, headers, idx } = getRespostasIndexMap_();
  const lastCol = sh.getLastColumn();
  const row = sh.getRange(linha, 1, 1, lastCol).getValues()[0];

  const toStr = v => {
    if (v instanceof Date) return formatDateBR(v);
    if (typeof v === 'number') return String(v);
    return v == null ? '' : String(v);
  };

  const cols = [];
  for (let i=0;i<lastCol;i++){
    const label = headers[i] ? String(headers[i]) : ('Coluna ' + (i+1));
    cols.push({ label, value: toStr(row[i]) });
  }

  const iProt = findColIdx_(headers, PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo','numero do protocolo','id','id protocolo');
  const protocolo = iProt >= 0 ? toStr(row[iProt]) : '';

  const status = isRowEntregue_(row, idx)
    ? 'ENTREGUE'
    : (String(idx.status>=0 ? row[idx.status] : 'PENDENTE').toUpperCase() || 'PENDENTE');

  const resumo = {
    linha,
    data: (function(){ const d = coerceSheetDate(row[idx.data]); return d ? formatDateBR(d) : ''; })(),
    unidade:  idx.unidade  >= 0 ? toStr(row[idx.unidade])  : '',
    demanda:  idx.beneficio>= 0 ? toStr(row[idx.beneficio]): '',
    nome:     idx.nome     >= 0 ? toStr(row[idx.nome])     : '',
    cpf:      idx.cpf      >= 0 ? formatCPF(toStr(row[idx.cpf])) : '',
    status,
    protocolo
  };

  return { linha, resumo, cols };
}

/************************
 * ATUALIZAR STATUS (Central)
 ************************/
function atualizarStatus(linha, status){
  status = String(status||'').trim().toUpperCase();
  if (!linha || !status) throw new Error('Parâmetros inválidos.');

  const { sh, idx } = getRespostasIndexMap_();

  let iStatus = idx.status;
  if (iStatus < 0) iStatus = ensureColumn_(sh, 'Status');

  const iEnt    = idx.entregueFlag;
  const iEntEm  = ensureColumn_(sh, ENTREGUE_EM_HDR);

  const rowVals = sh.getRange(linha, 1, 1, sh.getLastColumn()).getValues()[0];

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try{
    sh.getRange(linha, iStatus+1).setValue(status);
    if (status === 'ENTREGUE'){
      sh.getRange(linha, iEnt+1).setValue('ENTREGUE');
      if (!rowVals[iEntEm]) sh.getRange(linha, iEntEm+1).setValue(new Date());
    } else {
      sh.getRange(linha, iEnt+1).setValue('');
    }
  } finally { lock.releaseLock(); }
  return true;
}

/************************
 * PARECER TÉCNICO — atualizar/ler
 ************************/
function atualizarParecer(linha, payload){
  if (!linha) throw new Error('Linha não informada.');
  payload = payload || {};
  const { sh } = getRespostasIndexMap_();

  const iPar  = ensureColumn_(sh, 'PARECER TÉCNICO');
  const iNomT = ensureColumn_(sh, 'NOME TÉCNICO');
  const iMun  = ensureColumn_(sh, 'MUNICÍPIO PARECER');
  const iData = ensureColumn_(sh, 'DATA PARECER');
  const iAssB = ensureColumn_(sh, 'ASSINATURA BENEFICIÁRIO');
  const iAssT = ensureColumn_(sh, 'ASSINATURA/CARIMBO');

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    if (payload.parecer != null) sh.getRange(linha, iPar+1).setValue(String(payload.parecer||''));
    if (payload.nome_tecnico != null) sh.getRange(linha, iNomT+1).setValue(String(payload.nome_tecnico||''));
    if (payload.municipio_parecer != null) sh.getRange(linha, iMun+1).setValue(String(payload.municipio_parecer||''));
    if (payload.data_parecer != null){
      const d = parseISODateSafe(payload.data_parecer);
      sh.getRange(linha, iData+1).setValue(d ? formatDateBR(d) : '');
    }
    if (payload.assinatura_beneficiario != null) sh.getRange(linha, iAssB+1).setValue(String(payload.assinatura_beneficiario||''));
    if (payload.assinatura_carimbo != null) sh.getRange(linha, iAssT+1).setValue(String(payload.assinatura_carimbo||''));
  } finally {
    lock.releaseLock();
  }
  return true;
}

function lerParecer(linha){
  if (!linha) throw new Error('Linha não informada.');
  const { sh } = getRespostasIndexMap_();

  const iPar  = ensureColumn_(sh, 'PARECER TÉCNICO');
  const iNomT = ensureColumn_(sh, 'NOME TÉCNICO');
  const iMun  = ensureColumn_(sh, 'MUNICÍPIO PARECER');
  const iData = ensureColumn_(sh, 'DATA PARECER');
  const iAssB = ensureColumn_(sh, 'ASSINATURA BENEFICIÁRIO');
  const iAssT = ensureColumn_(sh, 'ASSINATURA/CARIMBO');

  const vals = sh.getRange(linha, 1, 1, sh.getLastColumn()).getValues()[0];

  return {
    parecer:                 vals[iPar]  || '',
    nome_tecnico:            vals[iNomT] || '',
    municipio_parecer:       vals[iMun]  || '',
    data_parecer:            vals[iData] || '',
    assinatura_beneficiario: vals[iAssB] || '',
    assinatura_carimbo:      vals[iAssT] || ''
  };
}

/************************
 * PDF — utilidades
 ************************/
function obterLinkPdf(linha){
  const { sh } = getRespostasIndexMap_();
  if (!linha) throw new Error('Linha não informada.');
  const iPdf = ensureColumn_(sh, 'PDF');
  const link = sh.getRange(linha, iPdf+1).getValue();
  return { link: link || '' };
}
function regerarPdf(linha){
  if (!linha) throw new Error('Linha não informada.');
  return gerarPdfDeRegistro(linha);
}

/************************
 * ENTREGA — desfazer
 ************************/
function desfazerEntrega(linha){
  if (!linha) throw new Error('Linha não informada.');
  const { sh, idx } = getRespostasIndexMap_();

  const iEnt   = idx.entregueFlag;
  const iEntEm = ensureColumn_(sh, ENTREGUE_EM_HDR);
  const iEntPor= ensureColumn_(sh, ENTREGUE_POR_HDR);
  const iEntUni= ensureColumn_(sh, ENTREGUE_UNID_HDR);
  const iEntObs= ensureColumn_(sh, ENTREGUE_OBS_HDR);

  let iStatus  = idx.status;
  if (iStatus < 0) iStatus = ensureColumn_(sh, 'Status');

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try{
    sh.getRange(linha, iEnt+1).setValue('');
    sh.getRange(linha, iEntEm+1).setValue('');
    sh.getRange(linha, iEntPor+1).setValue('');
    sh.getRange(linha, iEntUni+1).setValue('');
    sh.getRange(linha, iEntObs+1).setValue('');
    sh.getRange(linha, iStatus+1).setValue('PENDENTE');
  } finally {
    lock.releaseLock();
  }
  return true;
}

/************************
 * HISTÓRICO por CPF
 ************************/
function historicoPorCPF(cpf){
  const docpf = normalizeCPF(cpf);
  if (!docpf) return { registros: [] };

  const { sh, headers, idx } = getRespostasIndexMap_();
  const vals = sh.getDataRange().getValues();

  const iProt   = findColIdx_(headers, PROTOCOLO_HDR,'protocolo','nº protocolo','n° protocolo','numero do protocolo','id','id protocolo');
  const iEntEm  = ensureColumn_(sh, ENTREGUE_EM_HDR);
  const iEntPor = ensureColumn_(sh, ENTREGUE_POR_HDR);
  const iEntUni = ensureColumn_(sh, ENTREGUE_UNID_HDR);
  const iEntObs = ensureColumn_(sh, ENTREGUE_OBS_HDR);
  const iPdf    = idx.pdf >= 0 ? idx.pdf : ensureColumn_(sh, 'PDF');
  const iDocs   = idx.docs;

  const out = [];
  for (let i = 1; i < vals.length; i++){
    const row = vals[i];

    const cpfCell = (idx.cpf >= 0 && idx.cpf < row.length) ? row[idx.cpf] : '';
    const cpfRow  = normalizeCPF(cpfCell);
    if (!cpfRow || cpfRow !== docpf) continue;

    const dSolic  = (idx.data >= 0 && idx.data < row.length) ? coerceSheetDate(row[idx.data]) : null;

    const dEntRaw = (iEntEm >= 0 && iEntEm < row.length) ? row[iEntEm] : null;
    const dEnt = dEntRaw ? parseAnyDate(dEntRaw) : null;

    const statusCell = (idx.status >= 0 && idx.status < row.length) ? row[idx.status] : 'PENDENTE';

    const status = isRowEntregue_(row, idx) ? 'ENTREGUE' : String(statusCell || 'PENDENTE').toUpperCase();

    out.push({
      linha: i+1,
      protocolo:       (iProt >=0   && iProt   < row.length) ? (row[iProt]      || '') : '',
      nome:            (idx.nome>=0 && idx.nome< row.length) ? (row[idx.nome]   || '') : '',
      unidade:         (idx.unidade>=0 && idx.unidade<row.length) ? (row[idx.unidade] || '') : '',
      beneficio:       (idx.beneficio>=0 && idx.beneficio<row.length) ? (row[idx.beneficio] || '') : '',
      data:            dSolic ? formatDateBR(dSolic) : '',
      status,
      entregue_em:     dEnt ? Utilities.formatDate(dEnt, TZ, 'dd/MM/yyyy HH:mm') : '',
      unidade_entrega: (iEntUni>=0 && iEntUni<row.length) ? (row[iEntUni] || '') : '',
      entregue_por:    (iEntPor>=0 && iEntPor<row.length) ? (row[iEntPor] || '') : '',
      obs_entrega:     (iEntObs>=0 && iEntObs<row.length) ? (row[iEntObs] || '') : '',
      pdf:             (iPdf   >=0 && iPdf   <row.length) ? (row[iPdf]    || '') : '',
      docs:            (iDocs  >=0 && iDocs  <row.length) ? (row[iDocs]   || '') : ''
    });
  }

  sortList_(out, 'data_desc');
  return { registros: out };
}

/************************
 * EXPORTAÇÃO CSV
 ************************/
function exportarCsvDashboard(params){
  params = params || {};
  const statusFilter   = params.status || 'todos';
  const unidadeFiltro  = params.unidade || '';
  const beneficioFiltro= params.beneficio || '';
  const order          = params.order || 'data_desc';
  const pdfOnly        = !!params.pdfOnly;
  const hasProto       = !!params.hasProto;
  const dataInicio     = params.dataInicio || '';
  const dataFim        = params.dataFim || '';

  const pack = baixa_list(statusFilter, unidadeFiltro, beneficioFiltro, order, pdfOnly, hasProto, dataInicio, dataFim);
  const registros = pack.registros || [];
  const sep = ';';

  const header = ['Linha','Protocolo','Nome','CPF','Benefício','Unidade','Solicitado em','Status','PDF'];
  const linhas = [header.join(sep)];
  registros.forEach(r=>{
    linhas.push([
      r.rowRef, r.protocolo, r.nome, r.cpf, r.beneficio, r.unidade, r.solicitadoEm, r.status, r.pdf
    ].map(x => (String(x||'').includes(sep)
      ? `"${String(x).replace(/"/g,'""')}"`
      : String(x||''))).join(sep));
  });

  const blob = Utilities.newBlob(linhas.join('\n'), 'text/csv', 'export-semfas.csv');
  const file = DriveApp.createFile(blob);
  setPublicSharing_(file);
  return { fileId: file.getId(), link: buildDriveLinks_(file.getId()).download };
}

/************************
 * SAÚDE / VERSÃO
 ************************/
function healthCheck(){
  const info = baixa_ping();
  return {
    ok: true,
    sheet: info.sheet,
    rows: info.rows,
    cols: info.cols,
    idx: info.idx,
    tz: TZ,
    now: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss")
  };
}
function versaoScript(){
  return { version: 'v2026.01.09-TI-INTEGRADO', planilha: PLANILHA_ID };
}

/************************
 * MENU (opcional)
 ************************/
function onOpen(){
  try{
    SpreadsheetApp.getUi()
      .createMenu('SEMFAS')
      .addItem('Backfill Protocolos','garantirProtocolosPreenchidos')
      .addItem('Health Check','healthCheck')
      .addItem('Processar Fila de E-mails','processEmailQueue')
      .addToUi();
  }catch(_){}
}

/************************
 * CORREÇÃO DE STATUS DA FILA
 ************************/
function corrigirStatusOutbox(){
  const sh = ensureOutboxSheet_();
  const lr = sh.getLastRow();
  if (lr < 2) return { corrigidos: 0 };
  const rng = sh.getRange(2,1,lr-1,10);
  const vals = rng.getValues();
  let c=0;
  for (let i=0;i<vals.length;i++){
    if (vals[i][6] === 'SENTE'){
      vals[i][6] = 'SENT';
      c++;
    }
  }
  rng.setValues(vals);
  return { corrigidos: c };
}

/************************
 * LOGIN — CRUD DE USUÁRIOS (Tela ANALISTA)
 ************************/
function ensureLoginSheet_(){
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName('Login');
  if (!sh){
    sh = ss.insertSheet('Login');
    sh.appendRow(['Setor','Usuário','Senha','E-mail','Perfil']);
  }
  return sh;
}

function listarUsuariosLogin(){
  const sh = ensureLoginSheet_();
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return { usuarios: [], setores: [] };

  const usuarios = [];
  const setSet = new Set();

  for (let i=1;i<vals.length;i++){
    const row = vals[i];
    const setorRaw = row[0] || '';
    const usuarioRaw = row[1] || '';
    const emailRaw = row[3] || '';
    const roleRaw  = row[4] || 'usuario';

    const setor = cleanSectorLabel_(setorRaw);
    const usuario = String(usuarioRaw||'').trim();
    const email   = String(emailRaw||'').trim();
    let role      = String(roleRaw||'usuario').trim().toLowerCase() || 'usuario';

    if (setor) setSet.add(setor);

    usuarios.push({ linha: i+1, setor, usuario, email, role });
  }

  const setores = Array.from(setSet).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
  return { usuarios, setores };
}

function criarUsuarioLogin(payload){
  payload = payload || {};
  let setor   = cleanSectorLabel_(payload.setor || '');
  let usuario = String(payload.usuario || '').trim();
  let senha   = String(payload.senha   || '').toString();
  let email   = String(payload.email   || '').trim();
  let role    = String(payload.role    || 'usuario').trim().toLowerCase() || 'usuario';

  if (!setor || !usuario || !senha){
    return { ok:false, msg:'Setor, usuário e senha são obrigatórios.' };
  }

  const sh = ensureLoginSheet_();
  const vals = sh.getDataRange().getValues();
  const setorKey = canonicalSectorKey_(setor);
  const userKey  = canonicalSectorKey_(usuario);

  for (let i=1;i<vals.length;i++){
    const sRow = cleanSectorLabel_(vals[i][0] || '');
    const uRow = String(vals[i][1] || '');
    if (canonicalSectorKey_(sRow) === setorKey && canonicalSectorKey_(uRow) === userKey){
      return { ok:false, msg:'Já existe um usuário com esse login para este setor.' };
    }
  }

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try{ sh.appendRow([setor, usuario, senha, email, role]); }
  finally { lock.releaseLock(); }
  return { ok:true };
}

function atualizarUsuarioLogin(payload){
  payload = payload || {};
  const linha = parseInt(payload.linha, 10);
  if (!linha || linha < 2){ return { ok:false, msg:'Linha inválida.' }; }

  let setor   = cleanSectorLabel_(payload.setor || '');
  let usuario = String(payload.usuario || '').trim();
  let email   = String(payload.email   || '').trim();
  let role    = String(payload.role    || 'usuario').trim().toLowerCase() || 'usuario';

  if (!setor || !usuario){ return { ok:false, msg:'Setor e usuário são obrigatórios.' }; }

  const sh = ensureLoginSheet_();
  const lastRow = sh.getLastRow();
  if (linha > lastRow){ return { ok:false, msg:'Linha fora do intervalo da planilha.' }; }

  const vals = sh.getDataRange().getValues();
  const setorKey = canonicalSectorKey_(setor);
  const userKey  = canonicalSectorKey_(usuario);

  for (let i=1;i<vals.length;i++){
    const rowIndex = i+1;
    if (rowIndex === linha) continue;
    const sRow = cleanSectorLabel_(vals[i][0] || '');
    const uRow = String(vals[i][1] || '');
    if (canonicalSectorKey_(sRow) === setorKey && canonicalSectorKey_(uRow) === userKey){
      return { ok:false, msg:'Já existe um usuário com esse login para este setor.' };
    }
  }

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    sh.getRange(linha,1).setValue(setor);
    sh.getRange(linha,2).setValue(usuario);
    sh.getRange(linha,4).setValue(email);
    sh.getRange(linha,5).setValue(role);
  } finally { lock.releaseLock(); }
  return { ok:true };
}

function alterarSenhaLogin(linha, novaSenha){
  const row = parseInt(linha, 10);
  if (!row || row < 2){ return { ok:false, msg:'Linha inválida.' }; }
  novaSenha = String(novaSenha || '').toString();
  if (!novaSenha){ return { ok:false, msg:'Senha não pode ser vazia.' }; }

  const sh = ensureLoginSheet_();
  const lastRow = sh.getLastRow();
  if (row > lastRow){ return { ok:false, msg:'Linha fora do intervalo da planilha.' }; }

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try { sh.getRange(row,3).setValue(novaSenha); }
  finally { lock.releaseLock(); }
  return { ok:true };
}

function resetarSenhaLogin(linha){
  const row = parseInt(linha, 10);
  if (!row || row < 2){ return { ok:false, msg:'Linha inválida.' }; }

  const sh = ensureLoginSheet_();
  const lastRow = sh.getLastRow();
  if (row > lastRow){ return { ok:false, msg:'Linha fora do intervalo da planilha.' }; }

  const nova = 'S' + Math.floor(100000 + Math.random()*900000);

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try { sh.getRange(row,3).setValue(nova); }
  finally { lock.releaseLock(); }
  return { ok:true, senha:nova };
}

function excluirUsuarioLogin(linha){
  const row = parseInt(linha, 10);
  if (!row || row < 2){ return { ok:false, msg:'Linha inválida.' }; }

  const sh = ensureLoginSheet_();
  const lastRow = sh.getLastRow();
  if (row > lastRow){ return { ok:false, msg:'Linha fora do intervalo da planilha.' }; }

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try { sh.deleteRow(row); }
  finally { lock.releaseLock(); }
  return { ok:true };
}

/************************
 * ============================
 * TI CHAMADOS — API (ESQUEMA DA SUA PLANILHA) ✅
 * Aba: "Chamados"
 * Colunas esperadas:
 * Carimbo | Protocolo | Nome | Email | Telefone | Setor/Local | Categoria | Prioridade | Descrição | Status | Responsável | Atualizado em | Obs | Anexo (Link) | Anexo (Nome)
 ************************/
const TI_SHEET_NAME = 'Chamados';
const TI_HIST_SHEET = 'Chamados_Historico';
const TI_CHAT_SHEET = 'Chamados_Chat';
const TI_PUBLIC_TOKEN_PROP = 'TI_PUBLIC_TOKEN';

function ensureTiSheets_(){
  try { ti__ensureSheets_(); return true; }
  catch(e){ Logger.log('[TI] ensureTiSheets_ erro: ' + (e && e.message ? e.message : e)); return false; }
}

function ti__ss_(){
  return SpreadsheetApp.openById(PLANILHA_ID);
}
function ti__sheet_(name){
  const ss = ti__ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ti__defaultHeaders_(){
  return [
    'Carimbo','Protocolo','Nome','Email','Telefone','Setor/Local',
    'Categoria','Prioridade','Descrição','Status','Responsável',
    'Atualizado em','Obs','Anexo (Link)','Anexo (Nome)','Anexo (Pasta ID)',
    'TI Última Leitura','Solicitante Última Leitura'
  ];
}

/** Versionamento de cache para listas TI */
function ti__cacheVer_(){
  try{ return CacheService.getScriptCache().get('ti_list_ver') || '0'; }catch(_){ return '0'; }
}
function ti__bumpCacheVer_(){
  try{ const c = CacheService.getScriptCache(); const v = parseInt(c.get('ti_list_ver')||'0',10)+1; c.put('ti_list_ver',String(v),600); }catch(_){}
}

function ti__ensureHeaders_(sh, headers){
  const lc = Math.max(1, sh.getLastColumn());
  let row1 = sh.getRange(1,1,1,lc).getValues()[0] || [];
  const existing = row1.map(x=>String(x||'').trim());

  const empty = existing.every(x=>!x);
  if (sh.getLastRow() === 0 || empty){
    sh.clear();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return headers;
  }

  const set = new Set(existing.map(x=>x.toLowerCase()));
  headers.forEach(h=>{
    const k = String(h).toLowerCase();
    if (!set.has(k)){
      sh.insertColumnAfter(sh.getLastColumn());
      sh.getRange(1, sh.getLastColumn()).setValue(h);
      set.add(k);
    }
  });

  const lc2 = sh.getLastColumn();
  row1 = sh.getRange(1,1,1,lc2).getValues()[0] || [];
  return row1;
}

function ti__idx_(headers, name){
  const h = (headers||[]).map(x=>String(x||'').trim().toLowerCase());
  return h.indexOf(String(name||'').trim().toLowerCase());
}

function ti__idxAny_(headers){
  for (let i = 1; i < arguments.length; i++) {
    const idx = ti__idx_(headers, arguments[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

function ti__extractUrls_(text){
  const s = String(text || '');
  if (!s) return [];
  const m = s.match(/https?:\/\/[^\s)"']+/g);
  return (m || []).map(u => String(u || '').trim()).filter(Boolean);
}

function ti__extractUrlsFromFormula_(formula){
  const src = String(formula || '').trim();
  if (!src) return [];
  const urls = [];
  const quoted = src.match(/"https?:\/\/[^\"]+"/g) || [];
  quoted.forEach((q) => {
    const u = String(q || '').replace(/^"|"$/g, '').trim();
    if (u) urls.push(u);
  });
  if (!urls.length) return ti__extractUrls_(src);
  return urls;
}

function ti__extractUrlsFromCell_(range){
  const found = [];
  const add = (u) => {
    const v = String(u || '').trim();
    if (v) found.push(v);
  };

  try {
    const rt = range.getRichTextValue();
    if (rt) {
      const root = rt.getLinkUrl();
      if (root) add(root);
      const runs = rt.getRuns() || [];
      runs.forEach((r) => {
        try {
          const u = r.getLinkUrl();
          if (u) add(u);
        } catch(_){ }
      });
    }
  } catch(_){ }

  try {
    const formula = range.getFormula();
    ti__extractUrlsFromFormula_(formula).forEach(add);
  } catch(_){ }

  try {
    const display = range.getDisplayValue();
    ti__extractUrls_(display).forEach(add);
  } catch(_){ }

  const uniq = [];
  const seen = Object.create(null);
  found.forEach((u) => {
    if (seen[u]) return;
    seen[u] = true;
    uniq.push(u);
  });
  return uniq;
}

function ti__parseAnexosFromCell_(raw){
  const src = String(raw || '').trim();
  if (!src) return { links: [], nomes: [] };

  const lines = src.split(/\r?\n|\s*;\s*/).map(x => String(x || '').trim()).filter(Boolean);
  const links = [];
  const nomes = [];

  lines.forEach((line) => {
    const urls = ti__extractUrls_(line);
    if (!urls.length) return;

    urls.forEach((url) => {
      links.push(url);
      let nome = line.replace(url, '').replace(/[\-–—:|]+$/g, '').trim();
      if (!nome) nome = '';
      nomes.push(nome);
    });
  });

  if (!links.length) {
    const urls = ti__extractUrls_(src);
    return { links: urls, nomes: urls.map(() => '') };
  }

  return { links, nomes };
}

function ti__collectUrlsFromRow_(sh, rowRef){
  const urls = [];
  const add = (u) => {
    const v = String(u || '').trim();
    if (v && /^https?:\/\//i.test(v)) urls.push(v);
  };

  try {
    const lc = sh.getLastColumn();
    const range = sh.getRange(rowRef, 1, 1, lc);

    const values = range.getDisplayValues()[0] || [];
    values.forEach((v) => ti__extractUrls_(v).forEach(add));

    const formulas = range.getFormulas()[0] || [];
    formulas.forEach((f) => ti__extractUrlsFromFormula_(f).forEach(add));

    const rich = range.getRichTextValues()[0] || [];
    rich.forEach((rt) => {
      try {
        if (!rt) return;
        const root = rt.getLinkUrl();
        if (root) add(root);
        const runs = rt.getRuns() || [];
        runs.forEach((r) => {
          try {
            const u = r.getLinkUrl();
            if (u) add(u);
          } catch(_){ }
        });
      } catch(_){ }
    });
  } catch(_){ }

  const uniq = [];
  const seen = Object.create(null);
  urls.forEach((u) => {
    if (seen[u]) return;
    seen[u] = true;
    uniq.push(u);
  });
  return uniq;
}

function ti__findDriveAnexosByProtocolo_(protocolo, maxItems){
  const id = String(protocolo || '').trim();
  if (!id) return { links: [], nomes: [] };

  const limit = Math.min(Math.max(parseInt(maxItems || 10, 10) || 10, 1), 30);
  const links = [];
  const nomes = [];
  const seen = Object.create(null);
  const idUp = id.toUpperCase();
  const idCompact = idUp.replace(/[^A-Z0-9]/g, '');
  const idNoTi = idUp.replace(/^TI-/, '');
  const idNoTiCompact = idNoTi.replace(/[^A-Z0-9]/g, '');

  const hasProto = (nome) => {
    const n = String(nome || '').trim();
    if (!n) return false;
    const nUp = n.toUpperCase();
    const nCompact = nUp.replace(/[^A-Z0-9]/g, '');
    if (nUp.indexOf(idUp) >= 0) return true;
    if (idCompact && nCompact.indexOf(idCompact) >= 0) return true;
    if (idNoTi && nUp.indexOf(idNoTi) >= 0) return true;
    if (idNoTiCompact && nCompact.indexOf(idNoTiCompact) >= 0) return true;
    return false;
  };

  const addFile = (file, acceptAnyName) => {
    try {
      if (!file) return;
      const nome = String(file.getName() || '').trim();
      if (!nome) return;
      if (!acceptAnyName && !hasProto(nome)) return;
      const fileId = file.getId();
      if (!fileId || seen[fileId]) return;
      seen[fileId] = true;
      try { setPublicSharing_(file); } catch(_){ }
      const driveLinks = buildDriveLinks_(fileId);
      links.push(driveLinks.view || file.getUrl());
      nomes.push(nome);
    } catch(_){ }
  };

  try {
    const root = DriveApp.getFolderById(FOLDER_PDFS_ID);
    const byProto = root.getFoldersByName(id);
    if (byProto.hasNext()) {
      const pFolder = byProto.next();
      const fit = pFolder.getFiles();
      let scanned = 0;
      while (fit.hasNext() && links.length < limit && scanned < 120) {
        scanned++;
        addFile(fit.next(), true);
      }
    }
  } catch(_){ }

  const runSearch = (token) => {
    const q = String(token || '').trim().replace(/'/g, "\\'");
    if (!q) return;
    try {
      const it = DriveApp.searchFiles(`trashed = false and title contains '${q}'`);
      let scanned = 0;
      while (it.hasNext() && links.length < limit && scanned < 220) {
        scanned++;
        addFile(it.next());
      }
    } catch(_){ }
  };

  try {
    runSearch(id);
    if (links.length < limit && idNoTi && idNoTi !== id) runSearch(idNoTi);
  } catch(_){ }

  if (!links.length) {
    try {
      const folder = DriveApp.getFolderById(FOLDER_PDFS_ID);
      const it = folder.getFiles();
      let scanned = 0;
      while (it.hasNext() && links.length < limit && scanned < 180) {
        scanned++;
        addFile(it.next());
      }
    } catch(_){ }
  }

  return { links, nomes };
}

function ti_listarAnexosDrive(input){
  input = input || {};
  const protocolo = String(input.id || input.protocolo || '').trim();
  if (!protocolo) return { ok:false, msg:'Protocolo obrigatório.' };

  const maxItems = Math.min(Math.max(parseInt(input.limit || input.max || 20, 10) || 20, 1), 30);
  let folderId = String(input.folderId || input.pastaId || '').trim();

  if (!folderId) {
    try {
      const { sh, headers } = ti__ensureSheets_();
      const rowRef = ti__findRowByProtocolo_(sh, headers, protocolo);
      if (rowRef > 0) {
        const iFld = ti__idxAny_(headers, 'Anexo (Pasta ID)', 'Anexo Pasta ID', 'Pasta Anexo ID', 'Pasta ID');
        if (iFld >= 0) folderId = String(sh.getRange(rowRef, iFld + 1).getDisplayValue() || '').trim();
      }
    } catch(_){ }
  }

  let drivePack = { links: [], nomes: [] };
  if (folderId) drivePack = ti__listDriveFilesByFolderId_(folderId, maxItems);
  if (!drivePack.links || !drivePack.links.length) {
    drivePack = ti__findDriveAnexosByProtocolo_(protocolo, maxItems);
  }

  return {
    ok: true,
    id: protocolo,
    anexoFolderId: folderId || '',
    anexoLinks: drivePack.links || [],
    anexoNomes: drivePack.nomes || [],
    total: (drivePack.links || []).length
  };
}

function ti__chatNormalize_(s){
  return String(s || '').replace(/\s+/g, ' ').trim();
}

function ti__chatFindRecentDuplicate_(chat, idx, payload){
  const protocolo = String(payload.protocolo || '').trim();
  const autorTipo = String(payload.autorTipo || '').trim().toUpperCase();
  const autorNome = String(payload.autorNome || '').trim().toLowerCase();
  const mensagem  = ti__chatNormalize_(payload.mensagem || '');
  const now = payload.now instanceof Date ? payload.now : new Date();
  const windowMs = Math.max(500, parseInt(payload.windowMs || 4000, 10) || 4000);

  if (!protocolo || !mensagem) return null;

  const lastRow = chat.getLastRow();
  if (lastRow < 2) return null;

  const scanRows = Math.min(lastRow - 1, 120);
  const startRow = lastRow - scanRows + 1;
  const vals = chat.getRange(startRow, 1, scanRows, chat.getLastColumn()).getValues();

  for (let i = vals.length - 1; i >= 0; i--) {
    const row = vals[i];
    if (String(row[idx.prot] || '').trim() !== protocolo) continue;

    const t = String(row[idx.tipo] || '').trim().toUpperCase();
    if (t !== autorTipo) continue;

    const n = String(row[idx.nome] || '').trim().toLowerCase();
    if (n !== autorNome) continue;

    const m = ti__chatNormalize_(row[idx.msg] || '');
    if (m !== mensagem) continue;

    const when = row[idx.car] instanceof Date ? row[idx.car] : parseAnyDate(row[idx.car]);
    if (!when) continue;

    const delta = now.getTime() - when.getTime();
    if (delta < 0 || delta > windowMs) continue;

    return {
      quando: Utilities.formatDate(when, TZ, "yyyy-MM-dd'T'HH:mm:ss"),
      protocolo,
      autorTipo: t,
      autorNome: String(row[idx.nome] || '').trim() || payload.autorNome || '',
      mensagem: String(row[idx.msg] || '').trim(),
      canal: String(row[idx.can] || '').trim() || String(payload.canal || 'WEB')
    };
  }

  return null;
}

function ti__findRowByProtocolo_(sh, headers, protocolo){
  const id = String(protocolo || '').trim();
  if (!id) return -1;
  const iProt = ti__idx_(headers, 'Protocolo');
  if (iProt < 0 || sh.getLastRow() < 2) return -1;
  try {
    const tf = sh.getRange(2, iProt + 1, sh.getLastRow() - 1, 1)
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext();
    if (tf) return tf.getRow();
  } catch(_){ }
  return -1;
}

function ti__ensureSheets_(){
  const sh = ti__sheet_(TI_SHEET_NAME);
  const headers = ti__ensureHeaders_(sh, ti__defaultHeaders_());

  const hist = ti__sheet_(TI_HIST_SHEET);
  ti__ensureHeaders_(hist, ['Carimbo','Protocolo','Ação','Detalhes','Usuário','Setor']);

  const chat = ti__sheet_(TI_CHAT_SHEET);
  ti__ensureHeaders_(chat, ['Carimbo','Protocolo','Autor Tipo','Autor Nome','Mensagem','Canal']);

  return { sh, headers, hist, chat };
}

function ti__ctx_(){
  try{
    const up = PropertiesService.getUserProperties().getProperties();
    return { usuario: up.semfas_usuario || '', setor: up.semfas_setor || '' };
  }catch(_){
    return { usuario:'', setor:'' };
  }
}

function ti__pushHist_(protocolo, acao, detalhes){
  const { hist } = ti__ensureSheets_();
  const c = ti__ctx_();
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try { hist.appendRow([ new Date(), protocolo, acao, detalhes || '', c.usuario, c.setor ]); }
  finally { lock.releaseLock(); }
}

function ti__nextProtocolo_(){
  const year = new Date().getFullYear();
  const key = 'TI_SEQ_' + year;
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sp = PropertiesService.getScriptProperties();
    let n = parseInt(sp.getProperty(key) || '0', 10);
    n++;
    sp.setProperty(key, String(n));
    return `TI-${year}-${String(n).padStart(6,'0')}`;
  } finally { lock.releaseLock(); }
}

function ti__splitContato_(contato){
  const s = String(contato||'').trim();
  const email = (s.match(/[^\s@]+@[^\s@]+\.[^\s@]+/)||[])[0] || '';
  const tel = (s.replace(/[^\d()+\- ]/g,'').trim()) || '';
  return { email, tel };
}

function ti__normalizePrioridade_(p){
  const v = String(p||'').trim().toLowerCase();
  if (v === 'urgente') return 'URGENTE';
  if (v === 'alta') return 'ALTA';
  if (v === 'media' || v === 'média') return 'NORMAL';
  if (v === 'baixa') return 'NORMAL';
  return String(p||'NORMAL').trim().toUpperCase() || 'NORMAL';
}

function ti__cleanText_(v, maxLen){
  const lim = Math.max(1, parseInt(maxLen || 500, 10));
  return String(v || '').replace(/\s+/g, ' ').trim().slice(0, lim);
}

function ti__autorTipo_(v){
  const raw = String(v || '').trim().toUpperCase();
  if (raw === 'TI') return 'TI';
  if (raw === 'SISTEMA') return 'SISTEMA';
  return 'SOLICITANTE';
}

function ti__validarTokenPublico_(input){
  const sp = PropertiesService.getScriptProperties();
  const esperado = String(sp.getProperty(TI_PUBLIC_TOKEN_PROP) || '').trim();
  if (!esperado) return { ok: true, tokenRequired: false };

  const recebido = String(input.token || input.publicToken || input.apiToken || '').trim();
  if (recebido && recebido === esperado) return { ok: true, tokenRequired: true };
  return { ok: false, msg: 'Token inválido para API pública TI.' };
}

function ti_public_meta(input){
  input = input || {};
  const check = ti__validarTokenPublico_(input);
  if (!check.ok) return check;

  let setores = [];
  try { setores = listarSetores() || []; }
  catch(_) { setores = []; }

  return {
    ok: true,
    tokenRequired: !!check.tokenRequired,
    prioridades: ['NORMAL', 'ALTA', 'URGENTE'],
    categorias: ['Hardware', 'Software', 'Rede', 'Impressora', 'Sistema', 'Acesso', 'Outro'],
    setores,
    statusInicial: 'ABERTO',
    chat: {
      enabled: true,
      maxMensagem: 2000
    }
  };
}

function ti_public_create(payload){
  payload = payload || {};
  const check = ti__validarTokenPublico_(payload);
  if (!check.ok) return check;

  const res = ti_abrirChamado(payload);
  if (!res || !res.ok) return res || { ok:false, msg:'Falha ao abrir chamado.' };

  return {
    ok: true,
    protocolo: res.protocolo,
    anexo: !!res.anexo,
    anexos: res.anexos || 0,
    anexoLinks: res.anexoLinks || [],
    anexoNomes: res.anexoNomes || [],
    chatHabilitado: true
  };
}

function ti_chat_enviar(input){
  input = input || {};

  const protocolo = String(input.id || input.protocolo || input.chamadoId || '').trim();
  if (!protocolo) return { ok:false, msg:'Protocolo obrigatório.' };

  const mensagem = ti__cleanText_(input.mensagem || input.msg || input.texto || '', 2000);
  if (!mensagem) return { ok:false, msg:'Mensagem obrigatória.' };

  const { sh, headers, chat } = ti__ensureSheets_();
  const chamadoRow = ti__findRowByProtocolo_(sh, headers, protocolo);
  if (chamadoRow < 0) return { ok:false, msg:'Chamado não encontrado.' };

  const chatHeaders = chat.getRange(1,1,1,chat.getLastColumn()).getValues()[0] || [];
  const iCar = ti__idx_(chatHeaders,'Carimbo');
  const iProt= ti__idx_(chatHeaders,'Protocolo');
  const iTipo= ti__idx_(chatHeaders,'Autor Tipo');
  const iNome= ti__idx_(chatHeaders,'Autor Nome');
  const iMsg = ti__idx_(chatHeaders,'Mensagem');
  const iCan = ti__idx_(chatHeaders,'Canal');

  const autorTipo = ti__autorTipo_(input.autorTipo || input.tipoAutor);

  // Validar: somente sessão TI pode enviar como TI ou SISTEMA
  if (autorTipo === 'TI' || autorTipo === 'SISTEMA') {
    const ctx = ti__ctx_();
    const setor = String(ctx.setor || '').toUpperCase();
    if (setor && setor !== 'TI' && setor !== 'TECNOLOGIA' && setor !== 'INFORMATICA' && setor !== 'INFORMÁTICA' && setor !== 'T.I.') {
      return { ok: false, msg: 'Sem permissão para enviar como ' + autorTipo + '.' };
    }
  }

  const autorNome = ti__cleanText_(input.autorNome || input.nome || input.autor || '', 120) || (autorTipo === 'TI' ? 'Técnico' : 'Solicitante');
  const canal = ti__cleanText_(input.canal || 'WEB', 40).toUpperCase();
  const now = new Date();

  const row = new Array(chat.getLastColumn()).fill('');
  row[iCar] = now;
  row[iProt] = protocolo;
  row[iTipo] = autorTipo;
  row[iNome] = autorNome;
  row[iMsg] = mensagem;
  row[iCan] = canal;

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let duplicateItem = null;
  let appended = false;
  try{
    duplicateItem = ti__chatFindRecentDuplicate_(chat, {
      car: iCar,
      prot: iProt,
      tipo: iTipo,
      nome: iNome,
      msg: iMsg,
      can: iCan
    }, {
      protocolo,
      autorTipo,
      autorNome,
      mensagem,
      canal,
      now,
      windowMs: 4500
    });

    if (!duplicateItem) {
      chat.appendRow(row);
      appended = true;
    }
  }
  finally{ lock.releaseLock(); }

  if (duplicateItem) {
    return { ok: true, dedup: true, item: duplicateItem };
  }

  if (appended) ti__pushHist_(protocolo, 'CHAT_' + autorTipo, mensagem.slice(0,200));

  return {
    ok: true,
    item: {
      quando: Utilities.formatDate(now, TZ, "yyyy-MM-dd'T'HH:mm:ss"),
      protocolo,
      autorTipo,
      autorNome,
      mensagem,
      canal
    }
  };
}

function ti_chat_listar(input){
  input = input || {};
  const protocolo = String(input.id || input.protocolo || input.chamadoId || '').trim();
  if (!protocolo) return { ok:true, items: [] };

  const { sh, headers, chat } = ti__ensureSheets_();
  const iReadTI  = ti__idx_(headers, 'TI Última Leitura');
  const iReadSol = ti__idx_(headers, 'Solicitante Última Leitura');

  // Localizar linha do chamado na planilha Chamados
  const chamadoRow = ti__findRowByProtocolo_(sh, headers, protocolo);

  // Marcar leitura do lado que está consultando (grava na planilha Chamados)
  const leitor = String(input.leitor || '').trim().toUpperCase();
  if (chamadoRow > 0 && (leitor === 'TI' || leitor === 'SOLICITANTE')) {
    try {
      const col = leitor === 'TI' ? iReadTI : iReadSol;
      if (col >= 0) sh.getRange(chamadoRow, col + 1).setValue(new Date());
    } catch(_){}
  }

  // Ler timestamps de leitura de cada lado
  let readByTI = null, readBySolic = null;
  if (chamadoRow > 0) {
    try {
      if (iReadTI >= 0) {
        const v = sh.getRange(chamadoRow, iReadTI + 1).getValue();
        if (v instanceof Date) readByTI = v;
        else if (v) readByTI = new Date(v);
      }
      if (iReadSol >= 0) {
        const v = sh.getRange(chamadoRow, iReadSol + 1).getValue();
        if (v instanceof Date) readBySolic = v;
        else if (v) readBySolic = new Date(v);
      }
    } catch(_){}
  }

  const lastRow = chat.getLastRow();
  if (lastRow < 2) return { ok:true, items: [] };

  const chatHeaders = chat.getRange(1,1,1,chat.getLastColumn()).getValues()[0] || [];
  const iCar = ti__idx_(chatHeaders,'Carimbo');
  const iProt= ti__idx_(chatHeaders,'Protocolo');
  const iTipo= ti__idx_(chatHeaders,'Autor Tipo');
  const iNome= ti__idx_(chatHeaders,'Autor Nome');
  const iMsg = ti__idx_(chatHeaders,'Mensagem');
  const iCan = ti__idx_(chatHeaders,'Canal');

  const sinceRaw = String(input.since || '').trim();
  const sinceDate = sinceRaw ? (parseAnyDate(sinceRaw) || null) : null;
  const limit = Math.min(Math.max(parseInt(input.limit || 200, 10) || 200, 1), 1000);

  const totalRows = lastRow - 1;
  const scanRows = sinceDate
    ? Math.min(totalRows, Math.max(limit * 4, 240))
    : Math.min(totalRows, Math.max(limit * 6, 900));
  const startRow = lastRow - scanRows + 1;
  const vals = chat.getRange(startRow,1,scanRows,chat.getLastColumn()).getValues();
  const out = [];
  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    if (String(row[iProt] || '').trim() !== protocolo) continue;

    const quandoDate = row[iCar] instanceof Date ? row[iCar] : parseAnyDate(row[iCar]);
    if (sinceDate && quandoDate && quandoDate <= sinceDate) continue;

    const tipo = String(row[iTipo] || '');
    const tipoUp = tipo.toUpperCase();

    // Determinar se a mensagem foi lida pelo outro lado
    let lido = false;
    if (tipoUp === 'SOLICITANTE' && readByTI && quandoDate && quandoDate <= readByTI) lido = true;
    if (tipoUp === 'TI' && readBySolic && quandoDate && quandoDate <= readBySolic) lido = true;
    if (tipoUp === 'SISTEMA') lido = true;

    out.push({
      quando: quandoDate ? Utilities.formatDate(quandoDate, TZ, "yyyy-MM-dd'T'HH:mm:ss") : String(row[iCar] || ''),
      protocolo,
      autorTipo: tipo,
      autorNome: String(row[iNome] || ''),
      mensagem: String(row[iMsg] || ''),
      canal: String(row[iCan] || ''),
      lido
    });
  }

  out.sort((a,b)=>{
    const ta = parseAnyDate(a.quando) || new Date(0);
    const tb = parseAnyDate(b.quando) || new Date(0);
    return ta.getTime() - tb.getTime();
  });

  return { ok:true, items: out.slice(-limit) };
}

function ti_public_chat_send(input){
  input = input || {};
  const check = ti__validarTokenPublico_(input);
  if (!check.ok) return check;
  // Forçar tipo SOLICITANTE para chamadas públicas (segurança)
  input.autorTipo = 'SOLICITANTE';
  return ti_chat_enviar(input);
}

function ti_public_chat_list(input){
  input = input || {};
  const check = ti__validarTokenPublico_(input);
  if (!check.ok) return check;
  return ti_chat_listar(input);
}

// ✅ Chamados abertos no login
function ti_abrirChamado(payload){
  payload = payload || {};
  ensureTiSheets_();

  const nome = String(payload.nome || payload.solicitante || '').trim();
  const emailRaw = String(payload.email || '').trim();
  const email = emailRaw || 'naoinformado@semfas.local';
  const telefone = String(payload.telefone || payload.tel || '').trim();
  const setor = String(payload.setor || payload.unidade || '').trim();
  const categoria = String(payload.categoria || '').trim();
  const prioridade = ti__normalizePrioridade_(payload.prioridade || 'NORMAL');
  const descricao = String(payload.descricao || '').trim();

  if (!nome || !telefone || !setor || !categoria || !descricao){
    return { ok:false, msg:'Preencha todos os campos obrigatórios.' };
  }

  const { sh, headers } = ti__ensureSheets_();
  const iCar = ti__idx_(headers,'Carimbo');
  const iProt= ti__idx_(headers,'Protocolo');
  const iNom = ti__idx_(headers,'Nome');
  const iEm  = ti__idx_(headers,'Email');
  const iTel = ti__idx_(headers,'Telefone');
  const iUni = ti__idx_(headers,'Setor/Local');
  const iCat = ti__idx_(headers,'Categoria');
  const iPr  = ti__idx_(headers,'Prioridade');
  const iDes = ti__idx_(headers,'Descrição');
  const iSt  = ti__idx_(headers,'Status');
  const iResp= ti__idx_(headers,'Responsável');
  const iUp  = ti__idx_(headers,'Atualizado em');
  const iObs = ti__idx_(headers,'Obs');
  const iAnx = ti__idx_(headers,'Anexo (Link)');
  const iAnxN= ti__idx_(headers,'Anexo (Nome)');
  const iAnxFld = ti__idxAny_(headers,'Anexo (Pasta ID)','Anexo Pasta ID','Pasta Anexo ID','Pasta ID');

  const now = new Date();
  const protocolo = ti__nextProtocolo_();

  const anexosPack = [];
  const filePack = payload.file || payload.anexo || null;
  if (filePack && filePack.dataUrl) anexosPack.push(filePack);

  if (Array.isArray(payload.files)) {
    payload.files.forEach(f => {
      if (f && f.dataUrl) anexosPack.push(f);
    });
  }
  if (Array.isArray(payload.anexos)) {
    payload.anexos.forEach(f => {
      if (f && f.dataUrl) anexosPack.push(f);
    });
  }

  const anexoLinks = [];
  const anexoNomes = [];
  let anexoFolderId = '';
  if (anexosPack.length){
    const folder = ti__getOrCreateProtocoloFolder_(protocolo);
    anexoFolderId = String(folder.getId() || '').trim();
    const limite = Math.min(anexosPack.length, 10);
    for (let i=0; i<limite; i++){
      const anexo = anexosPack[i];
      try{
        const blob = dataUrlToBlob_(anexo.dataUrl, anexo.mimeType || MimeType.BINARY);
        if (!blob) continue;

        const baseName = String(anexo.name || ('anexo_' + (i+1)));
        const safeName = baseName.replace(/[^a-zA-Z0-9._\- ]/g,'_');
        const file = folder.createFile(blob).setName(safeName);
        setPublicSharing_(file);

        const links = buildDriveLinks_(file.getId());
        anexoLinks.push(links.view || file.getUrl());
        anexoNomes.push(file.getName());
      }catch(e){
        Logger.log('[TI] anexo erro: ' + (e && e.message ? e.message : e));
      }
    }
  }

  const row = new Array(sh.getLastColumn()).fill('');
  row[iCar]  = now;
  row[iProt] = protocolo;
  row[iNom]  = nome;
  row[iEm]   = email;
  row[iTel]  = telefone;
  row[iUni]  = setor;
  row[iCat]  = categoria;
  row[iPr]   = prioridade;
  row[iDes]  = descricao;
  row[iSt]   = 'ABERTO';
  row[iResp] = '';
  row[iUp]   = now;
  row[iObs]  = anexoLinks.length ? `Aberto pelo portal • ${anexoLinks.length} anexo(s)` : 'Aberto pelo portal';
  if (iAnx >= 0) row[iAnx] = anexoLinks.join('\n');
  if (iAnxN >= 0) row[iAnxN] = anexoNomes.join('\n');
  if (iAnxFld >= 0) row[iAnxFld] = anexoFolderId;

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{ sh.appendRow(row); }
  finally{ lock.releaseLock(); }

  ti__pushHist_(protocolo, 'CRIAR', descricao.slice(0,200));
  ti__bumpCacheVer_();

  return { ok:true, protocolo, anexo: anexoLinks.length > 0, anexos: anexoLinks.length, anexoLinks, anexoNomes };
}

/** Boot pro front */
function ti_boot(){
  // Cache de 60 segundos
  const cache = CacheService.getScriptCache();
  const cacheKey = 'ti_boot_' + Session.getActiveUser().getEmail();
  try{
    const cached = cache.get(cacheKey);
    if(cached) return JSON.parse(cached);
  }catch(e){}
  
  ensureTiSheets_();
  const ctx = ti__ctx_();
  const tecnicoAtual = ti_getTecnicoAtual_();
  const result = {
    ok: true,
    ctx,
    tecnicoAtual,
    tz: TZ,
    now: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss")
  };
  
  try{
    cache.put(cacheKey, JSON.stringify(result), 60);
  }catch(e){}
  
  return result;
}

/** PING leve — retorna apenas contagens para polling eficiente */
function ti_ping(){
  const cache = CacheService.getScriptCache();
  const cacheKey = 'ti_ping_v1';
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch(_) {}

  const { sh, headers } = ti__ensureSheets_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return { ok: true, total: 0, abertos: 0, andamento: 0, pendentes: 0, urgentes: 0, lastId: '', ts: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss") };
  }

  const lc = sh.getLastColumn();
  const iSt   = ti__idx_(headers, 'Status');
  const iPr   = ti__idx_(headers, 'Prioridade');
  const iProt = ti__idx_(headers, 'Protocolo');
  const iUp   = ti__idx_(headers, 'Atualizado em');

  const scanRows = Math.min(lastRow - 1, 500);
  const startRow = lastRow - scanRows + 1;
  const values = sh.getRange(startRow, 1, scanRows, lc).getValues();

  let total = 0, abertos = 0, andamento = 0, pendentes = 0, urgentes = 0;
  let lastId = '';
  let lastTs = null;

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const st = String(row[iSt] || '').trim().toUpperCase();
    const pr = String(row[iPr] || '').trim().toUpperCase();
    total++;
    if (st === 'ABERTO') abertos++;
    else if (st === 'EM ANDAMENTO') andamento++;
    else if (st === 'PENDENTE' || st === 'AGUARDANDO') pendentes++;
    if (pr === 'URGENTE' || pr === 'ALTA') urgentes++;
    const upDate = row[iUp] instanceof Date ? row[iUp] : parseAnyDate(row[iUp]);
    if (!lastTs || (upDate && upDate > lastTs)) {
      lastTs = upDate;
      lastId = String(row[iProt] || '').trim();
    }
  }

  const result = {
    ok: true, total, abertos, andamento, pendentes, urgentes,
    lastId,
    ts: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss")
  };

  try { cache.put(cacheKey, JSON.stringify(result), 8); } catch(_) {}
  return result;
}

function ti_getTecnicoAtual_(){
  try{
    const up = PropertiesService.getUserProperties();
    const t = String(up.getProperty('TI_TECNICO_ATUAL') || '').trim();
    if (t) return t;
  }catch(_){}
  const ctx = ti__ctx_();
  return ctx.usuario || '—';
}

function ti_setTecnicoAtual(input){
  input = input || {};
  const nome = String(input.nome || input.tecnico || input.value || '').trim();
  if (!nome) return { ok:false, msg:'Nome do técnico vazio.' };
  try{
    PropertiesService.getUserProperties().setProperty('TI_TECNICO_ATUAL', nome);
  }catch(e){
    return { ok:false, msg:'Falha ao salvar técnico: ' + e.message };
  }
  return { ok:true, tecnicoAtual: nome };
}

function ti_listarTecnicos(){
  const ctx = ti__ctx_();
  const lista = new Set();
  try{
    const sh = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Login');
    if (sh){
      const vals = sh.getDataRange().getValues();
      for (let i=1;i<vals.length;i++){
        const setor = String(vals[i][0]||'').trim();
        const user  = String(vals[i][1]||'').trim();
        const role  = String(vals[i][4]||'').trim().toLowerCase();

        const isTi = role === 'ti' || canonicalSectorKey_(setor) === 'ti' || canonicalSectorKey_(setor).includes('tecnologiainformacao');
        if (isTi && user) lista.add(user);
      }
    }
  }catch(e){
    Logger.log('[TI] listarTecnicos erro: ' + e.message);
  }
  if (ctx.usuario) lista.add(ctx.usuario);
  const tecnicos = Array.from(lista).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
  return { ok:true, tecnicos, tecnicoAtual: ti_getTecnicoAtual_() };
}

/**
 * LISTA CHAMADOS
 * params: { q, status, prioridade, unidade, categoria, limit }
 */
function ti_listarChamados(params){
  params = params || {};
  
  // Cache de 30 segundos para melhorar performance (versionado)
  const ver = ti__cacheVer_();
  const cacheKey = 'ti_lv' + ver + '_' + JSON.stringify(params).substring(0,200);
  const cache = CacheService.getScriptCache();
  try{
    const cached = cache.get(cacheKey);
    if(cached) return JSON.parse(cached);
  }catch(e){}
  
  const { sh, headers } = ti__ensureSheets_();

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok:true, items: [] };

  const lc = sh.getLastColumn();

  const iCar = ti__idx_(headers,'Carimbo');
  const iProt= ti__idx_(headers,'Protocolo');
  const iNom = ti__idx_(headers,'Nome');
  const iEm  = ti__idx_(headers,'Email');
  const iTel = ti__idx_(headers,'Telefone');
  const iUni = ti__idx_(headers,'Setor/Local');
  const iCat = ti__idx_(headers,'Categoria');
  const iPr  = ti__idx_(headers,'Prioridade');
  const iDes = ti__idx_(headers,'Descrição');
  const iSt  = ti__idx_(headers,'Status');
  const iResp= ti__idx_(headers,'Responsável');
  const iUp  = ti__idx_(headers,'Atualizado em');
  const iObs = ti__idx_(headers,'Obs');

  const q = String(params.q||'').trim().toLowerCase();
  const fStatus = String(params.status||'').trim().toUpperCase();
  const fPrio   = String(params.prioridade||'').trim().toUpperCase();
  const fUni    = String(params.unidade||'').trim().toLowerCase();
  const fCat    = String(params.categoria||'').trim().toLowerCase();
  const limit   = Math.min(Math.max(parseInt(params.limit||200,10) || 200, 1), 2000);
  const hasFilters = !!(q || fStatus || fPrio || fUni || fCat);

  const totalRows = lastRow - 1;
  let startRow = 2;
  let rowCount = totalRows;
  if (!hasFilters) {
    rowCount = Math.min(totalRows, Math.max(limit * 4, 800));
    startRow = lastRow - rowCount + 1;
  }

  const values = sh.getRange(startRow,1,rowCount,lc).getValues();

  const out = [];

  for (let r=0; r<values.length; r++){
    const row = values[r];

    const protocolo = String(row[iProt]||'').trim();
    const status    = String(row[iSt]||'').trim().toUpperCase();
    const prio      = String(row[iPr]||'').trim().toUpperCase();
    const unidade   = String(row[iUni]||'').trim();
    const categoria = String(row[iCat]||'').trim();
    const solicit   = String(row[iNom]||'').trim();
    const descricao = String(row[iDes]||'').trim();
    const obs       = String(row[iObs]||'').trim();
    const resp      = String(row[iResp]||'').trim();

    if (fStatus && status !== fStatus) continue;
    if (fPrio   && prio   !== fPrio) continue;
    if (fUni    && !unidade.toLowerCase().includes(fUni)) continue;
    if (fCat    && !categoria.toLowerCase().includes(fCat)) continue;

    if (q){
      const blob = `${protocolo} ${status} ${prio} ${unidade} ${categoria} ${solicit} ${descricao} ${obs}`.toLowerCase();
      if (!blob.includes(q)) continue;
    }

    const titulo = (descricao.split('\n')[0] || descricao || protocolo).slice(0,80);

    out.push({
      rowRef: startRow + r,
      id: protocolo,
      criadoEm: row[iCar] instanceof Date ? Utilities.formatDate(row[iCar], TZ, "yyyy-MM-dd'T'HH:mm:ss") : (row[iCar]||''),
      atualizadoEm: row[iUp] instanceof Date ? Utilities.formatDate(row[iUp], TZ, "yyyy-MM-dd'T'HH:mm:ss") : (row[iUp]||''),
      status: status || 'ABERTO',
      prioridade: prio || 'NORMAL',
      categoria,
      unidade,
      solicitante: solicit,
      contato: [String(row[iEm]||'').trim(), String(row[iTel]||'').trim()].filter(v => v && v !== 'naoinformado@semfas.local').join(' • '),
      titulo,
      descricao,
      responsavel: resp,
      previsao: '',
      concluidoEm: '',
      ultimaAcao: obs
    });
  }

  // ordena por atualizado desc
  out.sort((a,b)=>{
    const ta = parseAnyDate(a.atualizadoEm) || parseAnyDate(a.criadoEm) || new Date(0);
    const tb = parseAnyDate(b.atualizadoEm) || parseAnyDate(b.criadoEm) || new Date(0);
    return tb.getTime() - ta.getTime();
  });

  const result = { ok:true, items: out.slice(0, limit) };
  
  // Salva no cache por 30 segundos
  try{
    cache.put(cacheKey, JSON.stringify(result), 30);
  }catch(e){}
  
  return result;
}

/** GET chamado por Protocolo */
function ti_obterChamado(input){
  const id = (typeof input === 'object' && input) ? (input.id || input.ID || '') : String(input||'');
  const protocolo = String(id||'').trim();
  if (!protocolo) return { ok:false, msg:'ID vazio.' };

  const { sh, headers } = ti__ensureSheets_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok:false, msg:'Não encontrado.' };

  const iProt= ti__idx_(headers,'Protocolo');
  if (iProt < 0) return { ok:false, msg:'Coluna Protocolo não encontrada.' };

  const rowRef = ti__findRowByProtocolo_(sh, headers, protocolo);
  if (rowRef < 0) return { ok:false, msg:'Não encontrado.' };
  const row = sh.getRange(rowRef, 1, 1, sh.getLastColumn()).getValues()[0];

  const iCar = ti__idx_(headers,'Carimbo');
  const iNom = ti__idx_(headers,'Nome');
  const iEm  = ti__idx_(headers,'Email');
  const iTel = ti__idx_(headers,'Telefone');
  const iUni = ti__idx_(headers,'Setor/Local');
  const iCat = ti__idx_(headers,'Categoria');
  const iPr  = ti__idx_(headers,'Prioridade');
  const iDes = ti__idx_(headers,'Descrição');
  const iSt  = ti__idx_(headers,'Status');
  const iResp= ti__idx_(headers,'Responsável');
  const iUp  = ti__idx_(headers,'Atualizado em');
  const iObs = ti__idx_(headers,'Obs');
  const iAnx = ti__idxAny_(headers, 'Anexo (Link)', 'Anexo Link', 'Anexos');
  const iAnxN= ti__idxAny_(headers, 'Anexo (Nome)', 'Anexo Nome', 'Nome do Anexo');
  const iAnxFld = ti__idxAny_(headers, 'Anexo (Pasta ID)', 'Anexo Pasta ID', 'Pasta Anexo ID', 'Pasta ID');
  const iDocs= ti__idxAny_(headers, 'Documentos', 'Docs', 'Arquivos Anexos', 'Documentos/Anexos');

  const descricao = String(row[iDes]||'').trim();
  const status = String(row[iSt]||'').trim().toUpperCase();
  const prio = String(row[iPr]||'').trim().toUpperCase();
  const anexoFolderId = iAnxFld >= 0 ? String(row[iAnxFld] || '').trim() : '';
  let anexoLinks = [];
  let anexoNomes = [];

  if (iAnx >= 0) {
    const parsedLinks = ti__parseAnexosFromCell_(row[iAnx]);
    anexoLinks = parsedLinks.links;
    anexoNomes = parsedLinks.nomes;
  }

  if (iAnxN >= 0) {
    const nomesLinha = String(row[iAnxN] || '').split(/\r?\n/).map(x => String(x || '').trim()).filter(Boolean);
    if (nomesLinha.length) {
      anexoNomes = anexoLinks.map((_, i) => nomesLinha[i] || anexoNomes[i] || '');
    }
  }

  if (!anexoLinks.length && iDocs >= 0) {
    const parsedDocs = ti__parseAnexosFromCell_(row[iDocs]);
    anexoLinks = parsedDocs.links;
    anexoNomes = parsedDocs.nomes;
  }

  if (iAnx >= 0) {
    try {
      const urlsCell = ti__extractUrlsFromCell_(sh.getRange(rowRef, iAnx + 1));
      if (urlsCell.length) {
        const has = new Set(anexoLinks);
        urlsCell.forEach((u) => {
          if (has.has(u)) return;
          anexoLinks.push(u);
          anexoNomes.push('');
          has.add(u);
        });
      }
    } catch(_){ }
  }

  if (iDocs >= 0) {
    try {
      const urlsCell = ti__extractUrlsFromCell_(sh.getRange(rowRef, iDocs + 1));
      if (urlsCell.length) {
        const has = new Set(anexoLinks);
        urlsCell.forEach((u) => {
          if (has.has(u)) return;
          anexoLinks.push(u);
          anexoNomes.push('');
          has.add(u);
        });
      }
    } catch(_){ }
  }

  if (anexoLinks.length) {
    anexoNomes = anexoLinks.map((_, i) => {
      const n = String(anexoNomes[i] || '').trim();
      return n || ('Anexo ' + (i + 1));
    });
  }

  if (!anexoLinks.length) {
    const allUrls = ti__collectUrlsFromRow_(sh, rowRef);
    if (allUrls.length) {
      anexoLinks = allUrls;
      anexoNomes = allUrls.map((_, i) => 'Anexo ' + (i + 1));
    }
  }

  // IMPORTANTE: não faz varredura no Drive aqui para não travar abertura do modal.
  // A busca de anexos no Drive é feita sob demanda via API dedicada (ti_drive_anexos).

  const item = {
    rowRef,
    id: protocolo,
    criadoEm: row[iCar] instanceof Date ? Utilities.formatDate(row[iCar], TZ, "yyyy-MM-dd'T'HH:mm:ss") : (row[iCar]||''),
    atualizadoEm: row[iUp] instanceof Date ? Utilities.formatDate(row[iUp], TZ, "yyyy-MM-dd'T'HH:mm:ss") : (row[iUp]||''),
    status: status || 'ABERTO',
    prioridade: prio || 'NORMAL',
    categoria: String(row[iCat]||'').trim(),
    unidade: String(row[iUni]||'').trim(),
    solicitante: String(row[iNom]||'').trim(),
    contato: [String(row[iEm]||'').trim(), String(row[iTel]||'').trim()].filter(v => v && v !== 'naoinformado@semfas.local').join(' • '),
    titulo: (descricao.split('\n')[0] || descricao || protocolo).slice(0,80),
    descricao,
    responsavel: String(row[iResp]||'').trim(),
    previsao: '',
    concluidoEm: '',
    ultimaAcao: String(row[iObs]||'').trim(),
    anexoFolderId,
    anexoLinks,
    anexoNomes
  };

  return { ok:true, rowRef, item };
}

/** CREATE */
function ti_criarChamado(payload){
  payload = payload || {};

  // Validar campos mínimos
  const _solic = String(payload.solicitante||'').trim();
  const _titulo = String(payload.titulo||'').trim();
  const _desc = String(payload.descricao||'').trim();
  const _uni = String(payload.unidade||'').trim();
  if (!_solic) return { ok:false, msg:'Solicitante é obrigatório.' };
  if (!_titulo && !_desc) return { ok:false, msg:'Título ou Descrição é obrigatório.' };
  if (!_uni) return { ok:false, msg:'Unidade/Setor é obrigatório.' };

  const { sh, headers } = ti__ensureSheets_();

  const iCar = ti__idx_(headers,'Carimbo');
  const iProt= ti__idx_(headers,'Protocolo');
  const iNom = ti__idx_(headers,'Nome');
  const iEm  = ti__idx_(headers,'Email');
  const iTel = ti__idx_(headers,'Telefone');
  const iUni = ti__idx_(headers,'Setor/Local');
  const iCat = ti__idx_(headers,'Categoria');
  const iPr  = ti__idx_(headers,'Prioridade');
  const iDes = ti__idx_(headers,'Descrição');
  const iSt  = ti__idx_(headers,'Status');
  const iResp= ti__idx_(headers,'Responsável');
  const iUp  = ti__idx_(headers,'Atualizado em');
  const iObs = ti__idx_(headers,'Obs');

  const now = new Date();
  const protocolo = ti__nextProtocolo_();

  const contato = ti__splitContato_(payload.contato);

  const titulo = String(payload.titulo||'').trim();
  const desc   = String(payload.descricao||'').trim();
  const descricaoFinal = titulo && desc ? (titulo + '\n' + desc) : (desc || titulo);

  const row = new Array(sh.getLastColumn()).fill('');
  row[iCar]  = now;
  row[iProt] = protocolo;
  row[iNom]  = String(payload.solicitante||'').trim();
  row[iEm]   = contato.email;
  row[iTel]  = contato.tel;
  row[iUni]  = String(payload.unidade||'').trim();
  row[iCat]  = String(payload.categoria||'').trim();
  row[iPr]   = String(payload.prioridade||'NORMAL').trim().toUpperCase();
  row[iDes]  = descricaoFinal;
  row[iSt]   = String(payload.status||'ABERTO').trim().toUpperCase();
  row[iResp] = String(payload.responsavel||ti_getTecnicoAtual_()||'').trim();
  row[iUp]   = now;
  row[iObs]  = String(payload.ultimaAcao||'Criado').trim();

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{ sh.appendRow(row); }
  finally{ lock.releaseLock(); }

  ti__pushHist_(protocolo, 'CRIAR', titulo || desc || '');
  ti__bumpCacheVer_();
  return { ok:true, id: protocolo };
}

/** UPDATE */
function ti_atualizarChamado(payload){
  payload = payload || {};
  const protocolo = String(payload.id||'').trim();
  if (!protocolo) return { ok:false, msg:'ID inválido.' };

  const { sh, headers } = ti__ensureSheets_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok:false, msg:'Sem dados.' };

  const iProt= ti__idx_(headers,'Protocolo');
  const iNom = ti__idx_(headers,'Nome');
  const iEm  = ti__idx_(headers,'Email');
  const iTel = ti__idx_(headers,'Telefone');
  const iUni = ti__idx_(headers,'Setor/Local');
  const iCat = ti__idx_(headers,'Categoria');
  const iPr  = ti__idx_(headers,'Prioridade');
  const iDes = ti__idx_(headers,'Descrição');
  const iSt  = ti__idx_(headers,'Status');
  const iResp= ti__idx_(headers,'Responsável');
  const iUp  = ti__idx_(headers,'Atualizado em');
  const iObs = ti__idx_(headers,'Obs');

  const rowRef = ti__findRowByProtocolo_(sh, headers, protocolo);
  if (rowRef < 0) return { ok:false, msg:'Não encontrado.' };

  const contato = ti__splitContato_(payload.contato);

  const titulo = String(payload.titulo||'').trim();
  const desc   = String(payload.descricao||'').trim();
  const descricaoFinal = titulo && desc ? (titulo + '\n' + desc) : (desc || titulo);

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    if (payload.status != null)      sh.getRange(rowRef, iSt+1).setValue(String(payload.status).trim().toUpperCase());
    if (payload.prioridade != null)  sh.getRange(rowRef, iPr+1).setValue(String(payload.prioridade).trim().toUpperCase());
    if (payload.categoria != null)   sh.getRange(rowRef, iCat+1).setValue(String(payload.categoria||'').trim());
    if (payload.unidade != null)     sh.getRange(rowRef, iUni+1).setValue(String(payload.unidade||'').trim());
    if (payload.solicitante != null) sh.getRange(rowRef, iNom+1).setValue(String(payload.solicitante||'').trim());
    if (payload.contato != null){
      if (contato.email) sh.getRange(rowRef, iEm+1).setValue(contato.email);
      if (contato.tel)   sh.getRange(rowRef, iTel+1).setValue(contato.tel);
    }
    if (payload.titulo != null || payload.descricao != null){
      sh.getRange(rowRef, iDes+1).setValue(descricaoFinal);
    }
    if (payload.responsavel != null) sh.getRange(rowRef, iResp+1).setValue(String(payload.responsavel||'').trim());
    sh.getRange(rowRef, iUp+1).setValue(new Date());
    if (payload.ultimaAcao != null) sh.getRange(rowRef, iObs+1).setValue(String(payload.ultimaAcao||'').trim());
  } finally {
    lock.releaseLock();
  }

  ti__pushHist_(protocolo, 'ATUALIZAR', String(payload.ultimaAcao||payload.titulo||payload.descricao||'').slice(0,200));
  ti__bumpCacheVer_();
  return { ok:true };
}

/** HIST */
function ti_historicoChamado(input){
  const id = (typeof input === 'object' && input) ? (input.id || input.ID || '') : String(input||'');
  const protocolo = String(id||'').trim();
  if (!protocolo) return { ok:true, items: [] };

  const { hist } = ti__ensureSheets_();
  const lastRow = hist.getLastRow();
  if (lastRow < 2) return { ok:true, items: [] };

  const vals = hist.getRange(2,1,lastRow-1,hist.getLastColumn()).getValues();
  const out = [];
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][1]||'').trim() !== protocolo) continue;
    out.push({
      quando: vals[i][0] instanceof Date ? Utilities.formatDate(vals[i][0], TZ, 'dd/MM/yyyy HH:mm') : String(vals[i][0]||''),
      acao: String(vals[i][2]||''),
      detalhes: String(vals[i][3]||''),
      usuario: String(vals[i][4]||''),
      setor: String(vals[i][5]||'')
    });
  }
  out.sort((a,b)=> (parseAnyDate(b.quando)||0) - (parseAnyDate(a.quando)||0));
  return { ok:true, items: out.slice(0,500) };
}

/** REPORT */
function ti_relatorios(params){
  params = params || {};
  const repVer = ti__cacheVer_();
  const repKey = 'ti_rep_' + repVer + '_' + JSON.stringify({ i: params.inicio || '', f: params.fim || '' }).substring(0, 120);
  const repCache = CacheService.getScriptCache();
  try {
    const cached = repCache.get(repKey);
    if (cached) return JSON.parse(cached);
  } catch(_){ }

  const inicio = params.inicio ? parseISODateSafe(params.inicio) : null;
  const fim    = params.fim ? parseISODateSafe(params.fim) : null;

  const { sh, headers } = ti__ensureSheets_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok:true, total:0, porStatus:{}, porPrioridade:{}, porCategoria:{}, porUnidade:{}, porResponsavel:{}, serieDias:[], estatisticas:{} };

  const lc = sh.getLastColumn();
  const values = sh.getRange(2,1,lastRow-1,lc).getValues();

  const iCar = ti__idx_(headers,'Carimbo');
  const iSt  = ti__idx_(headers,'Status');
  const iPr  = ti__idx_(headers,'Prioridade');
  const iCat = ti__idx_(headers,'Categoria');
  const iUni = ti__idx_(headers,'Setor/Local');
  const iUp  = ti__idx_(headers,'Atualizado em');
  const iResp= ti__idx_(headers,'Responsável');
  const iProt= ti__idx_(headers,'Protocolo');
  const iNome= ti__idx_(headers,'Nome');
  let iTit = ti__idx_(headers,'Título');
  if (iTit < 0) iTit = ti__idx_(headers,'Titulo');

  const porStatus = {}, porPrioridade = {}, porCategoria = {}, porUnidade = {}, porResponsavel = {}, porDia = {};
  let total = 0;
  let temposResolucao = [];
  let temposAtendimento = [];
  const items = [];
  const tmaRespMap = {};
  const tmaUniMap = {};
  const resolCatMap = {};

  for (let r=0;r<values.length;r++){
    const cr = values[r][iCar] instanceof Date ? values[r][iCar] : parseAnyDate(values[r][iCar]);
    if (!cr) continue;

    if (inicio && cr < startOfDay(inicio)) continue;
    if (fim && cr > endOfDay(fim)) continue;

    total++;

    const st = String(values[r][iSt]||'').trim().toUpperCase() || 'SEM STATUS';
    const pr = String(values[r][iPr]||'').trim().toUpperCase() || 'SEM PRIORIDADE';
    const ca = String(values[r][iCat]||'').trim() || 'SEM CATEGORIA';
    const un = String(values[r][iUni]||'').trim() || 'SEM UNIDADE';
    const rp = String(values[r][iResp]||'').trim() || 'SEM RESPONSAVEL';

    porStatus[st] = (porStatus[st]||0)+1;
    porPrioridade[pr] = (porPrioridade[pr]||0)+1;
    porCategoria[ca] = (porCategoria[ca]||0)+1;
    porUnidade[un] = (porUnidade[un]||0)+1;
    porResponsavel[rp] = (porResponsavel[rp]||0)+1;

    const keyDia = Utilities.formatDate(cr, TZ, 'yyyy-MM-dd');
    porDia[keyDia] = (porDia[keyDia]||0)+1;

    // Calcular TMA (Tempo Médio de Atendimento)
    const up = values[r][iUp] instanceof Date ? values[r][iUp] : parseAnyDate(values[r][iUp]);
    if (up && cr){
      const diff = (up.getTime() - cr.getTime()) / (1000*60); // em minutos
      if (diff >= 0) temposAtendimento.push(diff);
      if (!tmaRespMap[rp]) tmaRespMap[rp] = [];
      tmaRespMap[rp].push(diff);
      if (!tmaUniMap[un]) tmaUniMap[un] = [];
      tmaUniMap[un].push(diff);
    }

    // Se status = CONCLUÍDO, calcular tempo de resolução
    let tempoResolucaoHoras = 0;
    if (st === 'CONCLUIDO' || st === 'CONCLUÍDO' || st === 'RESOLVIDO'){
      if (up && cr){
        const diff = (up.getTime() - cr.getTime()) / (1000*60*60); // em horas
        if (diff >= 0) {
          temposResolucao.push(diff);
          tempoResolucaoHoras = diff;
          if (!resolCatMap[ca]) resolCatMap[ca] = [];
          resolCatMap[ca].push(diff);
        }
      }
    }

    const protocolo = iProt >= 0 ? String(values[r][iProt]||'').trim() : '';
    const solicitante = iNome >= 0 ? String(values[r][iNome]||'').trim() : '';
    const titulo = iTit >= 0 ? String(values[r][iTit]||'').trim() : '';

    items.push({
      id: protocolo || '',
      status: st,
      prioridade: pr,
      categoria: ca,
      unidade: un,
      responsavel: rp,
      solicitante,
      titulo,
      criadoEm: cr instanceof Date ? Utilities.formatDate(cr, TZ, "yyyy-MM-dd'T'HH:mm:ss") : (cr||''),
      atualizadoEm: up instanceof Date ? Utilities.formatDate(up, TZ, "yyyy-MM-dd'T'HH:mm:ss") : (up||''),
      tmaMin: (up && cr) ? Math.max(0, Math.round(((up.getTime()-cr.getTime())/60000)*100)/100) : 0,
      resolucaoHoras: Math.round((tempoResolucaoHoras||0)*100)/100
    });
  }

  const dias = Object.keys(porDia).sort();
  const serieDias = dias.map(d=>({ dia:d, total: porDia[d] }));

  // Calcular estatísticas
  const calcEstatistica = (arr) => {
    if (arr.length === 0) return { media: 0, minimo: 0, maximo: 0, mediana: 0 };
    arr.sort((a,b)=>a-b);
    const media = arr.reduce((a,b)=>a+b,0) / arr.length;
    const minimo = arr[0];
    const maximo = arr[arr.length-1];
    const mid = Math.floor(arr.length/2);
    const mediana = arr.length % 2 === 0 ? (arr[mid-1]+arr[mid])/2 : arr[mid];
    return { media: Math.round(media*100)/100, minimo: Math.round(minimo*100)/100, maximo: Math.round(maximo*100)/100, mediana: Math.round(mediana*100)/100 };
  };

  const estatisticas = {
    tma: calcEstatistica(temposAtendimento), // em minutos
    tempoResolucao: calcEstatistica(temposResolucao), // em horas
    totalAberto: porStatus['ABERTO'] || 0,
    totalAndamento: porStatus['EM ANDAMENTO'] || 0,
    totalResolvido: (porStatus['CONCLUIDO']||0) + (porStatus['CONCLUÍDO']||0) + (porStatus['RESOLVIDO']||0),
    totalPendente: (porStatus['PENDENTE']||0) + (porStatus['AGUARDANDO']||0),
    taxaResolucao: total > 0 ? Math.round(((porStatus['CONCLUIDO']||0) + (porStatus['CONCLUÍDO']||0) + (porStatus['RESOLVIDO']||0)) / total * 10000)/100 : 0
  };

  const calcAvgMap = (mapObj) => {
    return Object.keys(mapObj || {}).map(k => {
      const arr = mapObj[k] || [];
      const media = arr.length ? (arr.reduce((a,b)=>a+b,0) / arr.length) : 0;
      return { nome: k, media: Math.round(media*100)/100, total: arr.length };
    }).sort((a,b)=>b.media - a.media);
  };

  const tmaPorResponsavel = calcAvgMap(tmaRespMap);
  const tmaPorUnidade = calcAvgMap(tmaUniMap);
  const resolucaoPorCategoria = calcAvgMap(resolCatMap);
  const rankingLocais = Object.keys(porUnidade || {}).map(k => ({ nome:k, total: porUnidade[k]||0 }))
    .sort((a,b)=>b.total - a.total);

  items.sort((a,b)=> (parseAnyDate(b.criadoEm)||0) - (parseAnyDate(a.criadoEm)||0));

  const detalhes = {
    items: items.slice(0, 500),
    tmaPorResponsavel,
    tmaPorUnidade,
    resolucaoPorCategoria,
    rankingLocais
  };

  const result = { ok:true, total, porStatus, porPrioridade, porCategoria, porUnidade, porResponsavel, serieDias, estatisticas, detalhes };
  try { repCache.put(repKey, JSON.stringify(result), 45); } catch(_){ }
  return result;
}

/** alias p/ apiDispatch */
function ti_list(params){ return ti_listarChamados(params); }
function ti_get(input){ return ti_obterChamado(input); }
function ti_create(payload){ return ti_criarChamado(payload); }
function ti_update(payload){ return ti_atualizarChamado(payload); }
function ti_hist(input){ return ti_historicoChamado(input); }
function ti_report(params){ return ti_relatorios(params); }
function ti_deleted_v2(){ return { ok:true, version:2 }; }

function maq_stats(){ return maq_estatisticasPorLocal(); }

// =========================================================================
// MÓDULO: MÁQUINAS (Cadastro e Rastreamento)
// =========================================================================

function maq__ensureSheets_(){
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName('Maquinas');
  
  if (!sh){
    sh = ss.insertSheet('Maquinas');
    const headers = ['ID','Código Governo','Tipo Máquina','Local Atual','Data Cadastro','Atualizado Em','Observações','Marca','Modelo'];
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1e293b').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  
  let shHist = ss.getSheetByName('Maquinas_Historico');
  if (!shHist){
    shHist = ss.insertSheet('Maquinas_Historico');
    const headers = ['Carimbo','ID Máquina','Local Anterior','Local Novo','Responsável','Motivo'];
    shHist.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1e293b').setFontColor('#fff');
    shHist.setFrozenRows(1);
  }
  
  let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const required = ['ID','Código Governo','Tipo Máquina','Local Atual','Data Cadastro','Atualizado Em','Observações','Marca','Modelo'];
  required.forEach(h=>{
    if (headers.indexOf(h) < 0) {
      const col = sh.getLastColumn() + 1;
      sh.insertColumnAfter(sh.getLastColumn());
      sh.getRange(1,col).setValue(h).setFontWeight('bold').setBackground('#1e293b').setFontColor('#fff');
      headers.push(h);
    }
  });

  return { sh, shHist, headers };
}

function maq__gerarID_(){
  const seq = getSequenceNumber('MAQ_SEQ');
  return 'MAQ-' + String(seq).padStart(6,'0');
}

function maq__idx_(headers, name){
  const idx = headers.indexOf(name);
  return idx >= 0 ? idx : -1;
}

function maq_listar(params){
  params = params || {};
  const { sh, headers } = maq__ensureSheets_();
  
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok:true, items:[] };
  
  const values = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  
  const iId = maq__idx_(headers,'ID');
  const iCod = maq__idx_(headers,'Código Governo');
  const iTipo = maq__idx_(headers,'Tipo Máquina');
  const iLocal = maq__idx_(headers,'Local Atual');
  const iCad = maq__idx_(headers,'Data Cadastro');
  const iUp = maq__idx_(headers,'Atualizado Em');
  const iObs = maq__idx_(headers,'Observações');
  const iMarca = maq__idx_(headers,'Marca');
  const iModelo = maq__idx_(headers,'Modelo');
  
  const q = String(params.q||'').trim().toLowerCase();
  const fLocal = String(params.local||'').trim().toLowerCase();
  
  const out = [];

  // Mapa de ultima movimentacao por maquina
  const { shHist } = maq__ensureSheets_();
  const histLast = {};
  const histLastRow = shHist.getLastRow();
  if (histLastRow > 1) {
    const hvals = shHist.getRange(2,1,histLastRow-1,shHist.getLastColumn()).getValues();
    for (let i=0;i<hvals.length;i++) {
      const row = hvals[i];
      const when = row[0] instanceof Date ? row[0] : parseAnyDate(row[0]);
      const idHist = String(row[1]||'').trim();
      if (!idHist || !when) continue;
      const cur = histLast[idHist];
      if (!cur || when > cur.when) {
        histLast[idHist] = { when, localNovo: String(row[3]||'').trim() };
      }
    }
  }
  
  for (let r=0; r<values.length; r++){
    const row = values[r];
    const id = String(row[iId]||'').trim();
    if(!id) continue;
    
    const codigo = String(row[iCod]||'').trim();
    const tipo = String(row[iTipo]||'').trim();
    const local = String(row[iLocal]||'').trim();
    const obs = String(row[iObs]||'').trim();
    const marca = iMarca >= 0 ? String(row[iMarca]||'').trim() : '';
    const modelo = iModelo >= 0 ? String(row[iModelo]||'').trim() : '';
    const cad = row[iCad];
    const cadDate = cad instanceof Date ? cad : parseAnyDate(cad);
    const up = row[iUp];
    const upDate = up instanceof Date ? up : parseAnyDate(up);
    const lastMove = histLast[id];
    const dataLocal = lastMove && lastMove.when ? lastMove.when : cadDate;
    
    if(fLocal && !local.toLowerCase().includes(fLocal)) continue;
    
    if(q){
      const blob = `${id} ${codigo} ${tipo} ${local} ${obs} ${marca} ${modelo}`.toLowerCase();
      if(!blob.includes(q)) continue;
    }
    
    out.push({
      rowRef: r+2,
      id,
      codigoGoverno: codigo,
      tipoMaquina: tipo,
      localAtual: local,
      cadastradoEm: cadDate ? Utilities.formatDate(cadDate, TZ, "yyyy-MM-dd'T'HH:mm:ss") : (cad||''),
      atualizadoEm: upDate ? Utilities.formatDate(upDate, TZ, "yyyy-MM-dd'T'HH:mm:ss") : (up||''),
      dataLocal: dataLocal ? Utilities.formatDate(dataLocal, TZ, "yyyy-MM-dd'T'HH:mm:ss") : '',
      marca,
      modelo,
      observacoes: obs
    });
  }
  
  // Ordena por atualizado desc
  out.sort((a,b)=>{
    const ta = parseAnyDate(a.atualizadoEm) || parseAnyDate(a.cadastradoEm) || new Date(0);
    const tb = parseAnyDate(b.atualizadoEm) || parseAnyDate(b.cadastradoEm) || new Date(0);
    return tb.getTime() - ta.getTime();
  });
  
  return { ok:true, items:out };
}

function maq_obter(input){
  const id = (typeof input === 'object' && input) ? (input.id || input.ID || '') : String(input||'');
  if(!id) return { ok:false, msg:'ID vazio.' };
  
  const pack = maq_listar({ q:id });
  const item = (pack.items||[]).find(x => x.id === id);
  
  if(!item) return { ok:false, msg:'Máquina não encontrada.' };
  return { ok:true, item };
}

function maq_criar(payload){
  payload = payload || {};
  const codigoGoverno = String(payload.codigoGoverno||'').trim();
  const tipoMaquina = String(payload.tipoMaquina||'').trim();
  const localAtual = String(payload.localAtual||'').trim();
  const marca = String(payload.marca||'').trim();
  const modelo = String(payload.modelo||'').trim();
  const observacoes = String(payload.observacoes||'').trim();
  
  if(!codigoGoverno) return { ok:false, msg:'Código de governo obrigatório.' };
  if(!tipoMaquina) return { ok:false, msg:'Tipo de máquina obrigatório.' };
  if(!localAtual) return { ok:false, msg:'Local obrigatório.' };
  
  const { sh, headers } = maq__ensureSheets_();
  const lock = LockService.getScriptLock();
  
  try{
    lock.waitLock(10000);
    
    const id = maq__gerarID_();
    const now = new Date();
    
    const row = new Array(headers.length).fill('');
    const iId = maq__idx_(headers,'ID');
    const iCod = maq__idx_(headers,'Código Governo');
    const iTipo = maq__idx_(headers,'Tipo Máquina');
    const iLocal = maq__idx_(headers,'Local Atual');
    const iCad = maq__idx_(headers,'Data Cadastro');
    const iUp = maq__idx_(headers,'Atualizado Em');
    const iObs = maq__idx_(headers,'Observações');
    const iMarca = maq__idx_(headers,'Marca');
    const iModelo = maq__idx_(headers,'Modelo');

    if (iId >= 0) row[iId] = id;
    if (iCod >= 0) row[iCod] = codigoGoverno;
    if (iTipo >= 0) row[iTipo] = tipoMaquina;
    if (iLocal >= 0) row[iLocal] = localAtual;
    if (iCad >= 0) row[iCad] = now;
    if (iUp >= 0) row[iUp] = now;
    if (iObs >= 0) row[iObs] = observacoes;
    if (iMarca >= 0) row[iMarca] = marca;
    if (iModelo >= 0) row[iModelo] = modelo;

    sh.appendRow(row);
    
    lock.releaseLock();
    return { ok:true, id, msg:'Máquina cadastrada com sucesso!' };
    
  }catch(e){
    if(lock && lock.hasLock()) lock.releaseLock();
    return { ok:false, msg:'Erro ao cadastrar: ' + e.message };
  }
}

function maq_atualizar(payload){
  payload = payload || {};
  const id = String(payload.id||'').trim();
  if(!id) return { ok:false, msg:'ID obrigatório.' };
  
  const { sh, headers } = maq__ensureSheets_();
  const pack = maq_obter(id);
  
  if(!pack.ok || !pack.item) return { ok:false, msg:'Máquina não encontrada.' };
  
  const rowRef = pack.item.rowRef;
  const iCod = maq__idx_(headers,'Código Governo');
  const iTipo = maq__idx_(headers,'Tipo Máquina');
  const iLocal = maq__idx_(headers,'Local Atual');
  const iUp = maq__idx_(headers,'Atualizado Em');
  const iObs = maq__idx_(headers,'Observações');
  const iMarca = maq__idx_(headers,'Marca');
  const iModelo = maq__idx_(headers,'Modelo');
  
  const lock = LockService.getScriptLock();
  
  try{
    lock.waitLock(10000);
    
    if(payload.codigoGoverno !== undefined && iCod >= 0){
      sh.getRange(rowRef, iCod+1).setValue(String(payload.codigoGoverno||'').trim());
    }
    if(payload.tipoMaquina !== undefined && iTipo >= 0){
      sh.getRange(rowRef, iTipo+1).setValue(String(payload.tipoMaquina||'').trim());
    }
    if(payload.localAtual !== undefined && iLocal >= 0){
      sh.getRange(rowRef, iLocal+1).setValue(String(payload.localAtual||'').trim());
    }
    if(payload.observacoes !== undefined && iObs >= 0){
      sh.getRange(rowRef, iObs+1).setValue(String(payload.observacoes||'').trim());
    }
    if(payload.marca !== undefined && iMarca >= 0){
      sh.getRange(rowRef, iMarca+1).setValue(String(payload.marca||'').trim());
    }
    if(payload.modelo !== undefined && iModelo >= 0){
      sh.getRange(rowRef, iModelo+1).setValue(String(payload.modelo||'').trim());
    }
    
    if(iUp >= 0){
      sh.getRange(rowRef, iUp+1).setValue(new Date());
    }
    
    lock.releaseLock();
    return { ok:true, msg:'Máquina atualizada!' };
    
  }catch(e){
    if(lock && lock.hasLock()) lock.releaseLock();
    return { ok:false, msg:'Erro ao atualizar: ' + e.message };
  }
}

function maq_excluir(payload){
  payload = payload || {};
  const id = String(payload.id||'').trim();
  if(!id) return { ok:false, msg:'ID obrigatório.' };

  const { sh } = maq__ensureSheets_();
  const pack = maq_obter(id);
  if(!pack.ok || !pack.item) return { ok:false, msg:'Máquina não encontrada.' };

  const rowRef = pack.item.rowRef;
  const lock = LockService.getScriptLock();

  try{
    lock.waitLock(10000);
    sh.deleteRow(rowRef);
    lock.releaseLock();
    return { ok:true, msg:'Máquina excluída!' };
  }catch(e){
    if(lock && lock.hasLock()) lock.releaseLock();
    return { ok:false, msg:'Erro ao excluir: ' + e.message };
  }
}

function maq_moverLocal(payload){
  payload = payload || {};
  const id = String(payload.id||'').trim();
  const novoLocal = String(payload.novoLocal||'').trim();
  const motivo = String(payload.motivo||'').trim();
  
  if(!id) return { ok:false, msg:'ID obrigatório.' };
  if(!novoLocal) return { ok:false, msg:'Novo local obrigatório.' };
  
  const pack = maq_obter(id);
  if(!pack.ok || !pack.item) return { ok:false, msg:'Máquina não encontrada.' };
  
  const localAnterior = pack.item.localAtual;
  if(localAnterior.toLowerCase() === novoLocal.toLowerCase()){
    return { ok:false, msg:'O novo local é igual ao atual.' };
  }
  
  const ctx = ti__ctx_();
  const responsavel = ctx.usuario || Session.getActiveUser().getEmail();
  
  const { sh, shHist, headers } = maq__ensureSheets_();
  const rowRef = pack.item.rowRef;
  const iLocal = maq__idx_(headers,'Local Atual');
  const iUp = maq__idx_(headers,'Atualizado Em');
  
  const lock = LockService.getScriptLock();
  
  try{
    lock.waitLock(10000);
    
    // Atualiza local na planilha principal
    if(iLocal >= 0){
      sh.getRange(rowRef, iLocal+1).setValue(novoLocal);
    }
    if(iUp >= 0){
      sh.getRange(rowRef, iUp+1).setValue(new Date());
    }
    
    // Registra no histórico
    shHist.appendRow([
      new Date(),
      id,
      localAnterior,
      novoLocal,
      responsavel,
      motivo
    ]);
    
    lock.releaseLock();
    return { ok:true, msg:'Máquina movida com sucesso!' };
    
  }catch(e){
    if(lock && lock.hasLock()) lock.releaseLock();
    return { ok:false, msg:'Erro ao mover: ' + e.message };
  }
}

/** ESTATÍSTICAS DE MÁQUINAS POR LOCAL */
function maq_estatisticasPorLocal(){
  try{
    maq__ensureSheets_();
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const shMaq = ss.getSheetByName('Maquinas');
    if (!shMaq) return { ok:true, locais: [] };

    const lastRow = shMaq.getLastRow();
    if (lastRow < 2) return { ok:true, locais: [] };

    const vals = shMaq.getRange(2, 1, lastRow-1, shMaq.getLastColumn()).getValues();
    
    // Usar headers dinâmicos em vez de índices hardcoded
    const maqHeaders = shMaq.getRange(1,1,1,shMaq.getLastColumn()).getValues()[0] || [];
    const iLocal = maq__idx_(maqHeaders, 'Local Atual');
    const iTipo = maq__idx_(maqHeaders, 'Tipo Máquina');
    if (iLocal < 0 || iTipo < 0) return { ok:true, locais: [] };

    const stats = {}; // { local: { tipo: count } }

    vals.forEach(row => {
      const local = String(row[iLocal] || '—').trim();
      const tipo = String(row[iTipo] || 'Outro').trim();
      
      if (!stats[local]) stats[local] = {};
      stats[local][tipo] = (stats[local][tipo] || 0) + 1;
    });

    // Converter para array ordenado
    const locais = Object.keys(stats).sort((a,b) => a.localeCompare(b, 'pt-BR', {sensitivity:'base'}));
    const resultado = locais.map(local => {
      const tipos = stats[local];
      const total = Object.values(tipos).reduce((a,b) => a+b, 0);
      const detalhe = Object.keys(tipos)
        .sort((a,b) => tipos[b] - tipos[a])
        .map(tipo => ({ tipo, quantidade: tipos[tipo] }));
      
      return { local, total, detalhes: detalhe };
    });

    return { ok:true, locais: resultado };
  }catch(e){
    return { ok:false, msg:'Erro ao obter estatísticas: ' + e.message };
  }
}

function maq_historico(input){
  const id = (typeof input === 'object' && input) ? (input.id || input.ID || '') : String(input||'');
  if(!id) return { ok:false, msg:'ID vazio.' };
  
  const { shHist } = maq__ensureSheets_();
  const lastRow = shHist.getLastRow();
  
  if(lastRow < 2) return { ok:true, historico:[] };
  
  const values = shHist.getRange(2,1,lastRow-1,shHist.getLastColumn()).getValues();
  const historico = [];
  
  for(let r=0; r<values.length; r++){
    const row = values[r];
    const maqId = String(row[1]||'').trim();
    
    if(maqId !== id) continue;
    
    historico.push({
      carimbo: row[0] instanceof Date ? Utilities.formatDate(row[0], TZ, "yyyy-MM-dd'T'HH:mm:ss") : (row[0]||''),
      localAnterior: String(row[2]||'').trim(),
      localNovo: String(row[3]||'').trim(),
      responsavel: String(row[4]||'').trim(),
      motivo: String(row[5]||'').trim()
    });
  }
  
  // Ordena por data desc
  historico.sort((a,b)=>{
    const ta = parseAnyDate(a.carimbo) || new Date(0);
    const tb = parseAnyDate(b.carimbo) || new Date(0);
    return tb.getTime() - ta.getTime();
  });
  
  return { ok:true, historico };
}

// =====================================================
// RMA - CORE BACKEND
// =====================================================

const RMA_SPREADSHEET_ID = PLANILHA_ID;
const SHEET_SUBMISSOES = 'SUBMISSOES';
const SHEET_BASE_FLAT = 'BASE_FLAT';
const RMA_EDIT_WINDOW_HOURS = 24;

const RMA_UNIDADES = [
  { id:'CT1', nome:'Conselho Tutelar 1', tipo:'CONSELHO_TUTELAR' },
  { id:'CT2', nome:'Conselho Tutelar 2', tipo:'CONSELHO_TUTELAR' },
  { id:'CT3', nome:'Conselho Tutelar 3', tipo:'CONSELHO_TUTELAR' },
  { id:'CT4', nome:'Conselho Tutelar 4', tipo:'CONSELHO_TUTELAR' },

  { id:'CRAS_ANGELA_MARIA', nome:'CRAS Angela Maria', tipo:'CRAS' },
  { id:'CRAS_DR_FRANKLIN', nome:'CRAS Dr. Franklin', tipo:'CRAS' },
  { id:'CRAS_JOAO_ALVES', nome:'CRAS Joao Alves', tipo:'CRAS' },
  { id:'CRAS_MARIA_JOSE', nome:'CRAS Maria Jose', tipo:'CRAS' },
  { id:'CRAS_PROF_MARIA_LUIZA', nome:'CRAS Prof Maria Luiza', tipo:'CRAS' },
  { id:'CRAS_ZILDA_ARNS', nome:'CRAS Zilda Arns', tipo:'CRAS' },

  { id:'CREAS_LEONEL_BRIZOLA', nome:'CREAS Leonel Brizola', tipo:'CREAS' },
  { id:'CREAS_MARCOS_FREIRE_I', nome:'CREAS Marcos Freire I', tipo:'CREAS' },

  { id:'CRAM', nome:'CRAM', tipo:'CRAM' },
  { id:'BEM', nome:'BEM', tipo:'BEM' },
  { id:'CASA_CONSELHOS', nome:'Casa dos Conselhos', tipo:'CASA_CONSELHOS' },
  { id:'CENTRAL_CAD', nome:'Central CadUnico', tipo:'CADUNICO_CENTRAL' },
  { id:'RESTAURANTE_POPULAR', nome:'Restaurante Popular', tipo:'RESTAURANTE_POPULAR' },

  { id:'UA_IRMA_VALMIRA', nome:'UA Irma Valmira', tipo:'ACOLHIMENTO' },
  { id:'UA_PROF_ROSINEIDE', nome:'UA Prof Rosineide', tipo:'ACOLHIMENTO' }
];

const RMA_FORMS_BY_TIPO = {
  CRAS: ['RMA_CRAS'],
  CREAS: ['RMA_CREAS'],
  CADUNICO_CENTRAL: ['CADUNICO_CENTRAL'],
  CONSELHO_TUTELAR: ['CONSELHO_TUTELAR'],
  CRAM: ['CRAM_MENSAL'],
  CASA_CONSELHOS: ['CASA_CONSELHOS_MENSAL'],
  BEM: ['BEM_GERAL'],
  ACOLHIMENTO: ['ACOLHIMENTO_MENSAL'],
  RESTAURANTE_POPULAR: ['RESTAURANTE_POPULAR_MENSAL']
};

const RMA_SUB_HEADERS = [
  'Carimbo',
  'Data_Preenchimento',
  'Competencia',
  'Ano',
  'Mes',
  'Unidade_ID',
  'Unidade',
  'Tipo_Unidade',
  'Form_Key',
  'Versao',
  'Dados_JSON',
  'Preenchido_por',
  'Email_Usuario',
  'Submission_ID',
  'Editavel_ate',
  'Atualizado_em',
  'Client_Request_ID',
  'Payload_Hash'
];

const RMA_BASE_HEADERS = [
  'Carimbo',
  'Competencia',
  'Ano',
  'Mes',
  'Unidade_ID',
  'Unidade',
  'Tipo_Unidade',
  'Form_Key',
  'Versao',
  'Submission_ID',
  'Preenchido_por',
  'Email_Usuario'
];

function ss_(){
  return SpreadsheetApp.openById(RMA_SPREADSHEET_ID);
}

function normStr_(v){
  return String(v || '').trim();
}

function normCompetencia_(v){
  const s = normStr_(v);
  if (!s) return '';
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{2})\/(\d{4})$/);
  if (m) return `${m[2]}-${m[1]}`;
  return s;
}

function asISODate_(v){
  if (!v) return '';
  const d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function now_(){
  return new Date();
}

function getColsMap_(sh){
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const map = {};
  header.forEach((h, i)=>{
    const key = String(h || '').trim();
    if (key) map[key] = i + 1;
  });
  return map;
}

function ensureSubmissoesSheet_(){
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_SUBMISSOES);
  if (!sh){
    sh = ss.insertSheet(SHEET_SUBMISSOES);
    sh.getRange(1,1,1,RMA_SUB_HEADERS.length).setValues([RMA_SUB_HEADERS]);
  } else if (sh.getLastRow() === 0){
    sh.getRange(1,1,1,RMA_SUB_HEADERS.length).setValues([RMA_SUB_HEADERS]);
  }
  return sh;
}

function ensureBaseFlatSheet_(){
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_BASE_FLAT);
  if (!sh){
    sh = ss.insertSheet(SHEET_BASE_FLAT);
    sh.getRange(1,1,1,RMA_BASE_HEADERS.length).setValues([RMA_BASE_HEADERS]);
  } else if (sh.getLastRow() === 0){
    sh.getRange(1,1,1,RMA_BASE_HEADERS.length).setValues([RMA_BASE_HEADERS]);
  }
  return sh;
}

function hashPayload_(payload){
  const s = JSON.stringify(payload || {});
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function flattenDados_(dados){
  const out = {};
  Object.keys(dados || {}).forEach(k=>{
    const v = dados[k];
    if (v === null || v === undefined) {
      out[k] = '';
      return;
    }
    if (Array.isArray(v) || typeof v === 'object'){
      out[k] = JSON.stringify(v);
      return;
    }
    out[k] = v;
  });
  return out;
}

function ensureBaseFlatColumns_(sh, keys){
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const existing = new Set(header.map(h=>String(h || '').trim()));
  const toAdd = keys.filter(k=>!existing.has(k));
  if (!toAdd.length) return;
  const startCol = header.length + 1;
  sh.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
}

function upsertBaseFlat_(payload, submissionId, carimbo){
  const sh = ensureBaseFlatSheet_();
  const cols = getColsMap_(sh);
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

  const dadosFlat = flattenDados_(payload.dados || {});
  ensureBaseFlatColumns_(sh, Object.keys(dadosFlat));

  const updatedHeader = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const updatedCols = getColsMap_(sh);

  const rowValues = new Array(updatedHeader.length).fill('');
  const setVal = (name, value)=>{
    const idx = updatedCols[name];
    if (idx) rowValues[idx-1] = value;
  };

  setVal('Carimbo', carimbo);
  setVal('Competencia', normCompetencia_(payload.competencia));
  setVal('Ano', payload.ano || '');
  setVal('Mes', payload.mes || '');
  setVal('Unidade_ID', payload.unidadeId || '');
  setVal('Unidade', payload.unidadeNome || '');
  setVal('Tipo_Unidade', payload.unidadeTipo || '');
  setVal('Form_Key', payload.formKey || '');
  setVal('Versao', payload.versao || 1);
  setVal('Submission_ID', submissionId);
  setVal('Preenchido_por', payload.preenchidoPor || '');
  setVal('Email_Usuario', payload.emailUsuario || '');

  Object.keys(dadosFlat).forEach(k=> setVal(k, dadosFlat[k]));

  const lastRow = sh.getLastRow();
  let targetRow = -1;
  if (lastRow >= 2){
    const colSubId = updatedCols['Submission_ID'];
    if (colSubId){
      const vals = sh.getRange(2, colSubId, lastRow-1, 1).getValues();
      for (let i=0;i<vals.length;i++){
        if (String(vals[i][0] || '') === submissionId){
          targetRow = i + 2;
          break;
        }
      }
    }
  }

  if (targetRow > 0){
    sh.getRange(targetRow,1,1,rowValues.length).setValues([rowValues]);
  } else {
    sh.appendRow(rowValues);
  }
}

function api_getBootstrap(){
  return {
    unidades: RMA_UNIDADES,
    formsByTipo: RMA_FORMS_BY_TIPO,
    serverTime: now_().toISOString()
  };
}

function api_saveSubmission(payload){
  payload = payload || {};
  ensureSubmissoesSheet_();
  ensureBaseFlatSheet_();

  const comp = normCompetencia_(payload.competencia);
  if (!comp) return { ok:false, error:'Competencia obrigatoria.' };
  if (!payload.unidadeId || !payload.formKey) return { ok:false, error:'Unidade e formulario obrigatorios.' };

  const sh = ss_().getSheetByName(SHEET_SUBMISSOES);
  const cols = getColsMap_(sh);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  const clientRequestId = payload.clientRequestId || '';
  const payloadHash = payload.payloadHash || hashPayload_(payload.dados || {});

  if (lastRow >= 2){
    const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
    for (let i=0;i<values.length;i++){
      const row = values[i];
      const existingReq = String(row[cols['Client_Request_ID']-1] || '');
      const existingHash = String(row[cols['Payload_Hash']-1] || '');
      if ((clientRequestId && existingReq === clientRequestId) || (payloadHash && existingHash === payloadHash)){
        return { ok:true, submissionId: row[cols['Submission_ID']-1], message:'Registro ja salvo.' };
      }
    }
  }

  const carimbo = now_();
  const submissionId = Utilities.getUuid();
  const editavelAte = new Date(carimbo.getTime() + RMA_EDIT_WINDOW_HOURS*60*60*1000);
  const emailUsuario = Session.getActiveUser().getEmail();

  const row = new Array(RMA_SUB_HEADERS.length).fill('');
  const setVal = (name, value)=>{
    const idx = RMA_SUB_HEADERS.indexOf(name);
    if (idx >= 0) row[idx] = value;
  };

  setVal('Carimbo', carimbo);
  setVal('Data_Preenchimento', payload.dataPreenchimento || '');
  setVal('Competencia', comp);
  setVal('Ano', payload.ano || '');
  setVal('Mes', payload.mes || '');
  setVal('Unidade_ID', payload.unidadeId || '');
  setVal('Unidade', payload.unidadeNome || '');
  setVal('Tipo_Unidade', payload.unidadeTipo || '');
  setVal('Form_Key', payload.formKey || '');
  setVal('Versao', payload.versao || 1);
  setVal('Dados_JSON', JSON.stringify(payload.dados || {}));
  setVal('Preenchido_por', payload.preenchidoPor || '');
  setVal('Email_Usuario', emailUsuario || '');
  setVal('Submission_ID', submissionId);
  setVal('Editavel_ate', editavelAte);
  setVal('Atualizado_em', carimbo);
  setVal('Client_Request_ID', clientRequestId);
  setVal('Payload_Hash', payloadHash);

  sh.appendRow(row);

  payload.emailUsuario = emailUsuario || '';
  upsertBaseFlat_(payload, submissionId, carimbo);

  const pastaInfo = criarPastaRMA_(payload);
  const pdfInfo = gerarPDFRMA_(payload, pastaInfo);

  return { ok:true, submissionId, message:'RMA salvo com sucesso!', pastaInfo, pdfInfo };
}

function api_updateSubmission(payload){
  payload = payload || {};
  ensureSubmissoesSheet_();
  ensureBaseFlatSheet_();

  const submissionId = payload.submissionId || '';
  if (!submissionId) return { ok:false, error:'Submission ID obrigatorio.' };

  const sh = ss_().getSheetByName(SHEET_SUBMISSOES);
  const cols = getColsMap_(sh);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return { ok:false, error:'Nada para atualizar.' };

  const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  let rowIndex = -1;
  let rowData = null;
  for (let i=0;i<values.length;i++){
    if (String(values[i][cols['Submission_ID']-1] || '') === submissionId){
      rowIndex = i + 2;
      rowData = values[i];
      break;
    }
  }
  if (rowIndex < 0) return { ok:false, error:'Registro nao encontrado.' };

  const editavelAte = rowData[cols['Editavel_ate']-1];
  const editavelAteDt = (editavelAte instanceof Date) ? editavelAte : new Date(editavelAte);
  if (editavelAteDt && now_().getTime() > editavelAteDt.getTime()){
    return { ok:false, error:'Prazo de edicao expirado.' };
  }

  const carimbo = now_();
  const emailUsuario = Session.getActiveUser().getEmail();
  const payloadHash = payload.payloadHash || hashPayload_(payload.dados || {});

  const setCell = (name, value)=>{
    const idx = cols[name];
    if (idx) sh.getRange(rowIndex, idx).setValue(value);
  };

  setCell('Data_Preenchimento', payload.dataPreenchimento || '');
  setCell('Competencia', normCompetencia_(payload.competencia));
  setCell('Ano', payload.ano || '');
  setCell('Mes', payload.mes || '');
  setCell('Unidade_ID', payload.unidadeId || '');
  setCell('Unidade', payload.unidadeNome || '');
  setCell('Tipo_Unidade', payload.unidadeTipo || '');
  setCell('Form_Key', payload.formKey || '');
  setCell('Versao', payload.versao || 1);
  setCell('Dados_JSON', JSON.stringify(payload.dados || {}));
  setCell('Preenchido_por', payload.preenchidoPor || '');
  setCell('Email_Usuario', emailUsuario || '');
  setCell('Atualizado_em', carimbo);
  setCell('Payload_Hash', payloadHash);

  payload.emailUsuario = emailUsuario || '';
  upsertBaseFlat_(payload, submissionId, carimbo);

  const pastaInfo = criarPastaRMA_(payload);
  const pdfInfo = gerarPDFRMA_(payload, pastaInfo);

  return { ok:true, submissionId, message:'RMA atualizado com sucesso!', pastaInfo, pdfInfo };
}

function api_getLatestSubmission(q){
  q = q || {};
  ensureSubmissoesSheet_();
  const sh = ss_().getSheetByName(SHEET_SUBMISSOES);
  const cols = getColsMap_(sh);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return { found:false };

  const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  const unidadeId = normStr_(q.unidadeId);
  const formKey = normStr_(q.formKey);

  for (let i=values.length-1; i>=0; i--){
    const row = values[i];
    if (unidadeId && normStr_(row[cols['Unidade_ID']-1]) !== unidadeId) continue;
    if (formKey && normStr_(row[cols['Form_Key']-1]) !== formKey) continue;

    const data = {};
    Object.keys(cols).forEach(k=>{ data[k] = row[cols[k]-1]; });
    const editavelAte = data['Editavel_ate'];
    const editavelDt = (editavelAte instanceof Date) ? editavelAte : new Date(editavelAte);
    const editavel = editavelDt && now_().getTime() <= editavelDt.getTime();
    return { found:true, data, editavel };
  }

  return { found:false };
}

function api_queryBaseFlat(q){
  q = q || {};
  ensureBaseFlatSheet_();
  const sh = ss_().getSheetByName(SHEET_BASE_FLAT);
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const cols = getColsMap_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { header, rows: [] };

  const values = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  const compFrom = normCompetencia_(q.competenciaFrom || '');
  const compTo = normCompetencia_(q.competenciaTo || '');
  const unidadeId = normStr_(q.unidadeId || '');
  const formKey = normStr_(q.formKey || '');
  const limit = Number(q.limit || 1000);

  const rows = [];
  values.forEach(r=>{
    const comp = normCompetencia_(r[cols['Competencia']-1]);
    if (compFrom && comp < compFrom) return;
    if (compTo && comp > compTo) return;
    if (unidadeId && unidadeId !== 'TODAS' && normStr_(r[cols['Unidade_ID']-1]) !== unidadeId) return;
    if (formKey && normStr_(r[cols['Form_Key']-1]) !== formKey) return;
    rows.push(r);
  });

  return { header, rows: rows.slice(0, limit) };
}

function api_getDashboardData(q){
  q = q || {};
  ensureBaseFlatSheet_();

  const sh = ss_().getSheetByName(SHEET_BASE_FLAT);
  const cols = getColsMap_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { series: [], byUnit: [], meta:{ totalRegistros:0, totalValor:0 } };

  const values = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  const compFrom = normCompetencia_(q.competenciaFrom || '');
  const compTo = normCompetencia_(q.competenciaTo || '');
  const unidadeId = normStr_(q.unidadeId || '');
  const formKey = normStr_(q.formKey || '');
  const indicadorKey = normStr_(q.indicadorKey || '');
  const indCol = cols[indicadorKey];

  if (!indCol) return { series: [], byUnit: [], meta:{ totalRegistros:0, totalValor:0 } };

  const byComp = {};
  const byUnit = {};
  let totalRegistros = 0;
  let totalValor = 0;

  values.forEach(r=>{
    const comp = normCompetencia_(r[cols['Competencia']-1]);
    if (compFrom && comp < compFrom) return;
    if (compTo && comp > compTo) return;
    if (unidadeId && unidadeId !== 'TODAS' && normStr_(r[cols['Unidade_ID']-1]) !== unidadeId) return;
    if (formKey && normStr_(r[cols['Form_Key']-1]) !== formKey) return;

    const raw = r[indCol-1];
    const val = parseNumber_(raw);
    if (!isNaN(val)){
      totalValor += val;
      byComp[comp] = (byComp[comp] || 0) + val;
      const unidadeNome = normStr_(r[cols['Unidade']-1]) || '—';
      byUnit[unidadeNome] = (byUnit[unidadeNome] || 0) + val;
    }
    totalRegistros += 1;
  });

  const series = Object.keys(byComp).sort().map(k=>({ competencia: k, valor: byComp[k] }));
  const byUnitArr = Object.keys(byUnit).sort((a,b)=> byUnit[b]-byUnit[a]).map(k=>({ unidade: k, valor: byUnit[k] }));

  return {
    series,
    byUnit: byUnitArr,
    meta: { totalRegistros, totalValor }
  };
}

function parseNumber_(v){
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/\./g,'').replace(',','.').replace(/[^0-9.-]/g,'');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

// =====================================================
// RMA - EXTENSAO (VIGILANCIA + PASTAS + PDF)
// =====================================================

const RMA_FORMS_LABEL = {
  RMA_CRAS: 'RMA CRAS',
  RMA_CREAS: 'RMA CREAS',
  CADUNICO_CENTRAL: 'CadUnico',
  CONSELHO_TUTELAR: 'Conselho Tutelar',
  CRAM_MENSAL: 'CRAM',
  CASA_CONSELHOS_MENSAL: 'Casa dos Conselhos',
  BEM_GERAL: 'BEM',
  ACOLHIMENTO_MENSAL: 'Acolhimento',
  RESTAURANTE_POPULAR_MENSAL: 'Restaurante Popular'
};

function api_queryVigilancia(q){
  if (typeof ensureSubmissoesSheet_ !== 'function' ||
      typeof ss_ !== 'function' ||
      typeof getColsMap_ !== 'function' ||
      typeof normCompetencia_ !== 'function' ||
      typeof normStr_ !== 'function' ||
      typeof asISODate_ !== 'function') {
    return { ok:false, error:'Backend RMA nao integrado no code.js.' };
  }

  ensureSubmissoesSheet_();

  q = q || {};
  const compFrom = q.competenciaFrom ? normCompetencia_(q.competenciaFrom) : '';
  const compTo = q.competenciaTo ? normCompetencia_(q.competenciaTo) : '';
  const unidadeId = q.unidadeId ? normStr_(q.unidadeId) : '';
  const formKey = q.formKey ? normStr_(q.formKey) : '';

  const sh = ss_().getSheetByName(SHEET_SUBMISSOES);
  const cols = getColsMap_(sh);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2) return { records: [], total: 0, stats: {} };

  const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  const records = [];

  values.forEach(r=>{
    const comp = normCompetencia_(r[cols['Competencia']-1] || r[cols['Competência']-1]);
    if (compFrom && comp < compFrom) return;
    if (compTo && comp > compTo) return;

    const uid = normStr_(r[cols['Unidade_ID']-1]);
    if (unidadeId && uid !== unidadeId) return;

    const fk = normStr_(r[cols['Form_Key']-1]);
    if (formKey && fk !== formKey) return;

    const carimbo = r[cols['Carimbo']-1];
    const carimboDt = (carimbo instanceof Date) ? carimbo : new Date(carimbo);
    const now = new Date();
    const diffMs = now.getTime() - carimboDt.getTime();
    const editavel = diffMs <= 24*60*60*1000;

    records.push({
      submissionId: r[cols['Submission_ID']-1] || '',
      competencia: comp,
      dataPreenchimento: r[cols['Data_Preenchimento']-1] ? asISODate_(r[cols['Data_Preenchimento']-1]) : '',
      unidadeId: uid,
      unidadeNome: r[cols['Unidade']-1] || '',
      unidadeTipo: r[cols['Tipo_Unidade']-1] || '',
      formKey: fk,
      formLabel: RMA_FORMS_LABEL[fk] || fk,
      preenchidoPor: r[cols['Preenchido_por']-1] || '',
      email: r[cols['Email_Usuario']-1] || '',
      carimbo: carimbo,
      carimboISO: carimboDt.toISOString(),
      editavel,
      editavelAte: r[cols['Editavel_ate']-1] ? asISODate_(r[cols['Editavel_ate']-1]) : ''
    });
  });

  records.sort((a,b)=> new Date(b.carimboISO) - new Date(a.carimboISO));

  return {
    records,
    total: records.length,
    stats: {
      por_unidade: groupBy_(records, 'unidadeNome'),
      por_tipo: groupBy_(records, 'formKey'),
      editaveis: records.filter(r=>r.editavel).length
    }
  };
}

function groupBy_(arr, key){
  const map = {};
  (arr || []).forEach(item=>{
    const k = item[key];
    map[k] = (map[k] || 0) + 1;
  });
  return map;
}

function criarPastaRMA_(payload){
  try {
    if (!RMA_OUTPUT_FOLDER_ID) {
      console.warn('RMA_OUTPUT_FOLDER_ID nao configurado.');
      return null;
    }

    const comp = (typeof normCompetencia_ === 'function')
      ? normCompetencia_(payload.competencia)
      : String(payload.competencia || '').trim();
    const unidade = (typeof normStr_ === 'function')
      ? normStr_(payload.unidadeNome)
      : String(payload.unidadeNome || '').trim();

    const rootFolder = DriveApp.getFolderById(RMA_OUTPUT_FOLDER_ID);

    let rmaFolder = null;
    const rmaFolders = rootFolder.getFoldersByName('RMA');
    if (rmaFolders.hasNext()){
      rmaFolder = rmaFolders.next();
    } else {
      rmaFolder = rootFolder.createFolder('RMA');
    }

    let mesFolder = null;
    const mesFolders = rmaFolder.getFoldersByName(comp);
    if (mesFolders.hasNext()){
      mesFolder = mesFolders.next();
    } else {
      mesFolder = rmaFolder.createFolder(comp);
    }

    let unidadeFolder = null;
    const unidadeFolders = mesFolder.getFoldersByName(unidade);
    if (unidadeFolders.hasNext()){
      unidadeFolder = unidadeFolders.next();
    } else {
      unidadeFolder = mesFolder.createFolder(unidade);
    }

    return {
      rootId: rootFolder.getId(),
      rmaId: rmaFolder.getId(),
      mesId: mesFolder.getId(),
      unidadeId: unidadeFolder.getId(),
      path: `RMA/${comp}/${unidade}/`
    };
  } catch(e){
    console.error('Erro ao criar pasta RMA:', e);
    return null;
  }
}

function gerarPDFRMA_(payload, pastaInfo){
  try {
    if (!pastaInfo || !pastaInfo.unidadeId) {
      console.warn('Info de pasta invalida.');
      return null;
    }

    const comp = (typeof normCompetencia_ === 'function')
      ? normCompetencia_(payload.competencia)
      : String(payload.competencia || '').trim();
    const unidade = payload.unidadeNome;
    const formKey = payload.formKey;
    const folder = DriveApp.getFolderById(pastaInfo.unidadeId);

    const fileName = `${unidade}_${formKey}_${comp}.pdf`;

    const docName = fileName.replace(/\.pdf$/i, '');
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    body.appendParagraph('RMA: ' + (RMA_FORMS_LABEL[payload.formKey] || payload.formKey));
    body.appendParagraph('Unidade: ' + (payload.unidadeNome || ''));
    body.appendParagraph('Competencia: ' + (payload.competencia || ''));
    body.appendParagraph('Data de Preenchimento: ' + (payload.dataPreenchimento || ''));
    body.appendParagraph('Preenchido por: ' + (payload.preenchidoPor || ''));
    body.appendParagraph('');
    body.appendParagraph('DADOS PREENCHIDOS:');
    body.appendParagraph(JSON.stringify(payload.dados || {}, null, 2));
    doc.saveAndClose();

    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs(MimeType.PDF).setName(fileName);
    const pdfFile = folder.createFile(pdfBlob);
    docFile.setTrashed(true);

    console.log('PDF gerado:', fileName);

    return {
      fileName,
      fileId: pdfFile.getId(),
      url: pdfFile.getUrl(),
      folderId: pastaInfo.unidadeId,
      path: `${pastaInfo.path}${fileName}`
    };

  } catch(e){
    console.error('Erro ao gerar PDF:', e);
    return null;
  }
}

function gerarPDFContent_(payload){
  const lines = [
    '='.repeat(80),
    `RMA: ${RMA_FORMS_LABEL[payload.formKey] || payload.formKey}`,
    '='.repeat(80),
    '',
    `Unidade: ${payload.unidadeNome}`,
    `Competencia: ${payload.competencia}`,
    `Data de Preenchimento: ${payload.dataPreenchimento}`,
    `Preenchido por: ${payload.preenchidoPor}`,
    '',
    'DADOS PREENCHIDOS:',
    JSON.stringify(payload.dados, null, 2),
    '',
    '='.repeat(80)
  ];

  return lines.join('\n');
}

function api_saveSubmission_WithPDF(payload) {
  if (typeof api_saveSubmission !== 'function') {
    return { ok:false, error:'api_saveSubmission nao encontrada no code.js.' };
  }

  const result = api_saveSubmission(payload);

  if (result.ok && result.submissionId){
    const pastaInfo = criarPastaRMA_(payload);
    const pdfInfo = gerarPDFRMA_(payload, pastaInfo);

    result.pastaInfo = pastaInfo;
    result.pdfInfo = pdfInfo;
    result.message = 'RMA salvo com sucesso! Pasta e PDF gerados.';
  }

  return result;
}

function api_updateSubmission_WithPDF(payload) {
  if (typeof api_updateSubmission !== 'function') {
    return { ok:false, error:'api_updateSubmission nao encontrada no code.js.' };
  }

  const result = api_updateSubmission(payload);

  if (result.ok && result.submissionId){
    const pastaInfo = criarPastaRMA_(payload);
    const pdfInfo = gerarPDFRMA_(payload, pastaInfo);

    result.pastaInfo = pastaInfo;
    result.pdfInfo = pdfInfo;
    result.message = 'RMA atualizado! PDF regenerado.';
  }

  return result;
}

function listarTemplatesRMA(){
  try {
    if (!RMA_TEMPLATES_FOLDER_ID) {
      return { erro: 'RMA_TEMPLATES_FOLDER_ID nao configurado.' };
    }

    const folder = DriveApp.getFolderById(RMA_TEMPLATES_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.MICROSOFT_WORD);
    const templates = [];

    while (files.hasNext()){
      const f = files.next();
      templates.push({
        id: f.getId(),
        nome: f.getName(),
        url: f.getUrl()
      });
    }

    return { templates, total: templates.length };
  } catch(e){
    return { erro: String(e) };
  }
}

function limparPastasRMA_DEBUG(){
  try {
    if (!RMA_OUTPUT_FOLDER_ID) return 'Pasta nao configurada';

    const folder = DriveApp.getFolderById(RMA_OUTPUT_FOLDER_ID);
    const subfolders = folder.getFolders();

    let count = 0;
    while (subfolders.hasNext()){
      const sub = subfolders.next();
      if (sub.getName() === 'RMA'){
        folder.removeFolder(sub);
        count++;
      }
    }

    return `${count} pastas removidas (DEBUG)`;
  } catch(e){
    return 'Erro: ' + e;
  }
}
