(function() {
'use strict';

/* ── Helpers ── */
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function shorten(u){try{return new URL(u).hostname}catch{return u}}
function initials(n){return(n||'?').split(/\s+/).map(function(w){return w[0]}).join('').toUpperCase().slice(0,2)||'?'}
function getVal(id){var el=document.getElementById(id);return el?el.value.trim():''}
function setVal(id,v){var el=document.getElementById(id);if(el)el.value=v}

var _tt;
function toast(msg){
  var el=document.getElementById('toast');
  if(!el){return;}
  el.textContent=msg;
  el.classList.add('show');
  clearTimeout(_tt);
  _tt=setTimeout(function(){el.classList.remove('show')},3600);
}

function setSt(id,type,text){
  var el=document.getElementById(id);
  if(!el){return;}
  var cls={ok:'s-ok',error:'s-error',loading:'s-loading',idle:'s-idle'};
  var dot={ok:'d-ok',error:'d-error',loading:'d-loading',idle:'d-idle'};
  el.className='sbox '+(cls[type]||'s-idle');
  el.innerHTML='<div class="sdot '+(dot[type]||'d-idle')+'"></div><span>'+text+'</span>';
}

var _progressHideTimer;
function showProgress(label,text,pct){
  var panel=document.getElementById('progressPanel');
  if(!panel){return;}
  document.getElementById('progressLabel').textContent=label||'Working...';
  document.getElementById('progressText').textContent=text||'';
  var safePct=Math.max(0,Math.min(100,Math.round(pct||0)));
  document.getElementById('progressPct').textContent=safePct+'%';
  document.getElementById('progressFill').style.width=safePct+'%';
  clearTimeout(_progressHideTimer);
  panel.style.display='block';
}

function hideProgress(delay){
  clearTimeout(_progressHideTimer);
  var run=function(){
    var panel=document.getElementById('progressPanel');
    if(panel)panel.style.display='none';
  };
  if(delay)_progressHideTimer=setTimeout(run,delay);
  else run();
}

function setSetupCardCollapsed(collapsed){
  var body=document.getElementById('setupCardBody');
  var btn=document.getElementById('btnToggleSetup');
  if(!body||!btn){return;}
  body.style.display=collapsed?'none':'block';
  btn.textContent=collapsed?'Expand':'Collapse';
  btn.setAttribute('aria-expanded',collapsed?'false':'true');
}

function focusModuleFiltersCard(){
  var card=document.getElementById('moduleFiltersCard');
  var modSel=document.getElementById('modSel');
  if(card&&card.scrollIntoView){
    card.scrollIntoView({behavior:'smooth',block:'start'});
  }
  if(modSel&&modSel.focus){
    setTimeout(function(){modSel.focus();},150);
  }
}

var _moduleProgressHideTimer;
function showModuleProgress(label,text,pct){
  var panel=document.getElementById('moduleProgressPanel');
  if(!panel){return;}
  document.getElementById('moduleProgressLabel').textContent=label||'Working...';
  document.getElementById('moduleProgressText').textContent=text||'';
  var safePct=Math.max(0,Math.min(100,Math.round(pct||0)));
  document.getElementById('moduleProgressPct').textContent=safePct+'%';
  document.getElementById('moduleProgressFill').style.width=safePct+'%';
  clearTimeout(_moduleProgressHideTimer);
  panel.style.display='block';
}

function hideModuleProgress(delay){
  clearTimeout(_moduleProgressHideTimer);
  var run=function(){
    var panel=document.getElementById('moduleProgressPanel');
    if(panel)panel.style.display='none';
  };
  if(delay)_moduleProgressHideTimer=setTimeout(run,delay);
  else run();
}

function generateId(){
  if(typeof crypto!=='undefined'&&crypto.randomUUID)return crypto.randomUUID();
  return 'id-'+Date.now()+'-'+Math.random().toString(36).slice(2);
}

var ROOT_SCOPE = typeof window !== 'undefined' ? window : globalThis;
ROOT_SCOPE.esc = esc;
ROOT_SCOPE.toast = toast;
ROOT_SCOPE.setSt = setSt;
ROOT_SCOPE.showProgress = showProgress;
ROOT_SCOPE.hideProgress = hideProgress;
ROOT_SCOPE.showModuleProgress = showModuleProgress;
ROOT_SCOPE.hideModuleProgress = hideModuleProgress;

/* ── Mode detection ── */
var IS_EXT = typeof chrome !== 'undefined' && !!chrome.runtime && !!chrome.runtime.id;
var IS_FILE = window.location.protocol === 'file:';

(function(){
  var el = document.getElementById('modeNote');
  if (!el) return;
  if(IS_EXT){
    el.innerHTML = '<span>&#x1F50C; <strong>Extension mode</strong> &mdash; data is fetched directly through your open D365 tabs using your existing browser session. No bearer token needed &mdash; just make sure each environment is open and you are logged in.</span>';
    el.className = 'note note-blue';
  } else if (IS_FILE) {
    el.innerHTML = '<span>&#x26A0;&#xFE0F; <strong>File mode</strong> &mdash; direct D365 calls are blocked by browser security. Use Launch-Tool.bat and open the localhost version instead.</span>';
    el.className = 'note';
  } else {
    el.innerHTML = '<span>&#x2705; <strong>HTTP mode</strong> &mdash; served over localhost. Make sure you are logged in to each D365 environment in this browser.</span>';
    el.className = 'note note-blue';
  }
  el.style.display = 'flex';
})();

function normUrl(url){
  return String(url || '').trim().replace(/\/+$/, '');
}

function isHttps(url){
  return /^https:\/\//i.test(String(url || '').trim());
}

function getSlotTitle(slot) {
  return slot === 'A' ? 'Source' : 'Target';
}

/* ── Storage helpers — chrome.storage.local in extension, localStorage elsewhere ── */
function _store(){
  return (IS_EXT && typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local)
    ? chrome.storage.local
    : null;
}

/* ── Profile helpers ── */
function getProfiles(){
  try{return JSON.parse(localStorage.getItem('d365_profiles')||'[]')}
  catch(e){return []}
}

function putProfiles(list){
  var json = JSON.stringify(list);
  localStorage.setItem('d365_profiles', json);
  var s = _store();
  if(!s) return Promise.resolve();
  return new Promise(function(resolve, reject){
    try{
      s.set({'d365_profiles': json}, function(){
        var err = chrome.runtime && chrome.runtime.lastError;
        if(err) reject(new Error(err.message || 'Failed to save profiles.'));
        else resolve();
      });
    }catch(e){
      reject(e);
    }
  });
}

// Load profiles from chrome.storage into localStorage then re-render.
// Returns a Promise that resolves when done.
function syncProfilesFromStorage(){
  var s = _store();
  if(!s) return Promise.resolve();
  return new Promise(function(resolve){
    s.get(['d365_profiles','d365_pick'], function(result){
      if(result.d365_profiles){
        localStorage.setItem('d365_profiles', result.d365_profiles);
      }
      if(result.d365_pick){
        localStorage.setItem('d365_pick', result.d365_pick);
      }
      resolve();
    });
  });
}

function renderProfileList(){
  var el = document.getElementById('pfList');
  if (!el) return;
  var list = getProfiles();
  if (!list.length) {
    el.innerHTML = '<div class="pf-empty">No profiles saved yet. Add one above.</div>';
    return;
  }
  el.innerHTML = list.map(function(p) {
    return '' +
      '<div class="pf-item">' +
        '<div class="pf-meta">' +
          '<div class="pf-name">' + esc(p.name) + '</div>' +
          '<div class="pf-url">' + esc(p.url) + '</div>' +
        '</div>' +
        '<div class="pf-actions">' +
          '<button class="pf-edit-btn" data-id="' + esc(p.id) + '">Edit</button>' +
          '<button class="pf-del-btn" data-id="' + esc(p.id) + '">Delete</button>' +
        '</div>' +
      '</div>';
  }).join('');
}

function refreshPickers(){
  var list = getProfiles();
  ['A','B'].forEach(function(slot){
    var sel = document.getElementById('picker' + slot);
    if (!sel) return;
    var current = slot === 'A' ? _selA : _selB;
    var html = '<option value="">— Select a saved profile —</option>' +
      list.map(function(p){
        return '<option value="' + esc(p.id) + '">' + esc(p.name + ' — ' + shorten(p.url)) + '</option>';
      }).join('');
    sel.innerHTML = html;
    if (current) sel.value = current;
  });
}

function refreshLegalEntityFilters(){
  return Promise.resolve();
}

async function addOrUpdateProfile(){
  var name=getVal('pfName').trim();
  var url=normUrl(getVal('pfUrl'));
  var eid=getVal('pfEditId');
  if(!name||!url){toast('⚠️ Enter both profile name and URL.');return}
  if(!isHttps(url)){toast('⚠️ URL must start with https://');return}
  var list=getProfiles();
  if(eid){
    var it=list.find(function(p){return p.id===eid});
    if(it){it.name=name;it.url=url}
    setVal('pfEditId','');
    document.getElementById('btnAddProfile').textContent='➕ Add Profile';
    toast('✅ Profile "'+name+'" updated.');
  } else {
    var ex=list.find(function(p){return p.url===url});
    if(ex){ex.name=name;toast('💾 Updated existing profile for this URL.');}
    else{list.push({id:generateId(),name:name,url:url});toast('✅ Profile "'+name+'" saved.');}
  }
  await putProfiles(list);
  setVal('pfName','');
  setVal('pfUrl','');
  var settings = document.getElementById('profileSettings');
  if (settings) settings.open = false;
  renderProfileList();
  refreshPickers();
}

function openProfileSettings(){
  var settings = document.getElementById('profileSettings');
  if (settings) settings.open = true;
}

function editProfile(id){
  var p=getProfiles().find(function(x){return x.id===id});
  if(!p)return;
  openProfileSettings();
  setVal('pfName',p.name);
  setVal('pfUrl',p.url);
  setVal('pfEditId',p.id);
  document.getElementById('btnAddProfile').textContent='💾 Update Profile';
  document.getElementById('pfName').focus();
}

async function deleteProfile(id){
  await putProfiles(getProfiles().filter(function(p){return p.id!==id}));
  if(_selA===id)clearSlot('A');
  if(_selB===id)clearSlot('B');
  renderProfileList();
  refreshPickers();
  toast('🗑️ Profile deleted.');
}

/* ── Slot selection ── */
var _selA=null,_selB=null;

function loadSlot(slot){
  var id=getVal('picker'+slot);
  if(!id){clearSlot(slot);return}
  var p=getProfiles().find(function(x){return x.id===id});
  if(!p)return;
  if(slot==='A')_selA=id;else _selB=id;
  setSt('st'+slot,'idle',(slot==='A'?'Source':'Target')+' — not connected');
  persist();
  toast('✅ "'+p.name+'" selected as '+getSlotTitle(slot));
}

function clearSlot(slot){
  if(slot==='A')_selA=null;else _selB=null;
  setVal('picker'+slot,'');
  setSt('st'+slot,'idle',(slot==='A'?'Source':'Target')+' — not connected');
  persist();
}

/* ── Get active URL/label for a slot ── */
function getEnvUrl(slot){
  var id=slot==='A'?_selA:_selB;
  if(!id)return'';
  var p=getProfiles().find(function(x){return x.id===id});
  return p?p.url:'';
}

function getEnvLabel(slot){
  var id=slot==='A'?_selA:_selB;
  if(!id)return getSlotTitle(slot);
  var p=getProfiles().find(function(x){return x.id===id});
  return p?p.name:getSlotTitle(slot);
}

function getCompany(){
  var companyEl = document.getElementById('company');
  if(!companyEl) return '';
  var value = companyEl.value || '';
  if(value === '__custom__') {
    var customEl = document.getElementById('companyCustom');
    return customEl ? customEl.value.trim().toUpperCase() : '';
  }
  return value;
}

function onCompanySelectChange(){
  var companyEl = document.getElementById('company');
  var customEl = document.getElementById('companyCustom');
  if(!companyEl || !customEl) return;
  var isCustom = companyEl.value === '__custom__';
  customEl.style.display = isCustom ? 'block' : 'none';
  if(isCustom) customEl.focus();
  else customEl.value = '';
}

/* ── Persist / restore last picks ── */
function persist(){
  var data = JSON.stringify({selA:_selA,selB:_selB,
    company:document.getElementById('company').value,
    companyCustom:document.getElementById('companyCustom').value
  });
  localStorage.setItem('d365_pick', data);
  var s = _store();
  if(s) s.set({'d365_pick': data});
}

function restore(){
  try{
    var d=JSON.parse(localStorage.getItem('d365_pick')||'{}');
    if(d.selA){_selA=d.selA;}
    if(d.selB){_selB=d.selB;}
    refreshPickers();
    if(d.company){document.getElementById('company').value=d.company;}
    if(d.companyCustom)document.getElementById('companyCustom').value=d.companyCustom;
    onCompanySelectChange();
  }catch(e){}
}

/* ── Normalise raw entity records from either endpoint into {name, module, category} ── */
function getEntityAotName(e) {
  return (e && (e.aotName || e.AotName || e.DataEntityAOTName || e.dataEntityAOTName || e.Name || e.TargetName)) || '';
}

function getEntityDmfName(e) {
  return (e && (e.dmfName || e.DmfName || e.EntityName || e.entityName)) || '';
}

function getEntityJoinKey(e) {
  return getEntityAotName(e) || (e && (e.Name || e.TargetName || e.PublicEntityName || e.name || e.url)) || '';
}

function normaliseEntities(raw) {
  var list = Array.isArray(raw) ? raw : (raw && raw.value ? raw.value : []);
  return list.map(function(e) {
    var serviceDoc = isServiceDocumentEntity(e);
    var aotName = getEntityAotName(e);
    var dmfName = getEntityDmfName(e);
    var name = aotName || dmfName || e.PublicEntityName || e.name || e.url || '';
    if (!name) return null;
    var moduleInfo = getModuleInfo(e);
    var category = e.EntityCategory || e.entityCategory || (serviceDoc ? 'serviceRoot' : '');
    var collection = e.PublicCollectionName || e.publicCollectionName || e.PublicCollection || e.publicCollection || '';
    if (!collection && serviceDoc) collection = e.url || e.name || '';
    return {
      name:     name,
      label:    e.PublicEntityName || e.EntityLabel || e.title || dmfName || name,
      aotName:  aotName || name,
      dmfName:  dmfName || '',
      module:   moduleInfo.name,
      moduleExact: moduleInfo.exact,
      moduleSource: moduleInfo.source,
      collection: collection,
      odataEnabled: typeof e.DataServiceEnabled === 'boolean' ? e.DataServiceEnabled : (typeof e.ODataEnabled === 'boolean' ? e.ODataEnabled : !!(collection || serviceDoc)),
      category: category,
      dmEnabled: !!(e.DataManagementEnabled || e.IsDataManagementEnabled),
      serviceDoc: serviceDoc
    };
  }).filter(Boolean).filter(function(e) {
    return isIncludedEntityCategory(e.category) || e.dmEnabled || e.serviceDoc;
  });
}

function isServiceDocumentEntity(e) {
  return !!(e && e.kind === 'EntitySet' && (e.name || e.url));
}

var INCLUDED_ENTITY_CATEGORIES = {
  master: true,
  reference: true,
  parameter: true,
  parameters: true
};

function isIncludedEntityCategory(category) {
  return !!INCLUDED_ENTITY_CATEGORIES[String(category || '').trim().toLowerCase()];
}

var MODULE_GROUP_UNCLASSIFIED = 'Unclassified';
var MODULE_GROUP_RAW = 'Raw Entity Sets';

function normaliseModuleToken(s) {
  return String(s || '').replace(/[^A-Za-z0-9]/g, '').toLowerCase();
}

function splitModuleTags(tags) {
  if (Array.isArray(tags)) {
    return tags.map(function(t) { return String(t || '').trim(); }).filter(Boolean);
  }
  return String(tags || '').split(/[;,|]+/).map(function(t) {
    return t.trim();
  }).filter(Boolean);
}

function findExactModuleFromTags(tags) {
  var parts = splitModuleTags(tags);
  for (var i = 0; i < parts.length; i++) {
    var token = humanizeModuleToken(parts[i]);
    if (normaliseModuleToken(token)) return token;
  }
  return '';
}

function humanizeModuleToken(token) {
  var text = String(token || '').trim();
  if (!text) return '';
  return text
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
    .replace(/[._-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeModuleName(token) {
  var raw = String(token || '').trim();
  if (!raw) return '';
  return humanizeModuleToken(raw);
}

var HEURISTIC_MODULE_RULES = [
  // ── Warehouse & Inventory ──
  { re:/^WHS/i,                                                                                     m:'Warehouse Management' },
  { re:/^(Invent|InventSite|InventTable|InventTrans|InventDim|InventLocation|InventSerial|InventBatch|InventJournal|InventOnhand|InventQuality|InventTransfer|InventCount)/i, m:'Inventory Management' },
  { re:/^(ItemArrival|TransportRoute|ReturnOrder)/i,                                                m:'Inventory Management' },

  // ── Product Information ──
  { re:/^(EcoRes|ReleasedProduct|DistinctProduct|UnitOfMeasure|ProductCategory|ProductMaster|ProductVariant|ProductAttribute|ProductDefault|ProductGroup|ItemGroup|BOMVersion|BOMCalc|ConfigLine)/i, m:'Product Information' },
  { re:/^(PdsBatch|PdsItem|PdsRebate|PdsCumulativeItem)/i,                                         m:'Product Information' },

  // ── Accounts Payable ──
  { re:/^(Vend|Vendor|VendorInvoice|VendorPayment|VendTable|VendTrans|VendInvoice|VendGroup|VendSettlement|VendParameters|VendAccount)/i, m:'Accounts Payable' },
  { re:/^(Paym|Payment|Tax1099|BankVend|VendorPortal)/i,                                            m:'Accounts Payable' },
  { re:/^(PurchLineDisc|PurchPrice|PurchAutoCharges|PurchOrderPool)/i,                              m:'Accounts Payable' },

  // ── Accounts Receivable ──
  { re:/^(Cust|Customer|CustTable|CustTrans|CustGroup|CustAccount|CustInvoice|CustPayment|CustSettlement|CustParameters|CustPosting|CustBalance|CustAgingReport|CustCollect)/i, m:'Accounts Receivable' },
  { re:/^(CollectionLetter|FreeTextInvoice|InterestNote)/i,                                        m:'Accounts Receivable' },

  // ── Sales & Marketing ──
  { re:/^(Sales|SalesOrder|SalesLine|SalesQuotation|SalesDelivery|SalesTable|SalesTrans|SalesGroup|SalesPolicy|SalesReturn)/i, m:'Sales & Marketing' },
  { re:/^(SalesOrderPool|SalesAutoCharges|SalesLineDisc|SalesPrice|SalesCampaign|SalesTarget|SalesTerrit)/i, m:'Sales & Marketing' },
  { re:/^(MCRSales|MCRCust|MCROrder|MCRReturn|MCRContinuity|MCRBroker|MCRInstallment|MCRPayment|MCRSourceCode|MCRCatalog)/i, m:'Call Center' },
  { re:/^(smmActivity|smmBusRel|smmCampaign|smmContact|smmLeads|smmOpportunity|smmQuotation|smmSales)/i, m:'Sales & Marketing' },

  // ── Procurement & Sourcing ──
  { re:/^(Purch|PurchOrder|PurchLine|PurchTable|PurchGroup|PurchReq|PurchRFQ|PurchRebate|PurchContract|PurchPolicy)/i, m:'Procurement' },
  { re:/^(Agreement|TradeAgreement|PriceDisc)/i,                                                   m:'Procurement' },

  // ── General Ledger ──
  { re:/^(Ledger|LedgerJournal|LedgerTrans|LedgerEntry|LedgerPeriod|LedgerAlloc|LedgerTrialBal)/i, m:'General Ledger' },
  { re:/^(MainAccount|Dimension|FiscalCalendar|FiscalYear|FiscalPeriod|ExchangeRate|CurrencyExchange|Currency)/i, m:'General Ledger' },
  { re:/^(Fiscal|Consol|Consolidat|LedgerFinancial|FinancialReport|FinancialStatement)/i,          m:'General Ledger' },
  { re:/^(Voucher|GeneralJournal|PeriodicJournal)/i,                                               m:'General Ledger' },

  // ── Finance (general) ──
  { re:/^(Financial|Finance|Subledger|AccountingDistrib|AccountingEvent|AccountEntry|FundAccounting)/i, m:'Finance' },

  // ── Tax ──
  { re:/^(Tax|TaxGroup|TaxCode|TaxTrans|TaxAdjust|TaxReport|TaxTable|TaxLedger|TaxWithhold|TaxDeclar|Excise|CustomsDuty|TaxRegistration)/i, m:'Tax' },

  // ── Cash & Bank ──
  { re:/^(Bank|BankAccount|BankTrans|BankJournal|BankReconcile|BankStatement|BankGroup|BankDeposit|BankCheck|BankLetter)/i, m:'Cash & Bank' },
  { re:/^(CashDisc|CashFlow|Cheque)/i,                                                             m:'Cash & Bank' },

  // ── Fixed Assets ──
  { re:/^(Asset|FixedAsset|AssetBook|AssetTrans|AssetGroup|AssetDepreciation|AssetDisposal|AssetAcquisition|AssetLease)/i, m:'Fixed Assets' },
  { re:/^(RAsset|ROU)/i,                                                                            m:'Fixed Assets' },

  // ── Budgeting ──
  { re:/^(Budget|BudgetCycle|BudgetPlan|BudgetControl|BudgetRegister|BudgetAlloc|BudgetReservation|BudgetForecast|Forecast)/i, m:'Budgeting' },

  // ── Project Management & Accounting ──
  { re:/^(Proj|Project|ProjTable|ProjTrans|ProjGroup|ProjForecast|ProjContract|ProjInvoice|ProjCategory|ProjBudget|ProjLine|ProjWorker|ProjPosting|ProjCost|ProjRevenue|ProjResource|ProjTimesheet)/i, m:'Project Management' },
  { re:/^(PSAContractLine|PSAProject)/i,                                                           m:'Project Management' },

  // ── Human Resources ──
  { re:/^(HRM|Hcm|HcmWorker|HcmPosition|HcmDepartment|HcmJob|HcmSkill|HcmLeave|HcmBenefit|HcmCompensation|HcmPerformance|HcmEmployment|HcmTraining|HcmCourse|HcmAbsence)/i, m:'Human Resources' },
  { re:/^(Worker|Employee|Applicant|Recruitment)/i,                                                 m:'Human Resources' },

  // ── Payroll ──
  { re:/^(Payroll|PayStatement|PayPeriod|PayCycle|EarningCode|EarningLine|BenefitAccrual|PayrollTax|PayrollWorker|PayrollDeduction)/i, m:'Payroll' },

  // ── Production Control ──
  { re:/^(Prod|ProdOrder|ProdTable|ProdBOM|ProdRoute|ProdJournal|ProdCostEstimation|ProdSubContract)/i, m:'Production Control' },
  { re:/^(BOM|Route|Kanban|WrkCtr|ProdCalc|ProdParameter|JobCard|ProdPickingList|ProductionFlow)/i, m:'Production Control' },
  { re:/^(LeanProductionFlow|LeanSchedule|KanbanJob)/i,                                            m:'Production Control' },

  // ── Master Planning ──
  { re:/^(Req|ReqPlan|ReqTrans|ReqForecast|ReqBOM|MRP|Coverage|CovPlan|IntercompanyPlanning)/i,   m:'Master Planning' },
  { re:/^(ForecastSupply|ForecastDemand|DemandForecast)/i,                                         m:'Master Planning' },

  // ── Retail & Commerce ──
  { re:/^(Retail|RetailStore|RetailChannel|RetailPos|RetailTerminal|RetailTransaction|RetailProduct|RetailCategory|RetailCatalog|RetailCust|RetailDiscount|RetailEmployee|RetailGift|RetailInventory|RetailLoyalty|RetailPricing|RetailShipping|RetailTender|RetailComm|RetailConnDist)/i, m:'Retail & Commerce' },

  // ── Transportation & Logistics ──
  { re:/^(TMSRoute|TMSLoad|TMSShipment|TMSCarrier|TMSFreight|TMSHub|TMSTransport|TMSEngine|TransportRoute)/i, m:'Transportation Management' },
  { re:/^(WHSShip|WHSLoad|WHSWork)/i,                                                              m:'Warehouse Management' },

  // ── Service Management ──
  { re:/^(SMAService|SMAContract|SMAObject|SMAOrder|SMAAgreement|SMASubscription|SMATemplate|SMARepair)/i, m:'Service Management' },

  // ── Global Address Book ──
  { re:/^(Dir|DirParty|DirPerson|DirOrg|Logistics|Address|CountryRegion|ZipCode|State|City|County|ContactPerson|Party)/i, m:'Global Address Book' },

  // ── Organization Administration ──
  { re:/^(OM|CompanyInfo|LegalEntity|DataArea|NumberSeq|ReasonCode|Organization|OrgUnit|OperatingUnit|Hierarchy|DimensionHierarchy)/i, m:'Organization Administration' },
  { re:/^(Sys|SystemParam|SysEmail|SysWorkflow|SysUser|BatchJob|BatchGroup|DocuType|Note)/i,       m:'System Administration' },
  { re:/^(UserGroup|SecurityRole|SecurityDuty|SecurityPrivilege|AccessRight)/i,                   m:'System Administration' },

  // ── Cost Management ──
  { re:/^(Cost|CostCategory|CostGroup|CostSheet|CostSharing|CostAdjust|CostInventory|InventCost)/i, m:'Cost Management' },

  // ── Credit & Collections ──
  { re:/^(Credit|CreditLimit|CreditHold|Aging|Collect)/i,                                         m:'Credit & Collections' },

  // ── Expense Management ──
  { re:/^(TrvExpense|TrvAdv|TrvPolicy|TrvMileage|TrvCash|TrvUnsettled|TrvParameters|Travel)/i,    m:'Expense Management' },

  // ── Asset Leasing ──
  { re:/^(AssetLease|ROU|IFRS16|LeaseBook|LeaseJournal)/i,                                        m:'Asset Leasing' },

  // ── Public Sector ──
  { re:/^(PsaPublic|PsaGrant|PsaFund|PublicSector|PSA)/i,                                         m:'Public Sector' },

  // ── Fleet Management ──
  { re:/^(FMVehicle|FMCustomer|FMRental|FMReservation|FMFacility|Fleet)/i,                        m:'Fleet Management' },

  // ── Intercompany ──
  { re:/^(Intercompany|ICust|IVend|InterComp)/i,                                                   m:'Intercompany' },

  // ── Electronic Reporting / Regulatory ──
  { re:/^(ERFormat|ERModel|ERConfig|ERSolution|ElectronicReport)/i,                                m:'Electronic Reporting' },
  { re:/^(Regulatory|RCS|GlobalizationStudio)/i,                                                   m:'Regulatory' },

  // ── Subscription Billing / Revenue Recognition ──
  { re:/^(SubBilling|SubscriptionBilling|RevRec|RevenueRecognition|RevenueSplit)/i,                m:'Subscription Billing' },

  // ── Rebate Management ──
  { re:/^(Rebate|RebateProg|PdsRebateProg|TAMRebate)/i,                                            m:'Rebate Management' },

  // ── Credit Management ──
  { re:/^(CreditMgmt|CreditManagement)/i,                                                          m:'Credit Management' },

  //── Landed Cost ──
  { re:/^(ITM|LandedCost|ItmVoyage|ItmContainer|ItmShipment|ItmFolio)/i,                          m:'Landed Cost' },

  // ── Advanced Bank Reconciliation ──
  { re:/^(BankStmtIso|BankReconcAdv|BankStmtFormat)/i,                                            m:'Cash & Bank' },

  // ── Questionnaire / Survey ──
  { re:/^(KM|KMQuestionnaire|KMQuestion|KMAnswer|KMForm|KMKnowledge)/i,                           m:'Questionnaire' },

  // ── Case Management ──
  { re:/^(Case|CaseDetail|CaseLog|CaseCategory|CaseAssociation)/i,                                 m:'Case Management' },

  // ── Vendor Portal / Collaboration ──
  { re:/^(VendorPortal|VendCollaboration|VendorCollab|PurchVendorPortal)/i,                       m:'Vendor Collaboration' },

  // ── Customer Portal ──
  { re:/^(CustPortal|CustomerPortal)/i,                                                             m:'Customer Collaboration' },

  // ── Compliance & Audit ──
  { re:/^(Audit|AuditPolicy|Compliance|PolicyViolation|PolicyRule)/i,                              m:'Compliance' },

  // ── Workflow ──
  { re:/^(Workflow|WFTracking|WFApproval|WorkflowElement|WorkflowTable)/i,                        m:'Workflow' },

  // ── Interoperability / Integration ──
  { re:/^(CDSVirtual|CDS|DualWrite|DualWriteMap|MicrosoftDataverse)/i,                             m:'Dual-Write / Dataverse' }

  // NOTE: ISV-specific prefixes (e.g. A365*, custom partner namespaces) are intentionally
  // NOT included here — they must come from the OData Module field returned by the D365
  // environment itself. If Module is blank for an ISV entity, it shows as Unclassified.
];


function inferModuleFromEntityName(name) {
  var value = String(name || '').trim();
  if (!value) return '';
  for (var i = 0; i < HEURISTIC_MODULE_RULES.length; i++) {
    if (HEURISTIC_MODULE_RULES[i].re.test(value)) return HEURISTIC_MODULE_RULES[i].m;
  }
  return '';
}

function hasExactModuleMetadata(e) {
  return !!(e.Modules || e.modules || e.AppModule || e.appModule || e.ApplicationModule || e.applicationModule || e.Module || e.module || e.ModuleName);
}

function getModuleInfo(e) {
  var serviceDoc = isServiceDocumentEntity(e);
  var entityName = e.EntityName || e.Name || e.PublicEntityName || e.name || e.url || '';
  var direct = e.Modules || e.modules || e.AppModule || e.appModule || e.ApplicationModule || e.Module || e.module || e.ModuleName || '';
  if (String(direct || '').trim()) {
    return { name: normalizeModuleName(direct), source: 'field', exact: true };
  }

  var fromTags = findExactModuleFromTags(e.Tags || e.tags || e.Tag || e.tag || '');
  if (String(fromTags || '').trim()) {
    return { name: normalizeModuleName(fromTags), source: 'tags', exact: false };
  }

  var inferred = inferModuleFromEntityName(entityName);
  if (String(inferred || '').trim()) {
    return { name: inferred, source: 'heuristic', exact: false };
  }

  // No more "Derived / Xyz" fragmentation — anything unrecognised goes to Unclassified.
  // This keeps the module dropdown clean with real module names only.
  return {
    name: serviceDoc ? MODULE_GROUP_RAW : MODULE_GROUP_UNCLASSIFIED,
    source: 'none',
    exact: false
  };
}

/* ── Rebuild module filter dropdown — grouped by first letter bucket or explicit group ── */
function rebuildModuleFilter(rows) {
  var modules = {};
  rows.forEach(function(r) {
    var moduleName = r.module || MODULE_GROUP_UNCLASSIFIED;
    if (!modules[moduleName]) modules[moduleName] = 0;
    modules[moduleName]++;
  });
  var sorted = Object.keys(modules).sort();
  var sel = document.getElementById('modSel');
  var prev = sel.value;
  var html = '<option value="">All Module Groups (' + rows.length + ' entities)</option>';

  sorted.forEach(function(m) {
    html += '<option value="' + esc(m) + '">' + esc(m) + ' (' + modules[m] + ')</option>';
  });

  sel.innerHTML = html;
  if (prev && modules[prev]) sel.value = prev;
  updateReportButtonState();
}

function updateReportButtonState() {
  var btn = document.getElementById('btnReport');
  var sel = document.getElementById('modSel');
  if (!btn || !sel) return;

  var moduleName = sel.value || '';
  var hasRows = Array.isArray(STATE.allRows) && STATE.allRows.length > 0;
  var enabled = !!moduleName && hasRows;

  btn.disabled = !enabled;
  btn.title = enabled
    ? 'Generate a standalone HTML comparison report for the selected module only'
    : 'Select a specific module group first. Full-report generation for All Module Groups is disabled.';
}

/* ── Derive module from a raw entity record using runtime metadata only ── */
function deriveModule(e) {
  return getModuleInfo(e).name;
}

/* ── Find D365 tab by hostname ── */
async function findD365Tab(envUrl) {
  var host = new URL(envUrl.replace(/\/+$/, '')).hostname;
  var tabs = await chrome.tabs.query({});
  var tab = tabs.find(function(t) {
    try { return new URL(t.url || '').hostname === host; }
    catch { return false; }
  });
  return { tab: tab, host: host };
}

/* ── Low-level: try sending a message, resolve 'NO_LISTENER' if content script absent ── */
function _trySendMessage(tabId, msgObj) {
  return new Promise(function(resolve) {
    chrome.tabs.sendMessage(tabId, msgObj, function(res) {
      if (chrome.runtime.lastError) {
        resolve('NO_LISTENER');
        return;
      }
      resolve(res);
    });
  });
}

/* ── Ensure content script is loaded in tab, then send message ── */
async function askTab(tabId, msgObj) {
  var res = await _trySendMessage(tabId, msgObj);
  if (res !== 'NO_LISTENER') return res;

  // Inject content.js and retry
  try {
    await chrome.scripting.executeScript({ target: { tabId: tabId }, files: ['content.js'] });
  } catch(e) {
    throw new Error('Could not inject into D365 tab: ' + e.message);
  }
  await new Promise(function(r){ setTimeout(r, 300); });

  var res2 = await _trySendMessage(tabId, msgObj);
  if (res2 === 'NO_LISTENER') throw new Error('Content script not responding. Reload the D365 tab.');
  return res2;
}

/* ── Probe a single endpoint — returns {ok, status, data, detail} ── */
async function probeEndpoint(tabId, endpoint) {
  return askTab(tabId, { type: 'FETCH_ENTITIES', endpoint: endpoint });
}

/* ── Build candidate endpoints ──
   OData first: /data/DataEntities returns a `Module` field per entity.
   Metadata:    /Metadata/DataEntities has richer entity metadata but NO Module field.
   We try OData first so module data is populated; fall back to Metadata if OData is blocked. ── */
function getODataCandidates(origin) {
  return [
    origin + '/data/DataEntities?$top=10000&cross-company=true',
    origin + '/data/DataEntities?$top=10000',
    origin + '/data/DataEntities',
    origin + '/data/dataentities?$top=10000&cross-company=true',
    origin + '/data/dataentities?$top=10000',
    origin + '/data/dataentities'
  ];
}

function getMetadataCandidates(origin) {
  return [
    origin + '/Metadata/DataEntities',
    origin + '/metadata/DataEntities',
    origin + '/metadata/dataentities'
  ];
}

function getDataManagementCandidates(origin) {
  return [
    origin + '/data/DataManagementEntities?$top=10000&cross-company=true',
    origin + '/data/DataManagementEntities?$top=10000',
    origin + '/data/DataManagementEntities'
  ];
}

function getServiceDocCandidates(origin) {
  return [
    origin + '/data',
    origin + '/data/'
  ];
}

function getCandidateEndpoints(origin) {
  // OData first so Module field is populated, then Metadata, then service doc as last resort
  return [].concat(
    getODataCandidates(origin),
    getMetadataCandidates(origin),
    getServiceDocCandidates(origin)
  );
}

/* ── Try a list of endpoints against a tab, return first raw JSON that yields usable entities ── */
async function _tryEndpointList(tabId, endpoints, host) {
  var lastErr = '';
  for (var i = 0; i < endpoints.length; i++) {
    var res = await probeEndpoint(tabId, endpoints[i]);
    if (!res) continue;
    if (res.ok) {
      var usable = normaliseEntities(res.data);
      if (usable.length) return { data: res.data, endpoint: endpoints[i] };
      lastErr = 'Endpoint returned 0 usable entities after filtering';
      continue;
    }
    if (res.status === 401) throw new Error('401 \u2014 Not authorised. Make sure you are logged in to ' + host);
    lastErr = 'HTTP ' + res.status + (res.detail ? ' \u2014 ' + res.detail.slice(0, 200) : '');
  }
  return { data: null, lastErr: lastErr };
}

async function _tryRawEndpointList(tabId, endpoints, host) {
  var lastErr = '';
  for (var i = 0; i < endpoints.length; i++) {
    var res = await probeEndpoint(tabId, endpoints[i]);
    if (!res) continue;
    if (res.ok) {
      return { data: res.data, endpoint: endpoints[i] };
    }
    if (res.status === 401) throw new Error('401 \u2014 Not authorised. Make sure you are logged in to ' + host);
    lastErr = 'HTTP ' + res.status + (res.detail ? ' \u2014 ' + res.detail.slice(0, 200) : '');
  }
  return { data: null, lastErr: lastErr };
}

/* ── Merge module info from an OData result into a Metadata result ──
   Metadata has richer entity metadata; OData has the Module field.
   We build a name→module map from OData and stamp it onto Metadata records. ── */
function _mergeModuleData(metaRaw, odataRaw) {
  var odataList = Array.isArray(odataRaw) ? odataRaw : (odataRaw && odataRaw.value ? odataRaw.value : []);
  var moduleMap = {};
  odataList.forEach(function(e) {
    var name = getEntityJoinKey(e);
    var mod = e.AppModule || e.appModule || e.ApplicationModule || e.Module || e.module || e.ModuleName || '';
    if (name && mod) moduleMap[name] = mod;
  });

  var metaList = Array.isArray(metaRaw) ? metaRaw : (metaRaw && metaRaw.value ? metaRaw.value : []);
  var merged = metaList.map(function(e) {
    var name = getEntityJoinKey(e);
    if (name && moduleMap[name] && !e.Module && !e.AppModule) {
      return Object.assign({}, e, { Module: moduleMap[name] });
    }
    return e;
  });
  return Array.isArray(metaRaw) ? merged : Object.assign({}, metaRaw, { value: merged });
}

function _mergeDataManagementData(entityRaw, dataManagementRaw) {
  var dataManagementList = Array.isArray(dataManagementRaw) ? dataManagementRaw : (dataManagementRaw && dataManagementRaw.value ? dataManagementRaw.value : []);
  var dmfMap = {};
  dataManagementList.forEach(function(e) {
    var targetName = (e && (e.TargetName || e.targetName)) || '';
    if (!targetName) return;
    dmfMap[targetName] = {
      dmfName: getEntityDmfName(e),
      modules: e.Modules || e.modules || '',
      isShared: e.IsShared,
      dataManagementEnabled: e.DataManagementEnabled
    };
  });

  var entityList = Array.isArray(entityRaw) ? entityRaw : (entityRaw && entityRaw.value ? entityRaw.value : []);
  var merged = entityList.map(function(e) {
    var joinKey = getEntityJoinKey(e);
    var dmfEntry = joinKey && dmfMap[joinKey];
    if (!dmfEntry) return e;
    var patch = {};
    if (dmfEntry.dmfName && !getEntityDmfName(e)) patch.DmfName = dmfEntry.dmfName;
    if (dmfEntry.modules && !e.Modules && !e.Module && !e.AppModule) patch.Modules = dmfEntry.modules;
    if (dmfEntry.isShared !== undefined && e.IsShared === undefined) patch.IsShared = dmfEntry.isShared;
    if (dmfEntry.dataManagementEnabled !== undefined && e.DataManagementEnabled === undefined) patch.DataManagementEnabled = dmfEntry.dataManagementEnabled;
    return Object.keys(patch).length ? Object.assign({}, e, patch) : e;
  });

  return Array.isArray(entityRaw) ? merged : Object.assign({}, entityRaw, { value: merged });
}

/* ── Fetch ── */
async function fetchEntities(envUrl, slot){
  var origin;
  try { origin = new URL(envUrl).origin; }
  catch(e) { throw new Error('Invalid URL: ' + envUrl); }

  if(IS_EXT){
    var found = await findD365Tab(envUrl);
    if(!found.tab) throw new Error(
      'No open tab for "' + found.host + '". Open ' + envUrl + ' in a Chrome tab and log in first.');

    var tabUrl = found.tab.url || '';
    if(tabUrl.includes('login.microsoftonline') || tabUrl.includes('login.live')) {
      throw new Error('The D365 tab is on the Microsoft login page. Please finish logging in first.');
    }

    var tabId = found.tab.id;
    var host = found.host;

    // ── Pass 1: try OData (has Module field) ──
    var odataResult = await _tryEndpointList(tabId, getODataCandidates(origin), host);

    // ── Pass 2: try Metadata (richer metadata, no Module field) ──
    var metaResult = await _tryEndpointList(tabId, getMetadataCandidates(origin), host);

    // ── Pass 3: try DataManagementEntities (TargetName -> EntityName / Modules) ──
    var dmfResult = await _tryRawEndpointList(tabId, getDataManagementCandidates(origin), host);

    var mergedData = null;

    if (odataResult.data && metaResult.data) {
      // Best case: merge module data from OData into the richer Metadata records
      mergedData = _mergeModuleData(metaResult.data, odataResult.data);
    } else if (odataResult.data) {
      // OData only — Module field is present directly
      mergedData = odataResult.data;
    } else if (metaResult.data) {
      // Metadata only — module will fall back to heuristics (no OData available)
      mergedData = metaResult.data;
    }

    if (mergedData && dmfResult.data) {
      mergedData = _mergeDataManagementData(mergedData, dmfResult.data);
    }
    if (mergedData) {
      return mergedData;
    }

    // ── Pass 4: last resort service doc ──
    var svcResult = await _tryEndpointList(tabId, getServiceDocCandidates(origin), host);
    if (svcResult.data) return svcResult.data;

    var lastErr = odataResult.lastErr || metaResult.lastErr || dmfResult.lastErr || svcResult.lastErr || 'Unknown error';
    throw new Error('All endpoints failed. Last error: ' + lastErr + '\nTried: ' + getCandidateEndpoints(origin).join(', '));

  } else {
    // ── Proxy mode: same two-pass strategy ──
    var token = getToken(slot);

    async function proxyFetch(url) {
      var proxyUrl = 'http://localhost:8888/proxy?url=' + encodeURIComponent(url);
      if (token) proxyUrl += '&token=' + encodeURIComponent(token);
      var r;
      try { r = await fetch(proxyUrl, { method: 'GET' }); }
      catch(e) { throw new Error('Cannot reach local server. Is Launch-Tool.bat running? (' + (e.message||e) + ')'); }
      if (r.status === 401) throw new Error('401 \u2014 Not authorised. Check your bearer token is correct and not expired.');
      if (r.ok) return r.json();
      return null;
    }

    async function proxyTryList(endpoints) {
      for (var i = 0; i < endpoints.length; i++) {
        try {
          var j = await proxyFetch(endpoints[i]);
          if (j) {
            var usable = normaliseEntities(j);
            if (usable.length) return j;
          }
        } catch(e) {
          if (/401|Cannot reach/.test(e.message)) throw e;
        }
      }
      return null;
    }

    async function proxyTryRawList(endpoints) {
      for (var i = 0; i < endpoints.length; i++) {
        try {
          var j = await proxyFetch(endpoints[i]);
          if (j) return j;
        } catch(e) {
          if (/401|Cannot reach/.test(e.message)) throw e;
        }
      }
      return null;
    }

    var odataData = await proxyTryList(getODataCandidates(origin));
    var metaData = await proxyTryList(getMetadataCandidates(origin));

    var dmfData = await proxyTryRawList(getDataManagementCandidates(origin));

    var mergedProxyData = null;

    if (odataData && metaData) mergedProxyData = _mergeModuleData(metaData, odataData);
    else if (odataData) mergedProxyData = odataData;
    else if (metaData) mergedProxyData = metaData;

    if (mergedProxyData && dmfData) {
      mergedProxyData = _mergeDataManagementData(mergedProxyData, dmfData);
    }
    if (mergedProxyData) return mergedProxyData;

    var svcData = await proxyTryList(getServiceDocCandidates(origin));
    if (svcData) return svcData;

    throw new Error('All proxy endpoints failed.');
  }
}

/* ── Validate ── */
async function validateAccess(){
  var urlA=getEnvUrl('A'), urlB=getEnvUrl('B');
  if(!urlA&&!urlB){toast('\u26A0\uFE0F Select or enter at least one environment.');return}
  refreshLegalEntityFilters().catch(function(){}); // Ensure filters are refreshed
  if(urlA&&!isHttps(urlA)){toast('\u26A0\uFE0F Source URL must start with https://');return}
  if(urlB&&!isHttps(urlB)){toast('\u26A0\uFE0F Target URL must start with https://');return}
  var btn=document.getElementById('btnVal'), sp=document.getElementById('spinVal');
  btn.disabled=true; sp.style.display='inline-block';

  async function check(url, stId, label, slot){
    setSt(stId,'loading',label+' \u2014 checking...');
    var origin;
    try { origin = new URL(url).origin; }
    catch(e) { setSt(stId,'error',label+' \u2014 Invalid URL'); return; }

    if(IS_EXT){
      var found = await findD365Tab(url);
      if(!found.tab){
        setSt(stId,'error',label+' \u2014 No open tab for '+found.host+'. Open & log in first.');
        return;
      }
      var tabUrl = found.tab.url || '';
      if(tabUrl.includes('login.microsoftonline') || tabUrl.includes('login.live')) {
        setSt(stId,'error',label+' \u2014 Tab is on login page. Finish logging in to D365 first.');
        return;
      }
      // Try all candidate endpoints — report which one worked or all errors
      var candidates = getCandidateEndpoints(origin).map(function(u){
        // Use $top=1 for a quick ping
        return u.replace('$top=10000','$top=1').replace(/&\$select=[^&]*/,'');
      });
      var working = null, lastStatus = '', lastDetail = '';
      for(var ci=0; ci<candidates.length; ci++){
        try{
          var res = await probeEndpoint(found.tab.id, candidates[ci]);
          if(!res) continue;
          if(res.ok){ working = candidates[ci]; break; }
          if(res.status===401){
            setSt(stId,'error',label+' \u2014 401 Not authorised. Log in to the D365 tab first.');
            return;
          }
          lastStatus = 'HTTP '+res.status;
          lastDetail = res.detail ? res.detail.slice(0,200) : '';
        }catch(e){ lastDetail = e.message; }
      }
      if(working){
        var viaPath = new URL(working).pathname;
        var approx = /\/data\/?$/i.test(viaPath) ? ' using OData service root fallback' : '';
        setSt(stId,'ok',label+' \u2014 Connected \u2713 via '+viaPath+approx+' ('+esc(found.tab.title||found.host)+')');
      } else {
        setSt(stId,'error',label+' \u2014 All endpoints failed ('+lastStatus+'). '+
          (lastDetail||'Open the D365 tab, log in, then click \u1F9EA Diagnose for details.'));
      }
    } else {
      var proxyUrl='http://localhost:8888/proxy?url='+encodeURIComponent(buildEndpoint(origin,'DataEntities?$top=1&$select=Name&cross-company=true'));
      var token=getToken(slot);
      if(token) proxyUrl+='&token='+encodeURIComponent(token);
      try{
        var r=await Promise.race([
          fetch(proxyUrl,{method:'GET'}),
          new Promise(function(_,rej){setTimeout(function(){rej(new Error('Timed out after 12s'))},12000)})
        ]);
        if(r.status===401) setSt(stId,'error',label+' \u2014 401 Not authorised. Paste a valid bearer token.');
        else if(!r.ok) setSt(stId,'error',label+' \u2014 HTTP '+r.status+' response from D365.');
        else setSt(stId,'ok',label+' \u2014 reachable and authenticated \u2713');
      }catch(e){ setSt(stId,'error',label+' \u2014 '+e.message); }
    }
  }

  var tasks=[];
  if(urlA) tasks.push(check(urlA,'stA',getEnvLabel('A'),'A'));
  if(urlB) tasks.push(check(urlB,'stB',getEnvLabel('B'),'B'));
  await Promise.all(tasks);
  btn.disabled=false; sp.style.display='none';
}

/* ── Load Entities ── */
async function loadEntities(){
  var urlA=getEnvUrl('A'),urlB=getEnvUrl('B');
  var lblA=getEnvLabel('A'),lblB=getEnvLabel('B');
  if(!urlA||!urlB){toast('\u26A0\uFE0F Select both Source and Target environments first.');return}
  var btn=document.getElementById('btnLoad'),sp=document.getElementById('spinLoad');
  btn.disabled=true;sp.style.display='inline-block';
  setSt('stA','loading',lblA+' \u2014 loading entities...');
  setSt('stB','loading',lblB+' \u2014 loading entities...');
  showProgress('Load & Compare', 'Preparing environment metadata...', 5);
  async function tryLoad(url,stId,label,slot){
    try{
      showProgress('Load & Compare', 'Loading data entities from ' + label + '...', slot==='A' ? 20 : 60);
      var raw = await fetchEntities(url, slot);
      var entities = normaliseEntities(raw);
      if(!entities.length) throw new Error('No entities returned — endpoint responded but list is empty.');
      setSt(stId,'ok',label+' \u2014 '+entities.length+' entities loaded \u2713');
      return { ok:true, entities:entities };
    }catch(e){ setSt(stId,'error',label+' \u2014 '+e.message); return { ok:false, entities:[] }; }
  }
  var resA = await tryLoad(urlA,'stA',lblA,'A');
  showProgress('Load & Compare', 'Merging Source entities and loading Target...', 50);
  var resB = await tryLoad(urlB,'stB',lblB,'B');
  var res=[resA,resB];
  btn.disabled=false;sp.style.display='none';
  if(!res[0].ok||!res[1].ok){showProgress('Load & Compare', 'Loading failed. Check the status boxes above.', 100); hideProgress(1800); toast('\u26A0\uFE0F One or both environments failed. See status above.');return}
  STATE.entitiesA=res[0].entities;
  STATE.entitiesB=res[1].entities;
  STATE.entityMapA=indexEntitiesByName(res[0].entities);
  STATE.entityMapB=indexEntitiesByName(res[1].entities);
  showProgress('Load & Compare', 'Building combined module list...', 85);
  STATE.allRows=buildRows(res[0].entities,res[1].entities);
  STATE.lblA=lblA;STATE.lblB=lblB;
  STATE.activeModule='';
  STATE.moduleDetailRows=[];
  STATE.visibleModuleDetailRows=[];
  document.getElementById('modDetailPanel').style.display='none';
  document.getElementById('entityDiffPanel').style.display='none';
  rebuildModuleFilter(STATE.allRows);
  updateReportButtonState();
  refreshLegalEntityFilters().catch(function(){});
  setSetupCardCollapsed(true);
  focusModuleFiltersCard();
  showProgress('Load & Compare', 'Finished. Module list is ready.', 100);
  hideProgress(1200);
  toast('\u2705 Loaded & compared successfully.');
}

/* ── Compare ── */
var STATE={allRows:[],filteredRows:[],lblA:'Source',lblB:'Target',entitiesA:[],entitiesB:[],entityMapA:{},entityMapB:{},activeModule:'',moduleDetailRows:[],visibleModuleDetailRows:[]};
function entityMetaScore(e) {
  if (!e) return -1;
  var score = 0;
  if (e.moduleExact) score += 4;
  else if (e.moduleSource === 'tags') score += 3;
  else if (e.moduleSource === 'heuristic') score += 2;
  if (e.collection) score += 2;
  if (e.odataEnabled) score += 1;
  return score;
}

function chooseBetterEntity(a, b) {
  if (!a) return b;
  if (!b) return a;
  return entityMetaScore(b) > entityMetaScore(a) ? b : a;
}

function getBestEntityLabel(a, b) {
  var candidates = [a, b, chooseBetterEntity(a, b)].filter(Boolean);
  for (var i = 0; i < candidates.length; i++) {
    var candidate = candidates[i];
    if (candidate.label && candidate.label !== candidate.name) return candidate.label;
  }
  var best = chooseBetterEntity(a, b) || a || b;
  return best ? (best.label || best.name || '') : '';
}

function indexEntitiesByName(list) {
  var map = {};
  list.forEach(function(e) {
    map[e.name] = chooseBetterEntity(map[e.name], e);
  });
  return map;
}

function buildRows(aList,bList){
  var map={};
  function upsert(list, sideKey) {
    list.forEach(function(e) {
      if (!map[e.name]) {
        map[e.name] = { entity: e };
      } else {
        map[e.name].entity = chooseBetterEntity(map[e.name].entity, e);
      }
      map[e.name][sideKey] = true;
      map[e.name][sideKey === 'inA' ? 'entityA' : 'entityB'] = e;
    });
  }
  upsert(aList, 'inA');
  upsert(bList, 'inB');
  return Object.entries(map).map(function(entry){
    var name=entry[0],it=entry[1],inA=!!it.inA,inB=!!it.inB;
    var moduleName =
      (it.entity && it.entity.module) ||
      (it.entityA && it.entityA.module) ||
      (it.entityB && it.entityB.module) ||
      MODULE_GROUP_UNCLASSIFIED;
    var moduleSource =
      (it.entity && it.entity.moduleSource) ||
      (it.entityA && it.entityA.moduleSource) ||
      (it.entityB && it.entityB.moduleSource) ||
      'none';
    var label =
      (it.entity && it.entity.label) ||
      (it.entityA && it.entityA.label) ||
      (it.entityB && it.entityB.label) ||
      name;
    var collection =
      (it.entity && it.entity.collection) ||
      (it.entityA && it.entityA.collection) ||
      (it.entityB && it.entityB.collection) ||
      '';
    return{
      name:name,
      label:label,
      aotName:(it.entity && it.entity.aotName) || (it.entityA && it.entityA.aotName) || (it.entityB && it.entityB.aotName) || name,
      dmfName:(it.entity && it.entity.dmfName) || (it.entityA && it.entityA.dmfName) || (it.entityB && it.entityB.dmfName) || name,
      publicCollectionName:collection,
      module:moduleName,
      moduleSource:moduleSource,
      collectionA:(it.entityA && it.entityA.collection) || '',
      collectionB:(it.entityB && it.entityB.collection) || '',
      inA:inA,
      inB:inB,
      status:inA&&inB?'Match':inA?'Only in Source':'Only in Target'
    };
  }).sort(function(a,b){return a.name.localeCompare(b.name)});
}

function isDifferenceStatus(status) {
  return status !== 'Match';
}

function stableJsonValue(value) {
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return '[' + value.map(stableJsonValue).join(',') + ']';
  var keys = Object.keys(value).sort();
  return '{' + keys.map(function(k) { return JSON.stringify(k) + ':' + stableJsonValue(value[k]); }).join(',') + '}';
}

function stringifyFieldValue(value) {
  if (value === undefined) return '—';
  if (value === null) return 'null';
  if (typeof value === 'object') return stableJsonValue(value);
  return String(value);
}

var TECHNICAL_FIELD_RE = /^(@|dataAreaId$|RecId$|_Etag|modifiedDateTime|createdDateTime)/i;

function isComparableField(name) {
  return !TECHNICAL_FIELD_RE.test(String(name || ''));
}

function formatNumber(value) {
  var num = Number(value || 0);
  return isFinite(num) ? num.toLocaleString() : '0';
}

function formatPercent(value) {
  var num = Number(value || 0);
  return (isFinite(num) ? num : 0).toFixed(1) + '%';
}

function clampPercent(value) {
  var num = Number(value || 0);
  if (!isFinite(num)) return 0;
  return Math.max(0, Math.min(100, num));
}

function fileSafeName(value) {
  return String(value || 'report').replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '') || 'report';
}

function summarizeEntityResult(result) {
  var total = result.matched + result.diffCount + result.missingInTarget + result.onlyInTarget;
  var alignment = total ? (result.matched / total) * 100 : 100;
  var issue = result.diffCount > 0 || result.missingInTarget > 0 || result.onlyInTarget > 0;
  return {
    total: total,
    alignmentPct: alignment,
    issue: issue,
    issueClass: issue ? 'issues' : 'ok'
  };
}

function shouldDisplayDiffEntity(row) {
  return !!row && row.status !== 'Match' && row.status !== 'No OData';
}

async function compareEntityRecords(row, idx, urlA, urlB) {
  var metaA = STATE.entityMapA[row.name] || null;
  var metaB = STATE.entityMapB[row.name] || null;
  var best = metaA || metaB;
  if (!best || !best.collection) {
    return {
      idx: idx + 1,
      name: row.name,
      label: row.label || getBestEntityLabel(metaA, metaB) || row.name,
      aotName: (best && best.aotName) || row.aotName || row.name,
      dmfName: (best && best.dmfName) || row.dmfName || row.label || row.name,
      publicCollectionName: '',
      module: row.module || '',
      countA: null,
      countB: null,
      matched: 0,
      diffCount: 0,
      missingInTarget: 0,
      onlyInTarget: 0,
      total: 0,
      alignmentPct: 0,
      status: 'No OData',
      detail: 'No OData collection',
      metaA: metaA,
      metaB: metaB,
      noOdata: true
    };
  }

  var resA = await fetchCollectionRows(urlA, 'A', metaA || metaB);
  var resB = await fetchCollectionRows(urlB, 'B', metaB || metaA);
  if (!resA.ok && !resB.ok) {
    return {
      idx: idx + 1,
      name: row.name,
      label: row.label || getBestEntityLabel(metaA, metaB) || row.name,
      aotName: (best && best.aotName) || row.aotName || row.name,
      dmfName: (best && best.dmfName) || row.dmfName || row.label || row.name,
      publicCollectionName: (best && best.collection) || '',
      module: row.module || '',
      countA: null,
      countB: null,
      matched: 0,
      diffCount: 0,
      missingInTarget: 0,
      onlyInTarget: 0,
      total: 0,
      alignmentPct: 0,
      status: 'No OData',
      detail: resA.detail || resB.detail || 'Entity query failed',
      metaA: metaA,
      metaB: metaB,
      noOdata: true
    };
  }
  if (!resA.ok || !resB.ok) {
    return {
      idx: idx + 1,
      name: row.name,
      label: row.label || getBestEntityLabel(metaA, metaB) || row.name,
      aotName: (best && best.aotName) || row.aotName || row.name,
      dmfName: (best && best.dmfName) || row.dmfName || row.label || row.name,
      publicCollectionName: (best && best.collection) || '',
      module: row.module || '',
      countA: resA.ok ? resA.count : null,
      countB: resB.ok ? resB.count : null,
      matched: 0,
      diffCount: 0,
      missingInTarget: 0,
      onlyInTarget: 0,
      total: Math.max(resA.count || 0, resB.count || 0),
      alignmentPct: 0,
      status: 'No OData',
      detail: (!resA.ok ? 'Source: ' + resA.detail : 'Target: ' + resB.detail),
      metaA: metaA,
      metaB: metaB,
      noOdata: true
    };
  }

  var pairs = findAllDifferentRowPairs(resA.rows || [], resB.rows || []);
  var diffCount = pairs.filter(function(p) { return p.rowA && p.rowB && p.fieldDiffs.length > 0; }).length;
  var missingInTarget = pairs.filter(function(p) { return p.rowA && !p.rowB; }).length;
  var onlyInTarget = pairs.filter(function(p) { return !p.rowA && p.rowB; }).length;
  var matched = Math.max(0, Math.min(
    resA.count - diffCount - missingInTarget,
    resB.count - diffCount - onlyInTarget
  ));
  var summary = summarizeEntityResult({
    matched: matched,
    diffCount: diffCount,
    missingInTarget: missingInTarget,
    onlyInTarget: onlyInTarget
  });
  var status = 'Match';
  var detail = '';
  if (diffCount > 0 || missingInTarget > 0 || onlyInTarget > 0) {
    if (matched === 0 && diffCount === 0 && missingInTarget > 0 && onlyInTarget === 0) {
      status = 'Only in Source';
      detail = missingInTarget + ' record(s) missing in target';
    } else if (matched === 0 && diffCount === 0 && onlyInTarget > 0 && missingInTarget === 0) {
      status = 'Only in Target';
      detail = onlyInTarget + ' record(s) only in target';
    } else {
      status = 'Diff';
      if (diffCount) detail += diffCount + ' value diff(s)';
      if (missingInTarget) detail += (detail ? ', ' : '') + missingInTarget + ' missing in target';
      if (onlyInTarget) detail += (detail ? ', ' : '') + onlyInTarget + ' only in target';
    }
  }

  return {
    idx: idx + 1,
    name: row.name,
    label: row.label || getBestEntityLabel(metaA, metaB) || row.name,
    aotName: (best && best.aotName) || row.aotName || row.name,
    dmfName: (best && best.dmfName) || row.dmfName || row.label || row.name,
    publicCollectionName: (best && best.collection) || '',
    module: row.module || '',
    countA: resA.count,
    countB: resB.count,
    matched: matched,
    diffCount: diffCount,
    missingInTarget: missingInTarget,
    onlyInTarget: onlyInTarget,
    total: summary.total,
    alignmentPct: summary.alignmentPct,
    status: status,
    detail: detail,
    metaA: metaA,
    metaB: metaB,
    noOdata: false,
    issueClass: summary.issueClass
  };
}

function aggregateModuleResults(detailRows) {
  var map = {};
  detailRows.forEach(function(row) {
    var key = row.module || '';
    if (!key) return;
    if (!map[key]) {
      map[key] = {
        module: key,
        total: 0,
        matched: 0,
        diffCount: 0,
        missingInTarget: 0,
        onlyInTarget: 0,
        entityCount: 0,
        noOdata: 0
      };
    }
    map[key].entityCount++;
    map[key].noOdata += row.noOdata ? 1 : 0;
    map[key].total += row.total || 0;
    map[key].matched += row.matched || 0;
    map[key].diffCount += row.diffCount || 0;
    map[key].missingInTarget += row.missingInTarget || 0;
    map[key].onlyInTarget += row.onlyInTarget || 0;
  });
  return Object.keys(map).sort().map(function(key) {
    var item = map[key];
    item.alignmentPct = item.total ? (item.matched / item.total) * 100 : 0;
    return item;
  });
}

function summarizeReport(detailRows) {
  var totals = {
    totalRecords: 0,
    matched: 0,
    diffCount: 0,
    missingInTarget: 0,
    onlyInTarget: 0,
    totalEntities: detailRows.length,
    identicalEntities: 0,
    differentEntities: 0,
    noOdataEntities: 0,
    totalModules: 0,
    alignmentPct: 0
  };
  detailRows.forEach(function(row) {
    totals.totalRecords += row.total || 0;
    totals.matched += row.matched || 0;
    totals.diffCount += row.diffCount || 0;
    totals.missingInTarget += row.missingInTarget || 0;
    totals.onlyInTarget += row.onlyInTarget || 0;
    if (row.noOdata) totals.noOdataEntities++;
    else if (row.diffCount === 0 && row.missingInTarget === 0 && row.onlyInTarget === 0) totals.identicalEntities++;
    else totals.differentEntities++;
  });
  totals.alignmentPct = totals.totalRecords ? (totals.matched / totals.totalRecords) * 100 : 0;
  return totals;
}

function buildReportHtml(detailRows) {
  var modules = aggregateModuleResults(detailRows);
  var totals = summarizeReport(detailRows);
  totals.totalModules = modules.length;
  var generatedAt = new Date().toISOString().slice(0, 19).replace('T', ' ');
  var sharedCompany = getCompany() || 'All legal entities';
  var diffRows = detailRows.filter(shouldDisplayDiffEntity);

  function pctClass(value) {
    if (value >= 95) return 'green';
    if (value >= 75) return 'amber';
    return 'red';
  }

  function badgeClass(status) {
    return { 'Match': 'bm', 'Diff': 'bd', 'Only in Source': 'bu', 'Only in Target': 'bs', 'No OData': 'bn' }[status] || 'bd';
  }

  function rowClass(status) {
    return { 'Match': 'data-row-M', 'Diff': 'data-row-D', 'Only in Source': 'data-row-S', 'Only in Target': 'data-row-T', 'No OData': 'data-row-N' }[status] || '';
  }

  function pct(n, d) { return d ? ((n / d) * 100).toFixed(1) : '0.0'; }
  function clamp(v) { return Math.min(100, Math.max(0, v || 0)); }

  // ── Module Summary rows ──────────────────────────────────────────
  var moduleRowsHtml = modules.map(function(item) {
    var matchPct = pct(item.matched, item.total);
    var cls = pctClass(item.alignmentPct);
    return '<tr class="data-row ' + rowClass('Match') + '" style="cursor:pointer" onclick="goToDetail(\'' + esc(item.module).replace(/'/g, "\\'") + '\')">' +
      '<td><span class="mod-tag">' + esc(item.module) + '</span></td>' +
      '<td class="val">' + formatNumber(item.total) + '</td>' +
      '<td class="green val">' + formatNumber(item.matched) + '</td>' +
      '<td class="amber val">' + formatNumber(item.diffCount) + '</td>' +
      '<td class="red val">' + formatNumber(item.missingInTarget) + '</td>' +
      '<td class="blue val">' + formatNumber(item.onlyInTarget) + '</td>' +
      '<td><div class="bar-wrap"><div class="bar bbar-' + cls + '" style="width:' + clamp(item.alignmentPct) + '%"></div></div><span style="font-size:10px;margin-left:4px">' + matchPct + '%</span></td>' +
    '</tr>';
  }).join('');

  // ── Data Entity Summary rows (ALL entities) ──────────────────────
  var allEntityRows = detailRows.slice().sort(function(a, b) {
    return (a.module || '').localeCompare(b.module || '') || (a.label || a.name).localeCompare(b.label || b.name);
  });
  var entitySummaryRowsHtml = allEntityRows.map(function(row) {
    var cls = pctClass(row.alignmentPct);
    var rc = rowClass(row.status);
    return '<tr class="data-row ' + rc + '" data-mod="' + esc(row.module) + '" data-status="' + esc(row.status) + '">' +
      '<td><span class="mod-tag">' + esc(row.module) + '</span></td>' +
      '<td>' + esc(row.label || row.name) + '</td>' +
      '<td class="val">' + (row.total ? formatNumber(row.total) : '—') + '</td>' +
      '<td class="green val">' + formatNumber(row.matched) + '</td>' +
      '<td class="amber val">' + formatNumber(row.diffCount) + '</td>' +
      '<td class="red val">' + formatNumber(row.missingInTarget) + '</td>' +
      '<td class="blue val">' + formatNumber(row.onlyInTarget) + '</td>' +
      '<td><div class="bar-wrap"><div class="bar bbar-' + cls + '" style="width:' + clamp(row.alignmentPct) + '%"></div></div><span style="font-size:10px;margin-left:4px">' + pct(row.matched, row.total) + '%</span></td>' +
    '</tr>';
  }).join('');

  // ── Full Detail rows (grouped by module, expandable) ─────────────
  var detailBodyHtml = modules.map(function(module) {
    var rows = detailRows.filter(function(r) { return r.module === module.module; });
    if (!rows.length) return '';
    var cls = pctClass(module.alignmentPct);
    var header = '<tr class="row-mod" data-mod-head="' + esc(module.module) + '">' +
      '<td colspan="8"><span class="chev">▼</span>' + esc(module.module) +
      '<span class="cnt-b">' + rows.length + ' entities</span>' +
      '<span class="pct-covered">Alignment ' + formatPercent(module.alignmentPct) + '</span></td></tr>';
    var entityRows = rows.map(function(row) {
      var rid = 'expand-' + row.idx;
      return '<tr class="row-det data-row ' + rowClass(row.status) + '" data-mod-row="' + esc(module.module) + '" data-status="' + esc(row.status) + '" onclick="toggleExpand(\'' + rid + '\')" style="cursor:pointer">' +
        '<td class="num-cell">' + row.idx + '</td>' +
        '<td><strong>' + esc(row.label || row.name) + '</strong><div class="sub">' + esc(row.aotName || row.name) + '</div></td>' +
        '<td><span class="badge ' + badgeClass(row.status) + '">' + esc(row.status) + '</span></td>' +
        '<td class="tc">' + (row.countA == null ? '—' : formatNumber(row.countA)) + '</td>' +
        '<td class="tc">' + (row.countB == null ? '—' : formatNumber(row.countB)) + '</td>' +
        '<td class="tc green">' + formatNumber(row.matched) + '</td>' +
        '<td class="tc amber">' + formatNumber(row.diffCount) + '</td>' +
        '<td class="tc red">' + formatNumber(row.missingInTarget) + '</td>' +
      '</tr>' +
      '<tr class="expand-row hidden-row" id="' + rid + '" data-mod-row="' + esc(module.module) + '">' +
        '<td colspan="8"><div class="expand-inner"><strong>Detail:</strong> ' + esc(row.detail || 'No differences detected') +
        (row.noOdata ? ' <span class="badge bn">No OData</span>' : '') + '</div></td>' +
      '</tr>';
    }).join('');
    return header + entityRows;
  }).join('');

  // ── Module filter options ─────────────────────────────────────────
  var modOptions = modules.map(function(m) {
    return '<option value="' + esc(m.module) + '">' + esc(m.module) + '</option>';
  }).join('');

  var lblA = esc(STATE.lblA);
  var lblB = esc(STATE.lblB);

  // ── CSS ───────────────────────────────────────────────────────────
  var css = [
    ':root{--primary:#1a237e;--primary2:#283593;--primary3:#3949ab;',
    '--green:#27ae60;--amber:#e67e22;--red:#e74c3c;--blue:#2980b9;--purple:#8e44ad;',
    '--green-bg:#eafaf1;--amber-bg:#fef9e7;--red-bg:#fdedec;--blue-bg:#ebf5fb;--purple-bg:#f5eef8;',
    '--border:#e0e4f0;--bg:#f0f3fa;--surface:#fff}',
    '*{box-sizing:border-box;margin:0;padding:0}',
    'body{font-family:"Segoe UI",system-ui,Arial,sans-serif;font-size:13px;background:var(--bg);color:#1a1a2e}',
    // Header
    '.app-header{background:linear-gradient(135deg,var(--primary) 0%,var(--primary2) 55%,var(--primary3) 100%);color:#fff;padding:18px 32px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;box-shadow:0 2px 10px rgba(0,0,0,.2)}',
    '.app-header h1{font-size:20px;font-weight:700;letter-spacing:.3px}',
    '.app-header p{font-size:11px;opacity:.75;margin-top:3px}',
    '.env-row{display:flex;gap:8px;align-items:center}',
    '.env-tag{padding:5px 16px;border-radius:20px;font-size:12px;font-weight:700;letter-spacing:.5px}',
    '.env-src{background:#f39c12;color:#fff}.env-tgt{background:#3498db;color:#fff}',
    // Tab bar
    '.tab-bar{display:flex;gap:0;background:var(--primary);padding:0 32px;border-bottom:3px solid var(--primary3);overflow-x:auto}',
    '.tab-btn-nav{padding:12px 20px;color:rgba(255,255,255,.68);font-size:12.5px;font-weight:600;border:none;background:none;cursor:pointer;white-space:nowrap;border-bottom:3px solid transparent;margin-bottom:-3px;transition:all .15s}',
    '.tab-btn-nav:hover{color:#fff;background:rgba(255,255,255,.08)}',
    '.tab-btn-nav.active{color:#fff;border-bottom-color:#ffd54f;background:rgba(255,255,255,.12)}',
    // Tab content
    '.tab-content{padding:24px 32px 48px}.tab-pane{display:none}.tab-pane.active{display:block}',
    '.section-title{font-size:14px;font-weight:700;color:var(--primary);margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid var(--border);display:flex;align-items:center;gap:8px}',
    '.cnt-badge{font-size:11px;background:var(--primary3);color:#fff;padding:2px 10px;border-radius:10px;font-weight:600}',
    // KPI cards
    '.kpi-row{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:24px}',
    '.kpi-card{flex:1;min-width:130px;background:var(--surface);border-radius:12px;padding:16px 20px;box-shadow:0 2px 8px rgba(0,0,0,.07);border-left:5px solid var(--primary3);transition:transform .15s}',
    '.kpi-card:hover{transform:translateY(-2px)}',
    '.kpi-card.kpi-green{border-left-color:var(--green);background:var(--green-bg)}',
    '.kpi-card.kpi-amber{border-left-color:var(--amber);background:var(--amber-bg)}',
    '.kpi-card.kpi-red{border-left-color:var(--red);background:var(--red-bg)}',
    '.kpi-card.kpi-blue{border-left-color:var(--blue);background:var(--blue-bg)}',
    '.kpi-card.kpi-purple{border-left-color:var(--purple);background:var(--purple-bg)}',
    '.kpi-icon{font-size:18px;opacity:.5;margin-bottom:6px}',
    '.kpi-val{font-size:26px;font-weight:800;color:var(--primary);line-height:1}',
    '.kpi-card.kpi-green .kpi-val{color:var(--green)}',
    '.kpi-card.kpi-amber .kpi-val{color:var(--amber)}',
    '.kpi-card.kpi-red .kpi-val{color:var(--red)}',
    '.kpi-card.kpi-blue .kpi-val{color:var(--blue)}',
    '.kpi-card.kpi-purple .kpi-val{color:var(--purple)}',
    '.kpi-lbl{font-size:10.5px;text-transform:uppercase;letter-spacing:.06em;color:#666;margin-top:5px;font-weight:600}',
    // Tables
    '.summary-tbl{width:100%;border-collapse:collapse;background:var(--surface);border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:16px}',
    '.summary-tbl thead tr{background:var(--primary);color:#fff}',
    '.summary-tbl thead th{padding:9px 12px;font-size:11px;font-weight:600;text-align:left}',
    '.summary-tbl td{padding:9px 12px;border-bottom:1px solid var(--border);font-size:12.5px}',
    '.summary-tbl tr:last-child td{border-bottom:none}',
    '.summary-tbl tr:hover td{background:#f7f8fd}',
    '.data-tbl{width:100%;border-collapse:collapse;background:var(--surface);border-radius:8px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,.08);margin-bottom:8px}',
    '.data-tbl thead tr{background:var(--primary);color:#fff}',
    '.data-tbl thead th{padding:10px 10px;font-size:11px;font-weight:600;text-align:left;white-space:nowrap;border-right:1px solid rgba(255,255,255,.1)}',
    '.data-tbl thead th:last-child{border-right:none}',
    '.data-tbl tbody td{padding:7px 10px;border-bottom:1px solid #f0f2fb;font-size:12px;vertical-align:middle}',
    '.data-tbl tbody tr.data-row{cursor:pointer;transition:background .1s}',
    '.data-tbl tbody tr.data-row-M:hover td{background:#d4f5e3}',
    '.data-tbl tbody tr.data-row-D:hover td{background:#fde8c8}',
    '.data-tbl tbody tr.data-row-S:hover td{background:#fcd9d9}',
    '.data-tbl tbody tr.data-row-T:hover td{background:#d0e8f8}',
    '.data-tbl tbody tr.data-row:hover td{background:#eef1fb}',
    // Expand rows
    '.expand-row td{padding:0!important;border-bottom:2px solid #c9cfe8}',
    '.expand-inner{padding:12px 20px 16px 20px;background:#f0f3fa;font-size:12px;color:#333}',
    // Misc
    '.val{text-align:right;font-weight:700}',
    '.tc{text-align:center}',
    '.green{color:var(--green)}.amber{color:var(--amber)}.red{color:var(--red)}.blue{color:var(--blue)}.purple{color:var(--purple)}',
    '.badge{display:inline-block;padding:3px 9px;border-radius:12px;font-size:10px;font-weight:700;white-space:nowrap;letter-spacing:.3px}',
    '.bm{background:#c8f5d8;color:#1a6e38}.bd{background:#fde8c8;color:#9a5000}.bu{background:#fcd9d9;color:#a02020}.bs{background:#d0e8f8;color:#0d3f6e}.bn{background:#ede9fe;color:#5b21b6}',
    '.mod-tag{font-size:10px;background:#e8eaf6;color:var(--primary);padding:2px 8px;border-radius:4px;white-space:nowrap;font-weight:600}',
    '.sub{font-size:10px;color:#667085;margin-top:2px}',
    '.bar-wrap{background:#e8ecf0;border-radius:4px;height:6px;overflow:hidden;display:inline-block;width:80px;vertical-align:middle}',
    '.bar{height:100%;border-radius:4px}',
    '.bbar-green{background:var(--green)}.bbar-amber{background:var(--amber)}.bbar-red{background:var(--red)}',
    // Toolbar
    '.toolbar-inner{display:flex;gap:8px;flex-wrap:wrap;align-items:center;padding:10px 0 14px 0;margin-bottom:16px}',
    '.toolbar-inner label{font-size:11.5px;font-weight:700;color:#555;white-space:nowrap}',
    '.sel-box{padding:6px 10px;border:1.5px solid #ccd;border-radius:6px;font-size:12px;background:#fafbff;color:#222;outline:none}',
    '.sel-box:focus{border-color:var(--primary)}',
    '.search-box{padding:6px 12px;border:1.5px solid #ccd;border-radius:6px;font-size:12px;background:#fafbff;min-width:220px;outline:none}',
    '.search-box:focus{border-color:var(--primary)}',
    '.filter-btn{padding:5px 12px;border:1.5px solid #ccd;border-radius:6px;font-size:11.5px;font-weight:600;cursor:pointer;background:#fff;color:#333;transition:all .15s}',
    '.filter-btn:hover,.filter-btn.active{background:var(--primary);color:#fff;border-color:var(--primary)}',
    '.tbl-wrap{overflow-x:auto;max-height:70vh;overflow-y:auto}',
    '.table-footer-count{font-size:11px;color:#666;padding:6px 2px 0 2px}',
    // Detail rows
    '.row-mod{background:linear-gradient(90deg,#1e2a6a,#2c3e80);color:#fff;font-weight:700;cursor:pointer}',
    '.row-mod td{padding:8px 10px;border-bottom:2px solid rgba(255,255,255,.12)}',
    '.row-det.hidden-row,.expand-row.hidden-row{display:none}',
    '.chev{display:inline-block;transition:transform .2s;margin-right:7px;font-size:9px;opacity:.8}',
    '.row-mod.collapsed .chev{transform:rotate(-90deg)}',
    '.cnt-b,.pct-covered{font-size:10px;background:rgba(255,255,255,.18);padding:2px 8px;border-radius:10px;margin-left:8px}',
    '.num-cell{width:36px;text-align:right;color:#aaa;font-size:11px}',
    '.hidden{display:none!important}',
    // Footer
    '.app-footer{text-align:center;font-size:11px;color:#999;padding:14px 32px;background:var(--surface);border-top:1px solid var(--border);margin-top:16px}',
    // Dual col layout
    '.dual-col{display:grid;grid-template-columns:1fr 1.4fr;gap:20px;margin-bottom:20px}@media(max-width:900px){.dual-col{grid-template-columns:1fr}}'
  ].join('');

  // ── KPI cards ─────────────────────────────────────────────────────
  var totRec = totals.totalRecords || 0;
  function kpiCard(extraClass, icon, val, lbl) {
    return '<div class="kpi-card' + (extraClass ? ' ' + extraClass : '') + '">' +
      '<div class="kpi-icon">' + icon + '</div>' +
      '<div class="kpi-val">' + val + '</div>' +
      '<div class="kpi-lbl">' + lbl + '</div>' +
    '</div>';
  }

  var kpiHtml =
    kpiCard('', '&#128203;', formatNumber(totRec), 'Total Records') +
    kpiCard('kpi-green', '&#9989;', formatNumber(totals.matched),
      'Match' + (totRec ? ' (' + pct(totals.matched, totRec) + '%)' : '')) +
    kpiCard('kpi-amber', '&#9889;', formatNumber(totals.diffCount),
      'Differences' + (totRec ? ' (' + pct(totals.diffCount, totRec) + '%)' : '')) +
    kpiCard('kpi-red', '&#128308;', formatNumber(totals.missingInTarget),
      lblA + ' Only' + (totRec ? ' (' + pct(totals.missingInTarget, totRec) + '%)' : '')) +
    kpiCard('kpi-blue', '&#128309;', formatNumber(totals.onlyInTarget),
      lblB + ' Only' + (totRec ? ' (' + pct(totals.onlyInTarget, totRec) + '%)' : ''));

  // ── HTML assembly ─────────────────────────────────────────────────
  return '<!DOCTYPE html>' +
  '<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
  '<title>D365 Config: ' + lblA + ' vs ' + lblB + '</title>' +
  '<style>' + css + '</style></head><body>' +

  // Header
  '<div class="app-header">' +
    '<div><h1>D365 Config: ' + lblA + ' vs ' + lblB + '</h1>' +
    '<p>Generated: ' + esc(generatedAt) + '</p></div>' +
    '<div class="env-row">' +
      '<span class="env-tag env-src">' + lblA + '</span>' +
      '<span style="font-size:18px;color:rgba(255,255,255,.6)">&#8594;</span>' +
      '<span class="env-tag env-tgt">' + lblB + '</span>' +
    '</div>' +
  '</div>' +

  // Tab bar
  '<div class="tab-bar">' +
    '<button class="tab-btn-nav active" data-tab="summary">&#128203; Overall Summary</button>' +
    '<button class="tab-btn-nav" data-tab="entity">&#128196; Data Entity Summary</button>' +
    '<button class="tab-btn-nav" data-tab="le">&#127970; Legal Entity Wise</button>' +
    '<button class="tab-btn-nav" data-tab="detail">&#128269; Full Detail</button>' +
  '</div>' +
  '<div class="tab-content">' +

  // ── Tab: Overall Summary ─────────────────────────────────────────
  '<div class="tab-pane active" id="tab-summary">' +
    '<div class="kpi-row">' + kpiHtml + '</div>' +
    '<div class="section-title">Module Summary <span class="cnt-badge">' + formatNumber(modules.length) + ' Modules</span></div>' +
    '<table class="summary-tbl"><thead><tr>' +
      '<th>Module</th><th>Total</th><th>Match</th><th>Diff</th>' +
      '<th>' + lblA + ' Only</th><th>' + lblB + ' Only</th><th>Match %</th>' +
    '</tr></thead><tbody>' + moduleRowsHtml + '</tbody></table>' +
    '<div style="margin-top:6px;font-size:11px;color:#888">Click a module row to view entity detail &rarr; Data Entity Summary tab</div>' +
  '</div>' +

  // ── Tab: Data Entity Summary ─────────────────────────────────────
  '<div class="tab-pane" id="tab-entity">' +
    '<div class="section-title">Data Entity Summary <span class="cnt-badge" id="ent-cnt-badge">' + formatNumber(detailRows.length) + ' Data Entities</span></div>' +
    '<div class="toolbar-inner">' +
      '<label>Module:</label><select class="sel-box" id="entModFilter"><option value="">All Modules</option>' + modOptions + '</select>' +
      '<label>Status:</label><select class="sel-box" id="entStatusFilter">' +
        '<option value="">All Statuses</option>' +
        '<option value="Match">Match</option>' +
        '<option value="Diff">Diff</option>' +
        '<option value="Only in Source">' + lblA + ' Only</option>' +
        '<option value="Only in Target">' + lblB + ' Only</option>' +
        '<option value="No OData">No OData</option>' +
      '</select>' +
      '<input type="search" class="search-box" id="entSearch" placeholder="Search entity...">' +
    '</div>' +
    '<div class="tbl-wrap"><table class="data-tbl" id="entTable"><thead><tr>' +
      '<th>Module</th><th>Data Entity</th><th>Total</th><th>Match</th><th>Diff</th>' +
      '<th>' + lblA + ' Only</th><th>' + lblB + ' Only</th><th>Match %</th>' +
    '</tr></thead><tbody id="entTbody">' + entitySummaryRowsHtml + '</tbody></table></div>' +
    '<div class="table-footer-count" id="entFooter">' + formatNumber(detailRows.length) + ' entities</div>' +
  '</div>' +

  // ── Tab: Legal Entity Wise ───────────────────────────────────────
  '<div class="tab-pane" id="tab-le">' +
    '<div class="section-title">Legal Entity Wise</div>' +
    '<table class="summary-tbl"><thead><tr><th>Environment</th><th>URL</th><th>Legal Entity Filter</th><th>Scope</th></tr></thead><tbody>' +
      '<tr><td><span class="env-tag env-src">' + lblA + '</span></td>' +
        '<td style="font-size:11px;word-break:break-all">' + esc(getEnvUrl('A') || '—') + '</td>' +
        '<td>' + esc(sharedCompany) + '</td>' +
        '<td>' + esc(sharedCompany === 'All legal entities' ? 'Cross-company' : 'Filtered') + '</td></tr>' +
      '<tr><td><span class="env-tag env-tgt">' + lblB + '</span></td>' +
        '<td style="font-size:11px;word-break:break-all">' + esc(getEnvUrl('B') || '—') + '</td>' +
        '<td>' + esc(sharedCompany) + '</td>' +
        '<td>' + esc(sharedCompany === 'All legal entities' ? 'Cross-company' : 'Filtered') + '</td></tr>' +
    '</tbody></table>' +
    '<div style="margin-top:16px;padding:14px 18px;background:var(--surface);border-radius:8px;border-left:4px solid var(--primary3);font-size:12px;color:#555;box-shadow:0 1px 4px rgba(0,0,0,.06)">' +
      '&#8505; Per-legal-entity record breakdowns are not available in this report. ' +
      'The comparison uses the legal entity filter set in the extension popup. ' +
      'To compare a specific legal entity, set the Company filter before generating the report.' +
    '</div>' +
  '</div>' +

  // ── Tab: Full Detail ─────────────────────────────────────────────
  '<div class="tab-pane" id="tab-detail">' +
    '<div class="section-title">Full Detail <span class="cnt-badge">' + formatNumber(detailRows.length) + ' Entities</span></div>' +
    '<div class="toolbar-inner">' +
      '<label>Module:</label><select class="sel-box" id="detModFilter"><option value="">All Modules</option>' + modOptions + '</select>' +
      '<label>Status:</label><select class="sel-box" id="detStatusFilter">' +
        '<option value="">All Statuses</option>' +
        '<option value="Match">Match</option>' +
        '<option value="Diff">Diff</option>' +
        '<option value="Only in Source">' + lblA + ' Only</option>' +
        '<option value="Only in Target">' + lblB + ' Only</option>' +
        '<option value="No OData">No OData</option>' +
      '</select>' +
      '<input type="search" class="search-box" id="detSearch" placeholder="Search entity...">' +
    '</div>' +
    '<div class="tbl-wrap"><table class="data-tbl"><thead><tr>' +
      '<th>#</th><th>Entity</th><th>Status</th>' +
      '<th class="tc">' + lblA + ' Rows</th><th class="tc">' + lblB + ' Rows</th>' +
      '<th class="tc">Matched</th><th class="tc">Diff</th><th class="tc">' + lblA + ' Only</th>' +
    '</tr></thead><tbody id="detTbody">' + (detailBodyHtml || '<tr><td colspan="8" style="text-align:center;color:#999;padding:24px">No entities found</td></tr>') + '</tbody></table></div>' +
    '<div class="table-footer-count" id="detFooter">' + formatNumber(detailRows.length) + ' entities | Click a row to expand detail</div>' +
  '</div>' +

  '</div>' + // tab-content

  // Footer
  '<div class="app-footer">D365 Configuration Comparison &nbsp;|&nbsp; ' + lblA + ' vs ' + lblB + ' &nbsp;|&nbsp; ' + esc(generatedAt) + '</div>' +

  // ── JavaScript ───────────────────────────────────────────────────
  '<script>(function(){' +
  // Tab switching
  'function switchTab(name){' +
    'document.querySelectorAll(".tab-pane").forEach(function(el){el.classList.toggle("active",el.id==="tab-"+name);});' +
    'document.querySelectorAll(".tab-btn-nav").forEach(function(btn){btn.classList.toggle("active",btn.dataset.tab===name);});' +
  '}' +
  'document.querySelectorAll(".tab-btn-nav").forEach(function(btn){btn.addEventListener("click",function(){switchTab(btn.dataset.tab);});});' +

  // goToDetail: module summary row click -> entity tab filtered
  'function goToDetail(modName){switchTab("entity");var sel=document.getElementById("entModFilter");if(sel){sel.value=modName;}filterEntityTable();}' +
  'window.goToDetail=goToDetail;' +

  // Toggle expand row in detail tab
  'function toggleExpand(id){var el=document.getElementById(id);if(el){el.classList.toggle("hidden-row");}}' +
  'window.toggleExpand=toggleExpand;' +

  // Entity table filter
  'function filterEntityTable(){' +
    'var search=(document.getElementById("entSearch").value||"").toLowerCase();' +
    'var mod=(document.getElementById("entModFilter").value||"");' +
    'var status=(document.getElementById("entStatusFilter").value||"");' +
    'var rows=document.querySelectorAll("#entTbody tr");' +
    'var visible=0;' +
    'rows.forEach(function(row){' +
      'var text=row.textContent.toLowerCase();' +
      'var ok=(!search||text.indexOf(search)!==-1)&&(!mod||row.dataset.mod===mod)&&(!status||row.dataset.status===status);' +
      'row.classList.toggle("hidden",!ok);' +
      'if(ok)visible++;' +
    '});' +
    'var badge=document.getElementById("ent-cnt-badge");if(badge)badge.textContent=visible+" Data Entities";' +
    'var footer=document.getElementById("entFooter");if(footer)footer.textContent=visible+" entities";' +
  '}' +
  'var es=document.getElementById("entSearch");if(es)es.addEventListener("input",filterEntityTable);' +
  'var em=document.getElementById("entModFilter");if(em)em.addEventListener("change",filterEntityTable);' +
  'var est=document.getElementById("entStatusFilter");if(est)est.addEventListener("change",filterEntityTable);' +

  // Detail table filter (module headers show/hide based on filter)
  'function filterDetailTable(){' +
    'var search=(document.getElementById("detSearch").value||"").toLowerCase();' +
    'var mod=(document.getElementById("detModFilter").value||"");' +
    'var status=(document.getElementById("detStatusFilter").value||"");' +
    'var visible=0;' +
    'document.querySelectorAll("#detTbody tr.row-mod").forEach(function(head){' +
      'var modName=head.dataset.modHead;' +
      'var show=!mod||modName===mod;' +
      'head.classList.toggle("hidden",!show);' +
    '});' +
    'document.querySelectorAll("#detTbody tr.row-det").forEach(function(row){' +
      'var text=row.textContent.toLowerCase();' +
      'var rmod=row.dataset.modRow||"";' +
      'var rst=row.dataset.status||"";' +
      'var ok=(!search||text.indexOf(search)!==-1)&&(!mod||rmod===mod)&&(!status||rst===status);' +
      'row.classList.toggle("hidden",!ok);' +
      'if(ok)visible++;' +
    '});' +
    'var footer=document.getElementById("detFooter");if(footer)footer.textContent=visible+" entities | Click a row to expand detail";' +
  '}' +
  'var ds=document.getElementById("detSearch");if(ds)ds.addEventListener("input",filterDetailTable);' +
  'var dm=document.getElementById("detModFilter");if(dm)dm.addEventListener("change",filterDetailTable);' +
  'var dst=document.getElementById("detStatusFilter");if(dst)dst.addEventListener("change",filterDetailTable);' +

  // Module group collapse/expand in detail tab
  'document.querySelectorAll(".row-mod").forEach(function(head){' +
    'head.addEventListener("click",function(e){' +
      'if(e.target.closest("select,input,button"))return;' +
      'var modName=head.dataset.modHead;' +
      'head.classList.toggle("collapsed");' +
      'var collapsed=head.classList.contains("collapsed");' +
      'document.querySelectorAll("[data-mod-row]").forEach(function(row){' +
        'if(row.dataset.modRow===modName){row.classList.toggle("hidden-row",collapsed);}' +
      '});' +
    '});' +
  '});' +

  '})();<\/script>' +
  '</body></html>';
}

function openGeneratedReport(html) {
  var blob = new Blob([html], { type: 'text/html' });
  var url = URL.createObjectURL(blob);
  if (typeof chrome !== 'undefined' && chrome.tabs) {
    chrome.tabs.create({ url: url });
  } else {
    window.open(url, '_blank');
  }
}

async function generateHtmlReport() {
  if (!STATE.allRows.length) {
    toast('⚠️ Load and compare entities first.');
    return;
  }
  var selectedModule = document.getElementById('modSel').value;
  if (!selectedModule) {
    toast('⚠️ Select a specific module group first. Full-report generation for All Module Groups is disabled.');
    return;
  }
  var urlA = getEnvUrl('A'), urlB = getEnvUrl('B');
  if (!urlA || !urlB) {
    return;
  }
  var btn = document.getElementById('btnReport');
  btn.disabled = true;
  btn.innerHTML = '<span class="spin"></span>Generating report...';
  try {
    var rows = STATE.allRows.filter(function(row) {
      return row.module === selectedModule;
    }).sort(function(a, b) {
      return a.module.localeCompare(b.module) || a.name.localeCompare(b.name);
    });
    if (!rows.length) {
      throw new Error('No entities found for the selected module.');
    }
    var detailRows = await mapLimit(rows, 4, async function(row, i) {
      showProgress('Generate Report', 'Comparing ' + row.module + ' / ' + row.name + '...', 5 + Math.round((i / Math.max(1, rows.length)) * 90));
      return compareEntityRecords(row, i, urlA, urlB);
    });
    var html = buildReportHtml(detailRows.filter(Boolean));
    openGeneratedReport(html);
    showProgress('Generate Report', 'Report generated successfully.', 100);
    hideProgress(1200);
    toast('✅ HTML report generated.');
  } catch (e) {
    hideProgress();
    toast('⚠️ Report generation failed: ' + (e && e.message ? e.message : String(e)));
  } finally {
    btn.innerHTML = '📄';
    updateReportButtonState();
  }
}

/* ── Retry-with-backoff for 429 Rate Limit responses ── */
async function fetchWithRetry(doFetch, label, maxRetries, baseDelay) {
  maxRetries = maxRetries || 3;
  baseDelay  = baseDelay  || 800;
  var delay = baseDelay;
  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    var result;
    try { result = await doFetch(); }
    catch(e) { throw e; } // non-HTTP errors (network down etc) — propagate immediately
    var is429 = result && result.ok === false && (result.status === 429 ||
                (result.detail && result.detail.indexOf('429') !== -1));
    if (!is429) return result;
    if (attempt === maxRetries) break;
    var waitMs = delay;
    if (result.retryAfter && !isNaN(Number(result.retryAfter))) {
      waitMs = Math.max(waitMs, Number(result.retryAfter) * 1000);
    }
    console.warn('[429] ' + (label || 'request') + ' rate-limited — retrying in ' + (waitMs / 1000).toFixed(1) + 's (attempt ' + (attempt + 1) + '/' + maxRetries + ')');
    await new Promise(function(r) { setTimeout(r, waitMs); });
    delay *= 2;
  }
  return { ok: false, status: 429, rows: [], count: 0,
    detail: 'HTTP 429 – Rate limited after ' + maxRetries + ' retries. Wait a moment and try again.' };
}

async function fetchCollectionRows(envUrl, slot, entityMeta) {
  if (!entityMeta || !entityMeta.collection) {
    return { ok: false, rows: [], count: 0, detail: 'No OData collection name', endpoint: '' };
  }
  var origin = new URL(envUrl).origin;
  var company = getCompany(slot);
  var companyFilter = company ? "&$filter=dataAreaId%20eq%20'" + encodeURIComponent(company) + "'" : '';
  var endpoint = origin + '/data/' + encodeURIComponent(entityMeta.collection) + '?$top=50&$count=true&cross-company=true' + companyFilter;

  return fetchWithRetry(async function() {
    var raw;
    if (IS_EXT) {
      var found = await findD365Tab(envUrl);
      if (!found.tab) return { ok: false, status: 0, rows: [], count: 0, detail: 'No open tab for ' + found.host+'. Open & log in first.', endpoint: endpoint };
      var res = await probeEndpoint(found.tab.id, endpoint);
      if (!res || !res.ok) {
        return { ok: false, status: res && res.status, rows: [], count: 0, detail: res ? (res.detail || ('HTTP ' + res.status)) : 'No response', endpoint: endpoint };
      }
      raw = res.data;
    } else {
      var token = getToken(slot);
      var proxyUrl = 'http://localhost:8888/proxy?url=' + encodeURIComponent(endpoint);
      if (token) proxyUrl += '&token=' + encodeURIComponent(token);
      var r = await fetch(proxyUrl);
      if (!r.ok) return { ok: false, status: r.status, retryAfter: r.headers.get('Retry-After'), detail: 'HTTP ' + r.status };
      raw = await r.json();
    }
    var rows = Array.isArray(raw) ? raw : (raw && raw.value ? raw.value : []);
    var count = raw && raw['@odata.count'] != null ? Number(raw['@odata.count']) : rows.length;
    if (!isFinite(count)) count = rows.length;
    return { ok: true, rows: rows, count: count, endpoint: endpoint };
  }, entityMeta.collection, 3, 800);
}

function findFirstDifferentRowPair(rowsA, rowsB) {
  var sortedA = rowsA.map(function(row) { return { raw: row, key: stableJsonValue(row) }; }).sort(function(a, b) { return a.key.localeCompare(b.key); });
  var sortedB = rowsB.map(function(row) { return { raw: row, key: stableJsonValue(row) }; }).sort(function(a, b) { return a.key.localeCompare(b.key); });
  var maxLen = Math.max(sortedA.length, sortedB.length);
  for (var i = 0; i < maxLen; i++) {
    var a = sortedA[i] || null;
    var b = sortedB[i] || null;
    if (!a || !b || a.key !== b.key) {
      return { rowA: a && a.raw, rowB: b && b.raw, index: i };
    }
  }
  return { rowA: sortedA[0] && sortedA[0].raw, rowB: sortedB[0] && sortedB[0].raw, index: 0 };
}

// Returns all differing row pairs, matched by business key (first non-odata field value).
// Records only in one side appear as { rowA, rowB:null } or { rowA:null, rowB }.
function findAllDifferentRowPairs(rowsA, rowsB) {

  /* ── Smart key detection ──────────────────────────────────────────────────
   * Priority order:
   *  1. Known D365 natural-key suffixes: Id, Code, Num, Key, Name, No, Ref
   *     (case-insensitive suffix match, prefer shorter/simpler field names)
   *  2. Fields named exactly: Id, Key, Code, Name, Number
   *  3. dataAreaId is EXCLUDED from the key (it's a partition, not identity)
   *  4. OData metadata fields (@odata.*) are always excluded
   *  5. If nothing qualifies → use ALL non-OData fields (full-row equality)
   * ─────────────────────────────────────────────────────────────────────── */
  var KEY_EXACT   = /^(id|key|code|name|number|num|no)$/i;
  var KEY_SUFFIX  = /(Id|Code|Num|Key|Name|No|Ref|Number)$/;

  function detectKeyFields(rows) {
    if (!rows || !rows.length) return null;

    // Gather all field names from the combined sample
    var fieldSet = {};
    rows.forEach(function(row) {
      Object.keys(row || {}).forEach(function(k) { fieldSet[k] = true; });
    });
    var allFields = Object.keys(fieldSet).filter(isComparableField).sort();

    if (!allFields.length) return null;

    // Score each field — higher = better key candidate
    function score(f) {
      if (KEY_EXACT.test(f)) return 100;
      if (KEY_SUFFIX.test(f)) {
        // Prefer shorter names (less compound) and names that appear early alphabetically
        return 50 + Math.max(0, 20 - f.length);
      }
      return 0;
    }

    var candidates = allFields
      .map(function(f) { return { f: f, s: score(f) }; })
      .filter(function(x) { return x.s > 0; })
      .sort(function(a, b) { return b.s - a.s || a.f.localeCompare(b.f); });

    if (!candidates.length) return null;

    // Verify uniqueness: the top candidate(s) must produce unique keys across the sample
    // Try the top candidate alone first, then combos of 2-3 if needed
    function tryFields(fields) {
      var seen = {};
      var ok = true;
      rows.forEach(function(row) {
        var key = fields.map(function(f) { return String(row[f] == null ? '' : row[f]); }).join('|');
        if (seen[key]) ok = false;
        seen[key] = true;
      });
      return ok;
    }

    // Try top-1
    if (tryFields([candidates[0].f])) return [candidates[0].f];

    // Try top-2 combo
    if (candidates.length >= 2 && tryFields([candidates[0].f, candidates[1].f]))
      return [candidates[0].f, candidates[1].f];

    // Try top-3 combo
    if (candidates.length >= 3 && tryFields([candidates[0].f, candidates[1].f, candidates[2].f]))
      return [candidates[0].f, candidates[1].f, candidates[2].f];

    // Fallback: use the best single candidate even if not perfectly unique
    return [candidates[0].f];
  }

  // Detect key from the combined pool so both sides agree on the same fields
  var combinedSample = (rowsA || []).concat(rowsB || []);
  var keyFields = detectKeyFields(combinedSample);

  function rowBusinessKey(row) {
    if (!row) return '';
    if (keyFields) {
      return keyFields.map(function(f) { return String(row[f] == null ? '' : row[f]); }).join('|');
    }
    // Ultimate fallback: hash all non-OData fields
    var keys = Object.keys(row).filter(isComparableField).sort();
    return keys.map(function(k) { return String(row[k]); }).join('|');
  }

  var mapA = {}, mapB = {};
  (rowsA || []).forEach(function(row) { var k = rowBusinessKey(row); mapA[k] = row; });
  (rowsB || []).forEach(function(row) { var k = rowBusinessKey(row); mapB[k] = row; });

  var allKeys = {};
  Object.keys(mapA).forEach(function(k) { allKeys[k] = true; });
  Object.keys(mapB).forEach(function(k) { allKeys[k] = true; });

  var pairs = [];
  Object.keys(allKeys).sort().forEach(function(k) {
    var a = mapA[k] || null;
    var b = mapB[k] || null;
    if (!a || !b) {
      pairs.push({ rowA: a, rowB: b, keyMatch: false, fieldDiffs: [], keyFields: keyFields });
    } else {
      var fields = Object.keys(Object.assign({}, a, b)).filter(isComparableField).sort();
      var diffs = fields.filter(function(f) {
        return stringifyFieldValue(a[f]) !== stringifyFieldValue(b[f]);
      });
      if (diffs.length > 0) {
        pairs.push({ rowA: a, rowB: b, keyMatch: true, fieldDiffs: diffs, keyFields: keyFields });
      }
    }
  });
  return pairs;
}

async function showEntityDiff(detailRow) {
  var panel = document.getElementById('entityDiffPanel');
  panel.style.display = 'block';
  panel.className = 'detail-panel';
  panel.innerHTML = '<div class="detail-head" style="color:#6b7280">⏳ Loading records…</div>';

  var metaA = detailRow.metaA || detailRow.metaB;
  var metaB = detailRow.metaB || detailRow.metaA;
  var resA = await fetchCollectionRows(getEnvUrl('A'), 'A', metaA || metaB);
  var resB = await fetchCollectionRows(getEnvUrl('B'), 'B', metaB || metaA);

  /* ── Load error ── */
  if (!resA.ok || !resB.ok) {
    panel.innerHTML =
      '<div class="detail-head">⚠ Cannot Load: ' + esc(detailRow.name) + '</div>' +
      '<div class="tbl-wrap"><table><tbody>' +
        '<tr><th style="width:140px;padding:8px 12px;border-bottom:1px solid rgba(0,0,0,.08);width:36px">#</th>' +
            '<td style="padding:8px 12px;color:' + (resA.ok ? '#166534' : '#991b1b') + '">' + esc(resA.ok ? 'OK' : resA.detail) + '</td></tr>' +
        '<tr><th style="padding:8px 12px">🟢 ' + esc(STATE.lblB) + '</th>' +
            '<td style="padding:8px 12px;color:' + (resB.ok ? '#166534' : '#991b1b') + '">' + esc(resB.ok ? 'OK' : resB.detail) + '</td></tr>' +
      '</tbody></table></div>';
    return;
  }

  var pairs       = findAllDifferentRowPairs(resA.rows || [], resB.rows || []);
  var detectedKeys = pairs.length && pairs[0].keyFields ? pairs[0].keyFields : null;
  var keyLabel    = detectedKeys ? detectedKeys.join(' + ') : 'full-row match';
  var missingA    = pairs.filter(function(p){ return p.rowA && !p.rowB; });
  var missingB    = pairs.filter(function(p){ return !p.rowA && p.rowB; });
  var valueDiffs  = pairs.filter(function(p){ return p.rowA && p.rowB && p.fieldDiffs.length > 0; });

  /* ── Record identity label from key fields ── */
  function recLabel(row) {
    if (!row) return '(absent)';
    var flds = detectedKeys
      ? detectedKeys
      : Object.keys(row).filter(function(k){ return !k.startsWith('@'); }).sort().slice(0,3);
    return flds.map(function(k){ return row[k] != null ? String(row[k]) : ''; }).filter(Boolean).join(' / ') || 'Record';
  }

  function recPreview(row) {
    if (!row) return '—';
    var priority = ['Name', 'Description', 'GroupName', 'PoolName', 'ItemName'];
    var used = {};
    (detectedKeys || []).forEach(function(key) { used[key] = true; });
    var parts = [];

    priority.forEach(function(field) {
      if (used[field] || !isComparableField(field)) return;
      var value = row[field];
      if (value == null || value === '') return;
      used[field] = true;
      parts.push(field + ': ' + String(value));
    });

    Object.keys(row).filter(function(field) {
      return isComparableField(field) && !used[field];
    }).sort().some(function(field) {
      var value = row[field];
      if (value == null || value === '') return false;
      parts.push(field + ': ' + String(value));
      return parts.length >= 2;
    });

    return parts.length ? parts.slice(0, 2).join(' | ') : 'No additional business fields';
  }

  function buildMissingTable(pairs, side) {
    var keyTitle = detectedKeys && detectedKeys.length ? detectedKeys.join(' + ') : 'Record key';
    return '<div style="margin-top:8px;font-size:11px;color:#6b7280">Key field: <strong>' + esc(keyTitle) + '</strong></div>' +
      '<div style="overflow:auto;margin-top:8px">' +
        '<table style="width:100%;border-collapse:collapse;font-size:11.5px;background:rgba(255,255,255,.6);border-radius:6px;overflow:hidden">' +
          '<thead><tr>' +
            '<th style="text-align:right;padding:6px 8px;border-bottom:1px solid rgba(0,0,0,.08);width:36px">#</th>' +
            '<th style="text-align:left;padding:6px 8px;border-bottom:1px solid rgba(0,0,0,.08);min-width:140px">Key</th>' +
            '<th style="text-align:left;padding:6px 8px;border-bottom:1px solid rgba(0,0,0,.08)">Preview</th>' +
          '</tr></thead>' +
          '<tbody>' + pairs.map(function(pair, index) {
            var row = side === 'A' ? pair.rowA : pair.rowB;
            return '<tr>' +
              '<td style="text-align:right;padding:6px 8px;border-bottom:1px solid rgba(0,0,0,.05);color:#6b7280">' + (index + 1) + '</td>' +
              '<td style="padding:6px 8px;border-bottom:1px solid rgba(0,0,0,.05);font-family:Consolas,monospace;font-weight:700">' + esc(recLabel(row)) + '</td>' +
              '<td style="padding:6px 8px;border-bottom:1px solid rgba(0,0,0,.05);color:#4b5563">' + esc(recPreview(row)) + '</td>' +
            '</tr>';
          }).join('') + '</tbody>' +
        '</table>' +
      '</div>';
  }

  /* ── Build tbody rows for one diff pair ── */
  function buildRows(p, diffOnly) {
    var diffSet = {};
    (p.fieldDiffs || []).forEach(function(f){ diffSet[f] = true; });

    var allFields = Object.keys(Object.assign({}, p.rowA || {}, p.rowB || {}))
      .filter(isComparableField)
      .sort(function(a, b){
        // diffs first, then alpha
        return (diffSet[a] ? 0 : 1) - (diffSet[b] ? 0 : 1) || a.localeCompare(b);
      });

    if (diffOnly) allFields = allFields.filter(function(f){ return diffSet[f]; });

    var hasSame   = !diffOnly && allFields.some(function(f){ return !diffSet[f]; });
    var shownSame = false;
    var html = '';

    allFields.forEach(function(f) {
      var isDiff = !!diffSet[f];
      // Insert a divider before the first unchanged row
      if (!isDiff && !shownSame && hasSame) {
        shownSame = true;
        html += '<tr class="divider-row"><td colspan="3">── Unchanged fields ──</td></tr>';
      }
      var vA = stringifyFieldValue(p.rowA ? p.rowA[f] : undefined);
      var vB = stringifyFieldValue(p.rowB ? p.rowB[f] : undefined);
      html +=
        '<tr class="' + (isDiff ? 'r-diff' : '') + '">' +
          '<td class="c-field">' + esc(f) + (isDiff ? '<span class="diff-pill">diff</span>' : '') + '</td>' +
          '<td class="c-src' + (isDiff ? ' changed' : '') + '">' + esc(vA) + '</td>' +
          '<td class="c-tgt' + (isDiff ? ' changed' : '') + '">' + esc(vB) + '</td>' +
        '</tr>';
    });
    return html || '<tr><td colspan="3" style="text-align:center;padding:20px;color:#9ca3af">No fields to show</td></tr>';
  }

  /* ── Render selected pair into the table ── */
  function renderPair(idx, diffOnly) {
    panel.querySelectorAll('.edv-rec').forEach(function(el, i){
      el.classList.toggle('active', i === idx);
    });
    var p = valueDiffs[idx];
    if (!p) return;
    var totalFields = Object.keys(Object.assign({}, p.rowA || {}, p.rowB || {}))
      .filter(isComparableField).length;
    document.getElementById('edv-panel-rec').textContent  = recLabel(p.rowA || p.rowB);
    document.getElementById('edv-panel-stat').textContent = p.fieldDiffs.length + ' of ' + totalFields + ' fields differ';
    document.getElementById('edv-tbody').innerHTML = buildRows(p, diffOnly);
  }

  /* ── No differences at all ── */
  if (!valueDiffs.length && !missingA.length && !missingB.length) {
    panel.innerHTML =
      '<div class="detail-head">✅ No Differences — ' + esc(detailRow.name) + '</div>' +
      '<div class="detail-sub">' +
        '🔵 ' + esc(STATE.lblA) + ': ' + resA.count + ' records &nbsp;|&nbsp; ' +
        '🟢 ' + esc(STATE.lblB) + ': ' + resB.count + ' records &nbsp;|&nbsp; ' +
        'Sampled top 50. No differences found.' +
      '</div>';
    return;
  }

  /* ── Build missing-records banners ── */
  var missingHtml = '';
  if (missingA.length || missingB.length) {
    missingHtml = '<div class="edv-missing">';
    if (missingA.length) {
      missingHtml +=
        '<div class="edv-missing-blk src">' +
          '<strong>✕ ' + missingA.length + ' record' + (missingA.length > 1 ? 's' : '') +
          ' exist in 🔵 ' + esc(STATE.lblA) + ' but are missing from 🟢 ' + esc(STATE.lblB) + '</strong>' +
          '<div style="margin-top:4px;font-size:11px;opacity:.9">These keys exist in the source environment only.</div>' +
          buildMissingTable(missingA, 'A') +
        '</div>';
    }
    if (missingB.length) {
      missingHtml +=
        '<div class="edv-missing-blk tgt">' +
          '<strong>✕ ' + missingB.length + ' record' + (missingB.length > 1 ? 's' : '') +
          ' exist in 🟢 ' + esc(STATE.lblB) + ' but are missing from 🔵 ' + esc(STATE.lblA) + '</strong>' +
          '<div style="margin-top:4px;font-size:11px;opacity:.9">These keys exist in the target environment only.</div>' +
          buildMissingTable(missingB, 'B') +
        '</div>';
    }
    missingHtml += '</div>';
  }

  /* ── Build sidebar nav items ── */
  var navHtml = valueDiffs.map(function(p, i){
    return '<div class="edv-rec' + (i === 0 ? ' active' : '') + '" data-idx="' + i + '">' +
      '<span class="edv-rec-key">' + esc(recLabel(p.rowA || p.rowB)) + '</span>' +
      '<span class="edv-rec-badge">' + p.fieldDiffs.length + ' diff' + (p.fieldDiffs.length > 1 ? 's' : '') + '</span>' +
    '</div>';
  }).join('');

  /* ── First pair data ── */
  var first = valueDiffs[0];
  var firstTotal = first
    ? Object.keys(Object.assign({}, first.rowA || {}, first.rowB || {})).filter(function(f){ return !f.startsWith('@'); }).length
    : 0;

  /* ── Assemble full HTML ── */
  panel.innerHTML =
    /* Header */
    '<div class="edv-head">' +
      '<div class="edv-head-title">' + esc(detailRow.label || detailRow.name) + '</div>' +
      '<div class="edv-counts">' +
        '<span class="edv-count-src">🔵 ' + esc(STATE.lblA) + ': ' + resA.count + ' records</span>' +
        '<span class="edv-count-tgt">🟢 ' + esc(STATE.lblB) + ': ' + resB.count + ' records</span>' +
      '</div>' +
      '<span class="edv-key">🔑 ' + esc(keyLabel) + '</span>' +
      (resA.count > 50 || resB.count > 50
        ? '<span style="font-size:10.5px;color:#b45309">⚠ top 50 sampled</span>' : '') +
    '</div>' +
    /* Body */
    '<div class="edv-body">' +
      /* Sidebar — only shown if there are value diffs */
      (valueDiffs.length > 0
        ? '<div class="edv-sidebar">' +
            '<div class="edv-sidebar-hdr">⚠ ' + valueDiffs.length + ' record' + (valueDiffs.length > 1 ? 's' : '') + ' differ</div>' +
            '<div class="edv-sidebar-list">' + navHtml + '</div>' +
          '</div>'
        : '') +
      /* Main comparison panel */
      (valueDiffs.length > 0
        ? '<div class="edv-panel">' +
            '<div class="edv-panel-bar">' +
              '<span class="edv-panel-rec" id="edv-panel-rec">' + esc(recLabel(first.rowA || first.rowB)) + '</span>' +
              '<span class="edv-panel-stat" id="edv-panel-stat">' + (first ? first.fieldDiffs.length : 0) + ' of ' + firstTotal + ' fields differ</span>' +
              '<label class="edv-difftoggle"><input type="checkbox" id="edv-diffsonly"/> Differences only</label>' +
            '</div>' +
            '<div class="edv-tbl-wrap">' +
              '<table class="edv-tbl">' +
                '<colgroup><col class="c-field"/><col class="c-val"/><col class="c-val"/></colgroup>' +
                '<thead><tr>' +
                  '<th class="h-field">Field</th>' +
                  '<th class="h-src">🔵 ' + esc(STATE.lblA) + ' &nbsp;<small style="font-weight:400;opacity:.7">(Source)</small></th>' +
                  '<th class="h-tgt">🟢 ' + esc(STATE.lblB) + ' &nbsp;<small style="font-weight:400;opacity:.7">(Target)</small></th>' +
                '</tr></thead>' +
                '<tbody id="edv-tbody">' + (first ? buildRows(first, false) : '') + '</tbody>' +
              '</table>' +
            '</div>' +
          '</div>'
        : '<div style="padding:16px;font-size:12.5px;color:#6b7280">No field-value differences — only missing records (see below).</div>') +
    '</div>' +
    /* Missing records */
    missingHtml +
  '</div>';

  /* ── Wire sidebar clicks ── */
  var diffOnly = false;
  panel.querySelectorAll('.edv-rec').forEach(function(el){
    el.addEventListener('click', function(){
      renderPair(parseInt(el.getAttribute('data-idx'), 10), diffOnly);
    });
  });

  /* ── Wire diff-only toggle ── */
  var chk = document.getElementById('edv-diffsonly');
  if (chk) {
    chk.addEventListener('change', function(){
      diffOnly = chk.checked;
      var active = panel.querySelector('.edv-rec.active');
      renderPair(active ? parseInt(active.getAttribute('data-idx'), 10) : 0, diffOnly);
    });
  }
}

/* ── Build OData endpoint ── */
function buildEndpoint(origin, path){
  return origin + '/data/' + path;
}

/* ── Render module detail table ── */
function renderModuleDetailTable(rows){
  var el=document.getElementById('modDetailTbody');

  // Update column headers to actual env names
  var thSrc=document.getElementById('thSrcOnly');
  var thTgt=document.getElementById('thTgtOnly');
  if(thSrc) thSrc.textContent=(STATE.lblA||'Source')+' Only';
  if(thTgt) thTgt.textContent=(STATE.lblB||'Target')+' Only';

  if(!rows||!rows.length){
    el.innerHTML='<tr><td colspan="10" style="text-align:center;padding:20px;color:#9ca3af">No matching entities for the current filters</td></tr>';
    return;
  }
  el.innerHTML=rows.map(function(r){
    var stCls={Match:'rm',Diff:'rd','Only in Source':'ru','Only in Target':'rs','No OData':''}[r.status]||'';
    var stBadge={Match:'<span class="badge bm">Match</span>',Diff:'<span class="badge bd">Diff</span>',
      'Only in Source':'<span class="badge bu">Only Source</span>','Only in Target':'<span class="badge bs">Only Target</span>',
      'No OData':'<span class="badge" style="background:#e5e7eb;color:#6b7280">No OData</span>'}[r.status]||r.status;
    var canDiff=r.status!=='No OData'&&(r.metaA||r.metaB);
    return'<tr class="'+stCls+(canDiff?' mod-row-clickable':'')+'"'+(canDiff?' style="cursor:pointer"':'')+' data-name="'+esc(r.name)+'">' +
      '<td style="width:26px;text-align:right;color:#aaa;font-size:11px">'+esc(r.idx)+'</td>'+
      '<td><strong>'+esc(r.aotName || r.name)+'</strong>'+(r.name&&r.aotName&&r.name!==r.aotName?'<br/><span style="font-size:10px;color:#888">'+esc(r.name)+'</span>':'')+'</td>'+
      '<td>'+esc(r.dmfName || r.label || r.name)+'</td>'+
      '<td>'+(r.publicCollectionName?'<span style="font-family:Consolas,monospace;color:#1d4ed8">'+esc(r.publicCollectionName)+'</span>':'<span style="color:#9ca3af">no entry</span>')+'</td>'+
      '<td style="text-align:center">'+esc(r.countA!=null?r.countA:'?')+'</td>'+
      '<td style="text-align:center">'+esc(r.countB!=null?r.countB:'?')+'</td>'+
      '<td style="text-align:center;color:#c0392b;font-weight:600">'+(r.srcOnly?r.srcOnly:'—')+'</td>'+
      '<td style="text-align:center;color:#2980b9;font-weight:600">'+(r.tgtOnly?r.tgtOnly:'—')+'</td>'+
      '<td>'+stBadge+'</td>'+
      '<td style="font-size:10.5px;color:#888;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+esc(r.detail||'')+'</td>'+
    '</tr>';
  }).join('');
}

function renderModuleDetailPlaceholder(message){
  var el=document.getElementById('modDetailTbody');
  if(!el)return;
  el.innerHTML='<tr><td colspan="10" style="text-align:center;padding:20px;color:#9ca3af">'+esc(message||'Loading module detail...')+'</td></tr>';
}

function clearModuleDetailResults(message){
  STATE.moduleDetailRows=[];
  STATE.visibleModuleDetailRows=[];
  document.getElementById('modDetailPanel').style.display='block';
  document.getElementById('modDetailSummary').innerHTML='';
  renderModuleDetailPlaceholder(message||'Loading module detail...');
}

function applyModuleDetailFilters(){
  var search=(document.getElementById('modDetailFilter').value||'').toLowerCase();
  var statusFilter=(document.getElementById('modDetailStatusFilter').value||'');
  var baseRows=STATE.moduleDetailRows.slice();
  STATE.visibleModuleDetailRows=baseRows.filter(function(r){
    var matchesStatus=!statusFilter||r.status===statusFilter;
    var matchesSearch=!search||
      r.name.toLowerCase().includes(search)||
      (r.label&&r.label.toLowerCase().includes(search))||
      r.status.toLowerCase().includes(search)||
      (r.detail&&r.detail.toLowerCase().includes(search));
    return matchesStatus&&matchesSearch;
  });
  renderModuleDetailTable(STATE.visibleModuleDetailRows);
}

/* ── Row click handler ── */
function handleModuleRowClick(tr){
  if(!tr)return;
  var name=tr.dataset.name;
  var row=STATE.visibleModuleDetailRows.find(function(r){return r.name===name;});
  if(!row)return;
  showEntityDiff(row).catch(function(e){
    toast('\u26A0\uFE0F Diff error: '+(e&&e.message?e.message:String(e)));
  });
}

/* ── Filter module detail ── */
window.filterModuleDetail=function(val){
  document.getElementById('modDetailFilter').value=val||'';
  applyModuleDetailFilters();
};

var _moduleFilterTimer;
function scheduleModuleCompareRefresh(delay){
  clearTimeout(_moduleFilterTimer);
  if(!STATE.allRows.length)return;
  _moduleFilterTimer=setTimeout(function(){
    loadModuleEntities().catch(function(e){
      hideModuleProgress();
      toast('\u26A0\uFE0F Compare failed: '+(e&&e.message?e.message:String(e)));
    });
  },delay||0);
}

/* ── mapLimit ── */
function mapLimit(items,limit,worker){
  return new Promise(function(resolve){
    var results=new Array(items.length);
    var idx=0,completed=0,active=0;
    var TIMEOUT_MS=3*60*1000;
    var dead=false;
    var timer=setTimeout(function(){
      dead=true;resolve(results);
      toast('\u23F1 Module compare timed out after 3 min. Partial results shown.');
    },TIMEOUT_MS);
    function next(){
      if(dead)return;
      while(active<limit&&idx<items.length){
        (function(i){
          active++;
          Promise.resolve().then(function(){return worker(items[i],i);}).then(function(r){
            if(!dead){results[i]=r;}
          }).catch(function(e){
            if(!dead){
              var item=items[i]||{};
              results[i]={
                name:item.name||'',
                label:item.label||item.name||'',
                module:item.module||'',
                aotName:item.aotName||item.name||'',
                dmfName:item.dmfName||item.label||item.name||'',
                publicCollectionName:item.publicCollectionName||'',
                idx:i+1,
                countA:null,
                countB:null,
                srcOnly:0,
                tgtOnly:0,
                status:'No OData',
                detail:e&&e.message?e.message:String(e),
                metaA:null,
                metaB:null
              };
            }
          }).then(function(){
            active--;completed++;
            if(!dead){showModuleProgress('Compare Module Details',completed+' / '+items.length+' entities',5+Math.round((completed/items.length)*90));}
            if(completed===items.length){clearTimeout(timer);resolve(results);}
            else next();
          });
        })(idx++);
      }
    }
    next();
  });
}

/* ── Load Module Entities ── */
async function loadModuleEntities(){
  var mod=document.getElementById('modSel').value;
  var rows=STATE.allRows.filter(function(r){
    return !mod || r.module===mod;
  });
  if(!rows.length){toast('\u26A0\uFE0F No entities for current filter.');return;}
  var urlA=getEnvUrl('A'),urlB=getEnvUrl('B');
  if(!urlA||!urlB){toast('\u26A0\uFE0F Select both environments first.');return;}
  clearModuleDetailResults('Loading '+(mod||'selected module')+'...');
  setSt('stA','loading',getEnvLabel('A')+' — comparing module data...');
  setSt('stB','loading',getEnvLabel('B')+' — comparing module data...');
  showModuleProgress('Compare Module Details','Starting\u2026',5);
  STATE.activeModule=mod;

  var detailRows=await mapLimit(rows,4,async function(row,i){
    var metaA=STATE.entityMapA[row.name]||null;
    var metaB=STATE.entityMapB[row.name]||null;
    var best=metaA||metaB;
    if(!best||!best.collection){
      return{name:row.name,label:(best&&best.label)||row.name,module:row.module,aotName:(best&&best.aotName)||row.aotName||row.name,dmfName:(best&&best.dmfName)||row.dmfName||row.label||row.name,publicCollectionName:(best&&best.collection)||row.publicCollectionName||'',
        idx:i+1,countA:null,countB:null,
        srcOnly:0,tgtOnly:0,
        status:'No OData',detail:'No OData collection',metaA:metaA,metaB:metaB};
    }
    if(i<4)await new Promise(function(r){setTimeout(r,i*50);});
    var resA=await fetchCollectionRows(urlA,'A',metaA||metaB);
    var resB=await fetchCollectionRows(urlB,'B',metaB||metaA);
    var status='Match',detail='';
    if(!resA.ok&&!resB.ok){status='No OData';detail=resA.detail||resB.detail;}
    else if(!resA.ok){status='No OData';detail='Src: '+resA.detail;}
    else if(!resB.ok){status='No OData';detail='Tgt: '+resB.detail;}
    else{
      var pairs=findAllDifferentRowPairs(resA.rows||[],resB.rows||[]);
      var vd=pairs.filter(function(p){return p.rowA&&p.rowB&&p.fieldDiffs.length>0;});
      var ma=pairs.filter(function(p){return p.rowA&&!p.rowB;});
      var mb=pairs.filter(function(p){return!p.rowA&&p.rowB;});
      if(vd.length>0){status='Diff';detail=vd.length+' record(s) differ';}
      else if(ma.length && !mb.length){
        status='Only in Source';
        detail=ma.length+' only in Source';
      } else if(!ma.length && mb.length){
        status='Only in Target';
        detail=mb.length+' only in Target';
      } else if(ma.length||mb.length){
        status='Diff';
        detail=(ma.length?ma.length+' only in Source':'')+(mb.length?(ma.length?', ':'')+mb.length+' only in Target':'');
      } else { status='Match'; }
    }
    return{name:row.name,label:(best&&best.label)||row.name,module:row.module,aotName:(best&&best.aotName)||row.aotName||row.name,dmfName:(best&&best.dmfName)||row.dmfName||row.label||row.name,publicCollectionName:(best&&best.collection)||row.publicCollectionName||'',
      idx:i+1,countA:resA.ok?resA.count:null,countB:resB.ok?resB.count:null,
      srcOnly:ma.length,tgtOnly:mb.length,
      status:status,detail:detail,metaA:metaA,metaB:metaB};
  });

  STATE.moduleDetailRows=detailRows.filter(Boolean);
  STATE.visibleModuleDetailRows=STATE.moduleDetailRows.slice();
  document.getElementById('modDetailPanel').style.display='block';

  var total=STATE.moduleDetailRows.length;
  var match=STATE.moduleDetailRows.filter(function(r){return r.status==='Match';}).length;
  var diff=STATE.moduleDetailRows.filter(function(r){return r.status==='Diff';}).length;
  var onlyA=STATE.moduleDetailRows.filter(function(r){return r.status==='Only in Source';}).length;
  var onlyB=STATE.moduleDetailRows.filter(function(r){return r.status==='Only in Target';}).length;
  var noOdata=STATE.moduleDetailRows.filter(function(r){return r.status==='No OData';}).length;
  var displayed=STATE.visibleModuleDetailRows.length;
  document.getElementById('modDetailSummary').innerHTML=
    '<span class="badge bd">\u25B3 Diff: '+diff+'</span> '+
    '<span class="badge bu">\u2212 Only Source: '+onlyA+'</span> '+
    '<span class="badge bs">+ Only Target: '+onlyB+'</span> '+
    '<span style="font-size:11px;color:#888;margin-left:6px">Compared '+total+' entities from Source and Target</span>'+
    '<span style="font-size:11px;color:#888;margin-left:6px">Showing '+displayed+' entities</span>'+
    ((match||noOdata)?'<span style="font-size:11px;color:#9ca3af;margin-left:6px">Includes '+match+' match and '+noOdata+' no OData</span>':'');

  var sf=document.getElementById('modDetailFilter');
  if(sf)sf.value='';
  var ssf=document.getElementById('modDetailStatusFilter');
  if(ssf)ssf.value='';
  applyModuleDetailFilters();
  showModuleProgress('Compare Module Details','Done!',100);
  hideModuleProgress(1000);
  setSt('stA','ok',getEnvLabel('A')+' — compared '+total+' entities ✓');
  setSt('stB','ok',getEnvLabel('B')+' — compared '+total+' entities ✓');
  toast('\u2705 Module compare complete.');
}

/* ── Diagnose ── */
async function runDiagnose(){
  var panel=document.getElementById('diagPanel');
  panel.style.display='block';
  panel.textContent='\uD83D\uDD0D Running diagnostics\u2026\n';
  var urlA=getEnvUrl('A'),urlB=getEnvUrl('B');
  async function probe(url,label){
    if(!url){panel.textContent+=label+': (no URL)\n';return;}
    panel.textContent+=label+': checking '+url+'\u2026\n';
    try{
      var origin=new URL(url).origin;
      if(IS_EXT){
        var found=await findD365Tab(url);
        if(!found.tab){panel.textContent+=label+': \u274C No open tab for '+found.host+'. Open & log in first.\n';return;}
        var candidates=getCandidateEndpoints(origin).map(function(u){
          return u.replace('$top=10000','$top=1').replace(/&\$select=[^&]*/,'');
        });
        for(var i=0;i<candidates.length;i++){
          var ep=candidates[i];
          var res=await probeEndpoint(found.tab.id,ep);
          panel.textContent+=label+': '+ep+' => '+(res&&res.ok?'\u2705 OK':'\u274C '+(res?('HTTP '+res.status+' '+(res.detail||'')).slice(0,220):'null'))+'\n';
          if(res&&res.ok) break;
        }
      }
    }catch(e){panel.textContent+=label+': \u274C '+e.message+'\n';}
  }
  await probe(urlA,'Source');
  await probe(urlB,'Target');
  panel.textContent+='Done.\n';
}

/* ═══════════════════════════════════════════
   INIT
═══════════════════════════════════════════ */
(function initApp(){
  // Pull profiles & picks from chrome.storage first, then render
  syncProfilesFromStorage().then(function(){
    renderProfileList();
    refreshPickers();
    restore();
  });

  document.getElementById('btnAddProfile').addEventListener('click',function(e){
    e.preventDefault();
    addOrUpdateProfile().catch(function(err){
      toast('⚠️ Could not save profile: '+(err && err.message ? err.message : String(err)));
    });
  });
  document.getElementById('pfList').addEventListener('click',function(e){
    var editBtn=e.target.closest('.pf-edit-btn');
    var delBtn=e.target.closest('.pf-del-btn');
    if(editBtn)editProfile(editBtn.dataset.id);
    else if(delBtn)deleteProfile(delBtn.dataset.id).catch(function(err){
      toast('⚠️ Could not delete profile: '+(err && err.message ? err.message : String(err)));
    });
  });
  document.getElementById('pickerA').addEventListener('change',function(){loadSlot('A');persist();});
  document.getElementById('pickerB').addEventListener('change',function(){loadSlot('B');persist();});

  ['modalCloseBtn','modalGotItBtn'].forEach(function(id){
    var el=document.getElementById(id);if(el)el.addEventListener('click',hideTokenModal);
  });
  document.getElementById('tokenModal').addEventListener('click',function(e){if(e.target===this)hideTokenModal();});
  document.getElementById('tokenCmd').addEventListener('click',copyTokenCmd);
  document.getElementById('btnToggleSetup').addEventListener('click',function(){
    var shouldCollapse=this.getAttribute('aria-expanded')==='true';
    setSetupCardCollapsed(shouldCollapse);
  });

  document.getElementById('btnVal').addEventListener('click',validateAccess);
  document.getElementById('btnLoad').addEventListener('click',function(){
    loadEntities().catch(function(e){toast('\u26A0\uFE0F Load failed: '+(e&&e.message?e.message:String(e)));});
  });
  document.getElementById('btnDiag').addEventListener('click',runDiagnose);

  document.getElementById('modSel').addEventListener('change',function(){
    document.getElementById('entityDiffPanel').style.display='none';
    clearModuleDetailResults('Loading '+(this.value||'selected module')+'...');
    updateReportButtonState();
    scheduleModuleCompareRefresh(0);
  });

  document.getElementById('btnReport').addEventListener('click',function(){
    generateHtmlReport().catch(function(e){
      hideProgress();
      toast('\u26A0\uFE0F Report failed: '+(e&&e.message?e.message:String(e)));
    });
  });

  document.getElementById('company').addEventListener('change',function(){onCompanySelectChange();persist();});
  document.getElementById('companyCustom').addEventListener('input',persist);

  document.getElementById('modDetailTbody').addEventListener('click',function(e){
    var tr=e.target.closest('tr.mod-row-clickable');
    if(tr)handleModuleRowClick(tr);
  });

  var mds=document.getElementById('modDetailFilter');
  if(mds)mds.addEventListener('input',function(){filterModuleDetail(this.value);});
  var mdss=document.getElementById('modDetailStatusFilter');
  if(mdss)mdss.addEventListener('change',applyModuleDetailFilters);

  updateReportButtonState();

  /* Full page */
  document.getElementById('btnFullPage').addEventListener('click',function(){
    if(typeof chrome!=='undefined'&&chrome.tabs){
      chrome.tabs.create({url:chrome.runtime.getURL('popup.html?fullpage=1')});
    } else {
      window.open(window.location.href.split('?')[0]+'?fullpage=1','_blank');
    }
  });

  if(window.location.search.includes('fullpage=1')){
    document.body.classList.add('fullpage');
  }

  // restore() is now called inside syncProfilesFromStorage().then() at top of initApp
})(); // end initApp

// ── Token & UI helpers ──
function getToken(slot){
  return '';
}

function showTokenHelp(){
  var m = document.getElementById('tokenModal');
  if(m) m.classList.add('show');
}

function hideTokenModal(){
  var m = document.getElementById('tokenModal');
  if(m) m.classList.remove('show');
}

function copyTokenCmd(){
  var el = document.getElementById('tokenCmd');
  if(!el) return;
  var text = el.textContent;
  if(navigator.clipboard && navigator.clipboard.writeText){
    navigator.clipboard.writeText(text).then(function(){toast('📋 Copied!');}).catch(function(){});
  } else {
    var ta = document.createElement('textarea');
    ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); toast('📋 Copied!'); } catch(e){}
    document.body.removeChild(ta);
  }
}

function filterModuleDetail(q){
  var tbody = document.getElementById('modDetailTbody');
  if(!tbody) return;
  var term = String(q||'').toLowerCase();
  Array.from(tbody.querySelectorAll('tr')).forEach(function(tr){
    if(!term){ tr.style.display=''; return; }
    var txt = tr.textContent.toLowerCase();
    tr.style.display = txt.includes(term) ? '' : 'none';
  });
}

})(); // end outer IIFE
