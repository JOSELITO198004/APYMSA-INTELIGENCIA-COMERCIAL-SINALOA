/* APYMSA: diagnostico de dependencias sin bloquear el renderizado. */
window.addEventListener('load', () => {
  const missing = [];
  if (!window.Papa) missing.push('PapaParse');
  if (!window.XLSX) missing.push('XLSX');
  if (!window.pako && !('DecompressionStream' in window)) missing.push('Pako');
  if (!window.Chart) missing.push('Chart.js');
  if (!window.ChartDataLabels) missing.push('ChartDataLabels');
  if (missing.length) {
    console.warn('Dependencias no disponibles:', missing.join(', '));
    const panel = document.getElementById('dependency-warning');
    const detail = document.getElementById('dependency-warning-detail');
    if (panel && detail) {
      detail.textContent = 'No respondieron: ' + missing.join(', ') + '. La interfaz basica sigue disponible. Recarga o cambia de red para usar todas las funciones.';
      panel.hidden = false;
    }
  }
});

function toggleFullScreen() {
        if (!document.fullscreenElement) { document.documentElement.requestFullscreen().catch(err => console.log(err.message)); } 
        else { if (document.exitFullscreen) document.exitFullscreen(); }
    }
    document.addEventListener('fullscreenchange', () => {
        const textSpan = document.getElementById('fs-text');
        const btn = document.getElementById('btnFullScreen');
        if (document.fullscreenElement) {
            textSpan.innerText = 'Salir Pantalla Completa';
            btn.classList.add('border-neon-cyan', 'text-neon-cyan');
            btn.classList.remove('border-slate-500', 'text-white');
        } else {
            textSpan.innerText = 'Pantalla Completa';
            btn.classList.remove('border-neon-cyan', 'text-neon-cyan');
            btn.classList.add('border-slate-500', 'text-white');
        }
    });

    const fileInput = document.getElementById('dataFile');
    const dropZone = document.getElementById('drop-zone');
    const catalogContainer = document.getElementById('catalog-container');
    const btnPrint = document.getElementById('btnPrint');
    const btnBack = document.getElementById('btnBack');
    const step1 = document.getElementById('step-1');
    const step2 = document.getElementById('step-2');
    const stepStrategy = document.getElementById('step-strategy');
    
    const regionalView = document.getElementById('regional-view');
    const advisorView = document.getElementById('advisor-view');
    const clientView = document.getElementById('client-view');
    const clientList = document.getElementById('client-list');
    const loaderIcon = document.getElementById('load-icon');
    const loaderSpinner = document.getElementById('loader-spinner');

    let companyTree = new Map();
    let allClientsFlat = new Map(); 
    let clientsById = new Map(); 
    let dataQuality = {
        ventasSinCliente: 0,
        ventasAmbiguas: 0,
        codigosSinMaestro: new Set(),
        asesoresSinJerarquia: new Set(),
        disciplinaSinCliente: 0,
        clientesDuplicadosCV: []
    };
    let selectedClient = null;
    let currentSelectedTerritorial = null;
    let currentSelectedRegional = null;
    let currentSelectedAdvisor = null; 
    let globalProducts = new Map();
    
    let globalSalesMn = {}; 
    let globalSalesPesos = {}; 
    let globalPiezas = {}; 
    let globalClientsPerSku = {};
    let globalPromociones = {}; 
    let globalPromoText = {}; 
    let globalComentarioPromo = {}; 
    let globalLowestPrice = {}; 
    let globalSkuStats = {}; 

    let cartItems = new Set();
    let manualSkus = []; 
    let manualSkusQty = {}; 
    let chartBrandsInstance = null;
    let chartTerritorialIntel=null,chartTerritorialDiscipline=null,chartRegionalIntel=null,chartAdvisorIntel=null,chartRegionalDiscipline=null,chartAdvisorDiscipline=null,chartDetailBrands=null;
    let disciplineReturnContext=null;
    let disciplineOpportunitySkus=[];
    let currentDisciplineClients=[];
    let noVisitRowsCache=[];
    let chartTrendInstance = null;
    const segmentChartInstances={};

    let currentStrategyLabel = "GENERAL";
    let appCurrentState = "LOAD";

    const NOW = new Date();
    const CURRENT_YEAR = NOW.getFullYear();
    const CURRENT_MONTH = NOW.getMonth();
    const MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

    function clearCacheAndReload() { loaderSpinner.classList.remove('hidden'); location.reload(); }

    function calcularDiasTranscurridos() {
        let diasAcumulados = 0;
        let hoy = new Date();
        for (let d = 1; d < hoy.getDate(); d++) {
            let fechaIterada = new Date(hoy.getFullYear(), hoy.getMonth(), d);
            let diaSemana = fechaIterada.getDay();
            if (diaSemana >= 1 && diaSemana <= 5) diasAcumulados += 1;
            else if (diaSemana === 6) diasAcumulados += 0.5;
        }
        return Math.max(diasAcumulados, 0);
    }

    function obtenerDiasEfectivosDelMes(anio, mes) {
        let dias = 0;
        let ultimoDia = new Date(anio, mes + 1, 0).getDate();
        for (let d = 1; d <= ultimoDia; d++) {
            let diaSemana = new Date(anio, mes, d).getDay();
            if (diaSemana >= 1 && diaSemana <= 5) dias += 1;
            else if (diaSemana === 6) dias += 0.5;
        }
        return dias;
    }

    const DIAS_EFECTIVOS_MES = obtenerDiasEfectivosDelMes(CURRENT_YEAR, CURRENT_MONTH); 
    const DIAS_TRANSCURRIDOS = calcularDiasTranscurridos();
    const marcasPropias = ["STAR", "BRESSER", "BRESSER LKW", "DYNAMIC", "TECNOFUEL", "VALUE", "ACOSA", "REWARD", "WEISCHLER", "CARFAN", "VERZE", "JULS CARMAN", "REIDEN"];
    
    function normalizeBrandName(value) { return String(value || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); }
    
    dropZone.addEventListener('click', e => { if (e.target.id !== 'btnSelectFile' && e.target.tagName !== 'BUTTON') fileInput.click(); });
    document.getElementById('btnSelectFile').addEventListener('click', e => { e.stopPropagation(); fileInput.click(); });
    fileInput.addEventListener('change', e => { const file = e.target.files[0]; if (file) handleFile(file); });
    document.getElementById('btnLoadCloud').addEventListener('click', async e => {
        e.stopPropagation();
        await tryLoadGzipJson();
    });

    function pickCloudArray(source, aliases) {
        if (!source || typeof source !== 'object') return [];
        for (const key of aliases) if (Array.isArray(source[key])) return source[key];
        const normalizedAliases = aliases.map(normalizeKey);
        for (const [key, value] of Object.entries(source)) {
            if (Array.isArray(value) && normalizedAliases.includes(normalizeKey(key))) return value;
        }
        return [];
    }

    function processCloudPayload(payload) {
        const root = payload && payload.data && typeof payload.data === 'object' ? payload.data : payload;
        const dataProductos = pickCloudArray(root, ['Productos', 'productos', 'dataProductos']);
        const dataClientes = pickCloudArray(root, ['Clientes', 'clientes', 'dataClientes']);
        const dataJerarquia = pickCloudArray(root, ['Jerarquia', 'Jerarquía', 'jerarquia', 'dataJerarquia']);
        const dataVentas = pickCloudArray(root, ['Ventas','Venta','ventas','venta','dataVentas']);
        const dataDisciplina = pickCloudArray(root, ['Disciplina','disciplina','Disciplina Comercial','dataDisciplina']);
        const dataAccesos = pickCloudArray(root, ['Accesos','accesos','Usuarios','usuarios']);
        const missingCollections = [];
        if(!dataProductos.length) missingCollections.push('Productos');
        if(!dataClientes.length) missingCollections.push('Clientes');
        if(!dataJerarquia.length) missingCollections.push('Jerarquia');
        if(!dataVentas.length) missingCollections.push('Ventas');
        if(!dataDisciplina.length) missingCollections.push('Disciplina');
        if(!dataAccesos.length) missingCollections.push('Accesos');
        if(missingCollections.length) throw new Error('Colecciones faltantes o vacías: ' + missingCollections.join(', ') + '. Claves encontradas: ' + Object.keys(root || {}).join(', '));
        const primeraJerarquia=normalizeRow(dataJerarquia[0]);
        if(!('gerente_territorial' in primeraJerarquia)||!String(primeraJerarquia.gerente_territorial||'').trim()){
            throw new Error('La hoja Jerarquía requiere la columna gerente_territorial con valores.');
        }
        processData(dataProductos, dataClientes, dataJerarquia, dataVentas, dataDisciplina, dataAccesos);
    }

    async function decodeGzipBytes(bytes){const u=bytes instanceof Uint8Array?bytes:new Uint8Array(bytes);if(u.length<2||u[0]!==31||u[1]!==139)throw new Error('El archivo no tiene una firma GZIP válida.');if('DecompressionStream' in window){const stream=new Blob([u]).stream().pipeThrough(new DecompressionStream('gzip'));return await new Response(stream).text()}if(!window.pako)throw new Error('Pako no disponible');return pako.ungzip(u,{to:'string'})}
    async function fetchWithTimeout(url,timeout=15000){const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),timeout);try{const r=await fetch(url,{cache:'no-store',signal:controller.signal});if(!r.ok)throw new Error(`HTTP ${r.status}`);return new Uint8Array(await r.arrayBuffer())}finally{clearTimeout(timer)}}
    async function tryLoadGzipJson(){
        const cloudUrl = './norte.json.gz';
        const button = document.getElementById('btnLoadCloud');
        loaderIcon.classList.add('hidden');
        loaderSpinner.classList.remove('hidden');
        button.disabled = true;
        document.getElementById('loader-text').innerText = 'Descargando norte.json.gz desde GitHub...';
        try {
            const bytes = await fetchWithTimeout(cloudUrl, 30000);
            if (bytes.length < 2 || bytes[0] !== 31 || bytes[1] !== 139) {
                throw new Error(`La descarga no es un archivo GZIP válido. Firma recibida: ${bytes[0] ?? 'sin dato'}, ${bytes[1] ?? 'sin dato'}.`);
            }
            document.getElementById('loader-text').innerText = 'Descomprimiendo y validando la base...';
            const jsonText = await decodeGzipBytes(bytes);
            if (!jsonText || !jsonText.trim()) throw new Error('El archivo GZIP se descargó, pero su contenido está vacío.');
            let payload;
            try {
                payload = JSON.parse(jsonText.replace(/^\uFEFF/, '').trim());
            } catch (parseError) {
                throw new Error('El GZIP se descargó correctamente, pero el JSON interno no es válido: ' + parseError.message);
            }
            processCloudPayload(payload);
        } catch (err) {
            console.error('Error cargando desde GitHub:', err);
            const message = err && err.name === 'AbortError'
                ? 'La descarga excedió el tiempo máximo de 30 segundos.'
                : (err && err.message ? err.message : String(err || 'Error desconocido'));
            showAlert('ERROR DE CARGA EN LA NUBE: ' + message);
            showLoadError(new Error(message + '\n\nFuente única: ' + cloudUrl), 'Carga desde GitHub');
            resetCarga();
        } finally {
            button.disabled = false;
        }
    }
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt => { dropZone.addEventListener(evt, e => { e.preventDefault(); e.stopPropagation(); }); });
    dropZone.addEventListener('dragover', () => { dropZone.classList.add('bg-[#00f3ff]/10', 'border-[#00f3ff]'); });
    dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('bg-[#00f3ff]/10', 'border-[#00f3ff]'); });
    dropZone.addEventListener('drop', e => { dropZone.classList.remove('bg-[#00f3ff]/10', 'border-[#00f3ff]'); if(e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]); });
    document.getElementById('clientSearch').addEventListener('input', () => renderClientList());

    function showAlert(msg) {
        const box = document.getElementById('alert-box');
        const text = String(msg || 'Error sin descripción');
        const label = document.getElementById('alert-msg');
        label.innerText = text;
        label.style.color = '#111827';
        box.style.background = '#ffffff';
        box.classList.remove('hidden');
        setTimeout(() => box.classList.add('hidden'), 8000);
    }
    function showLoadError(error, stage='Carga local') {
        toggleBackgroundLock(true);
        const detail = error && error.stack ? error.stack : (error && error.message ? error.message : String(error || 'Error desconocido'));
        document.getElementById('load-error-detail').textContent = `${stage}\n\n${detail}`;
        document.getElementById('load-error-panel').classList.remove('hidden');
    }
    function closeLoadError(){
        toggleBackgroundLock(false);
        document.getElementById('load-error-panel').classList.add('hidden');
    }
    async function copyLoadError(){try{await navigator.clipboard.writeText(document.getElementById('load-error-detail').textContent)}catch(e){console.error(e)}}

    function normalizeKey(key) { return String(key || "").toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9_]/g, ""); }
    function getSku(item) { return String(item.codigo || item.producto || "").trim(); }
    function getDescription(item) { return String(item.descripcion_producto || item.descripcionproducto || item.descripcion || "SIN DESCRIPCIÓN").trim(); }
    function getBrand(item) { return String(item.marca || "SIN MARCA").trim(); }
    function getFamily(item) { return String(item.familias || item.familia || "SIN FAMILIA").trim(); }
    function getSistema(item) { return String(item.sistema || "SIN SISTEMA").trim(); }
    function getTipoServicio(item) { return String(item.servicio || item.tipodeservicio || item.tipo_servicio || "SIN SERVICIO").trim(); }
    function getArmadora(item) { return String(item.armadora || "SIN ARMADORA").trim(); }
    function getProductImageUrl(item) { return String(item.urlimagenproducto || item.url_imagen_producto || "").trim(); }
    function getBrandLogoUrl(item) { return String(item.urlimagenmarca || item.url_imagen_marca || "").trim(); }

    function parseMoney(value) { 
        if (typeof value === 'number') return value;
        if (!value) return 0;
        let num = parseFloat(String(value).replace(/[^\d.-]/g, '').trim());
        return isNaN(num) ? 0 : num;
    }
    
    function parseDate(val) {
        if (!val) return null;
        if (typeof val === 'number' && !isNaN(val)) {
            let dateInfo = new Date((Math.floor(val - 25569) * 86400) * 1000);
            return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate(), 12, 0, 0);
        }
        const raw = String(val).trim(); if (!raw) return null;
        const isoMatch = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
        if (isoMatch) return new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10), 12, 0, 0);
        const localMatch = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
        if (localMatch) {
            let p1 = parseInt(localMatch[1], 10), p2 = parseInt(localMatch[2], 10), p3 = parseInt(localMatch[3], 10);
            if (p1 > 12) return new Date(p3, p2 - 1, p1, 12, 0, 0); 
            else return new Date(p3, p1 - 1, p2, 12, 0, 0); 
        }
        let d = new Date(raw); return isNaN(d.getTime()) ? null : d;
    }

    function financialPeriodKey(year, month) { return `${year}-${String(month + 1).padStart(2, '0')}`; }
    function extractFinancialHistory(item) {
        const history = {};
        const months = {enero:0,febrero:1,marzo:2,abril:3,mayo:4,junio:5,julio:6,agosto:7,septiembre:8,octubre:9,noviembre:10,diciembre:11};
        Object.entries(item || {}).forEach(([rawKey,value]) => {
            const key = normalizeKey(rawKey).replace(/_/g,'');
            const match = key.match(/^(venta|cuota|objetivoenfoque)(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)(20\d{2})$/);
            if(!match) return;
            const period = financialPeriodKey(Number(match[3]), months[match[2]]);
            if(!history[period]) history[period] = {venta:0,cuota:0,objetivoEnfoque:0};
            const metric = match[1] === 'objetivoenfoque' ? 'objetivoEnfoque' : match[1];
            history[period][metric] = parseMoney(value);
        });
        return history;
    }
    function financialValue(client,year,month,metric){const p=financialPeriodKey(year,month);return Number(client&&client.financiero&&client.financiero[p]?client.financiero[p][metric]||0:0)}
    function financialYTD(client,year,metric,throughMonth=CURRENT_MONTH){let total=0;for(let m=0;m<=throughMonth;m++)total+=financialValue(client,year,m,metric);return total}
    function refreshHierarchyFinancialTotals(){companyTree.forEach(reg=>{reg.totalVenta=0;reg.totalCuota=0;reg.asesores.forEach(adv=>{adv.totalVenta=0;adv.totalCuota=0;adv.clientes.forEach(c=>{c.venta=financialValue(c,CURRENT_YEAR,CURRENT_MONTH,'venta');c.cuota=financialValue(c,CURRENT_YEAR,CURRENT_MONTH,'cuota');c.objetivoEnfoque=financialValue(c,CURRENT_YEAR,CURRENT_MONTH,'objetivoEnfoque');adv.totalVenta+=c.venta;adv.totalCuota+=c.cuota;reg.totalVenta+=c.venta;reg.totalCuota+=c.cuota})})})}
    function clientPurchasedSkus(client){return new Set((client&&client.rawItems||[]).map(getSku).filter(Boolean).map(x=>String(x).trim().toUpperCase()))}
    function applyOpportunityFilter(products){const mode=document.getElementById('strategy-opportunity-filter')?.value||'all';const famMode=document.getElementById('strategy-familias-enfoque-filter')?.value||'all';let result=[...products];if(mode==='new'&&selectedClient){const bought=clientPurchasedSkus(selectedClient);result=result.filter(p=>!bought.has(getSku(p).toUpperCase()))}if(famMode!=='all'){result=result.filter(p=>String(p.familias_enfoque??p.familiasenfoque??'').trim().toLowerCase()===famMode)}return result}
    function isNewProduct(item){const v=String(item.codigo_nuevo??item.codigonuevo??item.CODIGO_NUEVO??'').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');return v.split(/[^a-z0-9]+/).includes('si')}
    function getSemaforoProyeccionClass(porcentaje) { return porcentaje >= 100 ? "glow-green" : (porcentaje >= 85 ? "glow-orange" : "glow-red"); }
    function getSemaforoProyeccionTextClass(porcentaje) { return porcentaje >= 100 ? "status-text-green" : (porcentaje >= 85 ? "status-text-orange" : "status-text-red"); }

    async function handleFile(file) {
        const fileName = String(file.name || '').trim();
        const lowerName = fileName.toLowerCase();



        loaderIcon.classList.add('hidden');
        loaderSpinner.classList.remove('hidden');
        document.getElementById('loader-text').innerText = 'Leyendo y descomprimiendo JSON.GZ local...';

        try {
            const compressed = new Uint8Array(await file.arrayBuffer());

            if (compressed.length < 2 || compressed[0] !== 0x1f || compressed[1] !== 0x8b) {
                throw new Error('El archivo seleccionado no tiene una firma GZIP valida.');
            }
            if (!window.pako) {
                throw new Error('No se pudo cargar la libreria Pako para descomprimir.');
            }

            const jsonText = pako.ungzip(compressed, { to: 'string' });
            if (!jsonText || !jsonText.trim()) {
                throw new Error('El contenido JSON descomprimido esta vacio.');
            }

            const payload = JSON.parse(jsonText.replace(/^\uFEFF/, '').trim());
            processCloudPayload(payload);
        } catch (err) {
            console.error('Error cargando JSON.GZ local:', err);
            const message = err && err.message ? err.message : String(err || 'Error desconocido');
            showAlert('ERROR JSON.GZ LOCAL: ' + message);
            showLoadError(err, 'Carga local de ' + fileName);
            resetCarga();
        }
    }

    function resetCarga() { loaderIcon.classList.remove('hidden'); loaderSpinner.classList.add('hidden'); fileInput.value = ''; }

    function normalizeRow(row) {
        if (!row || Object.keys(row).length === 0) return {};
        const item = {};
        for (let k in row) {
            if (!k) continue;
            const rawNorm = k.trim().toLowerCase();
            const normalizedK = normalizeKey(k);
            item[rawNorm] = row[k];
            item[normalizedK] = row[k];
        }
        return item;
    }


    let ACCESS_SESSION=null,CLOUD_ACCESS_USERS=[];
    const ACCESS_KEY='apymsa_access_users_v1',ACCESS_SESSION_KEY='apymsa_access_session_v1';
    function accessNorm(v){return normalizeBrandName(v).replace(/[^A-Z0-9]/g,'')}
    async function hashPin(pin){const data=new TextEncoder().encode(String(pin));const hash=await crypto.subtle.digest('SHA-256',data);return [...new Uint8Array(hash)].map(b=>b.toString(16).padStart(2,'0')).join('')}
    function accessUsers(){return CLOUD_ACCESS_USERS.slice()}
    function saveAccessUsers(v){console.warn('Los usuarios se administran en la hoja Accesos del Excel.')}
    function initializeAccessControl(){const overlay=document.getElementById('access-overlay');overlay.classList.remove('hidden');step1.classList.add('hidden');document.getElementById('access-title').innerText='Acceso Inteligencia Comercial';document.getElementById('access-help').innerText='Para asesores, el usuario es el asesor_id de Jerarquia. Para gerentes, usa el usuario definido en Accesos.'}
    async function submitAccess(){const user=document.getElementById('access-user').value.trim(),pin=document.getElementById('access-pin').value,err=document.getElementById('access-error');err.classList.add('hidden');if(!user||!pin){err.innerText='Captura usuario y PIN';err.classList.remove('hidden');return}const hash=await hashPin(pin),account=accessUsers().find(x=>accessNorm(x.user)===accessNorm(user)&&x.hash===hash&&x.active!==false);if(!account){err.innerText='Usuario o PIN incorrecto, o usuario inactivo';err.classList.remove('hidden');return}ACCESS_SESSION={user:account.user,role:account.role,entity:account.entity};document.getElementById('access-overlay').classList.add('hidden');document.getElementById('btnLogout').classList.remove('hidden');document.getElementById('btnAccessAdmin').classList.add('hidden');applyAccessLanding()}
    function applyAccessLanding(){
        step1.classList.add('hidden');step2.classList.remove('hidden');btnBack.classList.remove('hidden');
        document.getElementById('client-view').classList.add('hidden');document.getElementById('advisor-view').classList.add('hidden');document.getElementById('regional-view').classList.remove('hidden');
        if(ACCESS_SESSION.role==='TERRITORIAL'){renderTerritorialCards();return;}
        if(ACCESS_SESSION.role==='REGIONAL'){const reg=companyTree.get(ACCESS_SESSION.entity);if(!reg){showAlert('REGIÓN ASIGNADA NO ENCONTRADA');return;}currentSelectedTerritorial=reg.territorial;currentSelectedRegional=null;currentSelectedAdvisor=null;document.getElementById('directory-title').innerHTML='Región autorizada de <span class="text-neon-cyan">'+reg.territorial+'</span>';document.getElementById('directory-subtitle').innerText='Selecciona tu tarjeta regional para comenzar';document.getElementById('btnBackTerritories').classList.add('hidden');document.getElementById('territorial-dashboard').classList.remove('hidden');renderRegionalCards();switchTerritorialTab('regions');appCurrentState='INICIO_REGIONAL';return;}
        let found=null,reg=null;for(const r of companyTree.values()){if(r.asesores.has(ACCESS_SESSION.entity)){reg=r;found=r.asesores.get(ACCESS_SESSION.entity);break}}if(!found){showAlert('ASESOR ASIGNADO NO ENCONTRADO');return;}currentSelectedTerritorial=reg.territorial;currentSelectedRegional=reg;currentSelectedAdvisor=null;document.getElementById('regional-view').classList.add('hidden');document.getElementById('advisor-view').classList.remove('hidden');document.getElementById('current-regional-title').innerText=reg.name;renderAdvisorCards();switchRegionalTab('advisors');appCurrentState='INICIO_ASESOR';
    }
    function logoutAccess(){ACCESS_SESSION=null;location.reload()}
    function openAccessAdmin(){showAlert('LOS USUARIOS SE ADMINISTRAN EN LA HOJA ACCESOS DEL EXCEL Y SE ACTUALIZAN AL REGENERAR NORTE.JSON.GZ.')}
    function closeAccessAdmin(){document.getElementById('access-admin-modal').classList.add('hidden')}
    function refreshAccessEntityOptions(){const role=document.getElementById('new-access-role').value,sel=document.getElementById('new-access-entity');let vals=[];if(role==='TERRITORIAL')vals=['*'];else if(role==='REGIONAL')vals=[...companyTree.keys()].sort();else vals=[...new Set([...companyTree.values()].flatMap(r=>[...r.asesores.keys()]))].sort();sel.innerHTML=vals.map(v=>`<option value="${v}">${v==='*'?'ACCESO COMPLETO':v}</option>`).join('')}
    async function saveAccessUser(){const user=document.getElementById('new-access-user').value.trim(),pin=document.getElementById('new-access-pin').value,role=document.getElementById('new-access-role').value,entity=document.getElementById('new-access-entity').value;if(!user||pin.length<4){showAlert('USUARIO Y PIN DE AL MENOS 4 DÍGITOS.');return}const users=accessUsers().filter(x=>accessNorm(x.user)!==accessNorm(user));users.push({user,hash:await hashPin(pin),role,entity});saveAccessUsers(users);document.getElementById('new-access-user').value='';document.getElementById('new-access-pin').value='';renderAccessUsers()}
    function deleteAccessUser(user){const users=accessUsers();if(users.length<=1){showAlert('DEBE EXISTIR AL MENOS UN USUARIO.');return}saveAccessUsers(users.filter(x=>x.user!==user));renderAccessUsers()}
    function renderAccessUsers(){const box=document.getElementById('access-user-list');box.innerHTML=accessUsers().map(x=>`<div class="flex justify-between gap-3 border rounded-xl p-3"><div><b>${x.user}</b><div class="text-xs">${x.role} · ${x.entity}</div></div><button onclick="deleteAccessUser('${String(x.user).replace(/'/g,"\'")}')" class="text-red-700 font-black">Eliminar</button></div>`).join('')}
    function canOpenRegional(name){return !ACCESS_SESSION||ACCESS_SESSION.role==='TERRITORIAL'||(ACCESS_SESSION.role==='REGIONAL'&&ACCESS_SESSION.entity===name)}
    function canOpenAdvisor(name){return !ACCESS_SESSION||ACCESS_SESSION.role==='TERRITORIAL'||ACCESS_SESSION.role==='REGIONAL'||(ACCESS_SESSION.role==='ASESOR'&&ACCESS_SESSION.entity===name)}

    function processData(dataProductos, dataClientes, dataJerarquia, dataVentas, dataDisciplina, dataAccesos) {
        CLOUD_ACCESS_USERS=(dataAccesos||[]).map(normalizeRow).filter(x=>String(x.estatus||'ACTIVO').toUpperCase()==='ACTIVO').map(x=>({user:String(x.usuario||'').trim(),hash:String(x.pin_hash||x.pinhash||'').trim().toLowerCase(),role:String(x.rol||'').trim().toUpperCase(),entity:String(x.entidad||x.gerente_regional||x.asesor||x.gerente_territorial||'*').trim().toUpperCase(),active:true})).filter(x=>x.user&&x.hash&&['TERRITORIAL','REGIONAL','ASESOR'].includes(x.role));
        if(!CLOUD_ACCESS_USERS.length)throw new Error('La colección Accesos no contiene usuarios activos válidos.');
        companyTree.clear(); allClientsFlat.clear(); clientsById.clear(); globalProducts.clear();
        dataQuality={ventasSinCliente:0,ventasAmbiguas:0,codigosSinMaestro:new Set(),asesoresSinJerarquia:new Set(),disciplinaSinCliente:0,clientesDuplicadosCV:[]};
        globalSalesMn = {}; globalSalesPesos = {}; globalPiezas = {}; globalClientsPerSku = {};
        globalPromociones = {}; globalPromoText = {}; globalComentarioPromo = {}; globalLowestPrice = {};
        globalSkuStats = {};

        dataJerarquia.forEach(row => {
            const item = normalizeRow(row);
            const territorialName = String(item.gerente_territorial || item.gerenteterritorial || item.territorial || "SIN GERENTE TERRITORIAL").trim().toUpperCase();
            const regName = String(item.gerente_regional || item.gerente || item.regional || "SIN GERENTE REGIONAL").trim().toUpperCase();
            const zoneName = String(item.sucursal_vendedor || item.sucursal || "SIN SUCURSAL").trim().toUpperCase();
            const advName = String(item.asesor || item.vendedor || item.ejecutivo || "SIN ASESOR").trim().toUpperCase();

            if (!companyTree.has(regName)) {
                companyTree.set(regName, { name: regName, territorial: territorialName, region: zoneName, totalVenta: 0, totalCuota: 0, asesores: new Map() });
            }
            const regNode = companyTree.get(regName);
            if (territorialName !== "SIN GERENTE TERRITORIAL") regNode.territorial = territorialName;
            if (!regNode.asesores.has(advName)) {
                regNode.asesores.set(advName, { name: advName, sucursal: zoneName, totalVenta: 0, totalCuota: 0, clientes: new Map() });
            }
        });

        dataClientes.forEach(row => {
            const item=normalizeRow(row);
            const cliId=String(item.cliente_id||item.clienteid||'').trim().toUpperCase();
            const cliVendId=String(item.cliente_vendedor_id||item.clientevendedorid||'').trim().toUpperCase();
            if(!cliId||!cliVendId)return;
            if(allClientsFlat.has(cliVendId)){dataQuality.clientesDuplicadosCV.push(cliVendId);return}
            const advName=String(item.asesor||'SIN ASESOR').trim().toUpperCase();
            const cliName=String(item.razon_social||item.razonsocial||'SIN NOMBRE').trim();
            const ciudad=String(item.ciudad||'').trim();
            const tipoCliente=String(item.tipo_cliente||item.tipo||'GENERAL').trim().toUpperCase();
            let focus=String(item.enfoque||'GENERAL').trim();if(!focus||focus==='0'||focus.toUpperCase()==='NULL')focus='GENERAL';
            const financiero=extractFinancialHistory(item);
            let foundReg='SIN JERARQUIA ASIGNADA',advNode=null,regNode=null;
            for(const [rName,candidate] of companyTree.entries()){if(candidate.asesores.has(advName)){foundReg=rName;regNode=candidate;advNode=candidate.asesores.get(advName);break}}
            if(!advNode){dataQuality.asesoresSinJerarquia.add(advName);if(!companyTree.has(foundReg))companyTree.set(foundReg,{name:foundReg,territorial:'SIN GERENTE TERRITORIAL',region:'SIN SUCURSAL',totalVenta:0,totalCuota:0,asesores:new Map()});regNode=companyTree.get(foundReg);if(!regNode.asesores.has(advName))regNode.asesores.set(advName,{name:advName,sucursal:'SIN SUCURSAL',totalVenta:0,totalCuota:0,clientes:new Map()});advNode=regNode.asesores.get(advName)}
            const bonoSemanas=[1,2,3].map(n=>{const raw=item[`bono_semana_${n}`]??item[`bonosemana${n}`]??'';return{capturado:String(raw??'').trim()!=='',valor:parseMoney(raw)}});
            const cliObj={id:cliId,cliVendId,uniqueKey:cliVendId,name:cliName,ciudad,focus,advisor:advName,regional:foundReg,territorial:regNode.territorial,tipoCliente,financiero,bonoSemanas,venta:0,cuota:0,objetivoEnfoque:0,piezas:0,disciplina:{},rawItems:[]};
            advNode.clientes.set(cliVendId,cliObj);allClientsFlat.set(cliVendId,cliObj);if(!clientsById.has(cliId))clientsById.set(cliId,[]);clientsById.get(cliId).push(cliObj);
        });
        refreshHierarchyFinancialTotals();
        dataProductos.forEach(row => {
            const item = normalizeRow(row);
            const sku = getSku(item);
            if (!sku) return;
            const tempPromoVal = parseMoney(item.precio_solicitado ?? item.preciosolicitado ?? item["precio solicitado"] ?? 0);
            const rawPromoText = String(item["tipo promocion"] || item["tipopromocion"] || item["promocion"] || item["promo"] || "").trim();
            const rawComment = String(item["comentario"] || item["comentario promocion"] || item["comentariopromocion"] || "").trim();
            const isValidPromoText = value => { const normalized = String(value || "").trim().toLowerCase(); return normalized && !["0", "null", "undefined", "na", "n/a"].includes(normalized); };
            const tempPromoText = isValidPromoText(rawPromoText) ? rawPromoText : "";
            const tempPromoComment = isValidPromoText(rawComment) ? rawComment : "";

            if (tempPromoVal > 0) globalPromociones[sku] = tempPromoVal;
            if (tempPromoText) globalPromoText[sku] = tempPromoText;
            if (tempPromoComment) globalComentarioPromo[sku] = tempPromoComment;
            item._precio_solicitado_catalogo = tempPromoVal;
            item._nombre_promocion_catalogo = tempPromoComment;

            globalProducts.set(sku, item);
        });

        dataVentas.forEach(row=>{
            const item=normalizeRow(row),cliId=String(item.cliente_id||item.clienteid||'').trim().toUpperCase(),sku=getSku(item).toUpperCase();if(!cliId||!sku)return;
            const product=globalProducts.get(sku);if(!product)dataQuality.codigosSinMaestro.add(sku);item.marca=product?getBrand(product):'SIN MARCA';item.familia=product?getFamily(product):'SIN FAMILIA';
            const vpesos=parseMoney(item['venta en pesos']||item.venta_en_pesos||0),pz=parseMoney(item['venta en piezas']||item.venta_en_piezas||0);item._vmn_calculado=vpesos;item._venta_piezas_calculada=pz;item.fecha_doc=parseDate(item.fecha_factura||item.fecha);
            const rel=clientsById.get(cliId)||[];let cliNode=null;if(rel.length===1)cliNode=rel[0];else if(!rel.length)dataQuality.ventasSinCliente++;else dataQuality.ventasAmbiguas++;if(cliNode)cliNode.rawItems.push(item);
            if(pz>0&&vpesos>0){const unit=vpesos/pz;if(!globalLowestPrice[sku]||unit<globalLowestPrice[sku])globalLowestPrice[sku]=unit}
            globalSalesPesos[sku]=(globalSalesPesos[sku]||0)+vpesos;globalSalesMn[sku]=(globalSalesMn[sku]||0)+vpesos;globalPiezas[sku]=(globalPiezas[sku]||0)+pz;
            const tipo=cliNode?cliNode.tipoCliente:'SIN ASIGNACION';if(!globalSkuStats[sku])globalSkuStats[sku]={};if(!globalSkuStats[sku][tipo])globalSkuStats[sku][tipo]={mn:0,pz:0};globalSkuStats[sku][tipo].mn+=vpesos;globalSkuStats[sku][tipo].pz+=pz;if(!globalClientsPerSku[sku])globalClientsPerSku[sku]=new Set();globalClientsPerSku[sku].add(cliId);
        });
        const byCV=allClientsFlat;dataDisciplina.forEach(row=>{const i=normalizeRow(row),key=String(i.cliente_vendedor_id||i.clientevendedorid||'').trim().toUpperCase(),c=byCV.get(key);if(!c)return;MESES_NOMBRES.forEach(m=>{const n=normalizeKey(m);c.disciplina[n]={agendado:parseMoney(i[`agendado_${n}_2026`]||i[`agendado${n}2026`]||0)>0?1:0,visitado:Math.max(0,Math.trunc(parseMoney(i[`visitado_${n}_2026`]||i[`visitado${n}2026`]||0)))}})});
        loaderIcon.classList.remove('hidden'); loaderSpinner.classList.add('hidden');
        step1.classList.add('hidden'); step2.classList.remove('hidden'); btnBack.classList.remove('hidden');
        appCurrentState = "PANEL_REGIONALES";
        document.getElementById('summary-count').innerText = `${allClientsFlat.size} clientes | ${globalProducts.size} códigos`;
        populateSelectors();
        initializeAccessControl();
    }

    function dateOnly(d){return new Date(d.getFullYear(),d.getMonth(),d.getDate(),12,0,0)}
    function sameDate(a,b){return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate()}
    function previousOperationalDay(from){const d=dateOnly(from);d.setDate(d.getDate()-1);while(d.getDay()===0)d.setDate(d.getDate()-1);return d}
    function lastBusinessCutoff(){return previousOperationalDay(NOW)}
    function effectiveDaysToDate(c){if(c.getFullYear()!==CURRENT_YEAR||c.getMonth()!==CURRENT_MONTH)return 0;let n=0;for(let d=1;d<=c.getDate();d++){const w=new Date(CURRENT_YEAR,CURRENT_MONTH,d).getDay();if(w>=1&&w<=5)n++;else if(w===6)n+=.5}return n}
    function operationalMetrics(cs){const cutoff=lastBusinessCutoff(),prev=previousOperationalDay(cutoff),start=new Date(CURRENT_YEAR,CURRENT_MONTH,1,12),end=dateOnly(cutoff),prevEnd=dateOnly(prev);let ventaDia=0,ventaMes=0,cuota=0,cc=0,cp=0,rec=0;cs.forEach(c=>{cuota+=Number(c.cuota||0);ventaMes+=Number(c.venta||0);const s=(c.rawItems||[]).filter(i=>i.fecha_doc&&Number(i._vmn_calculado||0)>0).sort((a,b)=>a.fecha_doc-b.fecha_doc);s.forEach(i=>{if(sameDate(i.fecha_doc,cutoff))ventaDia+=Number(i._vmn_calculado||0)});if(s.some(i=>i.fecha_doc>=start&&i.fecha_doc<=end))cc++;if(s.some(i=>i.fecha_doc>=start&&i.fecha_doc<=prevEnd))cp++;const ms=s.filter(i=>i.fecha_doc>=start&&i.fecha_doc<=end),prior=s.filter(i=>i.fecha_doc<start);if(ms.length&&prior.length&&(ms[0].fecha_doc-prior[prior.length-1].fecha_doc)/86400000>90)rec++});const elapsed=effectiveDaysToDate(cutoff),total=obtenerDiasEfectivosDelMes(CURRENT_YEAR,CURRENT_MONTH),remaining=Math.max(total-elapsed,0),promedio=elapsed?ventaMes/elapsed:0,cuotaDiaria=total?cuota/total:0,recalculada=remaining?Math.max(cuota-ventaMes,0)/remaining:0,esperado=total?elapsed/total*100:0,real=cuota?ventaMes/cuota*100:0;return{cutoff,ventaDia,ventaMes,cuota,promedio,cuotaDiaria,recalculada,inc:Math.max(cc-cp,0),cc,cp,rec,esperado,real,diff:real-esperado,elapsed,total,remaining}}
    function renderOperational(cs,scope){const e=document.getElementById(scope+'-operational');if(!e)return;const m=operationalMetrics(cs),f=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0}),d=m.cutoff.toLocaleDateString('es-MX',{day:'2-digit',month:'short',year:'numeric'}).replace('.','').toUpperCase(),lamp=m.diff>=0?'green':m.diff>=-5?'amber':'red';e.innerHTML=`<section class="operational-block"><div class="operational-head"><div><div class="operational-title">Corte operativo diario</div></div><div class="operational-date">Corte al ${d} · último día hábil</div></div><div class="operational-grid"><div class="operational-card"><small>Venta último día hábil</small><b>${f.format(m.ventaDia)}</b></div><div class="operational-card"><small>Promedio diario del mes</small><b>${f.format(m.promedio)}</b></div><div class="operational-card white"><small>Cuota diaria original</small><b>${f.format(m.cuotaDiaria)}</b></div><div class="operational-card peach"><small>Cuota diaria recalculada</small><b>${f.format(m.recalculada)}</b></div><div class="operational-card"><small>Incremento clientes con venta</small><b>${m.inc}</b></div><div class="operational-card"><small>Clientes recuperados +90 días</small><b>${m.rec}</b></div><div class="operational-card white"><small>Avance esperado</small><b>${m.esperado.toFixed(1)}%</b></div><div class="operational-card white"><i class="operational-lamp ${lamp}"></i><small>Avance real</small><b>${m.real.toFixed(1)}%</b></div></div></section>`}
    function destroySegmentCharts(scope,type){['ENFOQUE','FUGA','GENERAL'].forEach(k=>{const id=`${scope}-${type}-${k}`;if(segmentChartInstances[id]){segmentChartInstances[id].destroy();delete segmentChartInstances[id]}})}
    
    function renderSegmentCharts(cs,scope,type){
        const box=document.getElementById(`${scope}-${type}-segment-charts`);
        if(!box)return;
        destroySegmentCharts(scope,type);
        const cfg={ENFOQUE:{title:'Clientes enfoque',color:'#245F70',cls:'segment-enfoque'},FUGA:{title:'Clientes en fuga',color:'#D92D20',cls:'segment-fuga'},GENERAL:{title:'Clientes generales',color:'#555B6E',cls:'segment-general'}};
        const g=groups(cs);
        const mesesTranscurridos=CURRENT_MONTH+1;
        box.innerHTML=['ENFOQUE','FUGA','GENERAL'].map(k=>`<div class="segment-chart-card ${cfg[k].cls}"><div class="segment-chart-title">${cfg[k].title}</div><div class="segment-chart-subtitle">${type==='intel'?'Venta mensual 2026':'Clientes visitados por mes 2026'} · línea promedio enero a ${MESES_NOMBRES[CURRENT_MONTH]}</div><div class="segment-chart-canvas"><canvas id="chart-${scope}-${type}-${k}"></canvas></div></div>`).join('');
        ['ENFOQUE','FUGA','GENERAL'].forEach(k=>{
            const values=Array(12).fill(0);
            if(type==='intel'){
                g[k].forEach(c=>{for(let m=0;m<12;m++)values[m]+=financialValue(c,CURRENT_YEAR,m,'venta')});
            }else{
                g[k].forEach(c=>{for(let m=0;m<12;m++)values[m]+=(dv(c,m,'visitado')>0?1:0)});
            }
            const acumuladoTranscurrido=values.slice(0,mesesTranscurridos).reduce((a,b)=>a+b,0);
            const avg=mesesTranscurridos>0?acumuladoTranscurrido/mesesTranscurridos:0;
            const id=`${scope}-${type}-${k}`,money=type==='intel';
            segmentChartInstances[id]=new Chart(document.getElementById('chart-'+id),{
                data:{labels:MESES_NOMBRES.map(x=>x.slice(0,3)),datasets:[
                    {type:'bar',label:money?'Venta mensual':'Clientes visitados',data:values,backgroundColor:cfg[k].color,borderColor:'#ffffff',borderWidth:2,borderRadius:6,maxBarThickness:54},
                    {type:'line',label:`Promedio ${mesesTranscurridos} meses`,data:Array(12).fill(avg),borderColor:'#00f3ff',backgroundColor:'#00f3ff',borderWidth:4,pointRadius:0,tension:0,fill:false}
                ]},
                plugins:[ChartDataLabels],
                options:{responsive:true,maintainAspectRatio:false,layout:{padding:{top:22}},scales:{y:{beginAtZero:true,ticks:{color:'#DCEAF0',font:{weight:'bold',size:11},precision:money?undefined:0,callback:v=>money?'$'+Intl.NumberFormat('es-MX',{notation:'compact'}).format(v):v},grid:{color:'rgba(85,91,110,.25)'}},x:{grid:{display:false},ticks:{color:'#DCEAF0',font:{weight:'900',size:11},autoSkip:false,maxRotation:0,minRotation:0,callback:function(value){return String(this.getLabelForValue(value)).slice(0,3).toUpperCase()}}}},plugins:{legend:{position:'top',labels:{color:'#DCEAF0',boxWidth:14,font:{size:11,weight:'bold'}}},tooltip:{callbacks:{afterBody:()=>`Promedio enero-${MESES_NOMBRES[CURRENT_MONTH]}: ${money?new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0}).format(avg):avg.toFixed(1)}`}},datalabels:{display:ctx=>ctx.dataset.type==='bar'&&ctx.raw>0,anchor:'end',align:'top',color:'#ffffff',textShadowBlur:6,textShadowColor:'#000000',font:{size:11,weight:'900'},formatter:v=>money?'$'+Intl.NumberFormat('es-MX',{notation:'compact'}).format(v):v}}}
            });
        });
    }

    function commercialBlock(c){const t=`${c.focus||''} ${c.tipoCliente||''}`.toUpperCase();if(t.includes('FUGA')||t.includes('SIN VENTA'))return'FUGA';if(t.includes('ENFOQUE'))return'ENFOQUE';return'GENERAL'}
    function groups(cs){const g={ENFOQUE:[],FUGA:[],GENERAL:[]};cs.forEach(c=>g[commercialBlock(c)].push(c));return g}
    function saleMonth(c,y,m){return financialValue(c,y,m,'venta')}
    function intelCard(t,cs,color){const v=cs.reduce((a,c)=>a+saleMonth(c,2026,CURRENT_MONTH),0),aa=cs.reduce((a,c)=>a+saleMonth(c,2025,CURRENT_MONTH),0),o=cs.reduce((a,c)=>a+financialValue(c,CURRENT_YEAR,CURRENT_MONTH,'cuota'),0),proy=DIAS_TRANSCURRIDOS>0?(v/DIAS_TRANSCURRIDOS)*DIAS_EFECTIVOS_MES:0,pct=o?(proy/o*100):0,lamp=pct>=100?'green':pct>=85?'amber':'red',f=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0});return`<div class="segment-kpi-card bg-black/50 border rounded-2xl p-5" style="border-color:${color}"><span class="segment-lamp segment-lamp-${lamp}" title="Proyección ${pct.toFixed(1)}% · ${pct>=100?'En ruta':pct>=85?'Atención':'Urgente'}"></span><h3 class="text-sm text-white font-black uppercase">${t}</h3><p class="text-[8px] text-slate-400 font-bold uppercase mb-4">${MESES_NOMBRES[CURRENT_MONTH]} 2026</p><div class="grid grid-cols-3 gap-2"><div><small>VENTA 2026</small><b class="block text-neon-cyan text-[11px]">${f.format(v)}</b></div><div><small>VENTA 2025</small><b class="block text-neon-green text-[11px]">${f.format(aa)}</b></div><div><small>CUOTA ${MESES_NOMBRES[CURRENT_MONTH]} 2026</small><b class="block text-neon-orange text-[11px]">${f.format(o)}</b></div></div><div class="segment-projection-footer"><span>${cs.length} CLIENTES</span><strong class="segment-projection-value segment-projection-${lamp}">PROYECCIÓN ${pct.toFixed(1)}%</strong></div></div>`}
    function intelTop(t,cs,color){const arr=worstProjectionTen(cs);return`<div class="bg-black/40 border border-white/10 rounded-2xl p-5"><h4 class="text-[12px] font-black uppercase mb-4" style="color:${color}">${t} · peores proyecciones</h4><div class="grid grid-cols-1 md:grid-cols-2 gap-4">${arr.map(projectionClientCard).join('')}</div></div>`}function renderIntel(cs,scope){renderOperational(cs,scope);const g=groups(cs);document.getElementById(`${scope}-intel-kpis`).innerHTML=intelCard('1. Clientes enfoque',g.ENFOQUE,'#00f3ff')+intelCard('2. Clientes en fuga',g.FUGA,'#ff004c')+intelCard('3. Clientes generales',g.GENERAL,'#64748b');renderSegmentCharts(cs,scope,'intel');document.getElementById(`${scope}-intel-tops`).innerHTML=intelTop('1. Principales clientes enfoque',g.ENFOQUE,'#00f3ff')+intelTop('2. Principales clientes en fuga',g.FUGA,'#ff004c')+intelTop('3. Principales clientes generales',g.GENERAL,'#cbd5e1')}
    function dv(c,m,t){const k=normalizeKey(MESES_NOMBRES[m]);return c.disciplina&&c.disciplina[k]?Number(c.disciplina[k][t]||0):0}function visits(c){let n=0;for(let m=0;m<12;m++)n+=dv(c,m,'visitado');return n}function recurrence(c){let current=0,max=0,total=0;for(let m=0;m<=CURRENT_MONTH;m++){if(!dv(c,m,'visitado')){current++;total++;max=Math.max(max,current)}else current=0}return{run:current,max,total}}
    function salesYTD(c,year){return financialYTD(c,year,'venta')}
    function annualQuotaYTD(c){return financialYTD(c,CURRENT_YEAR,'cuota')}
    function currentProjectionMetrics(c){const ventaMes=financialValue(c,CURRENT_YEAR,CURRENT_MONTH,'venta'),ventaAnterior=financialValue(c,CURRENT_YEAR-1,CURRENT_MONTH,'venta'),cuota=financialValue(c,CURRENT_YEAR,CURRENT_MONTH,'cuota'),proyeccion=DIAS_TRANSCURRIDOS>0?(ventaMes/DIAS_TRANSCURRIDOS)*DIAS_EFECTIVOS_MES:0,pct=cuota>0?proyeccion/cuota*100:0,crec=ventaAnterior>0?(ventaMes-ventaAnterior)/ventaAnterior*100:(ventaMes>0?100:0),rec=recurrence(c),ventaYTD=financialYTD(c,CURRENT_YEAR,'venta'),cuotaYTD=financialYTD(c,CURRENT_YEAR,'cuota'),deficitAnual=Math.max(cuotaYTD-ventaYTD,0),deficitMes=Math.max(cuota-ventaMes,0),visitadoMes=dv(c,CURRENT_MONTH,'visitado')>0;return{ventaMes,ventaAnterior,cuota,proyeccion,pct,crec,vis:visits(c),rec,ventaYTD,cuotaYTD,deficitAnual,deficitMes,visitadoMes}}
    function urgencyFor(m){if(!m.visitadoMes||m.rec.run>=2||m.pct<70)return{c:'red',l:'URGENTE'};if(m.rec.run===1||m.pct<100)return{c:'amber',l:'ATENCIÓN'};return{c:'green',l:'EN RUTA'}}
    function projectionClientCard(c,i){const m=currentProjectionMetrics(c),u=urgencyFor(m),f=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0}),growth=m.crec>=0?'text-neon-green':'text-neon-red';return`<button onclick="openDisciplineDetail('${c.cliVendId}')" class="apple-depth-card"><div class="depth-card-head"><span class="urgency-light urgency-${u.c}"></span><div class="depth-title"><span class="depth-id">#${i+1} · CLIENTE ${c.id}</span><span class="depth-name" title="${c.name}">${c.name}</span></div><span class="urgency-label ${u.c}">${u.l}</span></div><div class="depth-body"><div class="flex justify-end mb-2">${m.visitadoMes?'<span class="visited-badge">VISITADO ESTE MES</span>':'<span class="no-visit-badge">NO VISITADO ESTE MES</span>'}</div><div class="depth-grid"><div class="depth-metric deficit-metric"><small>PÉRDIDA CUOTA 2026 YTD</small><b>${f.format(m.deficitAnual)}</b></div><div class="depth-metric deficit-metric"><small>PÉRDIDA OBJETIVO MES</small><b>${f.format(m.deficitMes)}</b></div><div class="depth-metric"><small>VENTA 2026 YTD</small><b class="text-neon-cyan">${f.format(m.ventaYTD)}</b></div><div class="depth-metric"><small>CUOTA 2026 YTD</small><b class="text-neon-orange">${f.format(m.cuotaYTD)}</b></div><div class="depth-metric"><small>VENTA MES ACTUAL</small><b class="text-neon-cyan">${f.format(m.ventaMes)}</b></div><div class="depth-metric"><small>CUOTA MES ACTUAL</small><b class="text-neon-orange">${f.format(m.cuota)}</b></div><div class="depth-metric"><small>PROYECCIÓN DE CIERRE</small><b class="${m.pct>=100?'text-neon-green':m.pct>=85?'text-neon-orange':'text-neon-red'}">${f.format(m.proyeccion)}</b></div><div class="depth-metric"><small>CUMPLIMIENTO PROYECTADO</small><b class="${m.pct>=100?'text-neon-green':m.pct>=85?'text-neon-orange':'text-neon-red'}">${m.pct.toFixed(1)}%</b></div><div class="depth-metric"><small>CRECIMIENTO VS ${CURRENT_YEAR-1}</small><b class="${growth}">${m.crec>=0?'+':''}${m.crec.toFixed(1)}%</b></div><div class="depth-metric"><small>VISITAS 2026</small><b class="text-neon-green">${m.vis}</b></div><div class="depth-metric depth-wide"><small>URGENCIA COMERCIAL</small><b class="${u.c==='green'?'text-neon-green':u.c==='amber'?'text-neon-orange':'text-neon-red'}">${u.l} · ${m.rec.run} MES(ES) SIN VISITA</b></div></div></div></button>`}
    function worstProjectionTen(cs){return cs.map(c=>({c,m:currentProjectionMetrics(c)})).sort((a,b)=>b.m.deficitAnual-a.m.deficitAnual||b.m.deficitMes-a.m.deficitMes||Number(a.m.visitadoMes)-Number(b.m.visitadoMes)||a.m.pct-b.m.pct).slice(0,10).map(x=>x.c)}
    function disciplineCard(t,cs,color){
        const agendados=cs.filter(c=>dv(c,CURRENT_MONTH,'agendado')>0).length;
        const visitados=cs.filter(c=>dv(c,CURRENT_MONTH,'visitado')>0).length;
        const noVisitados=cs.filter(c=>dv(c,CURRENT_MONTH,'visitado')<=0).length;
        const cumplimiento=cs.length>0?visitados/cs.length*100:0;
        const lamp=cumplimiento>=100?'green':cumplimiento>=85?'amber':'red';
        return `<div class="discipline-kpi-card segment-kpi-card bg-black/50 border rounded-2xl p-5" style="border-color:${color}"><span class="segment-lamp segment-lamp-${lamp}" title="${cumplimiento>=100?'En ruta':cumplimiento>=85?'Atención':'Urgente'}"></span>
            <h3 class="text-sm text-white font-black uppercase leading-tight">${t}</h3>
            <p class="text-[8px] text-slate-400 mt-1">${MESES_NOMBRES[CURRENT_MONTH]} 2026 · RECUENTO DE CLIENTES</p>
            <div class="discipline-kpi-grid">
                <div class="discipline-kpi-item"><b class="discipline-kpi-value text-white">${cs.length}</b><small class="discipline-kpi-label">Total de clientes</small></div>
                <div class="discipline-kpi-item"><b class="discipline-kpi-value text-neon-orange">${agendados}</b><small class="discipline-kpi-label">Agendados</small></div>
                <div class="discipline-kpi-item"><b class="discipline-kpi-value text-neon-green">${visitados}</b><small class="discipline-kpi-label">Visitados</small></div>
                <div class="discipline-kpi-item"><b class="discipline-kpi-value text-neon-red">${noVisitados}</b><small class="discipline-kpi-label">No visitados</small></div>
            </div>
            <p class="discipline-kpi-coverage ${cumplimiento>=100?'text-neon-green':'text-neon-red'}">Cobertura de visita ${cumplimiento.toFixed(1)}%</p>
        </div>`
    }
    function disciplineList(title,cs,color,asc){const arr=worstProjectionTen(cs);return`<div class="bg-black/40 border border-white/10 rounded-2xl p-5"><h4 class="text-[11px] font-black uppercase mb-4" style="color:${color}">${title} · peores proyecciones</h4><div class="grid grid-cols-1 md:grid-cols-2 gap-4">${arr.map(projectionClientCard).join('')}</div></div>`}function disciplineBlock(label,cs,color){return`<div><div class="flex justify-between mb-3"><h3 class="text-lg font-black uppercase" style="color:${color}">${label}</h3><span class="text-[9px] text-slate-400 font-black">${cs.length} CLIENTES · MÁXIMO 10</span></div>${disciplineList('Prioridad: déficit anual, después déficit mensual',cs,color,false)}</div>`}
    function renderDiscipline(cs,scope){currentDisciplineClients=cs.slice();const g=groups(cs),recurrent=cs.filter(c=>recurrence(c).run>=2).length,noVisitados=cs.filter(c=>dv(c,CURRENT_MONTH,'visitado')<=0).length;document.getElementById(`${scope}-discipline-kpis`).innerHTML=disciplineCard('1. Clientes enfoque',g.ENFOQUE,'#00f3ff')+disciplineCard('2. Clientes en fuga',g.FUGA,'#ff004c')+disciplineCard('3. Clientes generales',g.GENERAL,'#64748b');document.getElementById(`${scope}-discipline-summary`).innerHTML=`<div class="bg-black/50 border border-neon-cyan rounded-2xl p-5"><small class="text-slate-400">TOTAL DE CLIENTES</small><b class="block text-3xl text-neon-cyan">${cs.length}</b></div><button onclick="openNoVisitModal()" class="text-left bg-red-50 border border-red-300 rounded-2xl p-5 hover:shadow-xl transition-all"><small class="text-red-700 font-black">CLIENTES NO VISITADOS ESTE MES</small><b class="block text-3xl text-red-700">${noVisitados}</b><span class="text-[8px] text-red-700 font-black">ABRIR LISTADO DETALLADO</span></button>`;renderSegmentCharts(cs,scope,'discipline');document.getElementById(`${scope}-discipline-tops`).innerHTML=disciplineBlock('Clientes enfoque',g.ENFOQUE,'#00f3ff')+disciplineBlock('Clientes en fuga',g.FUGA,'#ff004c')+disciplineBlock('Clientes generales',g.GENERAL,'#334155')}
    function clientBranch(c){const reg=companyTree.get(c.regional),adv=reg&&reg.asesores?reg.asesores.get(c.advisor):null;return adv?adv.sucursal:'SIN SUCURSAL'}
    
    function openNoVisitModal(){
        toggleBackgroundLock(true);
        noVisitRowsCache=currentDisciplineClients.filter(c=>dv(c,CURRENT_MONTH,'visitado')<=0).map(c=>{const m=currentProjectionMetrics(c);return{c,deficit:m.deficitAnual,sucursal:clientBranch(c)}}).sort((a,b)=>b.deficit-a.deficit);document.getElementById('no-visit-month').innerText=MESES_NOMBRES[CURRENT_MONTH]+' '+CURRENT_YEAR;document.getElementById('no-visit-context').innerText=currentSelectedAdvisor?'ASESOR: '+currentSelectedAdvisor.name:(currentSelectedRegional?'REGIONAL: '+currentSelectedRegional.name:(currentSelectedTerritorial?'TERRITORIO: '+currentSelectedTerritorial:'VISTA ACTUAL'));document.getElementById('no-visit-search').value='';document.getElementById('no-visit-modal').classList.remove('hidden');renderNoVisitRows()}
    
    function closeNoVisitModal(){
        toggleBackgroundLock(false);
        document.getElementById('no-visit-modal').classList.add('hidden')
    }
    
    function renderNoVisitRows(){const q=String(document.getElementById('no-visit-search').value||'').toLowerCase().trim(),money=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0});const rows=noVisitRowsCache.filter(x=>`${x.c.id} ${x.c.name} ${x.c.advisor} ${x.c.regional} ${x.sucursal}`.toLowerCase().includes(q));document.getElementById('no-visit-count').innerText=rows.length+' CLIENTES NO VISITADOS';document.getElementById('no-visit-rows').innerHTML=rows.map(x=>`<tr><td><b>${x.c.id}</b></td><td>${x.c.name}</td><td>${x.c.advisor}</td><td><b class="text-red-700">${money.format(x.deficit)}</b></td><td>${x.c.regional}</td><td>${x.sucursal}</td></tr>`).join('')||'<tr><td colspan="6" class="text-center font-black">Sin coincidencias</td></tr>'}

    function getFilteredNoVisitRows(){const q=String(document.getElementById('no-visit-search').value||'').toLowerCase().trim();return noVisitRowsCache.filter(x=>`${x.c.id} ${x.c.name} ${x.c.advisor} ${x.c.regional} ${x.sucursal}`.toLowerCase().includes(q))}
    function exportNoVisitExcel(){const rows=getFilteredNoVisitRows();if(!rows.length){showAlert('NO HAY CLIENTES NO VISITADOS PARA EXPORTAR');return}const data=rows.map(x=>({'cliente_id':x.c.id,'razón_social':x.c.name,'asesor':x.c.advisor,'déficit venta 2026 vs cuota 2026':x.deficit,'gerente_regional':x.c.regional,'sucursal_vendedor':x.sucursal}));const ws=XLSX.utils.json_to_sheet(data);ws['!cols']=[{wch:16},{wch:42},{wch:28},{wch:28},{wch:30},{wch:26}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Clientes no visitados');XLSX.writeFile(wb,`clientes_no_visitados_${MESES_NOMBRES[CURRENT_MONTH]}_${CURRENT_YEAR}.xlsx`)}

    function regionalClients(){return currentSelectedRegional?Array.from(currentSelectedRegional.asesores.values()).flatMap(a=>Array.from(a.clientes.values())):[]}function switchRegionalTab(tab){['intel','discipline'].forEach(n=>document.getElementById('regional-'+n).classList.toggle('hidden',n!==tab));document.getElementById('regional-advisors-list').classList.toggle('hidden',tab!=='advisors');if(tab==='intel')renderIntel(regionalClients(),'regional');if(tab==='discipline')renderDiscipline(regionalClients(),'regional')}
    function detailSummaryCard(label,value,color,caption){return`<div class="bg-black/50 border rounded-2xl p-4" style="border-color:${color}"><span class="text-[8px] text-slate-400 font-black uppercase">${label}</span><b class="block text-xl mt-2" style="color:${color}">${value}</b><small class="text-[8px] text-slate-500 font-bold uppercase">${caption||''}</small></div>`}
    function renderDisciplineExecutiveSummary(c){
        let agendados=0,visitados=0,lastVisit=-1,currentRun=0,maxRun=0;
        for(let m=0;m<12;m++){agendados+=dv(c,m,'agendado');visitados+=dv(c,m,'visitado');if(dv(c,m,'visitado'))lastVisit=m}
        for(let m=0;m<=CURRENT_MONTH;m++){if(!dv(c,m,'visitado')){currentRun++;maxRun=Math.max(maxRun,currentRun)}else currentRun=0}
        const cumplimiento=agendados>0?visitados/agendados*100:0;
        const ventaMes=financialValue(c,CURRENT_YEAR,CURRENT_MONTH,'venta'),objetivo=financialValue(c,CURRENT_YEAR,CURRENT_MONTH,'objetivoEnfoque'),alcance=objetivo>0?ventaMes/objetivo*100:0;
        const money=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0});
        const ultima=lastVisit>=0?MESES_NOMBRES[lastVisit]+' 2026':'Sin visita registrada';
        const alert=document.getElementById('discipline-detail-alert');
        alert.innerText=currentRun>=2?`Alerta: ${currentRun} meses consecutivos sin visita`:(currentRun===1?'Atención: mes actual sin visita':'Disciplina al corriente');
        alert.className=`text-[9px] font-black uppercase ${currentRun>=2?'text-neon-red':(currentRun===1?'text-neon-orange':'text-neon-green')}`;
        document.getElementById('discipline-detail-summary').innerHTML=
          detailSummaryCard('Meses agendados',agendados,'#ffbb00','Enero a diciembre 2026')+
          detailSummaryCard('visitas realizadas',visitados,'#00ff66','Enero a diciembre 2026')+
          detailSummaryCard('Cumplimiento agenda',cumplimiento.toFixed(1)+'%',cumplimiento>=100?'#00ff66':(cumplimiento>=80?'#ffbb00':'#ff004c'),`${visitados} visitas / ${agendados} agendas`)+
          detailSummaryCard('Última visita',ultima,lastVisit===CURRENT_MONTH?'#00ff66':'#ff004c',lastVisit>=0?`${CURRENT_MONTH-lastVisit} meses desde última visita`:'Requiere atención')+
          detailSummaryCard('Racha sin visita',currentRun+' meses',currentRun>=2?'#ff004c':(currentRun===1?'#ffbb00':'#00ff66'),`Máxima racha: ${maxRun}`)+
          detailSummaryCard('Venta mes actual',money.format(ventaMes),'#00f3ff',`${MESES_NOMBRES[CURRENT_MONTH]} 2026`)+
          detailSummaryCard('Objetivo enfoque',money.format(objetivo),'#a855f7',objetivo>0?`Alcance ${alcance.toFixed(1)}%`:'Sin objetivo capturado')+
          detailSummaryCard('Déficit al objetivo',money.format(Math.max(objetivo-ventaMes,0)),ventaMes>=objetivo&&objetivo>0?'#00ff66':'#ff004c',objetivo>0?(ventaMes>=objetivo?'Objetivo logrado':'Venta adicional requerida'):'No aplica');
    }
    function openDisciplineDetail(cv){const c=Array.from(allClientsFlat.values()).find(x=>String(x.cliVendId).toUpperCase()===String(cv).toUpperCase());if(!c)return;disciplineReturnContext=currentSelectedAdvisor?'advisor':'regional';selectedClient=c;step2.classList.add('hidden');document.getElementById('discipline-client-detail').classList.remove('hidden');document.getElementById('discipline-detail-name').innerText=c.name;document.getElementById('discipline-detail-meta').innerText=`CLIENTE ${c.id} · ${c.ciudad||'SIN CIUDAD'} · ${c.advisor}`;renderDisciplineExecutiveSummary(c);document.getElementById('discipline-detail-months').innerHTML=MESES_NOMBRES.map((m,i)=>`<div class="bg-black/40 border ${dv(c,i,'visitado')?'border-neon-green':(dv(c,i,'agendado')?'border-neon-orange':'border-white/10')} rounded-xl p-3"><div class="flex justify-between items-center"><b class="text-[9px] text-white uppercase">${m}</b><span class="text-[8px] font-black ${dv(c,i,'visitado')?'text-neon-green':(dv(c,i,'agendado')?'text-neon-orange':'text-slate-500')}">${dv(c,i,'visitado')?'VISITADO':(dv(c,i,'agendado')?'NO VISITADO':'SIN AGENDA')}</span></div><div class="text-[8px] mt-2">Agenda ${dv(c,i,'agendado')} · Visita ${dv(c,i,'visitado')}</div></div>`).join('');renderDetail360(c)}function closeDisciplineDetail(){document.getElementById('discipline-client-detail').classList.add('hidden');step2.classList.remove('hidden');if(disciplineReturnContext==='advisor')switchTab('discipline');else switchRegionalTab('discipline')}
    function generateDisciplineOpportunityCatalog(){
        if(!selectedClient){showAlert('NO HAY CLIENTE SELECCIONADO.');return;}
        const products=disciplineOpportunitySkus.map(sku=>globalProducts.get(sku)).filter(Boolean);
        if(!products.length){goToSelectedClientStrategies();return;}
        currentStrategyLabel='Oportunidades por Ciudad';
        document.getElementById('discipline-client-detail').classList.add('hidden');
        renderCatalog(products);
    }
    function goToSelectedClientStrategies(){
        if(!selectedClient){showAlert('NO HAY CLIENTE SELECCIONADO.');return;}
        document.getElementById('discipline-client-detail').classList.add('hidden');
        step2.classList.add('hidden');
        stepStrategy.classList.remove('hidden');
        catalogContainer.innerHTML='';
        hideCartBar();
        manualSkus=[];manualSkusQty={};const excelStatus=document.getElementById('strategy-excel-status');if(excelStatus)excelStatus.innerText='Sin archivo cargado';
        const fmt=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',minimumFractionDigits:0,maximumFractionDigits:0});
        document.getElementById('strategy-client-name').innerText=selectedClient.name;
        document.getElementById('strategy-client-meta').innerText=`ID: ${selectedClient.id} | Asesor: ${selectedClient.advisor} | Venta Mes: ${fmt.format(selectedClient.venta||0)}`;
        renderDashboards(selectedClient);
        appCurrentState='MENU_ESTRATEGIAS';
    }
    function renderDisciplineOpportunityActions(count){
        const box=document.getElementById('discipline-opportunity-actions');
        if(!box)return;
        if(count>0){
            box.innerHTML=`<div class="flex flex-col md:flex-row items-center justify-between gap-4 bg-[#001233]/70 border border-neon-cyan rounded-2xl p-4"><div><div class="text-[9px] text-slate-400 font-black uppercase">Oportunidades detectadas</div><div class="text-lg text-white font-black">${count} códigos sugeridos</div></div><button onclick="generateDisciplineOpportunityCatalog()" class="w-full md:w-auto bg-gradient-to-r from-[#00f3ff] to-[#008cff] text-black px-7 py-3 rounded-xl text-[10px] font-black uppercase shadow-[0_0_20px_rgba(0,243,255,0.35)] hover:scale-105 transition-all">Generar catálogo con estos códigos</button></div>`;
        }else{
            box.innerHTML=`<div class="flex flex-col md:flex-row items-center justify-between gap-4 bg-[#001233]/70 border border-neon-orange rounded-2xl p-4"><div><div class="text-[9px] text-neon-orange font-black uppercase">Sin códigos sugeridos por ciudad</div><div class="text-[10px] text-slate-300 font-bold uppercase mt-1">Continúa con las estrategias comerciales disponibles para este cliente.</div></div><button onclick="goToSelectedClientStrategies()" class="w-full md:w-auto bg-gradient-to-r from-[#ffbb00] to-[#ff7a00] text-black px-7 py-3 rounded-xl text-[10px] font-black uppercase shadow-[0_0_20px_rgba(255,187,0,0.35)] hover:scale-105 transition-all">Ir a estrategias comerciales</button></div>`;
        }
    }
    function renderDetail360(c){
        const brands={},months=Array.from({length:12},()=>new Map()),bought=new Set();c.rawItems.forEach(i=>{const d=i.fecha_doc;if(!d||d.getFullYear()!==2026)return;const sku=getSku(i),v=i._vmn_calculado||0,b=getBrand(i);brands[b]=(brands[b]||0)+v;bought.add(sku);months[d.getMonth()].set(sku,(months[d.getMonth()].get(sku)||0)+v)});if(chartDetailBrands)chartDetailBrands.destroy();const top=Object.entries(brands).sort((a,b)=>b[1]-a[1]).slice(0,10);
        chartDetailBrands=new Chart(document.getElementById('chart-detail-brands'),{type:'bar',data:{labels:top.map(x=>x[0]),datasets:[{data:top.map(x=>x[1]),backgroundColor:'#00f3ff',borderColor:'#ffffff',borderWidth:2,borderRadius:6}]},plugins:[ChartDataLabels],options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},datalabels:{color:'#ffffff',textShadowBlur:6,textShadowColor:'#000000',font:{weight:'900',size:11},formatter:v=>'$'+Intl.NumberFormat('es-MX',{notation:'compact'}).format(v)}},scales:{x:{grid:{color:'rgba(255,255,255,.15)'},ticks:{color:'#ffffff',font:{weight:'bold',size:11},callback:v=>'$'+Intl.NumberFormat('es-MX',{notation:'compact'}).format(v)}},y:{grid:{display:false},ticks:{color:'#ffffff',font:{weight:'bold',size:11}}}}}});
        document.getElementById('discipline-detail-skus').innerHTML=months.map((mp,i)=>`<div class="bg-white/5 rounded-xl p-3"><b class="text-neon-cyan text-[9px] uppercase">${MESES_NOMBRES[i]} · ${mp.size} códigos</b><div class="text-[8px] text-slate-300 mt-1">${[...mp.entries()].sort((a,b)=>b[1]-a[1]).slice(0,15).map(x=>x[0]).join(', ')||'Sin compras'}</div></div>`).join('');const cityClients=Array.from(allClientsFlat.values()).filter(x=>normalizeBrandName(x.ciudad)===normalizeBrandName(c.ciudad)&&x.id!==c.id),opp={};cityClients.forEach(x=>x.rawItems.forEach(i=>{const d=i.fecha_doc,sku=getSku(i);if(d&&d.getFullYear()===2026&&!bought.has(sku))opp[sku]=(opp[sku]||0)+(i._vmn_calculado||0)}));const f=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0});const opportunityEntries=Object.entries(opp).sort((a,b)=>b[1]-a[1]).filter(([sku])=>globalProducts.has(sku)).slice(0,50);disciplineOpportunitySkus=opportunityEntries.map(([sku])=>sku);renderDisciplineOpportunityActions(disciplineOpportunitySkus.length);document.getElementById('discipline-detail-opps').innerHTML=opportunityEntries.map(([sku,v])=>`<div class="bg-white/5 border border-white/10 rounded-xl p-3"><div class="flex justify-between gap-3"><b class="text-neon-cyan">${sku}</b><span class="text-[8px] text-neon-green font-black uppercase">Sugerido</span></div><div class="text-[9px] text-white uppercase truncate">${getDescription(globalProducts.get(sku))}</div><small class="text-slate-400">Venta ciudad 2026: ${f.format(v)}</small></div>`).join('')||'<p class="text-slate-400 text-xs">Sin oportunidades calculables para este cliente.</p>'}
    
    function territorialClients(){
        if(!currentSelectedTerritorial)return [];
        return Array.from(companyTree.values())
            .filter(reg=>reg.territorial===currentSelectedTerritorial)
            .flatMap(reg=>Array.from(reg.asesores.values()))
            .flatMap(adv=>Array.from(adv.clientes.values()));
    }
    function setTerritorialTabButton(active){
        ['intel','discipline','regions'].forEach(name=>{
            const btn=document.getElementById('territorial-tab-'+name);if(!btn)return;
            const on=name===active;
            btn.classList.toggle('bg-[#00f3ff]',on);btn.classList.toggle('text-black',on);
            btn.classList.toggle('bg-[#0f172a]',!on);btn.classList.toggle('text-white',!on);
        });
    }
    function switchTerritorialTab(tab){
        const intel=document.getElementById('territorial-intel'),discipline=document.getElementById('territorial-discipline'),regions=document.getElementById('regional-cards');
        intel.classList.toggle('hidden',tab!=='intel');discipline.classList.toggle('hidden',tab!=='discipline');regions.classList.toggle('hidden',tab!=='regions');
        setTerritorialTabButton(tab);
        const clients=territorialClients();
        if(tab==='intel')renderIntel(clients,'territorial');
        if(tab==='discipline')renderDiscipline(clients,'territorial');
    }
    function territorialNodes(){
        const map=new Map();
        companyTree.forEach(reg=>{
            const name=reg.territorial||"SIN GERENTE TERRITORIAL";
            if(!map.has(name))map.set(name,{name,totalVenta:0,totalCuota:0,regiones:[]});
            const node=map.get(name);node.totalVenta+=reg.totalVenta||0;node.totalCuota+=reg.totalCuota||0;node.regiones.push(reg);
        });
        return Array.from(map.values());
    }
    function renderTerritorialCards(){
        currentSelectedTerritorial=null;
        document.getElementById('directory-title').innerHTML='Directorio de <span class="text-neon-cyan">Territorios</span>';
        document.getElementById('directory-subtitle').innerText='Selecciona un gerente territorial para ver sus regiones';
        document.getElementById('btnBackTerritories').classList.add('hidden');
        document.getElementById('territorial-dashboard').classList.add('hidden');
        document.getElementById('regional-cards').classList.remove('hidden');
        const container=document.getElementById('regional-cards');container.innerHTML='';
        const fmt=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0});
        territorialNodes().filter(node=>!ACCESS_SESSION||ACCESS_SESSION.role!=='TERRITORIAL'||ACCESS_SESSION.entity==='*'||accessNorm(node.name)===accessNorm(ACCESS_SESSION.entity)).sort((a,b)=>b.totalVenta-a.totalVenta).forEach(node=>{
            const proy=DIAS_TRANSCURRIDOS>0?(node.totalVenta/DIAS_TRANSCURRIDOS)*DIAS_EFECTIVOS_MES:0;
            const pct=node.totalCuota>0?proy/node.totalCuota*100:0;
            const div=document.createElement('div');div.className='advisor-card-btn border rounded-2xl p-6 flex flex-col justify-between';div.onclick=()=>openTerritorialRegions(node.name);
            div.innerHTML=`<div class="flex justify-between items-start mb-5"><div><div class="text-[9px] text-slate-400 font-bold uppercase mb-1">Gerente Territorial</div><div class="text-xl font-black text-white uppercase break-words">${node.name}</div><div class="text-[8px] text-neon-cyan font-bold mt-1">${node.regiones.length} REGIÓN(ES)</div></div><div class="traffic-light ${getSemaforoProyeccionClass(pct)}"></div></div><div class="grid grid-cols-2 gap-3"><div><small class="text-slate-400 font-bold">VENTA MES</small><b class="block text-white">${fmt.format(node.totalVenta)}</b></div><div><small class="text-slate-400 font-bold">CUOTA MES</small><b class="block text-white">${fmt.format(node.totalCuota)}</b></div></div><div class="border-t mt-4 pt-4 flex justify-between"><span class="text-[9px] text-slate-400 font-bold">PROYECCIÓN</span><b class="${getSemaforoProyeccionTextClass(pct)}">${pct.toFixed(1)}%</b></div>`;
            container.appendChild(div);
        });
        appCurrentState='PANEL_TERRITORIALES';
    }
    function openTerritorialRegions(name){currentSelectedTerritorial=name;renderRegionalCards();document.getElementById('territorial-dashboard').classList.remove('hidden');switchTerritorialTab('intel');appCurrentState='PANEL_TERRITORIAL'}
    function backToTerritories(){renderTerritorialCards()}
    function renderRegionalCards() {
        const container = document.getElementById('regional-cards'); container.innerHTML = '';
        document.getElementById('directory-title').innerHTML='Regiones de <span class="text-neon-cyan">'+currentSelectedTerritorial+'</span>';
        document.getElementById('directory-subtitle').innerText='Selecciona una región para ver asesores';
        document.getElementById('btnBackTerritories').classList.remove('hidden');
        const fmt = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 });
        Array.from(companyTree.values()).filter(reg=>(!currentSelectedTerritorial||reg.territorial===currentSelectedTerritorial)&&(!ACCESS_SESSION||ACCESS_SESSION.role!=='REGIONAL'||accessNorm(reg.name)===accessNorm(ACCESS_SESSION.entity))).sort((a,b)=> b.totalVenta - a.totalVenta).forEach(reg => {
            const proyeccion = DIAS_TRANSCURRIDOS > 0 ? (reg.totalVenta / DIAS_TRANSCURRIDOS) * DIAS_EFECTIVOS_MES : 0;
            const porcentaje = reg.totalCuota > 0 ? (proyeccion / reg.totalCuota) * 100 : 0;
            const div = document.createElement('div');
            div.className = "advisor-card-btn border border-white/5 rounded-2xl p-6 flex flex-col justify-between";
            div.onclick = () => openRegionalAdvisors(reg.name);
            
            let zoneTitle = reg.name !== "SIN GERENTE" ? "Gerente Regional" : "Zonas sin Asignar";
            
            div.innerHTML = `
                <div class="flex justify-between items-start mb-5">
                    <div><div class="text-[9px] text-slate-400 font-bold uppercase tracking-[0.1em] mb-1">${zoneTitle}</div><div class="text-xl font-black text-white uppercase break-words">${reg.name}</div></div>
                    <div class="traffic-light ${getSemaforoProyeccionClass(porcentaje)} flex-shrink-0"></div>
                </div>
                <div class="grid grid-cols-2 gap-3 mb-4">
                    <div><div class="text-[8px] text-slate-400 font-bold uppercase">Venta Neta Mes</div><div class="text-[13px] font-black text-white">${fmt.format(reg.totalVenta)}</div></div>
                    <div><div class="text-[8px] text-slate-400 font-bold uppercase">Meta Acumulada</div><div class="text-[13px] font-black text-white">${fmt.format(reg.totalCuota)}</div></div>
                </div>
                <div class="border-t border-white/10 pt-4 flex justify-between items-center"><div class="text-[9px] text-slate-400 font-bold uppercase">Proyección de Cierre</div><div class="text-lg font-black ${getSemaforoProyeccionTextClass(porcentaje)}">${porcentaje.toFixed(1)}%</div></div>`;
            container.appendChild(div);
        });
    }

    function renderAdvisorCards() {
        const container = document.getElementById('advisor-cards'); container.innerHTML = '';
        if (!currentSelectedRegional) return;
        const fmt = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 });
        Array.from(currentSelectedRegional.asesores.values()).filter(adv=>!ACCESS_SESSION||ACCESS_SESSION.role!=='ASESOR'||accessNorm(adv.name)===accessNorm(ACCESS_SESSION.entity)).sort((a,b)=>b.totalVenta - a.totalVenta).forEach(adv => {
            const proyeccion = DIAS_TRANSCURRIDOS > 0 ? (adv.totalVenta / DIAS_TRANSCURRIDOS) * DIAS_EFECTIVOS_MES : 0;
            const porcentaje = adv.totalCuota > 0 ? (proyeccion / adv.totalCuota) * 100 : 0;
            const div = document.createElement('div');
            div.className = "advisor-card-btn border border-white/5 rounded-2xl p-6 flex flex-col justify-between";
            div.onclick = () => openAdvisorClients(adv.name);
            div.innerHTML = `
                <div class="flex justify-between items-start mb-5">
                    <div>
                        <div class="text-[9px] text-slate-400 font-bold uppercase mb-1">Asesor Comercial</div>
                        <div class="text-xl font-black text-white uppercase">${adv.name}</div>
                        <div class="text-[8px] text-neon-cyan font-bold uppercase mt-1">Sucursal: ${adv.sucursal || 'N/A'}</div>
                    </div>
                    <div class="traffic-light ${getSemaforoProyeccionClass(porcentaje)}"></div>
                </div>
                <div class="grid grid-cols-2 gap-3 mb-4">
                    <div><div class="text-[8px] text-slate-400 font-bold uppercase">Venta</div><div class="text-[13px] font-black text-white">${fmt.format(adv.totalVenta)}</div></div>
                    <div><div class="text-[8px] text-slate-400 font-bold uppercase">Objetivo</div><div class="text-[13px] font-black text-white">${fmt.format(adv.totalCuota)}</div></div>
                </div>
                <div class="border-t border-white/10 pt-4 flex justify-between items-center"><div class="text-[9px] text-slate-400 font-bold uppercase">Proyección</div><div class="text-lg font-black ${getSemaforoProyeccionTextClass(porcentaje)}">${porcentaje.toFixed(1)}%</div></div>`;
            container.appendChild(div);
        });
    }

    function renderClientList() {
        const container = document.getElementById('client-list'); container.innerHTML = '';
        if (!currentSelectedAdvisor) return;
        const search = document.getElementById('clientSearch').value.toLowerCase().trim();
        const fmt = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 });
        Array.from(currentSelectedAdvisor.clientes.values()).filter(c => c.id.toLowerCase().includes(search) || c.name.toLowerCase().includes(search)).sort((a,b)=>b.venta - a.venta).forEach(client => {
            const proyeccion = DIAS_TRANSCURRIDOS > 0 ? (client.venta / DIAS_TRANSCURRIDOS) * DIAS_EFECTIVOS_MES : 0;
            const porcentaje = client.cuota > 0 ? (proyeccion / client.cuota) * 100 : 0;
            
            let focusColor = "bg-white/5 text-white";
            if (client.focus === "URGENTE:SIN VENTA") focusColor = "bg-[#ff004c]/20 text-[#ff004c] border border-[#ff004c]/50 animate-pulse";
            
            const div = document.createElement('div');
            div.className = "p-5 bg-[#050b14]/80 border-2 border-white/5 rounded-2xl hover:border-neon-cyan cursor-pointer flex flex-col md:flex-row justify-between items-center gap-4 group";
            div.onclick = () => selectClient(client.uniqueKey);
            div.innerHTML = `
                <div class="flex-grow w-full">
                    <div class="client-identity-row"><span class="client-segment-badge">${client.focus || 'GENERAL'}</span><span class="client-id-badge">CLIENTE ID: ${client.id}</span></div>
                    <div class="text-xl font-black text-white uppercase group-hover:text-neon-cyan">${client.name}</div>
                    <div class="text-[10px] text-slate-400 font-bold uppercase mt-1"><span class="text-neon-cyan">📍</span> ${client.ciudad || 'CIUDAD NO ESPECIFICADA'}</div>
                    <div class="grid grid-cols-3 gap-4 mt-3 max-w-2xl bg-white/5 p-3 rounded-xl">
                        <div><span class="text-[8px] text-slate-400 font-bold block uppercase">Venta Neta Mes</span><span class="text-sm font-black text-white">${fmt.format(client.venta)}</span></div>
                        <div><span class="text-[8px] text-slate-400 font-bold block uppercase">Objetivo Mn</span><span class="text-sm font-black text-white">${fmt.format(client.cuota)}</span></div>
                        <div><span class="text-[8px] text-slate-400 font-bold block uppercase">Proyección</span><span class="text-sm font-black text-neon-cyan">${fmt.format(proyeccion)}</span></div>
                    </div>
                </div>
                <div class="flex items-center gap-5 flex-shrink-0 w-full md:w-auto justify-between md:justify-end">
                    <div class="text-right"><span class="text-[9px] text-slate-400 font-bold block uppercase">Proyección</span><span class="text-2xl font-black text-white">${porcentaje.toFixed(1)}%</span></div>
                    <div class="traffic-light ${getSemaforoProyeccionClass(porcentaje)} scale-125"></div>
                </div>`;
            container.appendChild(div);
        });
    }

    function switchTab(tabName){['intel','discipline','clientes','bonos'].forEach(n=>{const p=document.getElementById('tab-content-'+n);if(p)p.classList.toggle('hidden',n!==tabName)});if(tabName==='intel'&&currentSelectedAdvisor)renderIntel(Array.from(currentSelectedAdvisor.clientes.values()),'advisor');if(tabName==='discipline'&&currentSelectedAdvisor)renderDiscipline(Array.from(currentSelectedAdvisor.clientes.values()),'advisor');if(tabName==='clientes')renderClientList();if(tabName==='bonos')renderBonusView();}
    function renderBonusView() {
        const container = document.getElementById('bonus-cards-container');
        container.innerHTML = '';
        if (!currentSelectedAdvisor) return;

        const fmt = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 });
        const ventaAcumulada = currentSelectedAdvisor.totalVenta;
        const cuotaTotal = currentSelectedAdvisor.totalCuota;

        const metasSemanales = [
            { nombre: 'Semana 1', corte: 'Corte 1 al 11 de Julio', obj: 0.36 },
            { nombre: 'Semana 2', corte: 'Corte 1 al 18 de Julio', obj: 0.58 },
            { nombre: 'Semana 3', corte: 'Corte 1 al 25 de Julio', obj: 0.80 }
        ];

        const clientes = Array.from(currentSelectedAdvisor.clientes.values());
        metasSemanales.forEach((meta, semanaIndex) => {
            const registrosSemana = clientes.map(c => c.bonoSemanas?.[semanaIndex] || {capturado:false, valor:0});
            const semanaCerrada = registrosSemana.some(r => r.capturado);
            const ventaBono = semanaCerrada
                ? registrosSemana.reduce((total, r) => total + (r.capturado ? Number(r.valor || 0) : 0), 0)
                : ventaAcumulada;
            const metaSemanal = cuotaTotal * meta.obj;
            const deficit = metaSemanal > ventaBono ? metaSemanal - ventaBono : 0;
            const porcentajeAvance = metaSemanal > 0 ? (ventaBono / metaSemanal) * 100 : 0;
            const logrado = ventaBono >= metaSemanal;

            const colorBorder = logrado ? 'border-neon-green' : 'border-neon-red';
            const colorText = logrado ? 'text-neon-green' : 'text-neon-red';
            const shadow = logrado ? 'shadow-[0_0_15px_rgba(0,255,102,0.2)]' : 'shadow-[0_0_15px_rgba(255,0,76,0.2)]';

            container.innerHTML += `
                <div class="bg-black/60 border ${colorBorder} rounded-2xl p-6 ${shadow} flex flex-col justify-between relative overflow-hidden group hover:scale-105 transition-transform">
                    ${semanaCerrada ? `<div class="absolute -right-8 top-4 ${logrado ? 'bg-neon-green' : 'bg-neon-red'} text-black text-[9px] font-black uppercase py-1 px-10 rotate-45 shadow-lg">${logrado ? 'CERRADO · LOGRADO' : 'CERRADO · NO LOGRADO'}</div>` : `<div class="absolute -right-8 top-4 bg-neon-orange text-black text-[9px] font-black uppercase py-1 px-10 rotate-45 shadow-lg">EN CURSO</div>`}
                    
                    <div class="mb-4 border-b border-white/10 pb-4">
                        <div class="text-[10px] text-slate-400 font-bold uppercase tracking-widest">${meta.corte}</div>
                        <div class="text-2xl font-black text-white uppercase">${meta.nombre} <span class="text-lg text-neon-orange">(${Number((meta.obj * 100).toFixed(1)).toLocaleString('es-MX')}%)</span></div>
                    </div>

                    <div class="grid grid-cols-2 gap-4 mb-5">
                        <div>
                            <span class="text-[8px] text-slate-500 font-bold block uppercase">Meta Semanal</span>
                            <span class="text-sm font-black text-white">${fmt.format(metaSemanal)}</span>
                        </div>
                        <div>
                            <span class="text-[8px] text-slate-500 font-bold block uppercase">Avance (Venta Real)</span>
                            <span class="text-sm font-black text-white">${fmt.format(ventaBono)}</span>
                            <span class="text-[7px] ${semanaCerrada ? 'text-neon-green' : 'text-neon-orange'} font-black block uppercase mt-1">${semanaCerrada ? 'Valor cerrado desde bono_semana_' + (semanaIndex + 1) : 'Valor provisional de venta mensual'}</span>
                        </div>
                        <div class="col-span-2 bg-white/5 p-3 rounded-lg border border-white/5">
                            <span class="text-[9px] text-slate-500 font-bold block uppercase">Déficit para Lograr Bono</span>
                            <span class="text-xl font-black ${colorText}">${fmt.format(deficit)}</span>
                        </div>
                    </div>

                    <div>
                        <div class="flex justify-between items-center mb-1">
                            <span class="text-[9px] text-slate-400 font-bold uppercase">Progreso vs Semana</span>
                            <span class="text-[11px] font-black ${colorText}">${porcentajeAvance.toFixed(1)}%</span>
                        </div>
                        <div class="w-full bg-white/10 rounded-full h-2 overflow-hidden">
                            <div class="${logrado ? 'bg-neon-green' : 'bg-neon-red'} h-2 rounded-full" style="width: ${Math.min(porcentajeAvance, 100)}%"></div>
                        </div>
                    </div>
                </div>
            `;
        });
    }

    function openRegionalAdvisors(name) { if(!canOpenRegional(name)){showAlert('ACCESO RESTRINGIDO A TU REGIÓN.');return;} currentSelectedRegional = companyTree.get(name); if(!currentSelectedRegional) return; document.getElementById('regional-view').classList.add('hidden'); document.getElementById('advisor-view').classList.remove('hidden'); document.getElementById('current-regional-title').innerText = name; appCurrentState = "PANEL_ASESORES"; renderAdvisorCards(); switchRegionalTab('intel'); }
    function backToRegionals() { if(ACCESS_SESSION&&ACCESS_SESSION.role!=='TERRITORIAL'){applyAccessLanding();return;} currentSelectedRegional = null; document.getElementById('advisor-view').classList.add('hidden'); document.getElementById('regional-view').classList.remove('hidden'); renderRegionalCards(); appCurrentState = "PANEL_REGIONALES"; }
    
    function openAdvisorClients(name) { 
        if(!currentSelectedRegional) return; 
        currentSelectedAdvisor = currentSelectedRegional.asesores.get(name); 
        if(!currentSelectedAdvisor) return; 
        
        document.getElementById('advisor-view').classList.add('hidden'); 
        document.getElementById('client-view').classList.remove('hidden'); 
        document.getElementById('current-advisor-title').innerText = name; 
        
        switchTab('intel');
        appCurrentState = "PANEL_ASESOR"; 
    }
    
    function backToAdvisors() { if(ACCESS_SESSION&&ACCESS_SESSION.role==='ASESOR'){currentSelectedAdvisor=null;document.getElementById('client-view').classList.add('hidden');document.getElementById('advisor-view').classList.remove('hidden');renderAdvisorCards();switchRegionalTab('advisors');appCurrentState='INICIO_ASESOR';return;} currentSelectedAdvisor = null; document.getElementById('client-view').classList.add('hidden'); document.getElementById('advisor-view').classList.remove('hidden'); appCurrentState = "PANEL_ASESORES"; }

    function selectClient(id) {
        if(!currentSelectedAdvisor) return; selectedClient = currentSelectedAdvisor.clientes.get(id); if(!selectedClient) return;
        step2.classList.add('hidden'); stepStrategy.classList.remove('hidden'); catalogContainer.innerHTML = ''; hideCartBar(); 
        
        manualSkus = []; 
        manualSkusQty = {}; 

        const fmt = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 });
        document.getElementById('strategy-client-name').innerText = selectedClient.name;
        document.getElementById('strategy-client-meta').innerText = `CLIENTE ID: ${selectedClient.id} | Asesor: ${selectedClient.advisor} | Venta Mes: ${fmt.format(selectedClient.venta || 0)}`;
        renderDashboards(selectedClient); appCurrentState = "MENU_ESTRATEGIAS";
    }

    function renderDashboards(client) {
        let brandSales = {};
        client.rawItems.forEach(item => {
            let brand = normalizeBrandName(getBrand(item)), venta = item._vmn_calculado || 0;
            if(venta > 0) { if(!brandSales[brand]) brandSales[brand] = 0; brandSales[brand] += venta; }
        });
        let sortedBrands = Object.entries(brandSales).sort((a,b) => b[1] - a[1]);
        let topN = sortedBrands.slice(0, 6);
        let brandLabels = topN.map(b => b[0]);
        let brandData = topN.map(b => b[1]);

        let m1 = (CURRENT_MONTH - 1 + 12) % 12, y1 = CURRENT_MONTH - 1 < 0 ? CURRENT_YEAR - 1 : CURRENT_YEAR;
        let m2 = (CURRENT_MONTH - 2 + 12) % 12, y2 = CURRENT_MONTH - 2 < 0 ? CURRENT_YEAR - 1 : CURRENT_YEAR;
        let m3 = (CURRENT_MONTH - 3 + 12) % 12, y3 = CURRENT_MONTH - 3 < 0 ? CURRENT_YEAR - 1 : CURRENT_YEAR;

        let targets = [{ m: m3, y: y3, label: MESES_NOMBRES[m3], value: 0 }, { m: m2, y: y2, label: MESES_NOMBRES[m2], value: 0 }, { m: m1, y: y1, label: MESES_NOMBRES[m1], value: 0 }];
        client.rawItems.forEach(item => {
            let d = item.fecha_doc;
            if (d) { let target = targets.find(t => t.m === d.getMonth() && t.y === d.getFullYear()); if (target) target.value += (item._vmn_calculado || 0); }
        });

        if(chartBrandsInstance) chartBrandsInstance.destroy();
        if(chartTrendInstance) chartTrendInstance.destroy();
        
        Chart.defaults.color = '#8ab4f8'; 
        Chart.defaults.font.family = 'Montserrat';

        // 1. TOP MARCAS: Anillo Neón con texto negro sobrio
        const neonColors = ['#00f3ff', '#ff004c', '#39ff14', '#e6ff00', '#b026ff', '#00ffcc'];
        const neonBg = ['rgba(0,243,255,0.15)', 'rgba(255,0,76,0.15)', 'rgba(57,255,20,0.15)', 'rgba(230,255,0,0.15)', 'rgba(176,38,255,0.15)', 'rgba(0,255,204,0.15)'];

        const ctxBrands = document.getElementById('chart-brands').getContext('2d');
        chartBrandsInstance = new Chart(ctxBrands, { 
            type: 'doughnut', 
            data: { 
                labels: brandLabels, 
                datasets: [{ 
                    data: brandData, 
                    backgroundColor: neonBg, 
                    borderColor: neonColors, 
                    borderWidth: 2,
                    hoverOffset: 15
                }] 
            }, 
            plugins: [ChartDataLabels], 
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                cutout: '80%', 
                layout: { padding: 15 },
                plugins: { 
                    legend: { position: 'right', labels: { color: '#ffffff', boxWidth: 10, font: { size: 10, weight: '900' } } }, 
                    datalabels: { 
                        color: '#050505', // NEGRO SOBRIO
                        backgroundColor: function(context) { return context.dataset.borderColor[context.dataIndex]; },
                        borderRadius: 4,
                        padding: 4,
                        font: { weight: '900', size: 10 }, 
                        formatter: v => '$'+Intl.NumberFormat('es-MX',{notation:'compact'}).format(v),
                        display: 'auto'
                    } 
                } 
            } 
        });

        // 2. TENDENCIA: Curva Holográfica con texto negro sobrio
        const ctxTrend = document.getElementById('chart-trend').getContext('2d');
        let gradient = ctxTrend.createLinearGradient(0, 0, 0, 200);
        gradient.addColorStop(0, 'rgba(0, 243, 255, 0.5)'); 
        gradient.addColorStop(1, 'rgba(0, 243, 255, 0.0)'); 

        chartTrendInstance = new Chart(ctxTrend, { 
            type: 'line', 
            data: { 
                labels: targets.map(t=>t.label), 
                datasets: [{ 
                    label: 'Venta $', 
                    data: targets.map(t=>t.value), 
                    backgroundColor: gradient, 
                    borderColor: '#00f3ff', 
                    borderWidth: 3, 
                    fill: true,
                    tension: 0.4, 
                    pointBackgroundColor: '#000', 
                    pointBorderColor: '#00f3ff', 
                    pointBorderWidth: 2,
                    pointRadius: 6,
                    pointHoverRadius: 8,
                    pointHoverBackgroundColor: '#00f3ff'
                }] 
            }, 
            plugins: [ChartDataLabels], 
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                layout: { padding: { top: 30 } },
                scales: { 
                    y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)', tickLength: 0 }, ticks: { display: false } }, 
                    x: { grid: { display: false }, ticks: { color: '#ffffff', font: {weight: '900', size: 11} } } 
                }, 
                plugins: { 
                    legend: { display: false }, 
                    tooltip: { backgroundColor: 'rgba(0,0,0,0.8)', titleColor: '#00f3ff', bodyColor: '#fff', borderColor: '#00f3ff', borderWidth: 1 },
                    datalabels: { 
                        color: '#050505', // NEGRO SOBRIO
                        backgroundColor: '#00f3ff', // FONDO NEON
                        borderRadius: 4,
                        padding: { top: 4, bottom: 4, left: 6, right: 6 },
                        align: 'top',
                        offset: 8,
                        font: { weight: '900', size: 12 }, 
                        formatter: v => v > 0 ? new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', maximumFractionDigits: 0 }).format(v) : '' 
                    } 
                } 
            } 
        });
    }

    function confirmExitToLoader(){return window.confirm('¿Deseas salir de la sesión actual y volver a la pantalla de carga?\n\nSi continúas, deberás autenticarte nuevamente. La información cargada se cerrará.');}
    function exitToLoader(){if(!confirmExitToLoader())return false;step2.classList.add('hidden');stepStrategy.classList.add('hidden');document.getElementById('discipline-client-detail').classList.add('hidden');step1.classList.remove('hidden');btnBack.classList.add('hidden');fileInput.value='';ACCESS_SESSION=null;appCurrentState='LOAD';return true;}
    function backToSelection() {
        if (appCurrentState === "CATALOGO_LISTO") { catalogContainer.innerHTML = ''; stepStrategy.classList.remove('hidden'); btnPrint.setAttribute('disabled', 'true'); closeCartDrawer(); appCurrentState = "MENU_ESTRATEGIAS"; }
        else if (appCurrentState === "MENU_ESTRATEGIAS") { stepStrategy.classList.add('hidden'); step2.classList.remove('hidden'); hideCartBar(); if(ACCESS_SESSION&&ACCESS_SESSION.role==='ASESOR'){switchTab('intel');appCurrentState='PANEL_ASESOR';}else{switchTab('clientes');appCurrentState='LISTA_CLIENTES';} }
        else if (appCurrentState === 'INICIO_ASESOR' || appCurrentState === 'INICIO_REGIONAL') { exitToLoader(); }
        else if (appCurrentState === "LISTA_CLIENTES") backToAdvisors();
        else if (appCurrentState === "PANEL_ASESORES") backToRegionals();
        else if (appCurrentState === "PANEL_REGIONALES") { if(currentSelectedTerritorial){document.getElementById('advisor-view').classList.add('hidden');document.getElementById('regional-view').classList.remove('hidden');document.getElementById('territorial-dashboard').classList.remove('hidden');switchTerritorialTab('intel');appCurrentState='PANEL_TERRITORIAL';}else backToTerritories(); }
        else if (appCurrentState === "PANEL_TERRITORIAL") { backToTerritories(); }
        else if (appCurrentState === "PANEL_TERRITORIALES") { exitToLoader(); }
    }

    function getLastPriceForClient(sku, clientObj) {
        let lastPrice = null;
        if (clientObj && clientObj.rawItems) {
            const historyRecords = clientObj.rawItems.filter(raw => getSku(raw).toUpperCase() === sku.toUpperCase());
            let sortedHistory = [...historyRecords].sort((a,b)=> {
                let da = a.fecha_doc;
                let db = b.fecha_doc;
                let ta = da ? da.getTime() : 0;
                let tb = db ? db.getTime() : 0;
                return ta - tb;
            });
            for (let k = sortedHistory.length - 1; k >= 0; k--) {
                let pz = sortedHistory[k]._venta_piezas_calculada;
                let mn = sortedHistory[k]._vmn_calculado;
                if (pz > 0 && mn > 0) { 
                    lastPrice = mn / pz; 
                    break; 
                }
            }
        }
        return lastPrice;
    }

    function getPrecioSugerido(sku, tipoCli) {
        if (!globalSkuStats[sku]) return 0;
        let stats = globalSkuStats[sku][tipoCli || "GENERAL"];
        if (!stats) return 0;
        if (stats.pz > 0) return stats.mn / stats.pz;
        return 0;
    }
    
    function openManualModal() {
        toggleBackgroundLock(true);
        document.getElementById('manual-modal').classList.remove('hidden');
        renderManualSelectedList();
    }

    function closeManualModal() {
        toggleBackgroundLock(false);
        document.getElementById('manual-modal').classList.add('hidden');
        document.getElementById('manual-search-input').value = '';
        document.getElementById('manual-search-results').classList.add('hidden');
    }

    function handleManualSearch() {
        const query = document.getElementById('manual-search-input').value.toLowerCase().trim();
        const resultsBox = document.getElementById('manual-search-results');
        resultsBox.innerHTML = '';
        
        if (query.length < 2) { resultsBox.classList.add('hidden'); return; }
        
        let count = 0;
        for (let [sku, item] of globalProducts.entries()) {
            const desc = getDescription(item).toLowerCase();
            if (sku.toLowerCase().includes(query) || desc.includes(query)) {
                if(!manualSkus.includes(sku)) {
                    const div = document.createElement('div');
                    div.className = "p-3 hover:bg-neon-green hover:text-black cursor-pointer border-b border-white/10 text-[11px] font-bold text-slate-300 transition-colors flex justify-between";
                    div.innerHTML = `<span class="truncate pr-4">${sku} - ${getDescription(item)}</span> <span class="font-black text-lg leading-none">+</span>`;
                    div.onclick = () => {
                        manualSkus.push(sku);
                        manualSkusQty[sku] = 1; 
                        document.getElementById('manual-search-input').value = '';
                        resultsBox.classList.add('hidden');
                        renderManualSelectedList();
                        document.getElementById('manual-search-input').focus();
                    };
                    resultsBox.appendChild(div);
                    count++;
                    if (count > 30) break; 
                }
            }
        }
        
        if (count > 0) resultsBox.classList.remove('hidden');
        else resultsBox.classList.add('hidden');
    }

    function updateManualQty(sku, qty) {
        let val = parseInt(qty, 10);
        if (isNaN(val) || val < 1) val = 1; 
        manualSkusQty[sku] = val;
    }

    function removeManualSku(sku) {
        manualSkus = manualSkus.filter(s => s !== sku);
        delete manualSkusQty[sku];
        renderManualSelectedList();
    }

    function clearManualSkus() {
        if(manualSkus.length === 0) return;
        if(confirm("¿Estás seguro de vaciar la cotización manual?")) {
            manualSkus = [];
            manualSkusQty = {};
            renderManualSelectedList();
        }
    }

    function renderManualSelectedList() {
        const container = document.getElementById('manual-selected-list');
        container.innerHTML = '';
        
        if (manualSkus.length === 0) {
            container.innerHTML = '<div class="text-center text-slate-500 font-bold uppercase mt-10 text-[10px]">No hay códigos agregados a la lista</div>';
            return;
        }

        const fmt = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 });
        
        [...manualSkus].reverse().forEach(sku => {
            const prod = globalProducts.get(sku);
            
            let ultimoPrecio = getLastPriceForClient(sku, selectedClient);
            let precioPromo = globalPromociones[sku] || 0;
            let comentarioPromo = globalComentarioPromo[sku] || "";
            let psug = getPrecioSugerido(sku, selectedClient.tipoCliente);

            let txtComentario = "";
            if (comentarioPromo !== "") txtComentario = comentarioPromo;
            else if (ultimoPrecio !== null) txtComentario = "Precio Especial";
            else txtComentario = "No aplica";

            let txtSolicitado = precioPromo > 0 ? fmt.format(precioPromo) : "sin promoción";

            let txtSugerido = "No aplica";
            if (precioPromo === 0 && ultimoPrecio === null) {
                txtSugerido = psug > 0 ? fmt.format(psug) : "Sin historial global";
            }

            let txtUltimo = ultimoPrecio !== null ? fmt.format(ultimoPrecio) : "sin historial";

            container.innerHTML += `
                <div class="bg-black/60 border border-white/10 rounded-xl p-4 flex justify-between items-start group hover:border-neon-green transition-colors relative">
                    <button onclick="removeManualSku('${sku}')" class="absolute top-2 right-2 text-slate-500 hover:text-neon-red p-1 bg-white/5 rounded-full transition-colors border border-transparent hover:border-neon-red">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
                    </button>
                    
                    <div class="flex-grow w-full">
                        <div class="text-neon-green font-black text-sm mb-1">${sku}</div>
                        <div class="text-[9.5px] text-slate-300 uppercase font-bold truncate mb-3 pr-6">${getDescription(prod)}</div>
                        
                        <div class="flex items-center gap-2 mb-3">
                            <span class="text-[8px] text-neon-cyan font-black uppercase bg-[#001233] px-2 py-1 rounded border border-neon-cyan/30">Piezas:</span>
                            <input type="number" min="1" value="${manualSkusQty[sku] || 1}" onchange="updateManualQty('${sku}', this.value)" class="w-12 bg-[#030712] border border-neon-cyan/50 text-white text-[11px] font-black outline-none text-center rounded p-1 custom-scrollbar focus:border-neon-cyan">
                        </div>
                        
                        <div class="grid grid-cols-2 md:grid-cols-4 gap-2">
                            <div class="bg-black/40 border border-white/5 p-2 rounded-lg text-center flex flex-col justify-center min-h-[50px]">
                                <div class="text-[7px] text-slate-400 font-bold uppercase mb-1">Comentario</div>
                                <div class="text-[9px] text-neon-cyan font-black leading-tight">${txtComentario}</div>
                            </div>
                            <div class="bg-black/40 border border-white/5 p-2 rounded-lg text-center flex flex-col justify-center min-h-[50px]">
                                <div class="text-[7px] text-slate-400 font-bold uppercase mb-1">P. Solicitado</div>
                                <div class="text-[9px] text-white font-black leading-tight">${txtSolicitado}</div>
                            </div>
                            <div class="bg-black/40 border border-white/5 p-2 rounded-lg text-center flex flex-col justify-center min-h-[50px]">
                                <div class="text-[7px] text-slate-400 font-bold uppercase mb-1">P. Sugerido</div>
                                <div class="text-[9px] text-neon-orange font-black leading-tight">${txtSugerido}</div>
                            </div>
                            <div class="bg-black/40 border border-white/5 p-2 rounded-lg text-center flex flex-col justify-center min-h-[50px]">
                                <div class="text-[7px] text-slate-400 font-bold uppercase mb-1">Último Vendido</div>
                                <div class="text-[9px] text-neon-green font-black leading-tight">${txtUltimo}</div>
                            </div>
                        </div>
                    </div>
                </div>`;
        });
    }

    function exportManualExcel() {
        if (manualSkus.length === 0) { showAlert("No hay códigos capturados."); return; }
        const excelData = [];
        manualSkus.forEach(sku => {
            const ultimoPrecio = getLastPriceForClient(sku, selectedClient);
            const precioPromo = Number(globalPromociones[sku] || 0);
            const promedioTipoCliente = Number(getPrecioSugerido(sku, selectedClient.tipoCliente) || 0);
            const rawComentario = String(globalComentarioPromo[sku] || globalPromoText[sku] || '').trim();
            const comentarioValido = rawComentario && !['0','0.0','0.00','$0','$0.00','null','undefined','n/a','na'].includes(rawComentario.toLowerCase());

            let precioSolicitado = '';
            if (precioPromo > 0) precioSolicitado = Number(precioPromo.toFixed(2));
            else if (ultimoPrecio !== null && ultimoPrecio > 0) precioSolicitado = Number(ultimoPrecio.toFixed(2));
            else if (promedioTipoCliente > 0) precioSolicitado = Number(promedioTipoCliente.toFixed(2));

            let comentario = comentarioValido ? rawComentario : 'Precio Especial';
            if (comentario === '0' || Number(comentario) === 0) comentario = 'Precio Especial';

            excelData.push({
                "Producto": sku,
                "Cantidad": manualSkusQty[sku] || 1,
                "Precio solicitado": precioSolicitado,
                "Comentario": comentario
            });
        });

        const ws = XLSX.utils.json_to_sheet(excelData, { header: ["Producto", "Cantidad", "Precio solicitado", "Comentario"] });
        const wb = XLSX.utils.book_new();
        ws['!cols'] = [{wch: 20}, {wch: 12}, {wch: 20}, {wch: 40}];
        const headerStyle = { font: { color: { rgb: "FFFFFF" }, bold: true }, fill: { fgColor: { rgb: "17365D" } }, alignment: { horizontal: "center" } };
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const address = XLSX.utils.encode_col(C) + "1";
            if (ws[address]) ws[address].s = headerStyle;
        }
        XLSX.utils.book_append_sheet(wb, ws, "Cotización Manual");
        XLSX.writeFile(wb, `CotizacionManual_${selectedClient.id}_${new Date().getTime()}.xlsx`);
    }

    function printManualCatalog() {
        if (manualSkus.length === 0) { showAlert("No hay códigos capturados."); return; }
        closeManualModal();
        generateCatalogByStrategy('manual');
    }

    async function handleStrategyExcel(event) {
        const input=event.target,file=input.files&&input.files[0],status=document.getElementById('strategy-excel-status');
        if(!file)return;
        try{
            status.innerText='Leyendo códigos del Excel...';
            const wb=XLSX.read(await file.arrayBuffer(),{type:'array',raw:false});
            const firstSheet=wb.SheetNames[0];
            if(!firstSheet)throw new Error('El Excel no contiene hojas.');
            const matrix=XLSX.utils.sheet_to_json(wb.Sheets[firstSheet],{header:1,defval:'',raw:false,blankrows:false});
            if(!matrix.length)throw new Error('La primera hoja está vacía.');
            const nonEmptyColumns=new Set();
            matrix.forEach(row=>row.forEach((value,index)=>{if(String(value||'').trim())nonEmptyColumns.add(index)}));
            if(nonEmptyColumns.size>1)throw new Error('El Excel debe contener una sola columna con códigos.');
            const columnIndex=nonEmptyColumns.size?[...nonEmptyColumns][0]:0;
            let values=matrix.map(row=>String(row[columnIndex]||'').trim()).filter(Boolean);
            if(!values.length)throw new Error('No se encontraron códigos.');
            const headerNorm=normalizeKey(values[0]);
            if(['codigo','codigoproducto','producto','sku','codigodelproducto'].includes(headerNorm))values.shift();
            const uniqueCodes=[...new Set(values.map(x=>x.toUpperCase()))];
            const found=[],missing=[];
            uniqueCodes.forEach(code=>{
                const product=globalProducts.get(code)||Array.from(globalProducts.entries()).find(([sku])=>String(sku).trim().toUpperCase()===code)?.[1];
                if(product)found.push(product);else missing.push(code);
            });
            if(!found.length)throw new Error('Ningún código del Excel se encontró en Productos.');
            currentStrategyLabel='Carga Excel · '+found.length+' códigos';
            status.innerText=`${found.length} códigos encontrados${missing.length?' · '+missing.length+' no encontrados':''}`;
            if(missing.length)showAlert(`${missing.length} CÓDIGOS DEL EXCEL NO SE ENCONTRARON. SE GENERARÁ EL LISTADO CON ${found.length}.`);
            renderCatalog(found);
        }catch(err){
            console.error('Error en carga Excel:',err);
            status.innerText='Error: '+(err&&err.message?err.message:'No se pudo leer el Excel');
            showAlert(status.innerText.toUpperCase());
        }finally{
            input.value='';
        }
    }
    function generateCatalogByStrategy(strategyType) {
        if (!selectedClient) { showAlert("No hay cliente seleccionado."); return; }
        let targetProducts = [];
        const sortedGlobalSkusMn = Object.keys(globalSalesMn).sort((a, b) => (globalSalesMn[b] || 0) - (globalSalesMn[a] || 0));
        const sortedGlobalSkusPesos = Object.keys(globalSalesPesos).sort((a, b) => (globalSalesPesos[b] || 0) - (globalSalesPesos[a] || 0));
        
        if (strategyType === 'manual') {
            if (manualSkus.length === 0) { showAlert("No has capturado ningún código."); return; }
            currentStrategyLabel = "Cotización Manual";
            targetProducts = [...manualSkus].reverse().map(sku => globalProducts.get(sku)).filter(Boolean);
        } else if (strategyType === 'skuPorMes') {
            const targetMonth = parseInt(document.getElementById('sku-month-selector').value, 10);
            const targetYear = parseInt(document.getElementById('sku-year-selector').value, 10);
            const scope = document.getElementById('sku-scope-selector').value;
            currentStrategyLabel = `Top ${MESES_NOMBRES[targetMonth]} ${targetYear} (${scope === 'client' ? 'CLIENTE' : 'GLOBAL'})`;
            let skuVaaMap = {}, hasValidDates = false;
            const clientsToProcess = scope === 'client' ? [selectedClient] : Array.from(allClientsFlat.values());
            
            clientsToProcess.forEach(clientObj => {
                if (clientObj && clientObj.rawItems) {
                    clientObj.rawItems.forEach(item => {
                        const sku = getSku(item); if (!sku) return;
                        const fechaFact = item.fecha_doc;
                        if (fechaFact) {
                            hasValidDates = true;
                            if (fechaFact.getMonth() === targetMonth && fechaFact.getFullYear() === targetYear) {
                                let ventaPesos = item._vmn_calculado || 0;
                                if (ventaPesos > 0) { if (!skuVaaMap[sku]) skuVaaMap[sku] = 0; skuVaaMap[sku] += ventaPesos; }
                            }
                        }
                    });
                }
            });
            if (!hasValidDates) { showAlert("EL ARCHIVO NO TIENE FECHAS VÁLIDAS."); return; }
            let skusOrdenados = Object.keys(skuVaaMap).map(sku => ({ sku, totalVenta: skuVaaMap[sku] })).sort((a,b)=>b.totalVenta - a.totalVenta);
            targetProducts = skusOrdenados.slice(0, 50).map(x => globalProducts.get(x.sku)).filter(Boolean);
        } else if (strategyType === 'noComprados90') {
            currentStrategyLabel = "Recuperación +90 días";
            const limitDate = new Date(); limitDate.setDate(limitDate.getDate() - 90);
            const skuMaxDate = new Map();
            selectedClient.rawItems.forEach(item => {
                const sku = getSku(item), d = item.fecha_doc;
                if (d && (!skuMaxDate.has(sku) || d > skuMaxDate.get(sku))) skuMaxDate.set(sku, d);
            });
            const inactiveSkus = [];
            for (const [sku, lastDate] of skuMaxDate.entries()) { if (lastDate < limitDate) inactiveSkus.push(sku); }
            targetProducts = inactiveSkus.sort((a,b)=>(globalSalesPesos[b]||0)-(globalSalesPesos[a]||0)).slice(0, 50).map(sku => globalProducts.get(sku)).filter(Boolean);
        } else if (strategyType === 'topGlobal') {
            currentStrategyLabel = "Top Global Más Vendidos";
            targetProducts = sortedGlobalSkusMn.slice(0, 50).map(sku => globalProducts.get(sku)).filter(Boolean);
        } else if (strategyType === 'tipoPromocion') {
            const selectedPromo = document.getElementById('promo-selector').value; if (!selectedPromo) { showAlert("SELECCIONA UNA PROMO."); return; }
            currentStrategyLabel = `Promoción: ${selectedPromo}`;
            const selectedPromoNorm = normalizeBrandName(selectedPromo);
            targetProducts = Array.from(globalProducts.keys())
                .filter(sku => normalizeBrandName(globalPromoText[sku] || globalComentarioPromo[sku] || "") === selectedPromoNorm)
                .sort((a,b)=>(globalSalesPesos[b]||0)-(globalSalesPesos[a]||0))
                .map(sku => globalProducts.get(sku)).filter(Boolean);
        } else if (strategyType === 'topMarca') {
            currentStrategyLabel = "Top Marcas Propias";
            targetProducts = sortedGlobalSkusPesos.map(sku => globalProducts.get(sku)).filter(item => item && marcasPropias.map(b=>b.toUpperCase()).includes(normalizeBrandName(getBrand(item)))).slice(0, 50);
        } else if (strategyType === 'marcaFamilia') {
            const selectedBrand=document.getElementById('brand-selector').value,selectedFamily=document.getElementById('brand-family-selector')?.value||'';
            if(!selectedBrand){showAlert('SELECCIONA UNA MARCA.');return;}
            currentStrategyLabel=`Marca: ${selectedBrand}`+(selectedFamily?` | Familia: ${selectedFamily}`:' | Todas las familias');
            targetProducts=Array.from(globalProducts.values()).filter(item=>item&&normalizeBrandName(getBrand(item))===normalizeBrandName(selectedBrand)&&(!selectedFamily||normalizeBrandName(getFamily(item))===normalizeBrandName(selectedFamily)));
        } else if (strategyType === 'codigosNuevos') {
            const selectedBrand=document.getElementById('new-code-brand-selector')?.value||'',selectedFamily=document.getElementById('new-code-family-selector')?.value||'';
            currentStrategyLabel='Códigos Nuevos'+(selectedBrand?` | Marca: ${selectedBrand}`:' | Todas las marcas')+(selectedFamily?` | Familia: ${selectedFamily}`:' | Todas las familias');
            targetProducts=Array.from(globalProducts.values()).filter(item=>item&&isNewProduct(item)&&(!selectedBrand||normalizeBrandName(getBrand(item))===normalizeBrandName(selectedBrand))&&(!selectedFamily||normalizeBrandName(getFamily(item))===normalizeBrandName(selectedFamily)));
        } else if (strategyType === 'familyEspecifica') {
            const selectedFamily = document.getElementById('family-selector').value, selectedSistema = document.getElementById('sistema-selector').value;
            if (!selectedFamily) { showAlert("SELECCIONA UNA FAMILIA."); return; }
            currentStrategyLabel = `Familia Propia ${selectedFamily}` + (selectedSistema ? ` - ${selectedSistema}` : '');
            targetProducts = sortedGlobalSkusPesos.map(sku => globalProducts.get(sku)).filter(item => item && normalizeBrandName(getFamily(item)) === normalizeBrandName(selectedFamily) && marcasPropias.map(b=>normalizeBrandName(b)).includes(normalizeBrandName(getBrand(item))) && (!selectedSistema || normalizeBrandName(getSistema(item)) === normalizeBrandName(selectedSistema))).slice(0, 50);
        } else if (strategyType === 'tipoServicioFamilia') {
            const selectedServicio = document.getElementById('servicio-selector').value;
            const selectedArmadora = document.getElementById('servicio-family-selector').value;
            if (!selectedServicio) { showAlert("SELECCIONA UN TIPO DE SERVICIO."); return; }
            currentStrategyLabel = `Servicio: ${selectedServicio}` + (selectedArmadora ? ` | Armadora: ${selectedArmadora}` : '');
            targetProducts = sortedGlobalSkusPesos.map(sku => globalProducts.get(sku)).filter(item => item && normalizeBrandName(getTipoServicio(item)) === normalizeBrandName(selectedServicio) && (!selectedArmadora || normalizeBrandName(getArmadora(item)) === normalizeBrandName(selectedArmadora))).slice(0, 50);
        }

        if(strategyType!=='codigosNuevos')targetProducts=applyOpportunityFilter(targetProducts);
        targetProducts.sort((a,b)=>getSku(a).localeCompare(getSku(b),'es',{numeric:true,sensitivity:'base'}));
        if (targetProducts.length === 0) { showAlert("NO SE ENCONTRARON CÓDIGOS PARA ESTA ESTRATEGIA Y FILTRO."); return; }  
        renderCatalog(targetProducts);
    }

    function renderCatalog(products) {
        catalogContainer.innerHTML = ''; stepStrategy.classList.add('hidden'); appCurrentState = "CATALOGO_LISTO"; btnPrint.removeAttribute('disabled');
        const itemsPerPage = 9, totalPages = Math.ceil(products.length / itemsPerPage), moneyFmt = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 });
        const proyeccion = DIAS_TRANSCURRIDOS > 0 ? (selectedClient.venta / DIAS_TRANSCURRIDOS) * DIAS_EFECTIVOS_MES : 0;
        const porcentaje = selectedClient.cuota > 0 ? (proyeccion / selectedClient.cuota) * 100 : 0;

        for (let i = 0; i < totalPages; i++) {
            const pageProducts = products.slice(i * itemsPerPage, (i + 1) * itemsPerPage);
            const pageSheet = document.createElement('div'); pageSheet.className = "page-sheet";
            let headerHtml = `
                <div class="client-header-card">
                    <div class="flex-grow">
                        <div class="text-[8px] font-black tracking-widest opacity-60 uppercase mb-0.5">Análisis Comercial Estructurado</div>
                        <div class="neon-client-name uppercase">${selectedClient.id} - ${selectedClient.name}</div>
                        <div class="text-[8.5px] font-bold text-slate-200 mt-1.5 uppercase flex items-center gap-2">Estrategia: <span class="strategy-badge">${currentStrategyLabel}</span></div>
                    </div>
                    <div class="header-divider"></div>
                    <div class="text-center px-2 min-w-[120px]">
                        <div class="lbl-secundario mb-1">Proyección Cierre</div>
                        <div class="text-xl font-black ${getSemaforoProyeccionTextClass(porcentaje)}">${porcentaje.toFixed(1)}%</div>
                        <div class="text-[6.5px] font-bold text-slate-300 uppercase mt-0.5">${MESES_NOMBRES[CURRENT_MONTH]} (Día ${DIAS_TRANSCURRIDOS}/${DIAS_EFECTIVOS_MES})</div>
                    </div>
                    <div class="header-divider"></div>
                    <div class="text-center px-2 min-w-[110px]">
                        <div class="lbl-secundario mb-1">Venta / Meta</div>
                        <div class="text-[10px] font-black text-white">${moneyFmt.format(selectedClient.venta)}</div>
                        <div class="text-[7px] font-bold text-neon-red uppercase mt-0.5">Meta: ${moneyFmt.format(selectedClient.cuota)}</div>
                    </div>
                    <div class="header-divider"></div>
                    <div class="flex flex-col items-center justify-center"><div class="pag-badge">PÁG.<br>${i + 1}/${totalPages}</div></div>
                </div>
                <div class="grid-container">`;
            
            pageProducts.forEach(prod => {
                const skuCode = getSku(prod), imgUrl = getProductImageUrl(prod) || 'https://via.placeholder.com/150?text=Sin+Imagen', isSelectedClass = cartItems.has(skuCode) ? 'selected' : '', brandLogo = getBrandLogoUrl(prod);
                let brandDisplay = brandLogo ? `<img src="${brandLogo}" class="h-5 max-w-[60px] object-contain" alt="Logo" onerror="this.onerror=null; this.style.display='none'; this.nextElementSibling.style.display='inline';"><span style="display:none;" class="text-[7px] font-black text-slate-400 uppercase">${getBrand(prod)}</span>` : `<span class="text-[7px] font-black text-slate-400 uppercase">${getBrand(prod)}</span>`;
                
                const historyRecords = selectedClient.rawItems.filter(raw => getSku(raw) === skuCode);
                let uFecha = "SIN REGISTRO", uPrecio = 0;
                let fechasValidas = historyRecords.map(r => r.fecha_doc).filter(Boolean);
                if(fechasValidas.length > 0) uFecha = new Date(Math.max(...fechasValidas)).toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric' });
                
                let sortedHistory = [...historyRecords].sort((a,b)=> (a.fecha_doc||0) - (b.fecha_doc||0));
                for (let k = sortedHistory.length - 1; k >= 0; k--) {
                    let pz = sortedHistory[k]._venta_piezas_calculada, mn = sortedHistory[k]._vmn_calculado;
                    if (pz > 0) { uPrecio = mn / pz; break; }
                }

                const precioSolicitado = Number(globalPromociones[skuCode] || prod._precio_solicitado_catalogo || parseMoney(prod.precio_solicitado ?? prod.preciosolicitado ?? prod["precio solicitado"] ?? 0));
                const promoFinal = precioSolicitado > 0 ? moneyFmt.format(precioSolicitado) : "SIN PROMOCIÓN";
                const nombrePromoCompleto = String(globalComentarioPromo[skuCode] || prod._nombre_promocion_catalogo || prod.comentario || "").trim().toUpperCase();
                const abreviarPromo = (value, max=24) => value.length > max ? value.slice(0, max - 1).trimEnd() + "…" : value;
                const nombrePromoVisible = nombrePromoCompleto ? abreviarPromo(nombrePromoCompleto) : "PROMOCIÓN";
                const promoHtml = `<div class="spec-tag !border-[#ff004c]/30 mt-1 pb-1 flex justify-between items-center"><span class="text-[#ff004c] font-black max-w-[55%] text-[6.5px] leading-tight" title="${nombrePromoCompleto || 'PROMOCIÓN'}">🔥 ${nombrePromoVisible}</span><span class="spec-value text-[#ff004c] font-black text-[10px]">${promoFinal}</span></div>`;

                let estadoCompra = uPrecio > 0 ? `<div class="w-full bg-[#00ff66]/10 border border-[#00ff66] text-[#00b347] text-center py-1.5 rounded text-[8.5px] font-black uppercase">✔ Adquirido Previamente</div>` : `<div class="w-full bg-[#ff004c]/10 border border-[#ff004c] text-[#ff004c] text-center py-1.5 rounded text-[8.5px] font-black uppercase"><span class="animate-pulse">🔥</span> Nueva Oportunidad</div>`;

                headerHtml += `
                    <div id="card-${skuCode}" class="card cursor-pointer transition-all ${isSelectedClass}" onclick="toggleCart('${skuCode}')">
                        <div class="flex justify-between items-start"><span class="sku-badge">${skuCode}</span>${brandDisplay}</div>
                        <div class="card-img-container"><img src="${imgUrl}" class="card-img" onerror="this.onerror=null; this.src='https://via.placeholder.com/150?text=No+Disponible'"></div>
                        <div class="card-title" title="${getDescription(prod)}">${getDescription(prod)}</div>
                        <div class="space-y-0.5 mb-2">
                            <div class="spec-tag">Línea/Familia: <span class="spec-value">${getFamily(prod)}</span></div>
                            <div class="spec-tag">Última Compra: <span class="spec-value">${uFecha}</span></div>
                            ${promoHtml}
                        </div>
                        <div class="mt-auto pt-2 border-t border-dashed border-slate-200 w-full flex flex-col justify-center">${estadoCompra}</div>
                    </div>`;
            });

            if (pageProducts.length < itemsPerPage) { for (let j = 0; j < (itemsPerPage - pageProducts.length); j++) headerHtml += `<div class="card opacity-0 border-none pointer-events-none"></div>`; }
            headerHtml += `</div>`; pageSheet.innerHTML = headerHtml; catalogContainer.appendChild(pageSheet);
        }
    }

    function toggleCart(sku) {
        if (cartItems.has(sku)) { cartItems.delete(sku); const el = document.getElementById(`card-${sku}`); if(el) el.classList.remove('selected'); }
        else { cartItems.add(sku); const el = document.getElementById(`card-${sku}`); if(el) el.classList.add('selected'); }
        updateCartUI();
    }
    
    function removeSkuFromCart() {
        const input = document.getElementById('remove-sku-input'); const sku = input.value.trim().toUpperCase(); if (!sku) return;
        let found = false;
        for (let item of cartItems) { if (item.toUpperCase() === sku) { cartItems.delete(item); const el = document.getElementById(`card-${item}`); if (el) el.classList.remove('selected'); found = true; break; } }
        if (found) { updateCartUI(); input.value = ''; } else showAlert("El código " + sku + " no está en tu cotización.");
    }

    function clearCart() {
        if (cartItems.size === 0) return;
        if (confirm("¿Estás seguro de vaciar toda la cotización general?")) {
            cartItems.forEach(sku => { const el = document.getElementById(`card-${sku}`); if (el) el.classList.remove('selected'); });
            cartItems.clear(); updateCartUI();
        }
    }

    function updateCartUI() {
        const bar = document.getElementById('cart-bar'); document.getElementById('cart-count').innerText = cartItems.size;
        if (cartItems.size > 0) bar.classList.add('visible'); else { bar.classList.remove('visible'); closeCartDrawer(); }
        renderCartItems();
    }
    
    function hideCartBar() { cartItems.clear(); updateCartUI(); closeCartDrawer(); }
    
    function openCartDrawer() { 
        if(cartItems.size > 0) {
            toggleBackgroundLock(true);
            document.getElementById('cart-drawer').classList.remove('translate-x-full'); 
        }
    }
    function closeCartDrawer() { 
        toggleBackgroundLock(false);
        document.getElementById('cart-drawer').classList.add('translate-x-full'); 
    }

    function getCalculatedPriceForClient(sku, clientObj) {
        if (globalPromociones[sku] > 0) return globalPromociones[sku];
        let lowest = null;
        if (clientObj && clientObj.rawItems) {
            clientObj.rawItems.filter(raw => getSku(raw) === sku).forEach(r => {
                let pz = r._venta_piezas_calculada, mn = r._vmn_calculado;
                if (pz > 0 && mn > 0 && (lowest === null || (mn / pz) < lowest)) lowest = mn / pz;
            });
        }
        return lowest;
    }

    function renderCartItems() {
        const container = document.getElementById('cart-drawer-items'); container.innerHTML = '';
        if(cartItems.size === 0) { container.innerHTML = '<div class="text-center text-slate-500 font-bold uppercase mt-10 text-[10px]">El carrito está vacío</div>'; return; }
        const fmt = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 });

        Array.from(cartItems).reverse().forEach(sku => {
            const prod = globalProducts.get(sku); if (!prod) return;
            const calcP = getCalculatedPriceForClient(sku, selectedClient);
            const promo = globalComentarioPromo[sku] || (globalPromoText[sku] && globalPromoText[sku] !== "0" ? globalPromoText[sku] : (globalPromociones[sku] ? fmt.format(globalPromociones[sku]) : 'Precio sin IVA'));
            container.innerHTML += `
                <div class="bg-white/5 border border-white/10 hover:border-neon-cyan transition-colors rounded-xl p-4 relative group">
                    <button onclick="toggleCart('${sku}')" class="absolute top-3 right-3 text-slate-500 hover:text-neon-red bg-black/50 rounded-full p-1"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg></button>
                    <div class="text-neon-cyan font-black text-sm mb-1">${sku}</div>
                    <div class="text-[9.5px] text-slate-300 uppercase font-bold mb-3 line-clamp-2 pr-6">${getDescription(prod)}</div>
                    <div class="grid grid-cols-2 gap-3 mt-2 pt-3 border-t border-white/5">
                        <div class="bg-black/40 p-2 rounded border border-white/5"><span class="text-[7.5px] text-slate-400 font-bold block">Precio Histórico</span><span class="text-[11px] text-white font-black">${calcP !== null ? fmt.format(calcP) : ''}</span></div>
                        <div class="bg-black/40 p-2 rounded border border-white/5"><span class="text-[7.5px] text-slate-400 font-bold block">Estatus/Promo</span><span class="text-[11px] font-black text-neon-red line-clamp-1">${promo}</span></div>
                    </div>
                </div>`;
        });
    }

    function syncStrategyFamilySelector(brandId,familyId,onlyNew=false){const brand=document.getElementById(brandId)?.value||'',sel=document.getElementById(familyId);if(!sel)return;const fs=[...new Set(Array.from(globalProducts.values()).filter(p=>p&&(!onlyNew||isNewProduct(p))&&(!brand||normalizeBrandName(getBrand(p))===normalizeBrandName(brand))).map(getFamily).filter(f=>f&&f!=='SIN FAMILIA'))].sort((a,b)=>a.localeCompare(b,'es'));sel.innerHTML='<option value="">Todas las familias...</option>';fs.forEach(f=>{const o=document.createElement('option');o.value=f;o.innerText=f.toUpperCase();sel.appendChild(o)})}
    function syncServiceArmadoraSelector(){
        const servicio=document.getElementById('servicio-selector')?.value||'';
        const selector=document.getElementById('servicio-family-selector');
        if(!selector)return;
        const armadoras=[...new Set(Array.from(globalProducts.values())
            .filter(item=>item&&(!servicio||normalizeBrandName(getTipoServicio(item))===normalizeBrandName(servicio)))
            .map(item=>getArmadora(item)).filter(a=>a&&a!=='SIN ARMADORA'))]
            .sort((a,b)=>a.localeCompare(b,'es'));
        selector.innerHTML='<option value="">Seleccionar armadora (Opcional)...</option>';
        armadoras.forEach(a=>{const opt=document.createElement('option');opt.value=a;opt.innerText=a.toUpperCase();selector.appendChild(opt)});
    }
    function populateSelectors() {
        const uniqueBrands = [...new Set(Array.from(globalProducts.values()).map(item => getBrand(item)).filter(b => b && b !== "SIN MARCA"))].sort();
        const bSel = document.getElementById('brand-selector'); if(bSel) { bSel.innerHTML = '<option value="">Seleccionar marca...</option>'; uniqueBrands.forEach(b => { const opt = document.createElement('option'); opt.value = b; opt.innerText = b; bSel.appendChild(opt); }); }
        
        const newBrands=[...new Set(Array.from(globalProducts.values()).filter(isNewProduct).map(getBrand).filter(b=>b&&b!=='SIN MARCA'))].sort((a,b)=>a.localeCompare(b,'es'));
        const nb=document.getElementById('new-code-brand-selector');if(nb){nb.innerHTML='<option value="">Todas las marcas...</option>';newBrands.forEach(b=>{const o=document.createElement('option');o.value=b;o.innerText=b;nb.appendChild(o)})}
        syncStrategyFamilySelector('brand-selector','brand-family-selector',false);syncStrategyFamilySelector('new-code-brand-selector','new-code-family-selector',true);
        const uniqueOwnFamilies = [...new Set(Array.from(globalProducts.values()).filter(item => marcasPropias.map(b=>b.toUpperCase()).includes(normalizeBrandName(getBrand(item)))).map(item => getFamily(item)).filter(f => f && f !== "SIN FAMILIA"))].sort();
        const fSel = document.getElementById('family-selector'); if(fSel) { fSel.innerHTML = '<option value="">Seleccionar familia...</option>'; uniqueOwnFamilies.forEach(f => { const opt = document.createElement('option'); opt.value = f; opt.innerText = f.toUpperCase(); fSel.appendChild(opt); }); }

        syncServiceArmadoraSelector();

        const uniqueSistemas = [...new Set(Array.from(globalProducts.values()).map(item => getSistema(item)).filter(s => s && s !== "SIN SISTEMA"))].sort();
        const sSel = document.getElementById('sistema-selector'); if(sSel) { sSel.innerHTML = '<option value="">Seleccionar sistema (Opcional)...</option>'; uniqueSistemas.forEach(s => { const opt = document.createElement('option'); opt.value = s; opt.innerText = s.toUpperCase(); sSel.appendChild(opt); }); }

        const uniqueServicios = [...new Set(Array.from(globalProducts.values()).map(item => getTipoServicio(item)).filter(s => s && s !== "SIN SERVICIO"))].sort();
        const serSel = document.getElementById('servicio-selector'); if(serSel) { serSel.innerHTML = '<option value="">Seleccionar servicio...</option>'; uniqueServicios.forEach(s => { const opt = document.createElement('option'); opt.value = s; opt.innerText = s.toUpperCase(); serSel.appendChild(opt); }); }

        const uniqueArmadoras = [...new Set(Array.from(globalProducts.values()).map(item => getArmadora(item)).filter(s => s && s !== "SIN ARMADORA"))].sort();
        const armSel = document.getElementById('armadora-selector'); if(armSel) { armSel.innerHTML = '<option value="">Seleccionar armadora (Opcional)...</option>'; uniqueArmadoras.forEach(s => { const opt = document.createElement('option'); opt.value = s; opt.innerText = s.toUpperCase(); armSel.appendChild(opt); }); }

        const pSel = document.getElementById('promo-selector');
        if (pSel) {
            const uniquePromos = new Map();
            Array.from(globalProducts.keys()).forEach(sku => {
                const p = globalPromoText[sku] || globalComentarioPromo[sku] || null;
                if (p) uniquePromos.set(normalizeBrandName(p), p);
            });
            pSel.innerHTML = '<option value="">Seleccionar promo...</option>'; [...uniquePromos.values()].sort((a,b)=>a.localeCompare(b, 'es')).forEach(p => { const opt = document.createElement('option'); opt.value = p; opt.innerText = p.toUpperCase(); pSel.appendChild(opt); });
        }

        const ySel = document.getElementById('sku-year-selector');
        if (ySel && ySel.options.length === 0) {
            for (let i = CURRENT_YEAR - 2; i <= CURRENT_YEAR; i++) { const opt = document.createElement('option'); opt.value = i; opt.innerText = i; ySel.appendChild(opt); }
            ySel.value = CURRENT_YEAR; document.getElementById('sku-month-selector').value = CURRENT_MONTH;
        }
    }

    function exportCartToExcel() {
        if (cartItems.size === 0) { showAlert("No hay artículos seleccionados."); return; }
        let excelData = []; const fmt = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 });
        cartItems.forEach(sku => {
            const prod = globalProducts.get(sku); if (!prod) return;
            let comment = globalComentarioPromo[sku] || (globalPromoText[sku] && globalPromoText[sku] !== "0" ? globalPromoText[sku] : (globalPromociones[sku] ? fmt.format(globalPromociones[sku]) : "Precio Sin IVA"));
            const calcP = getCalculatedPriceForClient(sku, selectedClient);
            excelData.push({ "Producto": sku, "Cantidad": 1, "Precio solicitado": calcP !== null ? parseFloat(calcP.toFixed(2)) : "", "Comentario": comment });
        });
        const ws = XLSX.utils.json_to_sheet(excelData), wb = XLSX.utils.book_new();
        ws['!cols'] = [{wch: 20}, {wch: 15}, {wch: 25}, {wch: 40}];
        XLSX.utils.book_append_sheet(wb, ws, "Cotización");
        XLSX.writeFile(wb, `Cotizacion_${selectedClient.id}_${new Date().getTime()}.xlsx`);
    }

    function toggleBackgroundLock(lock) {
        const uiContainers = ['step-1', 'step-2', 'step-strategy', 'discipline-client-detail', 'catalog-container', 'cart-bar'];
        const header = document.querySelector('header');
        if(header) {
            header.style.pointerEvents = lock ? 'none' : 'auto';
            header.style.filter = lock ? 'blur(2px) brightness(0.6)' : 'none';
            header.style.transition = 'all 0.3s ease';
        }
        uiContainers.forEach(id => {
            const el = document.getElementById(id);
            if(el) {
                el.style.pointerEvents = lock ? 'none' : 'auto';
                el.style.filter = lock ? 'blur(4px) brightness(0.5)' : 'none';
                el.style.transition = 'all 0.3s ease';
            }
        });
        document.body.style.overflow = lock ? 'hidden' : 'auto';
    }

    window.addEventListener('beforeunload', function (e) {
        e.preventDefault();
        e.returnValue = '¿Estás seguro de que deseas salir de la aplicación?';
        return '¿Estás seguro de que deseas salir de la aplicación?';
    });

    document.addEventListener('contextmenu', e => e.preventDefault());
    document.addEventListener('selectstart', e => { 
        if(e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') {
            e.preventDefault(); 
        }
    });
    document.addEventListener('copy', e => { 
        if(e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') {
            e.preventDefault(); 
            showAlert("Función de copiado deshabilitada por seguridad."); 
        }
    });
    document.addEventListener('keydown', e => {
        if (e.key === 'F12') { e.preventDefault(); return false; }
        if (e.ctrlKey && e.shiftKey && ['I','J','C'].includes(e.key.toUpperCase())) { e.preventDefault(); return false; }
        if (e.ctrlKey && ['u','s','p'].includes(e.key.toLowerCase())) { e.preventDefault(); return false; }
        if (e.ctrlKey && e.key.toLowerCase() === 'c' && e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') { e.preventDefault(); showAlert("Copiar contenido deshabilitado."); return false; }
    });
