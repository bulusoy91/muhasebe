// 1. Dependencies & Locale Initialization
dayjs.locale('tr');
dayjs.extend(window.dayjs_plugin_isSameOrBefore);
dayjs.extend(window.dayjs_plugin_isSameOrAfter);

// --- GÜVENLİK AYARLARI (Şifresiz Mod) ---
let masterKey = 'baris1903'; // Verileri açmak için kullanılan anahtar
let payments = [];
let groupNotes = {};

// --- BULUT SENKRONİZASYON AYARLARI ---
let cloudSettings = JSON.parse(localStorage.getItem('cloud_sync_settings') || 'null');
const CLOUD_FILENAME = 'odemeler_yedek_sifreli.json';
// ------------------------------------

async function loadEncryptedData(enteredPass) {
    try {
        let encPayments = null;
        let encNotes = null;

        // 1. ÖNCE BULUTTAN ÇEKMEYİ DENE (Eğer ayarlar varsa)
        if (cloudSettings && cloudSettings.ghToken) {
            try {
                const cloudData = await fetchGitHubFile();
                if (cloudData && cloudData.encrypted) {
                    encPayments = cloudData.payments_enc;
                    encNotes = cloudData.grup_notlari_enc;
                    // Başarılıysa locale de yedekle
                    localStorage.setItem('odemeler_enc', encPayments);
                    localStorage.setItem('grup_notlari_enc', encNotes);
                    updateCloudStatus(true);
                }
            } catch (err) {
                console.error("Bulut çekme hatası:", err);
                updateCloudStatus(false);
            }
        }

        // 2. BULUT BAŞARISIZSA VEYA AYAR YOKSA LOCALSTORAGE'A BAK
        if (!encPayments) {
            encPayments = localStorage.getItem('odemeler_enc');
            encNotes = localStorage.getItem('grup_notlari_enc');
        }

        // 3. EĞER LOCALDE DE YOKSA STATİK JSON DOSYASINA BAK (Fallback)
        if (!encPayments) {
            try {
                const response = await fetch(CLOUD_FILENAME + '?t=' + Date.now());
                if (response.ok) {
                    const data = await response.json();
                    if (data.encrypted) {
                        encPayments = data.payments_enc;
                        encNotes = data.grup_notlari_enc;
                    }
                }
            } catch (e) {}
        }

        if (!encPayments) return true; // Hiç veri yoksa boş başla

        const bytesPay = CryptoJS.AES.decrypt(encPayments, enteredPass);
        const decryptedPay = bytesPay.toString(CryptoJS.enc.Utf8);
        
        if (!decryptedPay) return false; 

        payments = JSON.parse(decryptedPay);
        
        if (encNotes) {
            const bytesNotes = CryptoJS.AES.decrypt(encNotes, enteredPass);
            groupNotes = JSON.parse(bytesNotes.toString(CryptoJS.enc.Utf8));
        }
        
        masterKey = enteredPass;
        return true;
    } catch (e) {
        return false;
    }
}

async function fetchGitHubFile() {
    if (!cloudSettings) return null;
    const { ghUsername, ghRepo, ghToken } = cloudSettings;
    const url = `https://api.github.com/repos/${ghUsername}/${ghRepo}/contents/${CLOUD_FILENAME}`;
    
    const response = await fetch(url, {
        headers: { 'Authorization': `token ${ghToken}`, 'Accept': 'application/vnd.github.v3+json' },
        cache: 'no-store'
    });
    
    if (!response.ok) return null;
    const data = await response.json();
    return JSON.parse(atob(data.content));
}

async function updateGitHubFile(contentObj) {
    if (!cloudSettings) return false;
    const { ghUsername, ghRepo, ghToken } = cloudSettings;
    const url = `https://api.github.com/repos/${ghUsername}/${ghRepo}/contents/${CLOUD_FILENAME}`;
    
    // Önce mevcut dosyanın SHA değerini almalıyız
    const getRes = await fetch(url, {
        headers: { 'Authorization': `token ${ghToken}` }
    });
    
    let sha = null;
    if (getRes.ok) {
        const getData = await getRes.json();
        sha = getData.sha;
    }

    const putRes = await fetch(url, {
        method: 'PUT',
        headers: { 
            'Authorization': `token ${ghToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            message: `Update payments ${new Date().toLocaleString()}`,
            content: btoa(JSON.stringify(contentObj)),
            sha: sha
        })
    });

    return putRes.ok;
}

function updateCloudStatus(connected) {
    const dot = document.getElementById('cloudStatusDot');
    if (!dot) return;
    if (connected) {
        dot.className = 'absolute -top-1 -right-1 w-2.5 h-2.5 rounded-full bg-emerald-500 border-2 border-white animate-pulse';
    } else {
        dot.className = 'absolute -top-1 -right-1 w-2.5 h-2.5 rounded-full bg-red-500 border-2 border-white';
    }
}

async function savePayments() {
    if (!masterKey) return;
    
    // Verileri masterKey ile şifreleyip kaydet
    const encPay = CryptoJS.AES.encrypt(JSON.stringify(payments), masterKey).toString();
    const encNotes = CryptoJS.AES.encrypt(JSON.stringify(groupNotes), masterKey).toString();
    
    localStorage.setItem('odemeler_enc', encPay);
    localStorage.setItem('grup_notlari_enc', encNotes);

    const backupData = {
        payments_enc: encPay,
        grup_notlari_enc: encNotes,
        encrypted: true
    };

    // 1. BULUT SENKRONİZASYONU (Aktifse GitHub'a gönder)
    if (cloudSettings && cloudSettings.ghToken) {
        updateGitHubFile(backupData).then(success => {
            updateCloudStatus(success);
        });
    }

    // 2. MASAÜSTÜ YEDEĞİ (Eğer EXE içindeyse)
    if (typeof Neutralino !== 'undefined') {
        try {
            await Neutralino.filesystem.writeFile('./' + CLOUD_FILENAME, JSON.stringify(backupData));
        } catch (err) {}
    }
    
    renderCalendar();
    if (typeof allPaymentsModalOverlay !== 'undefined' && allPaymentsModalOverlay && !allPaymentsModalOverlay.classList.contains('opacity-0')) renderAllPayments();
    if (typeof monthlyTableModalOverlay !== 'undefined' && monthlyTableModalOverlay && !monthlyTableModalOverlay.classList.contains('opacity-0')) renderMonthlyTable();
}

// ------------------------------------

let currentWeekStart = dayjs().startOf('week'); 
let filterMode = 'week'; 
let filterStart = null;
let filterEnd = null;
let pendingDeleteId = null;
let expandedGroups = new Set(); 
let expandedSections = new Set(); 
let paymentDatePicker, jumpToDatePicker;

// DOM Elements
let calendarContainer, currentPeriodLabel, overallTotalDisplay, paymentModal, paymentModalOverlay, 
    addPaymentBtn, closeModalBtn, cancelPaymentBtn, paymentForm, showAllPaymentsBtn, 
    allPaymentsModalOverlay, allPaymentsModal, closeAllPaymentsModalBtn, allPaymentsContent, 
    confirmDeleteModalOverlay, confirmDeleteModal, deleteOnlyThisBtn, deleteAllGroupBtn, cancelDeleteBtn,
    groupDeleteActions, singleDeleteActions, deleteModalTitle, deleteModalText, confirmSingleDeleteBtn,
    showMonthlyTableBtn, monthlyTableModalOverlay, monthlyTableModal, monthlyTableContent,
    closeMonthlyTableModalBtn, exportMonthlyExcelBtn, exportMonthlyPdfBtn, monthlyTableTitle,
    monthlyMonthSelect, monthlyYearSelect, colorPickerContainer;



const categoryIcons = {
    'Kredi': 'credit-card',
    'Çek': 'layers',
    'Kira': 'home',
    'Fatura': 'file-text',
    'Maaş': 'users',
    'Diğer': 'layers'
};

const categoryColors = {
    'Kredi': 'bg-rose-500 text-white',
    'Çek': 'bg-amber-500 text-white',
    'Kira': 'bg-emerald-500 text-white',
    'Kredi Kartı': 'bg-sky-500 text-white',
    'Fatura': 'bg-orange-500 text-white',
    'Maaş': 'bg-teal-500 text-white',
    'Diğer': 'bg-slate-500 text-white'
};

const categoryHexColors = {
    'Kredi': '#f43f5e',
    'Çek': '#f59e0b',
    'Kira': '#10b981',
    'Kredi Kartı': '#0ea5e9',
    'Fatura': '#f97316',
    'Maaş': '#14b8a6',
    'Diğer': '#64748b'
};

const bankColors = {
    'AKBANK': 'bg-red-600 text-white',
    'DENİZBANK': 'bg-blue-600 text-white',
    'GARANTİ': 'bg-emerald-600 text-white',
    'HALK BANKASI': 'bg-sky-600 text-white',
    'İNG': 'bg-orange-600 text-white',
    'İŞ BANKASI': 'bg-blue-800 text-white',
    'KUVEYTTÜRK': 'bg-green-700 text-white',
    'QNB': 'bg-indigo-600 text-white',
    'ŞEKERBANK': 'bg-green-600 text-white',
    'TEB': 'bg-green-500 text-white',
    'VAKIFBANK': 'bg-yellow-600 text-white',
    'VAKIF KATILIM': 'bg-yellow-600 text-white',
    'YAPIKREDİ': 'bg-blue-600 text-white',
    'ZİRAAT': 'bg-red-700 text-white',
    'ZİRAAT KATILIM BANKASI': 'bg-red-700 text-white',
    'NAKİT KASA': 'bg-slate-500 text-white',
    'DİĞER': 'bg-slate-500 text-white'
};

const allHolidays = [
    "2026-01-01", "2026-03-20", "2026-03-21", "2026-03-22", "2026-03-23", "2026-04-23", "2026-05-01", 
    "2026-05-19", "2026-05-27", "2026-05-28", "2026-05-29", "2026-05-30", "2026-05-31", "2026-07-15", 
    "2026-08-30", "2026-10-29",
    "2027-01-01", "2027-03-09", "2027-03-10", "2027-03-11", "2027-04-23", "2027-05-01", 
    "2027-05-16", "2027-05-17", "2027-05-18", "2027-05-19", "2027-07-15", "2027-08-30", 
    "2027-10-29",
    "2028-01-01", "2028-02-27", "2028-02-28", "2028-02-29", "2028-04-23", "2028-05-01", 
    "2028-05-05", "2028-05-06", "2028-05-07", "2028-05-08", "2028-05-19", "2028-07-15", 
    "2028-08-30", "2028-10-29",
    "2029-01-01", "2029-02-15", "2029-02-16", "2029-02-17", "2029-04-23", "2029-04-24", 
    "2029-04-25", "2029-04-26", "2029-04-27", "2029-05-01", "2029-05-19", "2029-07-15", "2029-08-30", 
    "2029-10-29",
    "2030-01-01", "2030-02-04", "2030-02-05", "2030-02-06", "2030-04-13", 
    "2030-04-14", "2030-04-15", "2030-04-16", "2030-04-23", "2030-05-01", "2030-05-19", "2030-07-15", 
    "2030-08-30", "2030-10-29"
];

function isWorkDay(date) {
    const d = dayjs(date);
    if (d.day() === 0 || d.day() === 6) return false;
    return !allHolidays.includes(d.format('YYYY-MM-DD'));
}

function getNextWorkDay(date) {
    let d = dayjs(date);
    while (!isWorkDay(d)) d = d.add(1, 'day');
    return d;
}

function getPrevWorkDay(date) {
    let d = dayjs(date);
    while (!isWorkDay(d)) d = d.subtract(1, 'day');
    return d;
}

function updatePaymentFormUI() {
    const tp = document.getElementById('payType').value;
    const pDate = document.getElementById('payDate');
    const pInst = document.getElementById('payInstallments');
    const instWrapper = document.getElementById('installmentsWrapper');
    
    if (tp === 'first_working_day' || tp === 'last_working_day') {
        pDate.disabled = true;
        pDate.classList.add('bg-slate-100', 'cursor-not-allowed', 'opacity-60');
        if (pInst) pInst.disabled = true;
        if (instWrapper) instWrapper.classList.add('opacity-40', 'pointer-events-none');
    } else {
        pDate.disabled = false;
        pDate.classList.remove('bg-slate-100', 'cursor-not-allowed', 'opacity-60');
        if (pInst) pInst.disabled = false;
        if (instWrapper) instWrapper.classList.remove('opacity-40', 'pointer-events-none');
    }
}

function formatTL(num) {
    return new Intl.NumberFormat('tr-TR', { style: 'currency', currency: 'TRY' }).format(num);
}

// Removed old unencrypted savePayments

window.saveNote = function(id, note) {
    groupNotes[id] = note;
    const groupItems = payments.filter(p => p.groupId === id);
    if (groupItems.length > 0) groupItems.forEach(p => groupNotes[p.id] = note);
    localStorage.setItem('grup_notlari', JSON.stringify(groupNotes));
}

window.toggleSection = function(type) {
    if (expandedSections.has(type)) expandedSections.delete(type);
    else expandedSections.add(type);
    renderAllPayments();
};

window.toggleGroupDetails = function(id) {
    if (expandedGroups.has(id)) expandedGroups.delete(id);
    else expandedGroups.add(id);
    renderAllPayments();
};

window.openModal = function() {
    if (!paymentModalOverlay || !paymentForm) return;
    paymentModalOverlay.classList.remove('opacity-0', 'pointer-events-none');
    setTimeout(() => paymentModal.classList.remove('scale-95', 'opacity-0'), 10);
    paymentForm.reset();
    const editIdElem = document.getElementById('editingId');
    if (editIdElem) editIdElem.value = '';
    
    const editGroupElem = document.getElementById('editingGroupId');
    if (editGroupElem) editGroupElem.value = '';
    
    const mTitle = document.getElementById('modalTitle');
    if (mTitle) mTitle.innerText = 'Yeni Ödeme Ekle';
    
    const today = dayjs().format('DD.MM.YYYY');
    document.getElementById('payDate').value = today;
    if (paymentDatePicker) paymentDatePicker.setDate(today);
    
    document.getElementById('payType').value = 'installments';
    const pInst = document.getElementById('payInstallments');
    if (pInst) pInst.value = 1;
    const n = document.getElementById('payNote');
    if (n) n.value = '';
    updatePaymentFormUI();
}

window.closeModal = function() {
    if (!paymentModal) return;
    paymentModal.classList.add('scale-95', 'opacity-0');
    setTimeout(() => paymentModalOverlay.classList.add('opacity-0', 'pointer-events-none'), 300);
}

window.editPayment = function(id) {
    const pay = payments.find(p => p.id === id);
    if (!pay) return;
    window.openModal();
    const editIdElem = document.getElementById('editingId');
    if (editIdElem) editIdElem.value = pay.id;
    
    const editGroupElem = document.getElementById('editingGroupId');
    if (editGroupElem) editGroupElem.value = ''; 
    
    const mTitle = document.getElementById('modalTitle');
    if (mTitle) mTitle.innerText = 'Ödemeyi Düzenle';

    const pTitle = document.getElementById('payTitle');
    if (pTitle) pTitle.value = pay.title || '';
    
    const pAmt = document.getElementById('payAmount');
    if (pAmt) pAmt.value = pay.amount || 0;
    
    const pBank = document.getElementById('payBank');
    if (pBank) pBank.value = pay.bank || 'AKBANK';
    
    const pCat = document.getElementById('payCategory');
    if (pCat) pCat.value = pay.category || 'Diğer';
    
    const pDate = document.getElementById('payDate');
    if (pDate) {
        const d = dayjs(pay.date).format('DD.MM.YYYY');
        pDate.value = d;
        if (paymentDatePicker) paymentDatePicker.setDate(d);
    }
    
    const pPri = document.getElementById('payPriority');
    if (pPri) pPri.checked = !!pay.priority;
    
    const pType = document.getElementById('payType');
    if (pType) pType.value = 'installments'; 
    
    const pInst = document.getElementById('payInstallments');
    if (pInst) pInst.value = 1;
    const n = document.getElementById('payNote');
    if (n) n.value = pay.note || '';
    updatePaymentFormUI();
}

window.editGroup = function(groupId) {
    const groupPayments = payments.filter(p => p.groupId === groupId || (p.groupId === null && p.id === groupId));
    if (groupPayments.length === 0) return;
    const first = groupPayments[0];
    window.openModal();
    const editIdElem = document.getElementById('editingId');
    if (editIdElem) editIdElem.value = ''; 
    
    const editGroupElem = document.getElementById('editingGroupId');
    if (editGroupElem) editGroupElem.value = groupId; 
    
    const mTitle = document.getElementById('modalTitle');
    if (mTitle) mTitle.innerText = 'Grubu Toplu Düzenle';

    const pTitle = document.getElementById('payTitle');
    if (pTitle) pTitle.value = first.title.split(' (')[0] || '';
    
    const pAmt = document.getElementById('payAmount');
    if (pAmt) pAmt.value = first.amount || 0;
    
    const pBank = document.getElementById('payBank');
    if (pBank) pBank.value = first.bank || 'AKBANK';
    
    const pCat = document.getElementById('payCategory');
    if (pCat) pCat.value = first.category || 'Diğer';
    
    const pDate = document.getElementById('payDate');
    if (pDate) {
        const d = dayjs(first.date).format('DD.MM.YYYY');
        pDate.value = d;
        if (paymentDatePicker) paymentDatePicker.setDate(d);
    }
    
    const pPri = document.getElementById('payPriority');
    if (pPri) pPri.checked = !!first.priority;
    
    const pType = document.getElementById('payType');
    if (pType) pType.value = 'installments';
    
    const pInst = document.getElementById('payInstallments');
    if (pInst) pInst.value = groupPayments.length;
    const n = document.getElementById('payNote');
    if (n) n.value = first.note || '';
    updatePaymentFormUI();
}

window.deletePayment = function(id) {
    pendingDeleteId = id;
    const pay = payments.find(p => p.id === id);
    
    if (pay && pay.groupId) {
        deleteModalTitle.innerText = 'Taksitli Ödeme Silme';
        deleteModalText.innerText = 'Bu ödeme bir grubun (kredi/taksit) bir parçasıdır. Nasıl silmek istersiniz?';
        groupDeleteActions.classList.remove('hidden');
        singleDeleteActions.classList.add('hidden');
    } else {
        deleteModalTitle.innerText = 'Ödemeyi Sil';
        deleteModalText.innerText = 'Bu ödemeyi silmek istediğinizden emin misiniz?';
        groupDeleteActions.classList.add('hidden');
        singleDeleteActions.classList.remove('hidden');
    }
    
    confirmDeleteModalOverlay.classList.remove('opacity-0', 'pointer-events-none');
    setTimeout(() => confirmDeleteModal.classList.remove('scale-95', 'opacity-0'), 10);
}

window.deleteGroup = function(groupId) {
    const groupItem = payments.find(p => p.groupId === groupId);
    if (!groupItem) return;
    
    pendingDeleteId = groupItem.id; // Set this so deletion logic can find the group via groupId
    deleteModalTitle.innerText = 'Grubu Sil';
    deleteModalText.innerText = 'Taksit grubunu nasıl silmek istersiniz?';
    groupDeleteActions.classList.remove('hidden');
    singleDeleteActions.classList.add('hidden');
    confirmDeleteModalOverlay.classList.remove('opacity-0', 'pointer-events-none');
    setTimeout(() => confirmDeleteModal.classList.remove('scale-95', 'opacity-0'), 10);
}

window.closeConfirmDeleteModal = function() {
    confirmDeleteModal.classList.add('scale-95', 'opacity-0');
    setTimeout(() => confirmDeleteModalOverlay.classList.add('opacity-0', 'pointer-events-none'), 300);
}

window.togglePriority = function(id) {
    payments = payments.map(p => p.id === id ? { ...p, priority: !p.priority } : p);
    savePayments();
};

function renderCalendar() {
    if(!calendarContainer) return;
    calendarContainer.innerHTML = '';
    let daysToRender = [];
    if (filterMode === 'week') {
        currentPeriodLabel.innerText = `${currentWeekStart.format('D MMMM')} - ${currentWeekStart.add(4, 'day').format('D MMMM YYYY')}`;
        for (let i = 0; i < 5; i++) daysToRender.push(currentWeekStart.add(i, 'day'));
    }

    let overallTotal = 0;
    daysToRender.forEach(day => {
        const dStr = day.format('YYYY-MM-DD');
        const dPayments = payments.filter(p => p.date === dStr);
        overallTotal += dPayments.reduce((sum, p) => sum + p.amount, 0);

        const dayCol = document.createElement('div');
        dayCol.className = 'calendar-day border-r border-slate-100 last:border-0 p-4 min-w-[200px] bg-white';
        const isToday = day.isSame(dayjs(), 'day');
        const isHoliday = allHolidays.includes(dStr);

        dayCol.innerHTML = `
            <div class="mb-4">
                <div class="flex items-center justify-between mb-1">
                    <span class="text-[10px] font-black tracking-[0.2em] uppercase ${isToday ? 'text-brand-600' : 'text-slate-400'}">${day.format('dddd')}</span>
                    ${isToday ? '<span class="flex h-2 w-2 rounded-full bg-brand-500 animate-pulse"></span>' : ''}
                </div>
                <div class="flex items-baseline gap-2">
                    <span class="text-3xl font-black ${isToday ? 'text-brand-600' : 'text-slate-900'}">${day.format('D')}</span>
                    <span class="text-xs font-bold text-slate-400 uppercase">${day.format('MMMM')}</span>
                </div>
                <div class="mt-2 text-[10px] font-black text-slate-400 bg-slate-50 px-2.5 py-1 rounded-lg border border-slate-100/50 inline-block uppercase">TOPLAM: ${formatTL(dPayments.reduce((s,p)=>s+p.amount,0))}</div>
                ${isHoliday ? '<div class="mt-2 mb-2 px-3 py-1.5 rounded-lg bg-red-100 text-red-600 text-xs font-black uppercase tracking-widest text-center animate-pulse border border-red-200">RESMİ TATİL</div>' : ''}
            </div>
            <div class="space-y-3 calendar-payments-list min-h-[150px]" data-date="${dStr}">
                ${dPayments.sort((a,b) => b.priority - a.priority).map(p => `
                    <div class="payment-card group relative p-3 rounded-xl border-l-[6px] shadow-sm hover:shadow-lg transition-all active:scale-[0.98] cursor-grab ${p.priority ? 'bg-red-100 ring-2 ring-red-200 shadow-md shadow-red-200/50' : 'bg-slate-50/30'}" style="border-left-color: ${categoryHexColors[p.category] || categoryHexColors['Diğer']}" data-id="${p.id}">
                        <div class="flex items-start justify-between mb-2">
                             <div class="flex items-center gap-1.5">
                                <i data-lucide="${categoryIcons[p.category] || 'credit-card'}" class="w-3.5 h-3.5 text-slate-400"></i>
                                <span class="${bankColors[p.bank] || 'bg-slate-500 text-white'} px-1.5 py-0.5 rounded text-[8px] font-black uppercase truncate max-w-[70px] shadow-sm">${p.bank}</span>
                                <span class="${categoryColors[p.category] || 'bg-slate-500 text-white'} px-1.5 py-0.5 rounded text-[8px] font-black uppercase shadow-sm">${p.category || 'Diğer'}</span>
                             </div>
                             <div class="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                                <button onclick="window.togglePriority('${p.id}')" class="p-1 hover:bg-red-50 rounded shadow-sm ${p.priority ? 'text-red-500' : 'text-slate-300 hover:text-red-500'}" title="Önemli İşaretle/Kaldır"><i data-lucide="triangle-alert" class="w-3.5 h-3.5"></i></button>
                                <button onclick="window.editPayment('${p.id}')" class="p-1 hover:bg-white rounded shadow-sm text-slate-400 hover:text-blue-500"><i data-lucide="edit-2" class="w-2.5 h-2.5"></i></button>
                                <button onclick="window.deletePayment('${p.id}')" class="p-1 hover:bg-white rounded shadow-sm text-slate-400 hover:text-red-500"><i data-lucide="trash-2" class="w-2.5 h-2.5"></i></button>
                             </div>
                        </div>
                        <h4 class="text-xs font-bold text-slate-800 mb-1 leading-tight">${p.title}</h4>
                        <div class="flex items-center justify-between">
                            <span class="text-sm font-black text-slate-900">${formatTL(p.amount)}</span>
                            ${p.priority ? '<span class="text-[10px] font-bold bg-red-500 text-white px-3 py-1 rounded-lg uppercase shadow-sm animate-pulse">ÖNEMLİ</span>' : ''}
                        </div>
                    </div>
                `).join('')}
            </div>
        `;
        calendarContainer.appendChild(dayCol);
        new Sortable(dayCol.querySelector('.calendar-payments-list'), {
            group: 'shared', animation: 150, ghostClass: 'opacity-20',
            onEnd: (evt) => {
                const itemEl = evt.item; const newDate = evt.to.getAttribute('data-date'); const pid = itemEl.getAttribute('data-id');
                payments = payments.map(p => p.id === pid ? { ...p, date: newDate } : p); savePayments();
            }
        });
    });
    if (overallTotalDisplay) overallTotalDisplay.innerText = formatTL(overallTotal);
    lucide.createIcons();
}

function renderAllPayments() {
    allPaymentsContent.innerHTML = '';
    const sorted = [...payments].sort((a,b) => dayjs(a.date).diff(dayjs(b.date)));
    
    // Totals logic
    const today = dayjs().startOf('day');
    const totalAll = payments.reduce((s, p) => s + p.amount, 0);
    const totalRemaining = payments.filter(p => dayjs(p.date).isAfter(today)).reduce((s, p) => s + p.amount, 0);
    
    const listTotalAllElem = document.getElementById('listTotalAll');
    const listTotalRemainingElem = document.getElementById('listTotalRemaining');
    if (listTotalAllElem) listTotalAllElem.innerText = formatTL(totalAll);
    if (listTotalRemainingElem) listTotalRemainingElem.innerText = formatTL(totalRemaining);

    const groups = {};
    sorted.forEach(p => {
        const key = p.groupId || `p_${p.id}`;
        if (!groups[key]) groups[key] = { payments: [], title: p.title.split(' (')[0], id: key, isGroup: !!p.groupId };
        groups[key].payments.push(p);
    });

    const threshold = today.subtract(1, 'month');
    const singleThreshold = today.subtract(10, 'days');
    
    const allGroups = Object.values(groups).filter(g=>g.isGroup);
    const activeGroups = allGroups.filter(g => {
        const lastPayment = g.payments.reduce((latest, p) => dayjs(p.date).isAfter(latest) ? dayjs(p.date) : latest, dayjs(0));
        return lastPayment.isAfter(threshold);
    });
    const completedGroups = allGroups.filter(g => {
        const lastPayment = g.payments.reduce((latest, p) => dayjs(p.date).isAfter(latest) ? dayjs(p.date) : latest, dayjs(0));
        return lastPayment.isSameOrBefore(threshold);
    });

    const allSingles = Object.values(groups).filter(g=>!g.isGroup);
    const activeSingles = allSingles.filter(g => dayjs(g.payments[0].date).isAfter(singleThreshold));
    const completedSingles = allSingles.filter(g => dayjs(g.payments[0].date).isSameOrBefore(singleThreshold));

    [ 
        {id:'installments', title:'TAKSİTLİ ÖDEMELER', list:activeGroups, color:'indigo'}, 
        {id:'single', title:'TEKLİ ÖDEMELER', list:activeSingles, color:'violet'},
        {id:'completed_singles', title:'ÖDEMESİ YAPILMIŞ', list:completedSingles, color:'emerald'},
        {id:'completed', title:'BİTEN KREDİLER', list:completedGroups, color:'slate'} 
    ].forEach(sec => {
        if (sec.list.length === 0) return;
        const isExp = expandedSections.has(sec.id);
        const secWrap = document.createElement('div');
        secWrap.className = 'mb-6';
        secWrap.innerHTML = `
            <div class="flex items-center justify-between gap-2 mb-4 px-2 py-3 bg-slate-50 border border-slate-100 rounded-xl cursor-pointer hover:bg-slate-100" onclick="window.toggleSection('${sec.id}')">
                <div class="flex items-center gap-3">
                    <div class="h-6 w-1 bg-${sec.color}-600 rounded-full"></div>
                    <h3 class="text-xs font-black text-slate-900 uppercase tracking-widest">${sec.title}</h3>
                    <span class="text-[10px] font-bold text-slate-400 bg-slate-200/50 px-2 rounded-full">${sec.list.length}</span>
                </div>
                <i data-lucide="${isExp ? 'chevron-up' : 'chevron-down'}" class="w-4 h-4 text-slate-400"></i>
            </div>
            <div class="${isExp ? 'block' : 'hidden'} animate-fade-in"></div>
        `;
        const listDiv = secWrap.lastElementChild;
        sec.list.forEach(group => {
            const wrap = document.createElement('div');
            wrap.className = 'border border-slate-200 rounded-2xl mb-4 bg-white shadow-sm overflow-hidden';
            if (sec.id === 'single' || sec.id === 'completed_singles') {
                const p = group.payments[0];
                wrap.innerHTML = `
                    <div class="p-4 flex items-center justify-between">
                        <div class="flex items-center gap-4">
                            <div class="w-10 h-10 flex items-center justify-center bg-violet-50 text-violet-600 rounded-xl"><i data-lucide="credit-card" class="w-5 h-5"></i></div>
                            <div><div class="flex items-center gap-2 text-[10px]"><span class="font-black text-slate-700 bg-slate-100 px-2 rounded uppercase">${dayjs(p.date).format('DD MMM YYYY')}</span><span class="${bankColors[p.bank] || 'bg-slate-500 text-white'} px-2 py-0.5 rounded text-[9px] font-black uppercase shadow-sm">${p.bank}</span><span class="${categoryColors[p.category] || 'bg-slate-500 text-white'} px-2 rounded font-black uppercase text-[9px] py-0.5">${p.category || 'Diğer'}</span></div><h4 class="font-extrabold text-slate-900 text-base">${p.title}</h4></div>
                        </div>
                        <div class="flex items-center gap-6"><div class="text-right"><p class="text-[10px] font-bold text-slate-400 uppercase">Tutar</p><p class="text-lg font-black text-slate-900">${formatTL(p.amount)}</p></div><div class="flex items-center gap-2"><button onclick="window.togglePriority('${p.id}')" class="p-2.5 hover:bg-red-50 rounded-xl border border-slate-100 ${p.priority ? 'text-red-500' : 'text-slate-400 hover:text-red-500'}" title="Önemli İşareti"><i data-lucide="triangle-alert" class="w-4 h-4"></i></button><button onclick="window.editPayment('${p.id}')" class="p-2.5 hover:bg-blue-50 rounded-xl border border-slate-100"><i data-lucide="edit-2" class="w-4 h-4 text-slate-400 hover:text-blue-600"></i></button><button onclick="window.deletePayment('${p.id}')" class="p-2.5 hover:bg-red-50 rounded-xl border border-slate-100"><i data-lucide="trash-2" class="w-4 h-4 text-slate-400 hover:text-red-600"></i></button></div></div>
                    </div>
                `;
            } else {
                const isGExp = expandedGroups.has(group.id); 
                const note = groupNotes[group.id] || '';
                const today = dayjs().startOf('day');
                const groupPayments = group.payments;
                const groupTotal = groupPayments.reduce((a,c)=>a+c.amount,0);
                const groupRemaining = groupPayments.filter(p => dayjs(p.date).isAfter(today)).reduce((a,c)=>a+c.amount,0);
                
                // Taksit sayacı (Ödenen / Toplam)
                const paidCount = groupPayments.filter(p => dayjs(p.date).isSameOrBefore(today)).length;
                const totalCount = groupPayments.length;

                wrap.innerHTML = `
                    <div class="p-5 cursor-pointer flex items-center justify-between bg-slate-50/40" onclick="window.toggleGroupDetails('${group.id}')">
                        <div class="flex items-center gap-4">
                            <div class="w-10 h-10 flex items-center justify-center bg-indigo-50 text-indigo-600 rounded-xl"><i data-lucide="${isGExp ? 'chevron-down':'chevron-right'}" class="w-5 h-5"></i></div>
                            <div><div class="flex items-center gap-2 text-[10px] mb-1"><span class="${bankColors[group.payments[0].bank] || 'bg-slate-500 text-white'} px-2 py-0.5 rounded-lg text-[9px] font-black uppercase shadow-sm">${group.payments[0].bank}</span><span class="bg-indigo-600 text-white px-2 rounded-lg font-black uppercase">${paidCount} / ${totalCount} TAKSİT</span><span class="${categoryColors[group.payments[0].category] || 'bg-slate-500 text-white'} px-2 py-0.5 text-[9px] rounded-lg font-black uppercase">${group.payments[0].category || 'Diğer'}</span></div><h4 class="font-black text-slate-900 text-lg uppercase leading-none">${group.title}</h4></div>
                        </div>
                        <div class="flex items-center gap-4">
                            <div class="text-right">
                                <p class="text-[10px] font-bold text-slate-400 uppercase">Toplam</p>
                                <p class="text-base font-black text-slate-900">${formatTL(groupTotal)}</p>
                                ${groupRemaining > 0 ? `<p class="text-[10px] font-black text-indigo-600 mt-0.5 px-2 bg-indigo-50 rounded-lg border border-indigo-100/50 inline-block">KALAN: ${formatTL(groupRemaining)}</p>` : ''}
                            </div>
                            <div class="flex items-center gap-2">
                                <button onclick="window.editGroup('${group.id}')" class="p-2 hover:bg-white rounded-lg border border-slate-100"><i data-lucide="edit-2" class="w-4 h-4 text-slate-400"></i></button>
                                <button onclick="window.deleteGroup('${group.id}')" class="p-2 hover:bg-white rounded-lg border border-slate-100"><i data-lucide="trash-2" class="w-4 h-4 text-slate-400"></i></button>
                            </div>
                        </div>
                    </div>
                    <div class="p-4 border-t border-slate-100 bg-slate-50/20"><textarea onchange="window.saveNote('${group.id}', this.value)" class="w-full p-2 bg-white border border-slate-200 rounded-lg text-sm" rows="2" placeholder="Grup notu...">${note}</textarea></div>
                    ${isGExp ? `<div class="p-4 border-t border-slate-100 bg-slate-50/10 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                        ${group.payments.map(p => `
                            <div class="bg-white border rounded-lg p-3 shadow-sm">
                                <div class="flex justify-between mb-2">
                                    <span class="text-[10px] font-black bg-slate-100 px-1.5 rounded">${dayjs(p.date).format('DD MMM YYYY')}</span>
                                    <div class="flex items-center gap-2">
                                        <button onclick="window.togglePriority('${p.id}')" title="Öncelik"><i data-lucide="triangle-alert" class="w-3 ${p.priority ? 'text-red-500' : 'text-slate-300 hover:text-red-500'}"></i></button>
                                        <button onclick="window.editPayment('${p.id}')" title="Düzenle"><i data-lucide="edit-2" class="w-3 text-slate-300 hover:text-blue-500"></i></button>
                                        <button onclick="window.deletePayment('${p.id}')" title="Sil"><i data-lucide="trash-2" class="w-3 text-slate-300 hover:text-red-500"></i></button>
                                    </div>
                                </div>
                                <h5 class="text-[10px] font-bold truncate uppercase text-slate-400">${p.title}</h5>
                                <p class="text-xs font-black">${formatTL(p.amount)}</p>
                            </div>
                        `).join('')}
                    </div>` : ''}
                `;
            }
            listDiv.appendChild(wrap);
        });
        allPaymentsContent.appendChild(secWrap);
    });
    lucide.createIcons();
}

window.renderMonthlyTable = function() {
    if (!monthlyTableContent) return;
    const m = parseInt(monthlyMonthSelect.value);
    const y = parseInt(monthlyYearSelect.value);
    const filtered = payments.filter(p => {
        const d = dayjs(p.date);
        return d.month() === m && d.year() === y;
    }).sort((a,b) => dayjs(a.date).diff(dayjs(b.date)));

    if (filtered.length === 0) {
        monthlyTableContent.innerHTML = '<div class="text-center py-10 text-slate-500 font-medium">Bu aya ait kayıt bulunamadı.</div>';
        return;
    }

    let html = `
        <table class="w-full text-left border-collapse">
            <thead>
                <tr class="bg-slate-100 text-slate-600 text-[10px] uppercase font-black tracking-wider">
                    <th class="p-3 rounded-l-xl">Tarih</th>
                    <th class="p-3">Açıklama</th>
                    <th class="p-3">Banka</th>
                    <th class="p-3">Kategori</th>
                    <th class="p-3 text-center">Öncelik</th>
                    <th class="p-3 rounded-r-xl text-right">Tutar</th>
                </tr>
            </thead>
            <tbody class="divide-y divide-slate-100">
    `;
    let total = 0;
    filtered.forEach(p => {
        total += p.amount;
        html += `
            <tr class="hover:bg-slate-50 transition-colors">
                <td class="p-3 text-xs font-bold text-slate-700 whitespace-nowrap">${dayjs(p.date).format('DD.MM.YYYY')}</td>
                <td class="p-3 text-sm font-bold text-slate-900">${p.title}</td>
                <td class="p-3"><span class="${bankColors[p.bank] || 'bg-slate-500 text-white'} px-2 py-0.5 rounded text-[9px] font-black uppercase shadow-sm">${p.bank}</span></td>
                <td class="p-3"><span class="${categoryColors[p.category] || 'bg-slate-500 text-white'} px-2 rounded font-black uppercase text-[9px] py-0.5">${p.category || 'Diğer'}</span></td>
                <td class="p-3 text-center">${p.priority ? '<span class="text-[8px] bg-red-100 text-red-600 px-2 py-1 rounded font-black">ÖNEMLİ</span>':'-'}</td>
                <td class="p-3 text-right font-black text-slate-900">${formatTL(p.amount)}</td>
            </tr>
        `;
    });
    html += `
            </tbody>
            <tfoot>
                <tr class="bg-slate-900 text-white">
                    <td colspan="5" class="p-4 rounded-l-xl text-right font-bold text-sm">TOPLAM HACİM:</td>
                    <td class="p-4 rounded-r-xl text-right font-black text-lg">${formatTL(total)}</td>
                </tr>
            </tfoot>
        </table>
    `;
    monthlyTableContent.innerHTML = html;
    if (overallTotalDisplay) overallTotalDisplay.innerText = formatTL(total);
};

window.exportToExcel = function(data, filename, titleStr = '', periodStr = '') {
    if(data.length === 0) return alert('Dışa aktarılacak veri yok.');
    
    // Create Header info rows
    const aoa = [
        [titleStr],
        [periodStr],
        [''], // Empty spacer
        ['Tarih', 'Aciklama', 'Banka', 'Kategori', 'Tutar', 'Not']
    ];

    data.forEach(p => {
        aoa.push([
            dayjs(p.date).format('DD.MM.YYYY'),
            p.title,
            p.bank,
            p.category,
            p.amount,
            p.note || ''
        ]);
    });

    const totalAmount = data.reduce((sum, p) => sum + p.amount, 0);
    aoa.push(['', '', '', 'GENEL TOPLAM:', totalAmount, '']);

    // Summary Table Data
    const categories = ['Kredi', 'Çek', 'Kira', 'Kredi Kartı', 'Fatura', 'Maaş', 'Diğer'];
    const summaryAoa = [['Kategori', 'Tutar']];
    categories.forEach(cat => {
        const total = data.filter(p => p.category === cat).reduce((sum, p) => sum + p.amount, 0);
        summaryAoa.push([cat, total]);
    });
    summaryAoa.push(['GENEL TOPLAM:', totalAmount]);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    // Add Summary table to Column H (index 7) starting row 4 (index 3)
    XLSX.utils.sheet_add_aoa(ws, summaryAoa, { origin: { r: 3, c: 7 } });
    
    const b = { top: { style: 'thin', color: { rgb: "000000" } }, bottom: { style: 'thin', color: { rgb: "000000" } }, left: { style: 'thin', color: { rgb: "000000" } }, right: { style: 'thin', color: { rgb: "000000" } } };
    
    const titleStyle = { font: { bold: true, size: 14 } };
    const periodStyle = { font: { italic: true, size: 10 }, color: { rgb: "666666" } };
    const hdrStyle = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "1E293B" } }, alignment: { horizontal: "center", vertical: "center" }, border: b };
    
    const greenStyle = { fill: { fgColor: { rgb: "D9EAD3" } }, alignment: { vertical: "center" }, border: b };
    const greenCurStyle = { fill: { fgColor: { rgb: "D9EAD3" } }, alignment: { horizontal: "right" }, border: b, numFmt: "#,##0.00\ \"\u20BA\"" };
    const blueStyle = { fill: { fgColor: { rgb: "CFE2F3" } }, alignment: { vertical: "center" }, border: b };
    const blueCurStyle = { fill: { fgColor: { rgb: "CFE2F3" } }, alignment: { horizontal: "right" }, border: b, numFmt: "#,##0.00\ \"\u20BA\"" };
    
    const totalBase = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "1E293B" } }, alignment: { vertical: "center" }, border: b };
    const totalCur = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "1E293B" } }, alignment: { horizontal: "right" }, border: b, numFmt: "#,##0.00\ \"\u20BA\"" };
    const totalRight = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "1E293B" } }, alignment: { horizontal: "right", vertical: "center" }, border: b };

    if (ws['A1']) ws['A1'].s = titleStyle;
    if (ws['A2']) ws['A2'].s = periodStyle;

    const range = XLSX.utils.decode_range(ws['!ref']);
    const tableHeaderRow = 3; // 0-indexed row 3 (which is Row 4 in Excel)
    
    for(let R = tableHeaderRow; R <= range.e.r; ++R) {
        for(let C = range.s.c; C <= range.e.c; ++C) {
            const addr = XLSX.utils.encode_cell({r:R, c:C});
            if(!ws[addr]) continue;

            // Main Table Styling (A-F)
            if (C <= 5) {
                const isHeader = (R === tableHeaderRow);
                const isTotal = (ws[XLSX.utils.encode_cell({r:R, c:3})] && ws[XLSX.utils.encode_cell({r:R, c:3})].v === 'GENEL TOPLAM:');
                const dataRowIndex = R - tableHeaderRow;
                const isEven = !isHeader && !isTotal && (dataRowIndex % 2 === 0);

                if (isHeader) {
                    ws[addr].s = hdrStyle;
                } else if (isTotal) {
                    if (C === 4) ws[addr].s = totalCur;
                    else if (C === 3) ws[addr].s = totalRight;
                    else ws[addr].s = totalBase;
                } else if (isEven) {
                    if (C === 4) ws[addr].s = blueCurStyle;
                    else ws[addr].s = blueStyle;
                } else {
                    if (C === 4) ws[addr].s = greenCurStyle;
                    else ws[addr].s = greenStyle;
                }
            }
            
            // Summary Table Styling (H-I)
            if (C >= 7 && C <= 8) {
                const isHeader = (R === tableHeaderRow);
                const isTotal = (ws[XLSX.utils.encode_cell({r:R, c:7})] && ws[XLSX.utils.encode_cell({r:R, c:7})].v === 'GENEL TOPLAM:');
                const summaryRowIndex = R - tableHeaderRow;
                const isEven = !isHeader && !isTotal && (summaryRowIndex % 2 === 0);

                if (isHeader) {
                    ws[addr].s = hdrStyle;
                } else if (isTotal) {
                    if (C === 8) ws[addr].s = totalCur;
                    else ws[addr].s = totalBase;
                } else if (isEven) {
                    if (C === 8) ws[addr].s = blueCurStyle;
                    else ws[addr].s = blueStyle;
                } else {
                    if (C === 8) ws[addr].s = greenCurStyle;
                }
            }
        }
    }
    ws['!cols'] = [ {wpx: 80}, {wpx: 220}, {wpx: 100}, {wpx: 120}, {wpx: 100}, {wpx: 250}, {wpx: 20}, {wpx: 120}, {wpx: 100} ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, ws, "Ödemeler");
    XLSX.writeFile(workbook, filename);
};

window.exportToPdf = function(data, filename, titleStr) {
    if(data.length === 0) return alert('Dışa aktarılacak veri yok.');
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('l');
    const sanitize = (t) => (t||'').toString().replace(/Ğ/g,'G').replace(/ğ/g,'g').replace(/Ü/g,'U').replace(/ü/g,'u').replace(/Ş/g,'S').replace(/ş/g,'s').replace(/İ/g,'I').replace(/ı/g,'i').replace(/Ö/g,'O').replace(/ö/g,'o').replace(/Ç/g,'C').replace(/ç/g,'c');
    
    doc.setFontSize(16);
    doc.text(sanitize(titleStr), 14, 15);
    
    const tableData = data.map(p => [
        dayjs(p.date).format('DD.MM.YYYY'),
        sanitize(p.title),
        sanitize(p.bank),
        sanitize(p.category),
        p.amount.toLocaleString('tr-TR', { minimumFractionDigits: 2 }) + ' TL',
        sanitize(p.note || '')
    ]);

    doc.autoTable({
        startY: 25,
        head: [['Tarih', 'Aciklama', 'Banka', 'Kategori', 'Tutar', 'Not']],
        body: tableData,
        foot: [['', '', '', 'Toplam:', data.reduce((s,x)=>s+x.amount,0).toLocaleString('tr-TR', { minimumFractionDigits: 2 }) + ' TL', '']],
        theme: 'striped',
        headStyles: { fillColor: [15, 23, 42] },
        footStyles: { fillColor: [15, 23, 42], fontStyle: 'bold' }
    });
    doc.save(filename);
};

    const showMonthGrid = (picker, ui, year) => {
        const grid = document.createElement('div');
        grid.className = 'datepicker-grid-selector animate-in fade-in zoom-in duration-200';
        const months = ['Oca', 'Şub', 'Mar', 'Nis', 'May', 'Haz', 'Tem', 'Ağu', 'Eyl', 'Eki', 'Kas', 'Ara'];
        grid.innerHTML = `<div class="datepicker-grid-selector-header"><span class="font-bold text-slate-700">${year}</span></div>`;
        months.forEach((m, idx) => {
            const item = document.createElement('div');
            item.className = 'datepicker-grid-item';
            item.innerText = m;
            item.onclick = (e) => { e.stopPropagation(); picker.gotoDate(dayjs().year(year).month(idx).toDate()); grid.remove(); };
            grid.appendChild(item);
        });
        ui.querySelector('.container__main').appendChild(grid);
    };

    const showYearGrid = (picker, ui, baseYear) => {
        const grid = document.createElement('div');
        grid.className = 'datepicker-grid-selector animate-in fade-in zoom-in duration-200';
        const startYear = baseYear - (baseYear % 12);
        grid.innerHTML = `
            <div class="datepicker-grid-selector-header">
                <button class="p-1 hover:bg-slate-200 rounded transition-colors" id="prevDecade"><i data-lucide="chevron-left" class="w-4 h-4"></i></button>
                <span class="font-bold text-slate-700">${startYear} - ${startYear + 11}</span>
                <button class="p-1 hover:bg-slate-200 rounded transition-colors" id="nextDecade"><i data-lucide="chevron-right" class="w-4 h-4"></i></button>
            </div>
        `;
        
        grid.querySelector('#prevDecade').onclick = (e) => { e.stopPropagation(); grid.remove(); showYearGrid(picker, ui, baseYear - 12); };
        grid.querySelector('#nextDecade').onclick = (e) => { e.stopPropagation(); grid.remove(); showYearGrid(picker, ui, baseYear + 12); };

        for (let y = startYear; y < startYear + 12; y++) {
            const item = document.createElement('div');
            item.className = 'datepicker-grid-item';
            item.innerText = y;
            item.onclick = (e) => { e.stopPropagation(); grid.remove(); showMonthGrid(picker, ui, y); };
            grid.appendChild(item);
        }
        ui.querySelector('.container__main').appendChild(grid);
        lucide.createIcons();
    };

    const opts = { 
        lang: 'tr-TR', 
        format: 'YYYY-MM-DD', 
        setup: (p) => { 
            p.on('render', (ui) => { 
                const monthLabel = ui.querySelector('.month-item-name');
                const yearLabel = ui.querySelector('.month-item-year');
                if (monthLabel) monthLabel.onclick = (e) => { e.stopPropagation(); const d = p.getDate(); showMonthGrid(p, ui, d ? d.getFullYear() : dayjs().year()); };
                if (yearLabel) yearLabel.onclick = (e) => { e.stopPropagation(); const d = p.getDate(); showYearGrid(p, ui, d ? d.getFullYear() : dayjs().year()); };

                ui.querySelectorAll('.day-item').forEach(day => { 
                    const d = dayjs(parseInt(day.dataset.time)); 
                    if (d.day()===0||d.day()===6||allHolidays.includes(d.format('YYYY-MM-DD'))) { 
                        day.style.color='#ef4444'; 
                        day.style.fontWeight='800'; 
                    } 
                }); 
            }); 
        } 
    };

document.addEventListener('DOMContentLoaded', () => {
    const loginOverlay = document.getElementById('loginOverlay');
    const loginForm = document.getElementById('loginForm');
    const loginPassword = document.getElementById('loginPassword');
    const loginError = document.getElementById('loginError');

    loginForm.onsubmit = async (e) => {
        e.preventDefault();
        const entered = loginPassword.value;
        if (VALID_PASSWORDS.includes(entered)) {
            const success = await loadEncryptedData(entered);
            if (success) {
                loginOverlay.classList.add('opacity-0', 'pointer-events-none');
                masterKey = entered;
                initApp();
            } else {
                loginError.classList.remove('hidden');
            }
        } else {
            loginError.classList.remove('hidden');
        }
    };

    function initApp() {
        calendarContainer = document.getElementById('calendarContainer');
        currentPeriodLabel = document.getElementById('currentPeriodLabel');
        overallTotalDisplay = document.getElementById('overallTotalDisplay');
        paymentModal = document.getElementById('paymentModal');
        paymentModalOverlay = document.getElementById('paymentModalOverlay');
        addPaymentBtn = document.getElementById('addPaymentBtn');
        closeModalBtn = document.getElementById('closeModalBtn');
        cancelPaymentBtn = document.getElementById('cancelPaymentBtn');
        paymentForm = document.getElementById('paymentForm');
        showAllPaymentsBtn = document.getElementById('showAllPaymentsBtn');
        allPaymentsModalOverlay = document.getElementById('allPaymentsModalOverlay');
        allPaymentsModal = document.getElementById('allPaymentsModal');
        closeAllPaymentsModalBtn = document.getElementById('closeAllPaymentsModalBtn');
        allPaymentsContent = document.getElementById('allPaymentsContent');
        confirmDeleteModalOverlay = document.getElementById('confirmDeleteModalOverlay');
        confirmDeleteModal = document.getElementById('confirmDeleteModal');
        deleteOnlyThisBtn = document.getElementById('deleteOnlyThisBtn');
        deleteAllGroupBtn = document.getElementById('deleteAllGroupBtn');
        cancelDeleteBtn = document.getElementById('cancelDeleteBtn');
        groupDeleteActions = document.getElementById('groupDeleteActions');
        singleDeleteActions = document.getElementById('singleDeleteActions');
        deleteModalTitle = document.getElementById('deleteModalTitle');
        deleteModalText = document.getElementById('deleteModalText');
        confirmSingleDeleteBtn = document.getElementById('confirmSingleDeleteBtn');
        showMonthlyTableBtn = document.getElementById('showMonthlyTableBtn');
        monthlyTableModalOverlay = document.getElementById('monthlyTableModalOverlay');
        monthlyTableModal = document.getElementById('monthlyTableModal');
        monthlyTableContent = document.getElementById('monthlyTableContent');
        closeMonthlyTableModalBtn = document.getElementById('closeMonthlyTableModalBtn');
        exportMonthlyExcelBtn = document.getElementById('exportMonthlyExcelBtn');
        exportMonthlyPdfBtn = document.getElementById('exportMonthlyPdfBtn');
        monthlyTableTitle = document.getElementById('monthlyTableTitle');
        monthlyMonthSelect = document.getElementById('monthlyMonthSelect');
        monthlyYearSelect = document.getElementById('monthlyYearSelect');
        
        initEventListeners();
        renderCalendar();
        lucide.createIcons();
    }

    function initEventListeners() {
        const exportExcelBtn = document.getElementById('exportExcelBtn');
        if(exportExcelBtn) exportExcelBtn.onclick = () => {
            let weeklyPayments = [];
            for (let i = 0; i < 5; i++) {
                const dStr = currentWeekStart.add(i, 'day').format('YYYY-MM-DD');
                weeklyPayments.push(...payments.filter(p => p.date === dStr));
            }
            const pStr = `Donem : ${currentWeekStart.format('DD.MM.YYYY')} - ${currentWeekStart.add(4, 'day').format('DD.MM.YYYY')}`;
            window.exportToExcel(weeklyPayments.sort((a,b)=>dayjs(a.date).diff(dayjs(b.date))), `haftalik_odemeler_${currentWeekStart.format('DD_MM')}.xlsx`, 'Haftalik Odeme Tablosu', pStr);
        };
        
        const exportPdfBtn = document.getElementById('exportPdfBtn');
        if(exportPdfBtn) exportPdfBtn.onclick = () => {
            let weeklyPayments = [];
            for (let i = 0; i < 5; i++) {
                const dStr = currentWeekStart.add(i, 'day').format('YYYY-MM-DD');
                weeklyPayments.push(...payments.filter(p => p.date === dStr));
            }
            window.exportToPdf(weeklyPayments.sort((a,b)=>dayjs(a.date).diff(dayjs(b.date))), `haftalik_odemeler_${currentWeekStart.format('DD_MM')}.pdf`, `Haftalik Odemeler (${currentWeekStart.format('DD.MM.YYYY')})`);
        };

        jumpToDatePicker = new Litepicker({ element: document.getElementById('jumpToDate'), ...opts, format:'DD.MM.YYYY', autoApply:true, setup:(p)=>{ opts.setup(p); p.on('selected', (d) => { currentWeekStart = dayjs(d.dateInstance || d).startOf('week'); renderCalendar(); }); } });
        paymentDatePicker = new Litepicker({ element: document.getElementById('payDate'), ...opts, format:'DD.MM.YYYY' });

        document.getElementById('payType').onchange = updatePaymentFormUI;
        addPaymentBtn.onclick = window.openModal;
        closeModalBtn.onclick = window.closeModal;
        cancelPaymentBtn.onclick = window.closeModal;
        showAllPaymentsBtn.onclick = () => { expandedSections.clear(); expandedGroups.clear(); allPaymentsModalOverlay.classList.remove('opacity-0','pointer-events-none'); setTimeout(() => { allPaymentsModal.classList.remove('scale-95','opacity-0'); renderAllPayments(); }, 10); };
        closeAllPaymentsModalBtn.onclick = () => { allPaymentsModal.classList.add('scale-95','opacity-0'); setTimeout(()=>allPaymentsModalOverlay.classList.add('opacity-0','pointer-events-none'),300); };
        showMonthlyTableBtn.onclick = () => { monthlyTableModalOverlay.classList.remove('opacity-0','pointer-events-none'); monthlyMonthSelect.value=dayjs().month(); monthlyYearSelect.value=dayjs().year(); setTimeout(()=>{ monthlyTableModal.classList.remove('scale-95','opacity-0'); renderMonthlyTable(); },10); };
        closeMonthlyTableModalBtn.onclick = () => { monthlyTableModal.classList.add('scale-95','opacity-0'); setTimeout(()=>monthlyTableModalOverlay.classList.add('opacity-0','pointer-events-none'),300); };
        
        exportMonthlyExcelBtn.onclick = () => {
            const m = parseInt(monthlyMonthSelect.value); const y = parseInt(monthlyYearSelect.value);
            const filtered = payments.filter(p => dayjs(p.date).month() === m && dayjs(p.date).year() === y).sort((a,b)=>dayjs(a.date).diff(dayjs(b.date)));
            const start = dayjs().year(y).month(m).startOf('month').format('DD.MM.YYYY');
            const end = dayjs().year(y).month(m).endOf('month').format('DD.MM.YYYY');
            window.exportToExcel(filtered, `odemeler_${m+1}_${y}.xlsx`, 'Aylik Odeme Tablosu', `Donem : ${start} - ${end}`);
        };
        exportMonthlyPdfBtn.onclick = () => {
            const m = parseInt(monthlyMonthSelect.value); const y = parseInt(monthlyYearSelect.value);
            const filtered = payments.filter(p => dayjs(p.date).month() === m && dayjs(p.date).year() === y).sort((a,b)=>dayjs(a.date).diff(dayjs(b.date)));
            window.exportToPdf(filtered, `odemeler_${m+1}_${y}.pdf`, `Aylik Odeme Tablosu (${m+1}/${y})`);
        };

        document.getElementById('monthlyPrevMonthBtn').onclick = () => { let m=parseInt(monthlyMonthSelect.value)-1; let y=parseInt(monthlyYearSelect.value); if(m<0){m=11;y--;} monthlyMonthSelect.value=m; monthlyYearSelect.value=y; renderMonthlyTable(); };
        document.getElementById('monthlyNextMonthBtn').onclick = () => { let m=parseInt(monthlyMonthSelect.value)+1; let y=parseInt(monthlyYearSelect.value); if(m>11){m=0;y++;} monthlyMonthSelect.value=m; monthlyYearSelect.value=y; renderMonthlyTable(); };
        monthlyMonthSelect.onchange = renderMonthlyTable;
        monthlyYearSelect.onchange = renderMonthlyTable;
        document.getElementById('prevWeekBtn').onclick = () => { currentWeekStart=currentWeekStart.subtract(1,'week'); renderCalendar(); };
        document.getElementById('nextWeekBtn').onclick = () => { currentWeekStart=currentWeekStart.add(1,'week'); renderCalendar(); };
        document.getElementById('todayBtn').onclick = () => { currentWeekStart=dayjs().startOf('week'); renderCalendar(); };

        // Cloud Settings Listeners
        const cloudSettingsBtn = document.getElementById('cloudSettingsBtn');
        const loginCloudBtn = document.getElementById('loginCloudBtn');
        const cloudSettingsModalOverlay = document.getElementById('cloudSettingsModalOverlay');
        const cloudSettingsModal = document.getElementById('cloudSettingsModal');
        const closeCloudSettingsBtn = document.getElementById('closeCloudSettingsBtn');
        const saveCloudSettingsBtn = document.getElementById('saveCloudSettingsBtn');
        const disconnectCloudBtn = document.getElementById('disconnectCloudBtn');
        const cloudTestResult = document.getElementById('cloudTestResult');

        const openCloudModal = () => {
            cloudSettingsModalOverlay.classList.remove('opacity-0', 'pointer-events-none');
            setTimeout(() => cloudSettingsModal.classList.remove('scale-95', 'opacity-0'), 10);
            
            if (cloudSettings) {
                document.getElementById('ghUsername').value = cloudSettings.ghUsername || '';
                document.getElementById('ghRepo').value = cloudSettings.ghRepo || '';
                document.getElementById('ghToken').value = cloudSettings.ghToken || '';
            }
        };

        if (cloudSettingsBtn) cloudSettingsBtn.onclick = openCloudModal;
        if (loginCloudBtn) loginCloudBtn.onclick = openCloudModal;

        const closeCloudModal = () => {
            cloudSettingsModal.classList.add('scale-95', 'opacity-0');
            setTimeout(() => cloudSettingsModalOverlay.classList.add('opacity-0', 'pointer-events-none'), 300);
        };

        if (closeCloudSettingsBtn) closeCloudSettingsBtn.onclick = closeCloudModal;

        if (saveCloudSettingsBtn) saveCloudSettingsBtn.onclick = async () => {
            const settings = {
                ghUsername: document.getElementById('ghUsername').value.trim(),
                ghRepo: document.getElementById('ghRepo').value.trim(),
                ghToken: document.getElementById('ghToken').value.trim()
            };

            if (!settings.ghUsername || !settings.ghRepo || !settings.ghToken) {
                alert("Lütfen tüm alanları doldurun.");
                return;
            }

            cloudTestResult.innerText = "Bağlantı test ediliyor...";
            cloudTestResult.className = "text-center py-2 px-3 rounded-lg text-xs font-bold bg-slate-100 text-slate-600 block";
            cloudTestResult.classList.remove('hidden');

            cloudSettings = settings; // Temporary set to test
            const test = await fetchGitHubFile();
            if (test) {
                localStorage.setItem('cloud_sync_settings', JSON.stringify(settings));
                cloudTestResult.innerText = "BAŞARILI: Bulut bağlantısı aktif!";
                cloudTestResult.className = "text-center py-2 px-3 rounded-lg text-xs font-bold bg-emerald-100 text-emerald-600 block";
                updateCloudStatus(true);
                setTimeout(closeCloudModal, 1500);
            } else {
                cloudSettings = null;
                cloudTestResult.innerText = "HATA: Bağlantı kurulamadı. Bilgileri kontrol edin.";
                cloudTestResult.className = "text-center py-2 px-3 rounded-lg text-xs font-bold bg-red-100 text-red-600 block";
                updateCloudStatus(false);
            }
        };

        if (disconnectCloudBtn) disconnectCloudBtn.onclick = () => {
            if (confirm("Bulut bağlantısını kesmek istediğinize emin misiniz? Veriler sadece yerelde kalacaktır.")) {
                localStorage.removeItem('cloud_sync_settings');
                cloudSettings = null;
                updateCloudStatus(false);
                closeCloudModal();
                window.location.reload();
            }
        };

        // Initial cloud status check
        if (cloudSettings) updateCloudStatus(true);

        const forceUpdateBtn = document.getElementById('forceUpdateBtn');
        const manualSyncBtn = document.getElementById('manualSyncBtn');

        if (manualSyncBtn) manualSyncBtn.onclick = () => {
            manualSyncBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> GÜNCELLENİYOR...';
            lucide.createIcons();
            // Sayfayı yenileyerek buluttaki en yeni veriyi çekmeyi zorla
            window.location.reload(true);
        };

        if (forceUpdateBtn) forceUpdateBtn.onclick = async () => {
            if (confirm("Uygulama önbelleği temizlenecek ve en yeni sürüm yüklenecek. Sayfa yenilenecektir. Devam edilsin mi?")) {
                if ('serviceWorker' in navigator) {
                    const registrations = await navigator.serviceWorker.getRegistrations();
                    for (let registration of registrations) { await registration.unregister(); }
                }
                if ('caches' in window) {
                    const cacheNames = await caches.keys();
                    for (let name of cacheNames) { await caches.delete(name); }
                }
                window.location.reload(true);
            }
        };

        paymentForm.onsubmit = (e) => {
            e.preventDefault(); const eid = document.getElementById('editingId').value; const egid = document.getElementById('editingGroupId').value;
            const title = document.getElementById('payTitle').value; const amt = parseFloat(document.getElementById('payAmount').value);
            const [ds, bn, cat, tp, inst, pNote] = [document.getElementById('payDate').value, document.getElementById('payBank').value, document.getElementById('payCategory').value, document.getElementById('payType').value, parseInt(document.getElementById('payInstallments').value)||1, document.getElementById('payNote').value];
            const parts = ds.split('.'); const dStr = parts.length===3 ? `${parts[2]}-${parts[1]}-${parts[0]}` : ds;

            if (egid) {
                payments = payments.map(p => (p.groupId === egid || (p.groupId === null && p.id === egid)) ? { ...p, title: p.groupId ? `${title} ${p.title.substring(p.title.lastIndexOf('('))}` : title, amount:amt, bank:bn, category:cat, note:pNote } : p);
            } else if (eid) {
                payments = payments.map(p => p.id === eid ? { ...p, title, amount:amt, bank:bn, category:cat, note:pNote, date:dStr } : p);
            } else {
                const gid = (tp === 'installments' && inst > 1) || (tp !== 'installments') ? crypto.randomUUID() : null;
                if (tp === 'first_working_day' || tp === 'last_working_day') {
                    const startMonth = dayjs().month();
                    const currentYear = dayjs().year();
                    for (let m = startMonth; m <= 11; m++) {
                        let d = dayjs().year(currentYear).month(m);
                        if (tp === 'first_working_day') d = getNextWorkDay(d.startOf('month'));
                        else d = getPrevWorkDay(d.endOf('month'));
                        payments.push({ id: crypto.randomUUID(), groupId: gid, title: `${title} (${d.format('MMMM YYYY')})`, amount: amt, bank: bn, category: cat, date: d.format('YYYY-MM-DD'), priority: false, note: pNote });
                    }
                } else {
                    const baseDate = dayjs(dStr);
                    for (let i = 0; i < inst; i++) {
                        let d = getNextWorkDay(baseDate.add(i, 'month'));
                        payments.push({ id: crypto.randomUUID(), groupId: gid, title: inst > 1 ? `${title} (${i+1}/${inst})` : title, amount: amt, bank: bn, category: cat, date: d.format('YYYY-MM-DD'), priority: false, note: pNote });
                    }
                }
            }
            savePayments(); window.closeModal();
        };

        confirmSingleDeleteBtn.onclick = deleteOnlyThisBtn.onclick = () => { payments = payments.filter(p => p.id !== pendingDeleteId); savePayments(); window.closeConfirmDeleteModal(); };
        deleteAllGroupBtn.onclick = () => { const p = payments.find(x => x.id === pendingDeleteId); if (p && p.groupId) payments = payments.filter(x => x.groupId !== p.groupId); else payments = payments.filter(x => x.id !== pendingDeleteId); savePayments(); window.closeConfirmDeleteModal(); };
        cancelDeleteBtn.onclick = window.closeConfirmDeleteModal;
        window.onkeydown = (e) => { if (e.key === 'Escape') { window.closeModal(); window.closeConfirmDeleteModal(); } };
    }
});
