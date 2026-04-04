// 1. Dependencies & Locale Initialization
dayjs.locale('tr');
dayjs.extend(window.dayjs_plugin_isSameOrBefore);
dayjs.extend(window.dayjs_plugin_isSameOrAfter);

// --- GÜVENLİK AYARLARI (Gelişmiş Güvenlik) ---
let masterKey = null; // Kodda şifre tutulmaz, giriş anında set edilir.
let payments = [];
let incomes = []; // Gelir verileri
let groupNotes = {};

const incomeCategories = [
    "Satışlardan Elde Edilen Nakit",
    "Döviz Kuru Değerleme Kazancı",
    "Faiz Gelirleri",
    "Alınan Krediler",
    "Ortaklardan Alınan Bedel",
    "Sulama gelirleri",
    "Diğer Ödemeler"
];

// --- BULUT SENKRONİZASYON AYARLARI ---
let cloudSettings = JSON.parse(localStorage.getItem('cloud_sync_settings') || 'null');
const CLOUD_FILENAME = 'odemeler_yedek_sifreli.json';

// --- EXCEL IMPORT LOGIC ---
window.downloadSampleExcel = function () {
    const cats = Object.keys(categoryColors);
    const banks = Object.keys(bankColors);
    const subCats = ["Stok", "Hizmet", "Kredi Ödemesi", "Banka Giderleri", "Akaryakıt", "Kum", "Beton", "Nakliye", "Araç Kira", "Ev Kira", "Araç Giderleri", "Maaş+SGK", "Vergiler", "Ortaklara Ödenen", "İş Kazası", "Kredi Kartı", "İSG Harcaması", "Trafik Cezaları", "Diğer"];

    // Ana Taslak + H Sütununda Bankalar + I Sütununda Kategoriler
    const data = [
        ['Tarih', 'Aciklama', 'Banka', 'Ödeme Türü', 'Tutar', 'Kategori', 'Not', 'GEÇERLİ BANKALAR', 'KATEGORİLER'],
        [dayjs().format('DD.MM.YYYY'), 'Örnek Ödeme', 'AKBANK', 'Fatura', 1500.50, 'Kum', 'Örnek Not', banks[0] || '', subCats[0] || ''],
    ];

    const maxLen = Math.max(banks.length, subCats.length, 20);
    for (let i = 1; i < maxLen; i++) {
        data.push(['', '', '', '', '', '', '', banks[i] || '', subCats[i] || '']);
    }

    const ws = XLSX.utils.aoa_to_sheet(data);

    // SÜTUN GENİŞLİKLERİ
    ws['!cols'] = [
        { wch: 11 }, { wch: 11 }, { wch: 11 }, { wch: 11 }, { wch: 11 }, { wch: 11 }, // A, B, C, D, E, F
        { wch: 25 },                                                   // G (Not: 25)
        { wch: 10 },                                                   // H (Boşluk)
        { wch: 23 }, { wch: 23 }, { wch: 23 }                          // I, J, K (Rehber: 23)
    ];

    // SATIR YÜKSEKLİKLERİ
    ws['!rows'] = [{ hpt: 27 }];
    for (let i = 1; i < 30; i++) ws['!rows'].push({ hpt: 20 });

    // Stiller
    const b = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    const headerStyle = {
        font: { bold: true, size: 16 },
        fill: { fgColor: { rgb: "B2B2B2" } }, // Daha koyu gri
        border: b,
        alignment: { horizontal: "center", vertical: "center" }
    };
    const helperHeaderStyle = {
        font: { bold: true, size: 16, name: "Calibri" },
        fill: { fgColor: { rgb: "B2B2B2" } },
        border: b,
        alignment: { horizontal: "center", vertical: "center" }
    };

    const blueZebra = { fill: { fgColor: { rgb: "CFE2F3" } }, border: b, alignment: { vertical: "center" } };
    const greenZebra = { fill: { fgColor: { rgb: "D9EAD3" } }, border: b, alignment: { vertical: "center" } };

    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const addr = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[addr]) ws[addr] = { v: "" };

            if (R === 0) {
                // Header (1. Satır)
                if (C === 8 || C === 9 || C === 10) ws[addr].s = helperHeaderStyle;
                else if (C < 7) ws[addr].s = headerStyle;
            } else if (R < 30) {
                // Zebra (2-30. Satır)
                if (C < 7 || C === 8 || C === 9 || C === 10) {
                    ws[addr].s = (R % 2 === 0) ? blueZebra : greenZebra;
                }
            }
        }
    }

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, ws, "Taslak");
    XLSX.writeFile(workbook, "odeme_sablonu.xlsx");
};

let currentExcelData = null;

window.handleExcelFileSelect = function (event) {
    const file = event.target.files[0];
    const display = document.getElementById('excelFileNameDisplay');
    const uploadBtn = document.getElementById('uploadExcelBtn');

    if (file) {
        display.innerHTML = `<span class="text-sm font-black text-emerald-600 flex items-center justify-center gap-2"><i data-lucide="check-circle" class="w-4 h-4"></i> ${file.name} SEÇİLDİ</span>`;
        uploadBtn.classList.remove('hidden');
        lucide.createIcons();

        const reader = new FileReader();
        reader.onload = (e) => {
            const binary = new Uint8Array(e.target.result);
            const workbook = XLSX.read(binary, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            currentExcelData = XLSX.utils.sheet_to_json(firstSheet);
        };
        reader.readAsArrayBuffer(file);
    } else {
        display.innerHTML = `<span class="text-xs font-bold text-slate-500">Dosya Seçin veya Sürükleyin</span>`;
        uploadBtn.classList.add('hidden');
        currentExcelData = null;
    }
};

function parseExcelDate(val) {
    if (!val) return null;
    if (typeof val === 'number') {
        // Excel seri tarihi
        try {
            return dayjs(XLSX.SSF.format('YYYY-MM-DD', val)).format('YYYY-MM-DD');
        } catch (e) { return null; }
    }
    // String tarih: 31.03.2026, 31/03/2026, 31,03,2026
    const s = val.toString().replace(/[,\/]/g, '.');
    const parts = s.split('.');
    if (parts.length === 3) {
        const d = parts[0].padStart(2, '0');
        const m = parts[1].padStart(2, '0');
        const y = parts[2].length === 2 ? '20' + parts[2] : parts[2];
        const formatted = `${y}-${m}-${d}`;
        if (dayjs(formatted).isValid()) return formatted;
    }
    const fallback = dayjs(val);
    return fallback.isValid() ? fallback.format('YYYY-MM-DD') : null;
}

window.processExcelUpload = function () {
    if (!currentExcelData) return;

    let importedCount = 0;
    currentExcelData.forEach(row => {
        const title = row['Aciklama'] || row['Açıklama'] || row['Başlık'] || row['Baslik'];
        const amount = parseFloat(row['Tutar'] || row['Miktar'] || row['Amount']);

        // Desteklenen sütun isimleri varyasyonları
        let bank = (row['Banka'] || row['Bank'] || row['Hesap'] || 'DİĞER').toString().trim().toUpperCase();
        let paymentType = (row['Ödeme Türü'] || row['Odeme Turu'] || row['Kategori'] || row['Category'] || row['Tür'] || 'Diğer').toString().trim();
        let category = (row['Kategori'] || row['Category'] || 'Diğer').toString().trim();

        // Eğer hem Kategori hem de Ödeme Türü sütunları varsa, yeni yapıya göre ayır
        if (row['Ödeme Türü'] && row['Kategori']) {
             paymentType = row['Ödeme Türü'].toString().trim();
             category = row['Kategori'].toString().trim();
        }

        // Boş gelirse Diğer yap
        if (!paymentType || paymentType === "") paymentType = "Diğer";
        if (!category || category === "") category = "Diğer";

        const note = row['Not'] || row['Açıklama'] || row['Memo'] || '';
        const dateStr = row['Tarih'] || row['Date'];

        // Ödeme Türü (eski adıyla kategori) normalleştirme
        const normalize = (s) => s.toLowerCase().replace(/ı/g, 'i').replace(/İ/g, 'i');
        const matchedCat = Object.keys(categoryColors).find(k => normalize(k) === normalize(paymentType));
        if (matchedCat) paymentType = matchedCat;

        // Banka normalleştirme
        const matchedBank = Object.keys(bankColors).find(k => k.toUpperCase() === bank);
        if (matchedBank) bank = matchedBank;

        let date = parseExcelDate(dateStr);
        if (date) date = getNextWorkDay(date).format('YYYY-MM-DD');

        if (title && !isNaN(amount) && date) {
            payments.push({
                id: 'p_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5),
                title, amount, bank, category: paymentType, subCategory: category, note,
                date, priority: false
            });
            importedCount++;
        }
    });

    if (importedCount > 0) {
        savePayments();
        alert(`${importedCount} adet ödeme başarıyla yüklendi!`);
        window.closeModal();
        // Reset UI
        document.getElementById('excelFileNameDisplay').innerHTML = `<span class="text-xs font-bold text-slate-500">Dosya Seçin veya Sürükleyin</span>`;
        document.getElementById('uploadExcelBtn').classList.add('hidden');
        currentExcelData = null;
        document.getElementById('excelImportInput').value = '';
    } else {
        alert("Geçerli veri bulunamadı. Lütfen tarihin gg.aa.yyyy olduğundan ve sütun başlıklarının doğru olduğundan emin olun.");
    }
};
// ------------------------------------

async function loadEncryptedData(enteredPass) {
    try {
        let encSticky = null;

        // 1. ÖNCE BULUTTAN ÇEKMEYİ DENE (Eğer ayarlar varsa)
        if (cloudSettings && cloudSettings.ghToken) {
            try {
                const cloudData = await fetchGitHubFile();
                if (cloudData && cloudData.encrypted) {
                    encPayments = cloudData.payments_enc;
                    encIncomes = cloudData.incomes_enc;
                    encNotes = cloudData.grup_notlari_enc;
                    encSticky = cloudData.sticky_note_enc;
                    // Başarılıysa locale de yedekle
                    localStorage.setItem('odemeler_enc', encPayments);
                    if (encIncomes) localStorage.setItem('gelirler_enc', encIncomes);
                    localStorage.setItem('grup_notlari_enc', encNotes);
                    if (encSticky) localStorage.setItem('sticky_note_enc', encSticky);
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
            encSticky = localStorage.getItem('sticky_note_enc');
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
                        encSticky = data.sticky_note_enc;
                    }
                }
            } catch (e) { }
        }

        if (!encPayments) return true; // Hiç veri yoksa boş başla

        // --- ŞİFRE ÇÖZME VE KURTARMA DENEMESİ ---
        let bytesPay;
        try {
            bytesPay = CryptoJS.AES.decrypt(encPayments, enteredPass);
            const decryptedPay = bytesPay.toString(CryptoJS.enc.Utf8);
            
            // Eğer bulut verisi bu şifreyle çözülmüyorsa, yerel veriyi dene!
            if (!decryptedPay) {
                console.warn("Bulut verisi sifreyle uyusmadi, yerel yedek deneniyor...");
                const localEnc = localStorage.getItem('odemeler_enc');
                if (localEnc && localEnc !== encPayments) {
                    const bytesLocal = CryptoJS.AES.decrypt(localEnc, enteredPass);
                    const decLocal = bytesLocal.toString(CryptoJS.enc.Utf8);
                    if (decLocal) {
                        payments = JSON.parse(decLocal);
                        const ln = localStorage.getItem('grup_notlari_enc');
                        if (ln) groupNotes = JSON.parse(CryptoJS.AES.decrypt(ln, enteredPass).toString(CryptoJS.enc.Utf8));
                        const li = localStorage.getItem('gelirler_enc');
                        if (li) incomes = JSON.parse(CryptoJS.AES.decrypt(li, enteredPass).toString(CryptoJS.enc.Utf8));
                        masterKey = enteredPass;
                        return true; 
                    }
                }
                return false; 
            }
            
            payments = JSON.parse(decryptedPay);
        } catch(err) {
            return false;
        }

        if (encIncomes) {
            try {
                const bytesInc = CryptoJS.AES.decrypt(encIncomes, enteredPass);
                const decInc = bytesInc.toString(CryptoJS.enc.Utf8);
                if (decInc) incomes = JSON.parse(decInc);
            } catch (e) { }
        }

        if (encNotes) {
            const bytesNotes = CryptoJS.AES.decrypt(encNotes, enteredPass);
            groupNotes = JSON.parse(bytesNotes.toString(CryptoJS.enc.Utf8));
        }

        if (encSticky) {
            try {
                const decSticky = CryptoJS.AES.decrypt(encSticky, enteredPass).toString(CryptoJS.enc.Utf8);
                const stickyData = JSON.parse(decSticky);
                const stickyNote = document.getElementById('stickyNoteText');
                if (stickyNote) {
                    stickyNote.innerHTML = stickyData.content || '';
                }
            } catch (e) {
                console.error("Sticky note load error:", e);
            }
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
    const encIncomes = CryptoJS.AES.encrypt(JSON.stringify(incomes), masterKey).toString();
    const encNotes = CryptoJS.AES.encrypt(JSON.stringify(groupNotes), masterKey).toString();

    localStorage.setItem('odemeler_enc', encPay);
    localStorage.setItem('gelirler_enc', encIncomes);
    localStorage.setItem('grup_notlari_enc', encNotes);

    // NOT DEFTERİ VERİSİ
    const stickyNote = document.getElementById('stickyNoteText');
    const noteContent = stickyNote ? stickyNote.innerHTML : '';
    const encSticky = CryptoJS.AES.encrypt(JSON.stringify({ content: noteContent }), masterKey).toString();
    localStorage.setItem('sticky_note_enc', encSticky);

    const backupData = {
        payments_enc: encPay,
        incomes_enc: encIncomes,
        grup_notlari_enc: encNotes,
        sticky_note_enc: encSticky,
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
        } catch (err) { }
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
    'Maaş+SGK': 'users',
    'Diğer': 'layers'
};

const categoryColors = {
    'Kredi': 'bg-rose-500 text-white',
    'Çek': 'bg-sky-500 text-white',
    'Kira': 'bg-emerald-500 text-white',
    'Kredi Kartı': 'bg-amber-500 text-white',
    'Fatura': 'bg-orange-500 text-white',
    'Maaş+SGK': 'bg-teal-500 text-white',
    'Diğer': 'bg-slate-500 text-white'
};

const categoryHexColors = {
    'Kredi': '#f43f5e',
    'Çek': '#0ea5e9',
    'Kira': '#10b981',
    'Kredi Kartı': '#f59e0b',
    'Fatura': '#f97316',
    'Maaş+SGK': '#14b8a6',
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

// --- YARDIMCI GÖRÜNÜM FONKSİYONLARI ---
function getCategoryColorClass(cat) {
    if (!cat) return categoryColors['Diğer'];
    const matched = Object.keys(categoryColors).find(k => k.toLowerCase() === cat.trim().toLowerCase());
    return matched ? categoryColors[matched] : categoryColors['Diğer'];
}

function getCategoryHexColor(cat) {
    if (!cat) return categoryHexColors['Diğer'];
    const matched = Object.keys(categoryHexColors).find(k => k.toLowerCase() === cat.trim().toLowerCase());
    return matched ? categoryHexColors[matched] : categoryHexColors['Diğer'];
}

function getCategoryIcon(cat) {
    if (!cat) return categoryIcons['Diğer'];
    const matched = Object.keys(categoryIcons).find(k => k.toLowerCase() === cat.trim().toLowerCase());
    return matched ? categoryIcons[matched] : (categoryIcons['Diğer'] || 'credit-card');
}

function getBankColorClass(bank) {
    if (!bank) return bankColors['DİĞER'];
    const matched = Object.keys(bankColors).find(k => k.toUpperCase() === bank.trim().toUpperCase());
    return matched ? bankColors[matched] : bankColors['DİĞER'];
}

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

window.saveNote = function (id, note) {
    groupNotes[id] = note;
    const groupItems = payments.filter(p => p.groupId === id);
    if (groupItems.length > 0) groupItems.forEach(p => groupNotes[p.id] = note);
    localStorage.setItem('grup_notlari', JSON.stringify(groupNotes));
}

window.toggleSection = function (type) {
    if (expandedSections.has(type)) expandedSections.delete(type);
    else expandedSections.add(type);
    renderAllPayments();
};

window.toggleGroupDetails = function (id) {
    if (expandedGroups.has(id)) expandedGroups.delete(id);
    else expandedGroups.add(id);
    renderAllPayments();
};

window.openModal = function () {
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
    const pSubCat = document.getElementById('paySubCategory');
    if (pSubCat) pSubCat.value = 'Diğer';
    const n = document.getElementById('payNote');
    if (n) n.value = '';
    updatePaymentFormUI();
}

window.closeModal = function () {
    if (!paymentModal) return;
    paymentModal.classList.add('scale-95', 'opacity-0');
    setTimeout(() => paymentModalOverlay.classList.add('opacity-0', 'pointer-events-none'), 300);
}

window.editPayment = function (id) {
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

    const pSubCat = document.getElementById('paySubCategory');
    if (pSubCat) pSubCat.value = pay.subCategory || 'Diğer';

    const n = document.getElementById('payNote');
    if (n) n.value = pay.note || '';
    updatePaymentFormUI();
}

window.editGroup = function (groupId) {
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

    const pSubCat = document.getElementById('paySubCategory');
    if (pSubCat) pSubCat.value = first.subCategory || 'Diğer';

    const n = document.getElementById('payNote');
    if (n) n.value = first.note || '';
    updatePaymentFormUI();
}

window.deletePayment = function (id) {
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

window.deleteGroup = function (groupId) {
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

window.closeConfirmDeleteModal = function () {
    confirmDeleteModal.classList.add('scale-95', 'opacity-0');
    setTimeout(() => confirmDeleteModalOverlay.classList.add('opacity-0', 'pointer-events-none'), 300);
}

window.togglePriority = function (id) {
    payments = payments.map(p => p.id === id ? { ...p, priority: !p.priority } : p);
    savePayments();
};

function renderCalendar() {
    if (!calendarContainer) return;
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
        const isToday = day.isSame(dayjs(), 'day');
        const isHoliday = allHolidays.includes(dStr);

        const dayCol = document.createElement('div');
        dayCol.className = `calendar-day border-r border-slate-200 last:border-r-0 p-4 min-w-[200px] transition-all duration-300 ${isToday ? 'bg-blue-50/70 ring-4 ring-inset ring-blue-100 scale-[1.01] z-10 shadow-xl' : 'bg-white/40'}`;

        dayCol.innerHTML = `
            <div class="mb-4 pb-4 border-b-2 border-slate-900">
                <div class="flex items-center justify-between mb-1">
                    <span class="text-[10px] font-black tracking-[0.2em] uppercase ${isToday ? 'text-blue-600' : 'text-slate-400'}">${day.format('dddd')}</span>
                    ${isToday ? '<span class="text-[13px] font-black bg-red-600 text-white px-3 py-1 rounded-lg shadow-lg animate-pulse tracking-tighter">BUGÜN</span>' : ''}
                </div>
                <div class="flex items-baseline gap-2">
                    <span class="text-3xl font-black ${isToday ? 'text-blue-900' : 'text-slate-900'}">${day.format('D')}</span>
                    <span class="text-xs font-bold text-slate-400 uppercase">${day.format('MMMM')}</span>
                </div>
                <div class="mt-2 text-[10px] font-black text-slate-400 bg-slate-50 px-2.5 py-1 rounded-lg border border-slate-100/50 inline-block uppercase">TOPLAM: ${formatTL(dPayments.reduce((s, p) => s + p.amount, 0))}</div>
                ${isHoliday ? '<div class="mt-2 mb-2 px-3 py-1.5 rounded-lg bg-red-100 text-red-600 text-xs font-black uppercase tracking-widest text-center animate-pulse border border-red-200">RESMİ TATİL</div>' : ''}
            </div>
            <div class="space-y-3 calendar-payments-list min-h-[150px]" data-date="${dStr}">
                ${dPayments.sort((a, b) => b.priority - a.priority).map(p => `
                    <div class="payment-card group relative p-3 pl-5 rounded-2xl border border-slate-100 shadow-sm hover:shadow-lg transition-all active:scale-[0.98] cursor-grab ${p.priority ? 'bg-red-100 ring-4 ring-red-200 shadow-xl shadow-red-200/50' : 'bg-white'}" data-id="${p.id}">
                        <!-- Sol kavisli kalın bar -->
                        <div class="absolute left-0 top-0 bottom-0 w-2.5 rounded-l-2xl" style="background-color: ${getCategoryHexColor(p.category)}"></div>
                        <div class="flex items-start justify-between mb-2">
                             <div class="flex items-center gap-1.5">
                                <i data-lucide="${getCategoryIcon(p.category)}" class="w-3.5 h-3.5 ${p.priority ? 'text-red-500' : 'text-slate-400'}"></i>
                                <span class="${getBankColorClass(p.bank)} px-1.5 py-0.5 rounded text-[8px] font-black uppercase truncate max-w-[70px] shadow-sm">${p.bank}</span>
                                <span class="${getCategoryColorClass(p.category)} px-1.5 py-0.5 rounded text-[8px] font-black uppercase shadow-sm">${p.category || 'Diğer'}</span>
                             </div>
                             <div class="flex items-center gap-1 ${p.priority ? 'opacity-100' : 'opacity-0 group-hover:opacity-100'} transition-opacity">
                                <button onclick="window.togglePriority('${p.id}')" class="p-1 hover:bg-white rounded shadow-sm ${p.priority ? 'text-red-500' : 'text-slate-300 hover:text-red-500'}" title="Önemli İşaretle/Kaldır"><i data-lucide="triangle-alert" class="w-3.5 h-3.5"></i></button>
                                <button onclick="window.editPayment('${p.id}')" class="p-1 hover:bg-white rounded shadow-sm ${p.priority ? 'text-slate-400' : 'text-slate-400 hover:text-blue-500'}"><i data-lucide="edit-2" class="w-2.5 h-2.5"></i></button>
                                <button onclick="window.deletePayment('${p.id}')" class="p-1 hover:bg-white rounded shadow-sm ${p.priority ? 'text-slate-400' : 'text-slate-400 hover:text-red-500'}"><i data-lucide="trash-2" class="w-2.5 h-2.5"></i></button>
                             </div>
                        </div>
                        <h4 class="text-xs font-bold ${p.priority ? 'text-red-900' : 'text-slate-800'} mb-1 leading-tight uppercase">${p.title}${p.note ? ' - ' + p.note : ''}</h4>
                        <div class="flex items-center justify-between">
                            <span class="text-sm font-black ${p.priority ? 'text-red-950' : 'text-slate-900'}">${formatTL(p.amount)}</span>
                            ${p.priority ? '<span class="text-[10px] font-bold bg-red-600 text-white px-3 py-1 rounded-lg uppercase shadow-lg animate-pulse">ÖNEMLİ</span>' : ''}
                        </div>
                    </div>
                `).join('')}
            </div>
        `;
        calendarContainer.appendChild(dayCol);
        new Sortable(dayCol.querySelector('.calendar-payments-list'), {
            group: 'shared', animation: 150, ghostClass: 'opacity-20',
            onEnd: (evt) => {
                const pid = evt.item.getAttribute('data-id');
                const newDate = evt.to.getAttribute('data-date');

                // 1. Tarihi Güncelle
                payments = payments.map(p => p.id === pid ? { ...p, date: newDate } : p);

                // 2. Dikey Sırayı Kaydet (Aynı gün içindeki sıra)
                const itemsInCol = Array.from(evt.to.querySelectorAll('.payment-card'));
                const orderedIds = itemsInCol.map(el => el.getAttribute('data-id'));

                // O günün verilerini ayır ve yeni sırayla tekrar birleştir
                const otherPayments = payments.filter(p => !orderedIds.includes(p.id));
                const movedPaymentsOrdered = orderedIds.map(id => payments.find(p => p.id === id)).filter(Boolean);
                
                payments = [...otherPayments, ...movedPaymentsOrdered];

                savePayments();
                renderCalendar(); // Toplamları ve UI'ı güncelle
            }
        });
    });
    if (overallTotalDisplay) overallTotalDisplay.innerText = formatTL(overallTotal);
    lucide.createIcons();
}

function renderAllPayments() {
    allPaymentsContent.innerHTML = '';
    const sorted = [...payments].sort((a, b) => dayjs(a.date).diff(dayjs(b.date)));

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

    const allGroups = Object.values(groups).filter(g => g.isGroup);
    const activeGroups = allGroups.filter(g => {
        const lastPayment = g.payments.reduce((latest, p) => dayjs(p.date).isAfter(latest) ? dayjs(p.date) : latest, dayjs(0));
        return lastPayment.isAfter(threshold);
    });
    const completedGroups = allGroups.filter(g => {
        const lastPayment = g.payments.reduce((latest, p) => dayjs(p.date).isAfter(latest) ? dayjs(p.date) : latest, dayjs(0));
        return lastPayment.isSameOrBefore(threshold);
    });

    const allSingles = Object.values(groups).filter(g => !g.isGroup);
    const activeSingles = allSingles.filter(g => dayjs(g.payments[0].date).isAfter(singleThreshold));
    const completedSingles = allSingles.filter(g => dayjs(g.payments[0].date).isSameOrBefore(singleThreshold));

    [
        { id: 'installments', title: 'TAKSİTLİ ÖDEMELER', list: activeGroups, color: 'indigo' },
        { id: 'single', title: 'TEKLİ ÖDEMELER', list: activeSingles, color: 'violet' },
        { id: 'completed_singles', title: 'ÖDEMESİ YAPILMIŞ', list: completedSingles, color: 'emerald' },
        { id: 'completed', title: 'BİTEN KREDİLER', list: completedGroups, color: 'slate' }
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
                wrap.className = `border rounded-2xl mb-4 shadow-xl overflow-hidden transition-all ${p.priority ? 'bg-red-100 border-red-200 ring-2 ring-red-100' : 'bg-white border-slate-200 shadow-sm'}`;
                wrap.innerHTML = `
                    <div class="p-4 flex items-center justify-between">
                        <div class="flex items-center gap-4">
                            <div class="w-10 h-10 flex items-center justify-center ${p.priority ? 'bg-red-200 text-red-600' : 'bg-violet-50 text-violet-600'} rounded-xl"><i data-lucide="credit-card" class="w-5 h-5"></i></div>
                            <div><div class="flex items-center gap-2 text-[10px]"><span class="font-black ${p.priority ? 'text-red-800 bg-red-200' : 'text-slate-700 bg-slate-100'} px-2 rounded uppercase">${dayjs(p.date).format('DD MMM YYYY')}</span><span class="${getBankColorClass(p.bank)} px-2 py-0.5 rounded text-[9px] font-black uppercase shadow-sm">${p.bank}</span><span class="${getCategoryColorClass(p.category)} px-2 rounded font-black uppercase text-[9px] py-0.5">${p.category || 'Diğer'}</span></div><h4 class="font-extrabold ${p.priority ? 'text-red-900' : 'text-slate-900'} text-base uppercase">${p.title}</h4></div>
                        </div>
                        <div class="flex items-center gap-6">
                            <div class="text-right">
                                <p class="text-[10px] font-bold ${p.priority ? 'text-red-500' : 'text-slate-400'} uppercase">Tutar</p>
                                <p class="text-lg font-black ${p.priority ? 'text-red-950' : 'text-slate-900'}">${formatTL(p.amount)}</p>
                            </div>
                            <div class="flex items-center gap-2">
                                <button onclick="window.togglePriority('${p.id}')" class="p-2.5 hover:bg-white rounded-xl border ${p.priority ? 'border-red-100 text-red-500' : 'border-slate-100 text-slate-400 hover:text-red-500'}" title="Önemli İşareti"><i data-lucide="triangle-alert" class="w-4 h-4"></i></button>
                                <button onclick="window.editPayment('${p.id}')" class="p-2.5 hover:bg-white rounded-xl border ${p.priority ? 'border-red-100 text-slate-400' : 'border-slate-100 text-slate-400 hover:text-blue-600'}"><i data-lucide="edit-2" class="w-4 h-4"></i></button>
                                <button onclick="window.deletePayment('${p.id}')" class="p-2.5 hover:bg-white rounded-xl border ${p.priority ? 'border-red-100 text-slate-400' : 'border-slate-100 text-slate-400 hover:text-red-600'}"><i data-lucide="trash-2" class="w-4 h-4"></i></button>
                            </div>
                        </div>
                    </div>
                `;
            } else {
                const isGExp = expandedGroups.has(group.id);
                const note = groupNotes[group.id] || '';
                const today = dayjs().startOf('day');
                const groupPayments = group.payments;
                const groupTotal = groupPayments.reduce((a, c) => a + c.amount, 0);
                const groupRemaining = groupPayments.filter(p => dayjs(p.date).isAfter(today)).reduce((a, c) => a + c.amount, 0);

                // Taksit sayacı (Ödenen / Toplam)
                const paidCount = groupPayments.filter(p => dayjs(p.date).isSameOrBefore(today)).length;
                const totalCount = groupPayments.length;

                wrap.innerHTML = `
                    <div class="p-5 cursor-pointer flex items-center justify-between bg-slate-50/40" onclick="window.toggleGroupDetails('${group.id}')">
                        <div class="flex items-center gap-4">
                            <div class="w-10 h-10 flex items-center justify-center bg-indigo-50 text-indigo-600 rounded-xl"><i data-lucide="${isGExp ? 'chevron-down' : 'chevron-right'}" class="w-5 h-5"></i></div>
                            <div><div class="flex items-center gap-2 text-[10px] mb-1"><span class="${getBankColorClass(group.payments[0].bank)} px-2 py-0.5 rounded-lg text-[9px] font-black uppercase shadow-sm">${group.payments[0].bank}</span><span class="bg-indigo-600 text-white px-2 rounded-lg font-black uppercase">${paidCount} / ${totalCount} TAKSİT</span><span class="${getCategoryColorClass(group.payments[0].category)} px-2 py-0.5 text-[9px] rounded-lg font-black uppercase">${group.payments[0].category || 'Diğer'}</span></div><h4 class="font-black text-slate-900 text-lg uppercase leading-none">${group.title}</h4></div>
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

window.renderMonthlyTable = function () {
    if (!monthlyTableContent) return;
    const m = parseInt(monthlyMonthSelect.value);
    const y = parseInt(monthlyYearSelect.value);
    const filtered = payments.filter(p => {
        const d = dayjs(p.date);
        return d.month() === m && d.year() === y;
    }).sort((a, b) => dayjs(a.date).diff(dayjs(b.date)));

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
                <td class="p-3"><span class="${getBankColorClass(p.bank)} px-2 py-0.5 rounded text-[9px] font-black uppercase shadow-sm">${p.bank}</span></td>
                <td class="p-3"><span class="${getCategoryColorClass(p.category)} px-2 rounded font-black uppercase text-[9px] py-0.5">${p.category || 'Diğer'}</span></td>
                <td class="p-3 text-center">${p.priority ? '<span class="text-[8px] bg-red-600 text-white px-2 py-1 rounded font-black shadow-sm animate-pulse">ÖNEMLİ</span>' : '-'}</td>
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

window.exportToExcel = function (dataInput, filename, titleStr = '', periodStr = '', showBankDetails = false) {
    let data = dataInput;
    let isMonthly = (titleStr === 'AYLIK ÖDEME TABLOSU');
    
    // Eğer isMonthly true ise veya data bir dizi değilse (eski çağrılar için fallback)
    if (!Array.isArray(data)) {
        isMonthly = true; // İlk parametre eğer data değilse (örn isMonthly=true gelmişse)
    }

    // Ay ve Yıl Belirleme (Sadece modal açıkken oradakini al, yoksa bugünü al)
    const mSelect = document.getElementById('monthlyMonthSelect');
    const ySelect = document.getElementById('monthlyYearSelect');
    const modalOverlay = document.getElementById('monthlyTableModalOverlay');
    
    // Modal açık değilse (overlay-hidden ise) her zaman bugünün ayını al!
    const isModalVisible = modalOverlay && !modalOverlay.classList.contains('pointer-events-none');
    
    // ARTIK BAŞLIKTAN BAĞIMSIZ: Eğer modal açıksa oradaki tarihi, değilse bugünü baz al.
    const month = (isModalVisible && mSelect) ? parseInt(mSelect.value) : dayjs().month();
    const year = (isModalVisible && ySelect) ? parseInt(ySelect.value) : dayjs().year();

    if (isMonthly && !Array.isArray(dataInput)) {
        data = payments.filter(p => {
            const d = dayjs(p.date);
            return d.month() === month && d.year() === year;
        }).sort((a, b) => dayjs(a.date).diff(dayjs(b.date)));
    } else if (!isMonthly && !Array.isArray(dataInput)) {
        const start = dayjs().startOf('week');
        const end = dayjs().endOf('week');
        data = payments.filter(p => {
            const d = dayjs(p.date);
            return d.isSameOrAfter(start) && d.isSameOrBefore(end);
        }).sort((a, b) => dayjs(a.date).diff(dayjs(b.date)));
    }

    if (data.length === 0 && incomes.length === 0) return alert('Dışa aktarılacak veri bulunamadı.');

    // Create Header info rows
    const aoa = [
        [titleStr || "MUHASEBE TAKİP - ÖDEME RAPORU"],
        [periodStr || ("Dönem: " + (month + 1) + "/" + year)],
        [''], // Empty spacer
        ['Tarih', 'Aciklama', 'Banka', 'Ödeme Türü', 'Tutar', 'Kategori', 'Not']
    ];

    data.forEach(p => {
        aoa.push([
            dayjs(p.date).format('DD.MM.YYYY'),
            p.title,
            p.bank,
            p.category,
            p.amount,
            p.subCategory || 'Diğer',
            p.note || ''
        ]);
    });

    const totalAmount = data.reduce((sum, p) => sum + p.amount, 0);
    aoa.push(['', '', '', 'GENEL TOPLAM:', totalAmount, '', '']);

    // Summary Table Data
    const categories = ['Kredi', 'Çek', 'Kira', 'Kredi Kartı', 'Fatura', 'Maaş+SGK', 'Diğer'];
    const summaryAoa = [['Kategori', 'Tutar']];
    categories.forEach(cat => {
        const total = data.filter(p => p.category === cat).reduce((sum, p) => sum + p.amount, 0);
        summaryAoa.push([cat, total]);
    });
    summaryAoa.push(['GENEL TOPLAM:', totalAmount]);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    // Add Summary table to Column I (index 8) starting row 4 (index 3)
    XLSX.utils.sheet_add_aoa(ws, summaryAoa, { origin: { r: 3, c: 8 } });

    const b = { top: { style: 'thin', color: { rgb: "000000" } }, bottom: { style: 'thin', color: { rgb: "000000" } }, left: { style: 'thin', color: { rgb: "000000" } }, right: { style: 'thin', color: { rgb: "000000" } } };

    const titleStyle = { font: { bold: true, size: 14 } };
    const periodStyle = { font: { italic: true, size: 10 }, color: { rgb: "666666" } };
    const hdrStyle = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "475569" } }, alignment: { horizontal: "center", vertical: "center" }, border: b };

    const greenStyle = { fill: { fgColor: { rgb: "D9EAD3" } }, alignment: { vertical: "center" }, border: b };
    const greenCurStyle = { fill: { fgColor: { rgb: "D9EAD3" } }, alignment: { horizontal: "right" }, border: b, numFmt: "#,##0.00\ \"\u20BA\"" };
    const blueStyle = { fill: { fgColor: { rgb: "CFE2F3" } }, alignment: { vertical: "center" }, border: b };
    const blueCurStyle = { fill: { fgColor: { rgb: "CFE2F3" } }, alignment: { horizontal: "right" }, border: b, numFmt: "#,##0.00\ \"\u20BA\"" };

    const totalBase = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "475569" } }, alignment: { vertical: "center" }, border: b };
    const totalCur = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "475569" } }, alignment: { horizontal: "right" }, border: b, numFmt: "#,##0.00\ \"\u20BA\"" };
    const totalRight = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "475569" } }, alignment: { horizontal: "right", vertical: "center" }, border: b };

    if (ws['A1']) ws['A1'].s = titleStyle;
    if (ws['A2']) ws['A2'].s = periodStyle;

    const range = XLSX.utils.decode_range(ws['!ref']);
    const tableHeaderRow = 3; // 0-indexed row 3 (which is Row 4 in Excel)

    for (let R = tableHeaderRow; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const addr = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[addr]) continue;

            // Main Table Styling (A-G)
            if (C <= 6) {
                const isHeader = (R === tableHeaderRow);
                const isTotal = (ws[XLSX.utils.encode_cell({ r: R, c: 3 })] && ws[XLSX.utils.encode_cell({ r: R, c: 3 })].v === 'GENEL TOPLAM:');
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

            // Summary Table Styling (I-J)
            if (C >= 8 && C <= 9) {
                const isHeader = (R === tableHeaderRow);
                const isTotal = (ws[XLSX.utils.encode_cell({ r: R, c: 8 })] && ws[XLSX.utils.encode_cell({ r: R, c: 8 })].v === 'GENEL TOPLAM:');
                const summaryRowIndex = R - tableHeaderRow;
                const isEven = !isHeader && !isTotal && (summaryRowIndex % 2 === 0);

                if (isHeader) {
                    ws[addr].s = hdrStyle;
                } else if (isTotal) {
                    if (C === 9) ws[addr].s = totalCur;
                    else ws[addr].s = totalBase;
                } else if (isEven) {
                    if (C === 9) ws[addr].s = blueCurStyle;
                    else ws[addr].s = blueStyle;
                } else {
                    if (C === 9) ws[addr].s = greenCurStyle;
                    else ws[addr].s = greenStyle;
                }
            }
        }
    }
    ws['!cols'] = [{ wpx: 80 }, { wpx: 220 }, { wpx: 100 }, { wpx: 120 }, { wpx: 100 }, { wpx: 150 }, { wpx: 250 }, { wpx: 20 }, { wpx: 120 }, { wpx: 100 }];

    const workbook = XLSX.utils.book_new();

    // --- SHEET 1: ÖDEMELER ---
    XLSX.utils.book_append_sheet(workbook, ws, "Ödemeler");

    // --- SHEET 2: GELİRLER ---
    
    // --- GELİRLER SHEET (Kesin Seçili Ay ve Yıl Filtresi) ---
    const periodIncomes = incomes.filter(inc => {
        const d = dayjs(inc.date);
        return d.month() === month && d.year() === year;
    }).sort((a,b) => dayjs(a.date).diff(dayjs(b.date)));

    const incAoa = [
        ["BAŞLIK", "TUTAR", "BANKA", "KATEGORİ", "TARİH", "NOT"],
    ];
    let totalInc = 0;
    periodIncomes.forEach(i => {
        totalInc += i.amount;
        incAoa.push([i.title, i.amount, i.bank, i.category, dayjs(i.date).format('DD.MM.YYYY'), i.note || '']);
    });
    incAoa.push(["GENEL TOPLAM:", totalInc, "", "", "", ""]);

    const wsInc = XLSX.utils.aoa_to_sheet(incAoa);
    // Gelirler tablosu stili
    const rangeInc = XLSX.utils.decode_range(wsInc['!ref']);
    for(let R=0; R <= rangeInc.e.r; R++) {
        for(let C=0; C <= rangeInc.e.c; C++) {
            const addr = XLSX.utils.encode_cell({r:R, c:C});
            if(!wsInc[addr]) continue;
            wsInc[addr].s = (R === 0) ? hdrStyle : (C === 1 ? greenCurStyle : greenStyle);
            if(R === rangeInc.e.r) {
                wsInc[addr].s = (C === 1) ? totalCur : totalBase;
            }
        }
    }
    wsInc['!cols'] = [{wpx:200}, {wpx:100}, {wpx:100}, {wpx:150}, {wpx:100}, {wpx:200}];
    XLSX.utils.book_append_sheet(workbook, wsInc, "Gelirler");

    // --- SHEET 3: NAKİT AKIŞ (Özet) ---
    const flowAoa = [
        ["NAKİT AKIŞ ÖZETİ (" + (month+1) + "/" + year + ")", "", "", "", ""],
        ["", "", "", "", ""],
        ["GİDER KATEGORİSİ", "TOPLAM GİDER", "", "GELİR KATEGORİSİ", "TOPLAM GELİR"],
    ];

    // Gider Kategorileri (Tam Liste - Sadece Dropdown Seçenekleri)
    const expenseCats = [
        'Stok', 'Hizmet', 'Kredi Ödemesi', 'Banka Giderleri', 'Akaryakıt', 
        'Kum', 'Beton', 'Nakliye', 'Araç Kira', 'Ev Kira', 'Araç Giderleri', 
        'Maaş+SGK', 'Vergiler', 'Ortaklara Ödenen', 'İş Kazası', 'Kredi Kartı', 
        'İSG Harcaması', 'Trafik Cezaları', 'Diğer'
    ];
    // Gelir Kategorileri
    const incomeCats = [...incomeCategories];
    
    const maxRows = Math.max(expenseCats.length, incomeCats.length);
    for(let i=0; i < maxRows; i++) {
        const row = [];
        // Gider tarafı
        if(expenseCats[i]) {
            // Kullanıcının "F Sütunu" dediği alt kategorilere (subCategory) göre filtrele
            const sum = data.filter(p => (p.subCategory || 'Diğer') === expenseCats[i]).reduce((s,p) => s + p.amount, 0);
            row.push(expenseCats[i], sum);
        } else {
            row.push("", "");
        }
        row.push(""); // Boş orta sütun
        // Gelir tarafı
        if(incomeCats[i]) {
            const sum = periodIncomes.filter(inc => inc.category === incomeCats[i]).reduce((s,inc) => s + inc.amount, 0);
            row.push(incomeCats[i], sum);
        } else {
            row.push("", "");
        }
        flowAoa.push(row);
    }

    flowAoa.push(["TOPLAM GİDER:", totalAmount, "", "TOPLAM GELİR:", totalInc]);
    flowAoa.push(["", "", "", "", ""]);
    flowAoa.push(["", "", "", "NET DURUM:", totalInc - totalAmount]);

    const wsFlow = XLSX.utils.aoa_to_sheet(flowAoa);
    // Nakit Akış Stil
    const rangeFlow = XLSX.utils.decode_range(wsFlow['!ref']);
    for(let R=0; R <= rangeFlow.e.r; R++) {
        for(let C=0; C <= rangeFlow.e.c; C++) {
            const addr = XLSX.utils.encode_cell({r:R, c:C});
            if(!wsFlow[addr]) continue;
            if(R === 0) wsFlow[addr].s = titleStyle;
            else if(R === 2) wsFlow[addr].s = hdrStyle;
            else if(C === 1 || C === 4) {
                wsFlow[addr].s = (R >= rangeFlow.e.r - 2) ? totalCur : blueCurStyle;
                // Net durum özel renk
                if (R === rangeFlow.e.r && C === 4) {
                    const diff = totalInc - totalAmount;
                    wsFlow[addr].s = { ...totalCur, fill: { fgColor: { rgb: diff >= 0 ? "D9EAD3" : "F4CCCC" } }, font: { bold: true, color: { rgb: diff >= 0 ? "38761D" : "990000" } } };
                }
            }
            else if(C === 0 || C === 3) wsFlow[addr].s = (R >= rangeFlow.e.r - 2) ? totalBase : blueStyle;
        }
    }
    wsFlow['!cols'] = [{wpx:150}, {wpx:100}, {wpx:30}, {wpx:150}, {wpx:100}];
    XLSX.utils.book_append_sheet(workbook, wsFlow, "Nakit Akış");

    if (showBankDetails) {
        // Banka detayları kaldırıldı.
    }

    // DOSYAYI İNDİR
    try {
        const finalName = filename || `aylik_muhasebe_raporu_${dayjs().format('MM_YYYY')}.xlsx`;
        XLSX.writeFile(workbook, finalName);
    } catch (err) {
        console.error("Excel indirme hatası:", err);
        alert("Excel oluşturulurken bir hata oluştu.");
    }
};

window.exportToPdf = function (data, filename, titleStr) {
    if (data.length === 0) return alert('Dışa aktarılacak veri yok.');
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('l');
    const sanitize = (t) => (t || '').toString().replace(/Ğ/g, 'G').replace(/ğ/g, 'g').replace(/Ü/g, 'U').replace(/ü/g, 'u').replace(/Ş/g, 'S').replace(/ş/g, 's').replace(/İ/g, 'I').replace(/ı/g, 'i').replace(/Ö/g, 'O').replace(/ö/g, 'o').replace(/Ç/g, 'C').replace(/ç/g, 'c');

    doc.setFontSize(16);
    doc.text(sanitize(titleStr), 14, 15);

    const tableData = data.map(p => [
        dayjs(p.date).format('DD.MM.YYYY'),
        sanitize(p.title),
        sanitize(p.bank),
        sanitize(p.category), // Ödeme Türü
        sanitize(p.subCategory || 'Diğer'), // Kategori
        p.amount.toLocaleString('tr-TR', { minimumFractionDigits: 2 }) + ' TL',
        sanitize(p.note || '')
    ]);

    doc.autoTable({
        startY: 25,
        head: [['Tarih', 'Aciklama', 'Banka', 'Odeme Turu', 'Kategori', 'Tutar', 'Not']],
        body: tableData,
        foot: [['', '', '', '', 'Toplam:', data.reduce((s, x) => s + x.amount, 0).toLocaleString('tr-TR', { minimumFractionDigits: 2 }) + ' TL', '']],
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
                if (d.day() === 0 || d.day() === 6 || allHolidays.includes(d.format('YYYY-MM-DD'))) {
                    day.style.color = '#ef4444';
                    day.style.fontWeight = '800';
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

    // Cloud Settings Listeners (Moved here to work before login)
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

    const addIncomeBtn = document.getElementById('addIncomeBtn');
    if (addIncomeBtn) addIncomeBtn.onclick = window.openIncomeModal;
    const closeIncomeModalBtn = document.getElementById('closeIncomeModalBtn');
    if (closeIncomeModalBtn) closeIncomeModalBtn.onclick = window.closeIncomeModal;

    window.incDatePicker = new Litepicker({
        element: document.getElementById('incDate'),
        lang: 'tr-TR',
        format: 'DD.MM.YYYY',
        autoApply: true
    });

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

    // OTOMATİK VERİ YÜKLEME VE BENİ HATIRLA SİSTEMİ
    async function autoLoad() {
        console.log("Güvenli giriş kontrol ediliyor...");

        // Tarayıcı hafızasında kayıtlı şifre var mı?
        const savedKey = localStorage.getItem('saved_master_key');

        if (savedKey) {
            const success = await loadEncryptedData(savedKey);
            if (success) {
                console.log("Beni Hatırla: Otomatik giriş başarılı.");
                masterKey = savedKey;
                if (loginOverlay) loginOverlay.classList.add('hidden');
                initApp();
                return;
            } else {
                localStorage.removeItem('saved_master_key'); // Eski/Hatalı kayıtlı şifreyi sil
            }
        }

        // Eğer kayıtlı şifre yoksa veya hatalıysa giriş ekranını göster
        if (loginOverlay) loginOverlay.classList.remove('hidden');
    }

    autoLoad();

    // Giriş Formu Yönetimi
    loginForm.onsubmit = async (e) => {
        e.preventDefault();
        const entered = loginPassword.value;
        const rememberCheckbox = document.getElementById('rememberMe');
        const remember = rememberCheckbox ? rememberCheckbox.checked : false;

        if (loginError) loginError.classList.add('hidden');

        // Şifreyi doğrudan veri çözümü için dene
        const success = await loadEncryptedData(entered);
        if (success) {
            masterKey = entered;
            if (remember) {
                localStorage.setItem('saved_master_key', entered);
            }
            if (loginOverlay) loginOverlay.classList.add('hidden');
            initApp();
        } else {
            if (loginError) loginError.classList.remove('hidden');
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
        
        // Excel Butonundaki Tarih Etiketini Güncelle
        const excelDateLabel = document.getElementById('excelBtnDateLabel');
        if (excelDateLabel) {
            const currentMonthName = dayjs().format('MMMM').toUpperCase();
            const currentYearValue = dayjs().format('YYYY');
            excelDateLabel.innerText = `${currentMonthName} ${currentYearValue}`;
        }

        lucide.createIcons();

        // OTOMATİK BULUT KAYIT SİSTEMİ (Geri Sayımlı)
        let countdownSeconds = 120;
        const countdownLabel = document.getElementById('autoSaveCountdownLabel');

        setInterval(() => {
            countdownSeconds--;
            if (countdownSeconds <= 0) {
                if (masterKey && cloudSettings && cloudSettings.ghToken) {
                    console.log("Otomatik bulut yedeklemesi başlatıldı...");
                    savePayments();
                }
                countdownSeconds = 120;
            }
            if (countdownLabel) {
                countdownLabel.innerText = `KALAN SÜRE : ${countdownSeconds}s`;
            }
        }, 1000);
    }

    function initEventListeners() {
        // Anasayfa Yeni Excel Butonu
        const mainExportExcelBtn = document.getElementById('mainExportExcelBtn');
        if (mainExportExcelBtn) mainExportExcelBtn.onclick = () => window.exportToExcel(true, 'aylik_muhasebe_raporu.xlsx', 'AYLIK ÖDEME TABLOSU');

        // Aylık Tablo Modalı Butonları
        if (exportMonthlyExcelBtn) exportMonthlyExcelBtn.onclick = () => window.exportToExcel(true, 'aylik_muhasebe_raporu.xlsx', 'AYLIK ÖDEME TABLOSU');
        if (exportMonthlyPdfBtn) exportMonthlyPdfBtn.onclick = () => {
            const m = parseInt(monthlyMonthSelect.value);
            const y = parseInt(monthlyYearSelect.value);
            const filtered = payments.filter(p => {
                const d = dayjs(p.date);
                return d.month() === m && d.year() === y;
            }).sort((a, b) => dayjs(a.date).diff(dayjs(b.date)));
            window.exportToPdf(filtered, `aylik_odemeler_${m + 1}_${y}.pdf`, `${m + 1} / ${y} Aylık Ödeme Tablosu`);
        };

        jumpToDatePicker = new Litepicker({ element: document.getElementById('jumpToDate'), ...opts, format: 'DD.MM.YYYY', autoApply: true, setup: (p) => { opts.setup(p); p.on('selected', (d) => { currentWeekStart = dayjs(d.dateInstance || d).startOf('week'); renderCalendar(); }); } });
        paymentDatePicker = new Litepicker({ element: document.getElementById('payDate'), ...opts, format: 'DD.MM.YYYY' });

        document.getElementById('payType').onchange = updatePaymentFormUI;
        addPaymentBtn.onclick = window.openModal;
        closeModalBtn.onclick = window.closeModal;
        cancelPaymentBtn.onclick = window.closeModal;
        showAllPaymentsBtn.onclick = () => { expandedSections.clear(); expandedGroups.clear(); allPaymentsModalOverlay.classList.remove('opacity-0', 'pointer-events-none'); setTimeout(() => { allPaymentsModal.classList.remove('scale-95', 'opacity-0'); renderAllPayments(); }, 10); };
        closeAllPaymentsModalBtn.onclick = () => { allPaymentsModal.classList.add('scale-95', 'opacity-0'); setTimeout(() => allPaymentsModalOverlay.classList.add('opacity-0', 'pointer-events-none'), 300); };
        showMonthlyTableBtn.onclick = () => { monthlyTableModalOverlay.classList.remove('opacity-0', 'pointer-events-none'); monthlyMonthSelect.value = dayjs().month(); monthlyYearSelect.value = dayjs().year(); setTimeout(() => { monthlyTableModal.classList.remove('scale-95', 'opacity-0'); renderMonthlyTable(); }, 10); };
        closeMonthlyTableModalBtn.onclick = () => { monthlyTableModal.classList.add('scale-95', 'opacity-0'); setTimeout(() => monthlyTableModalOverlay.classList.add('opacity-0', 'pointer-events-none'), 300); };

        exportMonthlyExcelBtn.onclick = () => {
            const m = parseInt(monthlyMonthSelect.value); const y = parseInt(monthlyYearSelect.value);
            const filtered = payments.filter(p => dayjs(p.date).month() === m && dayjs(p.date).year() === y).sort((a, b) => dayjs(a.date).diff(dayjs(b.date)));
            const start = dayjs().year(y).month(m).startOf('month').format('DD.MM.YYYY');
            const end = dayjs().year(y).month(m).endOf('month').format('DD.MM.YYYY');
            window.exportToExcel(filtered, `odemeler_${m + 1}_${y}.xlsx`, 'Aylik Odeme Tablosu', `Donem : ${start} - ${end}`, true);
        };
        exportMonthlyPdfBtn.onclick = () => {
            const m = parseInt(monthlyMonthSelect.value); const y = parseInt(monthlyYearSelect.value);
            const filtered = payments.filter(p => dayjs(p.date).month() === m && dayjs(p.date).year() === y).sort((a, b) => dayjs(a.date).diff(dayjs(b.date)));
            window.exportToPdf(filtered, `odemeler_${m + 1}_${y}.pdf`, `Aylik Odeme Tablosu (${m + 1}/${y})`);
        };

        document.getElementById('monthlyPrevMonthBtn').onclick = () => { let m = parseInt(monthlyMonthSelect.value) - 1; let y = parseInt(monthlyYearSelect.value); if (m < 0) { m = 11; y--; } monthlyMonthSelect.value = m; monthlyYearSelect.value = y; renderMonthlyTable(); };
        document.getElementById('monthlyNextMonthBtn').onclick = () => { let m = parseInt(monthlyMonthSelect.value) + 1; let y = parseInt(monthlyYearSelect.value); if (m > 11) { m = 0; y++; } monthlyMonthSelect.value = m; monthlyYearSelect.value = y; renderMonthlyTable(); };
        monthlyMonthSelect.onchange = renderMonthlyTable;
        if (monthlyYearSelect) monthlyYearSelect.onchange = renderMonthlyTable;

        const uploadExcelBtn = document.getElementById('uploadExcelBtn');
        if (uploadExcelBtn) uploadExcelBtn.onclick = window.processExcelUpload;

        document.getElementById('prevWeekBtn').onclick = () => { currentWeekStart = currentWeekStart.subtract(1, 'week'); renderCalendar(); };
        document.getElementById('nextWeekBtn').onclick = () => { currentWeekStart = currentWeekStart.add(1, 'week'); renderCalendar(); };
        document.getElementById('todayBtn').onclick = () => { currentWeekStart = dayjs().startOf('week'); renderCalendar(); };


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
            const [ds, bn, cat, tp, inst, pNote, subCat] = [document.getElementById('payDate').value, document.getElementById('payBank').value, document.getElementById('payCategory').value, document.getElementById('payType').value, parseInt(document.getElementById('payInstallments').value) || 1, document.getElementById('payNote').value, document.getElementById('paySubCategory').value];
            const parts = ds.split('.'); const dStr = parts.length === 3 ? `${parts[2]}-${parts[1]}-${parts[0]}` : ds;

            if (egid) {
                payments = payments.map(p => (p.groupId === egid || (p.groupId === null && p.id === egid)) ? { ...p, title: p.groupId ? `${title} ${p.title.substring(p.title.lastIndexOf('('))}` : title, amount: amt, bank: bn, category: cat, subCategory: subCat, note: pNote } : p);
            } else if (eid) {
                payments = payments.map(p => p.id === eid ? { ...p, title, amount: amt, bank: bn, category: cat, subCategory: subCat, note: pNote, date: dStr } : p);
            } else {
                const gid = (tp === 'installments' && inst > 1) || (tp !== 'installments') ? crypto.randomUUID() : null;
                if (tp === 'first_working_day' || tp === 'last_working_day') {
                    const startMonth = dayjs().month();
                    const currentYear = dayjs().year();
                    for (let m = startMonth; m <= 11; m++) {
                        let d = dayjs().year(currentYear).month(m);
                        if (tp === 'first_working_day') d = getNextWorkDay(d.startOf('month'));
                        else d = getPrevWorkDay(d.endOf('month'));
                        payments.push({ id: crypto.randomUUID(), groupId: gid, title: `${title} (${d.format('MMMM YYYY')})`, amount: amt, bank: bn, category: cat, subCategory: subCat, date: d.format('YYYY-MM-DD'), priority: false, note: pNote });
                    }
                } else {
                    const baseDate = dayjs(dStr);
                    for (let i = 0; i < inst; i++) {
                        let d = getNextWorkDay(baseDate.add(i, 'month'));
                        payments.push({ id: crypto.randomUUID(), groupId: gid, title: inst > 1 ? `${title} (${i + 1}/${inst})` : title, amount: amt, bank: bn, category: cat, subCategory: subCat, date: d.format('YYYY-MM-DD'), priority: false, note: pNote });
                    }
                }
            }
            savePayments(); window.closeModal();
        };

        confirmSingleDeleteBtn.onclick = deleteOnlyThisBtn.onclick = () => { payments = payments.filter(p => p.id !== pendingDeleteId); savePayments(); window.closeConfirmDeleteModal(); };
        deleteAllGroupBtn.onclick = () => { const p = payments.find(x => x.id === pendingDeleteId); if (p && p.groupId) payments = payments.filter(x => x.groupId !== p.groupId); else payments = payments.filter(x => x.id !== pendingDeleteId); savePayments(); window.closeConfirmDeleteModal(); };
        cancelDeleteBtn.onclick = window.closeConfirmDeleteModal;
        // NOT DEFTERİ ETKİLEŞİMİ (MANUEL KAYIT)
        const stickyNote = document.getElementById('stickyNoteText');
        const saveNoteBtn = document.getElementById('saveNoteBtn');
        const boldNoteBtn = document.getElementById('boldNoteBtn');
        const strikeNoteBtn = document.getElementById('strikeNoteBtn');

        if (boldNoteBtn) boldNoteBtn.onclick = () => { document.execCommand('bold', false, null); };
        if (strikeNoteBtn) strikeNoteBtn.onclick = () => { document.execCommand('strikethrough', false, null); };

        if (stickyNote) {
            stickyNote.addEventListener('paste', (e) => {
                e.preventDefault();
                const text = e.clipboardData.getData('text/plain');
                document.execCommand('insertText', false, text);
            });
        }

        if (saveNoteBtn && stickyNote) {
            saveNoteBtn.onclick = async () => {
                const originalIcon = saveNoteBtn.innerHTML;
                saveNoteBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin text-indigo-600"></i>';
                lucide.createIcons();

                // Hemen kaydet
                await savePayments();

                // Başarı ikonu göster
                saveNoteBtn.innerHTML = '<i data-lucide="check" class="w-4 h-4 text-emerald-600"></i>';
                lucide.createIcons();

                setTimeout(() => {
                    saveNoteBtn.innerHTML = originalIcon;
                    lucide.createIcons();
                }, 2000);
            };
        }

        window.onkeydown = (e) => {
            if (e.key === 'Escape') {
                window.closeModal();
                window.closeConfirmDeleteModal();
                if (window.closeAllPaymentsModalBtn) window.closeAllPaymentsModalBtn.click();
                if (window.closeMonthlyTableModalBtn) window.closeMonthlyTableModalBtn.click();
                if (window.closeCloudModal) window.closeCloudModal();
            }
        };

        // ÇIKIŞTA KAYDETME VE UYARI SİSTEMİ
        window.onbeforeunload = (e) => {
            savePayments(); // Tarayıcı izin verdiği ölçüde hızlıca yerel/bulut kaydı tetikle
            const msg = "Tüm değişiklikler kaydedildi mi? Çıkış yapmak istediğinize emin misiniz?";
            e.returnValue = msg;
            return msg;
        };
    }
});
window.openIncomeModal = function () {
    const overlay = document.getElementById("incomeModalOverlay");
    const modal = document.getElementById("incomeModal");
    if (!overlay || !modal) return;
    overlay.classList.remove("opacity-0", "pointer-events-none");
    setTimeout(() => modal.classList.remove("scale-95", "opacity-0"), 10);
    document.getElementById("incomeForm").reset();
    document.getElementById("editingIncomeId").value = "";
    const today = dayjs().format("DD.MM.YYYY");
    document.getElementById("incDate").value = today;
    if (window.incDatePicker) window.incDatePicker.setDate(today);
    renderIncomes();
};

window.closeIncomeModal = function () {
    const overlay = document.getElementById("incomeModalOverlay");
    const modal = document.getElementById("incomeModal");
    if (!overlay || !modal) return;
    modal.classList.add("scale-95", "opacity-0");
    setTimeout(() => overlay.classList.add("opacity-0", "pointer-events-none"), 300);
};

window.saveIncome = function () {
    const eid = document.getElementById("editingIncomeId").value;
    const title = document.getElementById("incTitle").value;
    const amt = parseFloat(document.getElementById("incAmount").value);
    const dateStr = document.getElementById("incDate").value;
    const bank = document.getElementById("incBank").value;
    const cat = document.getElementById("incCategory").value;
    const note = document.getElementById("incNote").value;
    
    if (!title || isNaN(amt)) {
        alert("Lütfen başlık ve tutar girin.");
        return;
    }

    const parts = dateStr.split(".");
    const dStr = parts.length === 3 ? `${parts[2]}-${parts[1]}-${parts[0]}` : dateStr;

    if (eid) {
        incomes = incomes.map(i => i.id === eid ? { ...i, title, amount: amt, date: dStr, bank, category: cat, note } : i);
    } else {
        incomes.push({ id: "inc_" + Date.now(), title, amount: amt, date: dStr, bank, category: cat, note });
    }
    savePayments();
    renderIncomes();
    document.getElementById("incomeForm").reset();
    document.getElementById("editingIncomeId").value = "";
};

function renderIncomes() {
    const container = document.getElementById("incomeListContent");
    const totalLabel = document.getElementById("incMonthTotal");
    if (!container) return;
    container.innerHTML = "";
    
    const currentMonth = dayjs().month();
    const currentYear = dayjs().year();
    
    const filtered = incomes.filter(i => {
        const d = dayjs(i.date);
        return d.month() === currentMonth && d.year() === currentYear;
    }).sort((a, b) => dayjs(b.date).diff(dayjs(a.date)));

    let total = 0;
    filtered.forEach(inc => {
        total += inc.amount;
        const div = document.createElement("div");
        div.className = "bg-white p-3 rounded-xl border border-slate-200 shadow-sm flex items-center justify-between group";
        div.innerHTML = `
            <div class="flex items-center gap-3">
                <div class="bg-emerald-50 text-emerald-600 p-2 rounded-lg"><i data-lucide="trending-up" class="w-4 h-4"></i></div>
                <div>
                    <div class="flex items-center gap-2 mb-0.5">
                        <span class="text-[9px] font-black bg-slate-100 px-1.5 rounded uppercase">${dayjs(inc.date).format("DD MMM")}</span>
                        <span class="px-1.5 py-0.5 rounded text-[8px] font-bold uppercase bg-emerald-100 text-emerald-700">${inc.category || 'DİĞER'}</span>
                        <span class="px-1.5 py-0.5 rounded text-[8px] font-bold uppercase bg-slate-50 text-slate-600">${inc.bank}</span>
                    </div>
                    <h4 class="text-xs font-bold text-slate-900 uppercase leading-tight">
                        ${inc.title}${inc.note ? ` <span class="text-slate-400 font-bold">- ${inc.note}</span>` : ''}
                    </h4>
                </div>
            </div>
            <div class="flex items-center gap-4">
                <div class="text-right">
                    <div class="text-sm font-black text-slate-900">${formatTL(inc.amount)}</div>
                </div>
                <div class="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                    <button type="button" onclick="window.editIncome('${inc.id}')" class="p-1 hover:bg-slate-100 rounded text-slate-400 hover:text-blue-600"><i data-lucide="edit-2" class="w-3 h-3"></i></button>
                    <button type="button" onclick="window.deleteIncome('${inc.id}')" class="p-1 hover:bg-slate-100 rounded text-slate-400 hover:text-red-500"><i data-lucide="trash-2" class="w-3 h-3"></i></button>
                </div>
            </div>
        `;
        container.appendChild(div);
    });
    if (totalLabel) totalLabel.innerText = "Aylık Toplam: " + formatTL(total);
    lucide.createIcons();
}

window.editIncome = function (id) {
    const inc = incomes.find(i => i.id === id);
    if (!inc) return;

    // Formu doldur
    document.getElementById("editingIncomeId").value = inc.id;
    document.getElementById("incTitle").value = inc.title;
    document.getElementById("incAmount").value = inc.amount;
    document.getElementById("incBank").value = inc.bank;
    document.getElementById("incCategory").value = inc.category;
    document.getElementById("incNote").value = inc.note || "";
    
    const dStr = dayjs(inc.date).format("DD.MM.YYYY");
    document.getElementById("incDate").value = dStr;
    if (window.incDatePicker) window.incDatePicker.setDate(dStr);
    
    // Buton metnini mavi renkle güncelle
    const saveBtn = document.querySelector("#incomeForm button[onclick='window.saveIncome()']");
    if (saveBtn) {
        saveBtn.innerText = "GELİRİ GÜNCELLE";
        saveBtn.classList.add("bg-blue-600");
        saveBtn.classList.remove("bg-emerald-600");
    }

    // Forma odaklanmak için en üste kaydır
    document.getElementById("incomeModal").scrollTo({ top: 0, behavior: 'smooth' });
};

window.deleteIncome = function (id) {
    if (confirm("Bu gelir kaydını silmek istediğinize emin misiniz?")) {
        incomes = incomes.filter(i => i.id !== id);
        savePayments();
        renderIncomes();
    }
};

window.downloadSampleIncomeExcel = function () {
    const data = [
        ["BAŞLIK", "TUTAR", "BANKA", "KATEGORİ", "TARİH", "NOT", "", "GEÇERLİ BANKALAR"],
        ["Örnek Gelir", 5000, "ZİRAAT", "Satış Geliri", dayjs().format("DD.MM.YYYY"), "Örnek Not", "", "AKBANK"],
        ["", "", "", "", "", "", "", "DENİZBANK"],
        ["", "", "", "", "", "", "", "GARANTİ"],
        ["", "", "", "", "", "", "", "HALK BANKASI"],
        ["", "", "", "", "", "", "", "İNG"],
        ["", "", "", "", "", "", "", "İŞ BANKASI"],
        ["", "", "", "", "", "", "", "KUVEYTTÜRK"],
        ["", "", "", "", "", "", "", "QNB"],
        ["", "", "", "", "", "", "", "ŞEKERBANK"],
        ["", "", "", "", "", "", "", "TEB"],
        ["", "", "", "", "", "", "", "VAKIFBANK"],
        ["", "", "", "", "", "", "", "VAKIF KATILIM"],
        ["", "", "", "", "", "", "", "YAPIKREDİ"],
        ["", "", "", "", "", "", "", "ZİRAAT"],
        ["", "", "", "", "", "", "", "ZİRAAT KATILIM BANKASI"],
        ["", "", "", "", "", "", "", "NAKİT KASA"],
        ["", "", "", "", "", "", "", "DİĞER"]
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Gelir Taslak");
    XLSX.writeFile(wb, "gelir_yukleme_taslak.xlsx");
};

window.handleIncomeExcelSelect = function (event) {
    const file = event.target.files[0];
    const nameLabel = document.getElementById("incomeExcelFileName");
    const uploadBtn = document.getElementById("processIncExcelBtn");
    const removeBtn = document.getElementById("clearIncomeExcelBtn");
    if (!file) return;

    if (nameLabel) {
        nameLabel.innerText = file.name + " SEÇİLDİ";
        nameLabel.parentElement.classList.add("border-emerald-500", "bg-emerald-50");
    }
    if (uploadBtn) uploadBtn.classList.remove("hidden");
    if (removeBtn) removeBtn.classList.remove("hidden");
};

// Excel Temizleme Mantığı
document.getElementById("clearIncomeExcelBtn").onclick = function() {
    document.getElementById("incomeExcelInput").value = "";
    document.getElementById("incomeExcelFileName").innerText = "Dosya Seçin veya Sürükleyin";
    document.getElementById("incomeExcelFileName").parentElement.classList.remove("border-emerald-500", "bg-emerald-50");
    this.classList.add("hidden");
    document.getElementById("processIncExcelBtn").classList.add("hidden");
};

document.getElementById("clearPaymentExcelBtn").onclick = function() {
    document.getElementById("excelImportInput").value = "";
    document.getElementById("excelFileNameDisplay").innerText = "Dosya Seçin veya Sürükleyin";
    document.getElementById("excelFileNameDisplay").classList.remove("border-emerald-400", "bg-emerald-50");
    this.classList.add("hidden");
    document.getElementById("uploadExcelBtn").classList.add("hidden");
};

document.getElementById("processIncExcelBtn").onclick = function() {
    const file = document.getElementById("incomeExcelInput").files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        const binary = new Uint8Array(e.target.result);
        const workbook = XLSX.read(binary, { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(firstSheet);
        
        let imported = 0;
        data.forEach(row => {
            const title = row["Aciklama"] || row["Açıklama"] || row["Başlık"] || row["BAŞLIK"];
            const amount = parseFloat(row["Tutar"] || row["Miktar"] || row["TUTAR"]);
            const bank = (row["Banka"] || row["BANKA"] || "DİĞER").toString().toUpperCase();
            const category = row["Kategori"] || row["KATEGORİ"] || "Diğer Ödemeler";
            const dateStr = row["Tarih"] || row["TARİH"];
            const note = row["Not"] || row["NOT"] || "";
            
            let date = parseExcelDate(dateStr);
            if (title && !isNaN(amount) && date) {
                incomes.push({ id: "inc_" + Date.now() + Math.random(), title, amount, bank, category, date, note });
                imported++;
            }
        });
        if (imported > 0) {
            savePayments();
            renderIncomes();
            alert(imported + " adet gelir kalemi başarıyla yüklendi.");
            // Reset UI
            document.getElementById("clearIncomeExcelBtn").click();
        }
    };
    reader.readAsArrayBuffer(file);
};

// Ödeme Excel Listener (onchange için)
document.getElementById("incomeExcelInput").onchange = window.handleIncomeExcelSelect;
document.getElementById("excelImportInput").onchange = function(e) {
    const file = e.target.files[0];
    const nameLabel = document.getElementById("excelFileNameDisplay");
    const uploadBtn = document.getElementById("uploadExcelBtn");
    const removeBtn = document.getElementById("clearPaymentExcelBtn");
    if(!file) return;
    if(nameLabel) {
        nameLabel.innerText = file.name + " SEÇİLDİ";
        nameLabel.classList.add("border-emerald-400", "bg-emerald-50");
    }
    if(uploadBtn) uploadBtn.classList.remove("hidden");
    if(removeBtn) removeBtn.classList.remove("hidden");
};
