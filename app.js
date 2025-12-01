// Test if script is loading
console.log('App.js module started loading...');

// Import Supabase
import { createClient } from 'https://cdn.skypack.dev/@supabase/supabase-js@2'

// --- UTILITY FUNCTIONS ---
// Convert Khmer numbers to English numbers
const convertKhmerToEnglishNumbers = (text) => {
    if (!text) return text;
    
    const khmerToEnglish = {
        '០': '0', '១': '1', '២': '2', '៣': '3', '៤': '4',
        '៥': '5', '៦': '6', '៧': '7', '៨': '8', '៩': '9'
    };
    
    return text.replace(/[០-៩]/g, (match) => khmerToEnglish[match] || match);
};

// Initialize sort listeners early (with polling) so nested backup/restore exposure runs on load
(() => {
    if (document._backupInitPollerStarted) return;
    document._backupInitPollerStarted = true;
    const tryInit = () => {
        try {
            if (typeof setupLoanSortListeners === 'function') {
                setupLoanSortListeners();
            }
        } catch (_) { /* noop */ }
    };
    const startPoll = () => {
        let tries = 0;
        const max = 120; // ~6s
        const intv = setInterval(() => {
            // If backup already exposed, stop
            if (typeof window.backupAllData === 'function') {
                clearInterval(intv);
                return;
            }
            // Attempt initialization and check again
            tryInit();
            if (typeof window.backupAllData === 'function' || ++tries >= max) {
                clearInterval(intv);
            }
        }, 50);
    };
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => { tryInit(); startPoll(); });
    } else {
        tryInit();
        startPoll();
    }
})();

// Shared: ensure backup function exists, initialize if needed, and invoke once with dedupe
function ensureBackupAndInvoke() {
    const tryInvoke = () => {
        if (typeof window.backupAllData === 'function') { window.backupAllData(); return true; }
        if (typeof backupAllData === 'function') { backupAllData(); return true; }
        return false;
    };

    if (tryInvoke()) return;

    // Attempt to initialize exposure lazily
    try { if (typeof setupLoanSortListeners === 'function') { setupLoanSortListeners(); } } catch (_) { /* noop */ }
    try {
        if (typeof window.bindBackupRestoreControls === 'function') { window.bindBackupRestoreControls(); }
        else if (typeof bindBackupRestoreControls === 'function') { bindBackupRestoreControls(); }
    } catch (_) { /* noop */ }

    if (tryInvoke()) return;

    // Poll up to ~2s for availability
    let tries = 0; const max = 40; const intv = setInterval(() => {
        tries++;
        if (tryInvoke() || tries >= max) { clearInterval(intv); }
    }, 50);
}

function requestBackupOnce() {
    if (document._backupCallLock) return;
    document._backupCallLock = true;
    // Slightly over 1s to cover any double-firing sequences
    setTimeout(() => { document._backupCallLock = false; }, 1200);
    ensureBackupAndInvoke();
}

// Helper: format ISO date (YYYY-MM-DD) to DD/MM/YYYY for print titles
const formatDateDDMMYYYY = (iso) => {
    if (!iso || typeof iso !== 'string') return iso || '';
    const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return iso;
    const [, y, mo, d] = m;
    return `${d}/${mo}/${y}`;
};

console.log('Supabase import successful');

// Supabase configuration
const supabaseUrl = 'https://bcbwrymhpjcncgiwjllr.supabase.co'
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJjYndyeW1ocGpjbmNnaXdqbGxyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTU2MDA5MDEsImV4cCI6MjA3MTE3NjkwMX0.ZiU1uF9_5h2N9choQvNihTvKWfqtPlHdvQm2iPaI2jw'
const supabase = createClient(supabaseUrl, supabaseKey)

// Debug: Test Supabase connection
console.log('Supabase client initialized:', supabase);
console.log('Supabase URL:', supabaseUrl);

// Test if script is loading
console.log('App.js module loaded successfully');

// --- GLOBAL STATE & DOM ELEMENTS ---
let books = [];
let loans = [];
let classLoans = [];
let locations = [];
let students = [];
let readingLogs = [];
let settingsData = {};

// --- LOAN SORT HEADER CLICK LISTENERS ---
const setupLoanSortListeners = () => {
    const bind = (id, key, defaultDir = 'asc') => {
        const el = document.getElementById(id);
        if (!el) return;
        el.addEventListener('click', () => {
            if (loanSortKey === key) {
                // toggle direction
                loanSortDir = loanSortDir === 'asc' ? 'desc' : 'asc';
            } else {
                loanSortKey = key;
                loanSortDir = defaultDir;
            }
            renderLoans();
        });

// --- BACKUP & RESTORE ---
async function backupAllData() {
    console.log('[Backup] backupAllData called');
    // Ensure loading overlay exists
    try {
        let overlayEl = document.getElementById('loading-overlay');
        if (!overlayEl) {
            console.warn('[Backup] loading-overlay not found. Creating fallback element.');
            overlayEl = document.createElement('div');
            overlayEl.id = 'loading-overlay';
            overlayEl.className = 'fixed inset-0 bg-gray-900 bg-opacity-75 flex flex-col items-center justify-center z-40 hidden';
            overlayEl.innerHTML = '<i class="fas fa-spinner fa-spin text-white text-5xl"></i><p class="text-white text-xl mt-4">កំពុងដំណើរការ...</p>';
            document.body.appendChild(overlayEl);
        }
    } catch (e) { console.error('[Backup] Failed ensuring overlay:', e); }
    // Notify user and show overlay immediately
    try { if (typeof window.showToast === 'function') window.showToast('កំពុងបម្រុងទុកទិន្នន័យទាំងអស់...', 'bg-blue-600'); } catch (_) {}
    try {
        const overlayEl = document.getElementById('loading-overlay') || (typeof loadingOverlay !== 'undefined' ? loadingOverlay : null);
        if (overlayEl) overlayEl.classList.remove('hidden');
    } catch (_) {}

// Ensure loan-sort listeners (and nested backup/restore exposure) initialize on DOM ready
// This guarantees window.backupAllData is assigned soon after load
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        try { typeof setupLoanSortListeners === 'function' && setupLoanSortListeners(); } catch (_) { /* noop */ }
    });
} else {
    try { typeof setupLoanSortListeners === 'function' && setupLoanSortListeners(); } catch (_) { /* noop */ }
}

    if (!currentUserId) {
        try {
            const overlayEl = document.getElementById('loading-overlay') || (typeof loadingOverlay !== 'undefined' ? loadingOverlay : null);
            if (overlayEl) overlayEl.classList.add('hidden');
        } catch (_) {}
        console.warn('[Backup] No currentUserId');
        alert('មិនមានអ្នកប្រើប្រាស់! សូមចូលគណនីមុន។');
        return;
    }

    try {
        const fetchTable = async (table) => {
            // Settings are small; one-shot fetch is fine
            if (table === 'settings') {
                const { data, error } = await supabase
                    .from('settings')
                    .select('*')
                    .eq('user_id', currentUserId);
                if (error) throw error;
                return data || [];
            }

            // Paginate for large tables to avoid 1000-row cap
            const pageSize = 1000;
            const all = [];
            let from = 0;
            while (true) {
                const to = from + pageSize - 1;
                const { data, error } = await supabase
                    .from(table)
                    .select('*')
                    .eq('user_id', currentUserId)
                    .order('id', { ascending: true })
                    .range(from, to);
                if (error) throw error;
                const batch = data || [];
                all.push(...batch);
                if (batch.length < pageSize) break; // last page
                from += pageSize;
            }
            return all;
        };

        console.log('[Backup] Fetching tables for user', currentUserId);
        const [booksArr, loansArr, classLoansArr, readingLogsArr, locationsArr, studentsArr, settingsArr] = await Promise.all([
            fetchTable('books'),
            fetchTable('loans'),
            fetchTable('class_loans'),
            fetchTable('reading_logs'),
            fetchTable('locations'),
            fetchTable('students'),
            fetchTable('settings')
        ]);

        if (typeof XLSX === 'undefined') throw new Error('XLSX library not loaded');
        const wb = XLSX.utils.book_new();
        const addSheet = (name, arr) => {
            const ws = XLSX.utils.json_to_sheet(arr);
            XLSX.utils.book_append_sheet(wb, ws, name);
        };
        addSheet('Books', booksArr);
        addSheet('Loans', loansArr);
        addSheet('ClassLoans', classLoansArr);
        addSheet('ReadingLogs', readingLogsArr);
        addSheet('Locations', locationsArr);
        addSheet('Students', studentsArr);
        addSheet('Settings', settingsArr);

        const now = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        // ddmmyyyyhhmmss
        const stamp = `${pad(now.getDate())}${pad(now.getMonth()+1)}${now.getFullYear()}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
        const fname = `LibraryBackup_${stamp}.LBBCK`;
        console.log('[Backup] Writing workbook file', fname);
        XLSX.writeFile(wb, fname, { bookType: 'xlsx' });
        try { if (typeof window.showToast === 'function') window.showToast('បានបម្រុងទុកទិន្នន័យទាំងអស់!', 'bg-green-600'); } catch (_) {}
    } catch (e) {
        console.error('Backup failed:', e);
        alert('ការបម្រុងទុកបានបរាជ័យ: ' + (e?.message || 'Unknown'));
    } finally {
        try {
            const overlayEl = document.getElementById('loading-overlay') || (typeof loadingOverlay !== 'undefined' ? loadingOverlay : null);
            if (overlayEl) overlayEl.classList.add('hidden');
        } catch (_) {}
    }
}
// Expose backup function globally for inline/console/manual triggers
try { window.backupAllData = backupAllData; } catch (_) {}

async function restoreAllDataFromWorkbook(wb) {
    if (!currentUserId) { alert('មិនមានអ្នកប្រើប្រាស់! សូមចូលគណនីមុន។'); return; }
    try { if (typeof window.showToast === 'function') window.showToast('កំពុងស្ដារទិន្នន័យពី .LBBCK...', 'bg-blue-600'); } catch (_) {}
    try { if (typeof loadingOverlay !== 'undefined') loadingOverlay.classList.remove('hidden'); } catch (_) {}

    try {
        const toJson = (sheetName) => {
            const ws = wb.Sheets[sheetName];
            return ws ? XLSX.utils.sheet_to_json(ws) : [];
        };
        const Books = toJson('Books');
        const Loans = toJson('Loans');
        const ClassLoans = toJson('ClassLoans');
        const ReadingLogs = toJson('ReadingLogs');
        const Locations = toJson('Locations');
        const Students = toJson('Students');
        const Settings = toJson('Settings');

        // Ensure user_id on all records
        const setUid = (arr) => arr.map(r => ({ ...r, user_id: currentUserId }));
        const bArr = setUid(Books);
        const lArr = setUid(Loans);
        const clArr = setUid(ClassLoans);
        const rlArr = setUid(ReadingLogs);
        const locArr = setUid(Locations);
        const sArr = setUid(Students);
        const stArr = setUid(Settings);

        // Non-destructive restore: chunked upserts to preserve existing/newer data
        const upsertChunked = async (table, rows, onConflict) => {
            if (!rows || rows.length === 0) return;
            const chunkSize = 500;
            for (let i = 0; i < rows.length; i += chunkSize) {
                const chunk = rows.slice(i, i + chunkSize);
                let error;
                if (onConflict) {
                    ({ error } = await supabase
                        .from(table)
                        .upsert(chunk, { onConflict }));
                } else {
                    // Rely on table primary key (usually id)
                    ({ error } = await supabase
                        .from(table)
                        .upsert(chunk));
                }
                if (error) throw error;
                // Yield to the UI thread to avoid long blocking loops
                await new Promise((r) => setTimeout(r, 0));
            }
        };

        // Upsert in dependency-friendly order
        await upsertChunked('locations', locArr);
        await upsertChunked('books', bArr);
        await upsertChunked('students', sArr);
        await upsertChunked('class_loans', clArr);
        await upsertChunked('loans', lArr);
        await upsertChunked('reading_logs', rlArr);
        await upsertChunked('settings', stArr, 'key,user_id');

        // Reload local state and UI
        await Promise.all([
            loadBooks(currentUserId),
            loadLoans(currentUserId),
            loadClassLoans(currentUserId),
            loadReadingLogs(currentUserId),
            loadLocations(currentUserId),
            loadStudents(currentUserId),
            loadSettings(currentUserId)
        ]);
        renderAll();
        try { if (typeof window.showToast === 'function') window.showToast('ស្ដារទិន្នន័យដោយជោគជ័យ!', 'bg-green-600'); } catch (_) {}
    } catch (e) {
        console.error('Restore failed:', e);
        alert('ការស្ដារឡើងវិញបានបរាជ័យ: ' + (e?.message || 'Unknown'));
    } finally {
        try { if (typeof loadingOverlay !== 'undefined') loadingOverlay.classList.add('hidden'); } catch (_) {}
    }
}

function bindBackupRestoreControls() {
    const fileInput = document.getElementById('file-restore-lbbck');

    // One-time delegated click listener to avoid timing issues and duplicates
    if (!document._backupRestoreDelegated) {
        document._backupRestoreDelegated = true;
        console.log('[Backup] Binding backup/restore controls');
        document.addEventListener('click', (evt) => {
            const t = evt.target;
            if (!t) return;
            const backupBtn = t.id === 'btn-backup-all' ? t : (t.closest && t.closest('#btn-backup-all'));
            if (backupBtn) {
                evt.preventDefault();
                console.log('[Backup] btn-backup-all click (delegated)');
                try { requestBackupOnce(); } catch (err) { console.error('[Backup] Delegated click failed:', err); }
                return;
            }
            const restoreBtn = t.id === 'btn-restore-all' ? t : (t.closest && t.closest('#btn-restore-all'));
            if (restoreBtn) {
                evt.preventDefault();
                console.log('[Backup] btn-restore-all click (delegated)');
                const fi = document.getElementById('file-restore-lbbck');
                if (fi) fi.click();
            }
        }, true);
    }

    // Guarded change handler for the restore file input (bind when present)
    if (fileInput && fileInput.getAttribute('data-bound') !== 'true') {
        fileInput.setAttribute('data-bound', 'true');
        fileInput.addEventListener('change', async () => {
            const file = fileInput.files && fileInput.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    if (typeof XLSX === 'undefined') throw new Error('XLSX library not loaded');
                    const data = new Uint8Array(e.target?.result);
                    const wb = XLSX.read(data, { type: 'array' });
                    await restoreAllDataFromWorkbook(wb);
                } catch (err) {
                    console.error('Restore read failed:', err);
                    alert('មិនអាចអានឯកសារ .LBBCK បានទេ: ' + (err?.message || 'Unknown'));
                } finally {
                    fileInput.value = '';
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    // Observe DOM for the file input appearing later, bind once
    if (!document._backupRestoreObserver) {
        try {
            document._backupRestoreObserver = new MutationObserver(() => {
                const fi = document.getElementById('file-restore-lbbck');
                if (fi && fi.getAttribute('data-bound') !== 'true') {
                    fi.setAttribute('data-bound', 'true');
                    fi.addEventListener('change', async () => {
                        const file = fi.files && fi.files[0];
                        if (!file) return;
                        const reader = new FileReader();
                        reader.onload = async (e) => {
                            try {
                                if (typeof XLSX === 'undefined') throw new Error('XLSX library not loaded');
                                const data = new Uint8Array(e.target?.result);
                                const wb = XLSX.read(data, { type: 'array' });
                                await restoreAllDataFromWorkbook(wb);
                            } catch (err) {
                                console.error('Restore read failed:', err);
                                alert('មិនអាចអានឯកសារ .LBBCK បានទេ: ' + (err?.message || 'Unknown'));
                            } finally {
                                fi.value = '';
                            }
                        };
                        reader.readAsArrayBuffer(file);
                    });
                }
            });
            document._backupRestoreObserver.observe(document.documentElement, { childList: true, subtree: true });
        } catch (_) { /* noop */ }
    }
}
// Expose binder globally so post-login rebind can call it
try { window.bindBackupRestoreControls = bindBackupRestoreControls; } catch (_) {}

// Bind Backup & Restore controls on DOM ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bindBackupRestoreControls);
} else {
    bindBackupRestoreControls();
}

// Observe DOM for late-loaded Settings section and bind once buttons appear
(function setupBackupRestoreObserver() {
    try {
        const observer = new MutationObserver(() => {
            const btn = document.getElementById('btn-backup-all');
            if (btn && btn.getAttribute('data-bound') !== 'true') {
                bindBackupRestoreControls();
            }
        });
        observer.observe(document.body, { childList: true, subtree: true });
    } catch (_) { /* no-op */ }
})();

// Attempt to focus the Books search box when the Books page is visible
document.addEventListener('DOMContentLoaded', () => {
    const tryFocusSearchBooks = () => {
        const input = document.getElementById('search-books');
        const page = document.getElementById('page-books');
        if (!input || !page) return false;
        const pageVisible = !page.classList.contains('hidden');
        const inputVisible = !!(input.offsetParent);
        if (pageVisible && inputVisible) {
            input.focus();
            return true;
        }
        return false;
    };

    let attempts = 0;
    const maxAttempts = 40; // ~10 seconds total
    const intervalId = setInterval(() => {
        if (tryFocusSearchBooks() || ++attempts >= maxAttempts) {
            clearInterval(intervalId);
        }
    }, 250);
});

// Grouped Return Modal for Individual Loans
window.openGroupedReturnModal = (groupKey) => {
    const [borrower, loanDate] = groupKey.split('|');
    // Find all individual loans in this group
    const groupedLoans = loans.filter(l => 
        (l.borrower || '').trim() === borrower && 
        l.loan_date === loanDate && 
        !l.class_loan_id
    );
    if (groupedLoans.length === 0) return;

    // Populate header
    document.getElementById('grouped-return-borrower-name').textContent = borrower;
    document.getElementById('grouped-return-loan-date-ind').textContent = loanDate;

    // Build book list with available quantities to return
    const listEl = document.getElementById('grouped-return-books-list-ind');
    listEl.innerHTML = '';
    const map = new Map(); // book_id -> { title, available }
    groupedLoans.forEach(l => {
        if (l.status === 'ខ្ចី') {
            const key = String(l.book_id);
            if (!map.has(key)) map.set(key, { title: (books.find(b => String(b.id) === key)?.title) || 'សៀវភៅត្រូវបានលុប', available: 0 });
            map.get(key).available += 1;
        }
    });
    Array.from(map.entries()).forEach(([bookId, info]) => {
        const wrapper = document.createElement('div');
        wrapper.className = 'flex items-center justify-between bg-white p-2 rounded border';
        wrapper.innerHTML = `
            <label class="flex items-center space-x-2">
                <input type="checkbox" data-book-id="${bookId}" class="mr-2">
                <span>${info.title}</span>
            </label>
            <input type="number" class="w-20 border rounded px-2 py-1" data-qty-for="${bookId}" min="1" max="${info.available}" value="${info.available}">
        `;
        listEl.appendChild(wrapper);
    });

    document.getElementById('grouped-return-modal').classList.remove('hidden');
};

window.closeGroupedReturnModal = () => {
    document.getElementById('grouped-return-modal').classList.add('hidden');
};

document.getElementById('grouped-return-form-ind').addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentUserId) return;
    const borrower = document.getElementById('grouped-return-borrower-name').textContent;
    const loanDate = document.getElementById('grouped-return-loan-date-ind').textContent;
    const checkboxes = document.querySelectorAll('#grouped-return-books-list-ind input[type="checkbox"]:checked');
    if (checkboxes.length === 0) { alert('សូមជ្រើសរើសសៀវភៅយ៉ាងហោចណាស់មួយក្បាលដើម្បីសង។'); return; }
    const loadingOverlay = document.getElementById('loading-overlay');
    if (loadingOverlay) loadingOverlay.classList.remove('hidden');
    try {
        const today = new Date().toISOString().split('T')[0];
        for (const cb of checkboxes) {
            const bookId = cb.getAttribute('data-book-id');
            const qtyInput = document.querySelector(`#grouped-return-books-list-ind input[data-qty-for="${bookId}"]`);
            const qty = Math.max(1, Math.min(parseInt(qtyInput.value, 10) || 1, parseInt(qtyInput.max, 10) || 1));
            // Find active individual loans for this group and book
            const candidates = loans.filter(l => 
                !l.class_loan_id && l.status === 'ខ្ចី' && 
                (l.borrower || '').trim() === borrower && l.loan_date === loanDate && String(l.book_id) === String(bookId)
            ).slice(0, qty);
            for (const loan of candidates) {
                const { error } = await supabase
                    .from('loans')
                    .update({ status: 'សង', return_date: today })
                    .eq('id', loan.id)
                    .eq('user_id', currentUserId);
                if (error) throw error;
            }
        }
        await loadLoans(currentUserId);
        renderAll();
        try { if (typeof window.showToast === 'function') { window.showToast('សងសៀវភៅបានជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
        window.closeGroupedReturnModal();
    } catch (err) {
        console.error('Error processing grouped return (individual): ', err);
        alert('មានបញ្ហាក្នុងការរក្សាទុកការសង។');
    } finally {
        if (loadingOverlay) loadingOverlay.classList.add('hidden');
    }
});

// Ensure individual grouped return handlers are bound globally
(function ensureIndGroupedHandlers() {
    if (window.__indGroupedHandlersBound) return;

    window.openGroupedReturnModal = (groupKey) => {
        const [borrower, loanDate] = groupKey.split('|');
        const groupedLoans = loans.filter(l => (l.borrower || '').trim() === borrower && l.loan_date === loanDate && !l.class_loan_id);
        if (groupedLoans.length === 0) return;

        document.getElementById('grouped-return-borrower-name').textContent = borrower;
        document.getElementById('grouped-return-loan-date-ind').textContent = loanDate;

        const listEl = document.getElementById('grouped-return-books-list-ind');
        listEl.innerHTML = '';
        const map = new Map();
        groupedLoans.forEach(l => {
            if (l.status === 'ខ្ចី') {
                const key = String(l.book_id);
                if (!map.has(key)) map.set(key, { title: (books.find(b => String(b.id) === key)?.title) || 'សៀវភៅត្រូវបានលុប', available: 0 });
                map.get(key).available += 1;
            }
        });
        Array.from(map.entries()).forEach(([bookId, info]) => {
            const wrapper = document.createElement('div');
            wrapper.className = 'flex items-center justify-between bg-white p-2 rounded border';
            wrapper.innerHTML = `
                <label class="flex items-center space-x-2">
                    <input type="checkbox" data-book-id="${bookId}" class="mr-2">
                    <span>${info.title}</span>
                </label>
                <input type="number" class="w-20 border rounded px-2 py-1" data-qty-for="${bookId}" min="1" max="${info.available}" value="${info.available}">
            `;
            listEl.appendChild(wrapper);

            // Auto-fill to max when checked and keep editable; also clamp input
            const cb = wrapper.querySelector(`input[type="checkbox"][data-book-id="${bookId}"]`);
            const qtyInput = wrapper.querySelector(`input[data-qty-for="${bookId}"]`);
            cb.addEventListener('change', () => {
                if (cb.checked) {
                    qtyInput.value = String(info.available);
                    qtyInput.focus();
                }
            });
            qtyInput.addEventListener('input', () => {
                const max = parseInt(qtyInput.max, 10) || info.available || 1;
                const v = parseInt(qtyInput.value, 10) || 1;
                qtyInput.value = String(Math.max(1, Math.min(v, max)));
            });
        });

        document.getElementById('grouped-return-modal').classList.remove('hidden');
    };

    window.closeGroupedReturnModal = () => {
        document.getElementById('grouped-return-modal').classList.add('hidden');
    };

    const indForm = document.getElementById('grouped-return-form-ind');
    if (indForm) {
        indForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!currentUserId) return;
            const borrower = document.getElementById('grouped-return-borrower-name').textContent;
            const loanDate = document.getElementById('grouped-return-loan-date-ind').textContent;
            const checkboxes = document.querySelectorAll('#grouped-return-books-list-ind input[type="checkbox"]:checked');
            if (checkboxes.length === 0) { alert('សូមជ្រើសរើសសៀវភៅយ៉ាងហោចណាស់មួយក្បាលដើម្បីសង។'); return; }

            const loadingOverlay = document.getElementById('loading-overlay');
            if (loadingOverlay) loadingOverlay.classList.remove('hidden');
            try {
                const today = new Date().toISOString().split('T')[0];
                for (const cb of checkboxes) {
                    const bookId = cb.getAttribute('data-book-id');
                    const qtyInput = document.querySelector(`#grouped-return-books-list-ind input[data-qty-for="${bookId}"]`);
                    const qty = Math.max(1, Math.min(parseInt(qtyInput.value, 10) || 1, parseInt(qtyInput.max, 10) || 1));
                    const candidates = loans.filter(l => !l.class_loan_id && l.status === 'ខ្ចី' && (l.borrower || '').trim() === borrower && l.loan_date === loanDate && String(l.book_id) === String(bookId)).slice(0, qty);
                    for (const loan of candidates) {
                        const { error } = await supabase
                            .from('loans')
                            .update({ status: 'សង', return_date: today })
                            .eq('id', loan.id)
                            .eq('user_id', currentUserId);
                        if (error) throw error;
                    }
                }
                await loadLoans(currentUserId);
                renderAll();
                try { if (typeof window.showToast === 'function') { window.showToast('សងសៀវភៅបានជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
                window.closeGroupedReturnModal();
            } catch (err) {
                console.error('Error processing grouped return (individual): ', err);
                alert('មានបញ្ហាក្នុងការរក្សាទុកការសង។');
            } finally {
                if (loadingOverlay) loadingOverlay.classList.add('hidden');
            }
        });
    }

    // Grouped loan delete modal handlers (Individual Loans)
    window.openGroupedLoanDeleteModal = (groupKey) => {
        const modal = document.getElementById('grouped-loan-delete-modal');
        const msg = document.getElementById('grouped-loan-delete-message');
        const btn = document.getElementById('grouped-loan-delete-confirm-btn');
        if (!modal || !btn) return;
        const [borrower, loanDate] = groupKey.split('|');
        const count = loans.filter(l => !l.class_loan_id && (l.borrower || '') === borrower && l.loan_date === loanDate).length;
        if (msg) {
            const b = (borrower || '').trim();
            msg.textContent = `តើអ្នកពិតជាចង់លុបការខ្ចីរបស់ ${b} កាលបរិច្ឆេទ ${loanDate} ចំនួន ${count} កំណត់ត្រា មែនទេ?`;
        }
        btn.onclick = async () => { await window.performGroupedLoanDelete(groupKey); window.closeGroupedLoanDeleteModal(); };
        modal.classList.remove('hidden');
    };

    window.closeGroupedLoanDeleteModal = () => {
        const modal = document.getElementById('grouped-loan-delete-modal');
        const btn = document.getElementById('grouped-loan-delete-confirm-btn');
        if (btn) btn.onclick = null;
        if (modal) modal.classList.add('hidden');
    };

    window.performGroupedLoanDelete = async (groupKey) => {
        if (!currentUserId) return;
        const [borrower, loanDate] = groupKey.split('|');
        const toDeleteCount = loans.filter(l => !l.class_loan_id && (l.borrower || '') === borrower && l.loan_date === loanDate).length;
        try {
            const { error } = await supabase
                .from('loans')
                .delete()
                .eq('user_id', currentUserId)
                .eq('borrower', borrower)
                .eq('loan_date', loanDate)
                .is('class_loan_id', null);
            if (error) throw error;
            await loadLoans(currentUserId);
            renderAll();
            try {
                if (typeof window.showToast === 'function') {
                    const b = (borrower || '').trim();
                    window.showToast(`បានលុបការខ្ចីរបស់ ${b} កាលបរិច្ឆេទ ${loanDate} ចំនួន ${toDeleteCount} កំណត់ត្រា ដោយជោគជ័យ!`, 'bg-green-600');
                }
            } catch (_) { /* noop */ }
        } catch (e) {
            console.error('Error deleting grouped loans: ', e);
            alert('ការលុបបានបរាជ័យ។');
        }
    };

    window.deleteGroupedLoan = (groupKey) => {
        if (!currentUserId) return;
        window.openGroupedLoanDeleteModal(groupKey);
    };

    window.__indGroupedHandlersBound = true;
})();

// Delete all individual loans in a group
window.deleteGroupedLoan = (groupKey) => {
    if (!currentUserId) return;
    window.openGroupedLoanDeleteModal(groupKey);
};
    };
    bind('loan-sort-serial', 'serial', 'asc');
    bind('loan-sort-borrower', 'borrower', 'asc');
    bind('loan-sort-title', 'title', 'asc');
    bind('loan-sort-loan-date', 'loan_date', 'desc');
    bind('loan-sort-return-date', 'return_date', 'desc');
    bind('loan-sort-status', 'status', 'asc');
};
let currentUserId = null;
let unsubscribeBooks = () => {};
let unsubscribeLoans = () => {};
let unsubscribeClassLoans = () => {};
let unsubscribeLocations = () => {};
let unsubscribeStudents = () => {};
let unsubscribeSettings = () => {};
let unsubscribeReadingLogs = () => {};
let currentScannedBooks = [];
let currentClassLoanScannedBooks = []; // For class loan multiple books
let currentStudentGender = ''; // For gender tracking
let selectedLoanBooks = []; // For multiple book selection in individual loans
// --- LOAN LIST SORT STATE ---
let loanSortKey = 'loan_date'; // one of: 'serial','borrower','title','loan_date','return_date','status'
let loanSortDir = 'desc'; // 'asc' | 'desc'
let loanSortSetupDone = false; // prevent duplicate listeners

// --- READING LOG SORT STATE ---
let readingLogSortKey = 'date'; // 'date' | 'title'
let readingLogSortDir = 'desc'; // 'asc' | 'desc'
let readingLogSortSetupDone = false; // prevent duplicate listeners

// Helper: extract a displayable books string from a reading log
const getBooksReadString = (log) => {
    let booksRead = 'N/A';
    if (log && log.books) {
        try {
            const booksData = typeof log.books === 'string' ? JSON.parse(log.books) : log.books;
            if (Array.isArray(booksData)) {
                booksRead = booksData.map(b => b?.title || b).join(', ');
            } else if (booksData && typeof booksData === 'object') {
                booksRead = booksData.title || booksData.toString();
            } else {
                booksRead = booksData?.toString?.() || 'N/A';
            }
        } catch (_) {
            booksRead = log.books.toString();
        }
    }
    return booksRead;
};

// --- READING LOG SORT HEADER CLICK LISTENER ---
const setupReadingSortListeners = () => {
    const titleHeader = document.getElementById('reading-sort-title');
    if (!titleHeader) return;
    if (titleHeader.getAttribute('data-bound') === 'true') return;
    titleHeader.setAttribute('data-bound', 'true');

    titleHeader.addEventListener('click', () => {
        if (readingLogSortKey === 'title') {
            readingLogSortDir = readingLogSortDir === 'asc' ? 'desc' : 'asc';
        } else {
            readingLogSortKey = 'title';
            readingLogSortDir = 'asc';
        }
        renderReadingLogs();
    });
};

const authContainer = document.getElementById('auth-container');
const appContainer = document.getElementById('app-container');
const loadingOverlay = document.getElementById('loading-overlay');
const loginForm = document.getElementById('login-form');
const showLoginBtn = document.getElementById('show-login');
const logoutBtn = document.getElementById('logout-btn');
const sidebarSchoolName = document.getElementById('sidebar-school-name');

const pages = document.querySelectorAll('.page');
const navLinks = document.querySelectorAll('.nav-link');

// --- CLASS LOAN: scanned books list rendering/removal ---
const renderClassLoanScannedBooks = () => {
    const list = document.getElementById('class-loan-scanned-books-list');
    if (!list) return;
    // Clear existing
    list.innerHTML = '';
    // Render items
    currentClassLoanScannedBooks.forEach((book, index) => {
        const li = document.createElement('li');
        li.className = 'selected-book-item flex items-center justify-between p-2 bg-gray-100 rounded mb-1';
        li.innerHTML = `
            <div class="flex-1">
                <span class="font-medium">${book.title}</span>
                <span class="text-sm text-blue-600 ml-2">(នៅសល់: ${book.remaining})</span>
            </div>
            <button type="button" onclick="removeClassLoanScannedBook(${index})" class="text-red-500 hover:text-red-700 ml-2">
                <i class="fas fa-times"></i>
            </button>
        `;
        list.appendChild(li);
    });
};

window.removeClassLoanScannedBook = (index) => {
    if (index >= 0 && index < currentClassLoanScannedBooks.length) {
        currentClassLoanScannedBooks.splice(index, 1);
        renderClassLoanScannedBooks();
        setTimeout(() => document.getElementById('class-loan-isbn-input')?.focus(), 50);
    }
};

// --- AUTHENTICATION ---
// Check initial auth state
supabase.auth.getSession().then(({ data: { session } }) => {
    handleAuthState(session);
});

// Listen for auth changes
supabase.auth.onAuthStateChange((event, session) => {
    handleAuthState(session);
});

function handleAuthState(session) {
    if (session?.user) {
        currentUserId = session.user.id;
        document.getElementById('user-email').textContent = session.user.email;
        authContainer.classList.add('hidden');
        appContainer.classList.remove('hidden');
        setupRealtimeListeners(currentUserId);
        try {
            // Attempt an immediate bind if available; otherwise, rely on the polling safety net below
            if (typeof bindBackupRestoreControls === 'function') {
                bindBackupRestoreControls();
            } else if (typeof window.bindBackupRestoreControls === 'function') {
                window.bindBackupRestoreControls();
            }
            // Safety net: poll briefly until buttons exist, then bind once
            try {
                let attempts = 0;
                const maxAttempts = 20; // ~2s total
                const poll = setInterval(() => {
                    attempts++;
                    const btn = document.getElementById('btn-backup-all');
                    if (btn && btn.getAttribute('data-bound') === 'true') {
                        clearInterval(poll);
                        return;
                    }
                    if (btn) {
                        // Bind directly as an extra guarantee
                        if (btn.getAttribute('data-bound') !== 'true') {
                            btn.setAttribute('data-bound', 'true');
                            btn.addEventListener('click', (e) => {
                                e.preventDefault();
                                console.log('[Backup] btn-backup-all click (direct-poll)');
                                try { requestBackupOnce(); } catch (err) { console.error('[Backup] Direct poll click failed:', err); }
                            });
                        }
                        clearInterval(poll);
                        return;
                    }
                    if (typeof bindBackupRestoreControls === 'function') {
                        bindBackupRestoreControls();
                    }
                    if (attempts >= maxAttempts) clearInterval(poll);
                }, 100);
            } catch (_) { /* noop */ }
        } catch (e) { console.warn('[Backup] Rebind failed:', e); }
        
        // Determine target page: URL param/hash > last visited > settings default > 'home'
        let target = 'home';
        try {
            const url = new URL(window.location.href);
            const fromQuery = (url.searchParams.get('page') || '').trim();
            const fromHash = (url.hash || '').replace(/^#/, '').trim();
            const preferred = fromQuery || fromHash;
            if (preferred) {
                target = preferred;
                // Persist so subsequent navigations respect this choice
                localStorage.setItem('lastPage', target);
                // Optional: clean the URL (keep other params)
                try {
                    url.searchParams.delete('page');
                    history.replaceState(null, '', url.toString());
                } catch (_) { /* noop */ }
            } else {
                const saved = localStorage.getItem('lastPage');
                if (saved) target = saved;
                else if (settingsData && settingsData.default_page) target = settingsData.default_page;
            }
        } catch (e) { /* ignore */ }
        navigateTo(target);
    } else {
        currentUserId = null;
        appContainer.classList.add('hidden');
        authContainer.classList.remove('hidden');
        // Unsubscribe from all listeners
        unsubscribeBooks();
        unsubscribeLoans();
        unsubscribeClassLoans();
        unsubscribeLocations();
        unsubscribeStudents();
        unsubscribeSettings();
        unsubscribeReadingLogs();
        // Clear local data
        books = []; loans = []; classLoans = []; locations = []; students = []; readingLogs = [];
        renderAll();
    }
}

loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const email = document.getElementById('login-email').value;
    const password = document.getElementById('login-password').value;
    const errorP = document.getElementById('login-error');
    errorP.textContent = '';
    
    console.log('Attempting login with:', { email, password: '***' });
    
    try {
        const { data, error } = await supabase.auth.signInWithPassword({ email, password });
        console.log('Login response:', { data, error });
        
        if (error) {
            console.error('Login error:', error);
            errorP.textContent = 'ការចូលប្រើបានបរាជ័យ: ' + error.message;
        } else {
            console.log('Login successful:', data);
        }
    } catch (err) {
        console.error('Login exception:', err);
        errorP.textContent = 'ការចូលប្រើបានបរាជ័យ: ' + err.message;
    }
});

logoutBtn.addEventListener('click', async () => {
    try {
        // Clear last visited page on logout
        localStorage.removeItem('lastPage');
    } catch (e) { /* ignore */ }
    await supabase.auth.signOut();
});

// --- NAVIGATION ---
const navigateTo = (pageId) => {
    pages.forEach(page => page.classList.add('hidden'));
    const targetPage = document.getElementById(`page-${pageId}`);
    if (targetPage) {
        targetPage.classList.remove('hidden');
    }

    navLinks.forEach(nav => {
        if (nav.getAttribute('data-page') === pageId) {
            nav.classList.add('bg-gray-900', 'text-white');
            nav.classList.remove('text-gray-300');
        } else {
            nav.classList.remove('bg-gray-900', 'text-white');
            nav.classList.add('text-gray-300');
        }
    });

    // Persist last visited page
    try { localStorage.setItem('lastPage', pageId); } catch (e) { /* ignore */ }

    // Auto-focus or setup for specific pages
    if (pageId === 'books') {
        // Focus the search box shortly after the page becomes visible
        setTimeout(() => {
            const input = document.getElementById('search-books');
            if (input) { input.focus(); input.select(); }
        }, 100);
    }
    if (pageId === 'locations') {
        // Focus search on Locations page
        setTimeout(() => {
            const input = document.getElementById('search-locations');
            if (input) { input.focus(); input.select(); }
        }, 100);
    }
    if (pageId === 'class-loans') {
        populateClassLoanForm();
        setTimeout(() => document.getElementById('class-loan-isbn-input').focus(), 100);
    }
    if (pageId === 'loans') {
        window.clearLoanForm();
        if (!loanSortSetupDone) { setupLoanSortListeners(); loanSortSetupDone = true; }
    }
    if (pageId === 'reading-log') {
        window.clearReadingLogForm();
        setTimeout(() => document.getElementById('reading-log-student-id').focus(), 100);
        // Ensure reading log sort header is bound
        try { setupReadingSortListeners(); } catch (_) { /* noop */ }
    }
    // Ensure Backup & Restore buttons are bound when entering Settings
    if (pageId === 'settings') {
        try { bindBackupRestoreControls(); } catch (_) { /* noop */ }
    }
    if (pageId === 'students') {
        // Focus search on Students page
        setTimeout(() => {
            const input = document.getElementById('search-students');
            if (input) { input.focus(); input.select(); }
        }, 100);
    }
};

navLinks.forEach(link => {
    link.addEventListener('click', (e) => {
        const pageId = link.getAttribute('data-page');
        // Only prevent default and navigate if it's an internal page
        if (pageId) {
            e.preventDefault();
            navigateTo(pageId);
        }
        // If no data-page attribute, let the natural href work (for external links like stcard.html)
    });
});

// Refocus relevant inputs after printing finishes
window.addEventListener('afterprint', () => {
    const activePage = document.querySelector('.page:not(.hidden)');
    if (!activePage) return;
    if (activePage.id === 'page-locations') {
        const input = document.getElementById('search-locations');
        if (input) { input.focus(); input.select(); }
    }
    if (activePage.id === 'page-students') {
        const input = document.getElementById('search-students');
        if (input) { input.focus(); input.select(); }
    }
    if (activePage.id === 'page-reading-log') {
        const input = document.getElementById('reading-log-student-id');
        if (input) { input.focus(); input.select(); }
    }
    if (activePage.id === 'page-class-loans') {
        const input = document.getElementById('class-loan-isbn-input');
        if (input) { input.focus(); input.select(); }
    }
    if (activePage.id === 'page-loans') {
        const input = document.getElementById('loan-isbn-input');
        if (input) { input.focus(); input.select(); }
    }
    if (activePage.id === 'page-books') {
        const input = document.getElementById('search-books');
        if (input) { input.focus(); input.select(); }
    }
});

// --- BOOKS EXCEL IMPORT: wire button to hidden file input ---
function bindBookImportControls() {
    const btn = document.getElementById('upload-book-excel-btn');
    const fileInput = document.getElementById('book-excel-file-input');
    if (!btn || !fileInput) return false;
    if (btn.getAttribute('data-bound') === 'true') return true;
    btn.setAttribute('data-bound', 'true');
    btn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
        const file = fileInput.files && fileInput.files[0];
        if (!file) return;
        const lowerName = (file.name || '').toLowerCase();
        const isExcel = lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls') || /sheet/.test(file.type);

        try { if (typeof window.showToast === 'function') window.showToast('កំពុងត្រៀមនាំចូល Excel សម្រាប់សៀវភៅ...', 'bg-blue-600'); } catch (_) {}

        const finish = () => { try { loadingOverlay.classList.add('hidden'); } catch (_) {} fileInput.value = ''; };
        try { loadingOverlay.classList.remove('hidden'); } catch (_) {}

        const handleParsedAndSyncBooks = async (parsed) => {
            if (!parsed.length) {
                const m = 'មិនអាចញែកទិន្នន័យពី CSV/Excel បានទេ ឬឯកសារវិកល។';
                if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
                return;
            }
            const normalized = normalizeBooks(parsed);
            if (!normalized.length) {
                const m = 'រកមិនឃើញជួរឈរដែលត្រូវការសម្រាប់សៀវភៅនៅក្នុង CSV/Excel ទេ।';
                if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
                return;
            }
            await syncBooksToSupabase(normalized);
            await loadBooks(currentUserId);
            renderAll();
            if (typeof window.showToast === 'function') {
                window.showToast(`នាំចូលសៀវភៅ ${normalized.length} ក្បាលដោយជោគជ័យពី CSV/Excel!`, 'bg-green-600');
            } else {
                alert(`នាំចូលសៀវភៅ ${normalized.length} ក្បាលដោយជោគជ័យពី CSV/Excel!`);
            }
        };

        if (isExcel) {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    if (typeof XLSX === 'undefined') throw new Error('XLSX library not loaded');
                    const data = new Uint8Array(e.target?.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const csvText = XLSX.utils.sheet_to_csv(worksheet);
                    const parsed = parseCSV(csvText);
                    await handleParsedAndSyncBooks(parsed);
                } catch (err) {
                    console.error('Excel upload (books) failed:', err);
                    const m = 'ការនាំចូល Excel បានបរាជ័យ: ' + (err?.message || 'Unknown error');
                    if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
                } finally { finish(); }
            };
            reader.readAsArrayBuffer(file);
        } else {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const text = e.target?.result || '';
                    const parsed = parseCSV(String(text));
                    await handleParsedAndSyncBooks(parsed);
                } catch (err) {
                    console.error('CSV upload (books) failed:', err);
                    const m = 'ការនាំចូល CSV បានបរាជ័យ: ' + (err?.message || 'Unknown error');
                    if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
                } finally { finish(); }
            };
            reader.readAsText(file, 'utf-8');
        }
    });
    return true;
}

// --- LOCATIONS EXCEL/CSV IMPORT ---
function normalizeLocations(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return [];
    const sample = rows[0] || {};
    const keys = Object.keys(sample);

    const findKey = (aliases) => {
        const lowered = keys.map(k => ({ k, lk: k.toLowerCase().trim() }));
        for (const alias of aliases) {
            const a = alias.toLowerCase().trim();
            const exact = lowered.find(x => x.lk === a);
            if (exact) return exact.k;
        }
        for (const alias of aliases) {
            const a = alias.toLowerCase().trim();
            const incl = lowered.find(x => x.lk.includes(a));
            if (incl) return incl.k;
        }
        return null;
    };

    const keyName = findKey(['ឈ្មោះទីតាំង','ឈ្មោះ','name','location name','បន្ទប់','ធ្នើ']);
    const keySource = findKey(['ប្រភពធ្នើ','ប្រភព','source','origin']);
    const keyYear = findKey(['ឆ្នាំ','year','ឆ្នាំសិក្សា']);

    const norm = rows.map(r => {
        const name = keyName ? (r[keyName] ?? '').toString().trim() : '';
        const source = keySource ? (r[keySource] ?? '').toString().trim() : '';
        const year = keyYear ? (r[keyYear] ?? '').toString().trim() : '';
        return { name, source, year };
    }).filter(o => o.name);

    return norm;
}

async function syncLocationsToSupabase(locationData) {
    if (!currentUserId) return;
    // Delete existing locations for this user
    await supabase
        .from('locations')
        .delete()
        .eq('user_id', currentUserId);

    // Attach user_id
    const withUserId = locationData.map(loc => ({ ...loc, user_id: currentUserId }));
    const { error } = await supabase
        .from('locations')
        .insert(withUserId);
    if (error) {
        console.error('Error syncing locations:', error);
        throw error;
    }
}

const handleParsedAndSyncLocations = async (parsed) => {
    if (!parsed.length) {
        const m = 'មិនអាចញែកទិន្នន័យពី CSV/Excel បានទេ ឬឯកសារវិកល។';
        if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
        return;
    }
    const normalized = normalizeLocations(parsed);
    if (!normalized.length) {
        const m = 'រកមិនឃើញជួរឈរដែលត្រូវការសម្រាប់ទីតាំងនៅក្នុង CSV/Excel ទេ។';
        if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
        return;
    }
    await syncLocationsToSupabase(normalized);
    await loadLocations(currentUserId);
    renderAll();
    if (typeof window.showToast === 'function') {
        window.showToast(`នាំចូលទីតាំង ${normalized.length} ឯកសារដោយជោគជ័យពី CSV/Excel!`, 'bg-green-600');
    } else {
        alert(`នាំចូលទីតាំង ${normalized.length} ឯកសារដោយជោគជ័យពី CSV/Excel!`);
    }
};

function bindLocationImportControls() {
    const btn = document.getElementById('upload-location-excel-btn');
    const fileInput = document.getElementById('location-excel-file-input');
    if (!btn || !fileInput) return false;
    if (btn.getAttribute('data-bound') === 'true') return true;
    btn.setAttribute('data-bound', 'true');
    btn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
        const file = fileInput.files && fileInput.files[0];
        if (!file) return;

        const finish = () => { fileInput.value = ''; };

        if (/\.(xlsx|xls)$/i.test(file.name)) {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    if (typeof XLSX === 'undefined') throw new Error('XLSX library not loaded');
                    const data = new Uint8Array(e.target?.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const csvText = XLSX.utils.sheet_to_csv(worksheet);
                    const parsed = parseCSV(csvText);
                    await handleParsedAndSyncLocations(parsed);
                } catch (err) {
                    console.error('Excel upload (locations) failed:', err);
                    const m = 'ការនាំចូល Excel បានបរាជ័យ: ' + (err?.message || 'Unknown error');
                    if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
                } finally { finish(); }
            };
            reader.readAsArrayBuffer(file);
        } else {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const text = e.target?.result || '';
                    const parsed = parseCSV(String(text));
                    await handleParsedAndSyncLocations(parsed);
                } catch (err) {
                    console.error('CSV upload (locations) failed:', err);
                    const m = 'ការនាំចូល CSV បានបរាជ័យ: ' + (err?.message || 'Unknown error');
                    if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
                } finally { finish(); }
            };
            reader.readAsText(file, 'utf-8');
        }
    });
    return true;
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bindBookImportControls);
} else {
    // DOM already ready
    bindBookImportControls();
}

// Bind Locations import controls on DOM ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bindLocationImportControls);
} else {
    bindLocationImportControls();
}

// --- RENDERING FUNCTIONS ---
const renderAll = () => {
    renderBooks();
    renderLoans();
    renderClassLoans();
    renderLocations();
    renderStudents();
    renderReadingLogs();
    updateHomeReadingStats();
    updateDashboard();
    renderHomeLoans();
    populateBookDropdowns();
};

const populateAuthorFilter = () => {
    const authorFilter = document.getElementById('author-filter');
    if (!authorFilter) return;

    const authors = [...new Set(books.map(book => book.author).filter(Boolean))];
    
    // Clear existing options except the first one
    while (authorFilter.options.length > 1) {
        authorFilter.remove(1);
    }

    authors.sort().forEach(author => {
        const option = document.createElement('option');
        option.value = author;
        option.textContent = author;
        authorFilter.appendChild(option);
    });
};

const renderBooks = () => {
    const bookList = document.getElementById('book-list');
    const searchBooksInput = document.getElementById('search-books');
    const authorFilter = document.getElementById('author-filter');
    
    const searchTerm = searchBooksInput.value.toLowerCase();
    const selectedAuthor = authorFilter.value;

    const filteredBooks = books.filter(book => {
        const matchesSearch = book.title.toLowerCase().includes(searchTerm) || 
                              (book.author && book.author.toLowerCase().includes(searchTerm)) || 
                              (book.isbn && book.isbn.toLowerCase().includes(searchTerm));
        
        const matchesAuthor = !selectedAuthor || (book.author === selectedAuthor);

        return matchesSearch && matchesAuthor;
    });

    bookList.innerHTML = '';
    if (filteredBooks.length === 0) { bookList.innerHTML = `<tr><td colspan="9" class="text-center p-4 text-gray-500">រកមិនឃើញសៀវភៅទេ។</td></tr>`; return; }
    const sortedBooks = [...filteredBooks].sort((a, b) => {
        const aAuthor = (a.author || '').toLowerCase();
        const bAuthor = (b.author || '').toLowerCase();
        const byAuthor = aAuthor.localeCompare(bAuthor);
        if (byAuthor !== 0) return byAuthor;
        const aTitle = (a.title || '').toLowerCase();
        const bTitle = (b.title || '').toLowerCase();
        return aTitle.localeCompare(bTitle);
    });
    sortedBooks.forEach((book, index) => {
        // Badge after title should reflect only individual borrowings (exclude class-linked)
        const individualLoanedCount = loans.filter(loan => loan.book_id === book.id && loan.status === 'ខ្ចី' && !loan.class_loan_id).length;
        // Compute active class-loan count for this book (sum of remaining quantities per class-loan)
        const classActiveCount = classLoans
            .filter(cl => cl.book_id === book.id)
            .reduce((sum, cl) => {
                const rem = (cl.loaned_quantity || 0) - (cl.returned_count || 0);
                return sum + (rem > 0 ? rem : 0);
            }, 0);
        // Remaining stock should reflect all active borrowings (individual + class-linked)
        const activeLoanedCountAll = loans.filter(loan => loan.book_id === book.id && loan.status === 'ខ្ចី').length;
        const remaining = (book.quantity || 0) - activeLoanedCountAll;
        const location = locations.find(loc => loc.id === book.location_id);
        const row = document.createElement('tr');
        row.className = 'border-b';
        row.innerHTML = `
            <td class="p-3">${index + 1}</td>
            <td class="p-3">${book.title}
                ${individualLoanedCount > 0 ? `<span class=\"text-xs bg-yellow-200 text-yellow-800 rounded-full px-2 py-1 ml-2 no-print\">ខ្ចីបុគ្គល ${individualLoanedCount}</span>` : ''}
                ${classActiveCount > 0 ? `<span class=\"text-xs bg-blue-200 text-blue-800 rounded-full px-2 py-1 ml-2 no-print\">ខ្ចីថ្នាក់ ${classActiveCount}</span>` : ''}
            </td>
            <td class="p-3">${book.author || ''}</td>
            <td class="p-3">${book.isbn || ''}</td>
            <td class="p-3">${book.quantity || 0}</td>
            <td class="p-3 font-bold ${remaining > 0 ? 'text-green-600' : 'text-red-600'}">${remaining}</td>
            <td class="p-3">${location ? location.name : 'N/A'}</td>
            <td class="p-3">${book.source || ''}</td>
            <td class="p-3 no-print"><button onclick="window.editBook('${book.id}')" class="text-blue-500 hover:text-blue-700 mr-2"><i class="fas fa-edit"></i></button><button onclick="window.deleteBook('${book.id}')" class="text-red-500 hover:text-red-700"><i class="fas fa-trash"></i></button></td>
        `;
        bookList.appendChild(row);
    });
};

const renderLoans = () => {
    const loanList = document.getElementById('loan-list');
    const searchLoansInput = document.getElementById('search-loans');
    const loanSummary = document.getElementById('loan-summary');
    const startDate = document.getElementById('loan-filter-start-date').value;
    const endDate = document.getElementById('loan-filter-end-date').value;

    const individualLoans = loans.filter(loan => !loan.class_loan_id);
    
    // Date Filtering
    let dateFilteredLoans = individualLoans;
    if (startDate) {
        dateFilteredLoans = dateFilteredLoans.filter(loan => loan.loan_date >= startDate);
    }
    if (endDate) {
        dateFilteredLoans = dateFilteredLoans.filter(loan => loan.loan_date <= endDate);
    }

    // Group by borrower + loan_date to combine multiple books borrowed at once
    const groupKey = (l) => `${(l.borrower || '').trim()}|${l.loan_date}`;
    const groupsMap = new Map();
    dateFilteredLoans.forEach(l => {
        const key = groupKey(l);
        if (!groupsMap.has(key)) groupsMap.set(key, []);
        groupsMap.get(key).push(l);
    });

    // Build grouped array with aggregated info
    let grouped = Array.from(groupsMap.entries()).map(([key, loansArr]) => {
        const [borrower, loan_date] = key.split('|');
        const titles = loansArr.map(x => {
            const b = books.find(bb => bb.id === x.book_id);
            return b ? b.title : 'សៀវភៅត្រូវបានលុប';
        });
        const statuses = new Set(loansArr.map(x => x.status));
        const statusLabel = statuses.size === 1 ? loansArr[0].status : 'ចម្រុះ';
        const anyActive = loansArr.some(x => x.status === 'ខ្ចី');
        const allReturned = loansArr.every(x => x.status !== 'ខ្ចី');
        const returnDates = loansArr.map(x => x.return_date).filter(Boolean).sort((a,b) => new Date(b) - new Date(a));
        const displayReturnDate = allReturned ? (returnDates[0] || 'N/A') : 'N/A';
        const genderVotes = new Map();
        loansArr.forEach(x => {
            const g = (x.borrower_gender || '').trim();
            if (!g) return;
            genderVotes.set(g, (genderVotes.get(g) || 0) + 1);
        });
        // pick most frequent gender label if any
        let borrower_gender = '';
        let max = 0;
        for (const [g, c] of genderVotes.entries()) { if (c > max) { max = c; borrower_gender = g; } }
        return {
            borrower,
            loan_date,
            titles,
            status: statusLabel,
            anyActive,
            return_date: displayReturnDate,
            borrower_gender,
            loans: loansArr
        };
    });

    // Search Term Filtering against borrower and book titles
    const searchTerm = searchLoansInput.value.toLowerCase();
    if (searchTerm) {
        grouped = grouped.filter(g => {
            const borrowerMatch = g.borrower && g.borrower.toLowerCase().includes(searchTerm);
            const titleMatch = g.titles.some(t => (t || '').toLowerCase().includes(searchTerm));
            return borrowerMatch || titleMatch;
        });
    }

    // Gender Summary based on grouped borrowers
    let maleCount = 0;
    let femaleCount = 0;
    grouped.forEach(g => {
        if (g.borrower_gender === 'ប្រុស' || g.borrower_gender === 'M') maleCount++;
        if (g.borrower_gender === 'ស្រី' || g.borrower_gender === 'F') femaleCount++;
    });
    loanSummary.textContent = `សរុប: ${grouped.length} នាក់ (ប្រុស: ${maleCount} នាក់, ស្រី: ${femaleCount} នាក់)`;

    // Render grouped rows
    loanList.innerHTML = '';
    if (grouped.length === 0) { loanList.innerHTML = `<tr><td colspan="7" class="text-center p-4 text-gray-500">រកមិនឃើញកំណត់ត្រាខ្ចីទេ។</td></tr>`; return; }

    // Sorting
    const statusOrder = (s) => {
        // Order: ខ្ចី (Borrowed) < សង (Returned) < ចម្រុះ (Mixed)
        if (s === 'ខ្ចី') return 0;
        if (s === 'សង') return 1;
        return 2; // ចម្រុះ or others
    };
    const val = (g, key) => {
        switch (key) {
            case 'borrower': return (g.borrower || '').toLowerCase();
            case 'title': return (g.titles.join(', ') || '').toLowerCase();
            case 'loan_date': return g.loan_date || '';
            case 'return_date': return g.return_date || '';
            case 'status': return statusOrder(g.status);
            case 'serial':
            default:
                return g.loan_date || '';
        }
    };
    const cmp = (a, b) => {
        let A = val(a, loanSortKey);
        let B = val(b, loanSortKey);
        // Date comparisons for date keys
        if (loanSortKey === 'loan_date' || loanSortKey === 'return_date' || loanSortKey === 'serial') {
            // Treat empty/"N/A" as minimal for asc
            const aDate = (A && A !== 'N/A') ? new Date(A) : null;
            const bDate = (B && B !== 'N/A') ? new Date(B) : null;
            if (aDate && bDate) return aDate - bDate;
            if (aDate && !bDate) return 1;
            if (!aDate && bDate) return -1;
            return 0;
        }
        if (loanSortKey === 'status') {
            return A - B;
        }
        // String compare
        return String(A).localeCompare(String(B));
    };
    const sortedGroups = [...grouped].sort((a, b) => {
        const res = cmp(a, b);
        return loanSortDir === 'asc' ? res : -res;
    });
    sortedGroups.forEach((g, index) => {
        const row = document.createElement('tr');
        row.className = 'border-b';
        // Compose titles list with counts: Title(unreturned)(returned)
        const counts = new Map(); // book_id -> { title, unreturned, returned }
        g.loans.forEach(item => {
            const bookId = String(item.book_id);
            const title = (books.find(bb => String(bb.id) === bookId)?.title) || 'N/A';
            if (!counts.has(bookId)) counts.set(bookId, { title, unreturned: 0, returned: 0 });
            if (item.status === 'សង') counts.get(bookId).returned += 1; else counts.get(bookId).unreturned += 1;
        });
        const titlesText = Array.from(counts.values())
            .sort((a, b) => a.title.localeCompare(b.title))
            .map(v => `${v.title}(<span class="text-red-600">${v.unreturned}</span>) (<span class="text-green-600">${v.returned}</span>)`)
            .join(', ');
        // Status badge class
        const statusClass = g.status === 'ខ្ចី' ? 'bg-yellow-200 text-yellow-800' : (g.status === 'សង' ? 'bg-green-200 text-green-800' : 'bg-blue-200 text-blue-800');

        // Build actions: single grouped return and delete buttons
        let actionsHTML = '';
        const hasUnreturned = g.loans.some(item => item.status === 'ខ្ចី');
        const groupKey = `${g.borrower}|${g.loan_date}`;
        if (hasUnreturned) {
            actionsHTML += `<button onclick="window.openGroupedReturnModal('${groupKey}')" class="text-teal-500 hover:text-teal-700 mr-2" title="សងសៀវភៅជាក្រុម"><i class="fas fa-book-reader"></i></button>`;
        }
        actionsHTML += `<button onclick="window.deleteGroupedLoan('${groupKey}')" class="text-red-500 hover:text-red-700 mr-2" title="លុបការខ្ចីជាក្រុម"><i class="fas fa-trash"></i></button>`;

        row.innerHTML = `
            <td class="p-3">${index + 1}</td>
            <td class="p-3">${g.borrower}</td>
            <td class="p-3">${titlesText}</td>
            <td class="p-3">${g.loan_date ? formatDateDDMMYYYY(g.loan_date) : ''}</td>
            <td class="p-3">${g.return_date && g.return_date !== 'N/A' ? formatDateDDMMYYYY(g.return_date) : 'N/A'}</td>
            <td class="p-3"><span class="px-2 py-1 text-xs rounded-full ${statusClass}">${g.status}</span></td>
            <td class="p-3 no-print">${actionsHTML}</td>
        `;
        loanList.appendChild(row);
    });
};

const renderClassLoans = () => {
    const classLoanList = document.getElementById('class-loan-list');
    const selectedClass = document.getElementById('class-loan-filter-select').value;
    
    let filtered = classLoans;
    if (selectedClass) {
        filtered = filtered.filter(loan => loan.class_name === selectedClass);
    }

    classLoanList.innerHTML = '';
    if (filtered.length === 0) { classLoanList.innerHTML = `<tr><td colspan="6" class="text-center p-4 text-gray-500">មិនទាន់មានប្រវត្តិខ្ចីតាមថ្នាក់ទេ។</td></tr>`; return; }

    // Group by class_name + loan_date so multiple books for one class appear on one row
    const keyOf = (l) => `${(l.class_name || '').trim()}|${l.loan_date}`;
    const groupMap = new Map();
    filtered.forEach(l => {
        const key = keyOf(l);
        if (!groupMap.has(key)) groupMap.set(key, []);
        groupMap.get(key).push(l);
    });

    // Build grouped summaries
    const grouped = Array.from(groupMap.entries()).map(([key, arr]) => {
        const [class_name, loan_date] = key.split('|');
        const titles = arr.map(item => {
            const b = books.find(bb => bb.id === item.book_id);
            const bookTitle = b ? b.title : 'សៀវភៅត្រូវបានលុប';
            const loanedQty = item.loaned_quantity || 0;
            const returnedQty = item.returned_count || 0;
            const unreturned = Math.max(0, loanedQty - returnedQty);
            return `${bookTitle} (<span class="text-red-600">${unreturned}</span>) (<span class="text-green-600">${returnedQty}</span>)`;
        });
        const totalLoaned = arr.reduce((s, x) => s + (x.loaned_quantity || 0), 0);
        const totalReturned = arr.reduce((s, x) => s + (x.returned_count || 0), 0);
        const allReturned = totalReturned >= totalLoaned && totalLoaned > 0;
        return { class_name, loan_date, titles, totalLoaned, totalReturned, items: arr, allReturned };
    });

    // Sort by loan_date desc
    const sortedGroups = [...grouped].sort((a, b) => new Date(b.loan_date) - new Date(a.loan_date));

    // Render rows
    sortedGroups.forEach((g, index) => {
        const row = document.createElement('tr');
        row.className = 'border-b';

        // Build actions for grouped items - single smart return button + individual edit/delete
        let actionsHTML = '';
        
        // Add single grouped return button if any books are not fully returned
        const hasUnreturnedBooks = g.items.some(item => (item.returned_count || 0) < (item.loaned_quantity || 0));
        if (hasUnreturnedBooks) {
            const groupKey = `${g.class_name}|${g.loan_date}`;
            actionsHTML += `<button onclick="window.openGroupedClassReturnModal('${groupKey}')" class="text-teal-500 hover:text-teal-700 mr-2" title="សងសៀវភៅតាមថ្នាក់"><i class="fas fa-book-reader"></i></button>`;
        }
        
        // Add single grouped delete button
        const groupKey = `${g.class_name}|${g.loan_date}`;
        actionsHTML += `<button onclick="window.deleteGroupedClassLoan('${groupKey}')" class="text-red-500 hover:text-red-700 mr-2" title="លុបការខ្ចីតាមថ្នាក់"><i class="fas fa-trash"></i></button>`;

        row.innerHTML = `
            <td class="p-3">${index + 1}</td>
            <td class="p-3">${g.class_name}</td>
            <td class="p-3">${g.titles.join(', ')}</td>
            <td class="p-3">${g.loan_date ? formatDateDDMMYYYY(g.loan_date) : ''}</td>
            <td class="p-3">
                <span class="font-bold ${g.allReturned ? 'text-green-600' : 'text-orange-600'}">
                    សងបាន: ${g.totalReturned} / ${g.totalLoaned}
                </span>
            </td>
            <td class="p-3 no-print">${actionsHTML}</td>
        `;
        classLoanList.appendChild(row);
    });
};

const renderLocations = () => {
    const locationList = document.getElementById('location-list');
    const searchLocationsInput = document.getElementById('search-locations');
    const searchTerm = searchLocationsInput.value.toLowerCase();
    const filteredLocations = locations.filter(loc => loc.name.toLowerCase().includes(searchTerm));
    locationList.innerHTML = '';
    if (filteredLocations.length === 0) { locationList.innerHTML = `<tr><td colspan="5" class="text-center p-4 text-gray-500">រកមិនឃើញទីតាំងទេ។</td></tr>`; return; }
    const sortedLocations = [...filteredLocations].sort((a,b) => a.name.localeCompare(b.name));
    sortedLocations.forEach((loc, index) => {
        const row = document.createElement('tr');
        row.className = 'border-b';
        row.innerHTML = `<td class="p-3">${index + 1}</td><td class="p-3">${loc.name}</td><td class="p-3">${loc.source || ''}</td><td class="p-3">${loc.year || ''}</td><td class="p-3 no-print"><button onclick="window.editLocation('${loc.id}')" class="text-blue-500 hover:text-blue-700 mr-2"><i class="fas fa-edit"></i></button><button onclick="window.deleteLocation('${loc.id}')" class="text-red-500 hover:text-red-700"><i class="fas fa-trash"></i></button></td>`;
        locationList.appendChild(row);
    });
};

const renderStudents = () => {
    const studentList = document.getElementById('student-list');
    const studentListHeader = document.getElementById('student-list-header');
    const searchStudentsInput = document.getElementById('search-students');
    const classFilter = document.getElementById('student-class-filter'); // Get the filter
    const duplicateListContainer = document.getElementById('duplicate-student-list');
    
    duplicateListContainer.innerHTML = ''; // Clear previous results

    if (students.length > 0) {
        const idKey = Object.keys(students[0]).find(k => k.toLowerCase().includes('អត្តលេខ'));
        const lastNameKey = Object.keys(students[0]).find(k => k.includes('នាមត្រកូល'));
                const firstNameKey = Object.keys(students[0]).find(k => k.includes('នាមខ្លួន'));
        const classKey = Object.keys(students[0]).find(k => k.includes('ថ្នាក់'));
        const noKey = 'ល.រ';

        if (idKey) {
            const idMap = new Map();
            students.forEach(student => {
                const studentId = (student[idKey] || '').trim();
                if (studentId) {
                    if (!idMap.has(studentId)) {
                        idMap.set(studentId, []);
                    }
                    idMap.get(studentId).push(student);
                }
            });

            const duplicates = [];
            for (const studentGroup of idMap.values()) {
                if (studentGroup.length > 1) {
                    duplicates.push(...studentGroup);
                }
            }

            if (duplicates.length > 0) {
                let html = '<div class="font-sans"><strong>អត្តលេខស្ទួន៖</strong><ul>';
                duplicates.forEach(student => {
                    const fullName = `${student[lastNameKey] || ''} ${student[firstNameKey] || ''}`.trim();
                    const className = student[classKey] || 'N/A';
                    html += `<li class="ml-4">- ល.រ: ${student[noKey] || 'N/A'}, អត្តលេខ: ${student[idKey]}, ឈ្មោះ: ${fullName}, ថ្នាក់: ${className}</li>`;
                });
                html += '</ul></div>';
                duplicateListContainer.innerHTML = html;
            }
        }
    }

    const searchTerm = searchStudentsInput.value.toLowerCase();
    const selectedClass = classFilter.value; 
    const classKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('ថ្នាក់')) : null;

    let filteredStudents = students;

    // Filter by class first
    if (selectedClass && classKey) {
        filteredStudents = filteredStudents.filter(student => student[classKey] === selectedClass);
    }

    // Then filter by search term
    if (searchTerm) {
        filteredStudents = filteredStudents.filter(student => {
            return Object.values(student).some(value =>
                String(value).toLowerCase().includes(searchTerm)
            );
        });
    }

    studentList.innerHTML = '';
    studentListHeader.innerHTML = '';

    if (filteredStudents.length === 0) {
        studentList.innerHTML = `<tr><td colspan="9" class="text-center p-4 text-gray-500">មិនទាន់មានទិន្នន័យសិស្ស ឬរកមិនឃើញ។</td></tr>`;
        return;
    }

    // Show all available columns including new ones
    const headers = ['ល.រ', 'អត្តលេខ', 'នាមត្រកូល', 'នាមខ្លួន', 'ថ្នាក់', 'ភេទ', 'ថ្ងៃខែឆ្នាំកំណើត', 'រូបថត URL', 'សកម្មភាព'];
    
    const sortedStudents = [...filteredStudents].sort((a, b) => {
        const numA = parseInt(a['ល.រ'], 10) || 0;
        const numB = parseInt(b['ល.រ'], 10) || 0;
        return numA - numB;
    });

    // Render headers
    let headerHTML = '';
    headers.forEach(header => {
        const thClass = header === 'សកម្មភាព' ? 'p-3 no-print' : 'p-3';
        headerHTML += `<th class="${thClass}">${header}</th>`;
    });
    studentListHeader.innerHTML = headerHTML;

    // Render rows
    sortedStudents.forEach(student => {
        const row = document.createElement('tr');
        row.className = 'border-b';
        
        let rowHTML = '';
        headers.slice(0, -1).forEach(header => { // All headers except 'Actions'
            rowHTML += `<td class="p-3">${student[header] || ''}</td>`;
        });
        
        // Add actions cell
        rowHTML += `
            <td class="p-3 no-print">
                <button onclick="window.openStudentModal('${student.id}')" class="text-blue-500 hover:text-blue-700 mr-2"><i class="fas fa-edit"></i></button>
                <button onclick="window.deleteStudent('${student.id}')" class="text-red-500 hover:text-red-700"><i class="fas fa-trash"></i></button>
            </td>`;

        row.innerHTML = rowHTML;
        studentList.appendChild(row);
    });
};

const renderReadingLogs = () => {
    const readingLogHistory = document.getElementById('reading-log-history');
    const readingLogSummary = document.getElementById('reading-log-summary');
    const startDate = document.getElementById('reading-log-filter-start-date').value;
    const endDate = document.getElementById('reading-log-filter-end-date').value;

    let filtered = readingLogs;
    if (startDate) {
        filtered = filtered.filter(log => log.date_time.split('T')[0] >= startDate);
    }
    if (endDate) {
        filtered = filtered.filter(log => log.date_time.split('T')[0] <= endDate);
    }

    // Gender Summary
    let maleCount = 0;
    let femaleCount = 0;
    filtered.forEach(log => {
        if (log.student_gender === 'ប្រុស' || log.student_gender === 'M') maleCount++;
        if (log.student_gender === 'ស្រី' || log.student_gender === 'F') femaleCount++;
    });
    readingLogSummary.textContent = `សរុប: ${filtered.length} នាក់ (ប្រុស: ${maleCount} នាក់, ស្រី: ${femaleCount} នាក់)`;

    readingLogHistory.innerHTML = '';
    if(filtered.length === 0) {
        readingLogHistory.innerHTML = `<tr><td colspan="5" class="text-center p-4 text-gray-500">មិនទាន់មានប្រវត្តិការចូលអានទេ។</td></tr>`;
        return;
    }

    // Apply sorting
    const sortedLogs = [...filtered];
    if (readingLogSortKey === 'title') {
        sortedLogs.sort((a, b) => {
            const ta = getBooksReadString(a).toLowerCase();
            const tb = getBooksReadString(b).toLowerCase();
            const cmp = ta.localeCompare(tb);
            return readingLogSortDir === 'asc' ? cmp : -cmp;
        });
    } else {
        // default by date_time
        sortedLogs.sort((a, b) => {
            const cmp = new Date(a.date_time) - new Date(b.date_time);
            return readingLogSortDir === 'asc' ? cmp : -cmp;
        });
    }

    // Build groups by date (YYYY-MM-DD)
    const byDate = new Map();
    sortedLogs.forEach(log => {
        const key = (log.date_time || '').split('T')[0] || 'N/A';
        if (!byDate.has(key)) byDate.set(key, []);
        byDate.get(key).push(log);
    });

    // Sort dates according to current sort key/dir
    const dateKeys = Array.from(byDate.keys());
    const sortedDates = dateKeys.sort((a, b) => {
        // Default: sort by date
        const cmp = new Date(a) - new Date(b);
        return readingLogSortDir === 'asc' ? cmp : -cmp;
    });

    // Helper for extracting titles array from a log entry
    const getBooksArrayForAgg = (log) => {
        if (!log || !log.books) return [];
        try {
            const data = typeof log.books === 'string' ? JSON.parse(log.books) : log.books;
            if (Array.isArray(data)) {
                return data.map(b => (typeof b === 'string' ? b : (b.title || ''))).filter(Boolean);
            }
            if (typeof data === 'object' && data) {
                return [data.title || data.name || String(data)].filter(Boolean);
            }
            return [String(data)].filter(Boolean);
        } catch (e) {
            return [String(log.books)].filter(Boolean);
        }
    };

    // Render grouped: date header -> student rows
    let dateIndex = 0;
    sortedDates.forEach(dateKey => {
        dateIndex += 1;
        const logsInDate = byDate.get(dateKey) || [];

        // Aggregate by student within the date
        const byStudent = new Map(); // name -> array of { title, id }
        logsInDate.forEach(l => {
            const name = (l.student_name || 'មិនស្គាល់').trim();
            const titles = getBooksArrayForAgg(l);
            if (!byStudent.has(name)) byStudent.set(name, []);
            byStudent.get(name).push(...titles.map(t => ({ title: t, id: l.id })));
        });

        // Optionally sort students
        const studentsSorted = Array.from(byStudent.entries()).sort((a, b) => {
            if (readingLogSortKey === 'title') {
                const aStr = (a[1] || []).map(x => x.title).join(', ').toLowerCase();
                const bStr = (b[1] || []).map(x => x.title).join(', ').toLowerCase();
                return readingLogSortDir === 'asc' ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
            }
            // default sort by student name asc
            return a[0].localeCompare(b[0]);
        });

        // Date header row
        const headerTr = document.createElement('tr');
        headerTr.className = 'bg-gray-100';
        const displayDate = (() => {
            try {
                const dt = new Date(dateKey);
                const dd = String(dt.getDate()).padStart(2, '0');
                const mm = String(dt.getMonth() + 1).padStart(2, '0');
                const yyyy = dt.getFullYear();
                return `${dd}/${mm}/${yyyy}`;
            } catch (_) { return dateKey; }
        })();
        headerTr.innerHTML = `
            <td class="p-3 font-semibold">${dateIndex}</td>
            <td class="p-3 font-semibold no-wrap" colspan="4">កាលបរិច្ឆេទ: ${displayDate}</td>
        `;
        readingLogHistory.appendChild(headerTr);

        // Student rows under this date
        studentsSorted.forEach(([studentName, entries]) => {
            const booksStr = (entries || []).map(x => x.title).join(', ');
            const delBtn = `
                <button onclick="window.deleteStudentLogsForDateAndStudent('${encodeURIComponent(dateKey)}','${encodeURIComponent(studentName)}')"
                        class="text-red-500 hover:text-red-700 inline-block"
                        title="លុបកំណត់ត្រ សម្រាប់សិស្សម្នាក់ក្នុងថ្ងៃនេះ">
                    <i class="fas fa-trash"></i>
                </button>
            `;
            const tr = document.createElement('tr');
            tr.className = 'border-b';
            tr.innerHTML = `
                <td class="p-3"></td>
                <td class="p-3"></td>
                <td class="p-3 no-wrap">${studentName}</td>
                <td class="p-3">${booksStr}</td>
                <td class="p-3 no-print">${delBtn}</td>
            `;
            readingLogHistory.appendChild(tr);
        });
    });

    // Grouped-by-book rendering
    const groupedBody = document.getElementById('reading-log-grouped-body');
    if (groupedBody) {
        groupedBody.innerHTML = '';

        // Helper: extract array of titles from a log
        const getBooksArray = (log) => {
            if (!log || !log.books) return [];
            try {
                const data = typeof log.books === 'string' ? JSON.parse(log.books) : log.books;
                if (Array.isArray(data)) {
                    return data.map(b => (typeof b === 'string' ? b : (b.title || ''))).filter(Boolean);
                }
                if (typeof data === 'object' && data) {
                    return [data.title || data.name || String(data)].filter(Boolean);
                }
                return [String(data)].filter(Boolean);
            } catch (e) {
                return [String(log.books)].filter(Boolean);
            }
        };

        // Build: title -> student -> count
        const byTitle = new Map();
        filtered.forEach(log => {
            const titles = getBooksArray(log);
            const student = log.student_name || 'មិនស្គាល់';
            titles.forEach(title => {
                const key = title.trim();
                if (!key) return;
                if (!byTitle.has(key)) byTitle.set(key, new Map());
                const studentsMap = byTitle.get(key);
                studentsMap.set(student, (studentsMap.get(student) || 0) + 1);
            });
        });

        if (byTitle.size === 0) {
            groupedBody.innerHTML = `<tr><td colspan="3" class="text-center p-4 text-gray-500">មិនទាន់មានទិន្នន័យ</td></tr>`;
            return;
        }

        // Sort titles alphabetically
        const titlesSorted = Array.from(byTitle.keys()).sort((a, b) => a.localeCompare(b));
        titlesSorted.forEach((title, idx) => {
            const studentsMap = byTitle.get(title);
            // Sort students by count desc then name asc
            const studentsList = Array.from(studentsMap.entries())
                .sort((a, b) => (b[1] - a[1]) || a[0].localeCompare(b[0]))
                .map(([name, count]) => `${name} (${count}ដង)`);
            const tr = document.createElement('tr');
            tr.className = 'border-b';
            tr.innerHTML = `
                <td class="p-3">${idx + 1}</td>
                <td class="p-3">${title}</td>
                <td class="p-3">${studentsList.join(', ')}</td>
            `;
            groupedBody.appendChild(tr);
        });
    }
};

const updateDashboard = () => {
    document.getElementById('total-books').textContent = books.length;
    const totalQuantity = books.reduce((sum, book) => sum + (parseInt(book.quantity, 10) || 0), 0);
    document.getElementById('total-quantity').textContent = totalQuantity;
    // Active individual loans (not part of class loans)
    const individualActive = loans.filter(l => l.status === 'ខ្ចី' && !l.class_loan_id).length;

    // For class loans: prefer counting linked active individual loan rows if they exist; otherwise
    // fall back to outstanding quantity from class_loans (loaned_quantity - returned_count)
    const linkedActiveByClassId = new Map();
    loans.forEach(l => {
        if (l.status === 'ខ្ចី' && l.class_loan_id) {
            const key = String(l.class_loan_id);
            linkedActiveByClassId.set(key, (linkedActiveByClassId.get(key) || 0) + 1);
        }
    });

    let classActiveTotal = 0;
    classLoans.forEach(cl => {
        const key = String(cl.id);
        const linkedCount = linkedActiveByClassId.get(key) || 0;
        if (linkedCount > 0) {
            classActiveTotal += linkedCount;
        } else {
            const loaned = Number(cl.loaned_quantity || 0);
            const returned = Number(cl.returned_count || 0);
            const remaining = Math.max(0, loaned - returned);
            classActiveTotal += remaining;
        }
    });

    const totalActiveLoans = individualActive + classActiveTotal;
    document.getElementById('total-loans').textContent = totalActiveLoans;
};

// Update homepage reading log gender stats (total, male, female + bars) + top book
function updateHomeReadingStats() {
    const totalEl = document.getElementById('home-readings-total');
    const maleEl = document.getElementById('home-readings-male');
    const femaleEl = document.getElementById('home-readings-female');
    const maleBar = document.getElementById('home-readings-male-bar');
    const femaleBar = document.getElementById('home-readings-female-bar');
    const topBookEl = document.getElementById('home-top-book');
    if (!totalEl || !maleEl || !femaleEl || !maleBar || !femaleBar || !topBookEl) return;

    const total = readingLogs.length;
    let male = 0, female = 0;
    readingLogs.forEach(log => {
        const g = (log.student_gender || '').trim();
        if (g === 'ប្រុស' || g === 'M') male++;
        else if (g === 'ស្រី' || g === 'F') female++;
    });

    totalEl.textContent = String(total);
    maleEl.textContent = String(male);
    femaleEl.textContent = String(female);

    const denom = total > 0 ? total : 1;
    const malePct = Math.min(100, Math.round((male / denom) * 100));
    const femalePct = Math.min(100, Math.round((female / denom) * 100));
    maleBar.style.width = malePct + '%';
    femaleBar.style.width = femalePct + '%';

    // Find most read book with title and author
    const bookCounts = new Map();
    readingLogs.forEach(log => {
        if (log.books) {
            try {
                const booksData = typeof log.books === 'string' ? JSON.parse(log.books) : log.books;
                if (Array.isArray(booksData)) {
                    booksData.forEach(b => {
                        const title = b.title || b;
                        if (title) {
                            // Find book in books array to get author
                            const bookRecord = books.find(book => book.title === title);
                            const author = bookRecord?.author || '';
                            const key = title;
                            const displayText = author ? `${title} ${author}` : title;
                            
                            if (!bookCounts.has(key)) {
                                bookCounts.set(key, { count: 0, displayText });
                            }
                            bookCounts.get(key).count++;
                        }
                    });
                } else {
                    const title = booksData.title || booksData;
                    if (title) {
                        const bookRecord = books.find(book => book.title === title);
                        const author = bookRecord?.author || '';
                        const key = title;
                        const displayText = author ? `${title} ${author}` : title;
                        
                        if (!bookCounts.has(key)) {
                            bookCounts.set(key, { count: 0, displayText });
                        }
                        bookCounts.get(key).count++;
                    }
                }
            } catch (e) {
                const title = log.books.toString();
                if (title) {
                    const bookRecord = books.find(book => book.title === title);
                    const author = bookRecord?.author || '';
                    const key = title;
                    const displayText = author ? `${title} ${author}` : title;
                    
                    if (!bookCounts.has(key)) {
                        bookCounts.set(key, { count: 0, displayText });
                    }
                    bookCounts.get(key).count++;
                }
            }
        }
    });

    if (bookCounts.size === 0) {
        topBookEl.textContent = 'មិនទាន់មានទិន្នន័យ';
    } else {
        const topBook = Array.from(bookCounts.entries())
            .sort((a, b) => b[1].count - a[1].count)[0];
        topBookEl.innerHTML = `<i class="fas fa-star text-yellow-400 mr-1" aria-hidden="true"></i> សៀវភៅមើលច្រើនជាងគេ៖ ${topBook[1].displayText} (${topBook[1].count}ដង)`;
    }
}

// Update homepage most read books section
function updateHomeMostReadBooks() {
    const container = document.getElementById('home-most-read-books');
    if (!container) return;

    // Count book occurrences in reading logs
    const bookCounts = new Map();
    readingLogs.forEach(log => {
        if (log.books) {
            try {
                // Handle both JSON string and array formats
                const booksData = typeof log.books === 'string' ? JSON.parse(log.books) : log.books;
                if (Array.isArray(booksData)) {
                    booksData.forEach(b => {
                        const title = b.title || b;
                        if (title) {
                            bookCounts.set(title, (bookCounts.get(title) || 0) + 1);
                        }
                    });
                } else {
                    const title = booksData.title || booksData;
                    if (title) {
                        bookCounts.set(title, (bookCounts.get(title) || 0) + 1);
                    }
                }
            } catch (e) {
                // If parsing fails, try to display as string
                const title = log.books.toString();
                if (title) {
                    bookCounts.set(title, (bookCounts.get(title) || 0) + 1);
                }
            }
        }
    });

    // Sort by count descending and take top 5
    const sortedBooks = Array.from(bookCounts.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

    if (sortedBooks.length === 0) {
        container.innerHTML = '<p class="text-gray-500 text-sm">មិនទាន់មានទិន្នន័យ</p>';
        return;
    }

    // Render the list
    container.innerHTML = sortedBooks.map((entry, index) => {
        const [title, count] = entry;
        return `
            <div class="flex items-center justify-between py-2 border-b border-gray-100 last:border-b-0">
                <div class="flex items-center">
                    <span class="text-sm font-bold text-gray-500 mr-3">${index + 1}.</span>
                    <span class="text-gray-800">${title}</span>
                </div>
                <span class="text-sm font-semibold text-blue-600">${count} ដង</span>
            </div>
        `;
    }).join('');
}

// Render summary list of active individual loans on Home page (plain text)
const renderHomeLoans = () => {
    const leftEl = document.getElementById('home-loans-text-left');
    const rightEl = document.getElementById('home-loans-text-right');
    if (!leftEl || !rightEl) return;

    // Active individual loans only
    const activeIndLoans = loans.filter(l => !l.class_loan_id && l.status === 'ខ្ចី');

    if (activeIndLoans.length === 0) {
        leftEl.innerHTML = 'គ្មានទិន្នន័យ';
        rightEl.innerHTML = '';
        return;
    }

    // Group: borrower -> { books: Map(book_id->qty), latestDate: string }
    const borrowerMap = new Map();
    activeIndLoans.forEach(l => {
        const borrower = (l.borrower || '').trim();
        if (!borrowerMap.has(borrower)) borrowerMap.set(borrower, { books: new Map(), latestDate: null });
        const entry = borrowerMap.get(borrower);
        const key = String(l.book_id);
        entry.books.set(key, (entry.books.get(key) || 0) + 1);
        const dateStr = l.loan_date || l.loanDate || null;
        if (dateStr) {
            if (!entry.latestDate || new Date(dateStr) > new Date(entry.latestDate)) {
                entry.latestDate = dateStr;
            }
        }
    });

    // Build lines
    const formatKhmerDate = (dateStr) => {
        const d = new Date(dateStr);
        if (isNaN(d)) return dateStr || '';
        const day = String(d.getDate()).padStart(2, '0');
        const month = d.getMonth() + 1;
        const year = d.getFullYear();
        const khmerMonths = {
            1: 'មករា', 2: 'កុម្ភៈ', 3: 'មីនា', 4: 'មេសា', 5: 'ឧសភា', 6: 'មិថុនា',
            7: 'កក្កដា', 8: 'សីហា', 9: 'កញ្ញា', 10: 'តុលា', 11: 'វិច្ឆិកា', 12: 'ធ្នូ'
        };
        return `${day}/${khmerMonths[month]}/${year}`;
    };
    const borrowers = Array.from(borrowerMap.keys()).sort((a, b) => a.localeCompare(b));
    const lines = [];
    borrowers.forEach(name => {
        const entry = borrowerMap.get(name);
        const entries = Array.from(entry.books.entries()).map(([bookId, qty]) => {
            const book = books.find(b => b.id === bookId);
            const title = book && book.title ? book.title : 'សៀវភៅត្រូវបានលុប';
            return `<span class="text-black">${title}</span>(<span class="text-red-600">${qty}</span>)`;
        }).sort((a, b) => a.localeCompare(b));
        const dateSuffix = entry.latestDate ? `  <span class="text-black">ថ្ងៃខ្ចី: ${formatKhmerDate(entry.latestDate)}</span>` : '';
        lines.push(`- <span class="text-black">${name}</span>: ${entries.join(', ')}${dateSuffix}`);
    });

    const mid = Math.ceil(lines.length / 2);
    const leftLinesHTML = lines.slice(0, mid).join('<br>');
    const rightLinesHTML = lines.slice(mid).join('<br>');
    leftEl.innerHTML = leftLinesHTML;
    rightEl.innerHTML = rightLinesHTML;
};

const populateClassLoanFilter = () => {
    const classLoanFilterSelect = document.getElementById('class-loan-filter-select');
    classLoanFilterSelect.innerHTML = '<option value="">-- ថ្នាក់ទាំងអស់ --</option>';
    if (students.length > 0) {
        const classKey = Object.keys(students[0]).find(k => k.includes('ថ្នាក់'));
        if (classKey) {
            const uniqueClasses = [...new Set(students.map(s => s[classKey]))].filter(Boolean).sort();
            uniqueClasses.forEach(className => {
                const option = document.createElement('option');
                option.value = className;
                option.textContent = className;
                classLoanFilterSelect.appendChild(option);
            });
        }
    }
};

// START: NEW FUNCTION TO POPULATE STUDENT PAGE CLASS FILTER
const populateStudentClassFilter = () => {
    const classFilter = document.getElementById('student-class-filter');
    // Only populate if there are students and the filter is empty (or has only the 'all' option)
    if (!students.length || (classFilter.options.length > 1 && classFilter.value !== '')) { 
        return;
    }

    const classKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('ថ្នាក់')) : null;
    if (!classKey) return;

    // Get unique classes, filter out any empty values, and sort them naturally (e.g., 1, 2, 10)
    const uniqueClasses = [...new Set(students.map(s => s[classKey]))]
        .filter(Boolean)
        .sort((a, b) => a.localeCompare(b, undefined, {numeric: true})); 
    
    const currentVal = classFilter.value; // Save current selection
    classFilter.innerHTML = '<option value="">-- ថ្នាក់ទាំងអស់ --</option>'; 

    uniqueClasses.forEach(className => {
        const option = document.createElement('option');
        option.value = className;
        option.textContent = className;
        classFilter.appendChild(option);
    });

    // Set default to the smallest class if no class was previously selected
    if (!currentVal && uniqueClasses.length > 0) {
        classFilter.value = uniqueClasses[0];
    } else {
        classFilter.value = currentVal; // Restore previous selection
    }
};
// END: NEW FUNCTION

// --- STUDENT CARD PAGE ---
// MODIFIED FUNCTION: Creates paginated containers for printing
function renderStudentCards() {
    if (!document.getElementById('page-student-cards')) return;

    const container = document.getElementById("student-card-container");
    const loading = document.getElementById("loading-cards");
    const classFilter = document.getElementById("card-class-filter");
    const searchBox = document.getElementById("card-search-box");
    
    // Update the dynamic style for the card background
    let styleElement = document.getElementById('dynamic-card-style');
    if (!styleElement) {
        styleElement = document.createElement('style');
        styleElement.id = 'dynamic-card-style';
        document.head.appendChild(styleElement);
    }
    const cardBgUrl = settingsData.cardBgUrl || 'https://i.imgur.com/s46369v.png'; // Default MOEYS logo
    styleElement.innerHTML = `#page-student-cards .card::before { background-image: url('${cardBgUrl}'); }`;


    // Populate class filter only once
    if (classFilter.options.length <= 1 && students.length > 0) {
        const classKey = Object.keys(students[0]).find(k => k.includes('ថ្នាក់'));
        if (classKey) {
            const classSet = new Set(students.map(std => std[classKey]).filter(Boolean));
            const sortedClasses = [...classSet].sort((a, b) => a.localeCompare(b, undefined, {numeric: true}));
            
            sortedClasses.forEach(cls => {
                const option = document.createElement('option');
                option.value = cls;
                option.textContent = cls;
                classFilter.appendChild(option);
            });
            
            // Set default to the smallest class (first in sorted array)
            if (sortedClasses.length > 0) {
                classFilter.value = sortedClasses[0];
            }
        }
    }
    
    const selectedClass = classFilter.value;
    const keyword = searchBox.value.toLowerCase();
    const classKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('ថ្នាក់')) : null;
    const nameKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('នាមខ្លួន')) : null;
    const lastNameKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('នាមត្រកូល')) : null;
    const idKey = students.length > 0 ? Object.keys(students[0]).find(k => k.toLowerCase().includes('អត្តលេខ')) : null;

    const filteredStudents = students.filter(std => {
        const fullName = `${std[lastNameKey] || ''} ${std[nameKey] || ''}`.toLowerCase();
        const studentId = (std[idKey] || '').toLowerCase();
        const classMatch = !selectedClass || (classKey && std[classKey] === selectedClass);
        const keywordMatch = !keyword || fullName.includes(keyword) || studentId.includes(keyword);
        return classMatch && keywordMatch;
    });

    loading.style.display = "none";
    container.style.display = "flex"; // Keep flex for on-screen view
    container.innerHTML = ""; // Clear previous content

    if (filteredStudents.length === 0) {
        container.innerHTML = `<p class="text-center text-gray-500 w-full p-4">រកមិនឃើញទិន្នន័យសិស្សទេ។</p>`;
        return;
    }

    const schoolName = settingsData.schoolName || 'ឈ្មោះសាលា';
    const academicYear = settingsData.academicYear || 'YYYY-YYYY';
    const sealUrl = settingsData.sealImageUrl || '';

    const CARDS_PER_PAGE = 9;
    let pageContainer = null;

    filteredStudents.forEach((std, index) => {
        // Create a new page container for the first card and every 9th card after that
        if (index % CARDS_PER_PAGE === 0) {
            pageContainer = document.createElement('div');
            pageContainer.className = 'print-page-container'; // New class for print styling
            container.appendChild(pageContainer);
        }

        const cardDiv = document.createElement("div");
        cardDiv.className = "card";
        const qrId = `qr-${std.id}`;
        const barcodeId = `barcode-${std.id}`;
        
        const dobKey = Object.keys(std).find(k => k.includes('ថ្ងៃខែឆ្នាំកំណើត'));
        const genderKey = Object.keys(std).find(k => k.includes('ភេទ'));
        const photoUrlKey = Object.keys(std).find(k => k.includes('រូបថត URL'));
        
        const fullName = `${std[lastNameKey] || ''} ${std[nameKey] || ''}`.trim();
        const studentClass = classKey ? `${std[classKey]}` : '';
        const studentId = idKey ? std[idKey] : '';
        const dob = dobKey && std[dobKey] ? std[dobKey] : '';
        const gender = genderKey && std[genderKey] ? std[genderKey] : '';
        const photoUrl = photoUrlKey && std[photoUrlKey] ? std[photoUrlKey].trim() : null;

        const photoElement = photoUrl 
            ? `<img class="photo" src="${photoUrl}" alt="Photo" onerror="this.onerror=null;this.src='https://placehold.co/108x108/e2e8f0/7d7d7d?text=Error';"/>`
            : `<div class="photo no-photo"></div>`;

        cardDiv.innerHTML = `
          <div>
            <div class="school-name">${schoolName}</div>
            <div class="student-name">${fullName}</div>
            <div class="info-and-qr">
              <div class="info">
                <div><b>អត្តលេខ:</b> ${studentId}</div>
                <div><b>ថ្នាក់:</b> ${studentClass}</div>
                <div><b>ឆ្នាំសិក្សា:</b> ${academicYear}</div>
                <div><b>ភេទ:</b> ${gender}</div>
              </div>
              <div id="${qrId}" class="qr"></div>
            </div>
          </div>
          <div class="barcode-container">
             <svg id="${barcodeId}"></svg>
          </div>
          <div class="photo-wrapper">
            ${photoElement}
          </div>
          ${sealUrl ? `<img class="stamp" src="${sealUrl}" alt="Stamp" />` : ''}
        `;
        
        // Append the card to the current page container
        pageContainer.appendChild(cardDiv);
    });

    // Generate QR codes and barcodes after all cards are rendered with a longer delay
    // to ensure proper rendering before printing
    setTimeout(() => {
        console.log('Generating QR codes and barcodes for', filteredStudents.length, 'students');
        // Add a small delay for each student to prevent browser from freezing
        let processedCount = 0;
        
        function processNextStudent() {
            if (processedCount >= filteredStudents.length) {
                console.log('Finished generating all QR codes and barcodes');
                return;
            }
            
            const std = filteredStudents[processedCount];
            processedCount++;
            
            const qrId = `qr-${std.id}`;
            const barcodeId = `barcode-${std.id}`;
            const studentId = idKey ? std[idKey] : '';

            if (studentId && window.QRCode) {
                const qrElement = document.getElementById(qrId);
                if (qrElement) {
                    qrElement.innerHTML = ''; // Clear any existing content
                    try {
                        new QRCode(qrElement, {
                            text: studentId,
                            width: 70,
                            height: 70,
                            correctLevel: QRCode.CorrectLevel.H
                        });
                    } catch (e) {
                        console.error("QR code generation failed for ID:", studentId, e);
                    }
                }
            }
            
            if (studentId && window.JsBarcode) {
                try {
                    const barcodeElement = document.getElementById(barcodeId);
                    if (barcodeElement) {
                        JsBarcode(`#${barcodeId}`, studentId, {
                            format: "CODE128",
                            height: 30,
                            width: 1.8,
                            displayValue: false,
                            margin: 0
                        });
                    }
                } catch (e) {
                    console.error("Barcode generation failed for ID:", studentId, e);
                    const barcodeElement = document.getElementById(barcodeId);
                    if (barcodeElement) {
                        barcodeElement.style.display = 'none';
                    }
                }
            }
            
            // Process next student with a small delay to prevent browser freezing
            setTimeout(processNextStudent, 10);
        }
        
        // Start processing students
        processNextStudent();
    }, 500);
};

// --- SUPABASE REALTIME LISTENERS ---
const setupRealtimeListeners = async (userId) => {
    // Load initial data
    await loadInitialData(userId);
    
    // Set up real-time subscriptions
    const booksSubscription = supabase
        .channel('books')
        .on('postgres_changes', { event: '*', schema: 'public', table: 'books', filter: `user_id=eq.${userId}` }, 
            async () => {
                await loadBooks(userId);
                renderAll();
            })
        .subscribe();
    
    const loansSubscription = supabase
        .channel('loans')
        .on('postgres_changes', { event: '*', schema: 'public', table: 'loans', filter: `user_id=eq.${userId}` }, 
            async () => {
                await loadLoans(userId);
                renderAll();
            })
        .subscribe();
    
    const classLoansSubscription = supabase
        .channel('class_loans')
        .on('postgres_changes', { event: '*', schema: 'public', table: 'class_loans', filter: `user_id=eq.${userId}` }, 
            async () => {
                await loadClassLoans(userId);
                renderAll();
            })
        .subscribe();
    
    const locationsSubscription = supabase
        .channel('locations')
        .on('postgres_changes', { event: '*', schema: 'public', table: 'locations', filter: `user_id=eq.${userId}` }, 
            async () => {
                await loadLocations(userId);
                renderAll();
            })
        .subscribe();
    
    const studentsSubscription = supabase
        .channel('students')
        .on('postgres_changes', { event: '*', schema: 'public', table: 'students', filter: `user_id=eq.${userId}` }, 
            async () => {
                await loadStudents(userId);
                populateClassLoanFilter();
                populateStudentClassFilter();
                renderAll();
            })
        .subscribe();
    
    const readingLogsSubscription = supabase
        .channel('reading_logs')
        .on('postgres_changes', { event: '*', schema: 'public', table: 'reading_logs', filter: `user_id=eq.${userId}` }, 
            async () => {
                await loadReadingLogs(userId);
                renderAll();
            })
        .subscribe();
    
    const settingsSubscription = supabase
        .channel('settings')
        .on('postgres_changes', { event: '*', schema: 'public', table: 'settings', filter: `user_id=eq.${userId}` }, 
            async () => {
                await loadSettings(userId);
                renderAll();
            })
        .subscribe();
    
    // Store unsubscribe functions
    unsubscribeBooks = () => supabase.removeChannel(booksSubscription);
    unsubscribeLoans = () => supabase.removeChannel(loansSubscription);
    unsubscribeClassLoans = () => supabase.removeChannel(classLoansSubscription);
    unsubscribeLocations = () => supabase.removeChannel(locationsSubscription);
    unsubscribeStudents = () => supabase.removeChannel(studentsSubscription);
    unsubscribeReadingLogs = () => supabase.removeChannel(readingLogsSubscription);
    unsubscribeSettings = () => supabase.removeChannel(settingsSubscription);
};

// Load initial data functions
async function loadInitialData(userId) {
    await Promise.all([
        loadBooks(userId),
        loadLoans(userId),
        loadClassLoans(userId),
        loadLocations(userId),
        loadStudents(userId),
        loadReadingLogs(userId),
        loadSettings(userId)
    ]);
    // Ensure filters that depend on students are populated on first load
    populateClassLoanFilter();
    populateStudentClassFilter();
    renderAll();
}

async function loadBooks(userId) {
    const { data, error } = await supabase
        .from('books')
        .select('*')
        .eq('user_id', userId);

    if (error) {
        console.error('Error loading books:', error);
        books = [];
    } else {
        books = data;
        populateAuthorFilter();
    }
}

async function loadLoans(userId) {
    const { data, error } = await supabase
        .from('loans')
        .select('*')
        .eq('user_id', userId);
    if (!error) {
        loans = data || [];
    }
}

async function loadClassLoans(userId) {
    const { data, error } = await supabase
        .from('class_loans')
        .select('*')
        .eq('user_id', userId);
    if (!error) {
        classLoans = data || [];
    }
}

async function loadLocations(userId) {
    const { data, error } = await supabase
        .from('locations')
        .select('*')
        .eq('user_id', userId);
    if (!error) {
        locations = data || [];
    }
}

async function loadStudents(userId) {
    // Fetch all students in pages to avoid the default 1000-row limit
    const pageSize = 1000;
    let offset = 0;
    let allRows = [];
    while (true) {
        const { data, error } = await supabase
            .from('students')
            .select('*')
            .eq('user_id', userId)
            .range(offset, offset + pageSize - 1);
        if (error) { console.error('Error loading students:', error); break; }
        const batch = data || [];
        allRows = allRows.concat(batch);
        if (batch.length < pageSize) break; // last page
        offset += pageSize;
    }
    if (allRows.length >= 0) {
        // Normalize to Khmer keys expected by the UI
        const rows = allRows;
        students = rows.map((r) => {
            // Handle both old format (name field) and new format (separate fields)
            let lastName = r['នាមត្រកូល'] || '';
            let firstName = r['នាមខ្លួន'] || '';
            
            // Fallback to old format if new format is empty
            if (!lastName && !firstName && r.name) {
                const fullName = (r.name || '').trim();
                if (fullName) {
                    const parts = fullName.split(/\s+/);
                    lastName = parts.shift() || '';
                    firstName = parts.join(' ');
                }
            }
            
            return {
                id: r.id,
                user_id: r.user_id,
                'ល.រ': r['ល.រ'] || r.serial_number || '',
                'អត្តលេខ': r['អត្តលេខ'] || r.student_id || '',
                'នាមត្រកូល': lastName,
                'នាមខ្លួន': firstName,
                'ភេទ': r['ភេទ'] || r.gender || '',
                'ថ្នាក់': r['ថ្នាក់'] || r.class || '',
                'ថ្ងៃខែឆ្នាំកំណើត': r['ថ្ងៃខែឆ្នាំកំណើត'] || r.date_of_birth || '',
                'រូបថត URL': r['រូបថត URL'] || r.photo_url || ''
            };
        });
    }
}

async function loadReadingLogs(userId) {
    const { data, error } = await supabase
        .from('reading_logs')
        .select('*')
        .eq('user_id', userId);
    if (!error) {
        readingLogs = data || [];
    }
}

async function loadSettings(userId) {
    const { data, error } = await supabase
        .from('settings')
        .select('*')
        .eq('user_id', userId);
    
    // Convert key-value pairs back to object structure
    const settingsObj = {};
    if (!error && data && data.length > 0) {
        data.forEach(setting => {
            if (setting.key && setting.value !== null) {
                settingsObj[setting.key] = setting.value;
            }
        });
    }
    
    settingsData = settingsObj;
    updateSettingsUI(settingsObj);
    renderStudentCards();
}

function updateSettingsUI(data) {
    const schoolNameInput = document.getElementById('school-name-input');
    const schoolNameDisplay = document.getElementById('school-name-display');
    const saveSchoolNameBtn = document.getElementById('save-school-name-btn');
    const editSchoolNameBtn = document.getElementById('edit-school-name-btn');
    const deleteSchoolNameBtn = document.getElementById('delete-school-name-btn');
    const academicYearInput = document.getElementById('academic-year-input');
    const academicYearDisplay = document.getElementById('academic-year-display');
    const saveAcademicYearBtn = document.getElementById('save-academic-year-btn');
    const editAcademicYearBtn = document.getElementById('edit-academic-year-btn');
    const deleteAcademicYearBtn = document.getElementById('delete-academic-year-btn');
    const sealImageUrlInput = document.getElementById('seal-image-url');
    const sealImagePreview = document.getElementById('seal-image-preview');
    const googleSheetUrlInput = document.getElementById('google-sheet-url');
    const cardBgUrlInput = document.getElementById('card-bg-url');
    const cardBgPreview = document.getElementById('card-bg-preview');
    const printSchoolName = document.getElementById('print-school-name');

    // School Name handling with null checks
    if (schoolNameInput && schoolNameDisplay && saveSchoolNameBtn && editSchoolNameBtn && deleteSchoolNameBtn) {
        if (data.schoolName || data.school_name) {
            const schoolName = data.schoolName || data.school_name;
            if (sidebarSchoolName) sidebarSchoolName.textContent = schoolName;
            if (printSchoolName) printSchoolName.textContent = schoolName;
            schoolNameDisplay.textContent = schoolName;
            schoolNameInput.value = schoolName;
            schoolNameDisplay.classList.remove('hidden');
            schoolNameInput.classList.add('hidden');
            saveSchoolNameBtn.classList.add('hidden');
            editSchoolNameBtn.classList.remove('hidden');
            deleteSchoolNameBtn.classList.remove('hidden');
        } else {
            if (sidebarSchoolName) sidebarSchoolName.textContent = '';
            if (printSchoolName) printSchoolName.textContent = '';
            // Clear input when school name has been deleted
            schoolNameInput.value = '';
            schoolNameDisplay.classList.add('hidden');
            schoolNameInput.classList.remove('hidden');
            saveSchoolNameBtn.classList.remove('hidden');
            editSchoolNameBtn.classList.add('hidden');
            deleteSchoolNameBtn.classList.add('hidden');
        }
    }

    // Academic Year handling with null checks
    if (academicYearInput && academicYearDisplay && saveAcademicYearBtn && editAcademicYearBtn && deleteAcademicYearBtn) {
        if (data.academicYear || data.academic_year) {
            const academicYear = data.academicYear || data.academic_year;
            academicYearDisplay.textContent = academicYear;
            academicYearInput.value = academicYear;
            academicYearDisplay.classList.remove('hidden');
            academicYearInput.classList.add('hidden');
            saveAcademicYearBtn.classList.add('hidden');
            editAcademicYearBtn.classList.remove('hidden');
            deleteAcademicYearBtn.classList.remove('hidden');
        } else {
            // Clear input when academic year has been deleted
            academicYearInput.value = '';
            academicYearDisplay.classList.add('hidden');
            academicYearInput.classList.remove('hidden');
            saveAcademicYearBtn.classList.remove('hidden');
            editAcademicYearBtn.classList.add('hidden');
            deleteAcademicYearBtn.classList.add('hidden');
        }
    }

    // Seal Image URL handling with null checks
    if (sealImageUrlInput && sealImagePreview) {
        if (data.sealImageUrl || data.seal_image_url) {
            const sealImageUrl = data.sealImageUrl || data.seal_image_url;
            sealImageUrlInput.value = sealImageUrl;
            sealImagePreview.src = sealImageUrl;
        } else {
            sealImageUrlInput.value = '';
            sealImagePreview.src = 'https://placehold.co/100x100/e2e8f0/e2e8f0?text=Preview';
        }
    }

    // Card Background URL handling with null checks
    if (cardBgUrlInput && cardBgPreview) {
        if (data.cardBgUrl || data.card_bg_url) {
            const cardBgUrl = data.cardBgUrl || data.card_bg_url;
            cardBgUrlInput.value = cardBgUrl;
            cardBgPreview.src = cardBgUrl;
        } else {
            cardBgUrlInput.value = '';
            cardBgPreview.src = 'https://placehold.co/100x100/e2e8f0/e2e8f0?text=Preview';
        }
    }
    
    if (googleSheetUrlInput) {
        googleSheetUrlInput.value = data.googleSheetUrl || data.google_sheet_url || '';
    }
    
    // Default page functionality removed
}

// --- SETTINGS PAGE - NEW FUNCTION FOR DEFAULT PAGE DROPDOWN ---
const defaultPageOptions = [
    { id: 'home', name: 'ទំព័រដើម' },
    { id: 'books', name: 'គ្រប់គ្រងសៀវភៅ' },
    { id: 'loans', name: 'ខ្ចី-សង បុគ្គល' },
    { id: 'class-loans', name: 'ខ្ចី-សង តាមថ្នាក់' },
    { id: 'reading-log', name: 'ចូលអាន' },
    { id: 'locations', name: 'ទីតាំងរក្សាទុក' },
    { id: 'students', name: 'ទិន្នន័យសិស្ស' },
    { id: 'student-cards', name: 'កាតសិស្ស' },
    { id: 'settings', name: 'ការកំណត់' }
];

const populateDefaultPageSelect = (currentDefaultPage) => {
    const defaultPageSelect = document.getElementById('default-page-select');
    if (!defaultPageSelect) return;
    
    defaultPageSelect.innerHTML = '<option value="">-- សូមជ្រើសរើស --</option>';
    
    defaultPageOptions.forEach(page => {
        const option = document.createElement('option');
        option.value = page.id;
        option.textContent = page.name;
        defaultPageSelect.appendChild(option);
    });
    
    if (currentDefaultPage) {
        defaultPageSelect.value = currentDefaultPage;
    }
    
    // Add the listener once
    if (!defaultPageSelect.hasAttribute('data-listener')) {
        defaultPageSelect.addEventListener('change', saveDefaultPage);
        defaultPageSelect.setAttribute('data-listener', 'true');
    }
};

const saveDefaultPage = async () => {
    if (!currentUserId) return;
    const defaultPageSelect = document.getElementById('default-page-select');
    const pageId = defaultPageSelect.value;
    
    try {
        const updateData = pageId ? { default_page: pageId } : { default_page: null };
        
        const { error } = await supabase
            .from('settings')
            .upsert({
                user_id: currentUserId,
                ...updateData
            });
        
        if (!error) {
            if (pageId) {
                settingsData.default_page = pageId;
            } else {
                delete settingsData.default_page;
            }
            try {
                if (typeof window.showToast === 'function') {
                    window.showToast('បានរក្សាទុកទំព័រលំនាំដើម ដោយជោគជ័យ!', 'bg-green-600');
                }
            } catch (_) { /* noop */ }
        } else {
            console.error("Error saving default page: ", error);
            alert('ការរក្សាទុកបានបរាជ័យ។');
        }
    } catch (e) {
        console.error("Error saving default page: ", e);
        alert('ការរក្សាទុកបានបរាជ័យ។');
    }
};

// Google Sheets import removed: URL save/fetch handlers deleted; Excel/CSV upload is the only import path now.

// Local CSV Upload Fallback (no CORS)
(() => {
    const uploadBtn = document.getElementById('upload-csv-btn');
    const fileInput = document.getElementById('csv-file-input');
    if (!uploadBtn || !fileInput) return;

    uploadBtn.addEventListener('click', () => fileInput.click());

    fileInput.addEventListener('change', () => {
        const file = fileInput.files && fileInput.files[0];
        if (!file) return;

        const lowerName = (file.name || '').toLowerCase();
        const isExcel = lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls') || /sheet/.test(file.type);

        const handleParsedAndSync = async (parsed) => {
            if (!parsed.length) {
                const m = 'មិនអាចញែកទិន្នន័យពី CSV/Excel បានទេ ឬឯកសារវិកល។';
                if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
                return;
            }
            const normalized = normalizeStudents(parsed);
            if (!normalized.length) {
                const m = 'រកមិនឃើញជួរឈរដែលត្រូវការសម្រាប់សិស្សនៅក្នុង CSV/Excel ទេ។';
                if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
                return;
            }
            await syncStudentsToSupabase(normalized);
            if (typeof window.showToast === 'function') {
                window.showToast(`នាំចូលសិស្ស ${normalized.length} នាក់ដោយជោគជ័យពី CSV/Excel!`, 'bg-green-600');
                setTimeout(() => window.location.reload(), 1200);
            } else {
                alert(`នាំចូលសិស្ស ${normalized.length} នាក់ដោយជោគជ័យពី CSV/Excel!`);
                window.location.reload();
            }
        };

        loadingOverlay.classList.remove('hidden');
        if (isExcel) {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    if (typeof XLSX === 'undefined') throw new Error('XLSX library not loaded');
                    const data = new Uint8Array(e.target?.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    // Convert to CSV then reuse the same CSV parser
                    const csvText = XLSX.utils.sheet_to_csv(worksheet);
                    const parsed = parseCSV(csvText);
                    await handleParsedAndSync(parsed);
                } catch (err) {
                    console.error('Excel upload failed:', err);
                    const m = 'ការនាំចូល Excel បានបរាជ័យ: ' + (err?.message || 'Unknown error');
                    if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
                } finally {
                    loadingOverlay.classList.add('hidden');
                    fileInput.value = '';
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const text = e.target?.result || '';
                    const parsed = parseCSV(String(text));
                    await handleParsedAndSync(parsed);
                } catch (err) {
                    console.error('CSV upload failed:', err);
                    const m = 'ការនាំចូល CSV បានបរាជ័យ: ' + (err?.message || 'Unknown error');
                    if (typeof window.showToast === 'function') window.showToast(m, 'bg-red-600'); else alert(m);
                } finally {
                    loadingOverlay.classList.add('hidden');
                    fileInput.value = '';
                }
            };
            reader.readAsText(file, 'utf-8');
        }
    });
})();

function parseCSV(text) {
    const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
    if (lines.length < 2) return [];
    
    // Improved CSV parsing to handle quoted fields with commas
    function parseCsvLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            const nextChar = line[i + 1];
            
            if (char === '"') {
                if (inQuotes && nextChar === '"') {
                    // Escaped quote
                    current += '"';
                    i++; // Skip next quote
                } else {
                    // Toggle quote state
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                // Field separator
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        // Add the last field
        result.push(current.trim());
        return result;
    }
    
    const headers = parseCsvLine(lines[0]).map(h => h.replace(/^"|"$/g, '').trim());
    console.log('CSV Headers:', headers);
    
    const data = lines.slice(1).map((line, index) => {
        const values = parseCsvLine(line);
        const obj = {};
        headers.forEach((header, headerIndex) => {
            let value = values[headerIndex] || '';
            // Remove surrounding quotes if present
            value = value.replace(/^"|"$/g, '').trim();
            obj[header] = value;
        });
        return obj;
    }).filter(obj => {
        // Filter out empty rows
        return Object.values(obj).some(value => value.trim() !== '');
    });
    
    console.log('Parsed CSV data sample:', data.slice(0, 3));
    return data;
}

// Normalize arbitrary CSV columns into expected student schema with Khmer keys
function normalizeStudents(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return [];
    const sample = rows[0] || {};
    const keys = Object.keys(sample);

    // Helper: find a key by list of aliases (case-insensitive, trimmed, includes support)
    const findKey = (aliases) => {
        const lowered = keys.map(k => ({ k, lk: k.toLowerCase().trim() }));
        for (const alias of aliases) {
            const a = alias.toLowerCase().trim();
            // exact match first
            const exact = lowered.find(x => x.lk === a);
            if (exact) return exact.k;
        }
        // includes match fallback
        for (const alias of aliases) {
            const a = alias.toLowerCase().trim();
            const incl = lowered.find(x => x.lk.includes(a));
            if (incl) return incl.k;
        }
        return null;
    };

    const keySerial = findKey(['ល.រ', 'លរ', 'no', 'serial', 'serial number', 'stt']);
    const keyId = findKey(['អត្តលេខ', 'student id', 'student_id', 'code', 'id', 'student code']);
    const keyLast = findKey(['នាមត្រកូល', 'last name', 'lastname', 'surname']);
    const keyFirst = findKey(['នាមខ្លួន', 'first name', 'firstname', 'given name', 'name']);
    const keyGender = findKey(['ភេទ', 'gender', 'sex']);
    const keyClass = findKey(['ថ្នាក់', 'class', 'grade']);
    const keyDob = findKey(['ថ្ងៃខែឆ្នាំកំណើត', 'dob', 'date of birth', 'birthdate', 'birth date']);
    const keyPhoto = findKey(['រូបថត url', 'រូបថត', 'photo url', 'photo', 'image', 'image url']);

    // Require minimally ID and names to proceed
    const minimalOK = !!(keyId && (keyLast || keyFirst));
    if (!minimalOK) return [];

    const toNumber = (v) => {
        if (v == null) return null;
        const s = String(v).trim();
        if (!s) return null;
        const n = Number(s);
        return Number.isNaN(n) ? s : n;
    };

    const norm = rows.map(r => {
        const obj = {
            'ល.រ': keySerial ? toNumber(r[keySerial]) : null,
            'អត្តលេខ': keyId ? (r[keyId] ?? '').toString().trim() : null,
            'នាមត្រកូល': keyLast ? (r[keyLast] ?? '').toString().trim() : '',
            'នាមខ្លួន': keyFirst ? (r[keyFirst] ?? '').toString().trim() : '',
            'ភេទ': keyGender ? (r[keyGender] ?? '').toString().trim() : '',
            'ថ្នាក់': keyClass ? (r[keyClass] ?? '').toString().trim() : '',
            'ថ្ងៃខែឆ្នាំកំណើត': keyDob ? (r[keyDob] ?? '').toString().trim() : '',
            'រូបថត URL': keyPhoto ? (r[keyPhoto] ?? '').toString().trim() : ''
        };
        // If last/first combined only one exists: attempt simple split by space
        if (!keyLast && keyFirst) {
            const parts = (obj['នាមខ្លួន'] || '').split(/\s+/);
            obj['នាមត្រកូល'] = parts.shift() || '';
            obj['នាមខ្លួន'] = parts.join(' ');
        } else if (keyLast && !keyFirst) {
            const parts = (obj['នាមត្រកូល'] || '').split(/\s+/);
            obj['នាមត្រកូល'] = parts.shift() || '';
            obj['នាមខ្លួន'] = parts.join(' ');
        }
        return obj;
    }).filter(o => Object.values(o).some(v => (v ?? '').toString().trim() !== ''));

    return norm;
}

async function syncStudentsToSupabase(studentData) {
    if (!currentUserId) return;
    
    // Delete existing students for this user
    await supabase
        .from('students')
        .delete()
        .eq('user_id', currentUserId);
    
    // Add user_id to each student record
    const studentsWithUserId = studentData.map(student => ({
        ...student,
        user_id: currentUserId
    }));
    
    // Insert new students
    const { error } = await supabase
        .from('students')
        .insert(studentsWithUserId);
    
    if (error) {
        console.error('Error syncing students:', error);
        throw error;
    }
}

// Normalize arbitrary CSV/Excel columns into expected book schema
function normalizeBooks(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return [];
    const sample = rows[0] || {};
    const keys = Object.keys(sample);

    const findKey = (aliases) => {
        const lowered = keys.map(k => ({ k, lk: k.toLowerCase().trim() }));
        for (const alias of aliases) {
            const a = alias.toLowerCase().trim();
            const exact = lowered.find(x => x.lk === a);
            if (exact) return exact.k;
        }
        for (const alias of aliases) {
            const a = alias.toLowerCase().trim();
            const incl = lowered.find(x => x.lk.includes(a));
            if (incl) return incl.k;
        }
        return null;
    };

    const keyTitle = findKey(['ចំណងជើង', 'title', 'book title', 'name']);
    const keyAuthor = findKey(['អ្នកនិពន្ធ/ថ្នាក់', 'អ្នកនិពន្ធ', 'author', 'class', 'grade']);
    const keyIsbn = findKey(['isbn', 'លេខសម្គាល់', 'លេខសម្គាល់សៀវភៅ']);
    const keyQty = findKey(['ចំនួនសរុប', 'quantity', 'qty', 'total']);
    const keyLocation = findKey(['ទីតាំង', 'location', 'shelf', 'shelve', 'shelving']);
    const keySource = findKey(['ប្រភព', 'source']);

    const toNumber = (v) => {
        if (v == null) return 0;
        let s = String(v).trim();
        if (!s) return 0;
        s = convertKhmerToEnglishNumbers(s);
        const n = Number(s);
        return Number.isNaN(n) ? 0 : n;
    };

    const resolveLocationId = (val) => {
        if (!val) return null;
        const s = String(val).trim();
        // If numeric id provided
        if (/^\d+$/.test(s)) {
            const idNum = Number(s);
            return locations.find(l => Number(l.id) === idNum)?.id || null;
        }
        // Match by name (case-insensitive)
        const lower = s.toLowerCase();
        const found = locations.find(l => (l.name || '').toLowerCase().trim() === lower);
        return found ? found.id : null;
    };

    const norm = rows.map(r => {
        const title = keyTitle ? (r[keyTitle] ?? '').toString().trim() : '';
        const author = keyAuthor ? (r[keyAuthor] ?? '').toString().trim() : '';
        const isbn = keyIsbn ? (r[keyIsbn] ?? '').toString().trim() : '';
        const quantity = keyQty ? toNumber(r[keyQty]) : 0;
        const locVal = keyLocation ? (r[keyLocation] ?? '').toString().trim() : '';
        const source = keySource ? (r[keySource] ?? '').toString().trim() : '';
        const location_id = resolveLocationId(locVal);
        return { title, author, isbn, quantity, location_id, source };
    }).filter(o => (o.title || o.author || o.isbn));

    if (!norm.length) return [];
    return norm;
}

async function syncBooksToSupabase(bookData) {
    if (!currentUserId) return;
    // Delete existing books for this user
    await supabase
        .from('books')
        .delete()
        .eq('user_id', currentUserId);

    // Add user_id to each record
    const booksWithUserId = bookData.map(b => ({
        ...b,
        user_id: currentUserId
    }));

    const { error } = await supabase
        .from('books')
        .insert(booksWithUserId);
    if (error) {
        console.error('Error syncing books:', error);
        throw error;
    }
}

// --- BOOK MANAGEMENT ---
window.openBookModal = (id = null) => {
    const bookForm = document.getElementById('book-form');
    bookForm.reset(); document.getElementById('book-id').value = '';
    const locationSelect = document.getElementById('book-location-id');
    locationSelect.innerHTML = '<option value="">[គ្មាន]</option>';
    locations.forEach(loc => { const option = document.createElement('option'); option.value = loc.id; option.textContent = loc.name; locationSelect.appendChild(option); });
    if (id) {
        const book = books.find(b => b.id === id);
        if (book) {
            document.getElementById('book-modal-title').textContent = 'កែសម្រួលព័ត៌មានសៀវភៅ';
            document.getElementById('book-id').value = book.id;
            document.getElementById('title').value = book.title;
            document.getElementById('author').value = book.author || '';
            document.getElementById('isbn').value = book.isbn || '';
            document.getElementById('quantity').value = book.quantity || 0;
            document.getElementById('book-location-id').value = book.location_id || '';
            document.getElementById('source').value = book.source || '';
            document.getElementById('book-url').value = book.book_url || '';
        }
    } else { document.getElementById('book-modal-title').textContent = 'បន្ថែមសៀវភៅថ្មី'; }
    document.getElementById('book-modal').classList.remove('hidden');
    // Auto-focus ISBN when opening the form (both add and edit)
    setTimeout(() => {
        const isbnInput = document.getElementById('isbn');
        if (isbnInput) isbnInput.focus();
    }, 50);
};
window.closeBookModal = () => {
    const modal = document.getElementById('book-modal');
    if (modal) modal.classList.add('hidden');
    // Return focus to Books search box
    setTimeout(() => {
        const input = document.getElementById('search-books');
        if (input && !document.getElementById('book-modal')?.classList.contains('hidden')) return; // still open
        if (input) { input.focus(); input.select(); }
    }, 50);
};
window.editBook = (id) => openBookModal(id);
// Book Delete Confirmation Modal controls
window.openBookDeleteModal = (id) => {
    const modal = document.getElementById('book-delete-modal');
    const messageEl = document.getElementById('book-delete-message');
    const confirmBtn = document.getElementById('book-delete-confirm-btn');
    const bookToDelete = books.find(b => String(b.id) === String(id));
    const title = (bookToDelete?.title || '').trim();
    if (messageEl) {
        messageEl.textContent = title ? `តើអ្នកពិតជាចង់លុបសៀវភៅ "${title}" មែនទេ?` : 'តើអ្នកពិតជាចង់លុបសៀវភៅនេះមែនទេ?';
    }
    if (confirmBtn) {
        confirmBtn.onclick = async () => {
            await window.performDeleteBook(id);
            window.closeBookDeleteModal();
        };
    }
    if (modal) modal.classList.remove('hidden');
};

window.closeBookDeleteModal = () => {
    const modal = document.getElementById('book-delete-modal');
    const confirmBtn = document.getElementById('book-delete-confirm-btn');
    if (confirmBtn) confirmBtn.onclick = null;
    if (modal) modal.classList.add('hidden');
};

window.performDeleteBook = async (id) => {
    if (!currentUserId) return;
    const bookToDelete = books.find(b => String(b.id) === String(id));
    const isLoaned = loans.some(loan => loan.book_id === id && loan.status === 'ខ្ចី');
    if (isLoaned) { alert('មិនអាចលុបសៀវភៅនេះបានទេ ព្រោះកំពុងមានគេខ្ចី។'); return; }
    try {
        const { error } = await supabase
            .from('books')
            .delete()
            .eq('id', id)
            .eq('user_id', currentUserId);
        if (error) {
            console.error("Error deleting document: ", error);
        } else {
            await loadBooks(currentUserId);
            renderAll();
            try {
                if (typeof window.showToast === 'function') {
                    const t = (bookToDelete?.title || '').trim();
                    window.showToast(`បានលុបសៀវភៅ ${t} ដោយជោគជ័យ!`, 'bg-green-600');
                }
            } catch (_) { /* noop */ }
        }
    } catch (e) { console.error("Error deleting document: ", e); }
};
window.deleteBook = (id) => {
    if (!currentUserId) return;
    window.openBookDeleteModal(id);
};

// Add real-time Khmer to English number conversion for book ISBN field
document.getElementById('isbn').addEventListener('input', (e) => {
    const convertedValue = convertKhmerToEnglishNumbers(e.target.value);
    if (convertedValue !== e.target.value) {
        e.target.value = convertedValue;
    }
});

document.getElementById('book-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentUserId) return;
    
    const id = document.getElementById('book-id').value;
    const isbnValue = convertKhmerToEnglishNumbers(document.getElementById('isbn').value.trim());
    
    const bookData = {
        title: document.getElementById('title').value,
        author: document.getElementById('author').value,
        isbn: isbnValue,
        quantity: parseInt(document.getElementById('quantity').value, 10) || 0,
        location_id: document.getElementById('book-location-id').value || null,
        source: document.getElementById('source').value,
        book_url: document.getElementById('book-url').value,
    };

    if (bookData.isbn) {
        try {
            const { data: existingBooks, error } = await supabase
                .from('books')
                .select('id')
                .eq('user_id', currentUserId)
                .eq('isbn', bookData.isbn);
            
            if (error) {
                console.error("Error checking for duplicate ISBN:", error);
                alert("មានបញ្ហាក្នុងការត្រួតពិនិត្យ ISBN។ សូមព្យាយាមម្តងទៀត។");
                return;
            }
            
            let isDuplicate = false;
            if (existingBooks && existingBooks.length > 0) {
                if (id) { 
                    if (existingBooks[0].id !== id) {
                        isDuplicate = true;
                    }
                } else { 
                    isDuplicate = true;
                }
            }

            if (isDuplicate) {
                alert('លេខ ISBN "' + bookData.isbn + '" នេះមានក្នុងប្រព័ន្ធរួចហើយ។ សូមប្រើលេខផ្សេង។');
                return; 
            }
        } catch (err) {
            console.error("Error checking for duplicate ISBN:", err);
            alert("មានបញ្ហាក្នុងការត្រួតពិនិត្យ ISBN។ សូមព្យាយាមម្តងទៀត។");
            return;
        }
    }
    
    const bookDataWithUserId = {
        ...bookData,
        user_id: currentUserId
    };
    
    try {
        if (id) {
            const { error } = await supabase
                .from('books')
                .update(bookDataWithUserId)
                .eq('id', id)
                .eq('user_id', currentUserId);
            if (error) throw error;
        } else {
            const { error } = await supabase
                .from('books')
                .insert([bookDataWithUserId]);
            if (error) throw error;
        }
        // Reload list after save
        try {
            await loadBooks(currentUserId);
            renderAll();
        } catch (loadError) {
            console.error("Error loading books after save:", loadError);
            // Don't show error for load failure, book was saved successfully
        }
        // Show mini overlay (toast) for 1s: សៀវភៅ [title] + [author] + [quantity] បានរក្សាទុកជោគជ័យ
        const t = (bookData.title || '').trim() || 'មិនមានចំណងជើង';
        const a = (bookData.author || '').trim() || 'មិនមានអ្នកនិពន្ធ';
        const q = bookData.quantity ?? 0;
        if (typeof window.showToast === 'function') {
            window.showToast(`សៀវភៅ ${t} + ${a} + ${q} បានរក្សាទុកជោគជ័យ`, 'bg-green-600');
        }
        // After save: if editing, close modal; if adding, reset form and focus ISBN for quick entry
        if (id) {
            closeBookModal();
        } else {
            const bookForm = document.getElementById('book-form');
            bookForm.reset();
            document.getElementById('book-id').value = '';
            const locationSelect = document.getElementById('book-location-id');
            if (locationSelect) {
                locationSelect.innerHTML = '<option value="">[គ្មាន]</option>';
                locations.forEach(loc => { const option = document.createElement('option'); option.value = loc.id; option.textContent = loc.name; locationSelect.appendChild(option); });
            }
            const isbnInput = document.getElementById('isbn');
            if (isbnInput) isbnInput.focus();
        }
        //alert("រក្សាទុកសៀវភៅបានជោគជ័យ។");
    } catch (e) {
        console.error("Error adding/updating document: ", e);
        alert("ការរក្សាទុកទិន្នន័យបានបរាជ័យ។");
    }
});


// --- LOAN MANAGEMENT (INDIVIDUAL) ---
window.clearLoanForm = () => {
    console.log('clearLoanForm called');
    
    // Clear form elements one by one with error checking
    const elements = [
        'loan-form',
        'loan-book-id', 
        'loan-borrower-gender',
        'loan-book-error',
        'loan-student-error', 
        'loan-isbn-input',
        'loan-student-id-input',
        'loan-book-title-display',
        'loan-borrower-name-display'
    ];
    
    elements.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            if (id === 'loan-form') {
                element.reset();
            } else if (element.tagName === 'INPUT' || element.tagName === 'TEXTAREA' || element.tagName === 'SELECT') {
                element.value = '';
            } else {
                element.textContent = '';
            }
            console.log(`Cleared ${id}`);
        } else {
            console.log(`Element ${id} not found`);
        }
    });
    
    // Clear selected books list
    clearSelectedLoanBooks();
    
    setTimeout(() => {
        const isbnInput = document.getElementById('loan-isbn-input');
        if (isbnInput) {
            isbnInput.focus();
            console.log('Focus set to loan-isbn-input');
        }
    }, 100);
};

document.getElementById('loan-isbn-input').addEventListener('input', (e) => {
    // Convert Khmer numbers to English numbers
    const convertedValue = convertKhmerToEnglishNumbers(e.target.value);
    if (convertedValue !== e.target.value) {
        e.target.value = convertedValue;
    }
    
    const isbn = convertedValue.trim();
    const loanBookError = document.getElementById('loan-book-error');
    loanBookError.textContent = '';
    
    if (isbn) {
        const foundBook = books.find(book => book.isbn === isbn);
        if (foundBook) {
            const remaining = (foundBook.quantity || 0) - loans.filter(loan => loan.book_id === foundBook.id && loan.status === 'ខ្ចី').length;
            if (remaining <= 0) {
                loanBookError.textContent = 'សៀវភៅនេះអស់ស្តុកហើយ។';
                return;
            }
            
            const success = addSelectedLoanBook(foundBook.id);
            if (success) {
                // Clear the input for next scan
                document.getElementById('loan-isbn-input').value = '';
                loanBookError.textContent = '';
                // Always focus back to scan textbox for fast consecutive scans
                setTimeout(() => {
                    const isbnInput = document.getElementById('loan-isbn-input');
                    if (isbnInput) isbnInput.focus();
                }, 50);
            }
        } else { 
            loanBookError.textContent = 'សៀវភៅនេះអស់ពីស្តុកហើយ'; 
        }
    } else { 
        loanBookError.textContent = 'រកមិនឃើញ ISBN នេះទេ'; 
    }
});

document.getElementById('loan-student-id-input').addEventListener('input', () => {
    const studentId = document.getElementById('loan-student-id-input').value.trim();
    const loanBorrowerNameDisplay = document.getElementById('loan-borrower-name-display');
    const loanStudentError = document.getElementById('loan-student-error');
    const borrowerGenderInput = document.getElementById('loan-borrower-gender');
    loanBorrowerNameDisplay.value = ''; loanStudentError.textContent = ''; borrowerGenderInput.value = '';
    if (!studentId) return;
    const studentIdKey = students.length > 0 ? Object.keys(students[0]).find(k => k.toLowerCase().includes('អត្តលេខ')) : null;
    if (!studentIdKey) { loanStudentError.textContent = 'រកមិនឃើញជួរឈរ "អត្តលេខ" ទេ។'; return; }
    const foundStudent = students.find(s => s[studentIdKey] && s[studentIdKey].trim() === studentId);
    if (foundStudent) {
        const lastNameKey = Object.keys(foundStudent).find(k => k.includes('នាមត្រកូល'));
        const firstNameKey = Object.keys(foundStudent).find(k => k.includes('នាមខ្លួន'));
        const classKey = Object.keys(foundStudent).find(k => k.includes('ថ្នាក់'));
        const genderKey = Object.keys(foundStudent).find(k => k.includes('ភេទ'));
        const studentFullName = `${foundStudent[lastNameKey] || ''} ${foundStudent[firstNameKey] || ''}`.trim();
        const studentClass = foundStudent[classKey] || '';
        loanBorrowerNameDisplay.value = studentClass ? `${studentFullName} - ${studentClass}` : studentFullName;
        borrowerGenderInput.value = foundStudent[genderKey] || '';
    } else { 
        loanStudentError.textContent = 'រកមិនឃើញអត្តលេខសិស្សនេះទេ។'; 
    }
});

document.getElementById('loan-form').addEventListener('submit', async (e) => {
    e.preventDefault(); 
    if (!currentUserId) return;
    const borrower = document.getElementById('loan-borrower-name-display').value;
    
    // Check if books are selected and borrower name is provided
    if (selectedLoanBooks.length === 0 || !borrower) { 
        alert('សូមបំពេញព័ត៌មានឲ្យបានត្រឹមត្រូវ! (ជ្រើសរើសសៀវភៅ និងបញ្ចូលឈ្មោះអ្នកខ្ចី)'); 
        return; 
    }
    try {
        const borrowerGender = document.getElementById('loan-borrower-gender').value;
        const loanDate = new Date().toISOString().split('T')[0];
        
        // Create loans for all selected books
        const loansToInsert = selectedLoanBooks.map(book => ({
            book_id: book.id,
            borrower: borrower,
            borrower_gender: borrowerGender,
            loan_date: loanDate,
            return_date: null,
            status: 'ខ្ចី',
            user_id: currentUserId
        }));
        
        const { error } = await supabase
            .from('loans')
            .insert(loansToInsert);
        if (error) throw error;
        
        await loadLoans(currentUserId);
        renderAll();
        // Show centered overlay message for 3 seconds about successful loan save
        try {
            const bookTitles = selectedLoanBooks.map(b => (b.title || '').trim()).filter(Boolean).join(',');
            if (typeof window.showToast === 'function') {
                window.showToast(`បានរក្សាការខ្ចីសៀវភៅ របស់ សិស្សឈ្មោះ ${borrower}+${bookTitles} ដោយជោគជ័យ!`, 'bg-green-600');
            }
        } catch (_) { /* noop */ }
        clearLoanForm();
    } 
    catch(e) { console.error("Error adding loan: ", e); }
});

window.returnBook = async (id) => {
    if (!currentUserId) return;
    if (confirm('តើអ្នកពិតជាចង់សម្គាល់ថាសៀវភៅនេះត្រូវបានសងវិញមែនទេ?')) {
        try { 
            const { error } = await supabase
                .from('loans')
                .update({ 
                    status: 'សង', 
                    return_date: new Date().toISOString().split('T')[0] 
                })
                .eq('id', id)
                .eq('user_id', currentUserId);
            if (error) console.error("Error updating loan: ", error);
            else {
                await loadLoans(currentUserId);
                renderAll();
            }
        } catch (e) { console.error("Error updating loan: ", e); }
    }
};
// Loan Delete Confirmation Modal controls
window.openLoanDeleteModal = (id) => {
    const modal = document.getElementById('loan-delete-modal');
    const messageEl = document.getElementById('loan-delete-message');
    const confirmBtn = document.getElementById('loan-delete-confirm-btn');
    const loanToDelete = loans.find(l => String(l.id) === String(id));
    const borrower = (loanToDelete?.borrower || '').trim();
    const bookForLoan = loanToDelete ? books.find(b => String(b.id) === String(loanToDelete.book_id)) : null;
    const bookTitle = (bookForLoan?.title || '').trim();
    if (messageEl) {
        const detail = [borrower, bookTitle].filter(Boolean).join(' • ');
        messageEl.textContent = detail ? `តើអ្នកពិតជាចង់លុបកំណត់ត្រា៖ ${detail} មែនទេ?` : 'តើអ្នកពិតជាចង់លុបកំណត់ត្រានេះមែនទេ?';
    }
    if (confirmBtn) {
        confirmBtn.onclick = async () => {
            await window.performDeleteLoan(id);
            window.closeLoanDeleteModal();
        };
    }
    if (modal) modal.classList.remove('hidden');
};

window.closeLoanDeleteModal = () => {
    const modal = document.getElementById('loan-delete-modal');
    const confirmBtn = document.getElementById('loan-delete-confirm-btn');
    if (confirmBtn) confirmBtn.onclick = null;
    if (modal) modal.classList.add('hidden');
};

window.performDeleteLoan = async (id) => {
    if (!currentUserId) return;
    const loanToDelete = loans.find(l => String(l.id) === String(id));
    const bookForLoan = loanToDelete ? books.find(b => String(b.id) === String(loanToDelete.book_id)) : null;
    try {
        const { error } = await supabase
            .from('loans')
            .delete()
            .eq('id', id)
            .eq('user_id', currentUserId);
        if (error) {
            console.error('Error deleting loan: ', error);
        } else {
            await loadLoans(currentUserId);
            renderAll();
            try {
                if (typeof window.showToast === 'function') {
                    const borrower = (loanToDelete?.borrower || '').trim();
                    const bookTitle = (bookForLoan?.title || '').trim();
                    window.showToast(`បានលុបការខ្ចី ${borrower}+${bookTitle} ដោយជោគជ័យ!`, 'bg-green-600');
                }
            } catch (_) { /* noop */ }
        }
    } catch (e) { console.error('Error deleting loan: ', e); }
};

window.deleteLoan = (id) => {
    if (!currentUserId) return;
    window.openLoanDeleteModal(id);
};

// --- LOAN MANAGEMENT (CLASS) ---
const classLoanClassSelect = document.getElementById('class-loan-class-select');
const classInfoText = document.getElementById('class-info-text');

// START: NEW function to update loan count based on checkboxes
window.updateClassLoanCount = () => {
    const checkedStudents = document.querySelectorAll('#class-loan-student-list-container input[type="checkbox"]:checked');
    document.getElementById('class-loan-quantity').value = checkedStudents.length;
};
// END: NEW function

function populateClassLoanForm() {
    document.getElementById('class-loan-form').reset();
    document.getElementById('class-loan-book-id').value = '';
    document.getElementById('class-loan-book-error').textContent = '';
    document.getElementById('class-info-text').textContent = '';
    document.getElementById('class-loan-scanned-books-list').innerHTML = '';
    currentClassLoanScannedBooks = [];
    const studentListContainer = document.getElementById('class-loan-student-list-container');
    studentListContainer.innerHTML = '';
    studentListContainer.classList.add('hidden');
    window.updateClassLoanCount(); // Reset count to 0

    const classKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('ថ្នាក់')) : null;
    classLoanClassSelect.innerHTML = '<option value="">-- សូមជ្រើសរើសថ្នាក់ --</option>';
    if (classKey) {
        const uniqueClasses = [...new Set(students.map(s => s[classKey]))].filter(Boolean).sort();
        uniqueClasses.forEach(className => {
            const option = document.createElement('option');
            option.value = className; option.textContent = className;
            classLoanClassSelect.appendChild(option);
        });
    }
}

// Disable enter key submission for class loan form
document.getElementById('class-loan-form').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
        e.preventDefault();
    }
});

document.getElementById('class-loan-isbn-input').addEventListener('input', (e) => {
    // Convert Khmer numbers to English numbers
    const convertedValue = convertKhmerToEnglishNumbers(e.target.value);
    if (convertedValue !== e.target.value) {
        e.target.value = convertedValue;
    }
    
    const isbn = convertedValue.trim();
    const bookIdInput = document.getElementById('class-loan-book-id');
    const classLoanBookTitleDisplay = document.getElementById('class-loan-book-title-display');
    const classLoanBookError = document.getElementById('class-loan-book-error');
    classLoanBookError.textContent = '';
    
    if (isbn) {
        const foundBook = books.find(book => book.isbn === isbn);
        if (foundBook) {
            const remaining = (foundBook.quantity || 0) - loans.filter(loan => loan.book_id === foundBook.id && loan.status === 'ខ្ចី').length;
            if (remaining <= 0) {
                classLoanBookError.textContent = 'សៀវភៅនេះអស់ស្តុកហើយ។';
                return;
            }
            
            currentClassLoanScannedBooks.push({ id: foundBook.id, title: foundBook.title, remaining });
            renderClassLoanScannedBooks();
            
            // Clear the input for next scan
            document.getElementById('class-loan-isbn-input').value = '';
            classLoanBookTitleDisplay.value = `${foundBook.title} (បានបន្ថែម)`;
            // Focus back to scan textbox for rapid entry
            setTimeout(() => {
                const input = document.getElementById('class-loan-isbn-input');
                if (input) input.focus();
            }, 50);
        } else {
            classLoanBookError.textContent = 'សៀវភៅនេះបានស្កេនរួចហើយ';
        }
    } else { classLoanBookError.textContent = 'រកមិនឃើញ ISBN នេះទេ'; }
});

// START: MODIFIED event listener for class selection
classLoanClassSelect.addEventListener('change', () => {
    const className = classLoanClassSelect.value;
    const studentListContainer = document.getElementById('class-loan-student-list-container');
    
    studentListContainer.innerHTML = ''; // Clear previous list
    studentListContainer.classList.add('hidden'); // Hide by default

    if (!className) {
        classInfoText.textContent = '';
        window.updateClassLoanCount(); // Update count to 0
        return;
    }

    const classKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('ថ្នាក់')) : null;
    if (!classKey) {
        classInfoText.textContent = 'Error: Cannot find class data.';
        return;
    }
    const studentsInClass = students.filter(s => s[classKey] === className);

    classInfoText.textContent = `ថ្នាក់នេះមានសិស្ស ${studentsInClass.length} នាក់`;

    if (studentsInClass.length > 0) {
        studentListContainer.classList.remove('hidden');
        const lastNameKey = Object.keys(studentsInClass[0]).find(k => k.includes('នាមត្រកូល'));
        const firstNameKey = Object.keys(studentsInClass[0]).find(k => k.includes('នាមខ្លួន'));
        const genderKey = Object.keys(studentsInClass[0]).find(k => k.includes('ភេទ'));

        const listGrid = document.createElement('div');
        listGrid.className = 'grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2';

        studentsInClass.forEach(student => {
            const studentFullName = `${student[lastNameKey] || ''} ${student[firstNameKey] || ''}`.trim();
            const studentGender = student[genderKey] || '';

            const label = document.createElement('label');
            label.className = 'flex items-center space-x-2 p-1 rounded hover:bg-gray-200 cursor-pointer';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.className = 'form-checkbox h-4 w-4 text-cyan-600 border-gray-300 rounded focus:ring-cyan-500';
            checkbox.checked = true; // Default to checked
            checkbox.dataset.fullName = studentFullName;
            checkbox.dataset.gender = studentGender;
            
            checkbox.addEventListener('change', window.updateClassLoanCount);

            const span = document.createElement('span');
            span.className = 'text-gray-800 text-sm';
            span.textContent = studentFullName;

            label.appendChild(checkbox);
            label.appendChild(span);
            listGrid.appendChild(label);
        });
        studentListContainer.appendChild(listGrid);
    }
    // Initial count update
    window.updateClassLoanCount();
});
// END: MODIFIED event listener

// START: MODIFIED form submission logic
document.getElementById('class-loan-form').addEventListener('submit', async (e) => {
    e.preventDefault(); 
    if (!currentUserId) return;

    const className = classLoanClassSelect.value;
    const checkedStudentCheckboxes = document.querySelectorAll('#class-loan-student-list-container input[type="checkbox"]:checked');
    const quantity = checkedStudentCheckboxes.length;

    if (currentClassLoanScannedBooks.length === 0 || !className || quantity === 0) {
        alert('សូមស្កេនសៀវភៅយ៉ាងហោចណាស់មួយក្បាល, ជ្រើសរើសថ្នាក់, និងសិស្សយ៉ាងហោចណាស់ម្នាក់។');
        return;
    }
    
    // Check availability for all scanned books
    for (const scannedBook of currentClassLoanScannedBooks) {
        const selectedBook = books.find(b => b.id === scannedBook.id);
        if (!selectedBook) {
            alert(`រកមិនឃើញព័ត៌មានសៀវភៅ: ${scannedBook.title}`);
            return;
        }
        const currentlyLoanedCount = loans.filter(l => l.book_id === scannedBook.id && l.status === 'ខ្ចី').length;
        const availableCopies = selectedBook.quantity - currentlyLoanedCount;

        if (quantity > availableCopies) {
            alert(`សៀវភៅមិនគ្រប់គ្រាន់! សៀវភៅ "${scannedBook.title}" អ្នកចង់ខ្ចី ${quantity} ក្បាល, ប៉ុន្តែនៅសល់តែ ${availableCopies} ក្បាលសម្រាប់ខ្ចី។`);
            return;
        }
    }

    loadingOverlay.classList.remove('hidden');
    try {
        const today = new Date().toISOString().split('T')[0];
        
        // Create class loan records for each book
        for (const scannedBook of currentClassLoanScannedBooks) {
            const { data: classLoanData, error: classLoanError } = await supabase
                .from('class_loans')
                .insert([{
                    book_id: scannedBook.id, 
                    class_name: className, 
                    loaned_quantity: quantity, 
                    loan_date: today, 
                    returned_count: 0, 
                    status: 'ខ្ចី',
                    user_id: currentUserId
                }])
                .select();
            
            if (classLoanError) throw classLoanError;
            const classLoanId = classLoanData[0].id;
            
            // Create individual loan records for this book
            const loanRecords = [];
            checkedStudentCheckboxes.forEach(checkbox => {
                const borrowerText = `${checkbox.dataset.fullName} - ${className}`;
                loanRecords.push({ 
                    book_id: scannedBook.id, 
                    borrower: borrowerText, 
                    loan_date: today, 
                    return_date: null, 
                    status: 'ខ្ចី', 
                    class_loan_id: classLoanId, 
                    borrower_gender: checkbox.dataset.gender || '',
                    user_id: currentUserId
                });
            });
            
            const { error: loansError } = await supabase
                .from('loans')
                .insert(loanRecords);
            
            if (loansError) throw loansError;
        }
        await loadLoans(currentUserId);
        await loadClassLoans(currentUserId);
        renderAll();
        // Show centered overlay message for 3 seconds about successful class loan save
        try {
            const classNameText = className || '';
            const bookTitles = currentClassLoanScannedBooks.map(b => (b.title || '').trim()).filter(Boolean).join(',');
            const qty = quantity;
            if (typeof window.showToast === 'function') {
                window.showToast(`បានរក្សាការខ្ចីសៀវភៅ របស់ ថ្នាក់ ${classNameText}+${bookTitles} ចំនួន ${qty} នាក់ ដោយជោគជ័យ!`, 'bg-green-600');
            }
        } catch (_) { /* noop */ }
        populateClassLoanForm(); // Reset the form completely
    } catch (err) { 
        console.error("Error creating class loan: ", err); 
        alert("មានបញ្ហាកើតឡើងពេលបង្កើតកំណត់ត្រា។");
    } finally { 
        loadingOverlay.classList.add('hidden'); 
    }
});
// END: MODIFIED form submission logic

window.openClassReturnModal = (id) => {
    const classReturnForm = document.getElementById('class-return-form');
    classReturnForm.reset();
    const classLoan = classLoans.find(cl => cl.id === id);
    if (!classLoan) return;
    const book = books.find(b => b.id === classLoan.book_id);
    document.getElementById('class-return-loan-id').value = id;
    document.getElementById('class-return-book-title').textContent = book ? book.title : 'N/A';
    document.getElementById('class-return-class-name').textContent = classLoan.class_name;
    document.getElementById('class-return-total-students').textContent = classLoan.loaned_quantity;
    document.getElementById('class-return-already-returned').textContent = classLoan.returned_count || 0;
    const numberInput = document.getElementById('number-to-return');
    const maxReturn = classLoan.loaned_quantity - (classLoan.returned_count || 0);
    numberInput.max = maxReturn;
    numberInput.placeholder = `អតិបរមា ${maxReturn}`;
    document.getElementById('class-return-modal').classList.remove('hidden');
};

window.closeClassReturnModal = () => document.getElementById('class-return-modal').classList.add('hidden');

document.getElementById('class-return-form').addEventListener('submit', async (e) => {
    e.preventDefault(); if (!currentUserId) return;
    const id = document.getElementById('class-return-loan-id').value;
    const numberToReturn = parseInt(document.getElementById('number-to-return').value, 10);
    const classLoan = classLoans.find(cl => cl.id === id);
    if (!classLoan) return;
    const maxReturn = classLoan.loaned_quantity - (classLoan.returned_count || 0);
    if (isNaN(numberToReturn) || numberToReturn <= 0 || numberToReturn > maxReturn) { alert(`សូមបញ្ចូលចំនួនត្រឹមត្រូវចន្លោះពី 1 ដល់ ${maxReturn}។`); return; }
    loadingOverlay.classList.remove('hidden');
    try {
        const today = new Date().toISOString().split('T')[0];
        
        // Get loans to update
        const { data: loansToUpdate, error: queryError } = await supabase
            .from('loans')
            .select('id')
            .eq('class_loan_id', id)
            .eq('status', 'ខ្ចី')
            .eq('user_id', currentUserId)
            .limit(numberToReturn);
        
        if (queryError) throw queryError;
        
        // Update individual loans
        const loanIds = loansToUpdate.map(loan => loan.id);
        const { error: updateLoansError } = await supabase
            .from('loans')
            .update({ status: 'សង', return_date: today })
            .in('id', loanIds);
        
        if (updateLoansError) throw updateLoansError;
        
        // Update class loan
        const newReturnedCount = (classLoan.returned_count || 0) + numberToReturn;
        const newStatus = newReturnedCount >= classLoan.loaned_quantity ? 'សងហើយ' : 'សងខ្លះ';
        const { error: updateClassLoanError } = await supabase
            .from('class_loans')
            .update({ returned_count: newReturnedCount, status: newStatus })
            .eq('id', id)
            .eq('user_id', currentUserId);
        
        if (updateClassLoanError) throw updateClassLoanError;
        
        // Reload data to update the UI
        await loadLoans(currentUserId);
        await loadClassLoans(currentUserId);
        renderAll();
        closeClassReturnModal();
    } catch (err) { console.error("Error updating class loan return: ", err); alert("មានបញ្ហាក្នុងការរក្សាទុកការសង។");
    } finally { loadingOverlay.classList.add('hidden'); }
});

// Grouped Class Return Modal Functions
window.openGroupedClassReturnModal = (groupKey) => {
    const [className, loanDate] = groupKey.split('|');
    
    // Find all class loans for this group
    const groupedLoans = classLoans.filter(loan => 
        loan.class_name === className && loan.loan_date === loanDate
    );
    
    if (groupedLoans.length === 0) return;
    
    // Populate modal header
    document.getElementById('grouped-return-class-name').textContent = className;
    document.getElementById('grouped-return-loan-date').textContent = loanDate;
    
    // Build books list with checkboxes and quantity inputs
    const booksList = document.getElementById('grouped-return-books-list');
    booksList.innerHTML = '';
    
    groupedLoans.forEach(loan => {
        const book = books.find(b => b.id === loan.book_id);
        const bookTitle = book ? book.title : 'សៀវភៅត្រូវបានលុប';
        const loanedQty = loan.loaned_quantity || 0;
        const returnedQty = loan.returned_count || 0;
        const maxReturn = loanedQty - returnedQty;
        const displayTitle = `${bookTitle} (${loanedQty}) (${returnedQty})`;
        
        if (maxReturn > 0) {
            const bookItem = document.createElement('div');
            bookItem.className = 'flex items-center justify-between p-3 border rounded-md bg-white';
            bookItem.innerHTML = `
                <div class="flex items-center">
                    <input type="checkbox" id="book-${loan.id}" class="mr-3 h-4 w-4 text-teal-600 focus:ring-teal-500 border-gray-300 rounded">
                    <label for="book-${loan.id}" class="flex-grow">
                        <div class="font-medium text-gray-900">${displayTitle}</div>
                        <div class="text-sm text-gray-500">
                            នៅសល់ដែលអាចសង: ${maxReturn}
                        </div>
                    </label>
                </div>
                <div class="ml-4">
                    <label for="return-qty-${loan.id}" class="block text-xs text-gray-600 mb-1">ចំនួនសង</label>
                    <input type="number" id="return-qty-${loan.id}" data-loan-id="${loan.id}" 
                           class="w-16 px-2 py-1 border border-gray-300 rounded-md text-sm" 
                           min="1" max="${maxReturn}" value="${maxReturn}" disabled>
                </div>
            `;
            booksList.appendChild(bookItem);
            
            // Enable/disable quantity input based on checkbox
            const checkbox = bookItem.querySelector(`#book-${loan.id}`);
            const quantityInput = bookItem.querySelector(`#return-qty-${loan.id}`);
            
            checkbox.addEventListener('change', () => {
                quantityInput.disabled = !checkbox.checked;
                if (checkbox.checked) {
                    // Default to the remaining borrowed count, but allow edits
                    quantityInput.value = String(maxReturn);
                    quantityInput.focus();
                }
            });
        }
    });
    
    document.getElementById('grouped-class-return-modal').classList.remove('hidden');
};

window.closeGroupedClassReturnModal = () => {
    document.getElementById('grouped-class-return-modal').classList.add('hidden');
};

// Handle grouped return form submission
document.getElementById('grouped-class-return-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentUserId) return;
    
    // Collect selected books and quantities
    const selectedReturns = [];
    const checkboxes = document.querySelectorAll('#grouped-return-books-list input[type="checkbox"]:checked');
    
    if (checkboxes.length === 0) {
        alert('សូមជ្រើសរើសសៀវភៅយ៉ាងហោចណាស់មួយក្បាលដើម្បីសង។');
        return;
    }
    
    checkboxes.forEach(checkbox => {
        const loanId = checkbox.id.replace('book-', '');
        const quantityInput = document.getElementById(`return-qty-${loanId}`);
        const quantity = parseInt(quantityInput.value, 10);
        
        if (quantity > 0) {
            selectedReturns.push({ loanId, quantity });
        }
    });
    
    if (selectedReturns.length === 0) {
        alert('សូមបញ្ចូលចំនួនសងសម្រាប់សៀវភៅដែលបានជ្រើសរើស។');
        return;
    }
    
    loadingOverlay.classList.remove('hidden');
    
    try {
        const today = new Date().toISOString().split('T')[0];
        
        // Process each selected return
        for (const returnItem of selectedReturns) {
            const classLoan = classLoans.find(cl => cl.id === returnItem.loanId);
            if (!classLoan) continue;
            
            const numberToReturn = returnItem.quantity;
            const maxReturn = (classLoan.loaned_quantity || 0) - (classLoan.returned_count || 0);
            
            if (numberToReturn > maxReturn) {
                alert(`ចំនួនសងសម្រាប់សៀវភៅ "${books.find(b => b.id === classLoan.book_id)?.title || 'N/A'}" លើសពីចំនួនដែលអាចសងបាន។`);
                continue;
            }
            
            // Get loans to update for this class loan
            const { data: loansToUpdate, error: queryError } = await supabase
                .from('loans')
                .select('id')
                .eq('class_loan_id', returnItem.loanId)
                .eq('status', 'ខ្ចី')
                .eq('user_id', currentUserId)
                .limit(numberToReturn);
            
            if (queryError) throw queryError;
            
            // Update individual loans
            if (loansToUpdate.length > 0) {
                const loanIds = loansToUpdate.map(loan => loan.id);
                const { error: updateLoansError } = await supabase
                    .from('loans')
                    .update({ status: 'សង', return_date: today })
                    .in('id', loanIds);
                
                if (updateLoansError) throw updateLoansError;
            }
            
            // Update class loan
            const newReturnedCount = (classLoan.returned_count || 0) + numberToReturn;
            const newStatus = newReturnedCount >= classLoan.loaned_quantity ? 'សងហើយ' : 'សងខ្លះ';
            
            const { error: updateClassLoanError } = await supabase
                .from('class_loans')
                .update({ returned_count: newReturnedCount, status: newStatus })
                .eq('id', returnItem.loanId)
                .eq('user_id', currentUserId);
            
            if (updateClassLoanError) throw updateClassLoanError;
        }
        
        // Reload data and close modal
        await loadLoans(currentUserId);
        await loadClassLoans(currentUserId);
        renderAll();
        try { if (typeof window.showToast === 'function') { window.showToast('សងសៀវភៅតាមថ្នាក់បានជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
        closeGroupedClassReturnModal();
        
    } catch (err) {
        console.error("Error processing grouped return: ", err);
        alert("មានបញ្ហាក្នុងការរក្សាទុកការសង។");
    } finally {
        loadingOverlay.classList.add('hidden');
    }
});

// Grouped Class Loan Delete Function
// Centered modal handlers for Grouped Class Loan deletion
window.openClassLoanDeleteModal = (groupKey) => {
    const modal = document.getElementById('class-loan-delete-modal');
    const msg = document.getElementById('class-loan-delete-message');
    const listEl = document.getElementById('class-loan-delete-books');
    const btn = document.getElementById('class-loan-delete-confirm-btn');
    if (!modal || !btn) return;
    const [className, loanDate] = groupKey.split('|');
    const groupedLoans = classLoans.filter(loan => loan.class_name === className && loan.loan_date === loanDate);
    const bookTitles = groupedLoans.map(loan => {
        const book = books.find(b => b.id === loan.book_id);
        return book ? book.title : 'សៀវភៅត្រូវបានលុប';
    });
    if (msg) msg.textContent = `តើអ្នកពិតជាចង់លុបការខ្ចីតាមថ្នាក់ «${className}» កាលបរិច្ឆេទ ${loanDate} មែនទេ?`;
    if (listEl) {
        listEl.innerHTML = '';
        bookTitles.forEach(t => {
            const li = document.createElement('li');
            li.textContent = t;
            listEl.appendChild(li);
        });
    }
    btn.onclick = async () => { await window.performClassLoanDelete(groupKey, bookTitles); window.closeClassLoanDeleteModal(); };
    modal.classList.remove('hidden');
};

window.closeClassLoanDeleteModal = () => {
    const modal = document.getElementById('class-loan-delete-modal');
    const btn = document.getElementById('class-loan-delete-confirm-btn');
    if (btn) btn.onclick = null;
    if (modal) modal.classList.add('hidden');
};

window.performClassLoanDelete = async (groupKey, bookTitles) => {
    if (!currentUserId) return;
    const [className, loanDate] = groupKey.split('|');
    const groupedLoans = classLoans.filter(loan => loan.class_name === className && loan.loan_date === loanDate);
    if (groupedLoans.length === 0) return;
    loadingOverlay.classList.remove('hidden');
    try {
        for (const classLoan of groupedLoans) {
            const { error: deleteLoansError } = await supabase
                .from('loans')
                .delete()
                .eq('class_loan_id', classLoan.id)
                .eq('user_id', currentUserId);
            if (deleteLoansError) throw deleteLoansError;
            const { error: deleteClassLoanError } = await supabase
                .from('class_loans')
                .delete()
                .eq('id', classLoan.id)
                .eq('user_id', currentUserId);
            if (deleteClassLoanError) throw deleteClassLoanError;
        }
        await loadClassLoans(currentUserId);
        await loadLoans(currentUserId);
        renderAll();
        try {
            if (typeof window.showToast === 'function') {
                const titles = (bookTitles || []).filter(Boolean).join(', ');
                window.showToast(`បានលុបការខ្ចីតាមថ្នាក់ «${className}» កាលបរិច្ឆេទ ${loanDate} សៀវភៅ៖ ${titles} ដោយជោគជ័យ!`, 'bg-green-600');
            }
        } catch (_) { /* noop */ }
    } catch (e) {
        console.error('Error deleting grouped class loans: ', e);
        alert('ការលុបបានបរាជ័យ។');
    } finally {
        loadingOverlay.classList.add('hidden');
    }
};

window.deleteGroupedClassLoan = (groupKey) => {
    if (!currentUserId) return;
    window.openClassLoanDeleteModal(groupKey);
};

window.openClassLoanEditModal = (id) => {
    const classLoan = classLoans.find(cl => cl.id === id);
    if (!classLoan) return;
    const book = books.find(b => b.id === classLoan.book_id);
    document.getElementById('class-loan-edit-id').value = id;
    document.getElementById('class-loan-edit-book-title').textContent = book ? book.title : 'N/A';
    document.getElementById('class-loan-edit-class-name').textContent = classLoan.class_name;
    document.getElementById('class-loan-edit-date').value = classLoan.loan_date;
    document.getElementById('class-loan-edit-modal').classList.remove('hidden');
};

window.closeClassLoanEditModal = () => document.getElementById('class-loan-edit-modal').classList.add('hidden');

document.getElementById('class-loan-edit-form').addEventListener('submit', async (e) => {
    e.preventDefault(); if (!currentUserId) return;
    const id = document.getElementById('class-loan-edit-id').value;
    const newDate = document.getElementById('class-loan-edit-date').value;
    if (!newDate) { alert("សូមជ្រើសរើសកាលបរិច្ឆេទ។"); return; }
    loadingOverlay.classList.remove('hidden');
    try {
        // Update class loan date
        const { error: classLoanError } = await supabase
            .from('class_loans')
            .update({ loan_date: newDate })
            .eq('id', id)
            .eq('user_id', currentUserId);
        
        if (classLoanError) throw classLoanError;
        
        // Update individual loan dates
        const { error: loansError } = await supabase
            .from('loans')
            .update({ loan_date: newDate })
            .eq('class_loan_id', id)
            .eq('user_id', currentUserId);
        
        if (loansError) throw loansError;
        alert("បានកែប្រែកាលបរិច្ឆេទខ្ចីដោយជោគជ័យ។");
        closeClassLoanEditModal();
    } catch (err) { console.error("Error editing class loan date: ", err); alert("ការកែប្រែបានបរាជ័យ។");
    } finally { loadingOverlay.classList.add('hidden'); }
});

window.deleteClassLoan = async (id) => {
    if (!currentUserId) return;
    const classLoanInfo = classLoans.find(cl => String(cl.id) === String(id));
    const className = classLoanInfo?.class_name || '';
    const loanDate = classLoanInfo?.loan_date || '';
    const bookTitle = (() => { const b = books.find(x => String(x.id) === String(classLoanInfo?.book_id)); return (b?.title || '').trim(); })();
    if (confirm('តើអ្នកពិតជាចង់លុបប្រវត្តិនៃការខ្ចីតាមថ្នាក់នេះមែនទេ? ការធ្វើបែបនេះនឹងលុបកំណត់ត្រាខ្ចីរបស់សិស្សទាំងអស់ដែលពាក់ព័ន្ធនឹងការខ្ចីនេះ។')) {
        loadingOverlay.classList.remove('hidden');
        try {
            // Delete individual loans first
            const { error: deleteLoansError } = await supabase
                .from('loans')
                .delete()
                .eq('class_loan_id', id)
                .eq('user_id', currentUserId);
            
            if (deleteLoansError) throw deleteLoansError;
            
            // Delete class loan
            const { error: deleteClassLoanError } = await supabase
                .from('class_loans')
                .delete()
                .eq('id', id)
                .eq('user_id', currentUserId);
            
            if (deleteClassLoanError) throw deleteClassLoanError;
            
            // Refresh data and re-render after successful deletion
            await loadClassLoans(currentUserId);
            await loadLoans(currentUserId);
            renderAll();
            try {
                if (typeof window.showToast === 'function') {
                    window.showToast(`បានលុបការខ្ចីតាមថ្នាក់ «${className}» កាលបរិច្ឆេទ ${loanDate} សៀវភៅ៖ ${bookTitle} ដោយជោគជ័យ!`, 'bg-green-600');
                }
            } catch (_) { /* noop */ }
            
        } catch (e) { console.error("Error deleting class loan and associated loans: ", e); alert("ការលុបបានបរាជ័យ។");
        } finally { loadingOverlay.classList.add('hidden'); }
    }
};

// --- LOCATION MANAGEMENT ---
window.openLocationModal = (id = null) => {
    const locationForm = document.getElementById('location-form');
    locationForm.reset(); document.getElementById('location-id').value = '';
    if (id) {
        const loc = locations.find(l => l.id === id);
        if (loc) {
            document.getElementById('location-modal-title').textContent = 'កែសម្រួលទីតាំង';
            document.getElementById('location-id').value = loc.id;
            document.getElementById('location-name').value = loc.name;
            document.getElementById('location-source').value = loc.source || '';
            document.getElementById('location-year').value = loc.year || '';
        }
    } else { document.getElementById('location-modal-title').textContent = 'បន្ថែមទីតាំងថ្មី'; }
    document.getElementById('location-modal').classList.remove('hidden');
};
window.closeLocationModal = () => document.getElementById('location-modal').classList.add('hidden');
window.closeLocationModal = () => {
    const modal = document.getElementById('location-modal');
    if (modal) modal.classList.add('hidden');
    // Return focus to Locations search box
    setTimeout(() => {
        const input = document.getElementById('search-locations');
        if (input && !document.getElementById('location-modal')?.classList.contains('hidden')) return; // still open
        if (input) { input.focus(); input.select(); }
    }, 50);
};
window.editLocation = (id) => openLocationModal(id);

// Location Delete Confirmation Modal controls
window.openLocationDeleteModal = (id) => {
    const modal = document.getElementById('location-delete-modal');
    const messageEl = document.getElementById('location-delete-message');
    const confirmBtn = document.getElementById('location-delete-confirm-btn');
    const loc = locations.find(l => String(l.id) === String(id));
    const isUsed = books.some(book => String(book.location_id) === String(id));
    if (isUsed) { 
        if (typeof window.openLocationInUseModal === 'function') {
            window.openLocationInUseModal();
        } else {
            alert('មិនអាចលុបទីតាំងនេះបានទេ ព្រោះកំពុងប្រើប្រាស់ដោយសៀវភៅ។');
        }
        return; 
    }
    if (messageEl) {
        const name = (loc?.name || '').trim();
        messageEl.textContent = name ? `តើអ្នកពិតជាចង់លុបទីតាំង «${name}» មែនទេ?` : 'តើអ្នកពិតជាចង់លុបទីតាំងនេះមែនទេ?';
    }
    if (confirmBtn) {
        confirmBtn.onclick = async () => {
            await window.performDeleteLocation(id);
            window.closeLocationDeleteModal();
        };
    }
    if (modal) modal.classList.remove('hidden');
};

window.closeLocationDeleteModal = () => {
    const modal = document.getElementById('location-delete-modal');
    const confirmBtn = document.getElementById('location-delete-confirm-btn');
    if (confirmBtn) confirmBtn.onclick = null;
    if (modal) modal.classList.add('hidden');
};

window.performDeleteLocation = async (id) => {
    if (!currentUserId) return;
    const locToDelete = locations.find(l => String(l.id) === String(id));
    try { 
        const { error } = await supabase
            .from('locations')
            .delete()
            .eq('id', id)
            .eq('user_id', currentUserId);
        if (error) console.error('Error deleting location: ', error);
        else {
            await loadLocations(currentUserId);
            renderAll();
            try {
                if (typeof window.showToast === 'function') {
                    const name = (locToDelete?.name || '').trim();
                    window.showToast(`បានលុបទីតាំង «${name}» ដោយជោគជ័យ!`, 'bg-green-600');
                }
            } catch (_) { /* noop */ }
        }
    } catch (e) { console.error('Error deleting location: ', e); }
};

window.deleteLocation = (id) => {
    if (!currentUserId) return;
    window.openLocationDeleteModal(id);
};

// Location In-Use Info Modal controls
window.openLocationInUseModal = () => {
    const modal = document.getElementById('location-inuse-modal');
    if (modal) modal.classList.remove('hidden');
};

window.closeLocationInUseModal = () => {
    const modal = document.getElementById('location-inuse-modal');
    if (modal) modal.classList.add('hidden');
};

document.getElementById('location-form').addEventListener('submit', async (e) => {
    e.preventDefault(); if (!currentUserId) return;
    const id = document.getElementById('location-id').value;
    const locData = { 
        name: document.getElementById('location-name').value, 
        source: document.getElementById('location-source').value, 
        year: document.getElementById('location-year').value,
        user_id: currentUserId
    };
    try { 
        if (id) { 
            const { error } = await supabase
                .from('locations')
                .update(locData)
                .eq('id', id)
                .eq('user_id', currentUserId);
            if (error) throw error;
        } else { 
            const { error } = await supabase
                .from('locations')
                .insert([locData]);
            if (error) throw error;
        } 
        await loadLocations(currentUserId);
        renderAll();
        // Centered toast for successful location save
        try {
            const locName = (locData.name || '').trim();
            if (typeof window.showToast === 'function') {
                window.showToast(`បានរក្សាទុកទីតាំង ${locName} ដោយជោគជ័យ!`, 'bg-green-600');
            }
        } catch (_) { /* noop */ }
        closeLocationModal(); 
    } catch (e) { console.error("Error adding/updating location: ", e); }
});

// --- STUDENT MANAGEMENT ---
window.openStudentModal = (id = null) => {
    const studentForm = document.getElementById('student-form');
    studentForm.reset();
    document.getElementById('student-id').value = '';
    const modalTitle = document.getElementById('student-modal-title');
    
    if (id) {
        const student = students.find(s => s.id === id);
        if (student) {
            modalTitle.textContent = 'កែសម្រួលព័ត៌មានសិស្ស';
            document.getElementById('student-id').value = student.id;
            document.getElementById('student-no').value = student['ល.រ'] || '';
            document.getElementById('student-code').value = student['អត្តលេខ'] || '';
            document.getElementById('student-lastname').value = student['នាមត្រកូល'] || '';
            document.getElementById('student-firstname').value = student['នាមខ្លួន'] || '';
            document.getElementById('student-gender').value = student['ភេទ'] || 'ប្រុស';
            document.getElementById('student-dob').value = student['ថ្ងៃខែឆ្នាំកំណើត'] || '';
            document.getElementById('student-class').value = student['ថ្នាក់'] || '';
            document.getElementById('student-photo-url').value = student['រូបថត URL'] || '';
        }
    } else {
        modalTitle.textContent = 'បន្ថែមសិស្សថ្មី';
        // Intelligent "ល.រ" logic
        let nextStudentNo = 1;
        if (students.length > 0) {
            const existingNos = students
                .map(s => parseInt(s['ល.រ'], 10))
                .filter(n => !isNaN(n));
            
            const noSet = new Set(existingNos);
            let i = 1;
            while (true) {
                if (!noSet.has(i)) {
                    nextStudentNo = i;
                    break;
                }
                i++;
            }
        }
        document.getElementById('student-no').value = nextStudentNo;
    }
    document.getElementById('student-modal').classList.remove('hidden');
};

window.closeStudentModal = () => {
    const modal = document.getElementById('student-modal');
    if (modal) modal.classList.add('hidden');
    // Return focus to Students search box
    setTimeout(() => {
        const input = document.getElementById('search-students');
        if (input && !document.getElementById('student-modal')?.classList.contains('hidden')) return; // still open
        if (input) { input.focus(); input.select(); }
    }, 50);
};

// Student Delete Confirmation Modal controls
window.openStudentDeleteModal = (id) => {
    const modal = document.getElementById('student-delete-modal');
    const messageEl = document.getElementById('student-delete-message');
    const confirmBtn = document.getElementById('student-delete-confirm-btn');
    const std = students.find(s => String(s.id) === String(id));
    const fullName = `${(std?.['នាមត្រកូល'] || '').trim()} ${(std?.['នាមខ្លួន'] || '').trim()}`.trim();
    if (messageEl) {
        messageEl.textContent = fullName
            ? `តើអ្នកពិតជាចង់លុបព័ត៌មានសិស្សឈ្មោះ ${fullName} មែនទេ?`
            : 'តើអ្នកពិតជាចង់លុបព័ត៌មានសិស្សនេះមែនទេ?';
    }
    if (confirmBtn) {
        confirmBtn.onclick = async () => {
            await window.performDeleteStudent(id);
            window.closeStudentDeleteModal();
        };
    }
    if (modal) modal.classList.remove('hidden');
};

window.closeStudentDeleteModal = () => {
    const modal = document.getElementById('student-delete-modal');
    const confirmBtn = document.getElementById('student-delete-confirm-btn');
    if (confirmBtn) confirmBtn.onclick = null;
    if (modal) modal.classList.add('hidden');
};

window.performDeleteStudent = async (id) => {
    if (!currentUserId) return;
    const std = students.find(s => String(s.id) === String(id));
    try {
        const { error } = await supabase
            .from('students')
            .delete()
            .eq('id', id)
            .eq('user_id', currentUserId);
        if (error) throw error;
        await loadStudents(currentUserId);
        populateClassLoanFilter();
        populateStudentClassFilter();
        renderAll();
        try {
            if (typeof window.showToast === 'function') {
                const fullName = `${(std?.['នាមត្រកូល'] || '').trim()} ${(std?.['នាមខ្លួន'] || '').trim()}`.trim();
                window.showToast(`បានលុបព័ត៌មានសិស្សឈ្មោះ ${fullName} ដោយជោគជ័យ!`, 'bg-green-600');
            }
        } catch (_) { /* noop */ }
    } catch (e) {
        console.error('Error deleting student: ', e);
        alert('ការលុបបានបរាជ័យ។');
    }
};

window.deleteStudent = (id) => {
    if (!currentUserId) return;
    window.openStudentDeleteModal(id);
};

document.getElementById('student-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentUserId) {
        alert('សូមចូលគណនីមុនពេលរក្សាទុកទិន្នន័យ។');
        return;
    }

    const studentId = document.getElementById('student-id').value;
    // Note: Excluding columns that are not present in current DB schema
    const studentData = {
        'ល.រ': document.getElementById('student-no').value?.trim(),
        'អត្តលេខ': document.getElementById('student-code').value?.trim(),
        'នាមត្រកូល': document.getElementById('student-lastname').value?.trim(),
        'នាមខ្លួន': document.getElementById('student-firstname').value?.trim(),
        'ភេទ': document.getElementById('student-gender').value?.trim(),
        'ថ្នាក់': document.getElementById('student-class').value?.trim(),
        'ថ្ងៃខែឆ្នាំកំណើត': document.getElementById('student-dob').value?.trim(),
        'រូបថត URL': document.getElementById('student-photo-url').value?.trim()
    };

    if (!studentData['អត្តលេខ'] || !studentData['នាមត្រកូល'] || !studentData['នាមខ្លួន']) {
        alert('សូមបំពេញអត្តលេខ, នាមត្រកូល, និងនាមខ្លួន។');
        return;
    }

    // Convert empty strings to null to avoid DB type errors
    for (const k of Object.keys(studentData)) {
        if (studentData[k] === '') studentData[k] = null;
    }
    // Optionally coerce serial number to integer if provided
    if (studentData['ល.រ'] != null) {
        const n = Number(studentData['ល.រ']);
        studentData['ល.រ'] = Number.isNaN(n) ? studentData['ល.រ'] : n;
    }

    // Build Supabase payload using Khmer column names to match database
    const supabasePayload = {
        user_id: currentUserId,
        'អត្តលេខ': studentData['អត្តលេខ'],
        'នាមត្រកូល': studentData['នាមត្រកូល'],
        'នាមខ្លួន': studentData['នាមខ្លួន'],
        'ភេទ': studentData['ភេទ'],
        'ថ្នាក់': studentData['ថ្នាក់'],
        'ល.រ': studentData['ល.រ'],
        'ថ្ងៃខែឆ្នាំកំណើត': studentData['ថ្ងៃខែឆ្នាំកំណើត'],
        'រូបថត URL': studentData['រូបថត URL']
    };

    try {
        if (studentId) {
            // Update existing student
            const { error } = await supabase
                .from('students')
                .update(supabasePayload)
                .eq('id', studentId)
                .eq('user_id', currentUserId);
            if (error) throw error;
        } else {
            // Add new student
            const { error } = await supabase
                .from('students')
                .insert([supabasePayload]);
            if (error) throw error;
        }
        await loadStudents(currentUserId);
        populateClassLoanFilter();
        populateStudentClassFilter();
        closeStudentModal();
        renderAll();
        // Centered toast for successful student save
        try {
            const fullName = `${(supabasePayload['នាមត្រកូល'] || '').trim()} ${(supabasePayload['នាមខ្លួន'] || '').trim()}`.trim();
            if (typeof window.showToast === 'function') {
                window.showToast(`បានរក្សាទុកសិស្សឈ្មោះ ${fullName} ដោយជោគជ័យ!`, 'bg-green-600');
            }
        } catch (_) { /* noop */ }
    } catch (err) {
        console.error('Error saving student data: ', err);
        // Friendly messages for common cases
        if (err && (err.code === '23505' || (err.message && err.message.includes('duplicate key')))) {
            alert('ការរក្សាទុកបរាជ័យ: អត្តលេខនេះមានរួចហើយ។');
        } else if (err && err.message) {
            alert(`ការរក្សាទុកទិន្នន័យសិស្សបានបរាជ័យ៖ ${err.message}`);
        } else {
            alert('ការរក្សាទុកទិន្នន័យសិស្សបានបរាជ័យ។');
        }
    }
});


// --- READING LOG MANAGEMENT ---
let isbnScanTimer = null;

// MODIFIED: Prevent Enter key from submitting the form and handle scan flow
document.getElementById('reading-log-form').addEventListener('keydown', function(event) {
    if (event.key === 'Enter' || event.keyCode === 13) {
        event.preventDefault(); // Always prevent the default form submission on Enter

        const activeElement = document.activeElement;

        // If Enter is pressed in the student ID field, move focus to the ISBN field
        if (activeElement && activeElement.id === 'reading-log-student-id') {
            document.getElementById('reading-log-isbn-input').focus();
            return; 
        }

        // If Enter is pressed in the ISBN field, process the ISBN immediately
        if (activeElement && activeElement.id === 'reading-log-isbn-input') {
            clearTimeout(isbnScanTimer); // Stop any pending input event timer
            
            const isbnInput = document.getElementById('reading-log-isbn-input');
            // Convert Khmer numbers and normalize value
            const converted = convertKhmerToEnglishNumbers(isbnInput.value || '');
            if (converted !== isbnInput.value) isbnInput.value = converted;
            const isbn = converted.trim();
            if (!isbn) return;
            // Enforce single scanned book
            if (currentScannedBooks.length >= 1) {
                try { if (typeof window.showToast === 'function') window.showToast('សូមស្កេនតែមួយក្បាល បន្ទាប់មកចុច Enter ដើម្បីរក្សាទុក។', 'bg-yellow-600'); } catch(_){}
                isbnInput.value = '';
                return;
            }

            const foundBook = books.find(b => b.isbn && b.isbn.toLowerCase() === isbn.toLowerCase());
            if (foundBook) {
                if (!currentScannedBooks.some(b => b.id === foundBook.id)) {
                    currentScannedBooks.push({ id: foundBook.id, title: foundBook.title });
                    const li = document.createElement('li');
                    li.textContent = foundBook.title;
                    document.getElementById('scanned-books-list').appendChild(li);
                }
                // Disable further scans until save/clear
                isbnInput.value = '';
                isbnInput.disabled = true;
                isbnInput.placeholder = 'បានស្កេន 1 ក្បាល។ ចុច Enter ដើម្បីរក្សាទុក';
                // Trigger submit (will validate student fields inside submit handler)
                const form = document.getElementById('reading-log-form');
                if (form && typeof form.requestSubmit === 'function') form.requestSubmit(); else form.submit();
            }
            // If book is not found, do nothing, leave the value for correction.
        }
    }
});

window.clearReadingLogForm = () => {
    document.getElementById('reading-log-form').reset();
    document.getElementById('reading-log-student-name').value = '';
    document.getElementById('reading-log-student-obj-id').value = '';
    document.getElementById('reading-log-student-error').textContent = '';
    document.getElementById('scanned-books-list').innerHTML = '';
    currentScannedBooks = [];
    currentStudentGender = '';
    const isbnEl = document.getElementById('reading-log-isbn-input');
    if (isbnEl) { isbnEl.disabled = false; isbnEl.placeholder = 'ស្កេនសៀវភៅម្តងមួយ...'; }
    document.getElementById('reading-log-student-id').focus();
};

document.getElementById('reading-log-student-id').addEventListener('input', () => {
    const studentId = document.getElementById('reading-log-student-id').value.trim();
    const studentNameInput = document.getElementById('reading-log-student-name');
    const studentObjIdInput = document.getElementById('reading-log-student-obj-id');
    const studentError = document.getElementById('reading-log-student-error');
    studentNameInput.value = ''; studentObjIdInput.value = ''; studentError.textContent = ''; currentStudentGender = '';
    if (!studentId) return;
    const studentIdKey = students.length > 0 ? Object.keys(students[0]).find(k => k.toLowerCase().includes('អត្តលេខ')) : null;
    if (!studentIdKey) { studentError.textContent = 'រកមិនឃើញជួរឈរ "អត្តលេខ" ទេ។'; return; }
    const foundStudent = students.find(s => s[studentIdKey] && s[studentIdKey].trim() === studentId);
    if (foundStudent) {
        const lastNameKey = Object.keys(foundStudent).find(k => k.includes('នាមត្រកូល'));
        const firstNameKey = Object.keys(foundStudent).find(k => k.includes('នាមខ្លួន'));
        const classKey = Object.keys(foundStudent).find(k => k.includes('ថ្នាក់'));
        const genderKey = Object.keys(foundStudent).find(k => k.includes('ភេទ'));
        const studentFullName = `${foundStudent[lastNameKey] || ''} ${foundStudent[firstNameKey] || ''}`.trim();
        const studentClass = foundStudent[classKey] || '';
        studentNameInput.value = studentClass ? `${studentFullName} - ${studentClass}` : studentFullName;
        studentObjIdInput.value = foundStudent.id;
        currentStudentGender = foundStudent[genderKey] || '';
        document.getElementById('reading-log-isbn-input').focus();
    } else { studentError.textContent = 'រកមិនឃើញអត្តលេខសិស្សនេះទេ។'; }
});

document.getElementById('reading-log-isbn-input').addEventListener('input', (e) => {
    // Convert Khmer numbers to English numbers
    const convertedValue = convertKhmerToEnglishNumbers(e.target.value);
    if (convertedValue !== e.target.value) {
        e.target.value = convertedValue;
    }
    
    clearTimeout(isbnScanTimer);
    isbnScanTimer = setTimeout(() => {
        const isbnInput = document.getElementById('reading-log-isbn-input');
        const isbn = convertedValue.trim();
        // Enforce single scanned book
        if (currentScannedBooks.length >= 1) {
            try { if (typeof window.showToast === 'function') window.showToast('បានស្កេនរួចហើយ 1 ក្បាល។ សូមរក្សាទុកជាមុនសិន។', 'bg-yellow-600'); } catch(_){}
            if (isbnInput) isbnInput.value = '';
            return;
        }
        if (!isbn) return;
        const foundBook = books.find(b => b.isbn && b.isbn.toLowerCase() === isbn.toLowerCase());
        if (foundBook) {
            if (!currentScannedBooks.some(b => b.id === foundBook.id)) {
                currentScannedBooks.push({ id: foundBook.id, title: foundBook.title });
                const li = document.createElement('li');
                li.textContent = foundBook.title;
                document.getElementById('scanned-books-list').appendChild(li);
            }
            // Disable further scans until save/clear
            isbnInput.value = '';
            isbnInput.disabled = true;
            isbnInput.placeholder = 'បានស្កេន 1 ក្បាល។ ចុច Enter ដើម្បីរក្សាទុក';
        }
    }, 300); 
});

document.getElementById('reading-log-form').addEventListener('submit', async (e) => {
    e.preventDefault(); if (!currentUserId) return;
    const studentId = document.getElementById('reading-log-student-obj-id').value;
    const studentName = document.getElementById('reading-log-student-name').value;
    if (!studentId || !studentName) { alert('សូមស្កេនអត្តលេខសិស្សជាមុនសិន។'); return; }
    if (currentScannedBooks.length === 0) { alert('សូមស្កេនសៀវភៅយ៉ាងហោចណាស់មួយក្បាល។'); return; }
    const logData = { 
        student_id: studentId, 
        student_name: studentName, 
        student_gender: currentStudentGender, 
        books: currentScannedBooks, 
        date_time: new Date().toISOString(),
        user_id: currentUserId
    };
    try { 
        const { error } = await supabase
            .from('reading_logs')
            .insert([logData]);
        if (error) throw error;
        await loadReadingLogs(currentUserId);
        renderAll();
        // Centered toast for successful reading log save
        try {
            const titles = currentScannedBooks.map(b => (b.title || '').trim()).filter(Boolean).join(',');
            if (typeof window.showToast === 'function') {
                window.showToast(`បានរក្សាកំណត់ត្រាចូលអាន របស់ សិស្សឈ្មោះ ${studentName}+${titles} ដោយជោគជ័យ!`, 'bg-green-600');
            }
        } catch (_) { /* noop */ }
        window.clearReadingLogForm(); 
    } 
    catch (err) { console.error("Error saving reading log: ", err); alert("ការរក្សាទុកកំណត់ត្រាចូលអានបានបរាជ័យ។"); }
});

// Reading Log Delete Confirmation Modal controls
window.openReadingDeleteModal = (id) => {
    const modal = document.getElementById('reading-delete-modal');
    const messageEl = document.getElementById('reading-delete-message');
    const confirmBtn = document.getElementById('reading-delete-confirm-btn');
    const log = readingLogs.find(r => String(r.id) === String(id));
    const name = (log?.student_name || '').trim();
    if (messageEl) {
        messageEl.textContent = name
            ? `តើអ្នកពិតជាចង់លុបកំណត់ត្រាចូលអាន របស់ សិស្សឈ្មោះ ${name} មែនទេ?`
            : 'តើអ្នកពិតជាចង់លុបកំណត់ត្រាចូលអាននេះមែនទេ?';
    }
    if (confirmBtn) {
        confirmBtn.onclick = async () => {
            await window.performDeleteReadingLog(id);
            window.closeReadingDeleteModal();
        };
    }
    if (modal) modal.classList.remove('hidden');
};

window.closeReadingDeleteModal = () => {
    const modal = document.getElementById('reading-delete-modal');
    const confirmBtn = document.getElementById('reading-delete-confirm-btn');
    if (confirmBtn) confirmBtn.onclick = null;
    if (modal) modal.classList.add('hidden');
};

window.performDeleteReadingLog = async (id) => {
    if (!currentUserId) return;
    const log = readingLogs.find(r => String(r.id) === String(id));
    try {
        const { error } = await supabase
            .from('reading_logs')
            .delete()
            .eq('id', id)
            .eq('user_id', currentUserId);
        if (error) {
            console.error('Error deleting reading log: ', error);
        } else {
            await loadReadingLogs(currentUserId);
            renderAll();
            try {
                if (typeof window.showToast === 'function') {
                    const name = (log?.student_name || '').trim();
                    window.showToast(`បានលុបកំណត់ត្រាចូលអាន របស់ សិស្សឈ្មោះ ${name} ដោយជោគជ័យ!`, 'bg-green-600');
                }
            } catch (_) { /* noop */ }
        }
    } catch (e) { console.error('Error deleting reading log: ', e); alert('ការលុបបានបរាជ័យ។'); }
};

window.deleteReadingLog = (id) => {
    if (!currentUserId) return;
    window.openReadingDeleteModal(id);
};

// Delete all logs for a given date (YYYY-MM-DD) and student name
window.deleteStudentLogsForDateAndStudent = async (encodedDateKey, encodedStudentName) => {
    if (!currentUserId) return;
    const dateKey = decodeURIComponent(encodedDateKey);
    const studentName = (decodeURIComponent(encodedStudentName) || '').trim();
    // Collect matching log IDs
    const toDelete = (readingLogs || []).filter(r => {
        const d = (r.date_time || '').split('T')[0];
        const n = (r.student_name || '').trim();
        return d === dateKey && n === studentName;
    }).map(r => r.id);

    if (!toDelete.length) {
        alert('មិនមានកំណត់ត្រាត្រូវលុបទេ។');
        return;
    }

    // Format date for confirmation
    const displayDate = (() => {
        try {
            const dt = new Date(dateKey);
            const dd = String(dt.getDate()).padStart(2, '0');
            const mm = String(dt.getMonth() + 1).padStart(2, '0');
            const yyyy = dt.getFullYear();
            return `${dd}/${mm}/${yyyy}`;
        } catch { return dateKey; }
    })();

    const ok = confirm(`តើអ្នកពិតជាចង់លុបកំណត់ត្រាចូលអាន របស់ \nសិស្សឈ្មោះ ${studentName} \nនៅថ្ងៃ ${displayDate} \nចំនួន ${toDelete.length} កំណត់ត្រា មែនទេ?`);
    if (!ok) return;

    // Delete sequentially to reuse existing logic and toasts
    for (const id of toDelete) {
        try {
            await window.performDeleteReadingLog(id);
        } catch (e) {
            console.error('Batch delete error: ', e);
        }
    }
};

// --- EXPORT DATA TO EXCEL ---
const exportData = () => {
    const select = document.getElementById('export-data-select');
    const dataType = select.value;

    if (!dataType) {
        alert('សូមជ្រើសរើសប្រភេទទិន្នន័យដែលត្រូវនាំចេញជាមុនសិន។');
        return;
    }

    let dataToExport = [];
    let fileName = `${dataType}_export.xlsx`;
    
    switch (dataType) {
        case 'books':
            dataToExport = books.map(book => {
                const location = locations.find(loc => loc.id === book.location_id);
                return {
                    'ចំណងជើង': book.title,
                    'អ្នកនិពន្ធ': book.author,
                    'ISBN': book.isbn,
                    'ចំនួនសរុប': book.quantity,
                    'ទីតាំង': location ? location.name : 'N/A',
                    'ប្រភព': book.source
                };
            });
            break;
        case 'loans':
            dataToExport = loans.filter(l => !l.class_loan_id).map(loan => {
                const book = books.find(b => b.id === loan.book_id);
                return {
                    'សៀវភៅ': book ? book.title : 'N/A',
                    'អ្នកខ្ចី': loan.borrower,
                    'ភេទ': loan.borrower_gender,
                    'ថ្ងៃខ្ចី': loan.loan_date,
                    'ថ្ងៃសង': loan.return_date,
                    'ស្ថានភាព': loan.status
                };
            });
            break;
        case 'class-loans':
            dataToExport = classLoans.map(loan => {
                const book = books.find(b => b.id === loan.book_id);
                return {
                    'សៀវភៅ': book ? book.title : 'N/A',
                    'ថ្នាក់': loan.class_name,
                    'ថ្ងៃខ្ចី': loan.loan_date,
                    'ចំនួនខ្ចី': loan.loaned_quantity,
                    'ចំនួនបានសង': loan.returned_count || 0,
                    'ស្ថានភាព': loan.status
                };
            });
            break;
        case 'reading-logs':
            dataToExport = readingLogs.map(log => {
                // Handle books field - it might be a string (JSON) or already an array
                let booksArray = [];
                if (log.books) {
                    if (typeof log.books === 'string') {
                        try {
                            booksArray = JSON.parse(log.books);
                        } catch (e) {
                            booksArray = [];
                        }
                    } else if (Array.isArray(log.books)) {
                        booksArray = log.books;
                    }
                }
                
                return {
                    'កាលបរិច្ឆេទ': new Date(log.date_time).toLocaleString('en-GB'),
                    'ឈ្មោះសិស្ស': log.student_name,
                    'ភេទ': log.student_gender,
                    'សៀវភៅបានអាន': booksArray.map(b => b.title || b).join('; ')
                };
            });
            break;
        case 'locations':
            dataToExport = locations.map(loc => ({
                'ឈ្មោះទីតាំង': loc.name,
                'ប្រភពធ្នើ': loc.source,
                'ឆ្នាំ': loc.year
            }));
            break;
        case 'students':
            dataToExport = students.map(std => {
                // Create a new object to control property order
                const studentData = {};
                studentData['ល.រ'] = std['ល.រ'];
                studentData['អត្តលេខ'] = std['អត្តលេខ'];
                studentData['នាមត្រកូល'] = std['នាមត្រកូល'];
                studentData['នាមខ្លួន'] = std['នាមខ្លួន'];
                studentData['ភេទ'] = std['ភេទ'];
                studentData['ថ្នាក់'] = std['ថ្នាក់'];
                studentData['ថ្ងៃខែឆ្នាំកំណើត'] = std['ថ្ងៃខែឆ្នាំកំណើត'];
                studentData['រូបថត URL'] = std['រូបថត URL'];
                return studentData;
            });
            break;
        case 'settings':
            dataToExport = Object.entries(settingsData).map(([key, value]) => ({
                'ការកំណត់': key,
                'តម្លៃ': value
            }));
            break;
    }

    if (dataToExport.length === 0) {
        alert('មិនមានទិន្នន័យសម្រាប់នាំចេញទេ។');
        return;
    }

    // Using SheetJS to create and download the Excel file
    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
    XLSX.writeFile(workbook, fileName);
};


// --- SETTINGS & DATA DELETION ---
window.openPasswordConfirmModal = (collectionName) => {
    document.getElementById('collection-to-delete').value = collectionName;
    const modalEl = document.getElementById('password-confirm-modal');
    const inputEl = document.getElementById('user-password');
    const errorEl = document.getElementById('password-error');
    modalEl.classList.remove('hidden');
    errorEl.textContent = '';
    // Defer focus to ensure element is visible and focusable after unhide
    setTimeout(() => {
        try {
            inputEl.focus();
            if (typeof inputEl.select === 'function') inputEl.select();
        } catch (_) {}
    }, 0);
};

window.closePasswordConfirmModal = () => {
    document.getElementById('password-confirm-modal').classList.add('hidden');
    document.getElementById('password-confirm-form').reset();
};

document.getElementById('password-confirm-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentUserId) return;

    const password = document.getElementById('user-password').value;
    const collectionName = document.getElementById('collection-to-delete').value;
    const passwordError = document.getElementById('password-error');
    passwordError.textContent = '';

    if (!password || !collectionName) return;

    loadingOverlay.classList.remove('hidden');
    
    try {
        // For Supabase, we'll verify the password by attempting to sign in
        const { data, error } = await supabase.auth.signInWithPassword({
            email: (await supabase.auth.getUser()).data.user.email,
            password: password
        });
        
        if (error) throw error;
        
        // Re-authentication successful, proceed with deletion
        await deleteAllData(collectionName);
        closePasswordConfirmModal();
    } catch (error) {
        console.error("Re-authentication failed:", error);
        passwordError.textContent = "ពាក្យសម្ងាត់មិនត្រឹមត្រូវទេ។";
    } finally {
        loadingOverlay.classList.add('hidden');
    }
});

async function deleteAllData(collectionName) {
    if (!currentUserId) return;
    
    // Map collection names to actual table names
    const tableMapping = {
        'books': 'books',
        'loans': 'loans', 
        'classLoans': 'class_loans',
        'readingLogs': 'reading_logs',
        'locations': 'locations',
        'students': 'students'
    };
    
    const tableName = tableMapping[collectionName];
    if (!tableName) {
        console.error(`Unknown collection name: ${collectionName}`);
        alert(`មិនស្គាល់ប្រភេទទិន្នន័យ "${collectionName}"`);
        return;
    }
    
    try {
        const { error } = await supabase
            .from(tableName)
            .delete()
            .eq('user_id', currentUserId);
            
        if (error) throw error;
        
        // Clear local arrays based on collection type
        switch(collectionName) {
            case 'books': books = []; break;
            case 'loans': loans = []; break;
            case 'classLoans': classLoans = []; break;
            case 'readingLogs': readingLogs = []; break;
            case 'locations': locations = []; break;
            case 'students': students = []; break;
        }
        
        // Re-render all data
        renderAll();
        
        try { if (typeof window.showToast === 'function') { window.showToast(`បានលុបទិន្នន័យទាំងអស់ក្នុង "${collectionName}" ដោយជោគជ័យ!`, 'bg-green-600'); } } catch (_) { }
    } catch (error) {
        console.error(`Error deleting all documents from ${collectionName}:`, error);
        alert(`ការលុបទិន្នន័យក្នុង "${collectionName}" បានបរាជ័យ។`);
    }
}

// --- SETTINGS GENERAL INFO LISTENERS ---
const saveSchoolNameBtn = document.getElementById('save-school-name-btn');
const editSchoolNameBtn = document.getElementById('edit-school-name-btn');
const deleteSchoolNameBtn = document.getElementById('delete-school-name-btn');
const schoolNameInput = document.getElementById('school-name-input');
const schoolNameDisplay = document.getElementById('school-name-display');

saveSchoolNameBtn.addEventListener('click', async () => {
    const name = schoolNameInput.value.trim();
    if (!name) { alert('សូមបញ្ចូលឈ្មោះសាលា។'); return; }
    try {
        const { error } = await supabase
            .from('settings')
            .upsert({ 
                key: 'schoolName', 
                value: name, 
                user_id: currentUserId 
            }, {
                onConflict: 'key,user_id'
            });
        if (error) throw error;
        await loadSettings(currentUserId);
        renderAll();
        try { if (typeof window.showToast === 'function') { window.showToast('រក្សាទុកឈ្មោះសាលាបានជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
    } catch (e) { console.error(e); alert('ការរក្សាទុកបានបរាជ័យ។'); }
});

editSchoolNameBtn.addEventListener('click', () => {
    schoolNameDisplay.classList.add('hidden');
    schoolNameInput.classList.remove('hidden');
    saveSchoolNameBtn.classList.remove('hidden');
    editSchoolNameBtn.classList.add('hidden');
    deleteSchoolNameBtn.classList.add('hidden');
    schoolNameInput.focus();
});

// Settings: School Name delete modal handlers
window.openSchoolNameDeleteModal = () => {
    const modal = document.getElementById('schoolname-delete-modal');
    const confirmBtn = document.getElementById('schoolname-delete-confirm-btn');
    if (confirmBtn) {
        confirmBtn.onclick = async () => {
            await window.performDeleteSchoolName();
            window.closeSchoolNameDeleteModal();
        };
    }
    if (modal) modal.classList.remove('hidden');
};

window.closeSchoolNameDeleteModal = () => {
    const modal = document.getElementById('schoolname-delete-modal');
    const confirmBtn = document.getElementById('schoolname-delete-confirm-btn');
    if (confirmBtn) confirmBtn.onclick = null;
    if (modal) modal.classList.add('hidden');
};

window.performDeleteSchoolName = async () => {
    try {
        const { error } = await supabase
            .from('settings')
            .delete()
            .eq('key', 'schoolName')
            .eq('user_id', currentUserId);
        if (error) throw error;
        await loadSettings(currentUserId);
        renderAll();
        try { if (typeof window.showToast === 'function') { window.showToast('បានលុបឈ្មោះសាលា ដោយជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
    } catch (e) { console.error(e); alert('ការលុបបានបរាជ័យ។'); }
};

deleteSchoolNameBtn.addEventListener('click', () => {
    window.openSchoolNameDeleteModal();
});

// Academic Year Listeners
const saveAcademicYearBtn = document.getElementById('save-academic-year-btn');
const editAcademicYearBtn = document.getElementById('edit-academic-year-btn');
const deleteAcademicYearBtn = document.getElementById('delete-academic-year-btn');
const academicYearInput = document.getElementById('academic-year-input');
const academicYearDisplay = document.getElementById('academic-year-display');

saveAcademicYearBtn.addEventListener('click', async () => {
    const year = academicYearInput.value.trim();
    if (!year) { alert('សូមបញ្ចូលឆ្នាំសិក្សា។'); return; }
    try {
        const { error } = await supabase
            .from('settings')
            .upsert({ 
                key: 'academicYear', 
                value: year, 
                user_id: currentUserId 
            }, {
                onConflict: 'key,user_id'
            });
        if (error) throw error;
        await loadSettings(currentUserId);
        renderAll();
        try { if (typeof window.showToast === 'function') { window.showToast('រក្សាទុកឆ្នាំសិក្សាបានជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
    } catch (e) { console.error(e); alert('ការរក្សាទុកបានបរាជ័យ។'); }
});

editAcademicYearBtn.addEventListener('click', () => {
    academicYearDisplay.classList.add('hidden');
    academicYearInput.classList.remove('hidden');
    saveAcademicYearBtn.classList.remove('hidden');
    editAcademicYearBtn.classList.add('hidden');
    deleteAcademicYearBtn.classList.add('hidden');
    academicYearInput.focus();
});

// Settings: Academic Year delete modal handlers
window.openAcademicYearDeleteModal = () => {
    const modal = document.getElementById('academicyear-delete-modal');
    const confirmBtn = document.getElementById('academicyear-delete-confirm-btn');
    if (confirmBtn) {
        confirmBtn.onclick = async () => {
            await window.performDeleteAcademicYear();
            window.closeAcademicYearDeleteModal();
        };
    }
    if (modal) modal.classList.remove('hidden');
};

window.closeAcademicYearDeleteModal = () => {
    const modal = document.getElementById('academicyear-delete-modal');
    const confirmBtn = document.getElementById('academicyear-delete-confirm-btn');
    if (confirmBtn) confirmBtn.onclick = null;
    if (modal) modal.classList.add('hidden');
};

window.performDeleteAcademicYear = async () => {
    try {
        const { error } = await supabase
            .from('settings')
            .delete()
            .eq('key', 'academicYear')
            .eq('user_id', currentUserId);
        if (error) throw error;
        await loadSettings(currentUserId);
        renderAll();
        try { if (typeof window.showToast === 'function') { window.showToast('បានលុបឆ្នាំសិក្សា ដោយជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
    } catch (e) { console.error(e); alert('ការលុបបានបរាជ័យ។'); }
};

deleteAcademicYearBtn.addEventListener('click', () => {
    window.openAcademicYearDeleteModal();
});


const sealImageUrlInput = document.getElementById('seal-image-url');
const sealImagePreview = document.getElementById('seal-image-preview');
const saveSealUrlBtn = document.getElementById('save-seal-url-btn');

sealImageUrlInput.addEventListener('input', () => {
    const url = sealImageUrlInput.value.trim();
    if (url) {
        sealImagePreview.src = url;
    } else {
        sealImagePreview.src = 'https://placehold.co/100x100/e2e8f0/e2e8f0?text=Preview';
    }
});

sealImagePreview.onerror = () => {
    sealImagePreview.src = 'https://placehold.co/100x100/e2e8f0/ff0000?text=Error';
};

saveSealUrlBtn.addEventListener('click', async () => {
    const url = sealImageUrlInput.value.trim();
    try {
        if (url) {
            const { error } = await supabase
                .from('settings')
                .upsert({ 
                    key: 'sealImageUrl', 
                    value: url, 
                    user_id: currentUserId 
                }, {
                    onConflict: 'key,user_id'
                });
            if (error) throw error;
            try { if (typeof window.showToast === 'function') { window.showToast('រក្សាទុក URL ត្រាបានជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
        } else {
            const { error } = await supabase
                .from('settings')
                .delete()
                .eq('key', 'sealImageUrl')
                .eq('user_id', currentUserId);
            if (error) throw error;
            try { if (typeof window.showToast === 'function') { window.showToast('បានលុប URL ត្រា ដោយជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
        }
    } catch (e) { console.error(e); alert('ការរក្សាទុកបានបរាជ័យ។'); }
});

// Card Background Listeners
const cardBgUrlInput = document.getElementById('card-bg-url');
const cardBgPreview = document.getElementById('card-bg-preview');
const saveCardBgBtn = document.getElementById('save-card-bg-btn');

cardBgUrlInput.addEventListener('input', () => {
    const url = cardBgUrlInput.value.trim();
    cardBgPreview.src = url || 'https://placehold.co/100x100/e2e8f0/e2e8f0?text=Preview';
});

cardBgPreview.onerror = () => {
    cardBgPreview.src = 'https://placehold.co/100x100/e2e8f0/ff0000?text=Error';
};

saveCardBgBtn.addEventListener('click', async () => {
    const url = cardBgUrlInput.value.trim();
    try {
        if (url) {
            const { error } = await supabase
                .from('settings')
                .upsert({ 
                    key: 'cardBgUrl', 
                    value: url, 
                    user_id: currentUserId 
                }, {
                    onConflict: 'key,user_id'
                });
            if (error) throw error;
            try { if (typeof window.showToast === 'function') { window.showToast('រក្សាទុក URL ផ្ទៃខាងក្រោយកាតបានជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
        } else {
            const { error } = await supabase
                .from('settings')
                .delete()
                .eq('key', 'cardBgUrl')
                .eq('user_id', currentUserId);
            if (error) throw error;
            try { if (typeof window.showToast === 'function') { window.showToast('បានលុប URL ផ្ទៃខាងក្រោយកាត ដោយជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
        }
    } catch (e) { console.error(e); alert('ការរក្សាទុកបានបរាជ័យ។'); }
});


// --- CHANGE PASSWORD LISTENER ---
const changePasswordForm = document.getElementById('change-password-form');
changePasswordForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentUserId) return;

    const currentPassword = document.getElementById('current-password').value;
    const newPassword = document.getElementById('new-password').value;
    const confirmPassword = document.getElementById('confirm-new-password').value;
    const errorP = document.getElementById('change-password-error');
    const successP = document.getElementById('change-password-success');

    errorP.textContent = '';
    successP.textContent = '';

    if (newPassword !== confirmPassword) {
        errorP.textContent = 'ពាក្យសម្ងាត់ថ្មី និងការយឺនយំនមិនតរងគ្នាទេ។';
        return;
    }
    if (newPassword.length < 6) {
        errorP.textContent = 'ពាក្យសម្ងាត់ថ្មីត្រូវមានយ៉ាងហោចណាស់ 6 តួអក្សរ។';
        return;
    }

    loadingOverlay.classList.remove('hidden');

    try {
        // Update password using Supabase
        const { error } = await supabase.auth.updateUser({
            password: newPassword
        });
        
        if (error) throw error;
        
        successP.textContent = 'ផ្លាស់ប្តូរពាក្យសម្ងាត់បានជោគជ័យ!';
        try { if (typeof window.showToast === 'function') { window.showToast('ផ្លាស់ប្តូរពាក្យសម្ងាត់បានជោគជ័យ!', 'bg-green-600'); } } catch (_) { }
        changePasswordForm.reset();
    } catch (error) {
        console.error("Password change failed:", error);
        errorP.textContent = 'ការផ្លាស់ប្តូរពាក្យសម្ងាត់បានបរាជ័យ។';
    } finally {
        loadingOverlay.classList.add('hidden');
    }
});


// --- CREATE ACCOUNT RGB ANIMATION (Settings) ---
(() => {
    const btn = document.getElementById('btn-create-account-settings');
    if (!btn) return;
    // Ensure starting styles for readability
    btn.style.backgroundColor = 'rgb(99, 102, 241)'; // indigo-ish
    btn.style.color = '#fff';

    const randomInt = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
    const setRandomBg = () => {
        const r = randomInt(32, 255);
        const g = randomInt(32, 255);
        const b = randomInt(32, 255);
        btn.style.backgroundColor = `rgb(${r}, ${g}, ${b})`;
    };

    // Change color periodically; transition is handled by Tailwind classes in HTML
    setRandomBg();
    const intervalId = setInterval(setRandomBg, 1500);

    // Optional: pause on hover to reduce distraction
    btn.addEventListener('mouseenter', () => clearInterval(intervalId));
})();

// --- FILTER BUTTON LISTENERS ---
document.getElementById('loan-filter-btn').addEventListener('click', renderLoans);
document.getElementById('loan-reset-btn').addEventListener('click', () => {
    document.getElementById('loan-filter-start-date').value = '';
    document.getElementById('loan-filter-end-date').value = '';
    renderLoans();
});
document.getElementById('class-loan-filter-select').addEventListener('change', renderClassLoans);
document.getElementById('class-loan-reset-btn').addEventListener('click', () => {
    document.getElementById('class-loan-filter-select').value = '';
    renderClassLoans();
});
document.getElementById('reading-log-filter-btn').addEventListener('click', renderReadingLogs);
document.getElementById('reading-log-reset-btn').addEventListener('click', () => {
    document.getElementById('reading-log-filter-start-date').value = '';
    document.getElementById('reading-log-filter-end-date').value = '';
    renderReadingLogs();
});

// --- PRINTING ---
// Helper: format current date in Khmer
const formatKhmerDate = (dateObj = new Date()) => {
    const khMonths = ['មករា','កុម្ភះ','មិនា','មេសា','ឧសភា','មិថុនា','កក្កដា','សីហា','កញ្ញា','តុលា','វិច្ឆិកា','ធ្នូ'];
    const d = dateObj.getDate();
    const m = khMonths[dateObj.getMonth()];
    const y = dateObj.getFullYear();
    return `ថ្ងៃទី ${d} ខែ ${m} ឆ្នាំ ${y}`;
};

// Helper: create footer element with Khmer text
const createKhmerFooter = () => {
    const el = document.createElement('div');
    el.className = 'print-footer print-only';

    const left = document.createElement('div');
    const school = (settingsData && (settingsData.schoolName || settingsData.school_name)) || '';
    left.textContent = `ធ្វើនៅ ${school}, ${formatKhmerDate(new Date())}`;

    const right = document.createElement('div');
    right.className = 'print-footer-right';
    right.textContent = 'បណ្ណារក្ស';

    el.appendChild(left);
    el.appendChild(right);
    return el;
};

// Helper: map page id to its print area id
const getPrintAreaForPageId = (pageId) => {
    const map = {
        'page-books': 'books-print-area',
        'page-loans': 'loans-print-area',
        'page-class-loans': 'class-loans-print-area',
        'page-reading-log': 'reading-log-print-area',
        'page-locations': 'locations-print-area',
        'page-students': 'students-print-area',
    };
    const id = map[pageId];
    return id ? document.getElementById(id) : null;
};

const prepareAndPrint = (printClass) => {
    // Ensure school name is shown in header
    const schoolNameSpan = document.getElementById('print-school-name');
    if (schoolNameSpan) {
        schoolNameSpan.textContent = (settingsData && (settingsData.schoolName || settingsData.school_name)) || '';
    }
    document.body.classList.add(printClass);
    window.print();
};

// --- UI TOAST ---
// Lightweight toast for short success/error messages
window.__toastTimer = null;
window.showToast = (message, bgClass = 'bg-green-600') => {
    let el = document.getElementById('app-toast');
    if (!el) {
        el = document.createElement('div');
        el.id = 'app-toast';
        document.body.appendChild(el);
    }
    // Update styles/message each time
    el.className = `${bgClass} text-white px-5 py-3 rounded shadow-lg fixed z-50 transition-opacity duration-200`;
    el.style.opacity = '0';
    el.style.top = '50%';
    el.style.left = '50%';
    el.style.transform = 'translate(-50%, -50%)';
    el.style.pointerEvents = 'none';
    el.textContent = message || '';
    // Show
    requestAnimationFrame(() => { el.style.opacity = '1'; });
    // Clear existing timer and schedule hide
    if (window.__toastTimer) clearTimeout(window.__toastTimer);
    window.__toastTimer = setTimeout(() => { el.style.opacity = '0'; }, 3000);
};

window.printReport = () => {
    const activePage = document.querySelector('.page:not(.hidden)');
    if (activePage) {
        const titleSpan = document.getElementById('print-report-title');
        let title = '';
        const pageId = activePage.id;

        switch (pageId) {
            case 'page-books':
                title = 'បញ្ជីសៀវភៅ';
                break;
            case 'page-loans':
                const startDateLoans = document.getElementById('loan-filter-start-date').value;
                const endDateLoans = document.getElementById('loan-filter-end-date').value;
                title = 'បញ្ជីការខ្ចី-សងបុគ្គល';
                if (startDateLoans && endDateLoans) {
                    const s = formatDateDDMMYYYY(startDateLoans);
                    const e = formatDateDDMMYYYY(endDateLoans);
                    title += ` ពីថ្ងៃ ${s} ដល់ថ្ងៃ ${e}`;
                }
                break;
            case 'page-class-loans':
                title = 'បញ្ជីការខ្ចីតាមថ្នាក់';
                break;
            case 'page-reading-log':
                const startDateLogs = document.getElementById('reading-log-filter-start-date').value;
                const endDateLogs = document.getElementById('reading-log-filter-end-date').value;
                title = 'បញ្ជីឈ្មោះសិស្សចូលអានសរុបតាមថ្ងៃ';
                if (startDateLogs && endDateLogs) {
                    const s = formatDateDDMMYYYY(startDateLogs);
                    const e = formatDateDDMMYYYY(endDateLogs);
                    title += ` ពីថ្ងៃ ${s} ដល់ថ្ងៃ ${e}`;
                }
                break;
            case 'page-locations':
                title = 'បញ្ជីទីតាំងរក្សាទុកសៀវភៅ';
                break;
            case 'page-students':
                const academicYear = settingsData.academicYear || '';
                const studentClassFilter = document.getElementById('student-class-filter');
                const selectedClass = studentClassFilter.value;
                
                title = 'បញ្ជីឈ្មោះសិស្ស';
                
                if (selectedClass) {
                    title += ` ថ្នាក់ទី ${selectedClass}`;
                }

                if (academicYear) {
                    title += ` ឆ្នាំសិក្សា ${academicYear}`;
                }
                break;
        }
        titleSpan.textContent = title;

        // Inject Khmer footer into the correct print area
        try {
            // Remove any previous dynamic footers to avoid duplicates
            document.querySelectorAll('.print-footer').forEach(el => el.remove());
            const area = getPrintAreaForPageId(pageId);
            if (area) {
                area.appendChild(createKhmerFooter());
            }
        } catch (e) { console.warn('Footer injection failed:', e); }

        prepareAndPrint(`printing-${pageId}`);
    }
};

// NEW: Print grouped-by-book reading log summary
window.printReadingLogGrouped = () => {
    // Only applicable on Reading Log page
    const activePage = document.querySelector('.page:not(.hidden)');
    if (!activePage || activePage.id !== 'page-reading-log') return;

    const titleSpan = document.getElementById('print-report-title');
    const startDate = document.getElementById('reading-log-filter-start-date')?.value;
    const endDate = document.getElementById('reading-log-filter-end-date')?.value;
    let title = 'សរុបតាមសៀវភៅ';
    if (startDate && endDate) {
        const s = formatDateDDMMYYYY(startDate);
        const e = formatDateDDMMYYYY(endDate);
        title += ` ពីថ្ងៃ ${s} ដល់ថ្ងៃ ${e}`;
    }
    if (titleSpan) titleSpan.textContent = title;

    // Inject Khmer footer into grouped area, not the detailed table area
    try {
        // Remove any previous dynamic footers to avoid duplicates
        document.querySelectorAll('.print-footer').forEach(el => el.remove());
        const area = document.getElementById('reading-log-grouped-area');
        if (area) {
            area.appendChild(createKhmerFooter());
        }
    } catch (e) { console.warn('Footer injection failed (grouped):', e); }

    // Insert the required upper title above the existing title (print-only)
    try {
        if (titleSpan && titleSpan.parentElement) {
            // Remove any existing dynamic upper title to avoid duplicates
            const oldUpper = document.getElementById('print-upper-title');
            if (oldUpper) oldUpper.remove();

            const upper = document.createElement('div');
            upper.id = 'print-upper-title';
            upper.className = 'print-only print-upper-title';
            upper.textContent = 'បញ្ជីឈ្មោះសិស្សចូលអាន';
            // Place above the existing report title span
            titleSpan.parentElement.insertBefore(upper, titleSpan);
        }
    } catch (e) { /* no-op */ }

    // Add a temporary class to hide the detailed table during grouped printing
    const tempClass = 'printing-reading-log-grouped';
    document.body.classList.add(tempClass);
    const cleanup = () => {
        document.body.classList.remove(tempClass);
        // Remove the dynamically added upper title
        try {
            const oldUpper = document.getElementById('print-upper-title');
            if (oldUpper) oldUpper.remove();
        } catch (_) { /* noop */ }
    };
    const prevAfter = window.onafterprint;
    window.onafterprint = () => {
        cleanup();
        if (typeof prevAfter === 'function') prevAfter();
        // reset to avoid stacking
        window.onafterprint = prevAfter || null;
    };

    // Reuse the same printing class as reading log page
    prepareAndPrint('printing-page-reading-log');
};

// MODIFIED FUNCTION: Adds a longer delay to ensure all elements are rendered before printing.
window.printCards = () => {
    // Add the printing class to the body
    document.body.classList.add('printing-page-student-cards');
    
    // Use a timeout to allow the browser to render all QR codes and barcodes
    // before opening the print dialog. This fixes issues with missing elements on later pages.
    setTimeout(() => {
        window.print();
    }, 3000); // Increased delay to 3 seconds for larger datasets
};

window.onafterprint = () => {
    const printClasses = Array.from(document.body.classList).filter(
        c => c.startsWith('printing-')
    );
    document.body.classList.remove(...printClasses);
    // Remove dynamically injected footers
    document.querySelectorAll('.print-footer').forEach(el => el.remove());
};

// Navigation function for student cards page - Universal compatibility
function navigateToStudentCards() {
    // Use window.open with _self target for maximum compatibility
    const currentOrigin = window.location.origin;
    const currentPath = window.location.pathname;
    const basePath = currentPath.substring(0, currentPath.lastIndexOf('/') + 1);
    const fullUrl = currentOrigin + basePath + 'stcard.html';
    
    // Open in same tab with full URL for universal compatibility
    window.open(fullUrl, '_self');
}

// NEW FUNCTION: Print a detailed class loan report
document.getElementById('print-class-loan-list-btn').addEventListener('click', () => {
    const selectedClass = document.getElementById('class-loan-filter-select').value;
    if (!selectedClass) {
        if (typeof window.openClassLoanSelectRequiredModal === 'function') {
            window.openClassLoanSelectRequiredModal();
        } else {
            alert('សូមជ្រើសរើសថ្នាក់ដែលត្រូវបោះពុម្ពជាមុនសិន។');
        }
        return;
    }
    
    const classKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('ថ្នាក់')) : null;
    const lastNameKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('នាមត្រកូល')) : null;
    const firstNameKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('នាមខ្លួន')) : null;
    const genderKey = students.length > 0 ? Object.keys(students[0]).find(k => k.includes('ភេទ')) : null;

    if (!classKey || !lastNameKey || !firstNameKey || !genderKey) {
         alert('ទិន្នន័យសិស្សមិនគ្រប់គ្រាន់សម្រាប់បង្កើតរបាយការណ៍ទេ។');
         return;
    }

    const studentsInClass = students.filter(s => s[classKey] === selectedClass);
    // Use linked individual loan rows belonging to class loans for this class
    const classLoansForClass = classLoans.filter(l => l.class_name === selectedClass);
    const relevantClassLoanIds = new Set(classLoansForClass.map(l => String(l.id)));
    const linkedStudentLoans = loans.filter(l => l.class_loan_id && relevantClassLoanIds.has(String(l.class_loan_id)) && l.status === 'ខ្ចី');
    
    // Helper to normalize names (trim, collapse spaces, lowercase)
    const normalizeName = (s) => (s || '').toString().replace(/\s+/g, ' ').trim().toLowerCase();

    // Build student -> title -> count map, with fallback to class loan quantities
    const studentLoanMap = new Map(); // Map<normalized student name, Map<title, count>>
    const bookTitleSet = new Set();
    if (linkedStudentLoans.length > 0) {
        linkedStudentLoans.forEach(loan => {
            const borrowerStr = (loan.borrower || '').trim();
            const studentNameRaw = borrowerStr ? borrowerStr.split(' - ')[0].trim() : '';
            const studentName = normalizeName(studentNameRaw);
            const bookTitle = books.find(b => b.id === loan.book_id)?.title || 'N/A';
            if (!studentName || !bookTitle) return;
            bookTitleSet.add(bookTitle);
            if (!studentLoanMap.has(studentName)) studentLoanMap.set(studentName, new Map());
            const titleMap = studentLoanMap.get(studentName);
            titleMap.set(bookTitle, (titleMap.get(bookTitle) || 0) + 1);
        });
    } else {
        // Fallback: compute from classLoans outstanding quantities
        classLoansForClass.forEach(loan => {
            const borrowerStr = (loan.borrower || '').trim();
            const studentNameRaw = borrowerStr ? borrowerStr.split(' - ')[0].trim() : '';
            const studentName = normalizeName(studentNameRaw);
            const bookTitle = books.find(b => b.id === loan.book_id)?.title || 'N/A';
            const loanedQty = Number(loan.loaned_quantity || 1);
            const returnedQty = Number(loan.returned_count || 0);
            const outstanding = Math.max(0, loanedQty - returnedQty);
            if (!studentName || !bookTitle || outstanding === 0) return;
            bookTitleSet.add(bookTitle);
            if (!studentLoanMap.has(studentName)) studentLoanMap.set(studentName, new Map());
            const titleMap = studentLoanMap.get(studentName);
            titleMap.set(bookTitle, (titleMap.get(bookTitle) || 0) + outstanding);
        });
    }

    const allBookTitles = [...bookTitleSet].sort();

    // Build the table dynamically
    const tableContainer = document.getElementById('class-loan-students-table-container');
    tableContainer.innerHTML = '';
    
    const table = document.createElement('table');
    table.className = 'w-full text-left whitespace-no-wrap';
    
    // Build table header
    const thead = document.createElement('thead');
    thead.className = 'bg-gray-200';
    const headerRow = document.createElement('tr');
    ['ល.រ', 'ឈ្មោះ', 'ភេទ', ...allBookTitles, 'សរុប'].forEach(text => {
        const th = document.createElement('th');
        th.className = 'p-3 text-center';
        th.textContent = text;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Build table body
    const tbody = document.createElement('tbody');
    let totalStudents = 0;
    let totalMales = 0;
    let totalFemales = 0;
    const bookTotals = {};

    studentsInClass.forEach((student, index) => {
        totalStudents++;
        if (student[genderKey] === 'ប្រុស') totalMales++;
        if (student[genderKey] === 'ស្រី') totalFemales++;
        
        const last = (student[lastNameKey] || '').toString();
        const first = (student[firstNameKey] || '').toString();
        const studentFullName = `${last} ${first}`.trim();
        const keyLF = normalizeName(`${last} ${first}`);
        const keyFL = normalizeName(`${first} ${last}`);
        const studentTitleMap = studentLoanMap.get(keyLF) || studentLoanMap.get(keyFL) || new Map();
        
        const row = document.createElement('tr');
        row.className = 'border-b';
        
        const noCell = document.createElement('td');
        noCell.className = 'p-3 text-center';
        noCell.textContent = index + 1;
        row.appendChild(noCell);

        const nameCell = document.createElement('td');
        nameCell.className = 'p-3';
        nameCell.textContent = studentFullName;
        row.appendChild(nameCell);
        
        const genderCell = document.createElement('td');
        genderCell.className = 'p-3 text-center';
        genderCell.textContent = student[genderKey];
        row.appendChild(genderCell);

        let studentBookCount = 0;
        allBookTitles.forEach(bookTitle => {
            const count = studentTitleMap.get(bookTitle) || 0;
            const cell = document.createElement('td');
            cell.className = 'p-3 text-center';
            cell.textContent = count > 0 ? count : '';
            row.appendChild(cell);
            studentBookCount += count;
            bookTotals[bookTitle] = (bookTotals[bookTitle] || 0) + count;
        });
        
        const totalCell = document.createElement('td');
        totalCell.className = 'p-3 text-center font-bold';
        totalCell.textContent = studentBookCount;
        row.appendChild(totalCell);
        
        tbody.appendChild(row);
    });
    table.appendChild(tbody);
    tableContainer.appendChild(table);

    // Add summary directly under the table with custom layout
    const summaryDiv = document.createElement('div');
    summaryDiv.className = 'mt-4 text-sm';
    
    const todayDate = new Date();
    const day = todayDate.getDate();
    const monthNumber = todayDate.getMonth() + 1;
    const year = todayDate.getFullYear();
    
    // Convert month number to Khmer month name
    const khmerMonths = {
        1: 'មករា', 2: 'កុម្ភៈ', 3: 'មីនា', 4: 'មេសា', 
        5: 'ឧសភា', 6: 'មិថុនា', 7: 'កក្កដា', 8: 'សីហា',
        9: 'កញ្ញា', 10: 'តុលា', 11: 'វិច្ឆិកា', 12: 'ធ្នូ'
    };
    const month = khmerMonths[monthNumber] || monthNumber;
    
    // Line 1: Student count on left, school info on right
    const line1 = document.createElement('div');
    line1.className = 'flex justify-between mb-2';
    line1.innerHTML = `
        <div>សិស្សសរុប៖ ${totalStudents} នាក់ (ស្រី៖ ${totalFemales} នាក់)</div>
        <div>ធ្វើនៅ, ${settingsData.schoolName || 'ឈ្មោះសាលា'} ថ្ងៃទី ${day} ខែ ${month} ឆ្នាំ ${year}</div>
    `;
    
    // Line 2: បណ្ណារក្ស positioned 3 tabs from right
    const line2 = document.createElement('div');
    line2.className = 'mb-2';
    line2.style.paddingLeft = '80%';
    line2.innerHTML = `<div>បណ្ណារក្ស</div>`;
    
    // Create book rows in 3 columns
    const bookRows = [];
    for (let i = 0; i < allBookTitles.length; i += 3) {
        const book1 = allBookTitles[i] ? `${allBookTitles[i]}៖ ${bookTotals[allBookTitles[i]] || 0} ក្បាល` : '';
        const book2 = allBookTitles[i + 1] ? `${allBookTitles[i + 1]}៖ ${bookTotals[allBookTitles[i + 1]] || 0} ក្បាល` : '';
        const book3 = allBookTitles[i + 2] ? `${allBookTitles[i + 2]}៖ ${bookTotals[allBookTitles[i + 2]] || 0} ក្បាល` : '';
        
        const bookRow = document.createElement('div');
        bookRow.className = 'mb-1';
        
        // Create book line with single spaces
        const bookLine = [book1, book2, book3].filter(book => book).join(' ');
        
        bookRow.innerHTML = `<div>${bookLine}</div>`;
        bookRows.push(bookRow);
    }
    
    summaryDiv.appendChild(line1);
    summaryDiv.appendChild(line2);
    bookRows.forEach(row => summaryDiv.appendChild(row));
    tableContainer.appendChild(summaryDiv);

    // Set up the print title and subtitle
    const titleSpan = document.getElementById('print-report-title');
    titleSpan.textContent = `បញ្ជីឈ្មោះសិស្សខ្ចីសៀវភៅ`;
    const subtitle = document.getElementById('class-loan-report-subtitle');
    subtitle.textContent = `ថ្នាក់ទី ${selectedClass} ឆ្នាំសិក្សា ${settingsData.academicYear || 'N/A'}`;
    subtitle.classList.remove('hidden');

    // Add delay to ensure DOM is fully rendered before printing
    setTimeout(() => {
        // Do not inject standardized footer for this specific print
        prepareAndPrint('printing-page-class-loan-list');
    }, 500);

     // Clean up after print
    setTimeout(() => {
        subtitle.classList.add('hidden');
        tableContainer.innerHTML = '';
    }, 1000);
});

window.onafterprint = () => {
    const printClasses = Array.from(document.body.classList).filter(
        c => c.startsWith('printing-')
    );
    document.body.classList.remove(...printClasses);
};

// Default page functionality removed

// --- SEARCH EVENT LISTENERS ---
document.getElementById('search-books').addEventListener('input', renderBooks);
document.getElementById('author-filter').addEventListener('change', renderBooks);
document.getElementById('search-loans').addEventListener('input', renderLoans);
document.getElementById('search-locations').addEventListener('input', renderLocations);
document.getElementById('search-students').addEventListener('input', renderStudents);

// --- CLASS FILTER EVENT LISTENER ---
document.getElementById('student-class-filter').addEventListener('change', renderStudents);

// --- EXCEL EXPORT EVENT LISTENER ---
document.getElementById('export-excel-btn').addEventListener('click', exportData);

// --- BOOK DROPDOWN POPULATION ---
const populateBookDropdowns = () => {
    const loanBookSelect = document.getElementById('loan-book-select');
    const classLoanBookSelect = document.getElementById('class-loan-book-select');
    
    // Clear existing options except the first one
    loanBookSelect.innerHTML = '<option value="">-- ជ្រើសរើសសៀវភៅ --</option>';
    classLoanBookSelect.innerHTML = '<option value="">-- ជ្រើសរើសសៀវភៅ --</option>';
    
    // Sort books by title
    const sortedBooks = [...books].sort((a, b) => a.title.localeCompare(b.title));
    
    sortedBooks.forEach(book => {
        // Only show books that have available copies
        const loanedCount = loans.filter(loan => loan.book_id === book.id && loan.status === 'ខ្ចី').length;
        const remaining = (book.quantity || 0) - loanedCount;
        
        if (remaining > 0) {
            const displayText = `${book.title}${book.author ? ` - ${book.author}` : ''} (នៅសល់: ${remaining})`;
            
            // Add to individual loan dropdown
            const loanOption = document.createElement('option');
            loanOption.value = book.id;
            loanOption.textContent = displayText;
            loanBookSelect.appendChild(loanOption);
            
            // Add to class loan dropdown
            const classLoanOption = document.createElement('option');
            classLoanOption.value = book.id;
            classLoanOption.textContent = displayText;
            classLoanBookSelect.appendChild(classLoanOption);
        }
    });
};

// --- SELECTED BOOKS MANAGEMENT ---
const renderSelectedLoanBooks = () => {
    const container = document.getElementById('loan-selected-books-list');
    const noMessage = document.getElementById('loan-no-books-message');
    
    if (selectedLoanBooks.length === 0) {
        noMessage.style.display = 'block';
        // Hide all book items
        container.querySelectorAll('.selected-book-item').forEach(item => item.remove());
    } else {
        noMessage.style.display = 'none';
        // Clear existing items
        container.querySelectorAll('.selected-book-item').forEach(item => item.remove());
        
        selectedLoanBooks.forEach((book, index) => {
            const bookItem = document.createElement('div');
            bookItem.className = 'selected-book-item flex items-center justify-between bg-white p-2 rounded border mb-2';
            bookItem.innerHTML = `
                <div class="flex-1">
                    <span class="font-medium">${book.title}</span>
                    ${book.author ? `<span class="text-gray-600"> - ${book.author}</span>` : ''}
                    <span class="text-sm text-blue-600 ml-2">(នៅសល់: ${book.remaining})</span>
                </div>
                <button type="button" onclick="removeSelectedLoanBook(${index})" class="text-red-500 hover:text-red-700 ml-2">
                    <i class="fas fa-times"></i>
                </button>
            `;
            container.appendChild(bookItem);
        });
    }
};

const addSelectedLoanBook = (bookId) => {
    const book = books.find(b => String(b.id) === String(bookId));
    if (!book) return false;
    
    // Check if book is already selected
    if (selectedLoanBooks.find(b => String(b.id) === String(bookId))) {
        if (typeof window.openLoanDuplicateModal === 'function') {
            window.openLoanDuplicateModal();
        } else {
            alert('សៀវភៅនេះត្រូវបានជ្រើសរើសរួចហើយ។');
        }
        return false;
    }
    
    // Calculate remaining quantity
    const loanedCount = loans.filter(loan => String(loan.book_id) === String(book.id) && loan.status === 'ខ្ចី').length;
    const remaining = (book.quantity || 0) - loanedCount;
    
    if (remaining <= 0) {
        alert('សៀវភៅនេះអស់ស្តុកហើយ។');
        return false;
    }
    
    selectedLoanBooks.push({
        id: book.id,
        title: book.title,
        author: book.author,
        isbn: book.isbn,
        remaining: remaining
    });
    
    renderSelectedLoanBooks();
    return true;
};

window.removeSelectedLoanBook = (index) => {
    selectedLoanBooks.splice(index, 1);
    renderSelectedLoanBooks();
};

const clearSelectedLoanBooks = () => {
    selectedLoanBooks = [];
    renderSelectedLoanBooks();
};

// Duplicate Selected Book Info Modal controls
window.openLoanDuplicateModal = () => {
    const modal = document.getElementById('loan-duplicate-modal');
    if (modal) modal.classList.remove('hidden');
};

window.closeLoanDuplicateModal = () => {
    const modal = document.getElementById('loan-duplicate-modal');
    if (modal) modal.classList.add('hidden');
};
// Class Selection Required Info Modal controls (Class Loans)
window.openClassLoanSelectRequiredModal = () => {
    const modal = document.getElementById('class-loan-select-required-modal');
    if (modal) modal.classList.remove('hidden');
};

window.closeClassLoanSelectRequiredModal = () => {
    const modal = document.getElementById('class-loan-select-required-modal');
    if (modal) modal.classList.add('hidden');
};

// --- BOOK DROPDOWN EVENT HANDLERS ---
document.getElementById('loan-book-select').addEventListener('change', function() {
    const selectedBookId = this.value;
    if (selectedBookId) {
        if (addSelectedLoanBook(selectedBookId)) {
            // Reset dropdown after successful addition
            this.value = '';
            // Focus back to scan textbox for rapid entry
            setTimeout(() => {
                const isbnInput = document.getElementById('loan-isbn-input');
                if (isbnInput) isbnInput.focus();
            }, 50);
        }
    }
});

document.getElementById('class-loan-book-select').addEventListener('change', function() {
    const selectedBookId = this.value;
    if (selectedBookId) {
        const selectedBook = books.find(book => String(book.id) === String(selectedBookId));
        if (selectedBook) {
            const classLoanBookError = document.getElementById('class-loan-book-error');
            classLoanBookError.textContent = '';

            // Prevent duplicate
            if (currentClassLoanScannedBooks.some(b => String(b.id) === String(selectedBook.id))) {
                classLoanBookError.textContent = 'សៀវភៅនេះបានស្កេនរួចហើយ';
            } else {
                // Compute remaining
                const loanedCount = loans.filter(loan => String(loan.book_id) === String(selectedBook.id) && loan.status === 'ខ្ចី').length;
                const remaining = (selectedBook.quantity || 0) - loanedCount;

                // Add to scanned list and render with remove buttons
                currentClassLoanScannedBooks.push({ id: selectedBook.id, title: selectedBook.title, remaining });
                renderClassLoanScannedBooks();

                // Indicate added
                document.getElementById('class-loan-book-title-display').value = `${selectedBook.title} (បានបន្ថែម)`;
            }

            // Reset dropdown and focus back to scan
            this.value = '';
            setTimeout(() => {
                const input = document.getElementById('class-loan-isbn-input');
                if (input) input.focus();
            }, 50);
        }
    } else {
        // Clear fields when no book is selected
        document.getElementById('class-loan-isbn-input').value = '';
        document.getElementById('class-loan-book-title-display').value = '';
        document.getElementById('class-loan-book-id').value = '';
    }
});
