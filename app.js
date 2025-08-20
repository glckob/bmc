// Test if script is loading
console.log('App.js module started loading...');

// Import Supabase
import { createClient } from 'https://cdn.skypack.dev/@supabase/supabase-js@2'

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
let currentUserId = null;
let unsubscribeBooks = () => {};
let unsubscribeLoans = () => {};
let unsubscribeClassLoans = () => {};
let unsubscribeLocations = () => {};
let unsubscribeStudents = () => {};
let unsubscribeSettings = () => {};
let unsubscribeReadingLogs = () => {};
let currentScannedBooks = [];
let currentStudentGender = ''; // For gender tracking

const authContainer = document.getElementById('auth-container');
const appContainer = document.getElementById('app-container');
const loadingOverlay = document.getElementById('loading-overlay');
const loginForm = document.getElementById('login-form');
const showLoginBtn = document.getElementById('show-login');
const logoutBtn = document.getElementById('logout-btn');
const sidebarSchoolName = document.getElementById('sidebar-school-name');

const pages = document.querySelectorAll('.page');
const navLinks = document.querySelectorAll('.nav-link');

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
        
        // Navigate to home page by default
        navigateTo('home');
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

    // Auto-focus or setup for specific pages
    if (pageId === 'class-loans') {
        populateClassLoanForm();
        setTimeout(() => document.getElementById('class-loan-isbn-input').focus(), 100);
    }
    if (pageId === 'loans') {
        window.clearLoanForm();
    }
    if (pageId === 'reading-log') {
        window.clearReadingLogForm();
        setTimeout(() => document.getElementById('reading-log-student-id').focus(), 100);
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

// --- RENDERING FUNCTIONS ---
const renderAll = () => {
    renderBooks();
    renderLoans();
    renderClassLoans();
    renderLocations();
    renderStudents();
    renderReadingLogs();
    updateDashboard();
};

const renderBooks = () => {
    const bookList = document.getElementById('book-list');
    const searchBooksInput = document.getElementById('search-books');
    const searchTerm = searchBooksInput.value.toLowerCase();
    const filteredBooks = books.filter(book => book.title.toLowerCase().includes(searchTerm) || (book.author && book.author.toLowerCase().includes(searchTerm)) || (book.isbn && book.isbn.toLowerCase().includes(searchTerm)));
    bookList.innerHTML = '';
    if (filteredBooks.length === 0) { bookList.innerHTML = `<tr><td colspan="9" class="text-center p-4 text-gray-500">រកមិនឃើញសៀវភៅទេ។</td></tr>`; return; }
    const sortedBooks = [...filteredBooks].sort((a,b) => a.title.localeCompare(b.title));
    sortedBooks.forEach((book, index) => {
        const loanedCount = loans.filter(loan => loan.book_id === book.id && loan.status === 'ខ្ចី').length;
        const remaining = (book.quantity || 0) - loanedCount;
        const location = locations.find(loc => loc.id === book.location_id);
        const row = document.createElement('tr');
        row.className = 'border-b';
        row.innerHTML = `
            <td class="p-3">${index + 1}</td>
            <td class="p-3">${book.title} ${loanedCount > 0 ? `<span class="text-xs bg-yellow-200 text-yellow-800 rounded-full px-2 py-1 ml-2 no-print">ខ្ចី ${loanedCount}</span>` : ''}</td>
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

    // Gender Summary
    let maleCount = 0;
    let femaleCount = 0;
    dateFilteredLoans.forEach(loan => {
        if (loan.borrower_gender === 'ប្រុស' || loan.borrower_gender === 'M') maleCount++;
        if (loan.borrower_gender === 'ស្រី' || loan.borrower_gender === 'F') femaleCount++;
    });
    loanSummary.textContent = `សរុប: ${dateFilteredLoans.length} នាក់ (ប្រុស: ${maleCount} នាក់, ស្រី: ${femaleCount} នាក់)`;

    // Search Term Filtering
    const searchTerm = searchLoansInput.value.toLowerCase();
    const filteredLoans = dateFilteredLoans.filter(loan => { 
        const book = books.find(b => b.id === loan.book_id); 
        return (loan.borrower && loan.borrower.toLowerCase().includes(searchTerm)) || (book && book.title.toLowerCase().includes(searchTerm)); 
    });
    
    loanList.innerHTML = '';
    if (filteredLoans.length === 0) { loanList.innerHTML = `<tr><td colspan="7" class="text-center p-4 text-gray-500">រកមិនឃើញកំណត់ត្រាខ្ចីទេ។</td></tr>`; return; }
    
    const sortedLoans = [...filteredLoans].sort((a,b) => new Date(b.loan_date) - new Date(a.loan_date));
    sortedLoans.forEach((loan, index) => {
        const book = books.find(b => b.id === loan.book_id);
        const row = document.createElement('tr');
        row.className = 'border-b';
        row.innerHTML = `<td class="p-3">${index + 1}</td><td class="p-3">${book ? book.title : 'សៀវភៅត្រូវបានលុប'}</td><td class="p-3">${loan.borrower}</td><td class="p-3">${loan.loan_date}</td><td class="p-3">${loan.return_date || 'N/A'}</td><td class="p-3"><span class="px-2 py-1 text-xs rounded-full ${loan.status === 'ខ្ចី' ? 'bg-yellow-200 text-yellow-800' : 'bg-green-200 text-green-800'}">${loan.status}</span></td><td class="p-3 no-print">${loan.status === 'ខ្ចី' ? `<button onclick="window.returnBook('${loan.id}')" class="text-green-500 hover:text-green-700 mr-2" title="សម្គាល់ថាសងវិញ"><i class="fas fa-undo"></i></button>` : ''}<button onclick="window.deleteLoan('${loan.id}')" class="text-red-500 hover:text-red-700" title="លុបកំណត់ត្រា"><i class="fas fa-trash"></i></button></td>`;
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
    
    const sortedClassLoans = [...filtered].sort((a,b) => new Date(b.loan_date) - new Date(a.loan_date));
    sortedClassLoans.forEach((loan, index) => {
        const book = books.find(b => b.id === loan.book_id);
        const row = document.createElement('tr');
        row.className = 'border-b';
        const isFullyReturned = (loan.returned_count || 0) >= loan.loaned_quantity;
        row.innerHTML = `
            <td class="p-3">${index + 1}</td>
            <td class="p-3">${book ? book.title : 'សៀវភៅត្រូវបានលុប'}</td>
            <td class="p-3">${loan.class_name}</td>
            <td class="p-3">${loan.loan_date}</td>
            <td class="p-3">
                <span class="font-bold ${isFullyReturned ? 'text-green-600' : 'text-orange-600'}">
                    សងបាន: ${loan.returned_count || 0} / ${loan.loaned_quantity}
                </span>
            </td>
            <td class="p-3 no-print">
                ${!isFullyReturned ? `<button onclick="window.openClassReturnModal('${loan.id}')" class="text-teal-500 hover:text-teal-700 mr-2" title="សងសៀវភៅ"><i class="fas fa-book-reader"></i></button>` : ''}
                <button onclick="window.openClassLoanEditModal('${loan.id}')" class="text-blue-500 hover:text-blue-700 mr-2" title="កែប្រែ"><i class="fas fa-edit"></i></button>
                <button onclick="window.deleteClassLoan('${loan.id}')" class="text-red-500 hover:text-red-700" title="លុបប្រវត្តិនេះ"><i class="fas fa-trash"></i></button>
            </td>
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
                    html += `<li class="ml-4">- ល.រ: ${student[noKey] || 'N/A'}, អត្តលេខ: ${student[idKey]}, ឈ្មោះ: ${fullName}</li>`;
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
    const headers = ['ល.រ', 'អត្តលេខ', 'នាមត្រកូល', 'នាមខ្លួន', 'ភេទ', 'ថ្នាក់', 'ថ្ងៃខែឆ្នាំកំណើត', 'រូបថត URL', 'សកម្មភាព'];
    
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

    const sortedLogs = [...filtered].sort((a,b) => new Date(b.date_time) - new Date(a.date_time));
    sortedLogs.forEach((log, index) => {
        const row = document.createElement('tr');
        row.className = 'border-b';
        let booksRead = 'N/A';
        if (log.books) {
            try {
                // Handle both JSON string and array formats
                const booksData = typeof log.books === 'string' ? JSON.parse(log.books) : log.books;
                if (Array.isArray(booksData)) {
                    booksRead = booksData.map(b => b.title || b).join(', ');
                } else {
                    booksRead = booksData.title || booksData;
                }
            } catch (e) {
                // If parsing fails, try to display as string
                booksRead = log.books.toString();
            }
        }
        row.innerHTML = `
            <td class="p-3">${index + 1}</td>
            <td class="p-3">${new Date(log.date_time).toLocaleString('en-GB')}</td>
            <td class="p-3">${log.student_name}</td>
            <td class="p-3">${booksRead}</td>
            <td class="p-3 no-print">
                <button onclick="window.deleteReadingLog('${log.id}')" class="text-red-500 hover:text-red-700" title="លុប"><i class="fas fa-trash"></i></button>
            </td>
        `;
        readingLogHistory.appendChild(row);
    });
};

const updateDashboard = () => {
    document.getElementById('total-books').textContent = books.length;
    const totalQuantity = books.reduce((sum, book) => sum + (parseInt(book.quantity, 10) || 0), 0);
    document.getElementById('total-quantity').textContent = totalQuantity;
    const activeLoans = loans.filter(l => l.status === 'ខ្ចី').length;
    document.getElementById('total-loans').textContent = activeLoans;
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
    if (!error) {
        books = data || [];
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
    const { data, error } = await supabase
        .from('students')
        .select('*')
        .eq('user_id', userId);
    if (!error) {
        // Normalize to Khmer keys expected by the UI
        const rows = data || [];
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
            alert('បានរក្សាទុកទំព័រលំនាំដើម។ ទំព័រនេះនឹងត្រូវបានប្រើជាទំព័រដើមនៅពេលចូលប្រើប្រាស់លើកក្រោយ។');
        } else {
            console.error("Error saving default page: ", error);
            alert('ការរក្សាទុកបានបរាជ័យ។');
        }
    } catch (e) {
        console.error("Error saving default page: ", e);
        alert('ការរក្សាទុកបានបរាជ័យ។');
    }
};

// --- GOOGLE SHEET & STUDENT DATA ---
document.getElementById('save-url-btn').addEventListener('click', async () => {
    if (!currentUserId) return;
    const url = document.getElementById('google-sheet-url').value.trim();
    
    if (!url) { alert('សូមបញ្ចូល Link ជាមុនសិន'); return; }
    
    // Validate Google Sheets URL format
    if (!url.includes('docs.google.com/spreadsheets/')) {
        alert('Link មិនត្រឹមត្រូវ។ សូមប្រាកដថាជា Google Sheets Link។');
        return;
    }
    
    try {
        const { error } = await supabase
            .from('settings')
            .upsert({
                user_id: currentUserId,
                key: 'google_sheet_url',
                value: url
            }, {
                onConflict: 'user_id,key'
            });
        
        if (!error) {
            alert('រក្សាទុក Link បានជោគជ័យ!');
            // Update local settings
            settingsData.google_sheet_url = url;
        } else {
            console.error("Error saving URL: ", error);
            alert('ការរក្សាទុក Link បានបរាជ័យ។ Error: ' + error.message);
        }
    } catch (e) { 
        console.error("Error saving URL: ", e); 
        alert('ការរក្សាទុក Link បានបរាជ័យ។ Error: ' + e.message); 
    }
});

document.getElementById('fetch-data-btn').addEventListener('click', async () => {
    const url = document.getElementById('google-sheet-url').value.trim();
    if (!url) { alert('សូមបញ្ចូល Link Google Sheet ជាមុនសិន។'); return; }
    
    // Convert regular Google Sheets URL to CSV export URL if needed
    let csvUrl = url;
    if (url.includes('docs.google.com/spreadsheets/') && !url.includes('/pub?output=csv')) {
        // Extract sheet ID from URL
        const sheetIdMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (sheetIdMatch) {
            csvUrl = `https://docs.google.com/spreadsheets/d/${sheetIdMatch[1]}/pub?output=csv`;
        } else {
            alert('Link មិនត្រឹមត្រូវ។ សូមប្រាកដថាជា Google Sheets Link។');
            return;
        }
    }
    
    if (!csvUrl.includes('/pub?output=csv')) { 
        alert('Link មិនត្រឹមត្រូវ។ សូមប្រាកដថា Link បាន Publish ជា CSV ឬប្រើ Link ធម្មតា។'); 
        return; 
    }

    loadingOverlay.classList.remove('hidden');
    try {
        // Try multiple proxy services for better reliability
        const proxyUrls = [
            'https://api.allorigins.win/raw?url=',
            'https://cors-anywhere.herokuapp.com/',
            'https://api.codetabs.com/v1/proxy?quest='
        ];
        
        let response = null;
        let lastError = null;
        
        for (const proxyUrl of proxyUrls) {
            try {
                console.log(`Trying proxy: ${proxyUrl}`);
                response = await fetch(proxyUrl + encodeURIComponent(csvUrl), {
                    headers: {
                        'X-Requested-With': 'XMLHttpRequest'
                    }
                });
                if (response.ok) break;
            } catch (error) {
                console.log(`Proxy ${proxyUrl} failed:`, error);
                lastError = error;
                continue;
            }
        }
        
        if (!response || !response.ok) { 
            throw new Error(`All proxy services failed. Last error: ${lastError?.message || 'Network error'}`); 
        }
        
        const csvText = await response.text();
        console.log('CSV data received:', csvText.substring(0, 200) + '...');
        
        const parsedData = parseCSV(csvText);
        console.log('Parsed data:', parsedData.slice(0, 3));
        
        if (parsedData.length === 0) { 
            alert('មិនអាចញែកទិន្នន័យពី CSV បានទេ ឬក៏ Sheet មិនមានទិន្នន័យ។'); 
            return; 
        }
        
        await syncStudentsToSupabase(parsedData);
        alert(`បានទាញយក និងរក្សាទុកទិន្នន័យសិស្ស ${parsedData.length} នាក់ដោយជោគជ័យ។`);
        
    } catch (error) { 
        console.error('Failed to fetch or process student data:', error); 
        alert('ការទាញយកទិន្នន័យបានបរាជ័យ។\n\nបញ្ហាអាចមកពី:\n- Link មិនត្រឹមត្រូវ\n- Sheet មិនបាន Publish to web\n- ការតភ្ជាប់ Internet\n- CORS policy\n\nError: ' + error.message);
    } finally { 
        loadingOverlay.classList.add('hidden'); 
    }
});

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

// --- BOOK MANAGEMENT ---
window.openBookModal = (id = null) => {
    const bookForm = document.getElementById('book-form');
    bookForm.reset(); document.getElementById('book-id').value = '';
    const locationSelect = document.getElementById('book-location-id');
    locationSelect.innerHTML = '<option value="">-- សូមជ្រើសរើសទីតាំង --</option>';
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
        }
    } else { document.getElementById('book-modal-title').textContent = 'បន្ថែមសៀវភៅថ្មី'; }
    document.getElementById('book-modal').classList.remove('hidden');
};
window.closeBookModal = () => document.getElementById('book-modal').classList.add('hidden');
window.editBook = (id) => openBookModal(id);
window.deleteBook = async (id) => {
    if (!currentUserId) return;
    if (confirm('តើអ្នកពិតជាចង់លុបសៀវភៅនេះមែនទេ?')) {
        const isLoaned = loans.some(loan => loan.book_id === id && loan.status === 'ខ្ចី');
        if (isLoaned) { alert('មិនអាចលុបសៀវភៅនេះបានទេ ព្រោះកំពុងមានគេខ្ចី។'); return; }
        try { 
            const { error } = await supabase
                .from('books')
                .delete()
                .eq('id', id)
                .eq('user_id', currentUserId);
            if (error) console.error("Error deleting document: ", error);
            else {
                await loadBooks(currentUserId);
                renderAll();
            }
        } catch (e) { console.error("Error deleting document: ", e); }
    }
};

document.getElementById('book-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentUserId) return;
    
    const id = document.getElementById('book-id').value;
    const isbnValue = document.getElementById('isbn').value.trim();
    
    const bookData = {
        title: document.getElementById('title').value,
        author: document.getElementById('author').value,
        isbn: isbnValue,
        quantity: parseInt(document.getElementById('quantity').value, 10) || 0,
        location_id: document.getElementById('book-location-id').value,
        source: document.getElementById('source').value,
    };

    if (!bookData.location_id) {
        alert('សូមជ្រើសរើសទីតាំងរក្សាទុក!');
        return;
    }

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
        closeBookModal();
        try {
            await loadBooks(currentUserId);
            renderAll();
        } catch (loadError) {
            console.error("Error loading books after save:", loadError);
            // Don't show error for load failure, book was saved successfully
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
    
    setTimeout(() => {
        const isbnInput = document.getElementById('loan-isbn-input');
        if (isbnInput) {
            isbnInput.focus();
            console.log('Focus set to loan-isbn-input');
        }
    }, 100);
};

document.getElementById('loan-isbn-input').addEventListener('input', () => {
    const isbn = document.getElementById('loan-isbn-input').value.trim();
    const bookIdInput = document.getElementById('loan-book-id');
    const loanBookTitleDisplay = document.getElementById('loan-book-title-display');
    const loanBookError = document.getElementById('loan-book-error');
    loanBookTitleDisplay.value = ''; bookIdInput.value = ''; loanBookError.textContent = '';
    if (!isbn) return;
    const foundBook = books.find(b => b.isbn && b.isbn.toLowerCase() === isbn.toLowerCase());
    if(foundBook) {
        const loanedCount = loans.filter(loan => loan.book_id === foundBook.id && loan.status === 'ខ្ចី').length;
        const remaining = (foundBook.quantity || 0) - loanedCount;
        if (remaining > 0) { 
            loanBookTitleDisplay.value = `${foundBook.title}`; 
            bookIdInput.value = foundBook.id;
            // Auto-focus to student scan field after book is found
            setTimeout(() => document.getElementById('loan-student-id-input').focus(), 100);
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
    const bookId = document.getElementById('loan-book-id').value;
    const borrower = document.getElementById('loan-borrower-name-display').value;
    if (!bookId || !borrower) { 
        alert('សូមបំពេញព័ត៌មានឲ្យបានត្រឹមត្រូវ!'); 
        return; 
    }
    const newLoan = { 
        book_id: bookId, 
        borrower: borrower, 
        borrower_gender: document.getElementById('loan-borrower-gender').value,
        loan_date: new Date().toISOString().split('T')[0], 
        return_date: null, 
        status: 'ខ្ចី',
        user_id: currentUserId
    };
    try { 
        const { error } = await supabase
            .from('loans')
            .insert([newLoan]);
        if (error) throw error;
        await loadLoans(currentUserId);
        renderAll();
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
window.deleteLoan = async (id) => {
    if (!currentUserId) return;
    if (confirm('តើអ្នកពិតជាចង់លុបកំណត់ត្រានេះមែនទេ?')) {
        try { 
            const { error } = await supabase
                .from('loans')
                .delete()
                .eq('id', id)
                .eq('user_id', currentUserId);
            if (error) console.error("Error deleting loan: ", error);
            else {
                await loadLoans(currentUserId);
                renderAll();
            }
        } catch (e) { console.error("Error deleting loan: ", e); }
    }
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

document.getElementById('class-loan-isbn-input').addEventListener('input', () => {
    const isbn = document.getElementById('class-loan-isbn-input').value.trim();
    const bookIdInput = document.getElementById('class-loan-book-id');
    const classLoanBookTitleDisplay = document.getElementById('class-loan-book-title-display');
    const classLoanBookError = document.getElementById('class-loan-book-error');
    classLoanBookTitleDisplay.value = ''; bookIdInput.value = ''; classLoanBookError.textContent = '';
    if (!isbn) return;
    const foundBook = books.find(b => b.isbn && b.isbn.toLowerCase() === isbn.toLowerCase());
    if(foundBook) {
        const loanedCount = loans.filter(loan => loan.book_id === foundBook.id && loan.status === 'ខ្ចី').length;
        const remaining = (foundBook.quantity || 0) - loanedCount;
        classLoanBookTitleDisplay.value = `${foundBook.title} (នៅសល់ ${remaining})`;
        bookIdInput.value = foundBook.id;
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

    const bookId = document.getElementById('class-loan-book-id').value;
    const className = classLoanClassSelect.value;
    const checkedStudentCheckboxes = document.querySelectorAll('#class-loan-student-list-container input[type="checkbox"]:checked');
    const quantity = checkedStudentCheckboxes.length;

    if (!bookId || !className || quantity === 0) {
        alert('សូមជ្រើសរើសសៀវភៅ, ថ្នាក់, និងសិស្សយ៉ាងហោចណាស់ម្នាក់។');
        return;
    }
    
    const selectedBook = books.find(b => b.id === bookId);
    if (!selectedBook) {
        alert('រកមិនឃើញព័ត៌មានសៀវភៅទេ។');
        return;
    }
    const currentlyLoanedCount = loans.filter(l => l.book_id === bookId && l.status === 'ខ្ចី').length;
    const availableCopies = selectedBook.quantity - currentlyLoanedCount;

    if (quantity > availableCopies) {
        alert(`សៀវភៅមិនគ្រប់គ្រាន់! អ្នកចង់ខ្ចី ${quantity} ក្បាល, ប៉ុន្តែនៅសល់តែ ${availableCopies} ក្បាលសម្រាប់ខ្ចី។`);
        return;
    }

    loadingOverlay.classList.remove('hidden');
    try {
        const today = new Date().toISOString().split('T')[0];
        
        // Create class loan record
        const { data: classLoanData, error: classLoanError } = await supabase
            .from('class_loans')
            .insert([{
                book_id: bookId, 
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
        
        // Create individual loan records
        const loanRecords = [];
        checkedStudentCheckboxes.forEach(checkbox => {
            const borrowerText = `${checkbox.dataset.fullName} - ${className}`;
            loanRecords.push({ 
                book_id: bookId, 
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
        await loadLoans(currentUserId);
        await loadClassLoans(currentUserId);
        renderAll();
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
        closeClassReturnModal();
    } catch (err) { console.error("Error updating class loan return: ", err); alert("មានបញ្ហាក្នុងការរក្សាទុកការសង។");
    } finally { loadingOverlay.classList.add('hidden'); }
});

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
window.editLocation = (id) => openLocationModal(id);
window.deleteLocation = async (id) => {
    if (!currentUserId) return;
    if (confirm('តើអ្នកពិតជាចង់លុបទីតាំងនេះមែនទេ?')) {
        const isUsed = books.some(book => book.location_id === id);
        if (isUsed) { alert('មិនអាចលុបទីតាំងនេះបានទេ ព្រោះកំពុងប្រើប្រាស់ដោយសៀវភៅ។'); return; }
        try { 
            const { error } = await supabase
                .from('locations')
                .delete()
                .eq('id', id)
                .eq('user_id', currentUserId);
            if (error) console.error("Error deleting location: ", error);
            else {
                await loadLocations(currentUserId);
                renderAll();
            }
        } catch (e) { console.error("Error deleting location: ", e); }
    }
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
    document.getElementById('student-modal').classList.add('hidden');
};

window.deleteStudent = async (id) => {
    if (!currentUserId) return;
    if (confirm('តើអ្នកពិតជាចង់លុបព័ត៌មានសិស្សនេះមែនទេ?')) {
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
        } catch (e) {
            console.error("Error deleting student: ", e);
            alert("ការលុបបានបរាជ័យ។");
        }
    }
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
            const isbn = isbnInput.value.trim();
            if (!isbn) return;

            const foundBook = books.find(b => b.isbn && b.isbn.toLowerCase() === isbn.toLowerCase());
            if (foundBook) {
                if (!currentScannedBooks.some(b => b.id === foundBook.id)) {
                    currentScannedBooks.push({ id: foundBook.id, title: foundBook.title });
                    const li = document.createElement('li');
                    li.textContent = foundBook.title;
                    document.getElementById('scanned-books-list').appendChild(li);
                }
                isbnInput.value = ''; // Clear input after successful scan
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

document.getElementById('reading-log-isbn-input').addEventListener('input', () => {
    clearTimeout(isbnScanTimer);
    isbnScanTimer = setTimeout(() => {
        const isbnInput = document.getElementById('reading-log-isbn-input');
        const isbn = isbnInput.value.trim();
        if (!isbn) return;
        const foundBook = books.find(b => b.isbn && b.isbn.toLowerCase() === isbn.toLowerCase());
        if (foundBook) {
            if (!currentScannedBooks.some(b => b.id === foundBook.id)) {
                currentScannedBooks.push({ id: foundBook.id, title: foundBook.title });
                const li = document.createElement('li');
                li.textContent = foundBook.title;
                document.getElementById('scanned-books-list').appendChild(li);
            }
            isbnInput.value = '';
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
        window.clearReadingLogForm(); 
    } 
    catch (err) { console.error("Error saving reading log: ", err); alert("ការរក្សាទុកកំណត់ត្រាចូលអានបានបរាជ័យ។"); }
});

window.deleteReadingLog = async (id) => {
     if (!currentUserId) return;
     if (confirm('តើអ្នកពិតជាចង់លុបកំណត់ត្រាចូលអាននេះមែនទេ?')) {
         try { 
            const { error } = await supabase
                .from('reading_logs')
                .delete()
                .eq('id', id)
                .eq('user_id', currentUserId);
            if (error) console.error("Error deleting reading log: ", error);
            else {
                await loadReadingLogs(currentUserId);
                renderAll();
            }
         } catch (e) { console.error("Error deleting reading log: ", e); alert("ការលុបបានបរាជ័យ។"); }
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
            dataToExport = readingLogs.map(log => ({
                'កាលបរិច្ឆេទ': new Date(log.date_time).toLocaleString('en-GB'),
                'ឈ្មោះសិស្ស': log.student_name,
                'ភេទ': log.student_gender,
                'សៀវភៅបានអាន': log.books.map(b => b.title).join('; ')
            }));
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
    document.getElementById('password-confirm-modal').classList.remove('hidden');
    document.getElementById('user-password').focus();
    document.getElementById('password-error').textContent = '';
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
        
        alert(`បានលុបទិន្នន័យទាំងអស់ក្នុង "${collectionName}" ដោយជោគជ័យ។`);
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
        alert('រក្សាទុកឈ្មោះសាលាបានជោគជ័យ។');
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

deleteSchoolNameBtn.addEventListener('click', async () => {
    if (confirm('តើអ្នកពិតជាចង់លុបឈ្មោះសាលាមែនទេ?')) {
        try {
            const { error } = await supabase
                .from('settings')
                .delete()
                .eq('key', 'schoolName')
                .eq('user_id', currentUserId);
            if (error) throw error;
            await loadSettings(currentUserId);
            renderAll();
            alert('បានលុបឈ្មោះសាលា។');
        } catch (e) { console.error(e); alert('ការលុបបានបរាជ័យ។'); }
    }
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
        alert('រក្សាទុកឆ្នាំសិក្សាបានជោគជ័យ។');
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

deleteAcademicYearBtn.addEventListener('click', async () => {
    if (confirm('តើអ្នកពិតជាចង់លុបឆ្នាំសិក្សាមែនទេ?')) {
        try {
            const { error } = await supabase
                .from('settings')
                .delete()
                .eq('key', 'academicYear')
                .eq('user_id', currentUserId);
            if (error) throw error;
            await loadSettings(currentUserId);
            renderAll();
            alert('បានលុបឆ្នាំសិក្សា។');
        } catch (e) { console.error(e); alert('ការលុបបានបរាជ័យ។'); }
    }
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
            alert('រក្សាទុក URL ត្រាបានជោគជ័យ។');
        } else {
            const { error } = await supabase
                .from('settings')
                .delete()
                .eq('key', 'sealImageUrl')
                .eq('user_id', currentUserId);
            if (error) throw error;
            alert('បានលុប URL ត្រា។');
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
            alert('រក្សាទុក URL ផ្ទៃខាងក្រោយកាតបានជោគជ័យ។');
        } else {
            const { error } = await supabase
                .from('settings')
                .delete()
                .eq('key', 'cardBgUrl')
                .eq('user_id', currentUserId);
            if (error) throw error;
            alert('បានលុប URL ផ្ទៃខាងក្រោយកាត។');
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
        changePasswordForm.reset();
    } catch (error) {
        console.error("Password change failed:", error);
        errorP.textContent = 'ការផ្លាស់ប្តូរពាក្យសម្ងាត់បានបរាជ័យ។';
    } finally {
        loadingOverlay.classList.add('hidden');
    }
});


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
const prepareAndPrint = (printClass) => {
    document.body.classList.add(printClass);
    window.print();
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
                    title += ` ពីថ្ងៃ ${startDateLoans} ដល់ថ្ងៃ ${endDateLoans}`;
                }
                break;
            case 'page-class-loans':
                title = 'បញ្ជីការខ្ចីតាមថ្នាក់';
                break;
            case 'page-reading-log':
                const startDateLogs = document.getElementById('reading-log-filter-start-date').value;
                const endDateLogs = document.getElementById('reading-log-filter-end-date').value;
                title = 'បញ្ជីកត់ត្រាការចូលអាន';
                if (startDateLogs && endDateLogs) {
                    title += ` ពីថ្ងៃ ${startDateLogs} ដល់ថ្ងៃ ${endDateLogs}`;
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
        prepareAndPrint(`printing-${pageId}`);
    }
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
        alert('សូមជ្រើសរើសថ្នាក់ដែលត្រូវបោះពុម្ពជាមុនសិន។');
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
    const classLoansForClass = loans.filter(l => l.borrower.includes(selectedClass));
    
    // Group loans by student
    const studentLoanMap = new Map();
    classLoansForClass.forEach(loan => {
        const studentName = loan.borrower.split(' - ')[0];
        if (!studentLoanMap.has(studentName)) {
            studentLoanMap.set(studentName, []);
        }
        studentLoanMap.get(studentName).push(books.find(b => b.id === loan.book_id)?.title || 'N/A');
    });

    // Get all unique book titles
    const allBookTitles = [...new Set(classLoansForClass.map(l => books.find(b => b.id === l.book_id)?.title).filter(Boolean))].sort();

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
        
        const studentFullName = `${student[lastNameKey] || ''} ${student[firstNameKey] || ''}`.trim();
        const loanedBooks = studentLoanMap.get(studentFullName) || [];
        
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
            const count = loanedBooks.filter(title => title === bookTitle).length;
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
    
    // QR code after book summary
    const qrDiv = document.createElement('div');
    qrDiv.className = 'text-center mt-4';
    const reportUrl = `${window.location.origin}${window.location.pathname}?class=${selectedClass}&report=loan`;
    const qrId = `qr-code-container-${Date.now()}`;
    qrDiv.innerHTML = `<div id="${qrId}" class="inline-block"></div>`;
    
    // Generate QR code
    setTimeout(() => {
        const qrContainer = document.getElementById(qrId);
        if (qrContainer && window.QRCode) {
            new QRCode(qrContainer, {
                text: reportUrl,
                width: 80,
                height: 80,
                colorDark: "#000000",
                colorLight: "#ffffff"
            });
        }
    }, 100);
    
    summaryDiv.appendChild(qrDiv);
    tableContainer.appendChild(summaryDiv);

    // Set up the print title and subtitle
    const titleSpan = document.getElementById('print-report-title');
    titleSpan.textContent = `បញ្ជីឈ្មោះសិស្សខ្ចីសៀវភៅ`;
    const subtitle = document.getElementById('class-loan-report-subtitle');
    subtitle.textContent = `ថ្នាក់ទី ${selectedClass} ឆ្នាំសិក្សា ${settingsData.academicYear || 'N/A'}`;
    subtitle.classList.remove('hidden');

    // Add delay to ensure DOM is fully rendered before printing
    setTimeout(() => {
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
document.getElementById('search-loans').addEventListener('input', renderLoans);
document.getElementById('search-locations').addEventListener('input', renderLocations);
document.getElementById('search-students').addEventListener('input', renderStudents);

// --- CLASS FILTER EVENT LISTENER ---
document.getElementById('student-class-filter').addEventListener('change', renderStudents);

// --- EXCEL EXPORT EVENT LISTENER ---
document.getElementById('export-excel-btn').addEventListener('click', exportData);
