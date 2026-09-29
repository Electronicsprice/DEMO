// Credentials
const AUTH_USER = "QGE@1983";
const AUTH_PASS = "QGE@1122";

// Global Data State
let rawRatesData = [];
let filteredData = [];

// DOM Elements
const loginScreen = document.getElementById('login-screen');
const appScreen = document.getElementById('app');
const loginForm = document.getElementById('login-form');
const loginError = document.getElementById('login-error');
const logoutBtn = document.getElementById('logout-btn');

const searchInput = document.getElementById('search-input');
const companyFilter = document.getElementById('company-filter');
const productFilter = document.getElementById('product-filter');
const periodFilter = document.getElementById('period-filter');
const cardsGrid = document.getElementById('cards-grid');
const noResults = document.getElementById('no-results');
const resultsCount = document.getElementById('results-count');
const resetBtn = document.getElementById('reset-filters');
const mainHeader = document.getElementById('main-header');

// INITIALIZATION
document.addEventListener('DOMContentLoaded', () => {
  checkAuth();
  setupHeaderScroll();
});

// AUTHENTICATION LOGIC
loginForm.addEventListener('submit', (e) => {
  e.preventDefault();
  const u = document.getElementById('username').value.trim();
  const p = document.getElementById('password').value.trim();

  if (u === AUTH_USER && p === AUTH_PASS) {
    localStorage.setItem('sales_portal_auth', 'true');
    loginError.textContent = '';
    showApp();
  } else {
    loginError.textContent = 'Invalid Username or Password!';
  }
});

logoutBtn.addEventListener('click', () => {
  localStorage.removeItem('sales_portal_auth');
  location.reload();
});

function checkAuth() {
  if (localStorage.getItem('sales_portal_auth') === 'true') {
    showApp();
  }
}

function showApp() {
  loginScreen.classList.add('hidden');
  appScreen.classList.remove('hidden');
  loadExcelData();
}

// FETCH & PARSE RATES.XLSX FROM GITHUB REPO
async function loadExcelData() {
  try {
    const response = await fetch('rates.xlsx?v=' + new Date().getTime()); // Prevent caching
    if (!response.ok) throw new Error("File rates.xlsx not found");
    
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Parse sheet to JSON
    rawRatesData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
    
    populateFilters(rawRatesData);
    applyFilters();
  } catch (err) {
    console.error("Error loading rates data:", err);
    cardsGrid.innerHTML = `<div style="color:var(--primary-red); text-align:center; grid-column:1/-1; padding: 40px;">
      <i class="fa-solid fa-triangle-exclamation" style="font-size:32px; margin-bottom:10px;"></i><br>
      Failed to load 'rates.xlsx'. Make sure the file is uploaded in your GitHub repository root.
    </div>`;
  }
}

// POPULATE DROPDOWNS DYNAMICALLY
function populateFilters(data) {
  const companies = new Set();
  const products = new Set();
  const periods = new Set();

  data.forEach(item => {
    if (item.Company) companies.add(item.Company.toString().trim());
    if (item.Product) products.add(item.Product.toString().trim());
    
    const m = item.Month ? item.Month.toString().trim() : '';
    const y = item.Year ? item.Year.toString().trim() : '';
    if (m || y) periods.add(`${m} ${y}`.trim());
  });

  // Populate Company Dropdown
  companyFilter.innerHTML = '<option value="">All Companies</option>';
  Array.from(companies).sort().forEach(c => {
    companyFilter.innerHTML += `<option value="${c}">${c}</option>`;
  });

  // Populate Product Dropdown
  productFilter.innerHTML = '<option value="">All Products</option>';
  Array.from(products).sort().forEach(p => {
    productFilter.innerHTML += `<option value="${p}">${p}</option>`;
  });

  // Populate Period Dropdown
  periodFilter.innerHTML = '<option value="">All Months/Years</option>';
  Array.from(periods).sort().forEach(pr => {
    periodFilter.innerHTML += `<option value="${pr}">${pr}</option>`;
  });
}

// FILTERING ENGINE
function applyFilters() {
  const q = searchInput.value.toLowerCase().trim();
  const selectedComp = companyFilter.value;
  const selectedProd = productFilter.value;
  const selectedPeriod = periodFilter.value;

  filteredData = rawRatesData.filter(item => {
    const comp = (item.Company || '').toString();
    const prod = (item.Product || '').toString();
    const model = (item.Model || '').toString();
    const month = (item.Month || '').toString();
    const year = (item.Year || '').toString();
    const period = `${month} ${year}`.trim();

    // Text Search
    const matchesSearch = !q || 
      comp.toLowerCase().includes(q) || 
      prod.toLowerCase().includes(q) || 
      model.toLowerCase().includes(q);

    // Dropdown filters
    const matchesComp = !selectedComp || comp === selectedComp;
    const matchesProd = !selectedProd || prod === selectedProd;
    const matchesPeriod = !selectedPeriod || period === selectedPeriod;

    return matchesSearch && matchesComp && matchesProd && matchesPeriod;
  });

  renderCards(filteredData);
}

// RENDER CARDS TO DOM
function renderCards(data) {
  cardsGrid.innerHTML = '';
  resultsCount.textContent = `Showing ${data.length} products`;

  if (data.length === 0) {
    noResults.classList.remove('hidden');
    return;
  }
  noResults.classList.add('hidden');

  data.forEach(item => {
    const company = (item.Company || '').toString().trim();
    const isHaier = company.toUpperCase() === 'HAIER';
    const remarks = item.Remarks ? item.Remarks.toString().trim() : '';

    const card = document.createElement('div');
    card.className = `rate-card ${isHaier ? 'haier-card' : ''}`;

    // Helper for currency formatting
    const formatRate = (val) => {
      if (!val) return 'N/A';
      return isNaN(val) ? val : Number(val).toLocaleString();
    };

    let ratesHtml = `
      <div class="rate-row">
        <span class="rate-label">Cash Rate:</span>
        <span class="rate-val cash">Rs. ${formatRate(item.Cash_Rate)}</span>
      </div>
      <div class="rate-row">
        <span class="rate-label">Installment:</span>
        <span class="rate-val installment">Rs. ${formatRate(item.Installment_Rate)}</span>
      </div>
    `;

    // Special Requirement: HAIER Fix Rate
    if (isHaier) {
      ratesHtml += `
        <div class="rate-row">
          <span class="rate-label">Fix Rate:</span>
          <span class="rate-val fix">Rs. ${formatRate(item.Fix_Rate)}</span>
        </div>
      `;
    }

    // Remarks logic: Show ONLY if remarks exist
    let remarksHtml = '';
    if (remarks !== '') {
      remarksHtml = `
        <div class="remarks-box">
          <i class="fa-solid fa-circle-info"></i> ${remarks}
        </div>
      `;
    }

    card.innerHTML = `
      <div>
        <div class="card-header-top">
          <span class="company-badge">${company}</span>
          <span class="period-badge">${item.Month || ''} ${item.Year || ''}</span>
        </div>
        <div class="product-category">${item.Product || 'General'}</div>
        <div class="model-name">${item.Model || 'Standard Model'}</div>
        <div class="rates-container">
          ${ratesHtml}
        </div>
      </div>
      ${remarksHtml}
    `;

    cardsGrid.appendChild(card);
  });
}

// EVENT LISTENERS FOR FILTERS
searchInput.addEventListener('input', applyFilters);
companyFilter.addEventListener('change', applyFilters);
productFilter.addEventListener('change', applyFilters);
periodFilter.addEventListener('change', applyFilters);

resetBtn.addEventListener('click', () => {
  searchInput.value = '';
  companyFilter.value = '';
  productFilter.value = '';
  periodFilter.value = '';
  applyFilters();
});

// AUTO-HIDING HEADER ON SCROLL DOWN
function setupHeaderScroll() {
  let lastScrollTop = 0;
  window.addEventListener('scroll', () => {
    const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
    if (scrollTop > lastScrollTop && scrollTop > 60) {
      // Scroll Down -> Hide Header
      mainHeader.classList.add('header-hidden');
    } else {
      // Scroll Up -> Show Header
      mainHeader.classList.remove('header-hidden');
    }
    lastScrollTop = scrollTop <= 0 ? 0 : scrollTop;
  });
}