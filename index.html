<!DOCTYPE html>
<html lang="bn">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Permissions-Policy" content="unload=(self)">
  <title>ISP বিলিং, এক্সেল ও MikroTik লাইভ মনিটরিং সিস্টেম</title>
  
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <link href="https://fonts.googleapis.com/css2?family=Hind+Siliguri:wght@400;600;700&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <style>
    :root {
      --primary-gradient: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
    }
    body { 
      background-color: #f4f7f6; 
      font-family: 'Hind Siliguri', sans-serif; 
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }
    .content-wrapper { flex: 1; }
    .header-banner { background: var(--primary-gradient); color: white; padding: 20px 0; border-radius: 0 0 20px 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.15); }
    .card-custom { border: none; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
    .stat-card { border-left: 5px solid; }
    .stat-total { border-color: #0d6efd; }
    .stat-paid { border-color: #198754; }
    .stat-unpaid { border-color: #dc3545; }
    .stat-online { border-color: #0dcaf0; }
    .table-container { background: white; border-radius: 12px; padding: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
    .auth-container { max-width: 400px; margin: 60px auto; border-top: 5px solid #0d6efd; }
    .loading-overlay { display: none; position: fixed; top:0; left:0; width:100%; height:100%; background: rgba(0,0,0,0.5); z-index: 9999; justify-content: center; align-items: center; flex-direction: column; color: white; }
    .app-footer { background: var(--primary-gradient); color: white; text-align: center; padding: 12px 0; font-size: 0.9rem; margin-top: auto; }
    .upload-box { border: 2px dashed #0d6efd; border-radius: 10px; padding: 15px; text-align: center; background: #eef5ff; cursor: pointer; transition: 0.3s; }
    .upload-box:hover { background: #dbeaff; }
    .status-badge { font-size: 0.8rem; padding: 4px 10px; border-radius: 15px; }
    .pulse-online {
      display: inline-block;
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background-color: #198754;
      box-shadow: 0 0 0 rgba(25, 135, 84, 0.4);
      animation: pulse 1.5s infinite;
      margin-right: 4px;
    }
    @keyframes pulse {
      0% { box-shadow: 0 0 0 0 rgba(25, 135, 84, 0.7); }
      70% { box-shadow: 0 0 0 6px rgba(25, 135, 84, 0); }
      100% { box-shadow: 0 0 0 0 rgba(25, 135, 84, 0); }
    }
  </style>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ISP Billing & MikroTik Management Dashboard</title>
    <!-- Bootstrap 5 CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SweetAlert2 CSS -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.min.css">
    
    <style>
        body {
            background-color: #f4f6f9;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        .navbar-brand {
            font-weight: 700;
            letter-spacing: 0.5px;
        }
        .card-counter {
            box-shadow: 0 4px 8px rgba(0,0,0,0.05);
            padding: 20px;
            background-color: #fff;
            border-radius: 10px;
            transition: .3s;
        }
        .card-counter:hover {
            box-shadow: 0 8px 16px rgba(0,0,0,0.1);
            transform: translateY(-2px);
        }
        .card-counter i {
            font-size: 2.5em;
            opacity: 0.3;
        }
        .card-counter .count-numbers {
            font-size: 24px;
            font-weight: bold;
        }
        .card-counter .count-name {
            opacity: 0.6;
            text-transform: capitalize;
            font-size: 13px;
            font-weight: 600;
        }
        .status-badge {
            font-size: 0.75rem;
            padding: 5px 10px;
            border-radius: 50px;
        }
        .table-container {
            background: #fff;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.05);
        }
        .search-box {
            max-width: 300px;
        }
    </style>
</head>
<body>

<div id="loader" class="loading-overlay">
  <div class="spinner-border text-light mb-2" style="width: 3rem; height: 3rem;" role="status"></div>
  <h6 id="loaderText" class="fw-bold">ডাটা প্রসেস হচ্ছে, অনুগ্রহ করে অপেক্ষা করুন...</h6>
</div>

<div class="content-wrapper">
  <!-- Auth Section -->
  <div id="authSection" class="container">
    <div class="card card-custom p-4 auth-container bg-white">
      <div class="text-center mb-4">
        <i class="fa-solid fa-network-wired fa-3x text-primary mb-2"></i>
        <h4 class="fw-bold">ISP এডমিন লগইন</h4>
        <p class="text-muted small">বিলিং ও মাইক্রোটিক প্যানেলে প্রবেশ করুন</p>
      </div>
      <form id="authForm">
        <div class="mb-3">
          <label class="form-label fw-semibold">ইমেইল এড্রেস</label>
          <input type="email" id="authEmail" class="form-control" placeholder="admin@isp.com" required>
        </div>
        <div class="mb-3">
          <label class="form-label fw-semibold">পাসওয়ার্ড</label>
          <input type="password" id="authPassword" class="form-control" placeholder="••••••••" required>
        </div>
        <div id="authError" class="alert alert-danger py-2 small mb-3 text-center" style="display: none;"></div>
        <button type="submit" class="btn btn-primary w-100 fw-bold py-2 shadow-sm">প্রবেশ করুন</button>
      </form>
    </div>
  </div>

  <!-- Dashboard Section -->
  <div id="dashboardSection" style="display: none;">
    <div class="header-banner mb-4">
      <div class="container-fluid px-md-4 d-flex justify-content-between align-items-center">
        <div>
          <h3 class="fw-bold mb-0"><i class="fa-solid fa-wifi me-2"></i>ISP বিলিং ও মাইক্রোটিক মনিটরিং</h3>
          <p class="mb-0 opacity-75 small" id="currentMonthLabel">চলতি মাসের বিলিং রেকর্ডস</p>
        </div>
        <div class="d-flex align-items-center gap-2">
          <span id="userDisplayEmail" class="badge bg-light text-dark d-none d-md-inline-block py-2 px-3"></span>
          <button onclick="logout()" class="btn btn-outline-light btn-sm fw-bold">
            <i class="fa-solid fa-power-off me-1"></i> লগআউট
          </button>
        </div>
      </div>
    </div>

    <div class="container-fluid px-md-4">
      <!-- Summary Cards -->
      <div class="row g-3 mb-4 text-center">
        <div class="col-6 col-lg-3">
          <div class="card card-custom p-3 bg-white stat-card stat-total">
            <div class="text-muted small fw-bold">মোট গ্রাহক</div>
            <h4 class="text-primary fw-bold mt-2 mb-0" id="statTotalClients">0</h4>
          </div>
        </div>
        <div class="col-6 col-lg-3">
          <div class="card card-custom p-3 bg-white stat-card stat-paid">
            <div class="text-muted small fw-bold">মোট আদায় (Paid)</div>
            <h4 class="text-success fw-bold mt-2 mb-0" id="statTotalPaid">৳ 0.00</h4>
          </div>
        </div>
        <div class="col-6 col-lg-3">
          <div class="card card-custom p-3 bg-white stat-card stat-unpaid">
            <div class="text-muted small fw-bold">বকেয়া বিল (Unpaid)</div>
            <h4 class="text-danger fw-bold mt-2 mb-0" id="statTotalUnpaid">৳ 0.00</h4>
          </div>
        </div>
        <div class="col-6 col-lg-3">
          <div class="card card-custom p-3 bg-white stat-card stat-online">
            <div class="text-muted small fw-bold">মাইক্রোটিক অনলাইন</div>
            <h4 class="text-info fw-bold mt-2 mb-0" id="statOnlineCount">0 / 0</h4>
          </div>
    <!-- Top Navigation Bar -->
    <nav class="navbar navbar-expand-lg navbar-dark bg-primary sticky-top shadow-sm">
        <div class="container-fluid">
            <a class="navbar-brand" href="#"><i class="fa-solid fa-network-wired me-2"></i>ISP Billing Panel</a>
            <span class="badge bg-light text-primary fw-semibold" id="server-status">Connecting Backend...</span>
        </div>
      </div>

      <div class="row">
        <!-- Sidebar Controls -->
        <div class="col-lg-4 mb-4">
          <!-- Excel Import Card -->
          <div class="card card-custom p-4 mb-3">
            <h5 class="fw-bold text-dark mb-2"><i class="fa-solid fa-file-csv me-2 text-success"></i>Excel / CSV ফাইল আপলোড</h5>
            <p class="text-muted small mb-3">নতুন ফাইল আপলোড করলে আগের ডাটা হারাবে না, নতুন ও পুরাতন ডাটা মার্জ হয়ে যাবে।</p>
            <div class="upload-box" onclick="document.getElementById('excelFileInput').click()">
              <i class="fa-solid fa-cloud-arrow-up fa-2x text-primary mb-2"></i>
              <div class="fw-bold small">ফাইল নির্বাচন করতে ক্লিক করুন</div>
              <div class="text-muted small" id="fileNameDisplay">.csv বা .xlsx ফাইল</div>
    </nav>

    <!-- Main Container -->
    <div class="container-fluid my-4 px-4">
        
        <!-- Summary Cards -->
        <div class="row g-3 mb-4">
            <div class="col-md-3">
                <div class="card-counter border-start border-primary border-4">
                    <div class="d-flex justify-content-between align-items-center">
                        <div>
                            <span class="count-numbers text-primary" id="total-users">0</span>
                            <span class="count-name d-block">মোট গ্রাহক (Total Secrets)</span>
                        </div>
                        <i class="fa-solid fa-users text-primary"></i>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card-counter border-start border-success border-4">
                    <div class="d-flex justify-content-between align-items-center">
                        <div>
                            <span class="count-numbers text-success" id="online-users">0</span>
                            <span class="count-name d-block">অনলাইন গ্রাহক (Active)</span>
                        </div>
                        <i class="fa-solid fa-signal text-success"></i>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card-counter border-start border-secondary border-4">
                    <div class="d-flex justify-content-between align-items-center">
                        <div>
                            <span class="count-numbers text-secondary" id="offline-users">0</span>
                            <span class="count-name d-block">অফলাইন গ্রাহক</span>
                        </div>
                        <i class="fa-solid fa-power-off text-secondary"></i>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card-counter border-start border-danger border-4">
                    <div class="d-flex justify-content-between align-items-center">
                        <div>
                            <span class="count-numbers text-danger" id="disabled-users">0</span>
                            <span class="count-name d-block">ব্লক/ডিজেবলড লাইন</span>
                        </div>
                        <i class="fa-solid fa-ban text-danger"></i>
                    </div>
                </div>
            </div>
            <input type="file" id="excelFileInput" accept=".csv, .xlsx, .xls" style="display: none;" onchange="handleFileUpload(event)">
          </div>

          <!-- Add Single Client Form -->
          <div class="card card-custom p-4">
            <h5 class="fw-bold text-dark mb-3"><i class="fa-solid fa-user-plus me-2 text-primary"></i>নতুন ক্লায়েন্ট এন্ট্রি</h5>
            <form id="clientForm">
              <div class="mb-2">
                <label class="form-label fw-semibold small">MikroTik Username / ID</label>
                <input type="text" id="clientId" class="form-control form-control-sm" placeholder="যেমন: 41younus88" required>
              </div>
              <div class="mb-2">
                <label class="form-label fw-semibold small">গ্রাহকের নাম</label>
                <input type="text" id="clientName" class="form-control form-control-sm" placeholder="যেমন: Younus" required>
              </div>
              <div class="mb-2">
                <label class="form-label fw-semibold small">মোবাইল নম্বর</label>
                <input type="text" id="clientPhone" class="form-control form-control-sm" placeholder="01920060010" required>
              </div>
              <div class="mb-2">
                <label class="form-label fw-semibold small">প্যাকেজ (Package)</label>
                <input type="text" id="clientPackage" class="form-control form-control-sm" placeholder="যেমন: SAM_630TK_I_Vat" required>
              </div>
              <div class="mb-2">
                <label class="form-label fw-semibold small">মাসিক বিল (৳)</label>
                <input type="number" id="billAmount" class="form-control form-control-sm" placeholder="630" min="0" required>
              </div>
              <div class="mb-3">
                <label class="form-label fw-semibold small">বিলিং স্ট্যাটাস</label>
                <select id="paymentStatus" class="form-select form-select-sm" required>
                  <option value="Paid">Paid</option>
                  <option value="Unpaid">Unpaid</option>
                </select>
              </div>
              <button type="submit" class="btn btn-primary w-100 fw-bold shadow-sm py-2">
                <i class="fa-solid fa-cloud-arrow-up me-1"></i> সেভ করুন
              </button>
            </form>
          </div>
        </div>

        <!-- Table Display Area -->
        <div class="col-lg-8">
          <div class="table-container mb-4">
            <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3">
              <input type="text" id="searchInput" onkeyup="filterData()" class="form-control form-control-sm shadow-sm" style="width: 240px;" placeholder="🔍 খুঁজুন (Username, Phone)...">
              <div class="d-flex gap-2">
                <button onclick="fetchMikrotikLiveStatus()" class="btn btn-outline-info btn-sm fw-bold">
                  <i class="fa-solid fa-arrows-rotate me-1"></i> MikroTik Sync
                </button>
                <button onclick="exportExcel()" class="btn btn-outline-success btn-sm fw-bold">
                  <i class="fa-solid fa-file-excel me-1"></i> Excel ডাউনলোড
                </button>
              </div>
        <!-- Table Controls & Search -->
        <div class="table-container">
            <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
                <h5 class="fw-bold mb-0 text-secondary"><i class="fa-solid fa-list me-2"></i>গ্রাহক তালিকা ও লাইভ স্ট্যাটাস</h5>
                <div class="d-flex gap-2">
                    <input type="text" id="searchInput" onkeyup="filterTable()" class="form-control search-box" placeholder="ইউজারনেম বা IP দিয়ে খুঁজুন...">
                    <button class="btn btn-outline-primary" onclick="fetchMikrotikLiveStatus()"><i class="fa-solid fa-rotate me-1"></i> রিফ্রেশ</button>
                </div>
            </div>

            <!-- Client List Table -->
            <div class="table-responsive">
              <table class="table table-hover align-middle" id="clientTable">
                <thead class="table-light">
                  <tr>
                    <th>Username / ID</th>
                    <th>নাম & ফোন</th>
                    <th>প্যাকেজ</th>
                    <th>বিল</th>
                    <th>বিল স্ট্যাটাস</th>
                    <th>লাইভ নেট</th>
                    <th class="text-center">অ্যাকশন</th>
                  </tr>
                </thead>
                <tbody id="tableBody"></tbody>
              </table>
                <table class="table table-hover align-middle" id="clientTable">
                    <thead class="table-light">
                        <tr>
                            <th>#</th>
                            <th>ইউজারনেম (Secret)</th>
                            <th>IP এড্রেস</th>
                            <th>আপটাইম (Uptime)</th>
                            <th>স্ট্যাটাস</th>
                            <th>লাস্ট সিন (Last Seen)</th>
                            <th class="text-center">অ্যাকশন</th>
                        </tr>
                    </thead>
                    <tbody id="clientTableBody">
                        <tr>
                            <td colspan="7" class="text-center py-4 text-muted">
                                <i class="fa-solid fa-spinner fa-spin fa-2x mb-2 d-block"></i>
                                মাইক্রোটিক থেকে ডাটা লোড হচ্ছে...
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</div>

<footer class="app-footer">
  <div class="container">
    <p class="mb-0 fw-semibold">ISP Billing & MikroTik Live Management System</p>
  </div>
</footer>

<script src="https://www.gstatic.com/firebasejs/9.23.0/firebase-app-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/9.23.0/firebase-auth-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/9.23.0/firebase-database-compat.js"></script>

<script>
  const MIKROTIK_API_URL = 'http://localhost:5000/api/mikrotik/users-status';

  // Firebase Configuration
  const firebaseConfig = {
    apiKey: "AIzaSyAfq988qijqN69xjUiZPRdiXlfN3-Dtngs",
    authDomain: "isp-billing-92ce2.firebaseapp.com",
    databaseURL: "https://isp-billing-92ce2-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "isp-billing-92ce2",
    storageBucket: "isp-billing-92ce2.firebasestorage.app",
    messagingSenderId: "951649795315",
    appId: "1:951649795315:web:f9a80ca1520d41b22c0b86",
    measurementId: "G-FCSWY5X5B8"
  };

  firebase.initializeApp(firebaseConfig);
  const _auth = firebase.auth();
  const _db = firebase.database();

  let _user = null;
  let _cache = [];
  let _mikrotikStatusMap = {};

  const toggleLoader = (val, text = "ডাটা প্রসেস হচ্ছে...") => {
    document.getElementById('loaderText').innerText = text;
    document.getElementById('loader').style.display = val ? 'flex' : 'none';
  };

  function getCurrentMonthKey() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  }

  // Auth Listener
  _auth.onAuthStateChanged(async (u) => {
    toggleLoader(true);
    try {
      if (u) {
        _user = u;
        document.getElementById('authSection').style.display = 'none';
        document.getElementById('dashboardSection').style.display = 'block';
        document.getElementById('userDisplayEmail').innerText = _user.email;
        document.getElementById('currentMonthLabel').innerText = `মাসিক সেশন: ${getCurrentMonthKey()}`;
        loadClients();
        fetchMikrotikLiveStatus();
      } else {
        _user = null;
        document.getElementById('authSection').style.display = 'block';
        document.getElementById('dashboardSection').style.display = 'none';
      }
    } catch (err) {
      alert("ডাটাবেজ সমস্যা: " + err.message);
    } finally {
      toggleLoader(false);
    }
  });

  document.getElementById('authForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    toggleLoader(true);
    const errorEl = document.getElementById('authError');
    errorEl.style.display = 'none';

    try {
      await _auth.signInWithEmailAndPassword(
        document.getElementById('authEmail').value,
        document.getElementById('authPassword').value
      );
    } catch (e) {
      errorEl.style.display = 'block';
      errorEl.innerText = e.code === 'auth/operation-not-allowed' 
        ? "Firebase Console-এ Email/Password অপশনটি Enable করুন!" 
        : "লগইন তথ্য ভুল হয়েছে!";
    } finally {
      toggleLoader(false);
    }
  });

  function logout() { _auth.signOut(); }

  // Load Firebase Clients Real-time
  function loadClients() {
    if (_user) {
      _db.ref(`users/${_user.uid}/isp_clients`).on('value', (snap) => {
        const val = snap.val();
        _cache = val ? Object.values(val) : [];
        filterData();
      });
    }
  }

  // Fetch MikroTik Status
  async function fetchMikrotikLiveStatus() {
    try {
      const res = await fetch(MIKROTIK_API_URL);
      const result = await res.json();
      if (result.success) {
        _mikrotikStatusMap = {};
        result.data.forEach(item => {
          _mikrotikStatusMap[item.username] = item;
        });
        filterData();
      }
    } catch (e) {
      console.warn("MikroTik Node Server Connection Failed:", e.message);
    }
  }

  // Render Table
  function renderTable(records) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    let totalPaid = 0, totalUnpaid = 0;

    const renderLimit = Math.min(records.length, 200);

    for (let i = 0; i < renderLimit; i++) {
      const item = records[i];
      const mkData = _mikrotikStatusMap[item.id];
      const isOnline = mkData && mkData.status === 'Online';

      const billBadge = item.status === 'Paid' 
        ? '<span class="badge bg-success">Paid</span>' 
        : '<span class="badge bg-danger">Unpaid</span>';

      const liveBadge = isOnline
        ? `<span class="badge status-badge bg-success-subtle text-success border border-success"><i class="pulse-online"></i> Online</span>`
        : `<span class="badge status-badge bg-secondary-subtle text-secondary">Offline</span>`;

      const row = `
        <tr>
          <td>
            <div class="fw-bold text-dark">${item.id}</div>
            <small class="text-muted">${item.pop || ''}</small>
          </td>
          <td>
            <div class="fw-semibold">${item.name}</div>
            <small class="text-muted">${item.phone || 'N/A'}</small>
          </td>
          <td><span class="badge bg-info text-dark">${item.package}</span></td>
          <td class="fw-bold">৳ ${item.billAmount.toLocaleString('en-US', {minimumFractionDigits: 2})}</td>
          <td>${billBadge}</td>
          <td>${liveBadge}</td>
          <td class="text-center">
            <button onclick="toggleStatus('${item.id}', '${item.status}')" class="btn btn-sm btn-outline-primary me-1" title="স্ট্যাটাস পরিবর্তন">
              <i class="fa-solid fa-arrows-rotate"></i>
            </button>
            <button onclick="deleteClient('${item.id}')" class="btn btn-sm btn-outline-danger" title="ডিলিট">
              <i class="fa-solid fa-trash"></i>
            </button>
          </td>
        </tr>
      `;
      tbody.innerHTML += row;
    }

    records.forEach(item => {
      if (item.status === 'Paid') totalPaid += item.billAmount;
      else totalUnpaid += item.billAmount;
    });

    document.getElementById('statTotalClients').innerText = records.length;
    document.getElementById('statTotalPaid').innerText = `৳ ${totalPaid.toLocaleString('en-US', {minimumFractionDigits: 2})}`;
    document.getElementById('statTotalUnpaid').innerText = `৳ ${totalUnpaid.toLocaleString('en-US', {minimumFractionDigits: 2})}`;
    document.getElementById('statOnlineCount').innerText = `${Object.values(_mikrotikStatusMap).filter(u=>u.status==='Online').length} / ${records.length}`;
  }

  // Upload Large Excel File (MERGE EXISTING DATA USING UPDATE)
  function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    document.getElementById('fileNameDisplay').innerText = file.name;
    toggleLoader(true, "ফাইল ডাটা প্রসেস হচ্ছে...");

    const reader = new FileReader();
    reader.onload = async function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        if (jsonData.length === 0) return alert("ফাইলটি খালি!");

        const currentMonth = getCurrentMonthKey();
        const chunkSize = 1000;
        const totalRows = jsonData.length;

        for (let i = 0; i < totalRows; i += chunkSize) {
          const chunk = jsonData.slice(i, i + chunkSize);
          const updates = {};

          chunk.forEach(row => {
            const username = (row['Username'] || row['Customer ID'] || row['id'] || '').toString().trim();
            if (!username) return;

            const name = row['Name'] || 'N/A';
            const mobile = (row['Mobile'] || row['Phone'] || '').toString().trim();
            const pkg = row['Package'] || 'Basic';
            const billAmount = Number(parseFloat(row['Bill Amount'] || 0).toFixed(2)) || 0;
            const statusRaw = row['Status'] || 'Enabled';
            const status = (statusRaw.toLowerCase() === 'enabled') ? 'Paid' : 'Unpaid';

            // Firebase path targeted specifically per user to preserve other users
            updates[`users/${_user.uid}/isp_clients/${username}`] = {
              id: username,
              name: name,
              phone: mobile,
              package: pkg,
              billAmount: billAmount,
              status: status,
              address: row['Address'] || '',
              pop: row['POP'] || '',
              lastUpdatedMonth: currentMonth,
              updatedAt: Date.now()
            };
          });

          toggleLoader(true, `মার্জ ও সেভ হচ্ছে: ${Math.min(i + chunkSize, totalRows)} / ${totalRows}...`);
          
          // Using update() ensures existing records stay untouched unless overwritten by matching Username
          await _db.ref().update(updates);
    <!-- External JS Libraries -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>

    <script>
        // ব্যাকএন্ড API URL (আপনার পিসির Node.js Server)
        const BACKEND_URL = 'http://localhost:5000';

        let globalUserData = [];

        // ১. MikroTik ডাটা ফেচিং ফাংশন
        async function fetchMikrotikLiveStatus() {
            try {
                const response = await fetch(`${BACKEND_URL}/api/mikrotik/users-status`);
                const result = await response.json();

                if (result.success) {
                    document.getElementById('server-status').className = 'badge bg-success fw-semibold';
                    document.getElementById('server-status').innerText = 'Server Connected';

                    globalUserData = result.data;
                    renderTable(globalUserData);
                    updateCounters(globalUserData);
                } else {
                    showServerError(result.message);
                }
            } catch (error) {
                showServerError("Node.js ব্যাকএন্ড সার্ভারের সাথে কানেক্ট করা যাচ্ছে না!");
            }
        }

        alert(`সফলভাবে ${totalRows} জন ক্লায়েন্টের ডাটা Firebase-এ যুক্ত (Merge) হয়েছে! আগের ডাটা অক্ষত রয়েছে।`);
      } catch (err) {
        alert("ত্রুটি: " + err.message);
      } finally {
        toggleLoader(false);
        document.getElementById('excelFileInput').value = '';
      }
    };

    reader.readAsArrayBuffer(file);
  }

  // Form Submit (Single User)
  document.getElementById('clientForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!_user) return;
    toggleLoader(true);

    const clientData = {
      id: document.getElementById('clientId').value.trim(),
      name: document.getElementById('clientName').value.trim(),
      phone: document.getElementById('clientPhone').value.trim(),
      package: document.getElementById('clientPackage').value.trim(),
      billAmount: Number(parseFloat(document.getElementById('billAmount').value).toFixed(2)) || 0,
      status: document.getElementById('paymentStatus').value,
      lastUpdatedMonth: getCurrentMonthKey(),
      updatedAt: Date.now()
    };

    try {
      await _db.ref(`users/${_user.uid}/isp_clients/${clientData.id}`).update(clientData);
      document.getElementById('clientForm').reset();
    } catch (err) {
      alert("ত্রুটি: " + err.message);
    } finally {
      toggleLoader(false);
    }
  });
        // ২. টেবিল রেন্ডার করা
        function renderTable(data) {
            const tbody = document.getElementById('clientTableBody');
            tbody.innerHTML = '';

            if (data.length === 0) {
                tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted py-3">কোনো কাস্টমার ডাটা পাওয়া যায়নি।</td></tr>';
                return;
            }

            data.forEach((user, index) => {
                const isOnline = user.status === 'Online';
                const isDisabled = user.disabled;

                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${index + 1}</td>
                    <td class="fw-bold">${user.username}</td>
                    <td><code>${user.address}</code></td>
                    <td>${user.uptime}</td>
                    <td>
                        ${isDisabled 
                            ? '<span class="badge bg-danger status-badge"><i class="fa-solid fa-ban me-1"></i>Disabled</span>'
                            : (isOnline 
                                ? '<span class="badge bg-success status-badge"><i class="fa-solid fa-circle me-1"></i>Online</span>' 
                                : '<span class="badge bg-secondary status-badge"><i class="fa-solid fa-circle me-1"></i>Offline</span>')
                        }
                    </td>
                    <td><small class="text-muted">${user.lastSeen}</small></td>
                    <td class="text-center">
                        <button onclick="toggleLine('${user.username}', ${isDisabled ? '\'enable\'' : '\'disable\''})" 
                                class="btn btn-sm ${isDisabled ? 'btn-success' : 'btn-outline-danger'} me-1" 
                                title="${isDisabled ? 'লাইন চালু করুন' : 'লাইন বন্ধ করুন'}">
                            <i class="fa-solid ${isDisabled ? 'fa-plug' : 'fa-power-off'}"></i>
                        </button>
                        <button onclick="sendBillSMS('${user.username}')" 
                                class="btn btn-sm btn-outline-info" 
                                title="বকেয়া বিলের SMS পাঠান">
                            <i class="fa-solid fa-comment-sms"></i>
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

  async function toggleStatus(id, currentStatus) {
    if (!_user) return;
    const newStatus = currentStatus === 'Paid' ? 'Unpaid' : 'Paid';
    toggleLoader(true);
    try {
      await _db.ref(`users/${_user.uid}/isp_clients/${id}`).update({ status: newStatus });
    } catch (e) { alert("ত্রুটি: " + e.message); } 
    finally { toggleLoader(false); }
  }
        // ৩. সামারি কাউন্টার আপডেট
        function updateCounters(data) {
            const total = data.length;
            const online = data.filter(u => u.status === 'Online' && !u.disabled).length;
            const disabled = data.filter(u => u.disabled).length;
            const offline = total - online - disabled;

            document.getElementById('total-users').innerText = total;
            document.getElementById('online-users').innerText = online;
            document.getElementById('offline-users').innerText = offline;
            document.getElementById('disabled-users').innerText = disabled;
        }

  async function deleteClient(id) {
    if (confirm("ক্লায়েন্টের ডাটা ডিলিট করতে চান?")) {
      toggleLoader(true);
      try {
        await _db.ref(`users/${_user.uid}/isp_clients/${id}`).remove();
      } catch (e) { alert("ত্রুটি: " + e.message); } 
      finally { toggleLoader(false); }
    }
  }
        // ৪. মাইক্রোটিকে লাইন চালু/বন্ধ করা
        async function toggleLine(username, action) {
            const actionText = action === 'disable' ? 'বন্ধ (Disable)' : 'চালু (Enable)';
            
            const confirm = await Swal.fire({
                title: 'আপনি কি নিশ্চিত?',
                text: `${username}-এর ইন্টারনেট লাইন ${actionText} করা হবে!`,
                icon: 'warning',
                showCancelButton: true,
                confirmButtonColor: action === 'disable' ? '#d33' : '#198754',
                cancelButtonColor: '#6c757d',
                confirmButtonText: `হ্যাঁ, ${actionText} করুন`,
                cancelButtonText: 'বাতিল'
            });

            if (!confirm.isConfirmed) return;

            try {
                Swal.showLoading();
                const res = await fetch(`${BACKEND_URL}/api/mikrotik/toggle-user`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ username, action })
                });

                const result = await res.json();
                if (result.success) {
                    Swal.fire('সফল!', `কাস্টমার লাইন সফলভাবে ${actionText} করা হয়েছে।`, 'success');
                    fetchMikrotikLiveStatus();
                } else {
                    Swal.fire('এরর!', result.message, 'error');
                }
            } catch (e) {
                Swal.fire('ত্রুটি!', 'সার্ভারে কানেক্ট করা সম্ভব হয়নি।', 'error');
            }
        }

  function filterData() {
    const q = document.getElementById('searchInput').value.toLowerCase();
    const filtered = _cache.filter(item => 
      item.name.toLowerCase().includes(q) ||
      item.id.toLowerCase().includes(q) ||
      (item.phone && item.phone.includes(q))
    );
    renderTable(filtered);
  }
        // ৫. বিলিং SMS পাঠানো
        async function sendBillSMS(username) {
            const { value: phone } = await Swal.fire({
                title: `${username}-কে SMS পাঠান`,
                input: 'text',
                inputLabel: 'গ্রাহকের মোবাইল নম্বর দিন',
                inputPlaceholder: '017XXXXXXXX',
                showCancelButton: true,
                confirmButtonText: 'SMS পাঠান',
                cancelButtonText: 'বাতিল',
                inputValidator: (value) => {
                    if (!value) return 'মোবাইল নম্বর দেওয়া বাধ্যতামূলক!';
                }
            });

            if (!phone) return;

            const message = `প্রিয় ${username}, আপনার চলতি মাসের ইন্টারনেট বিল বকেয়া রয়েছে। বিচ্ছিন্নতা এড়াতে দ্রুাত পরিশোধ করুন। ধন্যবাদ!`;

            try {
                Swal.showLoading();
                const res = await fetch(`${BACKEND_URL}/api/sms/send`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ phone, message })
                });

                const result = await res.json();
                if (result.success) {
                    Swal.fire('পাঠানো হয়েছে!', 'গ্রাহকের মোবাইলে SMS চলে গেছে।', 'success');
                } else {
                    Swal.fire('এরর!', result.message, 'error');
                }
            } catch (e) {
                Swal.fire('ত্রুটি!', 'SMS পাঠানোর সার্ভিস অফলাইন রয়েছে।', 'error');
            }
        }

  function exportExcel() {
    if (_cache.length === 0) return alert("ডাটা নেই!");
    const dataToExport = _cache.map(item => ({
      "Customer ID": item.id,
      "Username": item.id,
      "Name": item.name,
      "Mobile": item.phone,
      "Package": item.package,
      "Bill Amount": item.billAmount,
      "Status": item.status === 'Paid' ? 'Enabled' : 'Disabled',
      "Address": item.address || '',
      "POP": item.pop || ''
    }));
    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "ISPClients");
    XLSX.writeFile(workbook, `radcheck_export_${getCurrentMonthKey()}.xlsx`);
  }
        // ৬. টেবিল ফিল্টার/সার্চ
        function filterTable() {
            const input = document.getElementById("searchInput").value.toLowerCase();
            const filtered = globalUserData.filter(u => 
                u.username.toLowerCase().includes(input) || 
                u.address.toLowerCase().includes(input)
            );
            renderTable(filtered);
        }

  setInterval(fetchMikrotikLiveStatus, 15000);
</script>
        // ৭. সার্ভার এরর প্রদর্শন
        function showServerError(msg) {
            document.getElementById('server-status').className = 'badge bg-danger fw-semibold';
            document.getElementById('server-status').innerText = 'Server Offline';
            document.getElementById('clientTableBody').innerHTML = `
                <tr>
                    <td colspan="7" class="text-center text-danger py-4">
                        <i class="fa-solid fa-triangle-exclamation fa-2x mb-2"></i><br>
                        ${msg}<br>
                        <small class="text-muted">দয়া করে আপনার পিসিতে <code>node server.js</code> রান করা আছে কি না চেক করুন।</small>
                    </td>
                </tr>
            `;
        }

        // পেজ লোড হওয়ার সাথে সাথে ডাটা ফেচ হবে এবং প্রতি ১৫ সেকেন্ড পর পর লাইভ রিফ্রেশ হবে
        fetchMikrotikLiveStatus();
        setInterval(fetchMikrotikLiveStatus, 15000);
    </script>
</body>
</html>
