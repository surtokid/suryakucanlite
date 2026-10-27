
    const { useState, useEffect, useRef, useMemo } = React;

    // SheetJS cukup besar; unduh hanya saat pengguna benar-benar mengekspor Excel.
    let xlsxLoadPromise = null;
    function ensureXLSXLoaded() {
      if (window.XLSX) return Promise.resolve(window.XLSX);
      if (xlsxLoadPromise) return xlsxLoadPromise;
      xlsxLoadPromise = new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
        script.async = true;
        script.onload = () => window.XLSX
          ? resolve(window.XLSX)
          : reject(new Error('Library Excel tidak tersedia setelah dimuat.'));
        script.onerror = () => reject(new Error('Library Excel gagal dimuat. Periksa koneksi internet.'));
        document.head.appendChild(script);
      }).catch(error => {
        xlsxLoadPromise = null;
        throw error;
      });
      return xlsxLoadPromise;
    }


    // PWA install prompt: browser hanya menyediakan prompt pada kondisi yang memenuhi syarat.
    let deferredInstallPrompt = null;
    window.addEventListener('beforeinstallprompt', (event) => {
      event.preventDefault();
      deferredInstallPrompt = event;
      window.dispatchEvent(new Event('pwa-install-available'));
    });
    window.addEventListener('appinstalled', () => {
      deferredInstallPrompt = null;
      window.dispatchEvent(new Event('pwa-app-installed'));
    });
    if ('serviceWorker' in navigator && /^https?:$/.test(window.location.protocol)) {
      window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js').catch((error) => {
          console.warn('Service worker tidak dapat didaftarkan:', error);
        });
      });
    }

    async function requestPwaInstall(showToast) {
      const isStandalone = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
      if (isStandalone) {
        showToast('Aplikasi sudah terpasang di perangkat ini');
        return;
      }
      if (deferredInstallPrompt) {
        const installEvent = deferredInstallPrompt;
        deferredInstallPrompt = null;
        try {
          await installEvent.prompt();
          const choice = await installEvent.userChoice;
          if (choice && choice.outcome === 'accepted') {
            showToast('Konfirmasi instalasi diterima browser. Jika ikon belum muncul, periksa daftar aplikasi atau menu browser.');
          } else {
            showToast('Instalasi dibatalkan. Anda bisa mencoba lagi kapan saja.', 'error');
          }
        } catch (error) {
          console.error('PWA install gagal:', error);
          showToast('Instalasi belum dapat dimulai. Buka menu browser lalu pilih Install app / Tambahkan ke layar utama.', 'error');
        }
        return;
      }
      const ua = navigator.userAgent || '';
      if (/iphone|ipad|ipod/i.test(ua)) {
        showToast('Di iPhone/iPad: tekan Bagikan, lalu pilih Tambahkan ke Layar Utama.');
      } else if (/android/i.test(ua)) {
        showToast('Buka menu ⋮ Chrome, lalu pilih Instal aplikasi atau Tambahkan ke layar utama. Pastikan halaman dibuka lewat HTTPS.');
      } else {
        showToast('Browser belum menyediakan tombol instal otomatis. Gunakan menu browser lalu pilih Install app / Instal halaman ini. Pastikan aplikasi dibuka lewat HTTPS.');
      }
    }

    // Gunakan tanggal lokal perangkat (zona waktu Indonesia), bukan UTC.
    // Ini mencegah tanggal presensi bergeser saat pergantian hari.
    function getLocalDateKey(date = new Date()) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    function formatDateDMY(dateKey) {
      if (!dateKey) return '';
      const match = String(dateKey).match(/^(\d{4})-(\d{2})-(\d{2})$/);
      return match ? `${match[3]}/${match[2]}/${match[1]}` : dateKey;
    }

    function shiftDateKey(dateKey, days) {
      const parts = String(dateKey || '').split('-').map(Number);
      if (parts.length !== 3 || parts.some(Number.isNaN)) return getLocalDateKey();
      const date = new Date(parts[0], parts[1] - 1, parts[2]);
      date.setDate(date.getDate() + days);
      return getLocalDateKey(date);
    }

    // Supabase & Google Drive via Google Apps Script (GAS) Configuration
    const SUPABASE_URL = "https://hxpoajftzysfsxgjcgex.supabase.co";
    const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imh4cG9hamZ0enlzZnN4Z2pjZ2V4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTI0MDA1OTUsImV4cCI6MjEwNzk3NjU5NX0.2_LNfRj9B54uINQNMO4zecN3no8B6mYl07fNJ0mYHqM";
    const GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbz2XNXvOnw-bL9SVKhcgJvKuYiRYJtLMqoNLeGSesIUss_s2rv6c83TO-ZB0QsP5J70OA/exec";
    const GOOGLE_DRIVE_FOLDER_ID = "10QkIZG7iJbdcHwRTyKZt1EKlqADJ8V36";

    // Supabase client dibuat secara lazy agar tidak bergantung pada urutan CDN.
    let supabase = null;
    let supabaseClientPromise = null;

    function loadSupabaseFallback() {
      return new Promise((resolve) => {
        const existing = document.querySelector('script[data-supabase-fallback="true"]');
        if (window.supabase?.createClient) return resolve(true);
        const script = existing || document.createElement('script');
        let settled = false;
        const finish = (ok) => {
          if (settled) return;
          settled = true;
          clearTimeout(timer);
          resolve(ok && !!window.supabase?.createClient);
        };
        const timer = setTimeout(() => finish(false), 4000);
        script.src = 'https://unpkg.com/@supabase/supabase-js@2/dist/umd/supabase.js';
        script.async = true;
        script.dataset.supabaseFallback = 'true';
        script.onload = () => finish(!!window.supabase?.createClient);
        script.onerror = () => finish(false);
        if (!existing) document.head.appendChild(script);
      });
    }

    async function getSupabaseClient() {
      if (supabase) return supabase;
      if (supabaseClientPromise) return supabaseClientPromise;

      supabaseClientPromise = (async () => {
        // Polling maksimal sekitar 3,5 detik untuk CDN utama.
        for (let i = 0; i < 14; i++) {
          if (window.supabase?.createClient) break;
          await new Promise(resolve => setTimeout(resolve, 250));
        }
        if (!window.supabase?.createClient) {
          await loadSupabaseFallback();
        }
        if (!window.supabase?.createClient) {
          throw new Error('Library Supabase tidak berhasil dimuat. Periksa koneksi/CDN.');
        }
        supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
        return supabase;
      })();

      try {
        return await supabaseClientPromise;
      } catch (error) {
        supabaseClientPromise = null;
        throw error;
      }
    }

    // Upload media ke Google Drive melalui Google Apps Script Web App.
    // Kontrak disamakan dengan Code.gs: payload.filename dan respons result.url.
    const uploadToGoogleDriveViaGAS = async (base64Image, fileName = '', subFolder = 'presensi') => {
      if (!base64Image || !base64Image.startsWith('data:image')) {
        throw new Error('Data foto tidak valid atau belum dipilih.');
      }
      if (!GAS_WEB_APP_URL || GAS_WEB_APP_URL.includes('YOUR_DEPLOYMENT_ID')) {
        throw new Error('URL Google Apps Script belum dikonfigurasi pada aplikasi.');
      }
      try {
        const payload = {
          image: base64Image,
          filename: fileName || `${subFolder}_${Date.now()}.jpg`,
          folderId: GOOGLE_DRIVE_FOLDER_ID
        };
        const response = await fetch(GAS_WEB_APP_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(payload),
          redirect: 'follow'
        });
        const responseText = await response.text();
        let result;
        try {
          result = JSON.parse(responseText);
        } catch (_) {
          throw new Error(`Respons Google Apps Script bukan JSON yang valid (HTTP ${response.status}). Pastikan URL Web App dan deployment benar.`);
        }
        if (!response.ok) {
          throw new Error(`Google Apps Script mengembalikan HTTP ${response.status}: ${result.message || response.statusText}`);
        }
        if (result && result.status === 'success' && (result.url || result.webViewLink)) {
          return result.url || result.webViewLink;
        }
        throw new Error((result && result.message) || 'Google Apps Script tidak mengembalikan URL file. Periksa izin folder dan deployment.');
      } catch (err) {
        console.error('Gagal mengunggah foto ke Google Drive:', err);
        throw err instanceof Error ? err : new Error(String(err));
      }
    };

    // Sound Notification with Web Audio API
    const playNotificationSound = (type = 'success') => {
      try {
        const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
        const osc = audioCtx.createOscillator();
        const gain = audioCtx.createGain();
        osc.connect(gain);
        gain.connect(audioCtx.destination);
        
        if (type === 'success') {
          osc.type = 'sine';
          osc.frequency.setValueAtTime(587.33, audioCtx.currentTime); // D5
          osc.frequency.exponentialRampToValueAtTime(880, audioCtx.currentTime + 0.15); // A5
          gain.gain.setValueAtTime(0.15, audioCtx.currentTime);
          gain.gain.exponentialRampToValueAtTime(0.01, audioCtx.currentTime + 0.3);
          osc.start(audioCtx.currentTime);
          osc.stop(audioCtx.currentTime + 0.3);
        } else if (type === 'error') {
          osc.type = 'sawtooth';
          osc.frequency.setValueAtTime(300, audioCtx.currentTime);
          osc.frequency.setValueAtTime(200, audioCtx.currentTime + 0.1);
          gain.gain.setValueAtTime(0.2, audioCtx.currentTime);
          gain.gain.exponentialRampToValueAtTime(0.01, audioCtx.currentTime + 0.3);
          osc.start(audioCtx.currentTime);
          osc.stop(audioCtx.currentTime + 0.3);
        }
      } catch (err) {
        console.warn("Audio Context error:", err);
      }
    };

    // Haversine Geofence Calculation Formula
    function calculateDistance(lat1, lon1, lat2, lon2) {
      const R = 6371e3; // metres
      const phi1 = (lat1 * Math.PI) / 180;
      const phi2 = (lat2 * Math.PI) / 180;
      const deltaPhi = ((lat2 - lat1) * Math.PI) / 180;
      const deltaLambda = ((lon2 - lon1) * Math.PI) / 180;

      const a = Math.sin(deltaPhi / 2) * Math.sin(deltaPhi / 2) +
                Math.cos(phi1) * Math.cos(phi2) *
                Math.sin(deltaLambda / 2) * Math.sin(deltaLambda / 2);
      const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
      return Math.round(R * c); // Distance in meters
    }

    // Google Drive URL Formatter
    function getCleanImageUrl(url) {
      if (!url) return '';
      if (url.startsWith('data:image')) return url;
      const match = url.match(/[-\w]{25,}/);
      if (match && url.includes('drive.google.com')) {
        return `https://lh3.googleusercontent.com/d/${match[0]}`;
      }
      return url;
    }

    // URL foto profil diberi versi agar browser tidak memakai gambar lama dari cache.
    function getProfileImageUrl(url, version) {
      const cleanUrl = getCleanImageUrl(url);
      if (!cleanUrl) return '';
      const versionKey = version || 'latest';
      try {
        const parsed = new URL(cleanUrl, window.location.href);
        parsed.searchParams.set('profile_v', String(versionKey));
        return parsed.toString();
      } catch (_) {
        return cleanUrl + (cleanUrl.includes('?') ? '&' : '?') + 'profile_v=' + encodeURIComponent(String(versionKey));
      }
    }

    // Format Indonesian Date
    function formatIndoDate(dateString) {
      const date = new Date(dateString);
      return date.toLocaleDateString('id-ID', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric'
      });
    }

const DEFAULT_MORE_APPS = [
  { label: 'Cek Harga Surtok', icon: 'fa-magnifying-glass-dollar', description: 'Pengecekan harga produk Surtok.', color: 'from-violet-50 to-purple-100', ink: 'text-violet-700' },
  { label: 'Cek Harga Surcan', icon: 'fa-tags', description: 'Pengecekan harga produk Surcan.', color: 'from-sky-50 to-cyan-100', ink: 'text-sky-700' },
  { label: 'Request Harga Coret', icon: 'fa-percent', description: 'Pengajuan harga coret.', color: 'from-amber-50 to-orange-100', ink: 'text-amber-700' },
  { label: 'Request Purchase Order', icon: 'fa-cart-shopping', description: 'Pengajuan purchase order.', color: 'from-emerald-50 to-teal-100', ink: 'text-emerald-700' },
  { label: 'Request Price Tag', icon: 'fa-ticket', description: 'Permintaan price tag.', color: 'from-fuchsia-50 to-pink-100', ink: 'text-fuchsia-700' },
  { label: 'Request Wobbler', icon: 'fa-sign-hanging', description: 'Permintaan materi wobbler.', color: 'from-fuchsia-50 to-pink-100', ink: 'text-fuchsia-700' },
  { label: 'Stock Opname', icon: 'fa-clipboard-check', description: 'Pencatatan dan pemeriksaan stok.', color: 'from-indigo-50 to-sky-100', ink: 'text-indigo-700' },
  { label: 'Kasir Surtok', icon: 'fa-cash-register', description: 'Kasir untuk event di luar toko.', color: 'from-pink-50 to-rose-100', ink: 'text-rose-700' },
  { label: 'Kasir Surcan', icon: 'fa-store', description: 'Kasir untuk event di luar toko.', color: 'from-cyan-50 to-teal-100', ink: 'text-teal-700' }
];

    // Dialog modern pengganti alert/confirm bawaan browser (This Page Says).
    function showModernDialog(message, options = {}) {
      return new Promise(resolve => {
        const old = document.getElementById('suryakucan-modern-dialog');
        if (old) old.remove();
        const confirmMode = !!options.confirm;
        const overlay = document.createElement('div');
        overlay.id = 'suryakucan-modern-dialog';
        overlay.style.cssText = 'position:fixed;inset:0;z-index:999999;display:flex;align-items:center;justify-content:center;padding:20px;background:rgba(15,23,42,.48);backdrop-filter:blur(7px);font-family:Inter,system-ui,sans-serif;';
        overlay.innerHTML = `<div role="dialog" aria-modal="true" style="width:min(100%,390px);background:#fff;border-radius:24px;padding:24px;box-shadow:0 25px 80px rgba(15,23,42,.28);border:1px solid #f1f5f9;transform:translateY(0);animation:scDialogIn .18s ease-out"><div style="width:48px;height:48px;border-radius:16px;background:${confirmMode?'#fff7ed':'#fdf2f8'};color:${confirmMode?'#c2410c':'#db2777'};display:flex;align-items:center;justify-content:center;font-size:23px;margin-bottom:16px">${confirmMode?'!':'✓'}</div><h3 style="font-size:18px;font-weight:800;color:#0f172a;margin:0 0 8px">${options.title || (confirmMode?'Konfirmasi':'Informasi')}</h3><p style="font-size:14px;line-height:1.6;color:#475569;white-space:pre-wrap;margin:0 0 22px"></p><div style="display:flex;justify-content:flex-end;gap:10px"><button data-cancel style="display:${confirmMode?'inline-flex':'none'};padding:10px 16px;border-radius:12px;border:1px solid #e2e8f0;background:#fff;color:#475569;font-weight:700;cursor:pointer">Batal</button><button data-ok style="padding:10px 18px;border-radius:12px;border:0;background:linear-gradient(135deg,#e03ea0,#822ab0);color:#fff;font-weight:800;cursor:pointer">${confirmMode?'Ya, lanjutkan':'Mengerti'}</button></div></div>`;
        overlay.querySelector('p').textContent = String(message ?? '');
        const finish = value => { overlay.remove(); resolve(value); };
        overlay.querySelector('[data-ok]').onclick = () => finish(true);
        overlay.querySelector('[data-cancel]').onclick = () => finish(false);
        overlay.addEventListener('click', e => { if (e.target === overlay && confirmMode) finish(false); });
        const style = document.createElement('style'); style.textContent='@keyframes scDialogIn{from{opacity:0;transform:translateY(8px) scale(.98)}to{opacity:1;transform:translateY(0) scale(1)}}'; overlay.appendChild(style);
        document.body.appendChild(overlay);
        overlay.querySelector('[data-ok]').focus();
      });
    }
    window.alert = (message) => { showModernDialog(message); };
    const modernConfirm = (message, title) => showModernDialog(message, { confirm: true, title: title || 'Konfirmasi tindakan' });

    function App() {
      const [currentRoute, setCurrentRoute] = useState(window.location.hash || '#karyawan');
    
      // Auth Session State
      const [authSession, setAuthSession] = useState(() => {
        const saved = localStorage.getItem('suryatokoku_session');
        return saved ? JSON.parse(saved) : null;
      });

      const [currentUser, setCurrentUser] = useState(() => {
        const savedUser = localStorage.getItem('suryatokoku_user');
        if (savedUser) return JSON.parse(savedUser);
        return null;
      });

      const [isAdmin, setIsAdmin] = useState(() => {
        return localStorage.getItem('suryatokoku_is_admin') === 'true';
      });

      // State murni dari Supabase (tanpa data dummy)
      const [karyawanList, setKaryawanList] = useState([]);
      const [lokasiList, setLokasiList] = useState([]);
      const [divisiList, setDivisiList] = useState([]);
      const [jabatanList, setJabatanList] = useState([]);
      const [shiftList, setShiftList] = useState([]);
      const [pengumumanList, setPengumumanList] = useState([]);
      const [adminList, setAdminList] = useState([]);
      const [isAdminTableEmpty, setIsAdminTableEmpty] = useState(false);
      const [toleransiMenit, setToleransiMenit] = useState(15);
      const [pendaftaranKaryawanAktif, setPendaftaranKaryawanAktif] = useState(false);
      const [moreAppsAccess, setMoreAppsAccess] = useState({});
      const [moreAppsCatalog, setMoreAppsCatalog] = useState(DEFAULT_MORE_APPS);
        
      const [presensiList, setPresensiList] = useState([]);
      const [isPresensiReady, setIsPresensiReady] = useState(false);
      const [lemburList, setLemburList] = useState([]);
      const [klaimList, setKlaimList] = useState([]);
      const [toast, setToast] = useState(null);
      const [showLogoutConfirm, setShowLogoutConfirm] = useState(false);
      const [isLoadingData, setIsLoadingData] = useState(false);

      // Sinkronisasi awal + Supabase Realtime + fallback saat aplikasi aktif.
      // Perubahan database memicu pengambilan ulang data tanpa reload seluruh halaman.
      useEffect(() => {
        let isMounted = true;
        let refreshTimer = null;
        let pollingTimer = null;
        let realtimeChannel = null;
        let isFetchingAll = false;
        let refreshQueued = false;

        const fetchAllFromSupabase = async ({ silent = false } = {}) => {
          // Jika refresh lain sedang berjalan, gabungkan permintaan berikutnya
          // menjadi satu refresh susulan agar tidak ada banyak query berat bersamaan.
          if (isFetchingAll) {
            refreshQueued = true;
            return;
          }
          isFetchingAll = true;
          if (!silent) setIsLoadingData(true);
          try {
            const client = await getSupabaseClient();
            const results = await Promise.all([
              client.from('master_karyawan').select('*'),
              client.from('master_lokasi').select('*'),
              client.from('master_divisi').select('*'),
              client.from('master_jabatan').select('*'),
              client.from('master_shift').select('*'),
              client.from('master_admin').select('*'),
              client.from('pengumuman').select('*').order('created_at', { ascending: false }),
              client.from('presensi').select('*').order('created_at', { ascending: false }),
              client.from('pengajuan_lembur').select('*').order('created_at', { ascending: false }),
              client.from('klaim_kehadiran').select('*'),
              client.from('settings').select('key,value').in('key', ['toleransi_menit', 'pendaftaran_karyawan_aktif', 'more_apps_access', 'more_apps_catalog'])
            ]);

            if (!isMounted) return;
            const [
              kryRes, lokRes, divRes, jabRes, shRes, admRes,
              pgmRes, prsRes, lmbRes, klmRes, settingsRes
            ] = results;

            const assignIfOk = (result, setter, label) => {
              if (result.error) {
                console.warn(`Gagal memuat ${label}:`, result.error.message);
                return;
              }
              setter(result.data || []);
            };

            assignIfOk(kryRes, setKaryawanList, 'master_karyawan');

            // Sinkronkan juga sesi karyawan yang sedang login, bukan hanya daftar master.
            // Tanpa ini, state currentUser/localStorage dapat menyimpan divisi, jabatan,
            // atau shift lama walaupun data master sudah berhasil diperbarui dari Realtime.
            if (!kryRes.error && authSession?.type === 'karyawan' && currentUser?.id != null) {
              const freshCurrentUser = (kryRes.data || []).find(
                row => String(row.id) === String(currentUser.id)
              );
              if (freshCurrentUser) {
                setCurrentUser(previousUser => {
                  if (!previousUser || String(previousUser.id) !== String(freshCurrentUser.id)) {
                    return previousUser;
                  }
                  const mergedUser = { ...previousUser, ...freshCurrentUser };
                  try {
                    localStorage.setItem('suryatokoku_user', JSON.stringify(mergedUser));
                  } catch (storageError) {
                    console.warn('Gagal memperbarui cache sesi karyawan:', storageError);
                  }
                  return mergedUser;
                });
              }
            }

            assignIfOk(lokRes, setLokasiList, 'master_lokasi');
            assignIfOk(divRes, setDivisiList, 'master_divisi');
            assignIfOk(jabRes, setJabatanList, 'master_jabatan');
            assignIfOk(shRes, setShiftList, 'master_shift');

            if (!admRes.error) {
              const admins = admRes.data || [];
              setAdminList(admins);
              setIsAdminTableEmpty(admins.length === 0);
            } else {
              console.warn('Gagal memuat master_admin:', admRes.error.message);
              setIsAdminTableEmpty(false);
            }

            assignIfOk(pgmRes, setPengumumanList, 'pengumuman');
            assignIfOk(prsRes, setPresensiList, 'presensi');
            // Jangan jalankan pemeriksaan presensi kemarin sebelum data presensi selesai dimuat.
            if (!prsRes.error) setIsPresensiReady(true);
            assignIfOk(lmbRes, setLemburList, 'pengajuan_lembur');
            assignIfOk(klmRes, setKlaimList, 'klaim_kehadiran');

            if (settingsRes.error) {
              console.warn('Gagal memuat pengaturan aplikasi:', settingsRes.error.message);
            } else {
              const settingsRows = settingsRes.data || [];
              const toleransiSetting = settingsRows.find(x => x.key === 'toleransi_menit');
              const pendaftaranSetting = settingsRows.find(x => x.key === 'pendaftaran_karyawan_aktif');
              const moreAppsSetting = settingsRows.find(x => x.key === 'more_apps_access');
              const moreAppsCatalogSetting = settingsRows.find(x => x.key === 'more_apps_catalog');
              if (moreAppsSetting?.value) { try { const parsed = typeof moreAppsSetting.value === 'string' ? JSON.parse(moreAppsSetting.value) : moreAppsSetting.value; setMoreAppsAccess(parsed && typeof parsed === 'object' ? parsed : {}); } catch (e) { console.warn('Pengaturan akses More Apps tidak valid:', e); setMoreAppsAccess({}); } }
              if (moreAppsCatalogSetting?.value) { try { const parsedCatalog = typeof moreAppsCatalogSetting.value === 'string' ? JSON.parse(moreAppsCatalogSetting.value) : moreAppsCatalogSetting.value; if (Array.isArray(parsedCatalog)) setMoreAppsCatalog(parsedCatalog); } catch (e) { console.warn('Katalog More Apps tidak valid:', e); } }
              if (toleransiSetting?.value != null && toleransiSetting.value !== '') {
                const value = Number(toleransiSetting.value);
                if (Number.isFinite(value)) setToleransiMenit(value);
              }
              setPendaftaranKaryawanAktif(String(pendaftaranSetting?.value || 'false').toLowerCase() === 'true');
            }
          } catch (err) {
            console.error('Gagal menginisialisasi/memuat Supabase:', err);
          } finally {
            isFetchingAll = false;
            if (isMounted && !silent) setIsLoadingData(false);
            if (isMounted && refreshQueued) {
              refreshQueued = false;
              if (refreshTimer) clearTimeout(refreshTimer);
              refreshTimer = setTimeout(() => {
                if (isMounted && document.visibilityState === 'visible' && navigator.onLine) {
                  fetchAllFromSupabase({ silent: true });
                }
              }, 300);
            }
          }
        };

        const scheduleSilentRefresh = () => {
          if (!isMounted) return;
          if (refreshTimer) clearTimeout(refreshTimer);
          // Gabungkan beberapa perubahan berdekatan agar tidak menembakkan query berulang.
          refreshTimer = setTimeout(() => {
            if (isMounted && document.visibilityState === 'visible') {
              fetchAllFromSupabase({ silent: true });
            }
          }, 400);
        };

        const tableNames = [
          'klaim_kehadiran',
          'master_admin',
          'master_divisi',
          'master_jabatan',
          'master_karyawan',
          'master_lokasi',
          'master_shift',
          'pengajuan_lembur',
          'pengumuman',
          'presensi',
          'settings'
        ];

        fetchAllFromSupabase();

        // Berlangganan perubahan INSERT / UPDATE / DELETE pada tabel aplikasi.
        // Tabel harus diaktifkan pada publication supabase_realtime di Supabase.
        getSupabaseClient().then(client => {
          if (!isMounted) return;
          realtimeChannel = client.channel('suryakucan-db-sync-v1');
          tableNames.forEach(table => {
            realtimeChannel = realtimeChannel.on(
              'postgres_changes',
              { event: '*', schema: 'public', table },
              () => scheduleSilentRefresh()
            );
          });
          realtimeChannel.subscribe(status => {
            if (status === 'SUBSCRIBED') {
              console.info('SuryaKuCan Realtime aktif untuk tabel aplikasi.');
            } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
              console.warn('Koneksi Supabase Realtime bermasalah:', status);
            }
          });
        }).catch(error => {
          console.warn('Realtime belum dapat diaktifkan; penyegaran saat aplikasi aktif tetap berjalan.', error);
        });

        // Fallback jika event Realtime terlewat atau belum tersambung.
        pollingTimer = setInterval(() => {
          if (isMounted && document.visibilityState === 'visible' && navigator.onLine) {
            fetchAllFromSupabase({ silent: true });
          }
        }, 60000);

        // Sinkronkan segera saat pengguna kembali ke aplikasi dari background.
        const handleVisibilityChange = () => {
          if (document.visibilityState === 'visible' && navigator.onLine) {
            fetchAllFromSupabase({ silent: true });
          }
        };
        const handleOnline = () => fetchAllFromSupabase({ silent: true });
        document.addEventListener('visibilitychange', handleVisibilityChange);
        window.addEventListener('online', handleOnline);

        return () => {
          isMounted = false;
          if (refreshTimer) clearTimeout(refreshTimer);
          if (pollingTimer) clearInterval(pollingTimer);
          document.removeEventListener('visibilitychange', handleVisibilityChange);
          window.removeEventListener('online', handleOnline);
          if (realtimeChannel) {
            getSupabaseClient().then(client => client.removeChannel(realtimeChannel)).catch(() => {});
          }
        };
      }, []);

    // Effect terpisah agar selalu memakai authSession/currentUser terbaru (bukan closure awal Realtime).
    useEffect(() => {
      if (authSession?.type !== 'karyawan' || currentUser?.id == null || !Array.isArray(karyawanList)) return;
      const freshCurrentUser = karyawanList.find(row => String(row.id) === String(currentUser.id));
      if (!freshCurrentUser) return;

      const mergedUser = { ...currentUser, ...freshCurrentUser };
      // Hindari render berulang jika data sesi sudah sama dengan data master.
      if (JSON.stringify(currentUser) === JSON.stringify(mergedUser)) return;

      setCurrentUser(mergedUser);
      try {
        localStorage.setItem('suryatokoku_user', JSON.stringify(mergedUser));
      } catch (storageError) {
        console.warn('Gagal memperbarui cache sesi karyawan:', storageError);
      }
    }, [karyawanList, authSession, currentUser]);

    const showToast = (message, type = 'success') => {
      setToast({ message, type, id: Date.now() });
      playNotificationSound(type);
      setTimeout(() => setToast(null), 3500);
    };

    // Pembungkus setter untuk memastikan perubahan CRUD dari modul admin
    // benar-benar tersimpan ke Supabase, bukan hanya berubah di state React.
    const syncedListRefs = useRef({});
    syncedListRefs.current = {
      karyawan: karyawanList,
      lokasi: lokasiList,
      divisi: divisiList,
      jabatan: jabatanList,
      shift: shiftList,
      admin: adminList,
      pengumuman: pengumumanList,
      lembur: lemburList,
      klaim: klaimList
    };

    const makeSyncedSetter = (key, table, originalSetter) => (action) => {
      const previous = Array.isArray(syncedListRefs.current[key]) ? syncedListRefs.current[key] : [];
      const next = typeof action === 'function' ? action(previous) : action;
      if (!Array.isArray(next)) {
        originalSetter(action);
        return;
      }
      // Update ref segera agar dua perubahan berurutan tidak memakai snapshot lama.
      syncedListRefs.current[key] = next;
      originalSetter(next);

      (async () => {
        try {
          const client = await getSupabaseClient();
          const oldById = new Map(previous.filter(x => x && x.id != null).map(x => [String(x.id), x]));
          const newById = new Map(next.filter(x => x && x.id != null).map(x => [String(x.id), x]));
          const removedIds = [...oldById.keys()].filter(id => !newById.has(id));
          const changedRows = [...newById.values()].filter(row => {
            const old = oldById.get(String(row.id));
            return !old || JSON.stringify(old) !== JSON.stringify(row);
          });

          if (changedRows.length) {
            // Properti alamat hanya untuk kebutuhan formulir; skema Supabase
            // menyimpan detail alamat dalam satu kolom `alamat` (JSON string),
            // bukan kolom terpisah `provinsi`, `kota`, atau `jalan`.
            const rowsForDatabase = table === 'master_karyawan'
              ? changedRows.map(row => {
                  const { provinsi, kota, kecamatan, desa, kelurahan, jalan, ...dbRow } = row;
                  // Bila ada detail alamat dari formulir, pastikan semuanya
                  // tetap tersimpan pada kolom alamat yang didukung skema.
                  if (provinsi !== undefined || kota !== undefined || kecamatan !== undefined || desa !== undefined || kelurahan !== undefined || jalan !== undefined) {
                    let alamat = {};
                    try {
                      alamat = typeof dbRow.alamat === 'string'
                        ? JSON.parse(dbRow.alamat || '{}')
                        : (dbRow.alamat || {});
                    } catch (_) { alamat = {}; }
                    dbRow.alamat = JSON.stringify({
                      ...alamat,
                      ...(provinsi !== undefined ? { provinsi } : {}),
                      ...(kota !== undefined ? { kota } : {}),
                      ...(kecamatan !== undefined ? { kecamatan } : {}),
                      ...(desa !== undefined ? { desa } : {}),
                      ...(kelurahan !== undefined ? { kelurahan } : {}),
                      ...(jalan !== undefined ? { jalan } : {})
                    });
                  }
                  return dbRow;
                })
              : changedRows;
            const { error } = await client.from(table).upsert(rowsForDatabase, { onConflict: 'id' });
            if (error) throw error;
          }
          for (const id of removedIds) {
            const { error } = await client.from(table).delete().eq('id', id);
            if (error) throw error;
          }
        } catch (error) {
          console.error(`Gagal sinkronisasi perubahan ke tabel ${table}:`, error);
          // Rollback state lokal bila perubahan yang gagal masih merupakan perubahan terbaru.
          if (syncedListRefs.current[key] === next) {
            syncedListRefs.current[key] = previous;
            originalSetter(previous);
          }
          showToast(`Gagal menyimpan ke Supabase (${table}): ${error.message || 'periksa koneksi, struktur kolom, dan RLS'}`, 'error');
        }
      })();
    };

    const setKaryawanListSynced = makeSyncedSetter('karyawan', 'master_karyawan', setKaryawanList);
    const setLokasiListSynced = makeSyncedSetter('lokasi', 'master_lokasi', setLokasiList);
    const setDivisiListSynced = makeSyncedSetter('divisi', 'master_divisi', setDivisiList);
    const setJabatanListSynced = makeSyncedSetter('jabatan', 'master_jabatan', setJabatanList);
    const setShiftListSynced = makeSyncedSetter('shift', 'master_shift', setShiftList);
    const setAdminListSynced = makeSyncedSetter('admin', 'master_admin', setAdminList);
    const setPengumumanListSynced = makeSyncedSetter('pengumuman', 'pengumuman', setPengumumanList);
    const setLemburListSynced = makeSyncedSetter('lembur', 'pengajuan_lembur', setLemburList);
    const setKlaimListSynced = makeSyncedSetter('klaim', 'klaim_kehadiran', setKlaimList);

    const setToleransiMenitSynced = async (valueOrUpdater) => {
      const value = Number(typeof valueOrUpdater === 'function' ? valueOrUpdater(toleransiMenit) : valueOrUpdater);
      if (!Number.isFinite(value) || value < 0) {
        throw new Error('Nilai toleransi harus berupa angka 0 atau lebih.');
      }
      const client = await getSupabaseClient();
      // Simpan ke Supabase terlebih dahulu; jangan memberi kesan berhasil jika database gagal.
      const { error } = await client.from('settings').upsert(
        { key: 'toleransi_menit', value: String(value) },
        { onConflict: 'key' }
      );
      if (error) {
        console.error('Gagal menyimpan settings toleransi:', error);
        throw error;
      }
      setToleransiMenit(value);
      return value;
    };

    const setPendaftaranKaryawanAktifSynced = async (enabled) => {
      const value = Boolean(enabled);
      const client = await getSupabaseClient();
      const { error } = await client.from('settings').upsert(
        { key: 'pendaftaran_karyawan_aktif', value: value ? 'true' : 'false' },
        { onConflict: 'key' }
      );
      if (error) throw error;
      setPendaftaranKaryawanAktif(value);
      return value;
    };

    const setMoreAppsAccessSynced = async (nextValue) => {
      const value = nextValue && typeof nextValue === 'object' ? nextValue : {};
      const client = await getSupabaseClient();
      const { error } = await client.from('settings').upsert(
        { key: 'more_apps_access', value: JSON.stringify(value) },
        { onConflict: 'key' }
      );
      if (error) throw error;
      setMoreAppsAccess(value);
      return value;
    };

    const setMoreAppsCatalogSynced = async (nextValue) => {
      const value = Array.isArray(nextValue) ? nextValue : [];
      const client = await getSupabaseClient();
      const { error } = await client.from('settings').upsert(
        { key: 'more_apps_catalog', value: JSON.stringify(value) },
        { onConflict: 'key' }
      );
      if (error) throw error;
      setMoreAppsCatalog(value);
      return value;
    };

    useEffect(() => {
      const handleHashChange = () => {
        setCurrentRoute(window.location.hash || '#karyawan');
      };
      window.addEventListener('hashchange', handleHashChange);
      return () => window.removeEventListener('hashchange', handleHashChange);
    }, []);

    const navigateTo = (hash) => {
      window.location.hash = hash;
      setCurrentRoute(hash);
    };

    const isDashboardRoute = currentRoute.startsWith('#dashboard');

    const handleLoginKaryawan = async (identifier, password) => {
      const trimmedId = String(identifier || '').trim().toLowerCase();
      const findUser = (list) => (Array.isArray(list) ? list : []).find(k =>
        (((k.username || '').toString().trim().toLowerCase() === trimmedId) ||
         ((k.no_wa || '').toString().trim().toLowerCase() === trimmedId) ||
         ((k.email || '').toString().trim().toLowerCase() === trimmedId)) &&
        String(k.password ?? '') === String(password ?? '')
      );

      // Cek cache state terlebih dahulu agar login tetap cepat.
      let user = findUser(karyawanList);
      // Jika tidak ditemukan, ambil data terbaru dari Supabase. Ini penting setelah
      // pendaftaran baru agar karyawan tidak perlu me-refresh halaman secara manual.
      if (!user) {
        try {
          const client = await getSupabaseClient();
          const { data, error } = await client.from('master_karyawan').select('*');
          if (error) throw error;
          const latestList = Array.isArray(data) ? data : [];
          setKaryawanList(latestList);
          user = findUser(latestList);
        } catch (error) {
          console.error('Gagal menyegarkan data karyawan saat login:', error);
          showToast('Data login belum dapat diperbarui dari server. Periksa koneksi lalu coba lagi.', 'error');
          return { success: false, message: 'Gagal memuat data terbaru' };
        }
      }

      if (user) {
        setCurrentUser(user);
        setIsAdmin(false);
        setAuthSession({ type: 'karyawan', id: user.id });
        localStorage.setItem('suryatokoku_user', JSON.stringify(user));
        localStorage.setItem('suryatokoku_session', JSON.stringify({ type: 'karyawan', id: user.id }));
        localStorage.setItem('suryatokoku_is_admin', 'false');
        navigateTo('#karyawan');
        showToast(`Selamat datang, ${user.nama_lengkap}!`);
        return { success: true };
      } else {
        showToast('Username/No.WA/Email atau password salah!', 'error');
        return { success: false, message: 'Kredensial tidak cocok' };
      }
    };

    const handleLoginAdmin = (identifier, password) => {
      const trimmedId = identifier.trim().toLowerCase();
      const admin = adminList.find(a =>
        ((a.username || '').toLowerCase() === trimmedId ||
         (a.email || '').toLowerCase() === trimmedId) &&
        a.password === password
      );

      // Login root sementara hanya aktif jika query master_admin berhasil
      // dan tabel benar-benar kosong. Tidak aktif saat koneksi/query gagal.
      const emergencyAdmin = isAdminTableEmpty && trimmedId === 'admin' && password === '123'
        ? { id: 'temporary-root-admin', username: 'admin', email: '', nama_lengkap: 'Administrator (Akses Awal)', password: '123' }
        : null;
      const authenticatedAdmin = admin || emergencyAdmin;

      if (authenticatedAdmin) {
        const admin = authenticatedAdmin;
        setIsAdmin(true);
        setCurrentUser(admin);
        setAuthSession({ type: 'admin', id: admin.id });
        localStorage.setItem('suryatokoku_user', JSON.stringify(admin));
        localStorage.setItem('suryatokoku_session', JSON.stringify({ type: 'admin', id: admin.id }));
        localStorage.setItem('suryatokoku_is_admin', 'true');
        navigateTo('#dashboard');
        showToast(`Selamat datang di Workspace HR, ${admin.nama_lengkap}!`);
        return { success: true };
      } else {
        showToast('Username/Email Admin atau password salah!', 'error');
        return { success: false, message: 'Kredensial Admin tidak cocok' };
      }
    };

    const handleLogout = () => {
      setShowLogoutConfirm(true);
    };

    const confirmLogout = () => {
      localStorage.removeItem('suryatokoku_user');
      localStorage.removeItem('suryatokoku_session');
      localStorage.removeItem('suryatokoku_is_admin');
      const wasAdmin = isAdmin;
      setCurrentUser(null);
      setAuthSession(null);
      setIsAdmin(false);
      setShowLogoutConfirm(false);
      if (wasAdmin) {
        navigateTo('#dashboard');
      } else {
        navigateTo('#karyawan');
      }
      showToast('Anda telah berhasil keluar sistem.');
    };

    // Validasi apakah sesi saat ini sesuai dengan portal yang sedang dibuka
    const isEmployeeAuthenticated = authSession?.type === 'karyawan' && currentUser && !isAdmin;
    const isAdminAuthenticated = authSession?.type === 'admin' && currentUser && isAdmin;

    return (
      <div className="min-h-screen flex flex-col justify-between">
        {/* Top Navigation Bar: Ditampilkan kembali */}
        <header className="bg-white/90 backdrop-blur-md border-b border-slate-200/90 text-slate-700 py-2.5 px-3 sm:px-6 flex justify-between items-center z-50 sticky top-0 shadow-xs">
          {/* Brand Logo */}
          <div 
            onClick={() => navigateTo(isDashboardRoute ? '#dashboard' : '#karyawan')}
            className="flex items-center space-x-2 cursor-pointer select-none"
          >
            <img src="Logo SuryaKuCan.png" alt="Logo SuryaKuCan" className="w-7 h-7 object-contain rounded-lg drop-shadow-xs" />
            <div className="flex flex-col">
              <span className="font-extrabold tracking-tight font-display bg-clip-text text-transparent bg-gradient-to-r from-violet-600 to-fuchsia-600 text-xs sm:text-sm">
                SURYATOKOKU HR
              </span>
              <span className="text-[9px] text-slate-400 font-semibold hidden sm:inline -mt-0.5">
                {isDashboardRoute ? 'Admin HR Workspace' : 'Employee Portal'}
              </span>
            </div>
          </div>

          {/* Switcher Tabs Navigasi: Akses Karyawan vs Admin (#dashboard) */}
          <div className="flex items-center bg-slate-100 p-1 rounded-full border border-slate-200/80 shadow-inner">
            <button
              type="button"
              onClick={() => navigateTo('#karyawan')}
              className={`px-3 sm:px-4 py-1.5 rounded-full text-[11px] font-bold transition flex items-center gap-1.5 ${
                !isDashboardRoute
                  ? 'bg-white text-violet-700 shadow-sm border border-slate-200/60'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              <i className="fa-solid fa-mobile-screen text-xs"></i>
              <span>Karyawan</span>
            </button>
            <button
              type="button"
              onClick={() => navigateTo('#dashboard')}
              className={`px-3 sm:px-4 py-1.5 rounded-full text-[11px] font-bold transition flex items-center gap-1.5 ${
                isDashboardRoute
                  ? 'bg-slate-900 text-white shadow-sm'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              <i className="fa-solid fa-shield-halved text-xs"></i>
              <span>Admin</span>
              <span className={`text-[9px] font-mono-numbers px-1.5 py-0.2 rounded-md ${isDashboardRoute ? 'bg-white/20 text-white' : 'bg-slate-200 text-slate-600'}`}>
                #dashboard
              </span>
            </button>
          </div>

          {/* User Status / Logout Button */}
          <div className="flex items-center space-x-2">
            {(isDashboardRoute ? isAdminAuthenticated : isEmployeeAuthenticated) ? (
              <>
                <span className="text-[11px] font-bold text-slate-600 hidden md:inline truncate max-w-[140px]">
                  <i className="fa-regular fa-user mr-1 text-violet-500"></i>{currentUser.nama_lengkap}
                </span>
                <button 
                  onClick={handleLogout}
                  className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 active:scale-95 font-bold rounded-full transition flex items-center gap-1 text-[11px] border border-rose-200 shadow-xs"
                  title="Keluar dari sesi ini"
                >
                  <i className="fa-solid fa-arrow-right-from-bracket text-xs"></i>
                  <span className="hidden sm:inline">Keluar</span>
                </button>
              </>
            ) : (
              <span className="px-2.5 py-1 bg-amber-50 border border-amber-200/80 rounded-full text-[10px] font-bold text-amber-700 hidden sm:flex items-center gap-1">
                <i className="fa-solid fa-lock text-[9px]"></i>
                <span>Silakan Masuk</span>
              </span>
            )}
          </div>
        </header>

        {/* Main Content: Render portal aktif atau form login yang sesuai */}
        <main className="flex-1 relative z-10">
          {currentRoute === '#daftar-karyawan' ? (
            <PendaftaranKaryawanPublik aktif={pendaftaranKaryawanAktif} navigateTo={navigateTo} showToast={showToast} />
          ) : isDashboardRoute ? (
            /* Rute #dashboard */
            isAdminAuthenticated ? (
              <AdminDashboard 
                karyawanList={karyawanList}
                setKaryawanList={setKaryawanListSynced}
                presensiList={presensiList}
                setPresensiList={setPresensiList}
                lokasiList={lokasiList}
                setLokasiList={setLokasiListSynced}
                divisiList={divisiList}
                setDivisiList={setDivisiListSynced}
                jabatanList={jabatanList}
                setJabatanList={setJabatanListSynced}
                shiftList={shiftList}
                setShiftList={setShiftListSynced}
                pengumumanList={pengumumanList}
                setPengumumanList={setPengumumanListSynced}
                adminList={adminList}
                setAdminList={setAdminListSynced}
                lemburList={lemburList}
                setLemburList={setLemburListSynced}
                klaimList={klaimList}
                setKlaimList={setKlaimListSynced}
                toleransiMenit={toleransiMenit}
                setToleransiMenit={setToleransiMenitSynced}
                pendaftaranKaryawanAktif={pendaftaranKaryawanAktif}
                setPendaftaranKaryawanAktif={setPendaftaranKaryawanAktifSynced}
                moreAppsAccess={moreAppsAccess}
                setMoreAppsAccess={setMoreAppsAccessSynced}
                moreAppsCatalog={moreAppsCatalog}
                setMoreAppsCatalog={setMoreAppsCatalogSynced}
                showToast={showToast}
                handleLogout={handleLogout}
                currentUser={currentUser}
              />
            ) : (
              /* Belum Login Admin pada #dashboard: Tampilkan Form Login Admin */
              <AuthPortalView 
                forcedRole="admin"
                karyawanList={karyawanList}
                adminList={adminList}
                onLoginKaryawan={handleLoginKaryawan}
                onLoginAdmin={handleLoginAdmin}
                showToast={showToast}
                navigateTo={navigateTo}
                pendaftaranKaryawanAktif={pendaftaranKaryawanAktif}
              />
            )
          ) : (
            /* Rute Karyawan (Halaman Utama / #karyawan / #profil dll) */
            isEmployeeAuthenticated ? (
              <KaryawanPortal 
                currentUser={currentUser}
                moreAppsAccess={moreAppsAccess}
                moreAppsCatalog={moreAppsCatalog}
                setCurrentUser={setCurrentUser}
                setKaryawanList={setKaryawanList}
                karyawanList={karyawanList}
                presensiList={presensiList}
                setPresensiList={setPresensiList}
                isPresensiReady={isPresensiReady}
                lokasiList={lokasiList}
                divisiList={divisiList}
                jabatanList={jabatanList}
                shiftList={shiftList}
                pengumumanList={pengumumanList}
                lemburList={lemburList}
                setLemburList={setLemburListSynced}
                klaimList={klaimList}
                setKlaimList={setKlaimListSynced}
                toleransiMenit={toleransiMenit}
                currentRoute={currentRoute}
                navigateTo={navigateTo}
                showToast={showToast}
                handleLogout={handleLogout}
              />
            ) : (
              /* Belum Login Karyawan di Halaman Utama: Tampilkan Form Login Karyawan */
              <AuthPortalView 
                forcedRole="karyawan"
                karyawanList={karyawanList}
                adminList={adminList}
                onLoginKaryawan={handleLoginKaryawan}
                onLoginAdmin={handleLoginAdmin}
                showToast={showToast}
                navigateTo={navigateTo}
                pendaftaranKaryawanAktif={pendaftaranKaryawanAktif}
              />
            )
          )}
        </main>

        {/* Modal Konfirmasi Logout yang Ramah Pengguna */}
        {showLogoutConfirm && (
          <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <div className="bg-white rounded-[2rem] max-w-xs w-full p-6 text-center shadow-2xl border border-slate-100">
              <div className="w-14 h-14 rounded-full bg-rose-50 text-rose-500 flex items-center justify-center mx-auto mb-4">
                <i className="fa-solid fa-arrow-right-from-bracket text-2xl"></i>
              </div>
              <h3 className="font-display font-bold text-base text-slate-800">Keluar Sistem?</h3>
              <p className="text-xs text-slate-500 mt-1 mb-5">
                Sesi aktif atas nama <span className="font-bold text-slate-700">{currentUser?.nama_lengkap}</span> akan diakhiri.
              </p>
              <div className="flex gap-2">
                <button 
                  onClick={() => setShowLogoutConfirm(false)}
                  className="w-1/2 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-full transition"
                >
                  Batal
                </button>
                <button 
                  onClick={confirmLogout}
                  className="w-1/2 py-3 bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs rounded-full transition shadow-md active:scale-95"
                >
                  Ya, Keluar
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Toast Notification Modern */}
        {toast && (
          <div className="fixed top-16 right-5 z-50 flex items-center space-x-3 bg-white/95 backdrop-blur-xl text-slate-800 px-5 py-3.5 rounded-2xl shadow-[0_8px_30px_rgb(0,0,0,0.12)] border border-slate-100 animate-bounce">
            <i className={`fa-solid ${toast.type === 'success' ? 'fa-circle-check text-emerald-500' : 'fa-circle-exclamation text-rose-500'} text-xl`}></i>
            <div className="text-sm font-semibold pr-2">{toast.message}</div>
            <button onClick={() => setToast(null)} className="text-slate-400 hover:text-slate-600">
              <i className="fa-solid fa-xmark"></i>
            </button>
          </div>
        )}
      </div>
    );
  }

  /* =========================================================
     FORM PENDAFTARAN KARYAWAN PUBLIK
     URL: #daftar-karyawan. Kolom master dikelola HRD.
     ========================================================= */
  function PendaftaranKaryawanPublik({ aktif, navigateTo, showToast }) {
    const [form, setForm] = useState({ nama_lengkap: '', tgl_lahir: '', no_wa: '', username: '', password: '', email: '', provinsi: '', kota: '', kecamatan: '', desa: '', jalan: '', foto_profil: '' });
    const [busy, setBusy] = useState(false);
    const [provinceOptions, setProvinceOptions] = useState([]);
    const [regencyOptions, setRegencyOptions] = useState([]);
    const [districtOptions, setDistrictOptions] = useState([]);
    const [villageOptions, setVillageOptions] = useState([]);
    const [provinceId, setProvinceId] = useState('');
    const [regencyId, setRegencyId] = useState('');
    const [districtId, setDistrictId] = useState('');
    const [villageId, setVillageId] = useState('');
    const [wilayahLoading, setWilayahLoading] = useState(false);
    const [wilayahError, setWilayahError] = useState('');
    const WILAYAH_API = 'https://www.emsifa.com/api-wilayah-indonesia/api';
    const fetchWilayah = async path => { const r = await fetch(`${WILAYAH_API}/${path}.json`); if (!r.ok) throw new Error('Data wilayah tidak dapat dimuat. Periksa koneksi internet.'); return r.json(); };
    useEffect(() => { let alive = true; fetchWilayah('provinces').then(x => { if (alive) setProvinceOptions(x); }).catch(e => { if (alive) setWilayahError(e.message); }); return () => { alive = false; }; }, []);
    const formatDateInput = value => { const d = String(value || '').replace(/\D/g, '').slice(0, 8); return d.length <= 2 ? d : d.length <= 4 ? `${d.slice(0,2)}/${d.slice(2)}` : `${d.slice(0,2)}/${d.slice(2,4)}/${d.slice(4)}`; };
    const toIsoDate = value => { const m = String(value || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/); if (!m) return ''; const [,dd,mm,yyyy]=m; const d=new Date(+yyyy,+mm-1,+dd); return d.getFullYear()===+yyyy && d.getMonth()===+mm-1 && d.getDate()===+dd ? `${yyyy}-${mm}-${dd}` : ''; };
    const updateWilayah = async (level, id) => {
      setWilayahError(''); setWilayahLoading(true);
      try {
        if (level === 'province') { const item=provinceOptions.find(x=>x.id===id); setProvinceId(id); setRegencyId(''); setDistrictId(''); setVillageId(''); setRegencyOptions([]); setDistrictOptions([]); setVillageOptions([]); setForm(p=>({...p,provinsi:item?.name||'',kota:'',kecamatan:'',desa:''})); if(id) setRegencyOptions(await fetchWilayah(`regencies/${id}`)); }
        if (level === 'regency') { const item=regencyOptions.find(x=>x.id===id); setRegencyId(id); setDistrictId(''); setVillageId(''); setDistrictOptions([]); setVillageOptions([]); setForm(p=>({...p,kota:item?.name||'',kecamatan:'',desa:''})); if(id) setDistrictOptions(await fetchWilayah(`districts/${id}`)); }
        if (level === 'district') { const item=districtOptions.find(x=>x.id===id); setDistrictId(id); setVillageId(''); setVillageOptions([]); setForm(p=>({...p,kecamatan:item?.name||'',desa:''})); if(id) setVillageOptions(await fetchWilayah(`villages/${id}`)); }
        if (level === 'village') { const item=villageOptions.find(x=>x.id===id); setVillageId(id); setForm(p=>({...p,desa:item?.name||''})); }
      } catch(e) { setWilayahError(e.message || 'Gagal memuat wilayah.'); } finally { setWilayahLoading(false); }
    };
    const handleSubmit = async e => {
      e.preventDefault();
      if (!aktif) { showToast('Silakan hubungi HRD terlebih dahulu.', 'info'); navigateTo('#karyawan'); return; }
      const tgl = toIsoDate(form.tgl_lahir); const wa=String(form.no_wa||'').trim();
      if (!tgl) return showToast('Tanggal lahir harus menggunakan format dd/mm/yyyy.', 'error');
      if (!/^08\d{8,13}$/.test(wa)) return showToast('Nomor WhatsApp harus diawali 08 dan berisi 10–15 digit.', 'error');
      if (!form.provinsi || !form.kota || !form.kecamatan || !form.desa) return showToast('Lengkapi provinsi, kabupaten/kota, kecamatan, dan desa/kelurahan.', 'error');
      setBusy(true);
      try {
        const client = await getSupabaseClient();
        const { data: existing, error: checkError } = await client.from('master_karyawan').select('id,username,no_wa,email').or(`username.eq.${form.username.trim()},no_wa.eq.${wa},email.eq.${form.email.trim()}`).limit(1);
        if (checkError) throw checkError;
        if ((existing || []).length) throw new Error('Username, nomor WhatsApp, atau email sudah terdaftar.');
        let foto = '';
        if (form.foto_profil.startsWith('data:image')) {
          foto = await uploadToGoogleDriveViaGAS(form.foto_profil, `profil_daftar_${Date.now()}.jpg`, 'profil');
          if (!foto || foto.startsWith('data:image')) throw new Error('Foto profil gagal diunggah. Silakan coba lagi atau hapus foto lalu kirim ulang.');
        }
        const row = {
          id: 'kry-' + Date.now(), nama_lengkap: form.nama_lengkap.trim(), tgl_lahir: tgl,
          no_wa: wa, username: form.username.trim(), password: form.password, email: form.email.trim(),
          alamat: JSON.stringify({provinsi:form.provinsi,kota:form.kota,kecamatan:form.kecamatan,desa:form.desa,jalan:form.jalan}),
          foto_profil: foto, nik: null, lokasi_absen: [], divisi: '', jabatan: '', shifts: [], sisa_libur: 4, sisa_cuti: 0
        };
        const { error } = await client.from('master_karyawan').insert(row);
        if (error) throw error;
        showToast('Pendaftaran berhasil dikirim. Silakan hubungi HRD untuk penempatan divisi, jabatan, dan lokasi absen.', 'success');
        navigateTo('#karyawan');
      } catch(err) { console.error('Pendaftaran karyawan gagal:', err); showToast(`Pendaftaran gagal: ${err.message || 'periksa koneksi dan izin Supabase.'}`, 'error'); }
      finally { setBusy(false); }
    };
    if (!aktif) return <div className="min-h-[calc(100vh-60px)] flex items-center justify-center p-5"><div className="max-w-md w-full bg-white rounded-[2rem] border border-slate-100 shadow-xl p-8 text-center"><div className="w-16 h-16 rounded-2xl bg-amber-50 text-amber-600 flex items-center justify-center mx-auto mb-4"><i className="fa-solid fa-lock text-2xl"></i></div><h2 className="font-display font-extrabold text-xl text-slate-800">Pendaftaran Belum Dibuka</h2><p className="text-sm text-slate-500 mt-2">Silakan hubungi HRD terlebih dahulu.</p><button onClick={()=>navigateTo('#karyawan')} className="mt-6 w-full py-3 rounded-2xl bg-violet-600 text-white font-bold text-sm">Kembali ke Login</button></div></div>;
    return <div className="min-h-[calc(100vh-60px)] flex items-center justify-center p-4 sm:p-6"><div className="max-w-xl w-full bg-white/95 rounded-[2rem] p-6 sm:p-8 border border-slate-100 shadow-xl"><div className="text-center mb-6"><img src="Logo SuryaKuCan.png" alt="Logo SuryaKuCan" className="w-16 h-16 object-contain mx-auto mb-3"/><h2 className="font-display font-extrabold text-2xl text-slate-800">Pendaftaran Karyawan</h2><p className="text-xs text-slate-500 mt-2">Isi data berikut dengan benar. Data penempatan akan dilengkapi oleh HRD.</p></div>
      <form onSubmit={handleSubmit} className="space-y-3"><div className="flex items-center gap-3 p-3 bg-slate-50 border border-slate-200 rounded-2xl"><div className="w-12 h-12 rounded-full bg-white border border-slate-200 overflow-hidden flex items-center justify-center">{form.foto_profil ? <img src={form.foto_profil} alt="Foto profil" className="w-full h-full object-cover"/> : <i className="fa-solid fa-user text-slate-300"/>}</div><div className="flex-1"><label className="text-[10px] font-bold text-slate-500 block mb-1">Foto Profil (opsional)</label><input type="file" accept="image/*" onChange={e=>{const f=e.target.files?.[0]; if(!f)return; const reader=new FileReader(); reader.onloadend=()=>setForm(p=>({...p,foto_profil:String(reader.result||'')})); reader.readAsDataURL(f);}} className="w-full text-[11px]"/></div></div>
      <input required maxLength={120} placeholder="Nama Lengkap" value={form.nama_lengkap} onChange={e=>setForm(p=>({...p,nama_lengkap:e.target.value}))} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 font-bold outline-none"/>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3"><div><label className="text-[10px] text-slate-500 font-bold block mb-1 ml-1">Tanggal Lahir</label><input required placeholder="dd/mm/yyyy" inputMode="numeric" maxLength={10} pattern="\d{2}/\d{2}/\d{4}" value={form.tgl_lahir} onChange={e=>setForm(p=>({...p,tgl_lahir:formatDateInput(e.target.value)}))} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 font-bold outline-none"/></div><div><label className="text-[10px] text-slate-500 font-bold block mb-1 ml-1">No. WhatsApp</label><input required placeholder="08xxxxxxxxxx" inputMode="numeric" value={form.no_wa} onChange={e=>setForm(p=>({...p,no_wa:e.target.value.replace(/\D/g,'').slice(0,15)}))} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 font-bold outline-none"/></div></div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3"><input required placeholder="Username" autoComplete="username" value={form.username} onChange={e=>setForm(p=>({...p,username:e.target.value.trimStart()}))} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 font-bold outline-none"/><input required type="password" minLength={6} autoComplete="new-password" placeholder="Password (min. 6 karakter)" value={form.password} onChange={e=>setForm(p=>({...p,password:e.target.value}))} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 font-bold outline-none"/></div>
      <input required type="email" placeholder="Alamat Email" value={form.email} onChange={e=>setForm(p=>({...p,email:e.target.value.trimStart()}))} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 font-bold outline-none"/>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3"><select required value={provinceId} onChange={e=>updateWilayah('province',e.target.value)} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold"><option value="">Pilih Provinsi</option>{provinceOptions.map(x=><option key={x.id} value={x.id}>{x.name}</option>)}</select><select required disabled={!provinceId||wilayahLoading} value={regencyId} onChange={e=>updateWilayah('regency',e.target.value)} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold"><option value="">Pilih Kabupaten/Kota</option>{regencyOptions.map(x=><option key={x.id} value={x.id}>{x.name}</option>)}</select><select required disabled={!regencyId||wilayahLoading} value={districtId} onChange={e=>updateWilayah('district',e.target.value)} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold"><option value="">Pilih Kecamatan</option>{districtOptions.map(x=><option key={x.id} value={x.id}>{x.name}</option>)}</select><select required disabled={!districtId||wilayahLoading} value={villageId} onChange={e=>updateWilayah('village',e.target.value)} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold"><option value="">Pilih Desa/Kelurahan</option>{villageOptions.map(x=><option key={x.id} value={x.id}>{x.name}</option>)}</select></div>
      <textarea rows={2} placeholder="Alamat jalan / nomor rumah (opsional)" value={form.jalan} onChange={e=>setForm(p=>({...p,jalan:e.target.value}))} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 font-bold outline-none resize-y"/>
      {wilayahLoading&&<p className="text-[10px] text-violet-600">Memuat data wilayah...</p>}{wilayahError&&<p className="text-[10px] text-rose-600">{wilayahError}</p>}
      <div className="bg-amber-50 border border-amber-100 rounded-xl p-3 text-[10px] text-amber-800">Nomor induk karyawan, lokasi absen, divisi, dan jabatan akan diisi atau diperbarui oleh HRD.</div>
      <button disabled={busy} type="submit" className="w-full py-3.5 rounded-2xl bg-gradient-to-r from-violet-600 to-fuchsia-600 text-white text-xs font-bold shadow-lg disabled:opacity-60">{busy?<><i className="fa-solid fa-spinner fa-spin mr-2"/>Mengirim Pendaftaran...</>:<><i className="fa-solid fa-paper-plane mr-2"/>Kirim Pendaftaran</>}</button><button type="button" onClick={()=>navigateTo('#karyawan')} className="w-full py-3 rounded-2xl border border-slate-200 text-slate-600 text-xs font-bold">Kembali ke Login</button></form></div></div>;
  }

  /* =========================================================
     KOMPONEN HALAMAN LOGIN (PORTAL KARYAWAN & ADMIN HRD)
     ========================================================= */
  function AuthPortalView({ forcedRole = 'karyawan', karyawanList, adminList, onLoginKaryawan, onLoginAdmin, showToast, navigateTo, pendaftaranKaryawanAktif = false }) {
    const [identifier, setIdentifier] = useState('');
    const [password, setPassword] = useState('');
    const [showPassword, setShowPassword] = useState(false);
    const [isLoading, setIsLoading] = useState(false);
    const [showRegistrationClosedModal, setShowRegistrationClosedModal] = useState(false);

    const loginRole = forcedRole;

    const handleSubmit = async (e) => {
      e.preventDefault();
      if (!identifier.trim() || !password.trim()) {
        showToast('Harap isi semua kolom login', 'error');
        return;
      }
      setIsLoading(true);
      try {
        if (loginRole === 'karyawan') {
          await onLoginKaryawan(identifier, password);
        } else {
          await onLoginAdmin(identifier, password);
        }
      } finally {
        setIsLoading(false);
      }
    };

    // Quick Login Demo Helper
    const applyDummyCredential = (user) => {
      setIdentifier(user.username);
      setPassword(user.password);
      showToast(`Kredensial ${user.nama_lengkap} disiapkan. Silakan klik Masuk.`, 'info');
    };

    return (
      <div className="min-h-[calc(100vh-60px)] flex items-center justify-center p-4 sm:p-6 relative z-10">
        <div className="max-w-md w-full bg-white/90 backdrop-blur-2xl rounded-[2.5rem] p-7 sm:p-9 border border-white/60 shadow-[0_25px_50px_-12px_rgba(0,0,0,0.08)] relative overflow-hidden">
          
          {/* Logo & Header Title */}
          <div className="text-center mb-6">
            <div className="w-16 h-16 mx-auto mb-3.5 flex items-center justify-center bg-white rounded-2xl p-2 shadow-md border border-slate-100">
              <img src="Logo SuryaKuCan.png" alt="Logo SuryaKuCan" className="w-full h-full object-contain" />
            </div>
            <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[10px] font-bold mb-2 border uppercase tracking-wider bg-slate-50 text-slate-600 border-slate-200">
              <span className={`w-2 h-2 rounded-full ${loginRole === 'admin' ? 'bg-indigo-600' : 'bg-violet-600'}`}></span>
              {loginRole === 'admin' ? 'Login Administrator' : 'Login Karyawan'}
            </div>
            <h2 className="font-display font-extrabold text-2xl text-slate-800 tracking-tight">
              {loginRole === 'admin' ? 'SURYAKUCAN' : 'SURYAKUCAN'}
            </h2>
            <p className="text-xs text-slate-500 font-medium mt-1">
              {loginRole === 'admin' 
                ? 'Akses Panel Administrasi Operational Toko' 
                : 'Sistem Digital Presensi & Manajemen Operasional Toko'}
            </p>
          </div>

          {/* Form Login */}
          <form onSubmit={handleSubmit} className="space-y-4">
            <div>
              <label className="text-[11px] font-bold text-slate-500 block mb-1.5 ml-1">
                {loginRole === 'karyawan' ? 'Username / No. WhatsApp / Email' : 'Username / Email Administrator'}
              </label>
              <div className="relative">
                <i className="fa-regular fa-user absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
                <input 
                  type="text" 
                  required
                  placeholder={loginRole === 'karyawan' ? "Ex : Awil / 0838xxxx / awilistrators@gmail.com" : "admin atau hrd@suryatokoku.com"}
                  value={identifier}
                  onChange={(e) => setIdentifier(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl pl-10 pr-4 py-3.5 font-bold text-slate-700 outline-none focus:border-violet-500 focus:bg-white transition"
                />
              </div>
            </div>

            <div>
              <div className="flex justify-between items-center mb-1.5 ml-1">
                <label className="text-[11px] font-bold text-slate-500">Kata Sandi</label>
                <span
                  className="text-[10px] text-violet-600 font-semibold cursor-pointer hover:text-violet-800 transition"
                  role="button"
                  tabIndex={0}
                  onClick={() => {
                    if (loginRole === 'karyawan') {
                      const pesan = 'Permisi Bu.. App presensi aku lupa kata sandi nya, mohon bantu reset ya bu 🙏';
                      window.open(`https://wa.me/6285324667155?text=${encodeURIComponent(pesan)}`, '_blank', 'noopener,noreferrer');
                    } else {
                      showToast('Silakan hubungi administrator untuk reset kata sandi.', 'info');
                    }
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault();
                      if (loginRole === 'karyawan') {
                        const pesan = 'Permisi Bu.. App presensi aku lupa kata sandi nya, mohon bantu reset ya bu 🙏';
                        window.open(`https://wa.me/6285324667155?text=${encodeURIComponent(pesan)}`, '_blank', 'noopener,noreferrer');
                      } else {
                        showToast('Silakan hubungi administrator untuk reset kata sandi.', 'info');
                      }
                    }
                  }}
                >
                  Lupa Sandi?
                </span>
              </div>
              <div className="relative">
                <i className="fa-solid fa-lock absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
                <input 
                  type={showPassword ? 'text' : 'password'}
                  required
                  placeholder="Masukan kata sandi, bukan kata awil..."
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl pl-10 pr-11 py-3.5 font-bold text-slate-700 outline-none focus:border-violet-500 focus:bg-white transition font-mono-numbers"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute right-3.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1"
                >
                  <i className={`fa-solid ${showPassword ? 'fa-eye-slash' : 'fa-eye'} text-xs`}></i>
                </button>
              </div>
            </div>

            <button 
              type="submit"
              disabled={isLoading}
              className={`w-full py-3.5 mt-2 text-white rounded-2xl font-bold text-xs shadow-lg active:scale-95 transition flex items-center justify-center gap-2 ${
                loginRole === 'admin' 
                  ? 'bg-slate-900 hover:bg-slate-800 shadow-slate-300' 
                  : 'bg-gradient-to-r from-violet-600 to-fuchsia-600 hover:from-violet-700 hover:to-fuchsia-700 shadow-violet-200'
              }`}
            >
              {isLoading ? (
                <>
                  <i className="fa-solid fa-circle-notch fa-spin text-sm"></i>
                  <span>Memverifikasi...</span>
                </>
              ) : (
                <>
                  <i className="fa-solid fa-right-to-bracket text-xs"></i>
                  <span>Masuk</span>
                </>
              )}
            </button>
          </form>

          {loginRole === 'admin' && (
            <p className="mt-4 text-[10px] text-slate-400 text-center">
              <i className="fa-regular fa-copyright mr-1"></i>Operational Development Suryatokoku 2026
            </p>
          )}

          {loginRole === 'karyawan' && (
            <button
              type="button"
              onClick={() => {
                if (pendaftaranKaryawanAktif) navigateTo('#daftar-karyawan');
                else setShowRegistrationClosedModal(true);
              }}
              className="w-full mt-3 py-3.5 px-4 rounded-2xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 font-bold text-xs transition active:scale-95 flex items-center justify-center gap-2"
            >
              <i className="fa-solid fa-user-plus text-sm"></i>
              <span>Daftar Akun</span>
            </button>
          )}

          {loginRole === 'karyawan' && showRegistrationClosedModal && (
            <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-950/45 backdrop-blur-sm" onClick={() => setShowRegistrationClosedModal(false)}>
              <div className="w-full max-w-sm rounded-[2rem] bg-white p-6 sm:p-7 shadow-2xl border border-white/80 animate-in fade-in zoom-in-95 duration-200" onClick={e => e.stopPropagation()}>
                <div className="w-16 h-16 rounded-2xl bg-gradient-to-br from-violet-100 to-fuchsia-100 text-violet-600 flex items-center justify-center mx-auto mb-4 shadow-inner"><i className="fa-solid fa-user-lock text-2xl"></i></div>
                <div className="text-center">
                  <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-amber-50 text-amber-700 text-[10px] font-extrabold tracking-wide mb-3"><i className="fa-solid fa-circle-info"></i> PENDAFTARAN DITUTUP</span>
                  <h3 className="font-display font-extrabold text-xl text-slate-800">Pendaftaran Belum Dibuka</h3>
                  <p className="text-sm leading-relaxed text-slate-500 mt-2">Silakan hubungi HRD terlebih dahulu untuk mendapatkan bantuan pendaftaran karyawan.</p>
                </div>
                <button type="button" onClick={() => setShowRegistrationClosedModal(false)} className="w-full mt-6 py-3.5 rounded-2xl bg-gradient-to-r from-violet-600 to-fuchsia-600 hover:from-violet-700 hover:to-fuchsia-700 text-white text-xs font-extrabold shadow-lg shadow-violet-200 transition active:scale-[0.98]"><i className="fa-solid fa-check mr-2"></i>Mengerti</button>
              </div>
            </div>
          )}

          {loginRole === 'karyawan' && (
            <div className="mt-4 pt-4 border-t border-slate-100">
              <button
                type="button"
                onClick={() => requestPwaInstall(showToast)}
                className="w-full py-3.5 px-4 rounded-2xl border border-violet-200 bg-violet-50 hover:bg-violet-100 text-violet-700 font-bold text-xs transition active:scale-95 flex items-center justify-center gap-2"
              >
                <i className="fa-solid fa-download text-sm"></i>
                <span>Install SuryaKuCan</span>
              </button>
              <a
                href="https://suryatokoku.my.id/suryakucan/SuryaKuCan.apk"
                download="SuryaKuCan.apk"
                className="w-full mt-2 py-3.5 px-4 rounded-2xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 font-bold text-xs transition active:scale-95 flex items-center justify-center gap-2"
              >
                <i className="fa-brands fa-android text-sm text-emerald-600"></i>
                <span>Download App SuryaKuCan</span>
              </a>
              <p className="text-[10px] text-slate-400 text-center mt-2">
                Pasang aplikasi agar lebih mudah dibuka dari layar utama.
              <span className="block mt-2"><i className="fa-regular fa-copyright mr-1"></i>Operational Development Suryatokoku 2026</span>
              </p>
            </div>
          )}

        </div>
      </div>
    );
  }

    function KaryawanPortal({
      currentUser,
      moreAppsAccess = {},
      moreAppsCatalog = DEFAULT_MORE_APPS,
      setCurrentUser,
      setKaryawanList,
      karyawanList,
      presensiList,
      setPresensiList,
      isPresensiReady,
      lokasiList,
      divisiList,
      shiftList,
      pengumumanList,
      lemburList,
      setLemburList,
      klaimList,
      setKlaimList,
      toleransiMenit,
      currentRoute,
      navigateTo,
      showToast,
      handleLogout
    }) {
      const [clock, setClock] = useState(new Date());
      const [showYesterdayCheck, setShowYesterdayCheck] = useState(false);
      const [selectedClaimShift, setSelectedClaimShift] = useState('');
      const [selectedAnnouncement, setSelectedAnnouncement] = useState(null);
      const [showWizardModal, setShowWizardModal] = useState(false);
      const [wizardType, setWizardType] = useState('masuk'); 
      const [showCutiLockModal, setShowCutiLockModal] = useState(false);
      const [showLiburConfirmModal, setShowLiburConfirmModal] = useState(false);
      const [showSimpleFormModal, setShowSimpleFormModal] = useState(null); 
      const [filterStatusRiwayat, setFilterStatusRiwayat] = useState('Semua');
      const [cutoffFilter, setCutoffFilter] = useState('Semua');
      const [showComingSoonApp, setShowComingSoonApp] = useState('');

      // Ambil ulang data master karyawan dari Supabase agar perubahan jatah oleh HR langsung tercermin,
      // tanpa perlu menghapus cookie atau data situs.
      useEffect(() => {
        let active = true;
        const refreshEmployeeMaster = async () => {
          try {
            const client = await getSupabaseClient();
            if (!client || !currentUser?.id) return;
            const { data, error } = await client.from('master_karyawan')
              .select('*').eq('id', currentUser.id).maybeSingle();
            if (error) throw error;
            if (active && data) {
              setCurrentUser(prev => ({ ...prev, ...data }));
              if (typeof setKaryawanList === 'function') {
                setKaryawanList(prev => (Array.isArray(prev) ? prev : []).map(k =>
                  String(k.id) === String(data.id) ? { ...k, ...data } : k
                ));
              }
              try { localStorage.setItem('suryatokoku_user', JSON.stringify({ ...currentUser, ...data })); } catch (_) {}
            }
          } catch (err) { console.warn('Gagal menyegarkan data master karyawan:', err); }
        };
        refreshEmployeeMaster();
        const onFocus = () => refreshEmployeeMaster();
        window.addEventListener('focus', onFocus);
        window.addEventListener('pageshow', onFocus);
        return () => { active = false; window.removeEventListener('focus', onFocus); window.removeEventListener('pageshow', onFocus); };
      }, [currentUser?.id]);

      useEffect(() => {
        const timer = setInterval(() => setClock(new Date()), 1000);
        return () => clearInterval(timer);
      }, []);

      useEffect(() => {
        // Jangan pernah membuka popup hanya berdasarkan daftar lokal yang mungkin masih kosong.
        // Tunggu muatan awal selesai, lalu verifikasi langsung ke Supabase untuk tanggal kemarin.
        if (!isPresensiReady || !currentUser?.id) return;

        let cancelled = false;
        const now = new Date();
        const todayKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
        // Versi key dinaikkan supaya hasil pemeriksaan lama tidak memengaruhi logika yang baru.
        const checkKey = `yesterday_checked_v2_${currentUser.id}_${todayKey}`;
        try {
          if (sessionStorage.getItem(checkKey)) return;
        } catch (_) {}

        const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
        const yStr = `${yesterday.getFullYear()}-${String(yesterday.getMonth() + 1).padStart(2, '0')}-${String(yesterday.getDate()).padStart(2, '0')}`;

        (async () => {
          try {
            const client = await getSupabaseClient();
            const { data, error } = await client
              .from('presensi')
              .select('id, karyawan_id, tanggal, status, tipe')
              .eq('karyawan_id', currentUser.id)
              .eq('tanggal', yStr)
              .limit(1);

            // Jika query gagal, jangan menampilkan popup dan jangan menandai pemeriksaan selesai.
            // Ini mencegah prompt keliru saat koneksi/RLS/data Supabase belum siap.
            if (cancelled || error) {
              if (error) console.warn('Verifikasi presensi kemarin gagal; popup ditunda:', error.message);
              return;
            }

            // Query langsung ke Supabase adalah sumber utama. Fallback daftar state menangani
            // perbedaan tipe ID atau format tanggal pada data yang sudah dimuat.
            const hasDirectRecord = Array.isArray(data) && data.length > 0;
            const hasLoadedRecord = presensiList.some(p =>
              String(p.karyawan_id) === String(currentUser.id) &&
              String(p.tanggal || '').slice(0, 10) === yStr
            );
            // Jika karyawan sudah mengajukan klaim untuk kemarin, jangan munculkan popup lagi.
            const { data: claimRows, error: claimError } = await client
              .from('klaim_kehadiran')
              .select('id, karyawan_id, tanggal_klaim')
              .eq('karyawan_id', currentUser.id)
              .eq('tanggal_klaim', yStr)
              .limit(1);
            if (cancelled || claimError) {
              if (claimError) console.warn('Verifikasi klaim kemarin gagal; popup ditunda:', claimError.message);
              return;
            }
            const hasDirectClaim = Array.isArray(claimRows) && claimRows.length > 0;
            const hasLoadedClaim = (Array.isArray(klaimList) ? klaimList : []).some(c =>
              String(c.karyawan_id) === String(currentUser.id) && String(c.tanggal_klaim || '').slice(0, 10) === yStr
            );

            if (!cancelled && !hasDirectRecord && !hasLoadedRecord && !hasDirectClaim && !hasLoadedClaim) {
              setShowYesterdayCheck(true);
            } else if (!cancelled) {
              try { sessionStorage.setItem(checkKey, 'true'); } catch (_) {}
              setShowYesterdayCheck(false);
            }
          } catch (err) {
            console.warn('Belum dapat memverifikasi presensi kemarin; popup ditunda:', err);
          }
        })();

        return () => { cancelled = true; };
      }, [currentUser?.id, isPresensiReady, klaimList, presensiList]);

      const payrollCycle = useMemo(() => {
        const now = new Date();
        const curDate = now.getDate();
        let startDate, endDate;
        if (curDate > 25) {
          startDate = new Date(now.getFullYear(), now.getMonth(), 26);
          endDate = new Date(now.getFullYear(), now.getMonth() + 1, 25);
        } else {
          startDate = new Date(now.getFullYear(), now.getMonth() - 1, 26);
          endDate = new Date(now.getFullYear(), now.getMonth(), 25);
        }
        return { startDate, endDate, label: `${startDate.getDate()} ${startDate.toLocaleDateString('id-ID', {month: 'short'})} - ${endDate.getDate()} ${endDate.toLocaleDateString('id-ID', {month: 'short', year: 'numeric'})}` };
      }, []);

      const isDateInPayrollCycle = (dateStr) => {
        if (!dateStr) return false;
        const [y, m, d] = dateStr.split('-').map(Number);
        const recordDate = new Date(y, m - 1, d);
        const s = new Date(payrollCycle.startDate); s.setHours(0, 0, 0, 0);
        const e = new Date(payrollCycle.endDate); e.setHours(23, 59, 59, 999);
        return recordDate >= s && recordDate <= e;
      };

      const periodActivities = useMemo(() => {
        const userRecords = presensiList.filter(p => String(p.karyawan_id) === String(currentUser.id) && isDateInPayrollCycle(String(p.tanggal || '').slice(0, 10)));
        const userPendingClaims = (Array.isArray(klaimList) ? klaimList : []).filter(c =>
          String(c.karyawan_id) === String(currentUser.id) &&
          ['pending', 'menunggu', ''].includes(String(c.status || 'Pending').toLowerCase()) &&
          isDateInPayrollCycle(String(c.tanggal_klaim || '').slice(0, 10))
        );
        const uniqueDates = [...new Set([...userRecords.map(p => String(p.tanggal || '').slice(0, 10)), ...userPendingClaims.map(c => String(c.tanggal_klaim || '').slice(0, 10))])].filter(Boolean).sort((a, b) => b.localeCompare(a));

        return uniqueDates.map(date => {
          const masukRec = userRecords.find(p => String(p.tanggal || '').slice(0, 10) === date && p.tipe === 'masuk');
          const keluarRec = userRecords.find(p => String(p.tanggal || '').slice(0, 10) === date && p.tipe === 'keluar');
          const pendingClaim = !masukRec && !keluarRec ? userPendingClaims.find(c => String(c.tanggal_klaim || '').slice(0, 10) === date) : null;
          const status = pendingClaim ? 'Pending' : (masukRec ? masukRec.status : (keluarRec ? keluarRec.status : 'Hadir'));
          const shift = pendingClaim?.shift_nama || masukRec?.shift_nama || keluarRec?.shift_nama || 'Reguler';
          const statusTanpaJam = pendingClaim || ['Libur', 'Sakit', 'Cuti', 'Izin'].includes(String(status || '').trim());
          const jamMasuk = statusTanpaJam ? '-' : (masukRec ? (masukRec.jam || '-') : '-');
          const jamPulang = statusTanpaJam ? '-' : (keluarRec ? (keluarRec.jam || '-') : '-');
          return {
            id: pendingClaim?.id || masukRec?.id || keluarRec?.id || date,
            tanggal: date, status, shift, jamMasuk, jamPulang,
            keterangan: pendingClaim ? 'Menunggu persetujuan HR' : (masukRec?.keterangan || keluarRec?.keterangan || '')
          };
        });
      }, [presensiList, klaimList, currentUser.id, payrollCycle]);

      const filteredPeriodActivities = useMemo(() => {
        if (cutoffFilter === 'Semua') return periodActivities;
        if (cutoffFilter === 'Masuk') return periodActivities.filter(a => a.status === 'Hadir' || a.status === 'Terlambat');
        return periodActivities.filter(a => a.status === cutoffFilter);
      }, [periodActivities, cutoffFilter]);

      const employeeStats = useMemo(() => {
        const userAttendances = presensiList.filter(p => p.karyawan_id === currentUser.id && isDateInPayrollCycle(p.tanggal));
        const masuk = userAttendances.filter(p => p.tipe === 'masuk' && (p.status === 'Hadir' || p.status === 'Terlambat')).length;
        const libur = userAttendances.filter(p => p.status === 'Libur').length;
        const izin = userAttendances.filter(p => p.status === 'Izin').length;
        const sakit = userAttendances.filter(p => p.status === 'Sakit').length;
        const cuti = userAttendances.filter(p => p.status === 'Cuti').length;
        return { masuk, libur, izin, sakit, cuti };
      }, [presensiList, currentUser.id, payrollCycle]);

      const todayStr = getLocalDateKey();
      const todayAttendanceMasuk = presensiList.find(p => p.karyawan_id === currentUser.id && p.tanggal === todayStr && p.tipe === 'masuk');
      const todayAttendanceKeluar = presensiList.find(p => p.karyawan_id === currentUser.id && p.tanggal === todayStr && p.tipe === 'keluar');
      const todayIsLibur = presensiList.some(p =>
        String(p.karyawan_id) === String(currentUser.id) && p.tanggal === todayStr && p.status === 'Libur'
      );
      // Satu status non-kehadiran saja yang boleh dicatat per karyawan per hari.
      // Setelah Sakit/Izin/Cuti/Libur tercatat, semua tombol status tersebut dikunci.
      const todaySpecialStatusRecord = presensiList.find(p =>
        String(p.karyawan_id) === String(currentUser.id) &&
        p.tanggal === todayStr &&
        ['Sakit', 'Izin', 'Cuti', 'Libur'].includes(String(p.status || '').trim())
      );
      const todayHasStatusAction = !!todaySpecialStatusRecord;
      // sisa_libur di master karyawan adalah jatah/acuan untuk periode aktif;
      // sisa aktual dihitung dari jatah dikurangi catatan Libur pada periode payroll.
      const sisaLiburAktif = Math.max(0, Number(currentUser.sisa_libur ?? 4) - employeeStats.libur);

      const userDivisiInfo = divisiList.find(d => d.nama_divisi === currentUser.divisi);
      const allowCuti = userDivisiInfo ? userDivisiInfo.allow_cuti : true;

      const handleActionClick = (type) => {
        if (type === 'lembur' && todayHasStatusAction) {
          showToast(`Pengajuan lembur tidak dapat dilakukan karena status hari ini sudah tercatat ${todaySpecialStatusRecord.status}.`, 'info');
          return;
        }
        if (['sakit', 'libur', 'izin', 'cuti'].includes(type) && todayHasStatusAction) {
          showToast(`Status hari ini sudah tercatat ${todaySpecialStatusRecord.status}. Pilihan Sakit, Libur, Izin, dan Cuti tidak dapat diubah lagi.`, 'info');
          return;
        }
        if (type === 'libur') {
          if (todayAttendanceMasuk || todayAttendanceKeluar) {
            showToast('Presensi hari ini sudah tercatat. Hubungi admin jika perlu koreksi.', 'error');
            return;
          }
          if (sisaLiburAktif <= 0) {
            showToast('Jatah libur sudah habis. Silakan gunakan Izin jika diperlukan.', 'info');
            return;
          }
          setShowLiburConfirmModal(true);
          return;
        }
        if (type === 'cuti' && !allowCuti) {
          setShowCutiLockModal(true);
          return;
        }
        setShowSimpleFormModal(type);
      };

      const handleMarkLibur = async () => {
        if (todayIsLibur) {
          showToast('Status hari ini sudah tercatat Libur.', 'info');
          return;
        }
        if (sisaLiburAktif <= 0) {
          showToast('Jatah libur sudah habis. Silakan gunakan Izin jika diperlukan.', 'info');
          return;
        }
        if (todayAttendanceMasuk || todayAttendanceKeluar) {
          showToast('Presensi hari ini sudah tercatat. Hubungi admin jika perlu koreksi.', 'error');
          return;
        }

        const now = new Date();
        const newRecord = {
          id: 'prs-' + Date.now(),
          karyawan_id: currentUser.id,
          karyawan_nama: currentUser.nama_lengkap,
          divisi: currentUser.divisi,
          tipe: 'masuk',
          status: 'Libur',
          shift_nama: 'Libur',
          lokasi_nama: 'Libur',
          latitude: 0,
          longitude: 0,
          foto_url: '',
          keterangan: 'Libur karyawan (menggunakan jatah libur)',
          tanggal: todayStr,
          jam: [now.getHours(), now.getMinutes(), now.getSeconds()].map(v => String(v).padStart(2, '0')).join(':'),
          created_at: now.toISOString()
        };

        try {
          const client = await getSupabaseClient();
          const { error } = await client.from('presensi').insert([newRecord]);
          if (error) throw error;
          setPresensiList(prev => [newRecord, ...prev.filter(p => !(String(p.karyawan_id) === String(currentUser.id) && p.tanggal === todayStr && p.status === 'Libur'))]);
          showToast('Status Libur berhasil dicatat. Jatah libur otomatis berkurang.');
        } catch (error) {
          console.error('Gagal mencatat status Libur:', error);
          showToast('Status Libur gagal disimpan. Silakan coba lagi atau hubungi admin.', 'error');
        }
      };

      const handleYesterdayConfirm = async (status, note) => {
        // Simpan ke tanggal kalender kemarin berdasarkan waktu lokal; jangan hanya mengubah state React.
        const nowLocal = new Date();
        const yesterday = new Date(nowLocal.getFullYear(), nowLocal.getMonth(), nowLocal.getDate() - 1);
        const yStr = `${yesterday.getFullYear()}-${String(yesterday.getMonth() + 1).padStart(2, '0')}-${String(yesterday.getDate()).padStart(2, '0')}`;

        // Cegah duplikasi dengan memeriksa Supabase langsung sebelum menyimpan.
        try {
          const client = await getSupabaseClient();
          const { data: existingRows, error: checkError } = await client
            .from('presensi')
            .select('id, karyawan_id, tanggal, status, tipe')
            .eq('karyawan_id', currentUser.id)
            .eq('tanggal', yStr)
            .limit(1);

          if (checkError) throw checkError;
          if (Array.isArray(existingRows) && existingRows.length > 0) {
            const existing = existingRows[0];
            // Catatan sudah ada di database; cukup tutup popup. Jangan memasukkan
            // object parsial hasil query ke state riwayat.
            setShowYesterdayCheck(false);
            try {
              const now = new Date();
              const todayKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
              sessionStorage.setItem(`yesterday_checked_v2_${currentUser.id}_${todayKey}`, 'true');
            } catch (_) {}
            showToast(`Presensi tanggal ${yStr} sudah tercatat sebagai ${existing.status || 'presensi'}.`, 'info');
            return;
          }

          const now = new Date();
          const newRecord = {
            id: 'prs-' + Date.now(),
            karyawan_id: currentUser.id,
            karyawan_nama: currentUser.nama_lengkap,
            divisi: currentUser.divisi,
            tipe: 'masuk',
            status,
            shift_nama: 'Konfirmasi Mandiri',
            lokasi_nama: 'Dikonfirmasi Kemarin',
            latitude: 0,
            longitude: 0,
            foto_url: '',
            keterangan: note || 'Konfirmasi presensi kemarin',
            tanggal: yStr,
            jam: '08:00:00',
            created_at: now.toISOString()
          };

          const { error: insertError } = await client.from('presensi').insert([newRecord]);
          if (insertError) throw insertError;

          // Update UI hanya setelah Supabase mengonfirmasi insert berhasil.
          setPresensiList(prev => [
            newRecord,
            ...prev.filter(p => !(
              String(p.karyawan_id) === String(currentUser.id) &&
              String(p.tanggal || '').slice(0, 10) === yStr
            ))
          ]);

          if (status === 'Libur' && Number(currentUser.sisa_libur) > 0) {
            setCurrentUser(prev => ({ ...prev, sisa_libur: Number(prev.sisa_libur) - 1 }));
          } else if (status === 'Cuti' && Number(currentUser.sisa_cuti) > 0) {
            setCurrentUser(prev => ({ ...prev, sisa_cuti: Number(prev.sisa_cuti) - 1 }));
          }

          setShowYesterdayCheck(false);
          try {
            const nowToday = new Date();
            const todayKey = `${nowToday.getFullYear()}-${String(nowToday.getMonth() + 1).padStart(2, '0')}-${String(nowToday.getDate()).padStart(2, '0')}`;
            sessionStorage.setItem(`yesterday_checked_v2_${currentUser.id}_${todayKey}`, 'true');
          } catch (_) {}
          showToast(`Presensi tanggal ${yStr} berhasil disimpan sebagai ${status}.`);
        } catch (error) {
          console.error('Gagal menyimpan konfirmasi presensi kemarin ke Supabase:', error);
          showToast(`Presensi belum tersimpan ke Supabase: ${error?.message || 'periksa koneksi, izin akses, dan struktur tabel.'}`, 'error');
          // Popup tetap terbuka agar karyawan bisa mencoba lagi setelah masalah diperbaiki.
        }
      };

      const allowedClaimShifts = (() => {
        const division = (Array.isArray(divisiList) ? divisiList : []).find(d =>
          String(d.nama_divisi || '').trim().toLocaleLowerCase('id') === String(currentUser?.divisi || '').trim().toLocaleLowerCase('id')
        );
        let names = [];
        const configured = division?.shifts;
        if (Array.isArray(configured)) names = configured.map(item => typeof item === 'string' ? item : item?.nama_shift).filter(Boolean);
        else if (typeof configured === 'string' && configured.trim()) {
          try { const parsed = JSON.parse(configured); names = Array.isArray(parsed) ? parsed.map(item => typeof item === 'string' ? item : item?.nama_shift).filter(Boolean) : []; }
          catch (_) { names = configured.split(',').map(x => x.trim()).filter(Boolean); }
        }
        const allowed = new Set(names.map(name => String(name).trim().toLocaleLowerCase('id')));
        return (Array.isArray(shiftList) ? shiftList : []).filter(shift => allowed.has(String(shift.nama_shift || '').trim().toLocaleLowerCase('id'))).sort((a,b) => String(a.jam_masuk || '').localeCompare(String(b.jam_masuk || '')));
      })();

      const handleYesterdayClaimHR = async () => {
        if (!selectedClaimShift) {
          showToast('Pilih shift yang Anda jalani kemarin terlebih dahulu.', 'error');
          return;
        }
        const nowLocal = new Date();
        const yesterday = new Date(nowLocal.getFullYear(), nowLocal.getMonth(), nowLocal.getDate() - 1);
        const yStr = `${yesterday.getFullYear()}-${String(yesterday.getMonth() + 1).padStart(2, '0')}-${String(yesterday.getDate()).padStart(2, '0')}`;
        const selectedShift = allowedClaimShifts.find(s => String(s.nama_shift) === String(selectedClaimShift));
        if (!selectedShift) {
          showToast('Shift yang dipilih tidak ditemukan pada Master Shift divisi Anda.', 'error');
          return;
        }
        const claim = {
          id: 'klm-' + Date.now(),
          karyawan_id: currentUser.id,
          karyawan_nama: currentUser.nama_lengkap,
          divisi: currentUser.divisi,
          tanggal_klaim: yStr,
          shift_nama: selectedShift.nama_shift,
          shift_jam_masuk: selectedShift.jam_masuk || '',
          shift_jam_pulang: selectedShift.jam_pulang || '',
          status: 'Pending',
          catatan: 'Menunggu verifikasi HR.',
          created_at: new Date().toISOString()
        };
        try {
          const client = await getSupabaseClient();
          const { error } = await client.from('klaim_kehadiran').insert([claim]);
          if (error) throw error;
          setKlaimList([claim, ...(Array.isArray(klaimList) ? klaimList : [])]);
          setShowYesterdayCheck(false);
          try {
            const today = new Date();
            const todayKey = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
            sessionStorage.setItem(`yesterday_checked_v2_${currentUser.id}_${todayKey}`, 'true');
          } catch (_) {}
          showToast('Klaim kehadiran berhasil dikirim ke HRD. Pop-up tidak akan muncul lagi untuk tanggal ini.');
        } catch (error) {
          console.error('Gagal mengirim klaim kehadiran:', error);
          showToast(`Klaim gagal dikirim: ${error?.message || 'periksa koneksi dan izin tabel klaim_kehadiran.'}`, 'error');
        }
      };

      return (
        <div className="max-w-md mx-auto min-h-screen bg-slate-50/70 flex flex-col pb-28 shadow-2xl relative">
          
          {/* Kartu profil dan jam hanya tampil di halaman Beranda karyawan */}
          {(currentRoute === '#karyawan' || currentRoute === '') && (
          <div className="bg-[#fff0f3] text-slate-800 p-6 rounded-b-[2.5rem] border-b border-rose-100 shadow-[0_15px_40px_-15px_rgba(0,0,0,0.05)] relative z-20">
            <div className="flex items-center justify-between mb-5 relative z-10">
              <div className="flex items-center space-x-3.5 min-w-0 flex-1">
                <div className="w-14 h-14 rounded-full bg-white p-1 border border-rose-200 shadow-sm flex items-center justify-center overflow-hidden">
                  {currentUser.foto_profil ? (
                    <img key={`${currentUser.foto_profil}_${currentUser.foto_profil_updated_at || 'initial'}`} src={getProfileImageUrl(currentUser.foto_profil, currentUser.foto_profil_updated_at || 'initial')} alt={currentUser.nama_lengkap} className="w-full h-full rounded-full object-cover" />
                  ) : (
                    <span className="font-display font-bold text-xl text-rose-500">
                      {currentUser.nama_lengkap.split(' ').map(n=>n[0]).slice(0, 2).join('')}
                    </span>
                  )}
                </div>
                <div className="min-w-0 flex-1">
                  <h1 className="font-display font-bold tracking-tight leading-snug text-slate-900 whitespace-nowrap overflow-hidden" style={{ fontSize: (String(currentUser.nama_lengkap || '').length > 24 ? '13px' : (String(currentUser.nama_lengkap || '').length > 18 ? '14px' : '18px')) }} title={currentUser.nama_lengkap}>{currentUser.nama_lengkap}</h1>
                  <div className="flex items-center gap-1.5 mt-0.5 flex-wrap">
                    <span className="px-2.5 py-0.5 bg-rose-200 text-rose-700 text-[11px] font-semibold rounded-full">
                      {currentUser.divisi}
                    </span>
                    {currentUser.jabatan && (
                      <span className="px-2.5 py-0.5 bg-indigo-100 text-indigo-700 text-[11px] font-semibold rounded-full">
                        {currentUser.jabatan}
                      </span>
                    )}
                  </div>
                </div>
              </div>

              {}
            </div>

            {/* Attendance Box */}
            <div className="bg-white/60 backdrop-blur-md rounded-[1.5rem] p-4 border border-rose-100 shadow-sm relative z-10">
              {/* Jam Waktu Operasional Center */}
              <div className="text-center mb-4">
                <p className="text-[10px] text-slate-500 uppercase tracking-widest font-bold mb-1">Waktu Operasional</p>
                <p className="text-3xl sm:text-4xl font-extrabold tracking-tight text-slate-800 font-mono-numbers drop-shadow-sm">
                  {clock.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}
                </p>
                <p className="text-[11px] font-medium text-slate-600 mt-0.5">{formatIndoDate(clock)}</p>
                {todayAttendanceMasuk && (() => {
                  const shiftNama = String(todayAttendanceMasuk.shift_nama || '').trim();
                  const shiftTerpilih = (Array.isArray(shiftList) ? shiftList : []).find(s =>
                    String(s.nama_shift || '').trim().toLocaleLowerCase('id') === shiftNama.toLocaleLowerCase('id')
                  );
                  return (
                    <div className="mt-2 rounded-xl border border-violet-100 bg-gradient-to-r from-violet-50 to-fuchsia-50 px-3 py-1.5 text-center shadow-sm">
                      <p className="text-[9px] leading-tight font-bold uppercase tracking-wider text-violet-500">Shift Operasional Anda</p>
                      {shiftTerpilih ? (
                        <p className="mt-0.5 text-[11px] leading-snug font-extrabold text-slate-800 break-words">
                          {shiftNama || 'Shift belum tercatat'} <span className="font-semibold text-slate-500">- {formatShiftTimeDisplay(shiftTerpilih.jam_masuk)} – {formatShiftTimeDisplay(shiftTerpilih.jam_pulang)} WIB</span>
                        </p>
                      ) : (
                        <p className="mt-0.5 text-[10px] leading-snug text-slate-500">{shiftNama || 'Shift belum tercatat'} - Jam operasional tidak ditemukan di Master Shift.</p>
                      )}
                    </div>
                  );
                })()}
              </div>

              {/* Tombol presensi: satu tombol nonaktif jika status Sakit/Izin/Cuti/Libur sudah tercatat hari ini */}
              {todayHasStatusAction ? (
                <button type="button" disabled aria-disabled="true" title={`Status hari ini: ${todaySpecialStatusRecord.status}`} className={`w-full py-3.5 px-4 rounded-2xl flex items-center justify-center gap-2 border cursor-not-allowed shadow-sm ${
                  todaySpecialStatusRecord.status === 'Libur' ? 'bg-sky-100 border-sky-200 text-sky-700' :
                  todaySpecialStatusRecord.status === 'Sakit' ? 'bg-amber-100 border-amber-200 text-amber-700' :
                  todaySpecialStatusRecord.status === 'Cuti' ? 'bg-purple-100 border-purple-200 text-purple-700' :
                  'bg-orange-100 border-orange-200 text-orange-700'
                }`}>
                  <i className={`fa-solid ${
                    todaySpecialStatusRecord.status === 'Libur' ? 'fa-umbrella-beach' :
                    todaySpecialStatusRecord.status === 'Sakit' ? 'fa-file-medical' :
                    todaySpecialStatusRecord.status === 'Cuti' ? 'fa-calendar-check' :
                    'fa-hand'
                  } text-base`}></i>
                  <span className="text-sm font-extrabold tracking-wide">{String(todaySpecialStatusRecord.status).toUpperCase()}</span>
                  <i className="fa-solid fa-circle-check text-sm"></i>
                </button>
              ) : (
              <div className="grid grid-cols-2 gap-3">
                <button
                  onClick={() => {
                    if (todayAttendanceMasuk) showToast('Anda sudah melakukan presensi masuk', 'info');
                    else { setWizardType('masuk'); setShowWizardModal(true); }
                  }}
                  disabled={!!todayAttendanceMasuk}
                  className={`py-2.5 px-3 rounded-2xl flex items-center justify-center gap-2 transition active:scale-[0.97] shadow-sm ${
                    todayAttendanceMasuk
                      ? 'bg-rose-100 border border-rose-200 text-rose-400 cursor-default'
                      : 'bg-white border border-rose-100 text-rose-600 hover:bg-rose-50'
                  }`}
                >
                  <i className={`fa-solid ${todayAttendanceMasuk ? 'fa-check-circle' : 'fa-arrow-right-to-bracket'} text-sm`}></i>
                  <div className="text-left leading-none">
                    <span className="block text-[10px] opacity-70 font-semibold">{todayAttendanceMasuk ? 'Tercatat' : 'Presensi'}</span>
                    <span className="block text-xs font-bold">{todayAttendanceMasuk ? todayAttendanceMasuk.jam : 'Masuk Kerja'}</span>
                  </div>
                </button>

                <button
                  onClick={() => {
                    if (!todayAttendanceMasuk) return showToast('Anda belum presensi masuk!', 'error');
                    if (todayAttendanceKeluar) return showToast('Anda sudah presensi keluar', 'info');
                    setWizardType('keluar'); setShowWizardModal(true);
                  }}
                  disabled={!todayAttendanceMasuk || !!todayAttendanceKeluar}
                  className={`py-2.5 px-3 rounded-2xl flex items-center justify-center gap-2 transition active:scale-[0.97] shadow-sm ${
                    todayAttendanceKeluar
                      ? 'bg-rose-100 border border-rose-200 text-rose-400 cursor-default'
                      : !todayAttendanceMasuk
                      ? 'bg-white/50 border border-rose-100/50 text-slate-400 cursor-not-allowed'
                      : 'bg-white border border-rose-100 text-fuchsia-600 hover:bg-rose-50'
                  }`}
                >
                  <i className={`fa-solid ${todayAttendanceKeluar ? 'fa-check-circle' : 'fa-arrow-right-from-bracket'} text-sm`}></i>
                  <div className="text-left leading-none">
                    <span className="block text-[10px] opacity-70 font-semibold">{todayAttendanceKeluar ? 'Tercatat' : 'Presensi'}</span>
                    <span className="block text-xs font-bold">{todayAttendanceKeluar ? todayAttendanceKeluar.jam : 'Pulang Kerja'}</span>
                  </div>
                </button>
              </div>
              )}
            </div>
          </div>
          )}

          {/* Conditional Sub-Pages inside Employee Mobile Portal */}
          {currentRoute === '#status-karyawan' ? (
            <StatusRekanView karyawanList={karyawanList} presensiList={presensiList} currentUser={currentUser} />
          ) : currentRoute === '#more-apps' ? (
            <div className="px-5 pt-4 pb-28 space-y-5">
              <div className="relative overflow-hidden rounded-3xl border border-violet-100 bg-gradient-to-r from-violet-50 via-white to-fuchsia-50 p-4 shadow-sm">
                <div className="pointer-events-none absolute -right-5 -top-7 h-24 w-24 rounded-full bg-fuchsia-200/40 blur-2xl"></div>
                <div className="relative flex items-start gap-3">
                  <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-2xl bg-gradient-to-br from-violet-600 to-fuchsia-500 text-white shadow-md shadow-violet-200"><i className="fa-solid fa-layer-group text-lg"></i></div>
                  <div className="min-w-0">
                    <h2 className="text-base font-extrabold leading-snug text-slate-900">App Center Suryatokoku &amp; Surya Cantik</h2>
                    <p className="mt-1 text-xs leading-relaxed text-slate-600">Akses semua aplikasi untuk mendukung aktivitas operasional toko dalam satu aplikasi.</p>
                  </div>
                </div>
              </div>
              <div className="grid grid-cols-2 gap-3">
                {(Array.isArray(moreAppsCatalog) ? moreAppsCatalog : DEFAULT_MORE_APPS).filter(app => app && typeof app === 'object' && typeof app.label === 'string' && app.label.trim()).map(app => ({ color: app.color || 'from-violet-50 to-purple-100', ink: app.ink || 'text-violet-700', icon: app.icon || 'fa-layer-group', description: app.description || 'Akses cepat untuk kebutuhan operasional.', ...app })).filter(app => Array.isArray(moreAppsAccess?.[app.label]) && moreAppsAccess[app.label].some(id => String(id) === String(currentUser?.id))).map((app) => (
                  <button key={app.label} onClick={() => { if (app.url && /^https?:\/\//i.test(String(app.url).trim())) { try { const target = new URL(String(app.url).trim(), window.location.href); window.location.assign(target.href); } catch (error) { setShowComingSoonApp(app.label); } } else { setShowComingSoonApp(app.label); } }} className={`group min-h-[142px] text-left border rounded-2xl p-4 transition active:scale-[0.98] ${app.url && /^https?:\/\//i.test(String(app.url).trim()) ? 'bg-white border-slate-100 shadow-sm hover:shadow-md hover:-translate-y-0.5' : 'bg-slate-100 border-slate-200 shadow-sm hover:bg-slate-200'}`}>
                    <div className={`w-11 h-11 rounded-2xl bg-gradient-to-br ${app.color} ${app.ink} flex items-center justify-center text-lg mb-3 group-hover:scale-105 transition`}><i className={`fa-solid ${app.icon}`}></i></div>
                    <div className="text-xs font-extrabold text-slate-800 leading-snug">{app.label}</div>
                    <div className="mt-1 text-[10px] leading-snug text-slate-500">{app.description || 'Akses cepat untuk kebutuhan operasional.'}</div>
                    <div className="mt-2 flex items-center justify-between">{!(app.url && /^https?:\/\//i.test(String(app.url).trim())) && <span className="text-[9px] font-bold uppercase tracking-wider text-slate-400">Coming soon</span>}{app.url && /^https?:\/\//i.test(String(app.url).trim()) && <span className="text-[9px] font-bold uppercase tracking-wider text-emerald-600">Buka aplikasi</span>}<i className={`fa-solid fa-arrow-up-right-from-square text-[10px] ${app.url && /^https?:\/\//i.test(String(app.url).trim()) ? 'text-slate-300' : 'text-slate-400'}`}></i></div>
                  </button>
                ))}
                {!( (Array.isArray(moreAppsCatalog) ? moreAppsCatalog : DEFAULT_MORE_APPS).filter(app => app && typeof app === 'object' && typeof app.label === 'string' && app.label.trim()).some(app => Array.isArray(moreAppsAccess?.[app.label]) && moreAppsAccess[app.label].some(id => String(id) === String(currentUser?.id))) ) && (
                  <div className="col-span-2 rounded-2xl border border-dashed border-slate-200 bg-white p-6 text-center"><div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-violet-50 text-violet-600"><i className="fa-solid fa-mobile-screen-button text-xl" aria-hidden="true"></i></div><p className="text-sm font-bold text-slate-700">Belum ada aplikasi yang ditetapkan</p><p className="mt-1 text-xs text-slate-500">Hubungi Operational Development Suryatokoku jika Anda memerlukan akses More Apps.</p></div>
                )}
              </div>
            </div>
          ) : currentRoute === '#daftar-pengumuman' ? (
            <DaftarPengumumanKaryawanView 
              pengumumanList={pengumumanList} 
              onSelectAnnouncement={setSelectedAnnouncement}
              navigateTo={navigateTo}
            />
          ) : currentRoute === '#profil' ? (
            <ProfilKaryawanView currentUser={currentUser} setCurrentUser={setCurrentUser} showToast={showToast} handleLogout={handleLogout} />
          ) : currentRoute === '#riwayat-aktivitas' ? (
            <RiwayatAktivitasView 
              currentUser={currentUser} 
              presensiList={presensiList} 
              klaimList={klaimList}
              lemburList={lemburList}
              payrollCycle={payrollCycle}
              filterStatus={filterStatusRiwayat}
              setFilterStatus={setFilterStatusRiwayat}
            />
          ) : currentRoute === '#ganti-password' ? (
            <GantiPasswordView currentUser={currentUser} setCurrentUser={setCurrentUser} setKaryawanList={setKaryawanList} showToast={showToast} navigateTo={navigateTo} />
          ) : (
            /* Home Beranda View */
            <div className="px-5 pt-6 space-y-6">
              
              {/* Kartu Pengumuman Terakhir Ditampilkan Penuh */}
              {pengumumanList.length > 0 && (() => {
                const latest = pengumumanList[0];
                return (
                  <div className="bg-white/95 backdrop-blur-xl rounded-[1.75rem] p-5 border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] space-y-3.5 relative overflow-hidden">
                    {/* Header: Label, Indikator & Tombol Lihat Semua */}
                    <div className="flex items-center justify-between">
                      <div className="flex items-center space-x-2">
                        <span className="w-2 h-2 rounded-full bg-violet-600 animate-pulse"></span>
                        <span className="text-[10px] font-extrabold text-violet-600 uppercase tracking-widest">Pengumuman Terbaru</span>
                      </div>
                      <button 
                        onClick={() => navigateTo('#daftar-pengumuman')}
                        className="text-[10px] font-bold text-slate-400 hover:text-violet-600 flex items-center gap-1 transition"
                      >
                        <span>Lihat Semua</span>
                        <i className="fa-solid fa-arrow-right text-[9px]"></i>
                      </button>
                    </div>

                    {/* Judul & Tanggal */}
                    <div className="space-y-1.5">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="text-[10px] font-mono-numbers text-slate-400 font-medium">
                          {new Date(latest.created_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}
                        </span>
                      </div>
                      <h4 className="font-display font-bold text-sm text-slate-800 leading-snug">
                        {latest.judul}
                      </h4>
                    </div>

                    {/* Isi Pengumuman Lengkap */}
                    <div className="bg-slate-50/80 p-3.5 rounded-2xl border border-slate-100/80">
                      <p className="text-xs text-slate-600 font-medium leading-relaxed whitespace-pre-line">
                        {latest.isi}
                      </p>
                    </div>

                    {/* Footer Penulis */}
                    <div className="flex items-center justify-between text-[10px] text-slate-400 pt-0.5 font-semibold">
                      <span className="flex items-center gap-1.5 text-slate-600">
                        <i className="fa-regular fa-building text-violet-500"></i>
                        {latest.author || 'HRD Suryatokoku'}
                      </span>
                      <span className="text-[9px] text-slate-400">
                        Resmi Perusahaan
                      </span>
                    </div>
                  </div>
                );
              })()}

              {/* Kartu Sisa Libur & Sisa Cuti (Minimalist Colorful Style) */}
              <div className="grid grid-cols-2 gap-4">
                <div className="bg-white/90 backdrop-blur-xl p-3.5 rounded-[1.5rem] border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] flex items-center justify-between transition">
                  <div className="flex items-center space-x-3">
                    <div className="w-10 h-10 rounded-full bg-sky-100 text-sky-500 flex items-center justify-center shadow-inner">
                      <i className="fa-solid fa-umbrella-beach text-sm"></i>
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-500 font-bold block leading-none mb-1">Sisa Libur</span>
                      <span className="font-display text-xl font-extrabold text-slate-800 leading-none block">
                        {sisaLiburAktif} <span className="text-[10px] font-semibold text-slate-400">Hari</span>
                      </span>
                    </div>
                  </div>
                </div>

                <div className="bg-white/90 backdrop-blur-xl p-3.5 rounded-[1.5rem] border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] flex items-center justify-between transition">
                  <div className="flex items-center space-x-3">
                    <div className="w-10 h-10 rounded-full bg-purple-100 text-purple-500 flex items-center justify-center shadow-inner">
                      <i className="fa-solid fa-plane-departure text-sm"></i>
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-500 font-bold block leading-none mb-1">Sisa Cuti</span>
                      <span className="font-display text-xl font-extrabold text-slate-800 leading-none block">
                        {currentUser.sisa_cuti} <span className="text-[10px] font-semibold text-slate-400">Hari</span>
                      </span>
                    </div>
                  </div>
                </div>
              </div>

              {/* Secondary 4-Grid Menu (Pill shapes & Colorful Icons) */}
              <div>
                <h3 className="text-[11px] font-bold text-slate-400 uppercase tracking-widest mb-3 px-2">Layanan Cepat</h3>
                <div className="grid grid-cols-4 gap-3">
                  <button onClick={() => handleActionClick('sakit')} disabled={todayHasStatusAction || !!todayAttendanceMasuk || !!todayAttendanceKeluar} title={(todayAttendanceMasuk || todayAttendanceKeluar) ? 'Presensi masuk/pulang sudah tercatat hari ini' : todayHasStatusAction ? `Hari ini sudah tercatat ${todaySpecialStatusRecord.status}` : 'Ajukan sakit'} className="flex flex-col items-center justify-center text-center active:scale-95 transition group disabled:opacity-50 disabled:cursor-not-allowed">
                    <div className="w-14 h-14 rounded-[1.25rem] bg-white border border-slate-100 shadow-[0_8px_20px_rgb(0,0,0,0.03)] text-amber-500 flex items-center justify-center mb-2 group-hover:bg-amber-50 transition">
                      <i className="fa-solid fa-file-medical text-xl"></i>
                    </div>
                    <span className="text-[11px] font-semibold text-slate-600">Sakit</span>
                  </button>

                  {sisaLiburAktif > 0 ? (
                    <button onClick={() => handleActionClick('libur')} disabled={todayHasStatusAction || !!todayAttendanceMasuk || !!todayAttendanceKeluar} title={todayHasStatusAction ? `Hari ini sudah tercatat ${todaySpecialStatusRecord.status}` : 'Catat libur'} className="flex flex-col items-center justify-center text-center active:scale-95 transition group disabled:opacity-50 disabled:cursor-not-allowed">
                      <div className="w-14 h-14 rounded-[1.25rem] bg-white border border-slate-100 shadow-[0_8px_20px_rgb(0,0,0,0.03)] text-sky-500 flex items-center justify-center mb-2 group-hover:bg-sky-50 transition">
                        <i className="fa-solid fa-calendar-day text-xl"></i>
                      </div>
                      <span className="text-[11px] font-semibold text-slate-600">Libur</span>
                    </button>
                  ) : (
                    <button onClick={() => handleActionClick('izin')} disabled={todayHasStatusAction || !!todayAttendanceMasuk || !!todayAttendanceKeluar} title={(todayAttendanceMasuk || todayAttendanceKeluar) ? 'Presensi masuk/pulang sudah tercatat hari ini' : todayHasStatusAction ? `Hari ini sudah tercatat ${todaySpecialStatusRecord.status}` : 'Ajukan izin'} className="flex flex-col items-center justify-center text-center active:scale-95 transition group disabled:opacity-50 disabled:cursor-not-allowed">
                      <div className="w-14 h-14 rounded-[1.25rem] bg-white border border-slate-100 shadow-[0_8px_20px_rgb(0,0,0,0.03)] text-orange-500 flex items-center justify-center mb-2 group-hover:bg-orange-50 transition">
                        <i className="fa-solid fa-hand text-xl"></i>
                      </div>
                      <span className="text-[11px] font-semibold text-slate-600">Izin</span>
                    </button>
                  )}

                  <button onClick={() => handleActionClick('cuti')} disabled={todayHasStatusAction || !!todayAttendanceMasuk || !!todayAttendanceKeluar} title={(todayAttendanceMasuk || todayAttendanceKeluar) ? 'Presensi masuk/pulang sudah tercatat hari ini' : todayHasStatusAction ? `Hari ini sudah tercatat ${todaySpecialStatusRecord.status}` : 'Ajukan cuti'} className="flex flex-col items-center justify-center text-center active:scale-95 transition group disabled:opacity-50 disabled:cursor-not-allowed">
                    <div className="w-14 h-14 rounded-[1.25rem] bg-white border border-slate-100 shadow-[0_8px_20px_rgb(0,0,0,0.03)] text-purple-500 flex items-center justify-center mb-2 group-hover:bg-purple-50 transition">
                      <i className="fa-solid fa-calendar-check text-xl"></i>
                    </div>
                    <span className="text-[11px] font-semibold text-slate-600">Cuti</span>
                  </button>

                  <button onClick={() => handleActionClick('lembur')} disabled={todayHasStatusAction} title={todayHasStatusAction ? `Lembur tidak tersedia karena status hari ini ${todaySpecialStatusRecord.status}` : 'Ajukan lembur'} className="flex flex-col items-center justify-center text-center active:scale-95 transition group disabled:opacity-50 disabled:cursor-not-allowed">
                    <div className="w-14 h-14 rounded-[1.25rem] bg-white border border-slate-100 shadow-[0_8px_20px_rgb(0,0,0,0.03)] text-indigo-500 flex items-center justify-center mb-2 group-hover:bg-indigo-50 transition">
                      <i className="fa-solid fa-business-time text-xl"></i>
                    </div>
                    <span className="text-[11px] font-semibold text-slate-600">Lembur</span>
                  </button>
                </div>
              </div>

              {/* Payroll Cycle Summary with In-Place Activity List */}
              <div className="bg-white/90 backdrop-blur-xl rounded-[1.5rem] p-5 border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)]">
                {/* Header Kartu */}
                <div className="flex justify-between items-center mb-4">
                  <div>
                    <h3 className="text-xs font-extrabold text-slate-800 uppercase tracking-widest font-display">Periode Cut-Off</h3>
                    <p className="text-[11px] text-slate-500 font-medium mt-0.5">{payrollCycle.label}</p>
                  </div>
                  <button 
                    onClick={() => navigateTo('#riwayat-aktivitas')} 
                    className="w-8 h-8 rounded-full bg-violet-50 text-violet-600 hover:bg-violet-100 flex items-center justify-center transition active:scale-95"
                    title="Buka Riwayat Lengkap"
                  >
                    <i className="fa-solid fa-chevron-right text-[10px]"></i>
                  </button>
                </div>

                {/* 5 Tombol Filter Cepat In-Place */}
                <div className="grid grid-cols-5 gap-2 text-center mb-4">
                  <div 
                    onClick={() => setCutoffFilter(prev => prev === 'Masuk' ? 'Semua' : 'Masuk')} 
                    className="cursor-pointer active:scale-95 transition"
                  >
                    <div className={`w-full py-1.5 rounded-xl border flex flex-col items-center justify-center mb-1 transition shadow-sm ${
                      cutoffFilter === 'Masuk' 
                        ? 'bg-emerald-600 border-emerald-600 text-white shadow-emerald-200' 
                        : 'bg-emerald-50 border-emerald-100 text-emerald-600 hover:bg-emerald-100/70'
                    }`}>
                      <span className="font-mono-numbers text-base font-bold">{employeeStats.masuk}</span>
                    </div>
                    <span className={`text-[9px] font-bold uppercase ${cutoffFilter === 'Masuk' ? 'text-emerald-700' : 'text-slate-500'}`}>Masuk</span>
                  </div>

                  <div 
                    onClick={() => setCutoffFilter(prev => prev === 'Libur' ? 'Semua' : 'Libur')} 
                    className="cursor-pointer active:scale-95 transition"
                  >
                    <div className={`w-full py-1.5 rounded-xl border flex flex-col items-center justify-center mb-1 transition shadow-sm ${
                      cutoffFilter === 'Libur' 
                        ? 'bg-sky-600 border-sky-600 text-white shadow-sky-200' 
                        : 'bg-sky-50 border-sky-100 text-sky-600 hover:bg-sky-100/70'
                    }`}>
                      <span className="font-mono-numbers text-base font-bold">{employeeStats.libur}</span>
                    </div>
                    <span className={`text-[9px] font-bold uppercase ${cutoffFilter === 'Libur' ? 'text-sky-700' : 'text-slate-500'}`}>Libur</span>
                  </div>

                  <div 
                    onClick={() => setCutoffFilter(prev => prev === 'Izin' ? 'Semua' : 'Izin')} 
                    className="cursor-pointer active:scale-95 transition"
                  >
                    <div className={`w-full py-1.5 rounded-xl border flex flex-col items-center justify-center mb-1 transition shadow-sm ${
                      cutoffFilter === 'Izin' 
                        ? 'bg-orange-600 border-orange-600 text-white shadow-orange-200' 
                        : 'bg-orange-50 border-orange-100 text-orange-600 hover:bg-orange-100/70'
                    }`}>
                      <span className="font-mono-numbers text-base font-bold">{employeeStats.izin}</span>
                    </div>
                    <span className={`text-[9px] font-bold uppercase ${cutoffFilter === 'Izin' ? 'text-orange-700' : 'text-slate-500'}`}>Izin</span>
                  </div>

                  <div 
                    onClick={() => setCutoffFilter(prev => prev === 'Sakit' ? 'Semua' : 'Sakit')} 
                    className="cursor-pointer active:scale-95 transition"
                  >
                    <div className={`w-full py-1.5 rounded-xl border flex flex-col items-center justify-center mb-1 transition shadow-sm ${
                      cutoffFilter === 'Sakit' 
                        ? 'bg-amber-600 border-amber-600 text-white shadow-amber-200' 
                        : 'bg-amber-50 border-amber-100 text-amber-600 hover:bg-amber-100/70'
                    }`}>
                      <span className="font-mono-numbers text-base font-bold">{employeeStats.sakit}</span>
                    </div>
                    <span className={`text-[9px] font-bold uppercase ${cutoffFilter === 'Sakit' ? 'text-amber-700' : 'text-slate-500'}`}>Sakit</span>
                  </div>

                  <div 
                    onClick={() => setCutoffFilter(prev => prev === 'Cuti' ? 'Semua' : 'Cuti')} 
                    className="cursor-pointer active:scale-95 transition"
                  >
                    <div className={`w-full py-1.5 rounded-xl border flex flex-col items-center justify-center mb-1 transition shadow-sm ${
                      cutoffFilter === 'Cuti' 
                        ? 'bg-purple-600 border-purple-600 text-white shadow-purple-200' 
                        : 'bg-purple-50 border-purple-100 text-purple-600 hover:bg-purple-100/70'
                    }`}>
                      <span className="font-mono-numbers text-base font-bold">{employeeStats.cuti}</span>
                    </div>
                    <span className={`text-[9px] font-bold uppercase ${cutoffFilter === 'Cuti' ? 'text-purple-700' : 'text-slate-500'}`}>Cuti</span>
                  </div>
                </div>

                {/* Sub-Header & Status Filter Aktif */}
                <div className="flex items-center justify-between pt-3 border-t border-slate-100 mb-2.5">
                  <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">
                    Aktivitas Periode ({cutoffFilter === 'Semua' ? 'Semua' : cutoffFilter})
                  </span>
                  {cutoffFilter !== 'Semua' && (
                    <button 
                      onClick={() => setCutoffFilter('Semua')}
                      className="text-[10px] font-bold text-violet-600 hover:text-violet-800 transition"
                    >
                      Reset Filter
                    </button>
                  )}
                </div>

                {/* Daftar Aktivitas Presensi Periode Aktif */}
                <div className="space-y-2.5 max-h-72 overflow-y-auto pr-1">
                  {filteredPeriodActivities.length > 0 ? (
                    filteredPeriodActivities.map(act => (
                      <div 
                        key={act.id} 
                        className="bg-slate-50/90 hover:bg-slate-50 p-3 rounded-2xl border border-slate-100/80 transition flex flex-col space-y-2"
                      >
                        {/* Baris Atas: Tanggal & Badge Status Absen */}
                        <div className="flex items-center justify-between">
                          <span className="text-[11px] font-bold text-slate-800 font-mono-numbers flex items-center gap-1.5">
                            <i className="fa-regular fa-calendar text-[10px] text-slate-400"></i>
                            {formatIndoDate(act.tanggal)}
                          </span>
                          <span className={`px-2.5 py-0.5 rounded-full text-[9px] font-bold ${
                            act.status === 'Hadir' ? 'bg-emerald-100 text-emerald-700 border border-emerald-200' :
                            act.status === 'Terlambat' ? 'bg-amber-100 text-amber-700 border border-amber-200' :
                            act.status === 'Libur' ? 'bg-sky-100 text-sky-700 border border-sky-200' :
                            act.status === 'Izin' ? 'bg-orange-100 text-orange-700 border border-orange-200' :
                            act.status === 'Sakit' ? 'bg-amber-100 text-amber-700 border border-amber-200' :
                            act.status === 'Cuti' ? 'bg-purple-100 text-purple-700 border border-purple-200' :
                            act.status === 'Pending' ? 'bg-amber-100 text-amber-800 border border-amber-200' :
                            'bg-slate-200 text-slate-700'
                          }`}>
                            {act.status}
                          </span>
                        </div>

                        {/* Baris Bawah: Shift Kerja, Jam Masuk, dan Jam Pulang */}
                        <div className="flex items-center justify-between pt-1 border-t border-slate-200/50 text-[10px]">
                          <div className="text-slate-500 font-semibold truncate pr-2 flex items-center gap-1">
                            <i className="fa-regular fa-clock text-[9px] text-violet-500"></i>
                            <span className="truncate max-w-[130px]">{act.shift}</span>
                          </div>
                          {['Libur', 'Sakit', 'Cuti', 'Izin'].includes(String(act.status || '').trim()) ? (
                            <span className="text-slate-600 bg-slate-100 px-2 py-1 rounded-md border border-slate-200 font-semibold">
                              Tidak ada jam presensi
                            </span>
                          ) : (
                            <div className="flex items-center space-x-2 shrink-0 font-mono-numbers font-bold">
                              <span className="text-emerald-600 bg-emerald-50 px-1.5 py-0.5 rounded-md border border-emerald-100" title="Jam Masuk">
                                <i className="fa-solid fa-arrow-right-to-bracket text-[8px] mr-1"></i>{act.jamMasuk}
                              </span>
                              <span className="text-rose-600 bg-rose-50 px-1.5 py-0.5 rounded-md border border-rose-100" title="Jam Pulang">
                                <i className="fa-solid fa-arrow-right-from-bracket text-[8px] mr-1"></i>{act.jamPulang}
                              </span>
                            </div>
                          )}
                        </div>
                        {act.status === 'Pending' && <div className="text-[10px] leading-relaxed text-amber-800 bg-amber-50 border border-amber-100 rounded-xl px-3 py-2"><i className="fa-regular fa-clock mr-1.5"></i>Menunggu persetujuan HR</div>}
                        {act.keterangan && /disetujui HR/i.test(act.keterangan) && <div className="text-[10px] leading-relaxed text-emerald-700 bg-emerald-50 border border-emerald-100 rounded-xl px-3 py-2"><i className="fa-solid fa-circle-check mr-1.5"></i>{act.keterangan}</div>}
                      </div>
                    ))
                  ) : (
                    <div className="text-center py-6 text-slate-400 bg-slate-50/50 rounded-2xl border border-dashed border-slate-200">
                      <i className="fa-regular fa-calendar-xmark text-xl mb-1 text-slate-300"></i>
                      <p className="text-[11px] font-medium">Tidak ada aktivitas {cutoffFilter !== 'Semua' ? cutoffFilter : ''} pada periode ini.</p>
                    </div>
                  )}
                </div>
              </div>

            </div>
          )}

          {/* Floating Pill Bottom Navigation (Always Floating in Viewport) */}
          <div className="fixed bottom-4 left-1/2 -translate-x-1/2 w-[calc(100%-1.25rem)] max-w-md z-50 pointer-events-auto">
            <nav className="bg-white/95 backdrop-blur-2xl border border-slate-200/80 px-2 py-2.5 rounded-full flex justify-around items-center gap-1 shadow-[0_12px_35px_rgba(0,0,0,0.12)]">
              <button onClick={() => navigateTo('#karyawan')} className={`flex-1 min-w-0 flex flex-col items-center space-y-0.5 transition ${currentRoute === '#karyawan' || currentRoute === '' ? 'text-violet-600' : 'text-slate-400 hover:text-slate-600'}`}>
                <i className="fa-solid fa-house text-lg"></i>
                <span className="text-[9px] font-bold">Beranda</span>
              </button>
              <button onClick={() => navigateTo('#status-karyawan')} className={`flex-1 min-w-0 flex flex-col items-center space-y-0.5 transition ${currentRoute === '#status-karyawan' ? 'text-violet-600' : 'text-slate-400 hover:text-slate-600'}`}>
                <i className="fa-solid fa-users text-lg"></i>
                <span className="text-[9px] font-bold">Rekan</span>
              </button>
              <button onClick={() => navigateTo('#more-apps')} className={`flex-1 min-w-0 flex flex-col items-center space-y-0.5 transition ${currentRoute === '#more-apps' ? 'text-violet-600' : 'text-slate-400 hover:text-slate-600'}`}>
                <i className={`fa-solid fa-layer-group text-lg ${currentRoute === '#more-apps' ? 'text-violet-600' : ''}`}></i>
                <span className="text-[9px] font-bold whitespace-nowrap">More Apps</span>
              </button>
              <button onClick={() => navigateTo('#riwayat-aktivitas')} className={`flex-1 min-w-0 flex flex-col items-center space-y-0.5 transition ${currentRoute === '#riwayat-aktivitas' ? 'text-violet-600' : 'text-slate-400 hover:text-slate-600'}`}>
                <i className="fa-solid fa-clock-rotate-left text-lg"></i>
                <span className="text-[9px] font-bold">Riwayat</span>
              </button>
              <button onClick={() => navigateTo('#profil')} className={`flex-1 min-w-0 flex flex-col items-center space-y-0.5 transition ${currentRoute === '#profil' ? 'text-violet-600' : 'text-slate-400 hover:text-slate-600'}`}>
                {currentUser?.foto_profil ? (
                  <img
                    key={`${currentUser.foto_profil}_${currentUser.foto_profil_updated_at || 'nav'}`}
                    src={getProfileImageUrl(currentUser.foto_profil, currentUser.foto_profil_updated_at || 'nav')}
                    alt="Foto profil"
                    className={`w-6 h-6 rounded-full object-cover border-2 ${currentRoute === '#profil' ? 'border-violet-600' : 'border-slate-300'}`}
                    onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }}
                  />
                ) : null}
                <i className={`fa-solid fa-circle-user text-lg ${currentUser?.foto_profil ? 'hidden' : ''}`}></i>
                <span className="text-[9px] font-bold">Profil</span>
              </button>
            </nav>
          </div>

          {/* Modal Coming Soon More Apps */}
          {showComingSoonApp && (
            <div className="fixed inset-0 z-[120] flex items-center justify-center p-5 bg-slate-950/45 backdrop-blur-md" onClick={() => setShowComingSoonApp('')}>
              <div className="relative w-full max-w-xs overflow-hidden rounded-[2rem] bg-white shadow-2xl border border-white/70 animate-in fade-in zoom-in duration-200" onClick={(e) => e.stopPropagation()}>
                <div className="h-2 bg-gradient-to-r from-violet-600 via-fuchsia-500 to-sky-400"></div>
                <div className="p-6 text-center">
                  <div className="relative mx-auto mb-4 w-20 h-20 rounded-[1.7rem] bg-gradient-to-br from-violet-100 via-fuchsia-50 to-sky-100 flex items-center justify-center text-violet-700 shadow-inner">
                    <i className="fa-solid fa-rocket text-3xl"></i>
                    <span className="absolute -right-1 -top-1 w-7 h-7 rounded-full bg-amber-400 text-white flex items-center justify-center border-4 border-white"><i className="fa-solid fa-sparkles text-[10px]"></i></span>
                  </div>
                  <p className="text-[10px] font-extrabold uppercase tracking-[0.2em] text-violet-600 mb-2">SuryaKuCan • More Apps</p>
                  <h3 className="text-xl font-extrabold text-slate-900">Coming Soon!</h3>
                  <p className="mt-2 text-sm leading-relaxed text-slate-500"><span className="font-bold text-slate-700">{showComingSoonApp}</span> sedang kami siapkan agar bisa digunakan. Nantikan ya!</p>
                  <button onClick={() => setShowComingSoonApp('')} className="mt-6 w-full rounded-2xl bg-gradient-to-r from-violet-600 to-fuchsia-600 py-3.5 text-sm font-extrabold text-white shadow-lg shadow-violet-200 transition hover:brightness-105 active:scale-[0.98]">Oke, Mengerti</button>
                </div>
              </div>
            </div>
          )}

          {/* Modal Wizard Presensi Masuk / Keluar */}
          {showWizardModal && (
            <PresensiWizardModal
              type={wizardType}
              currentUser={currentUser}
              lokasiList={lokasiList}
              divisiList={divisiList}
              shiftList={shiftList}
              toleransiMenit={toleransiMenit}
              onClose={() => setShowWizardModal(false)}
              onSuccess={(newRecord) => {
                setPresensiList(prev => [newRecord, ...prev.filter(p => String(p.id) !== String(newRecord.id))]);
                setShowWizardModal(false);
                showToast(`Presensi ${wizardType.toUpperCase()} berhasil dikirim!`);
              }}
            />
          )}

          {/* Konfirmasi dua langkah untuk mencatat Libur */}
          {showLiburConfirmModal && (
            <div className="fixed inset-0 bg-slate-950/50 backdrop-blur-md z-[70] flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-sm w-full p-6 shadow-2xl border border-slate-100">
                <div className="w-16 h-16 rounded-2xl bg-sky-50 text-sky-600 flex items-center justify-center mx-auto mb-4"><i className="fa-solid fa-calendar-check text-2xl"></i></div>
                <h3 className="font-display font-bold text-lg text-slate-900 text-center">Konfirmasi Libur</h3>
                <p className="text-xs text-slate-500 text-center leading-relaxed mt-2">Anda akan menggunakan 1 hari dari jatah libur. Setelah disimpan, status hari ini tidak dapat diubah melalui aplikasi.</p>
                <div className="my-4 rounded-2xl bg-sky-50 border border-sky-100 p-3 text-center"><p className="text-[10px] uppercase tracking-wide text-sky-700 font-bold">Sisa jatah libur</p><p className="text-2xl font-extrabold text-sky-800">{sisaLiburAktif} hari</p></div>
                <div className="grid grid-cols-2 gap-3">
                  <button type="button" onClick={() => setShowLiburConfirmModal(false)} className="py-3 rounded-full border border-slate-200 text-slate-600 text-xs font-bold hover:bg-slate-50">Batal</button>
                  <button type="button" onClick={async () => { setShowLiburConfirmModal(false); await handleMarkLibur(); }} className="py-3 rounded-full bg-sky-600 hover:bg-sky-700 text-white text-xs font-bold shadow-lg shadow-sky-200">Ya, Konfirmasi Libur</button>
                </div>
              </div>
            </div>
          )}

          {/* Modal Cuti Terkunci */}
          {showCutiLockModal && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-xs w-full p-7 text-center shadow-2xl border border-slate-100">
                <div className="w-16 h-16 bg-purple-50 text-purple-500 rounded-full flex items-center justify-center mx-auto mb-4">
                  <i className="fa-solid fa-lock text-2xl"></i>
                </div>
                <h3 className="font-display font-bold text-lg text-slate-800">Akses Cuti Terkunci</h3>
                <p className="text-xs text-slate-500 mt-2 leading-relaxed">
                  Divisi <span className="font-bold text-purple-600">{currentUser.divisi}</span> belum memiliki izin pengajuan cuti mandiri secara sistem.
                </p>
                <div className="bg-slate-50 rounded-2xl p-3.5 my-5 border border-slate-100 text-left">
                  <p className="text-[11px] text-slate-600 font-medium">
                    <i className="fa-solid fa-lightbulb text-amber-500 mr-2"></i>
                    Untuk kebutuhan izin darurat, hubungi HRD atau manajer Anda.
                  </p>
                </div>
                <button onClick={() => setShowCutiLockModal(false)} className="w-full py-3.5 bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs rounded-full transition shadow-md">
                  Dimengerti
                </button>
              </div>
            </div>
          )}

          {/* Smart Check Modal */}
          {showYesterdayCheck && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-sm w-full p-6 shadow-2xl border border-slate-100">
                <div className="flex items-center space-x-3 mb-4">
                  <div className="w-12 h-12 rounded-full bg-amber-50 text-amber-500 flex items-center justify-center">
                    <i className="fa-solid fa-triangle-exclamation text-xl"></i>
                  </div>
                  <div>
                    <h3 className="font-display font-bold text-base text-slate-800 leading-tight">Konfirmasi<br/>Presensi Kemarin</h3>
                  </div>
                </div>

                <p className="text-xs text-slate-500 mb-5 leading-relaxed">
                  Catatan presensi Anda kemarin belum tersedia. Mohon konfirmasi status kehadiran:
                </p>

                <div className="grid grid-cols-2 gap-3 mb-5">
                  <button onClick={() => handleYesterdayConfirm('Libur', 'Libur kemarin')} className="py-3 px-3 bg-white hover:bg-slate-50 text-slate-700 font-bold text-xs rounded-2xl border border-slate-200 shadow-sm flex items-center justify-center gap-2">
                    <i className="fa-solid fa-umbrella-beach text-sky-500"></i> Libur
                  </button>
                  <button onClick={() => handleYesterdayConfirm('Sakit', 'Sakit kemarin')} className="py-3 px-3 bg-white hover:bg-slate-50 text-slate-700 font-bold text-xs rounded-2xl border border-slate-200 shadow-sm flex items-center justify-center gap-2">
                    <i className="fa-solid fa-bed text-amber-500"></i> Sakit
                  </button>
                  <button onClick={() => handleYesterdayConfirm('Izin', 'Izin kemarin')} className="py-3 px-3 bg-white hover:bg-slate-50 text-slate-700 font-bold text-xs rounded-2xl border border-slate-200 shadow-sm flex items-center justify-center gap-2">
                    <i className="fa-solid fa-hand text-orange-500"></i> Izin
                  </button>
                  <button onClick={() => handleYesterdayConfirm('Cuti', 'Cuti kemarin')} className="py-3 px-3 bg-white hover:bg-slate-50 text-slate-700 font-bold text-xs rounded-2xl border border-slate-200 shadow-sm flex items-center justify-center gap-2">
                    <i className="fa-solid fa-calendar-check text-purple-500"></i> Cuti
                  </button>
                </div>

                <div className="mb-3 rounded-2xl border border-violet-100 bg-violet-50/70 p-3 text-left">
                  <label htmlFor="yesterdayClaimShift" className="block text-[11px] font-extrabold text-violet-800 mb-2">Pilih shift yang dijalani kemarin</label>
                  <select id="yesterdayClaimShift" value={selectedClaimShift} onChange={e => setSelectedClaimShift(e.target.value)} className="w-full rounded-xl border border-violet-200 bg-white px-3 py-3 text-xs font-bold text-slate-700 outline-none focus:ring-2 focus:ring-violet-300">
                    <option value="">-- Pilih Shift --</option>
                    {allowedClaimShifts.map(shift => <option key={shift.id || shift.nama_shift} value={shift.nama_shift}>{shift.nama_shift} ({formatShiftTimeDisplay(shift.jam_masuk)} - {formatShiftTimeDisplay(shift.jam_pulang)})</option>)}
                  </select>
                  {allowedClaimShifts.length === 0 && <p className="text-[10px] text-rose-600 mt-2">Belum ada shift yang ditetapkan untuk divisi Anda. Hubungi HRD.</p>}
                </div>
                <button onClick={handleYesterdayClaimHR} disabled={!selectedClaimShift || allowedClaimShifts.length === 0} className="w-full py-3.5 bg-slate-900 hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed text-white font-bold text-xs rounded-full shadow-md mb-2">
                  Kemarin Masuk (Klaim ke HR)
                </button>
                <button onClick={() => setShowYesterdayCheck(false)} className="w-full py-2.5 text-slate-400 hover:text-slate-600 text-[11px] font-bold">
                  Ingatkan Nanti
                </button>
              </div>
            </div>
          )}

          {selectedAnnouncement && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-sm w-full p-6 shadow-2xl border border-slate-100">
                <div className="flex justify-between items-center mb-4">
                  <span className="px-3 py-1 bg-violet-50 text-violet-600 text-[10px] font-bold rounded-full">
                    INFO HRD
                  </span>
                  <button onClick={() => setSelectedAnnouncement(null)} className="w-8 h-8 rounded-full bg-slate-50 text-slate-400 flex items-center justify-center">
                    <i className="fa-solid fa-xmark text-sm"></i>
                  </button>
                </div>
                <h3 className="font-display font-bold text-lg text-slate-900 mb-2">{selectedAnnouncement.judul}</h3>
                <p className="text-xs text-slate-600 leading-relaxed bg-slate-50 p-4 rounded-2xl border border-slate-100 mb-4 whitespace-pre-line">
                  {selectedAnnouncement.isi}
                </p>
                <div className="flex justify-between items-center text-[10px] font-bold text-slate-400">
                  <span>{selectedAnnouncement.author}</span>
                  <span>{new Date(selectedAnnouncement.created_at).toLocaleDateString('id-ID')}</span>
                </div>
              </div>
            </div>
          )}

          {showSimpleFormModal && (
            <SimpleActionModal
              type={showSimpleFormModal}
              currentUser={currentUser}
              setCurrentUser={setCurrentUser}
              onClose={() => setShowSimpleFormModal(null)}
              onSubmitAttendance={(rec) => {
                setPresensiList([rec, ...presensiList]);
                setShowSimpleFormModal(null);
                showToast(`Pengajuan ${showSimpleFormModal.toUpperCase()} berhasil dikirim!`);
              }}
              onSubmitLembur={(lmb) => {
                setLemburList([lmb, ...lemburList]);
                setShowSimpleFormModal(null);
                showToast('Pengajuan lembur dikirim!');
              }}
            />
          )}

        </div>
      );
    }

    function PresensiWizardModal({ type, currentUser, lokasiList, divisiList, shiftList, toleransiMenit, onClose, onSuccess }) {
      const [step, setStep] = useState(1);
      // Karyawan lama tanpa pengaturan lokasi tetap dapat memilih semua lokasi yang tersedia.
      const availableLokasiList = useMemo(() => {
        const all = Array.isArray(lokasiList) ? lokasiList : [];
        const hasAssignment = Array.isArray(currentUser?.lokasi_absen);
        const assigned = hasAssignment ? currentUser.lokasi_absen.map(String) : [];
        if (assigned.includes('*')) {
          const selectedIds = assigned.filter(id => id !== '*');
          // Unlock mengabaikan radius, bukan pilihan lokasi. Jika data lama hanya berisi '*', tampilkan semua lokasi sebagai kompatibilitas.
          return selectedIds.length ? all.filter(l => selectedIds.includes(String(l.id))) : all;
        }
        return hasAssignment ? all.filter(l => assigned.includes(String(l.id))) : all;
      }, [lokasiList, currentUser?.lokasi_absen]);
      const locationUnlocked = Array.isArray(currentUser?.lokasi_absen) && currentUser.lokasi_absen.map(String).includes('*');
      const [selectedLokasi, setSelectedLokasi] = useState(() => availableLokasiList[0] || null);

      // Jika daftar lokasi berubah, pastikan lokasi terpilih tetap termasuk lokasi yang diizinkan.
      useEffect(() => {
        if ((!selectedLokasi || !availableLokasiList.some(l => String(l.id) === String(selectedLokasi.id))) && availableLokasiList.length > 0) {
          setSelectedLokasi(availableLokasiList[0]);
        } else if (!availableLokasiList.length) setSelectedLokasi(null);
      }, [availableLokasiList, selectedLokasi]);
      const [userCoords, setUserCoords] = useState(null);
      const [distance, setDistance] = useState(null);
      const [inRadius, setInRadius] = useState(false);
      const [gpsLoading, setGpsLoading] = useState(true);
      const [gpsAccuracy, setGpsAccuracy] = useState(null);

      // Batasi pilihan shift berdasarkan konfigurasi Master Divisi untuk divisi karyawan.
      // Jika divisi belum ditemukan atau belum memiliki shift, jangan tampilkan semua shift
      // karena itu dapat membuat karyawan memilih shift yang tidak diizinkan.
      const currentDivision = (Array.isArray(divisiList) ? divisiList : []).find(d =>
        String(d.nama_divisi || '').trim().toLocaleLowerCase('id') ===
        String(currentUser?.divisi || '').trim().toLocaleLowerCase('id')
      );
      const allowedShiftNames = useMemo(() => {
        const configured = currentDivision?.shifts;
        let names = [];
        if (Array.isArray(configured)) {
          names = configured.map(item => typeof item === 'string' ? item : item?.nama_shift).filter(Boolean);
        } else if (typeof configured === 'string' && configured.trim()) {
          try {
            const parsed = JSON.parse(configured);
            names = Array.isArray(parsed) ? parsed.map(item => typeof item === 'string' ? item : item?.nama_shift).filter(Boolean) : [];
          } catch (_) { names = configured.split(',').map(x => x.trim()).filter(Boolean); }
        }
        return new Set(names.map(name => String(name).trim().toLocaleLowerCase('id')));
      }, [currentDivision]);
      const sortedShiftList = useMemo(() => {
        const all = Array.isArray(shiftList) ? shiftList : [];
        const allowed = all.filter(shift => allowedShiftNames.has(String(shift.nama_shift || '').trim().toLocaleLowerCase('id')));
        return allowed.sort((a, b) => {
          const toMinutes = (value) => {
            const match = String(value || '').match(/^(\d{1,2}):(\d{2})/);
            return match ? Number(match[1]) * 60 + Number(match[2]) : Number.MAX_SAFE_INTEGER;
          };
          return toMinutes(a.jam_masuk) - toMinutes(b.jam_masuk) ||
            String(a.nama_shift || '').localeCompare(String(b.nama_shift || ''), 'id');
        });
      }, [shiftList, allowedShiftNames]);
      const [selectedShift, setSelectedShift] = useState('');
      // Pilih shift paling pagi dari shift yang diizinkan saja.
      useEffect(() => {
        if (!sortedShiftList.some(s => s.nama_shift === selectedShift)) {
          setSelectedShift(sortedShiftList[0]?.nama_shift || '');
        }
      }, [sortedShiftList, selectedShift]);
      const [statusPulang, setStatusPulang] = useState('Pulang');
      const [alasanPulang, setAlasanPulang] = useState('');

      const videoRef = useRef(null);
      const [capturedPhoto, setCapturedPhoto] = useState(null);
      const [cameraStream, setCameraStream] = useState(null);
      const [isUploading, setIsUploading] = useState(false);

      const mapContainerRef = useRef(null);
      const mapInstanceRef = useRef(null);
      const userMarkerRef = useRef(null);
      const officeMarkerRef = useRef(null);
      const officeCircleRef = useRef(null);

      const getGPSLocation = () => {
        // Unlock Lokasi hanya melewati validasi radius; GPS dan peta tetap aktif untuk pratinjau posisi.
        if (!selectedLokasi || !Number.isFinite(Number(selectedLokasi.latitude)) || !Number.isFinite(Number(selectedLokasi.longitude))) {
          setGpsLoading(false);
          setUserCoords(null);
          setDistance(null);
          setInRadius(false);
          return;
        }
        setGpsLoading(true);

        if (!navigator.geolocation) {
          setUserCoords(null);
          setDistance(null);
          setInRadius(false);
          setGpsAccuracy(null);
          setGpsLoading(false);
          alert('Perangkat atau browser tidak mendukung GPS. Aktifkan lokasi atau gunakan browser yang mendukung geolocation.');
          return;
        }

        navigator.geolocation.getCurrentPosition(
          (pos) => {
            const { latitude, longitude, accuracy } = pos.coords;
            const dist = calculateDistance(latitude, longitude, Number(selectedLokasi.latitude), Number(selectedLokasi.longitude));
            setUserCoords({ lat: latitude, lng: longitude });
            setDistance(dist);
            setInRadius(dist <= Number(selectedLokasi.radius_meter));
            setGpsAccuracy(Math.round(accuracy || 0));
            setGpsLoading(false);
          },
          (err) => {
            console.warn('GPS gagal diperoleh:', err);
            setUserCoords(null);
            setDistance(null);
            setInRadius(false);
            setGpsAccuracy(null);
            setGpsLoading(false);
            const message = err && err.code === 1
              ? 'Izin lokasi ditolak. Aktifkan izin lokasi untuk melakukan presensi.'
              : err && err.code === 3
                ? 'Waktu mendapatkan lokasi habis. Periksa GPS dan coba deteksi ulang.'
                : 'Lokasi GPS tidak dapat diperoleh. Periksa pengaturan lokasi perangkat lalu coba lagi.';
            alert(message);
          },
          { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 }
        );
      };

      useEffect(() => {
        getGPSLocation();
      }, [selectedLokasi]);

      useEffect(() => {
        if (step !== 1 || !mapContainerRef.current || !selectedLokasi || !Number.isFinite(Number(selectedLokasi.latitude)) || !Number.isFinite(Number(selectedLokasi.longitude))) return;

        // Inisialisasi peta jika belum ada
        if (!mapInstanceRef.current) {
          const map = L.map(mapContainerRef.current, {
            zoomControl: false,
            attributionControl: true
          }).setView([selectedLokasi.latitude, selectedLokasi.longitude], 16);

          // Tile OpenStreetMap tidak memerlukan API key, sehingga menghindari pesan API KEY REQUIRED.
          L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            maxZoom: 19,
            attribution: '&copy; <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noopener noreferrer">OpenStreetMap</a> contributors'
          }).addTo(map);

          mapInstanceRef.current = map;
        }

        const map = mapInstanceRef.current;

        // Hapus layer lama untuk menghindari duplikasi
        if (officeCircleRef.current) map.removeLayer(officeCircleRef.current);
        if (officeMarkerRef.current) map.removeLayer(officeMarkerRef.current);
        if (userMarkerRef.current) map.removeLayer(userMarkerRef.current);

        // 1. Gambar Radius & Penanda Kantor
        const officeCircle = L.circle([selectedLokasi.latitude, selectedLokasi.longitude], {
          color: '#8B5CF6',
          fillColor: '#C4B5FD',
          fillOpacity: 0.25,
          weight: 2,
          radius: selectedLokasi.radius_meter
        }).addTo(map);
        officeCircleRef.current = officeCircle;

        const officeIcon = L.divIcon({
          className: 'custom-office-pin',
          html: `<div style="background-color:#7C3AED;color:white;width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;box-shadow:0 4px 10px rgba(124,58,237,0.4);border:2px solid white;"><i class="fa-solid fa-building text-[11px]"></i></div>`,
          iconSize: [28, 28],
          iconAnchor: [14, 14]
        });
        officeMarkerRef.current = L.marker([selectedLokasi.latitude, selectedLokasi.longitude], { icon: officeIcon })
          .addTo(map)
          .bindPopup(`<b>${selectedLokasi.nama_lokasi}</b><br>Radius: ${selectedLokasi.radius_meter}m`);

        // 2. Gambar Penanda Posisi Pengguna (User Marker)
        if (userCoords) {
          const userIcon = L.divIcon({
            className: 'custom-user-marker',
            html: `<div style="position:relative;width:24px;height:24px;display:flex;align-items:center;justify-content:center;">
              <div style="position:absolute;width:100%;height:100%;border-radius:50%;background-color:${inRadius ? '#10B981' : '#EF4444'};opacity:0.35;animation:ping 1.5s cubic-bezier(0,0,0.2,1) infinite;"></div>
              <div style="width:14px;height:14px;border-radius:50%;background-color:${inRadius ? '#10B981' : '#EF4444'};border:2.5px solid white;box-shadow:0 2px 8px rgba(0,0,0,0.3);z-index:2;"></div>
            </div>`,
            iconSize: [24, 24],
            iconAnchor: [12, 12]
          });

          userMarkerRef.current = L.marker([userCoords.lat, userCoords.lng], { icon: userIcon })
            .addTo(map)
            .bindPopup(`<b>Posisi Anda</b><br>Jarak: ${distance ?? 0}m (${inRadius ? 'Dalam Area' : 'Luar Area'})`);

          // Sesuaikan tampilan peta agar kantor dan posisi user muat di layar
          try {
            const group = L.featureGroup([officeCircle, userMarkerRef.current]);
            map.fitBounds(group.getBounds().pad(0.35), { maxZoom: 17 });
          } catch (e) {
            map.setView([userCoords.lat, userCoords.lng], 16);
          }
        } else {
          map.setView([selectedLokasi.latitude, selectedLokasi.longitude], 16);
        }

        // Koreksi ukuran kontainer Leaflet pada modal
        const resizeTimer = setTimeout(() => {
          if (mapInstanceRef.current) mapInstanceRef.current.invalidateSize();
        }, 300);

        return () => clearTimeout(resizeTimer);
      }, [step, userCoords, selectedLokasi, inRadius, distance]);

      useEffect(() => {
        return () => {
          if (mapInstanceRef.current) {
            mapInstanceRef.current.remove();
            mapInstanceRef.current = null;
          }
        };
      }, []);

      useEffect(() => {
        if (step === 3 && !capturedPhoto) {
          navigator.mediaDevices?.getUserMedia({ video: { facingMode: 'user' }, audio: false })
            .then(stream => { setCameraStream(stream); if (videoRef.current) videoRef.current.srcObject = stream; })
            .catch(err => console.warn("Camera access error:", err));
        }
        return () => { if (cameraStream) cameraStream.getTracks().forEach(track => track.stop()); };
      }, [step, capturedPhoto]);

      const captureSelfie = () => {
        const video = videoRef.current;
        if (!video || !video.videoWidth || !video.videoHeight) {
          showToast('Kamera belum siap. Tunggu sebentar lalu coba ambil foto lagi.', 'error');
          return;
        }

        // Gunakan rasio asli kamera agar hasil foto tidak gepeng atau memanjang.
        // Batasi sisi terpanjang maksimal 1280 px agar ukuran file tetap wajar.
        const sourceWidth = video.videoWidth;
        const sourceHeight = video.videoHeight;
        const scale = Math.min(1, 1280 / Math.max(sourceWidth, sourceHeight));
        const outputWidth = Math.max(1, Math.round(sourceWidth * scale));
        const outputHeight = Math.max(1, Math.round(sourceHeight * scale));
        const canvas = document.createElement('canvas');
        canvas.width = outputWidth;
        canvas.height = outputHeight;

        const ctx = canvas.getContext('2d');
        if (!ctx) {
          showToast('Tidak dapat memproses foto kamera. Silakan coba lagi.', 'error');
          return;
        }

        // Cerminkan horizontal seperti preview kamera depan, tanpa mengubah rasio vertikal.
        ctx.save();
        ctx.translate(canvas.width, 0);
        ctx.scale(-1, 1);
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        ctx.restore();

        setCapturedPhoto(canvas.toDataURL('image/jpeg', 0.88));
        if (cameraStream) cameraStream.getTracks().forEach(track => track.stop());
      };

      const handleFinalSubmit = async () => {
        if (isUploading) return;
        if (!selectedLokasi) { alert('Silakan pilih lokasi presensi terlebih dahulu.'); return; }
        if (!locationUnlocked && !inRadius) { alert('Posisi Anda masih di luar radius lokasi yang dipilih.'); return; }
        if (type === 'masuk' && !selectedShift) { alert('Silakan pilih shift kerja terlebih dahulu.'); return; }
        setIsUploading(true);
        try {
        let finalStatus = 'Hadir';
        if (type === 'masuk') {
          const currentShiftObj = shiftList.find(s => s.nama_shift === selectedShift);
          if (currentShiftObj) {
            const [shiftH, shiftM] = currentShiftObj.jam_masuk.split(':').map(Number);
            const shiftTime = new Date();
            shiftTime.setHours(shiftH, shiftM + toleransiMenit, 0, 0);
            if (new Date() > shiftTime) finalStatus = 'Terlambat';
          }
        } else {
          finalStatus = statusPulang;
        }

        const now = new Date();
        const dateStr = getLocalDateKey(now);
        const timeStr = [now.getHours(), now.getMinutes(), now.getSeconds()].map(v => String(v).padStart(2, '0')).join(':');

        // Unggah foto selfie ke Google Drive via GAS Web App
        let uploadedPhotoUrl = capturedPhoto;
        if (capturedPhoto && capturedPhoto.startsWith('data:image')) {
          uploadedPhotoUrl = await uploadToGoogleDriveViaGAS(
            capturedPhoto,
            `selfie_${currentUser.id}_${type}_${dateStr}.jpg`,
            'presensi'
          );
        }

        const newRecord = {
          id: 'prs-' + Date.now(),
          karyawan_id: currentUser.id,
          karyawan_nama: currentUser.nama_lengkap,
          divisi: currentUser.divisi,
          tipe: type,
          status: finalStatus,
          shift_nama: selectedShift,
          lokasi_id: selectedLokasi.id,
          lokasi_nama: locationUnlocked ? `${selectedLokasi.nama_lokasi} (Unlock Lokasi HR)` : selectedLokasi.nama_lokasi,
          latitude: userCoords?.lat ?? null,
          longitude: userCoords?.lng ?? null,
          foto_url: uploadedPhotoUrl,
          keterangan: type === 'keluar' && statusPulang !== 'Pulang' ? alasanPulang : (finalStatus === 'Terlambat' ? 'Terlambat masuk' : 'Presensi App'),
          tanggal: dateStr,
          jam: timeStr,
          created_at: now.toISOString()
        };

        // Simpan ke Supabase dan periksa error dari proses upload maupun insert.
        // Gunakan satu blok try/finally agar modal tidak tertinggal dalam kondisi loading.
          const client = await getSupabaseClient();
          const { error } = await client.from('presensi').insert([newRecord]);
          if (error) throw error;
          onSuccess(newRecord);
        } catch (err) {
          console.error('Gagal menyimpan presensi:', err);
          alert(`Presensi belum tersimpan: ${err?.message || 'Terjadi kesalahan saat mengunggah foto atau menyimpan ke database.'}. Silakan coba lagi.`);
        } finally {
          setIsUploading(false);
        }
      };

      return (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-[2rem] max-w-sm w-full p-6 shadow-2xl flex flex-col border border-slate-100">
            {/* Header Wizard */}
            <div className="flex justify-between items-center mb-4">
              <div>
                <span className="text-[10px] font-bold uppercase text-slate-400">Langkah {step} dari 3</span>
                <h3 className="font-display font-bold text-base text-slate-800">Presensi {type === 'masuk' ? 'Masuk' : 'Pulang'}</h3>
              </div>
              <button onClick={onClose} className="w-8 h-8 rounded-full bg-slate-50 text-slate-400 hover:text-slate-600">
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>

            {/* Stepper Bar */}
            <div className="flex items-center space-x-2 mb-4">
              <div className={`h-1.5 flex-1 rounded-full ${step >= 1 ? 'bg-violet-600' : 'bg-slate-100'}`}></div>
              <div className={`h-1.5 flex-1 rounded-full ${step >= 2 ? 'bg-violet-600' : 'bg-slate-100'}`}></div>
              <div className={`h-1.5 flex-1 rounded-full ${step >= 3 ? 'bg-violet-600' : 'bg-slate-100'}`}></div>
            </div>

            {/* LANGKAH 1: Validasi Lokasi & Geofence GPS */}
            {step === 1 && !selectedLokasi && (
              <div className="py-8 text-center space-y-3">
                <i className="fa-solid fa-location-dot text-3xl text-amber-500"></i>
                <p className="text-sm font-bold text-slate-700">Lokasi kantor belum tersedia</p>
                <p className="text-xs text-slate-500">Data lokasi sedang dimuat atau belum diatur oleh admin. Tutup jendela ini dan coba lagi setelah data tersedia.</p>
                <button onClick={onClose} className="w-full py-3 rounded-full bg-slate-900 text-white text-xs font-bold">Tutup</button>
              </div>
            )}
            {step === 1 && selectedLokasi && (<div className="space-y-3.5">
                {locationUnlocked && <div className="rounded-2xl bg-emerald-50 border border-emerald-200 p-3 text-[11px] leading-relaxed text-emerald-800"><b><i className="fa-solid fa-lock-open mr-1.5"></i>Unlock Lokasi aktif.</b> Pilih lokasi kantor dan lihat posisi Anda di peta seperti biasa. Anda tetap bisa melanjutkan presensi dari luar radius kantor.</div>}
                
                <div>
                  <label className="text-[10px] font-bold text-slate-400 block mb-1">Pilih Titik Kantor Tujuan</label>
                  <select 
                    value={String(selectedLokasi.id)} 
                    onChange={(e) => { const nextLocation = lokasiList.find(l => String(l.id) === e.target.value); if (nextLocation) setSelectedLokasi(nextLocation); }} 
                    className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold text-slate-700 outline-none"
                  >
                    {availableLokasiList.map(l => (
                      <option key={l.id} value={l.id}>{l.nama_lokasi} (Radius: {l.radius_meter}m)</option>
                    ))}
                  </select>
                </div>

                {/* Wadah Peta Leaflet */}
                <div className="relative h-44 w-full rounded-2xl overflow-hidden border border-slate-200 shadow-inner bg-slate-100">
                  <div ref={mapContainerRef} className="w-full h-full"></div>
                  <button
                    type="button"
                    onClick={() => {
                      const coords = userCoords || { lat: selectedLokasi.latitude, lng: selectedLokasi.longitude };
                      window.open(`https://www.google.com/maps?q=${coords.lat},${coords.lng}`, '_blank', 'noopener,noreferrer');
                    }}
                    className="absolute top-2 right-2 z-[450] bg-white hover:bg-slate-50 text-slate-700 border border-slate-200 rounded-lg px-2.5 py-1.5 text-[10px] font-bold shadow-md"
                    title="Buka titik lokasi di Google Maps"
                  >
                    <i className="fa-solid fa-map-location-dot mr-1.5 text-emerald-600"></i>Google Maps
                  </button>
                  {gpsLoading && (
                    <div className="absolute inset-0 bg-white/85 backdrop-blur-xs flex flex-col items-center justify-center text-xs font-bold text-slate-600 z-[400]">
                      <i className="fa-solid fa-location-crosshairs fa-spin text-xl text-violet-600 mb-1.5"></i>
                      <span>Mengunci Sinyal GPS...</span>
                    </div>
                  )}
                  {/* Legenda kecil di pojok peta */}
                  <div className="absolute bottom-2 left-2 z-[400] bg-white/90 backdrop-blur-md px-2 py-1 rounded-lg text-[9px] font-bold text-slate-600 border border-slate-200 flex items-center gap-2 shadow-xs">
                    <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-violet-600"></span> Kantor</span>
                    <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-emerald-500"></span> Anda</span>
                  </div>
                </div>

                {/* Kotak Status Radius & Jarak Geofencing */}
                <div className={`p-3.5 rounded-2xl flex items-center justify-between border ${(locationUnlocked || inRadius) ? 'bg-emerald-50 border-emerald-200 text-emerald-700' : 'bg-rose-50 border-rose-200 text-rose-700'}`}>
                  <div className="flex items-center space-x-3">
                    <div className={`w-9 h-9 rounded-full flex items-center justify-center ${(locationUnlocked || inRadius) ? 'bg-emerald-100 text-emerald-600' : 'bg-rose-100 text-rose-600'}`}>
                      <i className={`fa-solid ${locationUnlocked ? 'fa-lock-open' : (inRadius ? 'fa-circle-check' : 'fa-circle-xmark')} text-base`}></i>
                    </div>
                    <div>
                      <p className="text-xs font-bold">{locationUnlocked ? (gpsLoading ? 'Unlock Lokasi HR • Mencari GPS…' : 'Unlock Lokasi HR • Radius diabaikan') : (inRadius ? 'Dalam Radius Kantor' : 'Di Luar Radius Kantor')}</p>
                      <p className="text-[10px] font-mono-numbers opacity-85">
                        Jarak: <b>{distance !== null ? `${distance}m` : '...'}</b> / Maks: {selectedLokasi.radius_meter}m
                        {gpsAccuracy && ` • Akurasi: ±${gpsAccuracy}m`}
                      </p>
                    </div>
                  </div>
                  <button 
                    type="button"
                    onClick={() => getGPSLocation()} 
                    className="p-2 rounded-xl hover:bg-black/5 active:scale-95 transition"
                    title="Deteksi Ulang GPS Perangkat"
                  >
                    <i className="fa-solid fa-arrows-rotate text-xs"></i>
                  </button>
                </div>

                {/* Tombol Lanjut */}
                <button 
                  onClick={() => setStep(2)} 
                  disabled={!locationUnlocked && (!inRadius || gpsLoading)} 
                  className={`w-full py-3.5 rounded-full font-bold text-xs transition shadow-sm ${
                    (locationUnlocked || (inRadius && !gpsLoading))
                      ? 'bg-slate-900 text-white hover:bg-slate-800 active:scale-95' 
                      : 'bg-slate-100 text-slate-400 cursor-not-allowed'
                  }`}
                >
                  {locationUnlocked ? 'Lanjutkan Presensi (Unlock Lokasi)' : 'Lanjutkan Presensi'}
                </button>
              </div>)}

            {/* LANGKAH 2: Pemilihan Shift / Alasan Pulang */}
            {step === 2 && (
              <div className="space-y-4">
                {type === 'masuk' ? (
                  <div className="space-y-2.5">
                    <label className="text-[10px] font-bold text-slate-400 block mb-1">Pilih Shift Kerja Hari Ini</label>
                    {sortedShiftList.length > 0 ? sortedShiftList.map(s => (
                      <div 
                        key={s.id} 
                        onClick={() => setSelectedShift(s.nama_shift)} 
                        className={`p-3.5 rounded-2xl border text-xs cursor-pointer transition flex justify-between items-center ${
                          selectedShift === s.nama_shift 
                            ? 'bg-violet-50 border-violet-200 text-violet-700 shadow-sm' 
                            : 'bg-white border-slate-100 text-slate-600 hover:bg-slate-50'
                        }`}
                      >
                        <div>
                          <span className="font-bold block text-sm">{s.nama_shift}</span>
                          <span className="text-[10px] opacity-80">{formatShiftTimeDisplay(s.jam_masuk)} - {formatShiftTimeDisplay(s.jam_pulang)} WIB</span>
                        </div>
                        {selectedShift === s.nama_shift && <i className="fa-solid fa-circle-check text-violet-500 text-lg"></i>}
                      </div>
                    )) : (
                      <div className="rounded-2xl border border-amber-200 bg-amber-50 p-4 text-center">
                        <i className="fa-solid fa-triangle-exclamation text-amber-500 text-xl mb-2"></i>
                        <p className="text-xs font-bold text-amber-800">Belum ada shift yang ditetapkan untuk divisi {currentUser?.divisi || 'Anda'}.</p>
                        <p className="text-[10px] text-amber-700 mt-1">Silakan hubungi administrator untuk mengatur shift di menu Master Divisi.</p>
                      </div>
                    )}
                  </div>
                ) : (
                  <div className="space-y-3">
                    <label className="text-[10px] font-bold text-slate-400 block mb-1">Status Kepulangan</label>
                    <select 
                      value={statusPulang} 
                      onChange={(e) => setStatusPulang(e.target.value)} 
                      className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold text-slate-700 outline-none"
                    >
                      <option value="Pulang">Pulang Selesai Shift</option>
                      <option value="Izin Pulang Cepat">Pulang Lebih Cepat (Izin)</option>
                      <option value="Sakit Mendadak">Pulang Karena Sakit</option>
                    </select>
                    {statusPulang !== 'Pulang' && (
                      <textarea 
                        rows="3" 
                        value={alasanPulang} 
                        onChange={(e) => setAlasanPulang(e.target.value)} 
                        placeholder="Tuliskan keterangan lengkap..." 
                        className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 outline-none"
                        required
                      ></textarea>
                    )}
                  </div>
                )}

                <div className="flex gap-2 pt-2">
                  <button onClick={() => setStep(1)} className="w-1/3 py-3.5 rounded-full bg-slate-100 text-slate-600 font-bold text-xs">
                    Kembali
                  </button>
                  <button onClick={() => setStep(3)} className="w-2/3 py-3.5 rounded-full bg-slate-900 text-white font-bold text-xs shadow-md active:scale-95 transition">
                    Lanjut Selfie
                  </button>
                </div>
              </div>
            )}

            {/* LANGKAH 3: Pengambilan Foto Selfie */}
            {step === 3 && (
              <div className="space-y-4">
                <div className="relative w-full aspect-square bg-slate-100 rounded-3xl overflow-hidden flex items-center justify-center border border-slate-200 shadow-inner">
                  {!capturedPhoto ? (
                    <video ref={videoRef} autoPlay playsInline className="w-full h-full object-cover mirror-video"></video>
                  ) : (
                    <img src={capturedPhoto} alt="Hasil Selfie" className="w-full h-full object-cover" />
                  )}
                </div>

                {!capturedPhoto ? (
                  <button 
                    onClick={captureSelfie} 
                    className="w-full py-3.5 bg-violet-600 hover:bg-violet-700 text-white rounded-full font-bold text-xs shadow-md active:scale-95 transition flex items-center justify-center gap-2"
                  >
                    <i className="fa-solid fa-camera"></i>
                    <span>Ambil Foto Selfie</span>
                  </button>
                ) : (
                  <div className="flex gap-2">
                    <button 
                      onClick={() => setCapturedPhoto(null)} 
                      className="w-1/3 py-3.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-full font-bold text-xs transition"
                    >
                      Ulangi Foto
                    </button>
                    <button 
                      onClick={handleFinalSubmit} 
                      disabled={isUploading} 
                      className="w-2/3 py-3.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full font-bold text-xs shadow-md active:scale-95 transition flex items-center justify-center gap-2"
                    >
                      {isUploading ? (
                        <>
                          <i className="fa-solid fa-spinner fa-spin"></i>
                          <span>Menyimpan Presensi...</span>
                        </>
                      ) : (
                        <>
                          <i className="fa-solid fa-paper-plane"></i>
                          <span>Kirim Presensi</span>
                        </>
                      )}
                    </button>
                  </div>
                )}
                <button onClick={() => setStep(2)} className="w-full py-1 text-xs text-slate-400 font-bold hover:text-slate-600">
                  Kembali ke Langkah Sebelumnya
                </button>
              </div>
            )}
          </div>
        </div>
      );
    }

    function SimpleActionModal({ type, currentUser, setCurrentUser, onClose, onSubmitAttendance, onSubmitLembur }) {
      const [keterangan, setKeterangan] = useState('');
      const [durasiJam, setDurasiJam] = useState('');
      const [fotoBukti, setFotoBukti] = useState('');
      const [isSubmitting, setIsSubmitting] = useState(false);

      const handleSubmit = async () => {
        setIsSubmitting(true);
        const now = new Date();
        const dateStr = getLocalDateKey(now);

        // Unggah foto bukti ke Google Drive via GAS jika ada.
        // Jika upload gagal, pulihkan status tombol dan pertahankan preview agar bisa dicoba ulang.
        let uploadedBuktiUrl = fotoBukti;
        if (fotoBukti && fotoBukti.startsWith('data:image')) {
          try {
            const folderBukti = ['sakit', 'izin', 'cuti', 'lembur'].includes(String(type).toLowerCase())
              ? String(type).toLowerCase()
              : 'izin';
            uploadedBuktiUrl = await uploadToGoogleDriveViaGAS(
              fotoBukti,
              `bukti_${type}_${currentUser.id}_${dateStr}.jpg`,
              folderBukti
            );
            if (uploadedBuktiUrl) setFotoBukti(uploadedBuktiUrl);
          } catch (uploadError) {
            console.error('Upload bukti gagal:', uploadError);
            setIsSubmitting(false);
            alert(`Foto bukti belum berhasil diunggah. Preview tetap tersedia.

${uploadError.message || 'Silakan coba lagi.'}`);
            return;
          }
        }

        if (type === 'lembur') {
          const durasiValid = Number(durasiJam);
          if (!durasiJam || !Number.isFinite(durasiValid) || durasiValid < 1 || durasiValid > 12) {
            setIsSubmitting(false);
            alert('Isi durasi lembur terlebih dahulu (1–12 jam).');
            return;
          }
          const newLembur = {
            id: 'lmb-' + Date.now(),
            karyawan_id: currentUser.id,
            karyawan_nama: currentUser.nama_lengkap,
            divisi: currentUser.divisi,
            tanggal: dateStr,
            durasi_jam: durasiValid,
            keterangan: keterangan || 'Lembur',
            bukti_url: uploadedBuktiUrl,
            status: 'Pending',
            created_at: now.toISOString()
          };

          if (supabase) {
            try {
              await supabase.from('pengajuan_lembur').insert([newLembur]);
            } catch (err) {
              console.warn("Gagal simpan lembur ke Supabase:", err);
            }
          }

          setIsSubmitting(false);
          onSubmitLembur(newLembur);
          return;
        }

        let statusMap = { sakit: 'Sakit', libur: 'Libur', izin: 'Izin', cuti: 'Cuti' };
        const newAttendance = {
          id: 'prs-' + Date.now(),
          karyawan_id: currentUser.id,
          karyawan_nama: currentUser.nama_lengkap,
          divisi: currentUser.divisi,
          tipe: 'masuk',
          status: statusMap[type] || 'Hadir',
          shift_nama: 'Pengajuan Form',
          lokasi_nama: 'Mandiri',
          latitude: 0,
          longitude: 0,
          foto_url: uploadedBuktiUrl,
          keterangan: keterangan || `Pengajuan ${type}`,
          tanggal: dateStr,
          jam: [now.getHours(), now.getMinutes(), now.getSeconds()].map(v => String(v).padStart(2, '0')).join(':'),
          created_at: now.toISOString()
        };

        if (supabase) {
          try {
            await supabase.from('presensi').insert([newAttendance]);
          } catch (err) {
            console.warn("Gagal simpan absensi form ke Supabase:", err);
          }
        }

        if (type === 'libur' && currentUser.sisa_libur > 0) {
          const updatedUser = { ...currentUser, sisa_libur: currentUser.sisa_libur - 1 };
          setCurrentUser(updatedUser);
          if (supabase) supabase.from('master_karyawan').update({ sisa_libur: updatedUser.sisa_libur }).eq('id', currentUser.id);
        } else if (type === 'cuti' && currentUser.sisa_cuti > 0) {
          const updatedUser = { ...currentUser, sisa_cuti: currentUser.sisa_cuti - 1 };
          setCurrentUser(updatedUser);
          if (supabase) supabase.from('master_karyawan').update({ sisa_cuti: updatedUser.sisa_cuti }).eq('id', currentUser.id);
        }

        setIsSubmitting(false);
        onSubmitAttendance(newAttendance);
      };

      return (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-[2rem] max-w-sm w-full p-6 shadow-2xl border border-slate-100">
            <div className="flex justify-between items-center mb-5">
              <h3 className="font-display font-bold text-lg text-slate-800 capitalize">Form {type}</h3>
              <button onClick={onClose} className="w-8 h-8 rounded-full bg-slate-50 text-slate-400"><i className="fa-solid fa-xmark"></i></button>
            </div>
            <div className="space-y-4">
              {type === 'lembur' && (
                <div><label className="text-[11px] font-bold text-slate-500 block mb-1">Durasi (Jam)</label><input type="number" min="1" max="12" value={durasiJam} placeholder="Masukkan durasi jam" onChange={(e) => setDurasiJam(e.target.value)} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none" /></div>
              )}
              <div><label className="text-[11px] font-bold text-slate-500 block mb-1">Keterangan Lengkap</label><textarea rows="3" value={keterangan} onChange={(e) => setKeterangan(e.target.value)} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 outline-none"></textarea></div>
              <div>
                <label className="text-[11px] font-bold text-slate-500 block mb-1">Foto Bukti (Opsional)</label>
                <input
                  type="file"
                  accept="image/*"
                  onChange={(e) => {
                    const file = e.target.files && e.target.files[0];
                    if (!file) return;
                    if (!file.type || !file.type.startsWith('image/')) {
                      alert('Pilih file gambar yang valid.');
                      e.target.value = '';
                      return;
                    }
                    const reader = new FileReader();
                    reader.onloadend = () => setFotoBukti(reader.result);
                    reader.onerror = () => alert('Gambar gagal dibaca. Silakan pilih ulang.');
                    reader.readAsDataURL(file);
                  }}
                  className="w-full text-xs text-slate-500 file:mr-3 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-[10px] file:font-bold file:bg-violet-50 file:text-violet-600"
                />
                {fotoBukti && (
                  <div className="mt-3 rounded-2xl border border-violet-100 bg-violet-50/60 p-3">
                    <div className="flex items-center justify-between gap-2 mb-2">
                      <span className="text-[10px] font-bold text-violet-700 flex items-center gap-1.5">
                        <i className="fa-solid fa-image"></i>
                        Preview bukti {fotoBukti.startsWith('data:image') ? '(siap diunggah)' : '(tersimpan)'}
                      </span>
                      <button type="button" onClick={() => setFotoBukti('')} className="text-[10px] font-bold text-rose-500 hover:text-rose-700">
                        Hapus gambar
                      </button>
                    </div>
                    <img
                      src={fotoBukti}
                      alt={`Preview bukti ${type}`}
                      className="w-full max-h-56 object-contain rounded-xl bg-white border border-slate-100"
                      onError={(e) => { e.currentTarget.style.display = 'none'; }}
                    />
                  </div>
                )}
                <p className="text-[10px] text-slate-400 mt-1">Gambar akan langsung terlihat sebagai preview sebelum pengajuan dikirim.</p>
              </div>
              <button 
                onClick={handleSubmit} 
                disabled={isSubmitting} 
                className="w-full py-3.5 bg-slate-900 text-white rounded-full font-bold text-xs shadow-md active:scale-95 flex items-center justify-center gap-2"
              >
                {isSubmitting ? (
                  <>
                    <i className="fa-solid fa-spinner fa-spin"></i>
                    <span>Mengirim...</span>
                  </>
                ) : (
                  <span>Kirim Pengajuan</span>
                )}
              </button>
            </div>
          </div>
        </div>
      );
    }

    function DaftarPengumumanKaryawanView({ pengumumanList, onSelectAnnouncement, navigateTo }) {
      const [searchQuery, setSearchQuery] = useState('');
      const filteredList = useMemo(() => {
        return pengumumanList.filter(item => {
          const matchSearch = !searchQuery.trim() || 
            item.judul.toLowerCase().includes(searchQuery.toLowerCase()) ||
            item.isi.toLowerCase().includes(searchQuery.toLowerCase()) ||
            (item.author && item.author.toLowerCase().includes(searchQuery.toLowerCase()));
          return matchSearch;
        });
      }, [pengumumanList, searchQuery]);

      return (
        <div className="p-5 space-y-4 pb-28">
          {/* Header Bar */}
          <div className="flex items-center justify-between mb-2">
            <div className="flex items-center space-x-3">
              <button 
                onClick={() => navigateTo('#karyawan')}
                className="w-9 h-9 rounded-full bg-white border border-slate-200 flex items-center justify-center text-slate-600 shadow-sm active:scale-95"
                title="Kembali ke Beranda"
              >
                <i className="fa-solid fa-arrow-left text-xs"></i>
              </button>
              <div>
                <h2 className="font-display font-bold text-lg text-slate-800">Pengumuman</h2>
                <p className="text-[11px] text-slate-400 font-medium">Pemberitahuan resmi & informasi HR</p>
              </div>
            </div>
            <span className="px-3 py-1 bg-white border border-slate-100 shadow-sm text-slate-700 rounded-full text-[10px] font-bold">
              {filteredList.length} Pesan
            </span>
          </div>

          {/* Kotak Pencarian */}
          <div className="relative">
            <i className="fa-solid fa-magnifying-glass absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
            <input 
              type="text" 
              placeholder="Cari pengumuman..." 
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full bg-white border border-slate-200/80 rounded-2xl pl-9 pr-4 py-2.5 text-xs font-semibold outline-none focus:border-violet-500 shadow-sm"
            />
          </div>

          {/* List Kartu Pengumuman */}
          <div className="space-y-3">
            {filteredList.length > 0 ? (
              filteredList.map(item => (
                <div 
                  key={item.id}
                  onClick={() => onSelectAnnouncement(item)}
                  className="bg-white/95 backdrop-blur-xl p-4 rounded-3xl border border-slate-100 shadow-[0_4px_20px_rgb(0,0,0,0.03)] hover:shadow-md cursor-pointer transition active:scale-[0.99] space-y-2.5"
                >
                  <div className="flex items-center justify-end">
                    <span className="text-[10px] font-mono-numbers text-slate-400 font-medium">
                      {new Date(item.created_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' })}
                    </span>
                  </div>

                  <div>
                    <h3 className="font-display font-bold text-sm text-slate-800 leading-snug hover:text-violet-600 transition">
                      {item.judul}
                    </h3>
                    <p className="text-[11px] text-slate-500 font-medium line-clamp-2 mt-1 leading-relaxed">
                      {item.isi}
                    </p>
                  </div>

                  <div className="pt-2 border-t border-slate-100 flex items-center justify-between text-[10px] font-bold text-slate-400">
                    <span className="flex items-center gap-1.5 text-slate-600">
                      <i className="fa-regular fa-building text-violet-500"></i>
                      {item.author || 'HRD'}
                    </span>
                    <span className="text-violet-600 flex items-center gap-1 font-semibold">
                      Baca Detail <i className="fa-solid fa-chevron-right text-[8px]"></i>
                    </span>
                  </div>
                </div>
              ))
            ) : (
              <div className="text-center py-12 text-slate-400 bg-white/50 rounded-3xl border border-slate-100">
                <i className="fa-regular fa-bell-slash text-2xl mb-2 text-slate-300"></i>
                <p className="text-xs font-medium">Tidak ada pengumuman yang sesuai.</p>
              </div>
            )}
          </div>
        </div>
      );
    }

    function StatusRekanView({ karyawanList, presensiList, currentUser }) {
      const [filterStatus, setFilterStatus] = React.useState('Semua');
      const [searchNama, setSearchNama] = React.useState('');
      const todayStr = getLocalDateKey();

      // Hitung jumlah rekan untuk masing-masing filter secara real-time hari ini
      const counts = useMemo(() => {
        let masuk = 0;
        let belumMasuk = 0;
        let libur = 0;
        let izin = 0;
        let sakit = 0;
        let cuti = 0;

        karyawanList.forEach(k => {
          const att = presensiList.find(p => p.karyawan_id === k.id && p.tanggal === todayStr && p.tipe === 'masuk');
          const status = att ? att.status : 'Belum Absen';

          if (status === 'Hadir' || status === 'Terlambat') masuk++;
          else if (status === 'Belum Absen') belumMasuk++;
          else if (status === 'Libur') libur++;
          else if (status === 'Izin') izin++;
          else if (status === 'Sakit') sakit++;
          else if (status === 'Cuti') cuti++;
        });

        return {
          'Semua': karyawanList.length,
          'Masuk': masuk,
          'Belum Masuk': belumMasuk,
          'Libur': libur,
          'Izin': izin,
          'Sakit': sakit,
          'Cuti': cuti
        };
      }, [karyawanList, presensiList, todayStr]);

      const filters = ['Semua', 'Masuk', 'Belum Masuk', 'Libur', 'Izin', 'Sakit', 'Cuti'];

      const filteredList = karyawanList.filter(k => {
        const nama = String(k.nama_lengkap || '').toLocaleLowerCase('id-ID');
        const query = searchNama.trim().toLocaleLowerCase('id-ID');
        if (query && !nama.includes(query)) return false;

        const att = presensiList.find(p => String(p.karyawan_id) === String(k.id) && p.tanggal === todayStr && p.tipe === 'masuk');
        const currentStatus = att ? att.status : 'Belum Absen';
        
        if (filterStatus === 'Semua') return true;
        if (filterStatus === 'Masuk') return currentStatus === 'Hadir' || currentStatus === 'Terlambat';
        if (filterStatus === 'Belum Masuk') return currentStatus === 'Belum Absen';
        return currentStatus === filterStatus;
      }).sort((a, b) => String(a.nama_lengkap || '').localeCompare(String(b.nama_lengkap || ''), 'id', { sensitivity: 'base' }));

      return (
        <div className="p-5 space-y-4 pb-28">
          <div className="flex justify-between items-center mb-2">
            <div>
              <h2 className="font-display font-bold text-lg text-slate-800">Status Rekan</h2>
              <p className="text-[11px] text-slate-400 font-medium">Pantauan hari ini ({formatIndoDate(new Date())})</p>
            </div>
            <span className="px-3 py-1 bg-white border border-slate-100 shadow-sm text-slate-700 rounded-full text-[10px] font-bold">
              {filteredList.length} Orang
            </span>
          </div>

          <div className="relative">
            <i className="fa-solid fa-magnifying-glass absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 text-sm"></i>
            <input
              type="text"
              value={searchNama}
              onChange={e => setSearchNama(e.target.value)}
              placeholder="Cari nama karyawan..."
              aria-label="Cari nama karyawan"
              className="w-full rounded-2xl border border-slate-200 bg-white/90 py-3 pl-11 pr-10 text-sm text-slate-700 placeholder:text-slate-400 shadow-sm outline-none transition focus:border-violet-300 focus:ring-4 focus:ring-violet-100"
            />
            {searchNama && <button type="button" onClick={() => setSearchNama('')} aria-label="Hapus pencarian" className="absolute right-3 top-1/2 -translate-y-1/2 rounded-full p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"><i className="fa-solid fa-xmark"></i></button>}
          </div>

          <div className="flex flex-wrap gap-1.5 text-[10px] font-bold">
            {filters.map(f => (
              <button 
                key={f} 
                onClick={() => setFilterStatus(f)} 
                className={`px-3 py-1.5 rounded-full transition shadow-sm flex items-center gap-1.5 ${
                  filterStatus === f 
                    ? 'bg-slate-900 text-white' 
                    : 'bg-white text-slate-600 border border-slate-100 hover:bg-slate-50'
                }`}
              >
                <span>{f}</span>
                <span className={`px-1.5 py-0.2 rounded-full text-[9px] font-mono-numbers font-bold ${
                  filterStatus === f ? 'bg-white/20 text-white' : 'bg-slate-100 text-slate-600'
                }`}>
                  {counts[f] ?? 0}
                </span>
              </button>
            ))}
          </div>

          <div className="space-y-3">
            {filteredList.length > 0 ? filteredList.map((k, index) => {
              const att = presensiList.find(p => String(p.karyawan_id) === String(k.id) && p.tanggal === todayStr && p.tipe === 'masuk');
              const pulangAtt = presensiList.find(p => String(p.karyawan_id) === String(k.id) && p.tanggal === todayStr && p.tipe === 'keluar');
              const sudahMasuk = !!att && ['Hadir', 'Terlambat'].includes(String(att.status || '').trim());
              return (
                <div key={k.id} className="bg-white/90 backdrop-blur-xl p-4 rounded-3xl border border-slate-100 shadow-[0_4px_20px_rgb(0,0,0,0.03)]">
                  <div className="flex items-start gap-3 min-w-0">
                    <span className="w-5 shrink-0 pt-1 text-center text-[10px] font-bold text-slate-400">{index + 1}.</span>
                    <div className="w-11 h-11 shrink-0 rounded-full bg-slate-50 flex items-center justify-center overflow-hidden border border-slate-100">
                      {k.foto_profil ? <img src={getCleanImageUrl(k.foto_profil)} alt={`Foto ${k.nama_lengkap || 'karyawan'}`} className="w-full h-full object-cover" /> : <span className="font-bold text-sm text-slate-400">{String(k.nama_lengkap || '?').charAt(0)}</span>}
                    </div>
                    <div className="flex-1 min-w-0 space-y-1.5">
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <h4 className="text-sm font-bold text-slate-800">{k.nama_lengkap || 'Nama belum diisi'}</h4>
                        {currentUser && String(k.id) === String(currentUser.id) && <span className="px-1.5 py-0.5 rounded-full bg-indigo-50 text-indigo-600 text-[8px] font-extrabold">ANDA</span>}
                      </div>
                      <p className="text-[10px] text-slate-400 font-medium">
                        {k.divisi || 'Divisi belum diatur'}{k.jabatan ? ` - ${k.jabatan}` : ''}
                      </p>
                      <div className="pt-0.5 flex flex-wrap items-center gap-1.5">
                        {att ? (
                          <span className={`inline-flex px-3 py-1 rounded-full text-[9px] font-bold ${
                            att.status === 'Hadir' ? 'bg-emerald-50 text-emerald-600' :
                            att.status === 'Terlambat' ? 'bg-amber-50 text-amber-600' :
                            att.status === 'Libur' ? 'bg-sky-50 text-sky-600' :
                            att.status === 'Izin' ? 'bg-orange-50 text-orange-600' :
                            att.status === 'Sakit' ? 'bg-amber-50 text-amber-600' :
                            att.status === 'Cuti' ? 'bg-purple-50 text-purple-600' :
                            'bg-slate-50 text-slate-600'
                          }`}>
                            <i className="fa-solid fa-arrow-right-to-bracket mr-1 text-[8px]"></i>
                            {att.status}{!['Sakit', 'Libur', 'Izin', 'Cuti'].includes(String(att.status || '').trim()) && att.jam ? ` · ${att.jam}` : ''}
                          </span>
                        ) : <span className="inline-flex px-3 py-1 rounded-full text-[9px] font-bold bg-slate-50 text-slate-400">Belum Absen</span>}
                        {sudahMasuk && (
                          pulangAtt ? (
                            <span className="inline-flex px-3 py-1 rounded-full text-[9px] font-bold bg-rose-50 text-rose-600" title="Status dan jam pulang">
                              <i className="fa-solid fa-arrow-right-from-bracket mr-1 text-[8px]"></i>
                              Pulang{pulangAtt.jam ? ` · ${pulangAtt.jam}` : ''}
                            </span>
                          ) : (
                            <span className="inline-flex px-3 py-1 rounded-full text-[9px] font-bold bg-slate-50 text-slate-500" title="Belum melakukan presensi pulang">
                              <i className="fa-regular fa-clock mr-1 text-[8px]"></i>Belum Pulang
                            </span>
                          )
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              );
            }) : (
              <div className="text-center py-10 text-slate-400">
                <p className="text-xs font-medium">{searchNama.trim() ? `Tidak ada karyawan dengan nama “${searchNama.trim()}”` : `Tidak ada rekan dengan status ${filterStatus}`}</p>
              </div>
            )}
          </div>
        </div>
      );
    }

    function RiwayatAktivitasView({ currentUser, presensiList, klaimList = [], lemburList, payrollCycle, filterStatus, setFilterStatus }) {
      const [selectedPeriodIndex, setSelectedPeriodIndex] = useState(0);

      // Membuat daftar siklus 6 bulan ke belakang (Tanggal 26 s/d 25)
      const periods = useMemo(() => {
        const list = [];
        const now = new Date();
        let baseMonth = now.getDate() > 25 ? now.getMonth() + 1 : now.getMonth();
        let baseYear = now.getFullYear();

        for (let i = 0; i < 6; i++) {
          let start = new Date(baseYear, baseMonth - 1 - i, 26);
          let end = new Date(baseYear, baseMonth - i, 25);
          let label = `${start.getDate()} ${start.toLocaleDateString('id-ID', {month: 'short'})} - ${end.getDate()} ${end.toLocaleDateString('id-ID', {month: 'short', year: 'numeric'})}`;
          list.push({ start, end, label });
        }
        return list;
      }, []);

      const activePeriod = periods[selectedPeriodIndex];

      const isDateInPeriod = (dateStr) => {
        if (!dateStr) return false;
        const [y, m, d] = dateStr.split('-').map(Number);
        const recordDate = new Date(y, m - 1, d); // Menggunakan local timezone
        const s = new Date(activePeriod.start); s.setHours(0,0,0,0);
        const e = new Date(activePeriod.end); e.setHours(23,59,59,999);
        return recordDate >= s && recordDate <= e;
      };

      // Menggabungkan presensi masuk dan keluar per hari persis seperti kartu aktivitas periode
      const dailyActivities = useMemo(() => {
        const userPresensi = presensiList.filter(p => String(p.karyawan_id) === String(currentUser.id) && isDateInPeriod(String(p.tanggal || '').slice(0, 10)));
        const userPendingClaims = (Array.isArray(klaimList) ? klaimList : []).filter(c =>
          String(c.karyawan_id) === String(currentUser.id) &&
          ['pending', 'menunggu', ''].includes(String(c.status || 'Pending').toLowerCase()) &&
          isDateInPeriod(String(c.tanggal_klaim || '').slice(0, 10))
        );
        const uniqueDates = [...new Set([...userPresensi.map(p => String(p.tanggal || '').slice(0, 10)), ...userPendingClaims.map(c => String(c.tanggal_klaim || '').slice(0, 10))])].filter(Boolean).sort((a, b) => b.localeCompare(a));

        return uniqueDates.map(date => {
          const masukRec = userPresensi.find(p => String(p.tanggal || '').slice(0, 10) === date && p.tipe === 'masuk');
          const keluarRec = userPresensi.find(p => String(p.tanggal || '').slice(0, 10) === date && p.tipe === 'keluar');
          const pendingClaim = !masukRec && !keluarRec ? userPendingClaims.find(c => String(c.tanggal_klaim || '').slice(0, 10) === date) : null;
          const status = pendingClaim ? 'Pending' : (masukRec ? masukRec.status : (keluarRec ? keluarRec.status : 'Hadir'));
          const shift = pendingClaim?.shift_nama || masukRec?.shift_nama || keluarRec?.shift_nama || 'Reguler';
          const statusTanpaJam = pendingClaim || ['Libur', 'Sakit', 'Cuti', 'Izin'].includes(String(status || '').trim());
          const jamMasuk = statusTanpaJam ? '-' : (masukRec ? (masukRec.jam || '-') : '-');
          const jamPulang = statusTanpaJam ? '-' : (keluarRec ? (keluarRec.jam || '-') : '-');
          return {
            id: pendingClaim?.id || masukRec?.id || keluarRec?.id || date,
            tanggal: date, status, shift, jamMasuk, jamPulang,
            keterangan: pendingClaim ? 'Menunggu persetujuan HR' : (masukRec?.keterangan || keluarRec?.keterangan || '')
          };
        });
      }, [presensiList, klaimList, currentUser.id, activePeriod]);

      // Filter daftar lembur pada periode aktif
      const periodLemburList = useMemo(() => {
        return lemburList.filter(l => l.karyawan_id === currentUser.id && isDateInPeriod(l.tanggal));
      }, [lemburList, currentUser.id, activePeriod]);

      // Hitung jumlah riwayat per status dalam periode aktif
      const counts = useMemo(() => {
        return {
          'Semua': dailyActivities.length,
          'Hadir': dailyActivities.filter(p => p.status === 'Hadir').length,
          'Pending': dailyActivities.filter(p => p.status === 'Pending').length,
          'Terlambat': dailyActivities.filter(p => p.status === 'Terlambat').length,
          'Libur': dailyActivities.filter(p => p.status === 'Libur').length,
          'Izin': dailyActivities.filter(p => p.status === 'Izin').length,
          'Sakit': dailyActivities.filter(p => p.status === 'Sakit').length,
          'Cuti': dailyActivities.filter(p => p.status === 'Cuti').length,
          'Lembur': periodLemburList.length
        };
      }, [dailyActivities, periodLemburList]);

      const filteredActivities = useMemo(() => {
        if (filterStatus === 'Semua') return dailyActivities;
        return dailyActivities.filter(a => a.status === filterStatus);
      }, [dailyActivities, filterStatus]);

      return (
        <div className="p-5 space-y-4 pb-28">
          <div className="mb-2">
            <h2 className="font-display font-bold text-lg text-slate-800">Riwayat Presensi</h2>
            <div className="mt-1 flex items-center space-x-2">
              <i className="fa-regular fa-calendar text-slate-400 text-xs"></i>
              <select 
                value={selectedPeriodIndex} 
                onChange={(e) => setSelectedPeriodIndex(Number(e.target.value))}
                className="bg-white border border-slate-200 text-[10px] text-slate-600 font-bold rounded-lg px-2 py-1.5 outline-none shadow-sm cursor-pointer"
              >
                {periods.map((p, idx) => (
                  <option key={idx} value={idx}>{p.label}</option>
                ))}
              </select>
            </div>
          </div>
          
          <div className="flex flex-wrap gap-2 pb-2 text-[10px] font-bold">
            {['Semua', 'Hadir', 'Pending', 'Terlambat', 'Libur', 'Izin', 'Sakit', 'Cuti', 'Lembur'].map(f => (
              <button 
                key={f} 
                onClick={() => setFilterStatus(f)} 
                className={`px-3 py-1.5 rounded-full transition shadow-sm flex items-center gap-1.5 ${filterStatus === f ? 'bg-slate-900 text-white' : 'bg-white text-slate-500 border border-slate-100 hover:bg-slate-50'}`}
              >
                <span>{f}</span>
                <span className={`px-1.5 py-0.2 rounded-full text-[9px] font-mono-numbers font-bold ${filterStatus === f ? 'bg-white/20 text-white' : 'bg-slate-100 text-slate-600'}`}>
                  {counts[f] ?? 0}
                </span>
              </button>
            ))}
          </div>

          <div className="space-y-3">
            {filterStatus === 'Lembur' ? (
              periodLemburList.length > 0 ? (
                periodLemburList.map(r => (
                  <div 
                    key={r.id} 
                    className="bg-white/95 hover:bg-white p-3.5 rounded-2xl border border-slate-100 shadow-[0_4px_20px_rgb(0,0,0,0.03)] transition flex flex-col space-y-2.5"
                  >
                    {/* Baris Atas: Tanggal & Badge Status Lembur */}
                    <div className="flex items-center justify-between">
                      <span className="text-[11px] font-bold text-slate-800 font-mono-numbers flex items-center gap-1.5">
                        <i className="fa-regular fa-calendar text-[10px] text-slate-400"></i>
                        {formatIndoDate(r.tanggal)}
                      </span>
                      <span className={`px-2.5 py-0.5 rounded-full text-[9px] font-bold ${
                        r.status === 'Approved' ? 'bg-emerald-100 text-emerald-700 border border-emerald-200' :
                        r.status === 'Rejected' ? 'bg-rose-100 text-rose-700 border border-rose-200' :
                        'bg-amber-100 text-amber-700 border border-amber-200'
                      }`}>
                        {r.status === 'Approved' ? 'Disetujui' : r.status === 'Rejected' ? 'Ditolak' : 'Pending'}
                      </span>
                    </div>

                    {/* Baris Bawah: Keterangan Tugas & Durasi Lembur */}
                    <div className="flex items-center justify-between pt-1 border-t border-slate-100 text-[10px]">
                      <div className="text-slate-500 font-semibold truncate pr-2 flex items-center gap-1">
                        <i className="fa-solid fa-briefcase text-[9px] text-indigo-500"></i>
                        <span className="truncate max-w-[170px]">{r.keterangan || 'Pengajuan Lembur'}</span>
                      </div>
                      <span className="text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded-md border border-indigo-100 font-mono-numbers font-bold shrink-0">
                        <i className="fa-solid fa-business-time text-[8px] mr-1"></i>{r.durasi_jam} Jam
                      </span>
                    </div>
                    {act.keterangan && /disetujui HR/i.test(act.keterangan) && <div className="text-[10px] leading-relaxed text-emerald-700 bg-emerald-50 border border-emerald-100 rounded-xl px-3 py-2"><i className="fa-solid fa-circle-check mr-1.5"></i>{act.keterangan}</div>}
                  </div>
                ))
              ) : (
                <div className="text-center py-10 text-slate-400 bg-white/50 rounded-2xl border border-dashed border-slate-200">
                  <i className="fa-regular fa-calendar-xmark text-xl mb-1 text-slate-300"></i>
                  <p className="text-xs font-medium">Tidak ada catatan Lembur pada periode ini</p>
                </div>
              )
            ) : (
              filteredActivities.length > 0 ? (
                filteredActivities.map(act => (
                  <div 
                    key={act.id} 
                    className="bg-white/95 hover:bg-white p-3.5 rounded-2xl border border-slate-100 shadow-[0_4px_20px_rgb(0,0,0,0.03)] transition flex flex-col space-y-2.5"
                  >
                    {/* Baris Atas: Tanggal & Badge Status Absen */}
                    <div className="flex items-center justify-between">
                      <span className="text-[11px] font-bold text-slate-800 font-mono-numbers flex items-center gap-1.5">
                        <i className="fa-regular fa-calendar text-[10px] text-slate-400"></i>
                        {formatIndoDate(act.tanggal)}
                      </span>
                      <span className={`px-2.5 py-0.5 rounded-full text-[9px] font-bold ${
                        act.status === 'Hadir' ? 'bg-emerald-100 text-emerald-700 border border-emerald-200' :
                        act.status === 'Terlambat' ? 'bg-amber-100 text-amber-700 border border-amber-200' :
                        act.status === 'Libur' ? 'bg-sky-100 text-sky-700 border border-sky-200' :
                        act.status === 'Izin' ? 'bg-orange-100 text-orange-700 border border-orange-200' :
                        act.status === 'Sakit' ? 'bg-amber-100 text-amber-700 border border-amber-200' :
                        act.status === 'Cuti' ? 'bg-purple-100 text-purple-700 border border-purple-200' :
                        act.status === 'Pending' ? 'bg-amber-100 text-amber-800 border border-amber-200' :
                        'bg-slate-200 text-slate-700'
                      }`}>
                        {act.status}
                      </span>
                    </div>

                    {/* Baris Bawah: Shift Kerja, Jam Masuk, dan Jam Pulang */}
                    <div className="flex items-center justify-between pt-1 border-t border-slate-100 text-[10px]">
                      <div className="text-slate-500 font-semibold truncate pr-2 flex items-center gap-1">
                        <i className="fa-regular fa-clock text-[9px] text-violet-500"></i>
                        <span className="truncate max-w-[140px]">{act.shift}</span>
                      </div>
                      {['Libur', 'Sakit', 'Cuti', 'Izin'].includes(String(act.status || '').trim()) ? (
                        <span className="text-slate-600 bg-slate-100 px-2 py-1 rounded-md border border-slate-200 font-semibold">
                          Tidak ada jam presensi
                        </span>
                      ) : (
                        <div className="flex items-center space-x-2 shrink-0 font-mono-numbers font-bold">
                          <span className="text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded-md border border-emerald-100" title="Jam Masuk">
                            <i className="fa-solid fa-arrow-right-to-bracket text-[8px] mr-1"></i>{act.jamMasuk}
                          </span>
                          <span className="text-rose-600 bg-rose-50 px-2 py-0.5 rounded-md border border-rose-100" title="Jam Pulang">
                            <i className="fa-solid fa-arrow-right-from-bracket text-[8px] mr-1"></i>{act.jamPulang}
                          </span>
                        </div>
                      )}
                    </div>
                    {act.status === 'Pending' && <div className="text-[10px] leading-relaxed text-amber-800 bg-amber-50 border border-amber-100 rounded-xl px-3 py-2"><i className="fa-regular fa-clock mr-1.5"></i>Menunggu persetujuan HR</div>}
                    {act.keterangan && /disetujui HR/i.test(act.keterangan) && <div className="text-[10px] leading-relaxed text-emerald-700 bg-emerald-50 border border-emerald-100 rounded-xl px-3 py-2"><i className="fa-solid fa-circle-check mr-1.5"></i>{act.keterangan}</div>}
                  </div>
                ))
              ) : (
                <div className="text-center py-10 text-slate-400 bg-white/50 rounded-2xl border border-dashed border-slate-200">
                  <i className="fa-regular fa-calendar-xmark text-xl mb-1 text-slate-300"></i>
                  <p className="text-xs font-medium">Tidak ada catatan {filterStatus} pada periode ini</p>
                </div>
              )
            )}
          </div>
        </div>
      );
    }

    function ProfilKaryawanView({ currentUser, setCurrentUser, showToast, handleLogout }) {
      const [uploadingProfilePhoto, setUploadingProfilePhoto] = useState(false);
      const handleProfilePhotoChange = async (event) => {
        const file = event.target.files && event.target.files[0];
        if (!file) return;
        if (!file.type.startsWith('image/')) { showToast('Pilih file gambar yang valid', 'error'); return; }
        if (file.size > 8 * 1024 * 1024) { showToast('Ukuran foto maksimal 8 MB', 'error'); return; }
        setUploadingProfilePhoto(true);
        try {
          const dataUrl = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = () => reject(new Error('Gagal membaca file foto'));
            reader.readAsDataURL(file);
          });
          const uploadedUrl = await uploadToGoogleDriveViaGAS(dataUrl, `profil_${currentUser.id}_${Date.now()}.jpg`, 'profil');
          if (!uploadedUrl || uploadedUrl.startsWith('data:image')) {
            throw new Error('Foto belum berhasil diunggah ke Google Drive. Periksa deployment Google Apps Script dan izin aksesnya.');
          }
          const client = await getSupabaseClient();
          const { error } = await client.from('master_karyawan').update({ foto_profil: uploadedUrl }).eq('id', currentUser.id);
          if (error) throw error;
          const updatedAt = Date.now();
          const updatedUser = { ...currentUser, foto_profil: uploadedUrl, foto_profil_updated_at: updatedAt };
          setCurrentUser(updatedUser);
          try { localStorage.setItem('suryatokoku_user', JSON.stringify(updatedUser)); } catch (storageError) {
            console.warn('Tidak dapat memperbarui cache sesi lokal:', storageError);
          }
          showToast('Foto profil berhasil diperbarui');
        } catch (err) {
          console.error('Gagal menyimpan foto profil:', err);
          showToast(`Gagal menyimpan foto profil: ${err.message || 'kesalahan tidak diketahui'}`, 'error');
        } finally {
          setUploadingProfilePhoto(false);
          event.target.value = '';
        }
      };
      return (
        <div className="p-5 space-y-4 pb-28">
          <div className="bg-white/90 backdrop-blur-xl p-6 rounded-[2rem] border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] text-center">
            <div className="relative w-24 h-24 mx-auto mb-4">
              <div className="w-full h-full rounded-full bg-slate-50 border border-slate-100 overflow-hidden flex items-center justify-center">
                {currentUser.foto_profil ? <img key={`${currentUser.foto_profil}_${currentUser.foto_profil_updated_at || 'initial'}`} src={getProfileImageUrl(currentUser.foto_profil, currentUser.foto_profil_updated_at || 'initial')} onError={(e) => { if (e.currentTarget.dataset.retry !== '1') { e.currentTarget.dataset.retry = '1'; e.currentTarget.src = getProfileImageUrl(currentUser.foto_profil, Date.now()); } }} className="w-full h-full object-cover" /> : <span className="font-bold text-2xl text-slate-300">{currentUser.nama_lengkap[0]}</span>}
              </div>
              <label className="absolute bottom-0 right-0 w-8 h-8 rounded-full bg-slate-900 text-white flex items-center justify-center cursor-pointer shadow-md">
                <i className={`fa-solid ${uploadingProfilePhoto ? 'fa-spinner fa-spin' : 'fa-camera'} text-xs`}></i>
                <input type="file" accept="image/*" onChange={handleProfilePhotoChange} disabled={uploadingProfilePhoto} className="hidden" />
              </label>
            </div>
            <h3 className="font-display font-bold text-lg text-slate-800">{currentUser.nama_lengkap}</h3>
            <p className="text-[11px] font-medium text-slate-500 mt-1">{currentUser.email} &bull; {currentUser.no_wa}</p>
          </div>

          <div className="bg-white/90 backdrop-blur-xl rounded-[2rem] p-5 border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] space-y-2">
            <a href="#ganti-password" className="flex items-center justify-between p-3 rounded-2xl hover:bg-slate-50 transition">
              <div className="flex items-center space-x-3">
                <div className="w-9 h-9 rounded-full bg-slate-50 text-slate-600 flex items-center justify-center"><i className="fa-solid fa-key text-xs"></i></div>
                <span className="text-xs font-bold text-slate-700">Ganti Sandi</span>
              </div>
              <i className="fa-solid fa-chevron-right text-xs text-slate-300"></i>
            </a>
            <button onClick={() => requestPwaInstall(showToast)} className="w-full flex items-center justify-between p-3 rounded-2xl hover:bg-slate-50 transition">
              <div className="flex items-center space-x-3">
                <div className="w-9 h-9 rounded-full bg-slate-50 text-slate-600 flex items-center justify-center"><i className="fa-solid fa-download text-xs"></i></div>
                <span className="text-xs font-bold text-slate-700">Install App</span>
              </div>
              <i className="fa-solid fa-arrow-down text-xs text-slate-300"></i>
            </button>
            <button onClick={handleLogout} className="w-full flex items-center justify-between p-3 rounded-2xl hover:bg-rose-50 text-rose-500 transition mt-2">
              <div className="flex items-center space-x-3">
                <div className="w-9 h-9 rounded-full bg-rose-50 text-rose-500 flex items-center justify-center"><i className="fa-solid fa-arrow-right-from-bracket text-xs"></i></div>
                <span className="text-xs font-bold">Keluar Akun</span>
              </div>
            </button>
          </div>
        </div>
      );
    }

    function GantiPasswordView({ currentUser, setCurrentUser, setKaryawanList, showToast, navigateTo }) {
      const [oldPass, setOldPass] = useState('');
      const [newPass, setNewPass] = useState('');
      const [confirmPass, setConfirmPass] = useState('');
      const [isSaving, setIsSaving] = useState(false);

      const handleSave = async (e) => {
        e.preventDefault();
        if (isSaving) return;
        if (oldPass !== String(currentUser?.password ?? '')) {
          showToast('Sandi lama salah!', 'error');
          return;
        }
        if (newPass.length < 6) {
          showToast('Sandi baru minimal 6 karakter.', 'error');
          return;
        }
        if (newPass !== confirmPass) {
          showToast('Konfirmasi sandi tidak cocok!', 'error');
          return;
        }
        if (!currentUser?.id) {
          showToast('ID karyawan tidak ditemukan. Silakan login ulang.', 'error');
          return;
        }

        setIsSaving(true);
        try {
          const client = await getSupabaseClient();
          // Simpan ke database lebih dahulu. Jangan tampilkan notifikasi sukses
          // sebelum Supabase mengonfirmasi baris karyawan benar-benar diperbarui.
          const { data, error } = await client
            .from('master_karyawan')
            .update({ password: newPass })
            .eq('id', currentUser.id)
            .select('id')
            .maybeSingle();

          if (error) throw error;
          if (!data) {
            throw new Error('Tidak ada baris karyawan yang diperbarui. Periksa ID karyawan dan izin UPDATE/RLS tabel master_karyawan di Supabase.');
          }

          const updatedUser = { ...currentUser, password: newPass };
          setCurrentUser(updatedUser);
          if (typeof setKaryawanList === 'function') {
            setKaryawanList(prev => (Array.isArray(prev) ? prev : []).map(k => String(k.id) === String(currentUser.id) ? { ...k, password: newPass } : k));
          }
          try {
            localStorage.setItem('suryatokoku_user', JSON.stringify(updatedUser));
          } catch (storageError) {
            console.warn('Sesi lokal tidak dapat diperbarui:', storageError);
          }
          showToast('Kata sandi berhasil diperbarui di database.');
          navigateTo('#profil');
        } catch (error) {
          console.error('Gagal memperbarui kata sandi karyawan:', error);
          showToast(`Kata sandi gagal disimpan ke Supabase: ${error?.message || 'Periksa koneksi dan izin database.'}`, 'error');
        } finally {
          setIsSaving(false);
        }
      };

      return (
        <div className="p-5 pb-28">
          <div className="bg-white/90 backdrop-blur-xl p-6 rounded-[2rem] border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)]">
            <h3 className="font-display font-bold text-lg text-slate-800 mb-5">Ganti Sandi</h3>
            <form onSubmit={handleSave} className="space-y-4">
              <input type="password" autoComplete="current-password" placeholder="Sandi Lama" value={oldPass} onChange={e => setOldPass(e.target.value)} required className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 outline-none font-bold" />
              <input type="password" autoComplete="new-password" minLength={6} placeholder="Sandi Baru (minimal 6 karakter)" value={newPass} onChange={e => setNewPass(e.target.value)} required className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 outline-none font-bold" />
              <input type="password" autoComplete="new-password" minLength={6} placeholder="Konfirmasi Sandi" value={confirmPass} onChange={e => setConfirmPass(e.target.value)} required className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 outline-none font-bold" />
              <button type="submit" disabled={isSaving} className="w-full py-3.5 bg-slate-900 text-white rounded-full font-bold text-xs shadow-md mt-2 disabled:opacity-60 disabled:cursor-not-allowed">{isSaving ? 'Menyimpan...' : 'Simpan Sandi'}</button>
            </form>
          </div>
        </div>
      );
    }

    /* ==========================================
       ADMIN DASHBOARD WORKSPACE (Minimalist Light)
       ========================================== */
    // Helper global agar penanda kelengkapan data bisa dipakai di dashboard admin maupun Master Karyawan.
    function isKaryawanDataIncompleteGlobal(k) {
      const parseAlamat = (value) => {
        try { return typeof value === 'string' ? JSON.parse(value || '{}') : (value || {}); }
        catch (_) { return {}; }
      };
      const alamat = parseAlamat(k?.alamat);
      const requiredValues = [
        k?.nama_lengkap, k?.tgl_lahir, k?.username, k?.password,
        k?.no_wa, k?.email, k?.divisi, k?.jabatan,
        alamat?.provinsi, alamat?.kota, alamat?.kecamatan, alamat?.desa || alamat?.kelurahan
      ];
      return requiredValues.some(value => value === null || value === undefined || String(value).trim() === '');
    }

    const ADMIN_MENU_OPTIONS = [
      { id: 'hrd', label: 'HRD' },
      { id: 'admin', label: 'Master User' },
      { id: 'more_apps', label: 'Master More Apps' },
      { id: 'pengumuman', label: 'Pengumuman' }
    ];
    const ADMIN_MENU_ALL = ADMIN_MENU_OPTIONS.map(item => item.id);
    const encodeAdminMenuAccess = (ids) => 'MENU_ACCESS:' + ADMIN_MENU_ALL.filter(id => (ids || []).includes(id)).join(',');
    const decodeAdminMenuAccess = (role) => {
      if (typeof role === 'string' && role.startsWith('MENU_ACCESS:')) {
        const saved = role.slice('MENU_ACCESS:'.length).split(',').map(v => v.trim()).filter(Boolean);
        return ADMIN_MENU_ALL.filter(id => saved.includes(id));
      }
      // Akun lama dengan role teks biasa tetap mendapat akses penuh agar tidak kehilangan akses setelah update.
      return [...ADMIN_MENU_ALL];
    };
    const getAdminMenuLabels = (role) => ADMIN_MENU_OPTIONS.filter(item => decodeAdminMenuAccess(role).includes(item.id)).map(item => item.label);

    function AdminDashboard({
      karyawanList, setKaryawanList, presensiList, setPresensiList, lokasiList, setLokasiList,
      divisiList, setDivisiList, jabatanList, setJabatanList, shiftList, setShiftList, pengumumanList, setPengumumanList,
      adminList, setAdminList,
      lemburList, setLemburList, klaimList, setKlaimList, toleransiMenit, setToleransiMenit, pendaftaranKaryawanAktif, setPendaftaranKaryawanAktif, moreAppsAccess = {}, setMoreAppsAccess, moreAppsCatalog = DEFAULT_MORE_APPS, setMoreAppsCatalog, showToast, handleLogout, currentUser
    }) {
      const [activeModule, setActiveModule] = useState('harian'); 
      const [sidebarPinned, setSidebarPinned] = useState(true);
      const [hrdExpanded, setHrdExpanded] = useState(true);
      const hrdModuleIds = ['harian', 'periode', 'shift', 'divisi', 'jabatan', 'lokasi', 'karyawan', 'libur_cuti', 'lembur', 'klaim_kehadiran'];
      const activeAdminRecord = adminList.find(a => String(a.id) === String(currentUser?.id)) || currentUser || {};
      const allowedAdminMenus = decodeAdminMenuAccess(activeAdminRecord.role);
      const canAccessHrd = allowedAdminMenus.includes('hrd');
      useEffect(() => {
        const activeTopLevel = hrdModuleIds.includes(activeModule) ? 'hrd' : activeModule;
        if (!allowedAdminMenus.includes(activeTopLevel)) {
          const fallback = allowedAdminMenus.includes('hrd') ? 'harian' :
            allowedAdminMenus.includes('admin') ? 'admin' :
            allowedAdminMenus.includes('more_apps') ? 'more_apps' :
            allowedAdminMenus.includes('pengumuman') ? 'pengumuman' : 'no_access';
          setActiveModule(fallback);
        }
      }, [activeModule, activeAdminRecord.role]);
      const [selectedDate, setSelectedDate] = useState(getLocalDateKey());
      const [dateDraftHarian, setDateDraftHarian] = useState(formatDateDMY(getLocalDateKey()));
      // Sinkronkan tanggal yang terlihat dengan tanggal filter yang aktif,
      // termasuk saat pengguna menekan panah kiri/kanan atau memilih kalender.
      useEffect(() => {
        setDateDraftHarian(formatDateDMY(selectedDate));
      }, [selectedDate]);
      const [filterStatusHarian, setFilterStatusHarian] = useState('Semua');
      const kalenderHarianRef = useRef(null);
      const [previewPhotoUrl, setPreviewPhotoUrl] = useState(null);
      const [statusEvidenceModal, setStatusEvidenceModal] = useState(null);
      const [editingPresensiModal, setEditingPresensiModal] = useState(null);
      const [deleteConfirmState, setDeleteConfirmState] = useState(null);
      const [isDeleting, setIsDeleting] = useState(false);
      const [showExportHarianModal, setShowExportHarianModal] = useState(false);
      const [exportTanggalMulai, setExportTanggalMulai] = useState(getLocalDateKey());
      const [exportTanggalSelesai, setExportTanggalSelesai] = useState(getLocalDateKey());
      const [exportTanggalMulaiDraft, setExportTanggalMulaiDraft] = useState(formatDateDMY(getLocalDateKey()));
      const [exportTanggalSelesaiDraft, setExportTanggalSelesaiDraft] = useState(formatDateDMY(getLocalDateKey()));
      const [isExportingHarian, setIsExportingHarian] = useState(false);

      const exportToExcel = async (data, fileName) => {
        try {
          const XLSX = await ensureXLSXLoaded();
          const ws = XLSX.utils.json_to_sheet(data);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, "Data");
          XLSX.writeFile(wb, `${fileName}_${Date.now()}.xlsx`);
          showToast('Excel diunduh!');
        } catch (error) {
          console.error('Gagal mengekspor Excel:', error);
          showToast(error.message || 'Gagal membuat file Excel.', 'error');
        }
      };

      const handleExportRekapanHarian = async () => {
        if (!exportTanggalMulai || !exportTanggalSelesai) {
          showToast('Pilih tanggal mulai dan tanggal selesai.', 'error');
          return;
        }
        if (exportTanggalMulai > exportTanggalSelesai) {
          showToast('Tanggal mulai tidak boleh melewati tanggal selesai.', 'error');
          return;
        }
        setIsExportingHarian(true);
        try {
          const XLSX = await ensureXLSXLoaded();
          const dateKeys = [];
          const cursor = new Date(`${exportTanggalMulai}T12:00:00`);
          const end = new Date(`${exportTanggalSelesai}T12:00:00`);
          while (cursor <= end) {
            const yyyy = cursor.getFullYear();
            const mm = String(cursor.getMonth() + 1).padStart(2, '0');
            const dd = String(cursor.getDate()).padStart(2, '0');
            dateKeys.push(`${yyyy}-${mm}-${dd}`);
            cursor.setDate(cursor.getDate() + 1);
          }

          const rows = [];
          const sortedKaryawan = [...karyawanList].sort((a, b) =>
            String(a.nama_lengkap || '').localeCompare(String(b.nama_lengkap || ''), 'id')
          );
          const recordsByEmployeeDate = new Map();
          presensiList.forEach(item => {
            const key = `${String(item.karyawan_id)}|${item.tanggal}`;
            if (!recordsByEmployeeDate.has(key)) recordsByEmployeeDate.set(key, []);
            recordsByEmployeeDate.get(key).push(item);
          });
          dateKeys.forEach(tanggal => {
            sortedKaryawan.forEach(k => {
                const records = recordsByEmployeeDate.get(`${String(k.id)}|${tanggal}`) || [];
                const masuk = records.find(item => item.tipe === 'masuk');
                const pulang = records.find(item => item.tipe === 'keluar');
                const special = records.find(item => ['Sakit', 'Izin', 'Cuti', 'Libur'].includes(String(item.status || '').trim()));
                const noTimes = !!special;
                rows.push({
                  'Tanggal': formatDateDMY(tanggal),
                  'Nama Karyawan': k.nama_lengkap || '',
                  'Nomor Induk Karyawan': k.nik || '',
                  'Divisi': k.divisi || '',
                  'Jabatan': k.jabatan || '',
                  'Shift': (masuk?.shift_nama || pulang?.shift_nama || (k.shifts && k.shifts[0]) || ''),
                  'Jam Masuk': noTimes ? '-' : (masuk?.jam || '-'),
                  'Jam Pulang': noTimes ? '-' : (pulang?.jam || '-'),
                  'Status': special?.status || masuk?.status || pulang?.status || 'Belum Masuk',
                  'Keterangan': special?.keterangan || masuk?.keterangan || pulang?.keterangan || '',
                  'Foto Bukti / Masuk': noTimes ? '-' : (masuk?.foto_url || ''),
                  'Foto Pulang': noTimes ? '-' : (pulang?.foto_url || '')
                });
              });
          });
          const ws = XLSX.utils.json_to_sheet(rows);
          ws['!cols'] = [
            { wch: 13 }, { wch: 28 }, { wch: 20 }, { wch: 18 }, { wch: 20 },
            { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 36 },
            { wch: 48 }, { wch: 48 }
          ];
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, 'Rekapan Harian');
          XLSX.writeFile(wb, `Rekapan_Presensi_Harian_${exportTanggalMulai}_sampai_${exportTanggalSelesai}.xlsx`);
          setShowExportHarianModal(false);
          showToast(`Excel berhasil diekspor (${rows.length} baris).`);
        } catch (error) {
          console.error('Gagal export rekapan harian:', error);
          showToast('Gagal membuat file Excel. Silakan coba lagi.', 'error');
        } finally {
          setIsExportingHarian(false);
        }
      };

      const requestHapusPresensi = (karyawanId, karyawanNama) => {
        setDeleteConfirmState({ karyawanId, karyawanNama, tanggal: selectedDate });
      };

      const executeHapusPresensi = async () => {
        if (!deleteConfirmState) return;
        const { karyawanId, karyawanNama, tanggal } = deleteConfirmState;
        setIsDeleting(true);

        try {
          // 1. Hapus dari tabel presensi di Supabase
          if (supabase) {
            const { error } = await supabase
              .from('presensi')
              .delete()
              .eq('karyawan_id', karyawanId)
              .eq('tanggal', tanggal);

            if (error) {
              console.error("Gagal menghapus presensi di Supabase:", error);
              showToast('Gagal menghapus dari Supabase: ' + error.message, 'error');
              setIsDeleting(false);
              setDeleteConfirmState(null);
              return;
            }
          }

          // 2. Perbarui state lokal setelah sukses
          setPresensiList(prev => prev.filter(p => !(p.karyawan_id === karyawanId && p.tanggal === tanggal)));
          showToast(`Presensi ${karyawanNama} tanggal ${tanggal} berhasil dihapus dari sistem!`);
        } catch (err) {
          console.error("Error eksekusi hapus presensi:", err);
          showToast('Terjadi kesalahan saat menghapus data.', 'error');
        } finally {
          setIsDeleting(false);
          setDeleteConfirmState(null);
        }
      };

      const handleSimpanEditPresensi = async (e) => {
        e.preventDefault();
        const { karyawan, masuk, pulang, formMasuk, formPulang } = editingPresensiModal;

        let updatedList = [...presensiList];
        const newRecordsForSupabase = [];

        // Update / create presensi Masuk
        if (formMasuk.status !== 'Belum' || formMasuk.jam) {
          const masukRecord = {
            id: masuk ? masuk.id : 'prs-' + Date.now() + '-m',
            karyawan_id: karyawan.id,
            karyawan_nama: karyawan.nama_lengkap,
            divisi: karyawan.divisi,
            tipe: 'masuk',
            status: formMasuk.status || 'Hadir',
            shift_nama: formMasuk.shift_nama || shiftList[0]?.nama_shift || 'Shift Pagi',
            lokasi_nama: masuk?.lokasi_nama || lokasiList[0]?.nama_lokasi || 'Kantor',
            latitude: masuk?.latitude || 0,
            longitude: masuk?.longitude || 0,
            foto_url: masuk?.foto_url || '',
            keterangan: formMasuk.keterangan || 'Disesuaikan oleh Admin',
            tanggal: selectedDate,
            jam: formMasuk.jam || '08:00:00',
            created_at: masuk?.created_at || new Date().toISOString()
          };
          updatedList = updatedList.filter(p => !(p.karyawan_id === karyawan.id && p.tanggal === selectedDate && p.tipe === 'masuk'));
          updatedList.unshift(masukRecord);
          newRecordsForSupabase.push(masukRecord);
        } else if (masuk && formMasuk.status === 'Belum') {
          updatedList = updatedList.filter(p => p.id !== masuk.id);
          if (supabase) {
            try { await supabase.from('presensi').delete().eq('id', masuk.id); } catch(err){}
          }
        }

        // Update / create presensi Keluar/Pulang
        if (formPulang.status !== 'Belum' || formPulang.jam) {
          const pulangRecord = {
            id: pulang ? pulang.id : 'prs-' + Date.now() + '-p',
            karyawan_id: karyawan.id,
            karyawan_nama: karyawan.nama_lengkap,
            divisi: karyawan.divisi,
            tipe: 'keluar',
            status: formPulang.status || 'Pulang',
            shift_nama: formMasuk.shift_nama || shiftList[0]?.nama_shift || 'Shift Pagi',
            lokasi_nama: pulang?.lokasi_nama || lokasiList[0]?.nama_lokasi || 'Kantor',
            latitude: pulang?.latitude || 0,
            longitude: pulang?.longitude || 0,
            foto_url: pulang?.foto_url || '',
            keterangan: formPulang.keterangan || 'Disesuaikan oleh Admin',
            tanggal: selectedDate,
            jam: formPulang.jam || '17:00:00',
            created_at: pulang?.created_at || new Date().toISOString()
          };
          updatedList = updatedList.filter(p => !(p.karyawan_id === karyawan.id && p.tanggal === selectedDate && p.tipe === 'keluar'));
          updatedList.unshift(pulangRecord);
          newRecordsForSupabase.push(pulangRecord);
        } else if (pulang && formPulang.status === 'Belum') {
          updatedList = updatedList.filter(p => p.id !== pulang.id);
          if (supabase) {
            try { await supabase.from('presensi').delete().eq('id', pulang.id); } catch(err){}
          }
        }

        // Sinkronisasi record baru/edit ke Supabase
        if (supabase && newRecordsForSupabase.length > 0) {
          try {
            await supabase.from('presensi').upsert(newRecordsForSupabase);
          } catch(err) {
            console.warn("Gagal menyimpan perubahan presensi ke Supabase:", err);
          }
        }

        setPresensiList(updatedList);
        setEditingPresensiModal(null);
        showToast('Data presensi berhasil diperbarui!');
      };

      return (
        <div className="flex min-h-[calc(100vh-50px)] bg-slate-50">
          {/* Sidebar Admin Modern Minimalist */}
          <aside className={`bg-white text-slate-600 flex flex-col justify-between transition-all duration-300 border-r border-slate-200 z-30 ${sidebarPinned ? 'w-64' : 'w-20'}`}>
            <div>
              <div className="p-4 flex items-center justify-between border-b border-slate-100">
                {sidebarPinned && (
                  <div className="flex items-center space-x-2">
                    <img src="Logo SuryaKuCan.png" alt="Logo SuryaKuCan" className="w-8 h-8 object-contain rounded-lg drop-shadow-xs" />
                    <div><h2 className="font-bold text-sm text-slate-800">Admin Panel</h2></div>
                  </div>
                )}
                <button onClick={() => setSidebarPinned(!sidebarPinned)} className="w-8 h-8 rounded-lg bg-slate-50 hover:bg-slate-100 flex items-center justify-center mx-auto text-slate-400"><i className={`fa-solid ${sidebarPinned ? 'fa-bars-staggered' : 'fa-bars'} text-xs`}></i></button>
              </div>
              <nav className="p-3 space-y-1.5 text-xs font-bold">
                {canAccessHrd && <>
                <button
                  onClick={() => setHrdExpanded(prev => !prev)}
                  aria-expanded={hrdExpanded}
                  className={`w-full flex items-center gap-3 px-3 py-3 rounded-xl transition ${hrdModuleIds.includes(activeModule) ? 'bg-violet-50 text-violet-700 shadow-sm border border-violet-100/50' : 'hover:bg-slate-50 text-slate-500'}`}
                  title="HRD"
                >
                  <i className={`fa-solid fa-people-group text-sm w-5 text-center ${hrdModuleIds.includes(activeModule) ? 'text-violet-500' : ''}`}></i>
                  {sidebarPinned && <><span className="truncate flex-1 text-left">HRD</span><i className={`fa-solid fa-chevron-${hrdExpanded ? 'down' : 'right'} text-[10px]`}></i></>}
                </button>
                {hrdExpanded && (
                  <div className="ml-3 pl-2 border-l border-slate-200 space-y-1">
                    {[
                      { id: 'harian', label: 'Rekapan Harian', icon: 'fa-calendar-day' },
                      { id: 'periode', label: 'Rekapan Periode', icon: 'fa-calendar-week' },
                      { id: 'shift', label: 'Master Shift', icon: 'fa-clock' },
                      { id: 'divisi', label: 'Master Divisi', icon: 'fa-layer-group' },
                      { id: 'jabatan', label: 'Master Jabatan', icon: 'fa-id-badge' },
                      { id: 'lokasi', label: 'Master Lokasi', icon: 'fa-map-pin' },
                      { id: 'karyawan', label: 'Master Karyawan', icon: 'fa-users' },
                      { id: 'libur_cuti', label: 'Master Libur & Cuti', icon: 'fa-umbrella-beach' },
                      { id: 'lembur', label: 'Pengajuan Lembur', icon: 'fa-business-time' },
                      { id: 'klaim_kehadiran', label: 'Klaim Kehadiran', icon: 'fa-file-circle-check' }
                    ].map(item => (
                      <button key={item.id} onClick={() => setActiveModule(item.id)} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition ${activeModule === item.id ? 'bg-violet-50 text-violet-700 shadow-sm border border-violet-100/50' : 'hover:bg-slate-50 text-slate-500'}`}>
                        <i className={`fa-solid ${item.icon} text-xs w-5 text-center ${activeModule === item.id ? 'text-violet-500' : ''}`}></i>
                        {sidebarPinned && <span className="truncate flex-1 text-left">{item.label}</span>}
                      </button>
                    ))}
                  </div>
                )}
                </>}
                {ADMIN_MENU_OPTIONS.filter(item => item.id !== 'hrd' && allowedAdminMenus.includes(item.id)).map(item => ({ ...item, icon: item.id === 'admin' ? 'fa-user-shield' : item.id === 'more_apps' ? 'fa-layer-group' : 'fa-bullhorn' })).map(item => (
                  <button key={item.id} onClick={() => setActiveModule(item.id)} className={`w-full flex items-center gap-3 px-3 py-3 rounded-xl transition ${activeModule === item.id ? 'bg-violet-50 text-violet-700 shadow-sm border border-violet-100/50' : 'hover:bg-slate-50 text-slate-500'}`}>
                    <i className={`fa-solid ${item.icon} text-sm w-5 text-center ${activeModule === item.id ? 'text-violet-500' : ''}`}></i>
                    {sidebarPinned && <span className="truncate flex-1 text-left">{item.label}</span>}
                  </button>
                ))}
              </nav>
            </div>
            
            <div className="p-4 border-t border-slate-100">
              <button 
                onClick={handleLogout}
                className="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl hover:bg-rose-50 text-rose-500 transition text-xs font-bold"
                title="Keluar dari sesi Administrator"
              >
                <i className="fa-solid fa-arrow-right-from-bracket text-sm w-5 text-center"></i>
                {sidebarPinned && <span className="truncate">Keluar Sesi</span>}
              </button>
            </div>
          </aside>

          {/* Admin Workspace Content */}
          <main className="flex-1 p-6 lg:p-8 overflow-x-auto relative">
            
            {activeModule === 'karyawan' && (
              <MasterKaryawanView 
                karyawanList={karyawanList} 
                setKaryawanList={setKaryawanList} 
                divisiList={divisiList} 
                jabatanList={jabatanList} 
                lokasiList={lokasiList}
                pendaftaranKaryawanAktif={pendaftaranKaryawanAktif}
                setPendaftaranKaryawanAktif={setPendaftaranKaryawanAktif}
                showToast={showToast} 
                setPreviewPhotoUrl={setPreviewPhotoUrl}
              />
            )}
            
            {activeModule === 'more_apps' && (
              <MasterMoreAppsView karyawanList={karyawanList} moreAppsAccess={moreAppsAccess} setMoreAppsAccess={setMoreAppsAccess} moreAppsCatalog={moreAppsCatalog} setMoreAppsCatalog={setMoreAppsCatalog} showToast={showToast} />
            )}

            {activeModule === 'no_access' && (
              <div className="bg-white rounded-3xl border border-slate-100 p-8 text-center shadow-sm"><div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-amber-50 text-amber-600"><i className="fa-solid fa-lock text-xl"></i></div><h2 className="text-lg font-extrabold text-slate-800">Belum Ada Menu yang Diizinkan</h2><p className="mt-2 text-sm text-slate-500">Hubungi administrator untuk mengatur hak akses menu akun ini.</p></div>
            )}
            {activeModule === 'harian' && (
              <div className="space-y-6">
                <div className="flex flex-wrap justify-between items-center gap-4 bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
                  <div>
                    <h2 className="font-display font-bold text-xl text-slate-800">Rekapan Presensi Harian</h2>
                    <p className="text-xs text-slate-400 font-medium">Tinjauan shift, jam, foto selfie, dan status kehadiran karyawan harian</p>
                  </div>
                  <div className="flex items-center gap-2 flex-wrap">
                  <div className="flex items-center space-x-2 bg-slate-50 border border-slate-200 p-1 rounded-2xl shadow-sm">
                    {/* Tombol Panah Kiri (Hari Sebelumnya) */}
                    <button 
                      onClick={() => {
                        setSelectedDate(shiftDateKey(selectedDate, -1));
                      }}
                      className="w-9 h-9 rounded-xl bg-white hover:bg-slate-100 text-slate-600 flex items-center justify-center transition active:scale-95 shadow-sm"
                      title="Hari Sebelumnya"
                    >
                      <i className="fa-solid fa-chevron-left text-xs"></i>
                    </button>

                    {/* Tanggal ditampilkan konsisten dalam format Indonesia dd/mm/yyyy */}
                    <input
                      type="text"
                      inputMode="numeric"
                      value={dateDraftHarian}
                      title="Tanggal (dd/mm/yyyy)"
                      aria-label="Filter tanggal, format dd/mm/yyyy"
                      placeholder="dd/mm/yyyy"
                      onChange={(e) => setDateDraftHarian(e.target.value)}
                      onBlur={() => {
                        const match = dateDraftHarian.trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                        if (match) {
                          const day = Number(match[1]), month = Number(match[2]), year = Number(match[3]);
                          const candidate = new Date(year, month - 1, day);
                          if (candidate.getFullYear() === year && candidate.getMonth() === month - 1 && candidate.getDate() === day) {
                            setSelectedDate(`${year}-${match[2]}-${match[1]}`);
                          } else setDateDraftHarian(formatDateDMY(selectedDate));
                        } else setDateDraftHarian(formatDateDMY(selectedDate));
                      }}
                      className="w-[100px] bg-transparent border-0 text-xs text-center font-bold rounded-xl px-1 py-1.5 outline-none font-mono-numbers text-slate-700"
                    />
                    <input
                      ref={kalenderHarianRef}
                      type="date"
                      value={selectedDate}
                      onChange={(e) => e.target.value && setSelectedDate(e.target.value)}
                      className="sr-only"
                      tabIndex={-1}
                      aria-hidden="true"
                    />
                    <button
                      type="button"
                      onClick={() => { if (kalenderHarianRef.current) { if (kalenderHarianRef.current.showPicker) kalenderHarianRef.current.showPicker(); else kalenderHarianRef.current.click(); } }}
                      className="w-8 h-8 rounded-lg text-violet-600 hover:bg-violet-50 flex items-center justify-center"
                      title="Pilih tanggal dari kalender"
                    ><i className="fa-regular fa-calendar text-sm"></i></button>

                    {/* Tombol Panah Kanan (Hari Berikutnya) */}
                    <button 
                      onClick={() => {
                        setSelectedDate(shiftDateKey(selectedDate, 1));
                      }}
                      className="w-9 h-9 rounded-xl bg-white hover:bg-slate-100 text-slate-600 flex items-center justify-center transition active:scale-95 shadow-sm"
                      title="Hari Berikutnya"
                    >
                      <i className="fa-solid fa-chevron-right text-xs"></i>
                    </button>
                  </div>
                  <button
                    type="button"
                    onClick={() => { setExportTanggalMulai(selectedDate); setExportTanggalSelesai(selectedDate); setExportTanggalMulaiDraft(formatDateDMY(selectedDate)); setExportTanggalSelesaiDraft(formatDateDMY(selectedDate)); setShowExportHarianModal(true); }}
                    className="inline-flex items-center gap-2 px-4 py-3 rounded-2xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold shadow-sm transition active:scale-95"
                  >
                    <i className="fa-solid fa-file-excel"></i> Export Excel
                  </button>
                  </div>
                </div>
                {(() => {
                  const statusFor = (k) => {
                    const att = presensiList.find(item => item.karyawan_id === k.id && item.tanggal === selectedDate && item.tipe === 'masuk');
                    if (!att) return 'Belum masuk';
                    if (att.status === 'Hadir' || att.status === 'Terlambat' || att.status === 'Masuk') return 'Masuk';
                    const normalized = String(att.status || '').toLowerCase();
                    if (normalized === 'belum absen' || normalized === 'belum masuk' || normalized === 'belum') return 'Belum masuk';
                    if (normalized === 'libur') return 'Libur';
                    if (normalized === 'izin') return 'Izin';
                    if (normalized === 'sakit') return 'Sakit';
                    if (normalized === 'cuti') return 'Cuti';
                    return 'Belum masuk';
                  };
                  const statusFilters = ['Masuk', 'Belum masuk', 'Libur', 'Izin', 'Sakit', 'Cuti'];
                  const statusCounts = statusFilters.reduce((acc, status) => {
                    acc[status] = karyawanList.filter(k => statusFor(k) === status).length;
                    return acc;
                  }, {});
                  const visibleEmployees = [...karyawanList]
                    .filter(k => filterStatusHarian === 'Semua' || statusFor(k) === filterStatusHarian)
                    .sort((a, b) => a.nama_lengkap.localeCompare(b.nama_lengkap));
                  return <>
                    <div className="grid grid-cols-2 sm:grid-cols-3 xl:grid-cols-6 gap-3">
                      {statusFilters.map(status => {
                        const active = filterStatusHarian === status;
                        const palette = {
                          'Masuk': { active: 'bg-emerald-600 border-emerald-600', hover: 'hover:border-emerald-200', number: 'text-emerald-600' },
                          'Belum masuk': { active: 'bg-slate-800 border-slate-800', hover: 'hover:border-slate-300', number: 'text-slate-700' },
                          'Libur': { active: 'bg-sky-600 border-sky-600', hover: 'hover:border-sky-200', number: 'text-sky-600' },
                          'Izin': { active: 'bg-orange-500 border-orange-500', hover: 'hover:border-orange-200', number: 'text-orange-600' },
                          'Sakit': { active: 'bg-amber-500 border-amber-500', hover: 'hover:border-amber-200', number: 'text-amber-600' },
                          'Cuti': { active: 'bg-purple-600 border-purple-600', hover: 'hover:border-purple-200', number: 'text-purple-600' }
                        }[status];
                        return (
                          <button
                            key={status}
                            type="button"
                            onClick={() => setFilterStatusHarian(current => current === status ? 'Semua' : status)}
                            aria-pressed={active}
                            className={`text-left p-4 rounded-3xl border transition shadow-sm ${active ? `${palette.active} text-white shadow-md` : `bg-white text-slate-700 border-slate-100 ${palette.hover}`}`}
                          >
                            <span className="text-[10px] font-bold uppercase tracking-wider block opacity-75">{status}</span>
                            <span className={`font-mono-numbers text-2xl font-extrabold mt-1 block ${active ? 'text-white' : palette.number}`}>{statusCounts[status]}</span>
                          </button>
                        );
                      })}
                    </div>
                    {filterStatusHarian !== 'Semua' && (
                      <div className="flex justify-end -mt-2">
                        <button type="button" onClick={() => setFilterStatusHarian('Semua')} className="px-3 py-2 rounded-full text-xs font-bold text-violet-600 hover:bg-violet-50 transition">
                          <i className="fa-solid fa-rotate-left mr-1.5"></i>Reset filter
                        </button>
                      </div>
                    )}
                    <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
                      <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                          <th className="p-4 text-center">No</th>
                          <th className="p-4">Karyawan</th>
                          <th className="p-4">Shift Kerja</th>
                          <th className="p-4">Masuk</th>
                          <th className="p-4">Pulang</th>
                          <th className="p-4 text-center">Foto Masuk</th>
                          <th className="p-4 text-center">Foto Pulang</th>
                          <th className="p-4">Status</th>
                          <th className="p-4 text-center">Aksi</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-50">
                        {visibleEmployees
                          .map((k, i) => {
                            const m = presensiList.find(p => p.karyawan_id === k.id && p.tanggal === selectedDate && p.tipe === 'masuk');
                            const p = presensiList.find(p => p.karyawan_id === k.id && p.tanggal === selectedDate && p.tipe === 'keluar');
                            const shiftTerpakai = m?.shift_nama || p?.shift_nama || (k.shifts && k.shifts[0]) || 'Reguler';
                            const statusRecord = [m, p].find(item => item && ['Sakit', 'Izin', 'Cuti', 'Libur'].includes(String(item.status || '').trim())) || m || p;
                            const isStatusTanpaJam = !!statusRecord && ['Sakit', 'Izin', 'Cuti', 'Libur'].includes(String(statusRecord.status || '').trim());
                            const needsEvidenceButton = !!statusRecord && ['Sakit', 'Izin', 'Cuti'].includes(String(statusRecord.status || '').trim());
                            
                            return (
                              <tr key={k.id} className="transition hover:bg-slate-50/50">

                                <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{i + 1}</td>
                                <td className="p-4">
                                  <div className="font-bold text-slate-800">{k.nama_lengkap || 'Nama belum diisi'}</div>
                                  <div className="text-[10px] text-slate-400 font-medium">{k.divisi} {k.jabatan ? `• ${k.jabatan}` : ''}</div>
                                </td>
                                <td className="p-4">
                                  <span className="px-2.5 py-1 rounded-full bg-violet-50 text-violet-700 text-[10px] font-bold inline-flex items-center gap-1">
                                    <i className="fa-regular fa-clock text-[9px]"></i>
                                    {shiftTerpakai}
                                  </span>
                                </td>
                                <td className="p-4 font-mono-numbers font-semibold text-slate-700">{!isStatusTanpaJam && m ? (m.jam || '-') : '-'}</td>
                                <td className="p-4 font-mono-numbers font-semibold text-slate-700">{!isStatusTanpaJam && p ? (p.jam || '-') : '-'}</td>
                                
                                {/* Foto Selfie Masuk */}
                                <td className="p-4 text-center">
                                  {!isStatusTanpaJam && m?.foto_url ? (
                                    <button 
                                      onClick={() => setPreviewPhotoUrl({ url: getCleanImageUrl(m.foto_url), label: `Foto Masuk - ${k.nama_lengkap || 'Nama belum diisi'} (${m.jam})` })}
                                      className="relative group inline-block"
                                      title="Klik untuk perbesar"
                                    >
                                      <img 
                                        src={getCleanImageUrl(m.foto_url)} 
                                        alt="Foto Masuk" 
                                        className="w-10 h-10 rounded-xl object-cover border border-emerald-200 shadow-sm group-hover:scale-105 group-hover:border-emerald-400 transition" 
                                      />
                                      <span className="absolute -bottom-1 -right-1 w-4 h-4 rounded-full bg-emerald-500 text-white text-[8px] flex items-center justify-center shadow">
                                        <i className="fa-solid fa-magnifying-glass"></i>
                                      </span>
                                    </button>
                                  ) : (
                                    <span className="text-[10px] text-slate-300 font-medium">-</span>
                                  )}
                                </td>

                                {/* Foto Selfie Pulang */}
                                <td className="p-4 text-center">
                                  {!isStatusTanpaJam && p?.foto_url ? (
                                    <button 
                                      onClick={() => setPreviewPhotoUrl({ url: getCleanImageUrl(p.foto_url), label: `Foto Pulang - ${k.nama_lengkap || 'Nama belum diisi'} (${p.jam})` })}
                                      className="relative group inline-block"
                                      title="Klik untuk perbesar"
                                    >
                                      <img 
                                        src={getCleanImageUrl(p.foto_url)} 
                                        alt="Foto Pulang" 
                                        className="w-10 h-10 rounded-xl object-cover border border-rose-200 shadow-sm group-hover:scale-105 group-hover:border-rose-400 transition" 
                                      />
                                      <span className="absolute -bottom-1 -right-1 w-4 h-4 rounded-full bg-rose-500 text-white text-[8px] flex items-center justify-center shadow">
                                        <i className="fa-solid fa-magnifying-glass"></i>
                                      </span>
                                    </button>
                                  ) : (
                                    <span className="text-[10px] text-slate-300 font-medium">-</span>
                                  )}
                                </td>

                                {/* Status */}
                                <td className="p-4">
                                  <span className={`px-2.5 py-1 rounded-full text-[10px] font-bold ${
                                    m?.status === 'Hadir' ? 'bg-emerald-50 text-emerald-600' :
                                    m?.status === 'Terlambat' ? 'bg-amber-50 text-amber-600' :
                                    m?.status === 'Libur' ? 'bg-sky-50 text-sky-600' :
                                    m?.status === 'Izin' ? 'bg-orange-50 text-orange-600' :
                                    m?.status === 'Sakit' ? 'bg-amber-50 text-amber-600' :
                                    m?.status === 'Cuti' ? 'bg-purple-50 text-purple-600' :
                                    'bg-slate-100 text-slate-400'
                                  }`}>
                                    {m ? m.status : 'Belum Masuk'}
                                  </span>
                                </td>

                                {/* Ikon Edit dan Hapus */}
                                <td className="p-4 text-center">
                                  <div className="flex items-center justify-center space-x-1">
                                    {needsEvidenceButton && (
                                      <button
                                        type="button"
                                        onClick={() => setStatusEvidenceModal({
                                          karyawan: k,
                                          tanggal: selectedDate,
                                          status: statusRecord.status,
                                          foto_url: statusRecord.foto_url || '',
                                          keterangan: statusRecord.keterangan || ''
                                        })}
                                        className="px-3 h-8 rounded-full bg-violet-50 text-violet-700 hover:bg-violet-100 active:scale-95 transition text-[10px] font-bold whitespace-nowrap"
                                        title="Lihat foto bukti dan keterangan"
                                      >
                                        <i className="fa-solid fa-file-circle-check mr-1"></i>Lihat Bukti
                                      </button>
                                    )}
                                    <button 
                                      onClick={() => setEditingPresensiModal({
                                        karyawan: k,
                                        masuk: m,
                                        pulang: p,
                                        formMasuk: {
                                          status: m?.status || 'Hadir',
                                          jam: m?.jam || '08:00:00',
                                          shift_nama: m?.shift_nama || shiftTerpakai,
                                          keterangan: m?.keterangan || ''
                                        },
                                        formPulang: {
                                          status: p?.status || (p ? 'Pulang' : 'Belum'),
                                          jam: p?.jam || '17:00:00',
                                          keterangan: p?.keterangan || ''
                                        }
                                      })}
                                      className="w-8 h-8 rounded-full bg-slate-50 text-indigo-500 hover:bg-indigo-50 active:scale-95 transition flex items-center justify-center"
                                      title="Edit Presensi"
                                    >
                                      <i className="fa-solid fa-pen-to-square text-xs"></i>
                                    </button>

                                    <button 
                                      onClick={() => requestHapusPresensi(k.id, k.nama_lengkap)}
                                      disabled={!m && !p}
                                      className={`w-8 h-8 rounded-full flex items-center justify-center transition ${
                                        m || p 
                                          ? 'bg-slate-50 text-rose-500 hover:bg-rose-50 active:scale-95' 
                                          : 'bg-slate-50 text-slate-200 cursor-not-allowed'
                                      }`}
                                      title={m || p ? "Hapus Presensi" : "Tidak ada data presensi"}
                                    >
                                      <i className="fa-solid fa-trash text-xs"></i>
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                        })}
                      </tbody>
                    </table>
                      </div>
                    </div>
                  </>;
                })()}
              </div>
            )}

            {}
            {/* Modal Export Rekapan Presensi Harian */}
            {showExportHarianModal && (
              <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-[70] flex items-center justify-center p-4" onClick={() => !isExportingHarian && setShowExportHarianModal(false)}>
                <div className="bg-white rounded-[2rem] max-w-md w-full p-6 shadow-2xl border border-slate-100" onClick={e => e.stopPropagation()}>
                  <div className="flex items-start justify-between gap-3 mb-5">
                    <div>
                      <h3 className="text-lg font-extrabold text-slate-800"><i className="fa-solid fa-file-excel text-emerald-600 mr-2"></i>Export Rekapan Harian</h3>
                      <p className="text-xs text-slate-500 mt-1">Pilih rentang tanggal yang ingin diunduh ke Excel.</p>
                    </div>
                    <button type="button" disabled={isExportingHarian} onClick={() => setShowExportHarianModal(false)} className="w-8 h-8 rounded-full bg-slate-100 text-slate-500 flex items-center justify-center"><i className="fa-solid fa-xmark"></i></button>
                  </div>
                  <div className="space-y-4">
                    <label className="block">
                      <span className="block text-xs font-bold text-slate-600 mb-2">Tanggal Mulai</span>
                      <input type="text" inputMode="numeric" placeholder="dd/mm/yyyy" value={exportTanggalMulaiDraft} onChange={e => setExportTanggalMulaiDraft(e.target.value)} onBlur={() => { const m = exportTanggalMulaiDraft.trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/); if (m) { const d = Number(m[1]), mo = Number(m[2]), y = Number(m[3]); const dt = new Date(y, mo - 1, d); if (dt.getFullYear() === y && dt.getMonth() === mo - 1 && dt.getDate() === d) { const key = `${y}-${m[2]}-${m[1]}`; setExportTanggalMulai(key); if (exportTanggalSelesai && key > exportTanggalSelesai) { setExportTanggalSelesai(key); setExportTanggalSelesaiDraft(m[1] + '/' + m[2] + '/' + m[3]); } return; } } setExportTanggalMulaiDraft(formatDateDMY(exportTanggalMulai)); }} className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-3 text-sm font-semibold text-slate-700 outline-none focus:border-emerald-400" required />
                    </label>
                    <label className="block">
                      <span className="block text-xs font-bold text-slate-600 mb-2">Tanggal Selesai</span>
                      <input type="text" inputMode="numeric" placeholder="dd/mm/yyyy" value={exportTanggalSelesaiDraft} onChange={e => setExportTanggalSelesaiDraft(e.target.value)} onBlur={() => { const m = exportTanggalSelesaiDraft.trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/); if (m) { const d = Number(m[1]), mo = Number(m[2]), y = Number(m[3]); const dt = new Date(y, mo - 1, d); if (dt.getFullYear() === y && dt.getMonth() === mo - 1 && dt.getDate() === d) { const key = `${y}-${m[2]}-${m[1]}`; setExportTanggalSelesai(key); if (exportTanggalMulai && key < exportTanggalMulai) { setExportTanggalMulai(key); setExportTanggalMulaiDraft(m[1] + '/' + m[2] + '/' + m[3]); } return; } } setExportTanggalSelesaiDraft(formatDateDMY(exportTanggalSelesai)); }} className="w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-3 text-sm font-semibold text-slate-700 outline-none focus:border-emerald-400" required />
                    </label>
                    <p className="text-[11px] text-slate-400 leading-relaxed">File berisi daftar semua karyawan untuk setiap tanggal dalam rentang yang dipilih, termasuk yang belum masuk. Status Sakit, Izin, Cuti, dan Libur tidak menampilkan jam atau foto presensi.</p>
                    <div className="flex gap-2 pt-1">
                      <button type="button" disabled={isExportingHarian} onClick={() => setShowExportHarianModal(false)} className="flex-1 py-3 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold">Batal</button>
                      <button type="button" disabled={isExportingHarian} onClick={handleExportRekapanHarian} className="flex-1 py-3 rounded-full bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white text-xs font-bold shadow-sm">{isExportingHarian ? <><i className="fa-solid fa-spinner fa-spin mr-2"></i>Memproses...</> : <><i className="fa-solid fa-download mr-2"></i>Unduh Excel</>}</button>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Modal Pratinjau Foto Selfie */}
            {previewPhotoUrl && (
              <div 
                onClick={() => setPreviewPhotoUrl(null)} 
                className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-50 flex items-center justify-center p-4"
              >
                <div 
                  onClick={e => e.stopPropagation()} 
                  className="bg-white rounded-[2rem] max-w-sm w-full p-5 shadow-2xl border border-slate-100 flex flex-col items-center"
                >
                  <div className="w-full flex justify-between items-center mb-3">
                    <span className="text-xs font-bold text-slate-800 truncate pr-2">{previewPhotoUrl.label}</span>
                    <button 
                      onClick={() => setPreviewPhotoUrl(null)} 
                      className="w-8 h-8 rounded-full bg-slate-100 text-slate-500 hover:bg-slate-200 flex items-center justify-center"
                    >
                      <i className="fa-solid fa-xmark text-sm"></i>
                    </button>
                  </div>
                  <div className="w-full aspect-square rounded-2xl overflow-hidden bg-slate-100 border border-slate-200 shadow-inner flex items-center justify-center">
                    <img 
                      src={previewPhotoUrl.url} 
                      alt="Pratinjau Foto" 
                      className="w-full h-full object-cover" 
                      onError={(e) => { e.target.src = 'https://placehold.co/400x400/f8fafc/94a3b8?text=Gagal+Memuat+Foto'; }}
                    />
                  </div>
                  <button 
                    onClick={() => setPreviewPhotoUrl(null)} 
                    className="w-full mt-4 py-2.5 bg-slate-900 text-white rounded-full font-bold text-xs shadow-md"
                  >
                    Tutup
                  </button>
                </div>
              </div>
            )}

            {/* Modal Bukti Sakit/Izin/Cuti */}
            {statusEvidenceModal && (
              <div
                className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[60] flex items-center justify-center p-4"
                onClick={() => setStatusEvidenceModal(null)}
              >
                <div
                  className="bg-white rounded-[2rem] max-w-md w-full p-5 shadow-2xl border border-slate-100"
                  onClick={e => e.stopPropagation()}
                >
                  <div className="flex items-start justify-between gap-3 mb-4">
                    <div>
                      <h3 className="text-base font-extrabold text-slate-800">Bukti {statusEvidenceModal.status}</h3>
                      <p className="text-xs text-slate-500 mt-1">{statusEvidenceModal.karyawan?.nama_lengkap} · {statusEvidenceModal.tanggal}</p>
                    </div>
                    <button type="button" onClick={() => setStatusEvidenceModal(null)} className="w-8 h-8 rounded-full bg-slate-100 text-slate-500 hover:bg-slate-200 flex items-center justify-center" aria-label="Tutup">
                      <i className="fa-solid fa-xmark"></i>
                    </button>
                  </div>
                  <div className="mb-4">
                    <div className="text-[10px] uppercase tracking-wider font-bold text-slate-400 mb-2">Keterangan</div>
                    <div className="rounded-2xl bg-slate-50 border border-slate-100 p-3 text-sm text-slate-700 whitespace-pre-wrap break-words">{statusEvidenceModal.keterangan?.trim() || 'Tidak ada keterangan yang diisi.'}</div>
                  </div>
                  <div>
                    <div className="text-[10px] uppercase tracking-wider font-bold text-slate-400 mb-2">Foto Bukti</div>
                    {statusEvidenceModal.foto_url ? (
                      <button type="button" className="block w-full" onClick={() => { setPreviewPhotoUrl({ url: getCleanImageUrl(statusEvidenceModal.foto_url), label: `Bukti ${statusEvidenceModal.status} - ${statusEvidenceModal.karyawan?.nama_lengkap}` }); setStatusEvidenceModal(null); }}>
                        <img src={getCleanImageUrl(statusEvidenceModal.foto_url)} alt={`Bukti ${statusEvidenceModal.status}`} className="w-full max-h-80 object-contain rounded-2xl bg-slate-50 border border-slate-200" onError={e => { e.currentTarget.style.display = 'none'; }} />
                        <span className="text-[10px] text-violet-600 font-bold mt-2 inline-block"><i className="fa-solid fa-magnifying-glass-plus mr-1"></i>Klik foto untuk memperbesar</span>
                      </button>
                    ) : (
                      <div className="rounded-2xl bg-slate-50 border border-dashed border-slate-200 p-6 text-center text-xs text-slate-400"><i className="fa-regular fa-image text-2xl block mb-2"></i>Foto bukti tidak tersedia.</div>
                    )}
                  </div>
                  <button type="button" onClick={() => setStatusEvidenceModal(null)} className="w-full mt-5 py-3 bg-slate-900 hover:bg-slate-800 text-white rounded-full font-bold text-xs">Tutup</button>
                </div>
              </div>
            )}

            {/* Modal Konfirmasi Hapus Presensi Harian dari Supabase */}
            {deleteConfirmState && (
              <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
                <div className="bg-white rounded-[2rem] max-w-xs w-full p-6 text-center shadow-2xl border border-slate-100">
                  <div className="w-14 h-14 rounded-full bg-rose-50 text-rose-500 flex items-center justify-center mx-auto mb-4">
                    <i className="fa-solid fa-trash-can text-2xl"></i>
                  </div>
                  <h3 className="font-display font-bold text-base text-slate-800">Hapus Presensi?</h3>
                  <p className="text-xs text-slate-500 mt-1 mb-5 leading-relaxed">
                    Hapus data presensi <span className="font-bold text-slate-700">{deleteConfirmState.karyawanNama}</span> pada tanggal <span className="font-bold text-slate-700">{deleteConfirmState.tanggal}</span>? Data akan langsung terhapus dari Supabase.
                  </p>
                  <div className="flex gap-2">
                    <button 
                      type="button"
                      disabled={isDeleting}
                      onClick={() => setDeleteConfirmState(null)}
                      className="w-1/2 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-full transition"
                    >
                      Batal
                    </button>
                    <button 
                      type="button"
                      disabled={isDeleting}
                      onClick={executeHapusPresensi}
                      className="w-1/2 py-3 bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs rounded-full transition shadow-md active:scale-95 flex items-center justify-center gap-1.5"
                    >
                      {isDeleting ? (
                        <>
                          <i className="fa-solid fa-spinner fa-spin"></i>
                          <span>Menghapus...</span>
                        </>
                      ) : (
                        <span>Ya, Hapus</span>
                      )}
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Modal Edit Presensi Harian */}
            {editingPresensiModal && (
              <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
                <div className="bg-white rounded-[2rem] max-w-md w-full p-6 shadow-2xl border border-slate-100">
                  <div className="flex justify-between items-center mb-4">
                    <div>
                      <span className="text-[10px] font-bold uppercase text-slate-400">Edit Presensi Harian</span>
                      <h3 className="font-display font-bold text-base text-slate-800">{editingPresensiModal.karyawan.nama_lengkap}</h3>
                    </div>
                    <button 
                      onClick={() => setEditingPresensiModal(null)} 
                      className="w-8 h-8 rounded-full bg-slate-50 text-slate-400 hover:text-slate-600"
                    >
                      <i className="fa-solid fa-xmark"></i>
                    </button>
                  </div>

                  <form onSubmit={handleSimpanEditPresensi} className="space-y-4">
                    {/* Shift Terkait */}
                    <div>
                      <label className="text-[10px] font-bold text-slate-400 block mb-1">Penugasan Shift</label>
                      <select 
                        value={editingPresensiModal.formMasuk.shift_nama} 
                        onChange={(e) => setEditingPresensiModal({
                          ...editingPresensiModal,
                          formMasuk: { ...editingPresensiModal.formMasuk, shift_nama: e.target.value }
                        })}
                        className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold text-slate-700 outline-none"
                      >
                        {shiftList.map(s => (
                          <option key={s.id} value={s.nama_shift}>{s.nama_shift}</option>
                        ))}
                      </select>
                    </div>

                    {/* Form Presensi Masuk */}
                    <div className="p-3.5 bg-slate-50 rounded-2xl border border-slate-100 space-y-2.5">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                          <i className="fa-solid fa-arrow-right-to-bracket text-emerald-500"></i> Presensi Masuk
                        </span>
                        <span className="text-[10px] text-slate-400">{formatDateDMY(selectedDate)}</span>
                      </div>
                      <div className="grid grid-cols-2 gap-2">
                        <div>
                          <label className="text-[9px] font-bold text-slate-400 block mb-1">Status Kehadiran</label>
                          <select 
                            value={editingPresensiModal.formMasuk.status} 
                            onChange={(e) => setEditingPresensiModal({
                              ...editingPresensiModal,
                              formMasuk: { ...editingPresensiModal.formMasuk, status: e.target.value }
                            })}
                            className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold text-slate-700 outline-none"
                          >
                            <option value="Hadir">Hadir</option>
                            <option value="Terlambat">Terlambat</option>
                            <option value="Libur">Libur</option>
                            <option value="Izin">Izin</option>
                            <option value="Sakit">Sakit</option>
                            <option value="Cuti">Cuti</option>
                            <option value="Belum">Belum Masuk</option>
                          </select>
                        </div>
                        <div>
                          <label className="text-[9px] font-bold text-slate-400 block mb-1">Jam Masuk</label>
                          <input 
                            type="text" 
                            placeholder="08:00:00"
                            value={editingPresensiModal.formMasuk.jam} 
                            onChange={(e) => setEditingPresensiModal({
                              ...editingPresensiModal,
                              formMasuk: { ...editingPresensiModal.formMasuk, jam: e.target.value }
                            })}
                            className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold outline-none font-mono-numbers"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Form Presensi Pulang */}
                    <div className="p-3.5 bg-slate-50 rounded-2xl border border-slate-100 space-y-2.5">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                          <i className="fa-solid fa-arrow-right-from-bracket text-rose-500"></i> Presensi Pulang
                        </span>
                        <span className="text-[10px] text-slate-400">{formatDateDMY(selectedDate)}</span>
                      </div>
                      <div className="grid grid-cols-2 gap-2">
                        <div>
                          <label className="text-[9px] font-bold text-slate-400 block mb-1">Status Pulang</label>
                          <select 
                            value={editingPresensiModal.formPulang.status} 
                            onChange={(e) => setEditingPresensiModal({
                              ...editingPresensiModal,
                              formPulang: { ...editingPresensiModal.formPulang, status: e.target.value }
                            })}
                            className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold text-slate-700 outline-none"
                          >
                            <option value="Pulang">Pulang</option>
                            <option value="Izin Pulang Cepat">Izin Pulang Cepat</option>
                            <option value="Sakit Mendadak">Sakit Mendadak</option>
                            <option value="Belum">Belum Pulang</option>
                          </select>
                        </div>
                        <div>
                          <label className="text-[9px] font-bold text-slate-400 block mb-1">Jam Pulang</label>
                          <input 
                            type="text" 
                            placeholder="17:00:00"
                            value={editingPresensiModal.formPulang.jam} 
                            onChange={(e) => setEditingPresensiModal({
                              ...editingPresensiModal,
                              formPulang: { ...editingPresensiModal.formPulang, jam: e.target.value }
                            })}
                            className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold outline-none font-mono-numbers"
                          />
                        </div>
                      </div>
                    </div>

                    <div className="flex gap-2 pt-2 border-t border-slate-100">
                      <button 
                        type="button" 
                        onClick={() => setEditingPresensiModal(null)} 
                        className="w-1/3 py-3.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-full font-bold text-xs"
                      >
                        Batal
                      </button>
                      <button 
                        type="submit" 
                        className="w-2/3 py-3.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full font-bold text-xs shadow-md active:scale-95 transition"
                      >
                        Simpan Perubahan
                      </button>
                    </div>
                  </form>
                </div>
              </div>
            )}

            {/* Additional modules map simply to standard modern minimalist white cards... */}
            {activeModule === 'periode' && (
              <RekapanPeriodeView 
                karyawanList={karyawanList}
                presensiList={presensiList}
                lemburList={lemburList}
                exportToExcel={exportToExcel}
                setPreviewPhotoUrl={setPreviewPhotoUrl}
              />
            )}

            {}
            {activeModule === 'libur_cuti' && (
              <MasterLiburCutiView
                karyawanList={karyawanList}
                setKaryawanList={setKaryawanList}
                presensiList={presensiList}
                exportToExcel={exportToExcel}
                showToast={showToast}
              />
            )}

            {activeModule === 'lembur' && (
              <PengajuanLemburAdminView 
                lemburList={lemburList}
                setLemburList={setLemburList}
                showToast={showToast}
                setPreviewPhotoUrl={setPreviewPhotoUrl}
                exportToExcel={exportToExcel}
              />
            )}

            {activeModule === 'klaim_kehadiran' && (
              <KlaimKehadiranAdminView
                klaimList={klaimList}
                setKlaimList={setKlaimList}
                presensiList={presensiList}
                setPresensiList={setPresensiList}
                shiftList={shiftList}
                showToast={showToast}
              />
            )}

            {activeModule === 'lokasi' && <MasterLokasiView lokasiList={lokasiList} setLokasiList={setLokasiList} showToast={showToast} />}
            {activeModule === 'shift' && <MasterShiftView shiftList={shiftList} setShiftList={setShiftList} divisiList={divisiList} setDivisiList={setDivisiList} toleransiMenit={toleransiMenit} setToleransiMenit={setToleransiMenit} showToast={showToast} />}
            {activeModule === 'divisi' && <MasterDivisiView divisiList={divisiList} setDivisiList={setDivisiList} shiftList={shiftList} showToast={showToast} />}
            {activeModule === 'jabatan' && <MasterJabatanView jabatanList={jabatanList} setJabatanList={setJabatanList} divisiList={divisiList} showToast={showToast} />}
            {activeModule === 'admin' && <MasterAdminView adminList={adminList} setAdminList={setAdminList} showToast={showToast} />}
            {activeModule === 'pengumuman' && (
              <MasterPengumumanView 
                pengumumanList={pengumumanList} 
                setPengumumanList={setPengumumanList} 
                showToast={showToast}
                currentUser={activeAdminRecord}
              />
            )}

          </main>
        </div>
      );
    }

    function KlaimKehadiranAdminView({ klaimList = [], setKlaimList, presensiList = [], setPresensiList, shiftList = [], showToast }) {
      const [filterStatus, setFilterStatus] = useState('Pending');
      const [processingId, setProcessingId] = useState(null);
      const [selectedClaimIds, setSelectedClaimIds] = useState([]);
      const claims = (Array.isArray(klaimList) ? klaimList : []).filter(item => {
        const status = String(item.status || 'Pending').toLowerCase();
        if (filterStatus === 'Semua') return true;
        if (filterStatus === 'Pending') return ['pending', 'menunggu', ''].includes(status);
        if (filterStatus === 'Disetujui') return ['disetujui', 'disetujui hr', 'approved'].includes(status);
        return false;
      }).sort((a, b) => String(a.karyawan_nama || '').localeCompare(String(b.karyawan_nama || ''), 'id', { sensitivity: 'base' }));

      const deleteApprovedClaim = async (claim) => {
        if (!claim?.id || processingId) return;
        if (!await modernConfirm(`Hapus riwayat pengajuan klaim ${claim.karyawan_nama || 'karyawan'}? Catatan presensi yang sudah disetujui di Rekapan Harian dan Riwayat karyawan TIDAK akan dihapus.`)) return;
        setProcessingId(claim.id);
        try {
          const client = await getSupabaseClient();
          const { error } = await client.from('klaim_kehadiran').delete().eq('id', claim.id);
          if (error) throw error;
          setKlaimList(prev => (Array.isArray(prev) ? prev : []).filter(item => String(item.id) !== String(claim.id)));
          setSelectedClaimIds(prev => prev.filter(id => id !== String(claim.id)));
          showToast('Riwayat pengajuan klaim dihapus. Catatan presensi yang telah disetujui tetap tersimpan.');
        } catch (error) {
          console.error('Gagal menghapus riwayat klaim yang disetujui:', error);
          showToast(`Gagal menghapus riwayat klaim: ${error?.message || 'periksa izin Supabase.'}`, 'error');
        } finally {
          setProcessingId(null);
        }
      };

      const processClaim = async (claim, action, silent = false) => {
        if (!claim?.id || processingId) return false;
        setProcessingId(claim.id);
        try {
          const client = await getSupabaseClient();
          const safeClaimId = String(claim.id).replace(/[^a-zA-Z0-9_-]/g, '');
          const masukId = 'prs-claim-' + safeClaimId;
          const keluarId = 'prs-claim-' + safeClaimId + '-keluar';
          const claimDate = String(claim.tanggal_klaim || '').slice(0, 10);
          if (action === 'Ditolak') {
            // Bersihkan catatan parsial yang mungkin sempat dibuat oleh versi sebelumnya.
            const { error: cleanupError } = await client.from('presensi').delete().in('id', [masukId, keluarId]);
            if (cleanupError) throw cleanupError;
            const { error } = await client.from('klaim_kehadiran').delete().eq('id', claim.id);
            if (error) throw error;
            if (typeof setPresensiList === 'function') setPresensiList(prev => (Array.isArray(prev) ? prev : []).filter(p => ![masukId, keluarId].includes(String(p.id))));
            setKlaimList((prev) => (Array.isArray(prev) ? prev : []).filter(item => String(item.id) !== String(claim.id)));
            if (!silent) showToast('Klaim ditolak dan dihapus.');
            return true;
          }

          if (!claimDate) throw new Error('Tanggal klaim tidak valid.');
          const selectedShift = (Array.isArray(shiftList) ? shiftList : []).find(s => String(s.nama_shift || '') === String(claim.shift_nama || ''));
          const shiftMasuk = String(claim.shift_jam_masuk || selectedShift?.jam_masuk || '08:00').slice(0, 8);
          const shiftPulang = String(claim.shift_jam_pulang || selectedShift?.jam_pulang || '16:00').slice(0, 8);
          const toTime = value => value.length === 5 ? `${value}:00` : value;
          const nowIso = new Date().toISOString();

          // Periksa apakah catatan masuk yang sudah ada berasal dari klaim versi lama.
          // Catatan klaim parsial boleh diperbaiki; presensi normal tidak boleh diduplikasi.
          const { data: existingRows, error: checkError } = await client.from('presensi')
            .select('*').eq('karyawan_id', claim.karyawan_id).eq('tanggal', claimDate);
          if (checkError) throw checkError;
          const existingMasuk = (existingRows || []).find(p => p.tipe === 'masuk');
          const existingIsClaim = existingMasuk && (String(existingMasuk.id) === masukId || String(existingMasuk.lokasi_nama || '').toLowerCase().includes('klaim kehadiran'));
          if (existingMasuk && !existingIsClaim) {
            throw new Error('Sudah ada presensi masuk normal pada tanggal ini. Periksa Rekapan Harian agar tidak terjadi duplikasi.');
          }

          const common = {
            karyawan_id: claim.karyawan_id,
            karyawan_nama: claim.karyawan_nama,
            divisi: claim.divisi || '',
            status: 'Hadir',
            shift_nama: claim.shift_nama || 'Shift tidak tercatat',
            lokasi_nama: 'Klaim Kehadiran - Disetujui HR',
            latitude: 0,
            longitude: 0,
            foto_url: '',
            keterangan: 'Kehadiran diklaim karyawan dan disetujui HR.',
            tanggal: claimDate,
            created_at: nowIso
          };
          const attendanceRows = [
            { ...common, id: masukId, tipe: 'masuk', jam: toTime(shiftMasuk) },
            { ...common, id: keluarId, tipe: 'keluar', jam: toTime(shiftPulang) }
          ];
          const { error: upsertError } = await client.from('presensi').upsert(attendanceRows, { onConflict: 'id' });
          if (upsertError) throw upsertError;

          // Supabase constraint klaim_kehadiran_status_check menggunakan status enum
          // berbahasa Inggris (Pending / Approved / Rejected). Simpan 'Approved' di DB,
          // sedangkan label UI tetap ditampilkan sebagai 'Disetujui'.
          const updated = { ...claim, status: 'Approved', catatan: 'Disetujui HR. Kehadiran tercatat sesuai shift pada Rekapan Harian.', diproses_at: nowIso };
          const { error: updateClaimError } = await client.from('klaim_kehadiran').update({ status: 'Approved', catatan: updated.catatan, diproses_at: nowIso }).eq('id', claim.id);
          if (updateClaimError) throw updateClaimError;

          if (typeof setPresensiList === 'function') setPresensiList(prev => {
            const list = Array.isArray(prev) ? prev : [];
            const filtered = list.filter(p => ![masukId, keluarId].includes(String(p.id)));
            return [...attendanceRows, ...filtered];
          });
          setKlaimList((prev) => (Array.isArray(prev) ? prev : []).map(item => String(item.id) === String(claim.id) ? updated : item));
          if (!silent) showToast('Klaim disetujui. Jam masuk dan pulang sesuai shift sudah tercatat, dan status klaim diperbarui.');
          return true;
        } catch (error) {
          console.error('Gagal memproses klaim kehadiran:', error);
          if (!silent) showToast(`Klaim gagal diproses: ${error?.message || 'periksa koneksi, skema, dan izin Supabase.'}`, 'error');
          return false;
        } finally {
          setProcessingId(null);
        }
      };

      const pendingClaims = (Array.isArray(klaimList) ? klaimList : []).filter(item => ['pending', 'menunggu', ''].includes(String(item.status || 'Pending').toLowerCase()));
      const pendingCount = pendingClaims.length;
      const visiblePendingIds = claims.filter(item => ['pending', 'menunggu', ''].includes(String(item.status || 'Pending').toLowerCase())).map(item => String(item.id));
      const selectedPending = pendingClaims.filter(item => selectedClaimIds.includes(String(item.id)));
      const toggleClaimSelection = (id) => setSelectedClaimIds(prev => prev.includes(String(id)) ? prev.filter(x => x !== String(id)) : [...prev, String(id)]);
      const toggleVisiblePending = () => setSelectedClaimIds(prev => visiblePendingIds.every(id => prev.includes(id)) ? prev.filter(id => !visiblePendingIds.includes(id)) : [...new Set([...prev, ...visiblePendingIds])]);
      const processBulk = async (action) => {
        const targets = selectedPending.slice().sort((a, b) => String(a.karyawan_nama || '').localeCompare(String(b.karyawan_nama || ''), 'id', { sensitivity: 'base' }));
        if (!targets.length) { showToast('Pilih minimal satu klaim berstatus Pending.', 'error'); return; }
        const label = action === 'Disetujui' ? 'menyetujui' : 'menolak dan menghapus';
        if (!await modernConfirm(`Yakin ingin ${label} ${targets.length} klaim yang dipilih?`)) return;
        let success = 0;
        let failed = 0;
        for (const claim of targets) {
          const ok = await processClaim(claim, action, true);
          if (ok) success++; else failed++;
        }
        setSelectedClaimIds([]);
        showToast(`Proses massal selesai: ${success} berhasil${failed ? `, ${failed} gagal` : ''}.`, failed ? 'error' : 'success');
      };
      return (
        <div className="space-y-5">
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
            <div><h2 className="text-xl font-black text-slate-800">Klaim Kehadiran</h2><p className="text-xs text-slate-500 mt-1">Tinjau pengajuan kehadiran kemarin. Daftar karyawan diurutkan A–Z.</p></div>
            <div className="px-4 py-2 rounded-xl bg-amber-50 text-amber-700 border border-amber-100 text-xs font-bold">{pendingCount} klaim menunggu</div>
          </div>
          <div className="flex flex-wrap gap-2">{['Pending', 'Disetujui', 'Semua'].map(status => <button key={status} onClick={() => { setFilterStatus(status); setSelectedClaimIds([]); }} className={`px-4 py-2 rounded-xl text-xs font-bold border transition ${filterStatus === status ? 'bg-violet-600 text-white border-violet-600' : 'bg-white text-slate-600 border-slate-200 hover:border-violet-300'}`}>{status}</button>)}</div>
          <div className="flex flex-wrap items-center gap-2 rounded-xl border border-violet-100 bg-violet-50/60 p-3">
            <span className="text-xs font-bold text-slate-600 mr-auto">{selectedPending.length} klaim Pending dipilih</span>
            <button onClick={() => processBulk('Disetujui')} disabled={!selectedPending.length || !!processingId} className="px-3 py-2 rounded-lg bg-emerald-600 text-white text-xs font-bold disabled:opacity-40"><i className="fa-solid fa-check-double mr-1"></i>Setujui Massal</button>
            <button onClick={() => processBulk('Ditolak')} disabled={!selectedPending.length || !!processingId} className="px-3 py-2 rounded-lg bg-white border border-rose-200 text-rose-600 text-xs font-bold disabled:opacity-40"><i className="fa-solid fa-xmark mr-1"></i>Tolak Massal</button>
          </div>
          <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
            <div className="overflow-x-auto"><table className="w-full text-left border-collapse text-xs min-w-[900px]">
              <thead><tr className="bg-slate-50 text-slate-500 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider"><th className="p-3 text-center w-10"><input type="checkbox" aria-label="Pilih semua klaim Pending yang terlihat" checked={visiblePendingIds.length > 0 && visiblePendingIds.every(id => selectedClaimIds.includes(id))} onChange={toggleVisiblePending} disabled={!visiblePendingIds.length || !!processingId} className="h-4 w-4 accent-violet-600" /></th><th className="p-4 text-center w-14">No.</th><th className="p-4">Karyawan</th><th className="p-4">Tanggal Pengajuan Kehadiran</th><th className="p-4">Shift yang Dipilih</th><th className="p-4 text-center">Status</th><th className="p-4 text-center">Aksi</th></tr></thead>
              <tbody className="divide-y divide-slate-100">
                {claims.map((claim, index) => {
                  const status = String(claim.status || 'Pending');
                  const statusLabel = ['approved', 'disetujui', 'disetujui hr'].includes(status.toLowerCase()) ? 'Disetujui' : ['rejected', 'ditolak', 'ditolak hr'].includes(status.toLowerCase()) ? 'Ditolak' : status;
                  const pending = ['pending', 'menunggu', ''].includes(status.toLowerCase());
                  const shift = (Array.isArray(shiftList) ? shiftList : []).find(s => String(s.nama_shift || '') === String(claim.shift_nama || ''));
                  const shiftTime = claim.shift_jam_masuk && claim.shift_jam_pulang ? `${formatShiftTimeDisplay(claim.shift_jam_masuk)} - ${formatShiftTimeDisplay(claim.shift_jam_pulang)}` : shift ? `${formatShiftTimeDisplay(shift.jam_masuk)} - ${formatShiftTimeDisplay(shift.jam_pulang)}` : '';
                  return <tr key={claim.id} className="hover:bg-violet-50/30">
                    <td className="p-3 text-center">{pending ? <input type="checkbox" aria-label={`Pilih klaim ${claim.karyawan_nama || "karyawan"}`} checked={selectedClaimIds.includes(String(claim.id))} onChange={() => toggleClaimSelection(claim.id)} disabled={!!processingId} className="h-4 w-4 accent-violet-600" /> : null}</td>
                    <td className="p-4 text-center text-slate-500">{index + 1}</td>
                    <td className="p-4"><div className="font-extrabold text-slate-800">{claim.karyawan_nama || 'Nama tidak tersedia'}</div><div className="text-[10px] text-slate-400 mt-1">{claim.divisi || '-'}</div></td>
                    <td className="p-4 text-slate-600">{claim.tanggal_klaim ? formatDateDMY(claim.tanggal_klaim) : '-'}</td>
                    <td className="p-4"><div className="font-bold text-slate-700">{claim.shift_nama || '-'}</div>{shiftTime && <div className="text-[10px] text-slate-400 mt-1">{shiftTime} WIB</div>}</td>
                    <td className="p-4 text-center"><span className={`inline-flex px-3 py-1.5 rounded-full text-[10px] font-black ${pending ? 'bg-amber-50 text-amber-700' : status.toLowerCase().includes('setuju') ? 'bg-emerald-50 text-emerald-700' : 'bg-rose-50 text-rose-700'}`}>{statusLabel}</span></td>
                    <td className="p-4"><div className="flex justify-center gap-2">{pending ? <><button disabled={processingId === claim.id} onClick={() => processClaim(claim, 'Disetujui')} className="px-3 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-[10px] font-bold"><i className="fa-solid fa-check mr-1"></i>Setujui</button><button disabled={processingId === claim.id} onClick={() => processClaim(claim, 'Ditolak')} className="px-3 py-2 rounded-lg bg-white hover:bg-rose-50 border border-rose-200 text-rose-600 text-[10px] font-bold disabled:opacity-50"><i className="fa-solid fa-xmark mr-1"></i>Tolak</button></> : statusLabel === 'Disetujui' ? <button disabled={!!processingId} onClick={() => deleteApprovedClaim(claim)} className="px-3 py-2 rounded-lg bg-white hover:bg-rose-50 border border-rose-200 text-rose-600 text-[10px] font-bold disabled:opacity-50"><i className="fa-solid fa-trash-can mr-1"></i>Hapus Klaim</button> : <span className="text-[10px] text-slate-400">Selesai diproses</span>}</div></td>
                  </tr>;
                })}
                {claims.length === 0 && <tr><td colSpan="7" className="p-10 text-center"><i className="fa-solid fa-clipboard-check text-3xl text-slate-300 mb-3"></i><p className="font-bold text-slate-600">Tidak ada klaim pada kategori ini.</p><p className="text-xs text-slate-400 mt-1">Pengajuan klaim baru akan muncul di tabel ini.</p></td></tr>}
              </tbody>
            </table></div>
          </div>
        </div>
      );
    }

    function MasterMoreAppsView({ karyawanList = [], moreAppsAccess = {}, setMoreAppsAccess, moreAppsCatalog = DEFAULT_MORE_APPS, setMoreAppsCatalog, showToast }) {
      const appItems = (Array.isArray(moreAppsCatalog) ? moreAppsCatalog : DEFAULT_MORE_APPS).filter(app => app && typeof app === 'object' && typeof app.label === 'string' && app.label.trim());
      const [showAppForm, setShowAppForm] = useState(false);
      const [editingLabel, setEditingLabel] = useState('');
      const [appForm, setAppForm] = useState({ label: '', description: '', icon: 'fa-layer-group', url: '' });
      const [savingCatalog, setSavingCatalog] = useState(false);
      const openNewApp = () => { setEditingLabel(''); setAppForm({ label: '', description: '', icon: 'fa-layer-group', url: '' }); setShowAppForm(true); };
      const openEditApp = app => { setEditingLabel(app.label); setAppForm({ label: app.label, description: app.description || '', icon: app.icon || 'fa-layer-group', url: app.url || '' }); setShowAppForm(true); };
      const saveApp = async () => {
        const label = appForm.label.trim();
        if (!label) { showToast('Masukkan nama aplikasi terlebih dahulu.', 'error'); return; }
        const urlValue = appForm.url.trim();
        if (urlValue) { try { const parsedUrl = new URL(urlValue); if (!['http:', 'https:'].includes(parsedUrl.protocol)) throw new Error('protocol'); } catch (_) { showToast('URL tujuan harus lengkap dan diawali https:// atau http://.', 'error'); return; } }
        if (appItems.some(a => a.label.toLowerCase() === label.toLowerCase() && a.label !== editingLabel)) { showToast('Nama aplikasi sudah digunakan.', 'error'); return; }
        if (typeof setMoreAppsCatalog !== 'function') { showToast('Fungsi penyimpanan katalog belum tersedia.', 'error'); return; }
        const palette = ['from-violet-50 to-purple-100|text-violet-700','from-sky-50 to-cyan-100|text-sky-700','from-amber-50 to-orange-100|text-amber-700','from-emerald-50 to-teal-100|text-emerald-700','from-fuchsia-50 to-pink-100|text-fuchsia-700'];
        const [color, ink] = palette[appItems.length % palette.length].split('|');
        const next = editingLabel
          ? appItems.map(a => a.label === editingLabel ? { ...a, label, description: appForm.description.trim(), icon: appForm.icon.trim() || 'fa-layer-group', url: urlValue } : a)
          : [...appItems, { label, description: appForm.description.trim(), icon: appForm.icon.trim() || 'fa-layer-group', url: urlValue, color, ink }];
        setSavingCatalog(true);
        try {
          await setMoreAppsCatalog(next);
          if (editingLabel && editingLabel !== label) {
            const nextAccess = { ...(moreAppsAccess || {}) };
            nextAccess[label] = nextAccess[editingLabel] || [];
            delete nextAccess[editingLabel];
            await setMoreAppsAccess(nextAccess);
          }
          setSelectedApp(label); setShowAppForm(false);
          showToast(editingLabel ? 'Aplikasi berhasil diperbarui.' : 'Aplikasi baru berhasil ditambahkan. Default akses kosong; setelah karyawan diberi akses, tombolnya menampilkan Coming Soon.');
        } catch (e) { showToast(`Gagal menyimpan katalog aplikasi: ${e.message || 'Periksa izin Supabase.'}`, 'error'); }
        finally { setSavingCatalog(false); }
      };
      const deleteApp = async app => {
        if (!await modernConfirm(`Hapus aplikasi "${app.label}" dari More Apps?`)) return;
        try {
          await setMoreAppsCatalog(appItems.filter(a => a.label !== app.label));
          const nextAccess = { ...(moreAppsAccess || {}) }; delete nextAccess[app.label];
          await setMoreAppsAccess(nextAccess);
          if (selectedApp === app.label) setSelectedApp((appItems.find(a => a.label !== app.label) || {}).label || '');
          showToast('Aplikasi berhasil dihapus.');
        } catch (e) { showToast(`Gagal menghapus aplikasi: ${e.message || 'Periksa izin Supabase.'}`, 'error'); }
      };
      const [selectedApp, setSelectedApp] = useState(appItems[0]?.label || '');
      const [search, setSearch] = useState('');
      const [draftIds, setDraftIds] = useState((moreAppsAccess?.[appItems[0]?.label] || []).map(String));
      const [saving, setSaving] = useState(false);
      useEffect(() => { setDraftIds((moreAppsAccess?.[selectedApp] || []).map(String)); }, [selectedApp, moreAppsAccess]);
      const filtered = karyawanList.filter(k => `${k.nama_lengkap || ''} ${k.nik || ''} ${k.divisi || ''} ${k.jabatan || ''}`.toLowerCase().includes(search.toLowerCase()));
      const setAll = checked => setDraftIds(checked ? karyawanList.map(k => String(k.id)) : []);
      const save = async () => {
        if (typeof setMoreAppsAccess !== 'function') { showToast('Fungsi penyimpanan pengaturan tidak tersedia.', 'error'); return; }
        setSaving(true);
        try {
          const next = { ...(moreAppsAccess || {}), [selectedApp]: draftIds.map(String) };
          await setMoreAppsAccess(next);
          showToast(`Akses ${selectedApp} berhasil disimpan untuk ${draftIds.length} karyawan.`);
        } catch (e) {
          console.error('Gagal menyimpan akses More Apps:', e);
          showToast(`Gagal menyimpan akses More Apps: ${e.message || 'Periksa koneksi dan izin tabel settings.'}`, 'error');
        } finally { setSaving(false); }
      };
      return (
        <div className="space-y-5">
          <div className="rounded-3xl bg-gradient-to-r from-violet-600 to-fuchsia-600 p-6 text-white shadow-lg">
            <div className="flex items-center gap-3"><div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-white/20 text-xl"><i className="fa-solid fa-layer-group"></i></div><div><h2 className="text-xl font-extrabold">Master More Apps</h2><p className="mt-1 text-xs text-white/80">Atur karyawan yang dapat melihat setiap menu di halaman More Apps.</p></div></div>
          </div>
          <section className="rounded-3xl border border-slate-100 bg-white p-4 shadow-sm">
            <div className="flex flex-wrap items-center justify-between gap-3 mb-3"><div><h3 className="text-sm font-extrabold text-slate-800">Kelola Aplikasi</h3><p className="mt-1 text-xs text-slate-500">Tambah aplikasi baru, ubah nama, deskripsi, ikon, atau hapus aplikasi.</p></div><button type="button" onClick={openNewApp} className="rounded-full bg-violet-600 px-4 py-2.5 text-xs font-bold text-white hover:bg-violet-700"><i className="fa-solid fa-plus mr-1.5"></i>Buat Aplikasi</button></div>
            <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-2">
              {appItems.map(app => <div key={app.label} className="flex items-center gap-3 rounded-2xl border border-slate-100 p-3"><div className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-xl ${app.color || 'bg-violet-50'} ${app.ink || 'text-violet-700'}`}><i className={`fa-solid ${app.icon || 'fa-layer-group'}`}></i></div><div className="min-w-0 flex-1"><p className="truncate text-xs font-extrabold text-slate-800">{app.label}</p><p className="truncate text-[10px] text-slate-500">{app.description || 'Tanpa deskripsi'}</p></div><button title="Edit aplikasi" onClick={() => openEditApp(app)} className="h-8 w-8 rounded-lg text-violet-600 hover:bg-violet-50"><i className="fa-solid fa-pen"></i></button><button title="Hapus aplikasi" onClick={() => deleteApp(app)} className="h-8 w-8 rounded-lg text-rose-500 hover:bg-rose-50"><i className="fa-solid fa-trash"></i></button></div>)}
            </div>
          </section>
          {showAppForm && <div className="fixed inset-0 z-[150] flex items-center justify-center bg-slate-950/50 p-4 backdrop-blur-sm"><div className="w-full max-w-md rounded-3xl bg-white p-6 shadow-2xl"><div className="mb-5 flex items-center justify-between"><div><h3 className="text-lg font-extrabold text-slate-900">{editingLabel ? 'Edit Aplikasi' : 'Buat Aplikasi Baru'}</h3><p className="mt-1 text-xs text-slate-500">Tombol karyawan otomatis menggunakan popup Coming Soon.</p></div><button onClick={() => setShowAppForm(false)} className="h-9 w-9 rounded-full bg-slate-100 text-slate-500"><i className="fa-solid fa-xmark"></i></button></div><div className="space-y-4"><label className="block text-xs font-bold text-slate-600">Nama aplikasi<input value={appForm.label} onChange={e => setAppForm(v => ({...v,label:e.target.value}))} className="mt-1 w-full rounded-xl border border-slate-200 p-3 text-sm outline-none focus:border-violet-400" placeholder="Contoh: Cek Promo" /></label><label className="block text-xs font-bold text-slate-600">Deskripsi aplikasi<textarea value={appForm.description} onChange={e => setAppForm(v => ({...v,description:e.target.value}))} className="mt-1 w-full rounded-xl border border-slate-200 p-3 text-sm outline-none focus:border-violet-400" rows="2" placeholder="Jelaskan fungsi aplikasi" /></label><label className="block text-xs font-bold text-slate-600">Ikon Font Awesome<input value={appForm.icon} onChange={e => setAppForm(v => ({...v,icon:e.target.value}))} className="mt-1 w-full rounded-xl border border-slate-200 p-3 text-sm outline-none focus:border-violet-400" placeholder="fa-layer-group" /><span className="mt-1 block text-[10px] text-slate-400">Masukkan nama kelas ikon, misalnya fa-box, fa-tags, atau fa-calculator.</span></label><label className="block text-xs font-bold text-slate-600">URL tujuan aplikasi (opsional)<input type="url" value={appForm.url || ''} onChange={e => setAppForm(v => ({...v,url:e.target.value}))} className="mt-1 w-full rounded-xl border border-slate-200 p-3 text-sm outline-none focus:border-violet-400" placeholder="https://contoh-domain.com" /><span className="mt-1 block text-[10px] text-slate-400">Jika diisi, tombol More Apps akan membuka URL ini. Jika dikosongkan, tombol menampilkan popup Coming Soon.</span></label><div className="flex items-center gap-3 rounded-xl bg-violet-50 p-3 text-violet-700"><i className={`fa-solid ${appForm.icon || 'fa-layer-group'} text-xl`}></i><span className="text-xs font-bold">{appForm.label || 'Pratinjau ikon aplikasi'}</span></div><div className="flex gap-2 pt-2"><button type="button" disabled={savingCatalog} onClick={() => setShowAppForm(false)} className="flex-1 rounded-full bg-slate-100 py-3 text-xs font-bold text-slate-600">Batal</button><button type="button" disabled={savingCatalog} onClick={saveApp} className="flex-1 rounded-full bg-violet-600 py-3 text-xs font-bold text-white disabled:opacity-50">{savingCatalog ? 'Menyimpan...' : 'Simpan Aplikasi'}</button></div></div></div></div>}
          <div className="grid grid-cols-1 gap-5 lg:grid-cols-[minmax(240px,0.8fr)_minmax(320px,1.2fr)]">
            <section className="rounded-3xl border border-slate-100 bg-white p-4 shadow-sm"><h3 className="mb-3 text-sm font-extrabold text-slate-800">Pilih Menu / Aplikasi</h3><div className="space-y-2">{appItems.map(app => <button key={app.label} onClick={() => setSelectedApp(app.label)} className={`flex w-full items-center gap-3 rounded-2xl border p-3 text-left transition ${selectedApp === app.label ? 'border-violet-200 bg-violet-50 text-violet-800' : 'border-slate-100 hover:bg-slate-50 text-slate-600'}`}><span className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-xl ${selectedApp === app.label ? 'bg-violet-600 text-white' : 'bg-slate-100 text-slate-500'}`}><i className={`fa-solid ${app.icon}`}></i></span><span className="min-w-0 flex-1"><span className="block text-xs font-extrabold">{app.label}</span><span className="mt-0.5 block text-[10px] text-slate-500">{app.description}</span></span><i className={`fa-solid fa-chevron-right text-[10px] ${selectedApp === app.label ? 'text-violet-500' : 'text-slate-300'}`}></i></button>)}</div></section>
            <section className="overflow-hidden rounded-3xl border border-slate-100 bg-white shadow-sm"><div className="border-b border-slate-100 p-5"><div className="flex flex-wrap items-start justify-between gap-3"><div><h3 className="text-base font-extrabold text-slate-800">Daftar Karyawan</h3><p className="mt-1 text-xs text-slate-500">Akses untuk: <strong>{selectedApp}</strong></p></div><span className="rounded-full bg-violet-50 px-3 py-1.5 text-xs font-extrabold text-violet-700">{draftIds.length} dipilih</span></div><div className="mt-4"><input value={search} onChange={e => setSearch(e.target.value)} placeholder="Cari nama, NIK, divisi..." className="w-full rounded-xl border border-slate-200 px-4 py-3 text-xs outline-none focus:border-violet-400 focus:ring-2 focus:ring-violet-100" /></div><div className="mt-3 flex flex-wrap gap-2"><button type="button" onClick={() => setAll(true)} className="rounded-full border border-violet-200 bg-violet-50 px-3 py-2 text-[11px] font-bold text-violet-700 hover:bg-violet-100"><i className="fa-solid fa-check-double mr-1"></i>Select All</button><button type="button" onClick={() => setAll(false)} className="rounded-full border border-slate-200 bg-white px-3 py-2 text-[11px] font-bold text-slate-600 hover:bg-slate-50"><i className="fa-solid fa-xmark mr-1"></i>Unselect All</button></div></div><div className="max-h-[52vh] overflow-y-auto p-3">{filtered.map(k => { const checked = draftIds.includes(String(k.id)); return <label key={k.id} className={`mb-1 flex cursor-pointer items-center gap-3 rounded-xl p-3 transition ${checked ? 'bg-violet-50' : 'hover:bg-slate-50'}`}><input type="checkbox" checked={checked} onChange={e => setDraftIds(prev => e.target.checked ? [...new Set([...prev, String(k.id)])] : prev.filter(id => id !== String(k.id)))} className="h-4 w-4 accent-violet-600"/><span className="min-w-0 flex-1"><span className="block truncate text-xs font-bold text-slate-800">{k.nama_lengkap || 'Tanpa nama'}</span><span className="mt-0.5 block truncate text-[10px] text-slate-500">{[k.nik, k.divisi, k.jabatan].filter(Boolean).join(' • ') || 'Data karyawan'}</span></span></label>; })}{filtered.length === 0 && <p className="p-6 text-center text-xs text-slate-400">Karyawan tidak ditemukan.</p>}</div><div className="border-t border-slate-100 bg-slate-50 p-4"><button disabled={saving} onClick={save} className="flex w-full items-center justify-center gap-2 rounded-xl bg-violet-600 px-4 py-3 text-xs font-extrabold text-white shadow-sm transition hover:bg-violet-700 disabled:opacity-60"><i className={`fa-solid ${saving ? 'fa-spinner fa-spin' : 'fa-floppy-disk'}`}></i>{saving ? 'Menyimpan...' : 'Simpan Pengaturan Akses'}</button></div></section>
          </div>
        </div>
      );
    }

    /* =========================================================
       MODUL MASTER LIBUR & CUTI (JATAH LIBUR & JATAH CUTI)
       ========================================================= */
    function MasterLiburCutiView({ karyawanList, setKaryawanList, presensiList, exportToExcel, showToast }) {
      const [selectedPeriodIndex, setSelectedPeriodIndex] = useState(0);
      const [searchQuery, setSearchQuery] = useState('');
      const [editingModal, setEditingModal] = useState(null);

      // Membuat daftar siklus 6 periode 26 s/d 25 ke belakang
      const periods = useMemo(() => {
        const list = [];
        const now = new Date();
        let baseMonth = now.getDate() > 25 ? now.getMonth() + 1 : now.getMonth();
        let baseYear = now.getFullYear();

        for (let i = 0; i < 6; i++) {
          let start = new Date(baseYear, baseMonth - 1 - i, 26);
          let end = new Date(baseYear, baseMonth - i, 25);
          let label = `${start.getDate()} ${start.toLocaleDateString('id-ID', { month: 'short' })} - ${end.getDate()} ${end.toLocaleDateString('id-ID', { month: 'short', year: 'numeric' })}`;
          list.push({ start, end, label });
        }
        return list;
      }, []);

      const activePeriod = periods[selectedPeriodIndex];

      const isDateInRange = (dateStr, start, end) => {
        if (!dateStr) return false;
        const [y, m, d] = dateStr.split('-').map(Number);
        const recordDate = new Date(y, m - 1, d);
        const s = new Date(start); s.setHours(0, 0, 0, 0);
        const e = new Date(end); e.setHours(23, 59, 59, 999);
        return recordDate >= s && recordDate <= e;
      };

      // Memproses data master libur & cuti berdasarkan acuan data awal di profil karyawan
      const processedData = useMemo(() => {
        return [...karyawanList]
          .sort((a, b) => a.nama_lengkap.localeCompare(b.nama_lengkap))
          .map(k => {
            // Jatah Libur acuan awal dari k.sisa_libur (atau default 4)
            const jatahLiburAcuan = (k.sisa_libur !== undefined && k.sisa_libur !== null) ? Number(k.sisa_libur) : 4;
            const liburTerpakai = presensiList.filter(att => 
              att.karyawan_id === k.id && 
              att.status === 'Libur' && 
              isDateInRange(att.tanggal, activePeriod.start, activePeriod.end)
            ).length;
            const sisaLiburAktif = Math.max(0, jatahLiburAcuan - liburTerpakai);

            // Jatah Cuti acuan awal dari k.sisa_cuti (default 0)
            const jatahCutiAcuan = (k.sisa_cuti !== undefined && k.sisa_cuti !== null) ? Number(k.sisa_cuti) : 0;
            const cutiTerpakai = presensiList.filter(att => 
              att.karyawan_id === k.id && 
              att.status === 'Cuti' && 
              isDateInRange(att.tanggal, activePeriod.start, activePeriod.end)
            ).length;
            const sisaCutiAktif = Math.max(0, jatahCutiAcuan - cutiTerpakai);

            return {
              karyawan: k,
              jatahLiburAcuan,
              liburTerpakai,
              sisaLiburAktif,
              jatahCutiAcuan,
              cutiTerpakai,
              sisaCutiAktif
            };
          });
      }, [karyawanList, presensiList, activePeriod]);

      const filteredData = useMemo(() => {
        if (!searchQuery.trim()) return processedData;
        const q = searchQuery.toLowerCase();
        return processedData.filter(d => 
          d.karyawan.nama_lengkap.toLowerCase().includes(q) ||
          d.karyawan.divisi.toLowerCase().includes(q) ||
          (d.karyawan.jabatan && d.karyawan.jabatan.toLowerCase().includes(q))
        );
      }, [processedData, searchQuery]);

      // Simpan perubahan Jatah Libur & Jatah Cuti
      const handleSaveAdjustment = (e) => {
        e.preventDefault();
        const { karyawan, newLibur, newCuti } = editingModal;

        setKaryawanList(prev => prev.map(k => {
          if (k.id === karyawan.id) {
            return {
              ...k,
              sisa_libur: parseInt(newLibur) || 0,
              sisa_cuti: parseInt(newCuti) || 0
            };
          }
          return k;
        }));

        setEditingModal(null);
        showToast(`Jatah libur & cuti ${karyawan.nama_lengkap} diperbarui!`);
      };

      const handleExport = () => {
        const rows = filteredData.map((d, idx) => ({
          'No': idx + 1,
          'Nama Karyawan': d.karyawan.nama_lengkap,
          'Divisi': d.karyawan.divisi,
          'Periode': activePeriod.label,
          'Jatah Libur': d.jatahLiburAcuan,
          'Libur Terpakai': d.liburTerpakai,
          'Sisa Libur': d.sisaLiburAktif,
          'Jatah Cuti': d.jatahCutiAcuan,
          'Cuti Terpakai': d.cutiTerpakai,
          'Sisa Cuti': d.sisaCutiAktif
        }));
        exportToExcel(rows, `Master_Libur_Cuti_${activePeriod.label.replace(/\s+/g, '_')}`);
      };

      return (
        <div className="space-y-6">
          {/* Header Panel */}
          <div className="flex flex-wrap justify-between items-center gap-4 bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
            <div>
              <div className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full bg-sky-500"></span>
                <span className="text-[10px] font-bold uppercase tracking-widest text-sky-600">Leave & Time-Off Management</span>
              </div>
              <h2 className="font-display font-bold text-xl text-slate-800 mt-1">Master Libur & Cuti</h2>
              <p className="text-xs text-slate-400 font-medium">
                Kelola data acuan awal jatah libur dan jatah cuti (default 0) per karyawan
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-3">
              <div className="flex items-center space-x-2 bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-2xl shadow-sm">
                <i className="fa-regular fa-calendar-days text-slate-400 text-xs"></i>
                <select 
                  value={selectedPeriodIndex} 
                  onChange={(e) => setSelectedPeriodIndex(Number(e.target.value))}
                  className="bg-transparent border-0 text-xs font-bold text-slate-700 outline-none cursor-pointer"
                >
                  {periods.map((p, idx) => (
                    <option key={idx} value={idx}>{p.label}</option>
                  ))}
                </select>
              </div>

              <button 
                onClick={handleExport}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-2xl text-xs font-bold shadow-md active:scale-95 transition flex items-center gap-2"
              >
                <i className="fa-solid fa-file-excel"></i>
                <span>Ekspor Excel</span>
              </button>
            </div>
          </div>

          {/* Tabel Master */}
          <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between gap-4">
              <div className="relative flex-1 max-w-sm">
                <i className="fa-solid fa-magnifying-glass absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
                <input 
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Cari nama karyawan atau divisi..."
                  className="w-full bg-slate-50 border border-slate-200 rounded-2xl pl-9 pr-4 py-2 text-xs font-semibold outline-none focus:border-sky-500"
                />
              </div>

              <div className="text-[11px] font-bold text-slate-500">
                Total: <span className="text-sky-600 font-mono-numbers">{filteredData.length}</span> Karyawan
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                    <th className="p-4 text-center">No</th>
                    <th className="p-4">Karyawan</th>
                    <th className="p-4 text-center">Jatah Libur (Acuan)</th>
                    <th className="p-4 text-center">Libur Terpakai</th>
                    <th className="p-4 text-center">Sisa Libur</th>
                    <th className="p-4 text-center">Jatah Cuti (Acuan)</th>
                    <th className="p-4 text-center">Cuti Terpakai</th>
                    <th className="p-4 text-center">Sisa Cuti</th>
                    <th className="p-4 text-center">Aksi Edit</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {filteredData.length > 0 ? (
                    filteredData.map((d, idx) => (
                      <tr key={d.karyawan.id} className="hover:bg-sky-50/30 transition">
                        <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{idx + 1}</td>
                        <td className="p-4">
                          <div className="font-bold text-slate-800">{d.karyawan.nama_lengkap}</div>
                          <div className="text-[10px] text-slate-400 font-medium">
                            {d.karyawan.divisi} {d.karyawan.jabatan ? `• ${d.karyawan.jabatan}` : ''}
                          </div>
                        </td>
                        <td className="p-4 text-center font-mono-numbers font-bold text-sky-700 bg-sky-50/30">
                          {d.jatahLiburAcuan} <span className="text-[9px] font-normal text-slate-400">hr</span>
                        </td>
                        <td className="p-4 text-center font-mono-numbers text-amber-600">
                          {d.liburTerpakai}
                        </td>
                        <td className="p-4 text-center font-mono-numbers font-bold text-slate-700">
                          {d.sisaLiburAktif} <span className="text-[9px] font-normal text-slate-400">hr</span>
                        </td>
                        <td className="p-4 text-center font-mono-numbers font-bold text-purple-700 bg-purple-50/30">
                          {d.jatahCutiAcuan} <span className="text-[9px] font-normal text-slate-400">hr</span>
                        </td>
                        <td className="p-4 text-center font-mono-numbers text-orange-600">
                          {d.cutiTerpakai}
                        </td>
                        <td className="p-4 text-center font-mono-numbers font-bold text-slate-700">
                          {d.sisaCutiAktif} <span className="text-[9px] font-normal text-slate-400">hr</span>
                        </td>
                        <td className="p-4 text-center">
                          <button 
                            onClick={() => setEditingModal({
                              karyawan: d.karyawan,
                              newLibur: d.jatahLiburAcuan,
                              newCuti: d.jatahCutiAcuan
                            })}
                            className="w-8 h-8 rounded-full bg-slate-50 text-indigo-500 hover:bg-indigo-50 active:scale-95 transition flex items-center justify-center mx-auto"
                            title="Edit Jatah Libur & Cuti"
                          >
                            <i className="fa-solid fa-pen-to-square text-xs"></i>
                          </button>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="9" className="p-8 text-center text-slate-400 font-medium">
                        Tidak ditemukan data karyawan.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* Modal Edit Acuan Jatah */}
          {editingModal && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-sm w-full p-6 shadow-2xl border border-slate-100">
                <div className="flex justify-between items-center mb-4">
                  <div>
                    <span className="text-[10px] font-bold uppercase text-slate-400">Edit Data Acuan Awal</span>
                    <h3 className="font-display font-bold text-base text-slate-800">{editingModal.karyawan.nama_lengkap}</h3>
                  </div>
                  <button onClick={() => setEditingModal(null)} className="w-8 h-8 rounded-full bg-slate-50 text-slate-400 hover:text-slate-600">
                    <i className="fa-solid fa-xmark"></i>
                  </button>
                </div>

                <form onSubmit={handleSaveAdjustment} className="space-y-4">
                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 space-y-3">
                    <div>
                      <label className="text-[10px] font-bold text-slate-600 block mb-1">
                        Jatah Libur (Acuan Awal)
                      </label>
                      <input 
                        type="number" 
                        min="0" 
                        max="30"
                        value={editingModal.newLibur}
                        onChange={(e) => setEditingModal({ ...editingModal, newLibur: e.target.value })}
                        className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2.5 font-bold outline-none font-mono-numbers"
                        required
                      />
                    </div>

                    <div>
                      <label className="text-[10px] font-bold text-slate-600 block mb-1">
                        Jatah Cuti (Acuan Awal, Default: 0)
                      </label>
                      <input 
                        type="number" 
                        min="0" 
                        max="30"
                        value={editingModal.newCuti}
                        onChange={(e) => setEditingModal({ ...editingModal, newCuti: e.target.value })}
                        className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2.5 font-bold outline-none font-mono-numbers"
                        required
                      />
                    </div>
                  </div>

                  <div className="flex gap-2 pt-2 border-t border-slate-100">
                    <button type="button" onClick={() => setEditingModal(null)} className="w-1/3 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-full font-bold text-xs">
                      Batal
                    </button>
                    <button type="submit" className="w-2/3 py-3 bg-slate-900 hover:bg-slate-800 text-white rounded-full font-bold text-xs shadow-md active:scale-95 transition">
                      Simpan Acuan
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}
        </div>
      );
    }

    /* =========================================================
       MODUL REKAPAN PERIODE (26 - 25) CUT-OFF PAYROLL
       ========================================================= */
    function RekapanPeriodeView({ karyawanList, presensiList, lemburList, exportToExcel, setPreviewPhotoUrl }) {
      const [selectedPeriodIndex, setSelectedPeriodIndex] = useState(0);
      const [searchQuery, setSearchQuery] = useState('');
      const [selectedKaryawanDetail, setSelectedKaryawanDetail] = useState(null);
      const [filterDetailStatus, setFilterDetailStatus] = useState('Semua');

      // Siklus 6 bulan ke belakang (Tanggal 26 s/d 25)
      const periods = useMemo(() => {
        const list = [];
        const now = new Date();
        let baseMonth = now.getDate() > 25 ? now.getMonth() + 1 : now.getMonth();
        let baseYear = now.getFullYear();

        for (let i = 0; i < 6; i++) {
          let start = new Date(baseYear, baseMonth - 1 - i, 26);
          let end = new Date(baseYear, baseMonth - i, 25);
          let label = `${start.getDate()} ${start.toLocaleDateString('id-ID', { month: 'short' })} - ${end.getDate()} ${end.toLocaleDateString('id-ID', { month: 'short', year: 'numeric' })}`;
          list.push({ start, end, label, key: `${start.toISOString().split('T')[0]}_${end.toISOString().split('T')[0]}` });
        }
        return list;
      }, []);

      const activePeriod = periods[selectedPeriodIndex];

      const isDateInPeriod = (dateStr) => {
        if (!dateStr) return false;
        const [y, m, d] = dateStr.split('-').map(Number);
        const recordDate = new Date(y, m - 1, d);
        const s = new Date(activePeriod.start); s.setHours(0, 0, 0, 0);
        const e = new Date(activePeriod.end); e.setHours(23, 59, 59, 999);
        return recordDate >= s && recordDate <= e;
      };

      // Rekapitulasi per karyawan dalam periode aktif
      const summaryData = useMemo(() => {
        return [...karyawanList]
          .sort((a, b) => a.nama_lengkap.localeCompare(b.nama_lengkap))
          .map(k => {
            const userPresensi = presensiList.filter(p => p.karyawan_id === k.id && isDateInPeriod(p.tanggal));
            const masuk = userPresensi.filter(p => p.tipe === 'masuk' && (p.status === 'Hadir' || p.status === 'Terlambat')).length;
            const tepatWaktu = userPresensi.filter(p => p.tipe === 'masuk' && p.status === 'Hadir').length;
            const terlambat = userPresensi.filter(p => p.tipe === 'masuk' && p.status === 'Terlambat').length;
            const libur = userPresensi.filter(p => p.status === 'Libur').length;
            const sakit = userPresensi.filter(p => p.status === 'Sakit').length;
            const izin = userPresensi.filter(p => p.status === 'Izin').length;
            const cuti = userPresensi.filter(p => p.status === 'Cuti').length;

            const userLembur = lemburList.filter(l => l.karyawan_id === k.id && isDateInPeriod(l.tanggal) && l.status === 'Approved');
            const totalJamLembur = userLembur.reduce((acc, curr) => acc + Number(curr.durasi_jam || 0), 0);

            return {
              karyawan: k,
              masuk,
              tepatWaktu,
              terlambat,
              libur,
              sakit,
              izin,
              cuti,
              lemburCount: userLembur.length,
              totalJamLembur: Math.round(totalJamLembur * 10) / 10
            };
          });
      }, [karyawanList, presensiList, lemburList, activePeriod]);

      const filteredSummary = useMemo(() => {
        if (!searchQuery.trim()) return summaryData;
        const q = searchQuery.toLowerCase();
        return summaryData.filter(item => 
          item.karyawan.nama_lengkap.toLowerCase().includes(q) ||
          item.karyawan.divisi.toLowerCase().includes(q) ||
          (item.karyawan.jabatan && item.karyawan.jabatan.toLowerCase().includes(q))
        );
      }, [summaryData, searchQuery]);

      // Ekspor Rekapan Periode ke Excel
      const handleExportRekapan = () => {
        const rows = filteredSummary.map((item, idx) => ({
          'No': idx + 1,
          'Nama Karyawan': item.karyawan.nama_lengkap,
          'Divisi': item.karyawan.divisi,
          'Jabatan': item.karyawan.jabatan || '-',
          'Total Masuk': item.masuk,
          'Tepat Waktu': item.tepatWaktu,
          'Terlambat': item.terlambat,
          'Libur': item.libur,
          'Sakit': item.sakit,
          'Izin': item.izin,
          'Cuti': item.cuti,
          'Total Lembur (Jam)': item.totalJamLembur
        }));
        exportToExcel(rows, `Rekapan_Presensi_${activePeriod.label.replace(/\s+/g, '_')}`);
      };

      // Data detail untuk drill-down karyawan terpilih
      const detailPresensiKaryawan = useMemo(() => {
        if (!selectedKaryawanDetail) return [];
        return presensiList.filter(p => p.karyawan_id === selectedKaryawanDetail.id && isDateInPeriod(p.tanggal));
      }, [selectedKaryawanDetail, presensiList, activePeriod]);

      // Data lembur untuk drill-down karyawan terpilih
      const detailLemburKaryawan = useMemo(() => {
        if (!selectedKaryawanDetail) return [];
        return lemburList.filter(l => l.karyawan_id === selectedKaryawanDetail.id && isDateInPeriod(l.tanggal));
      }, [selectedKaryawanDetail, lemburList, activePeriod]);

      const filteredDetailPresensi = useMemo(() => {
        if (filterDetailStatus === 'Semua') return detailPresensiKaryawan;
        if (filterDetailStatus === 'Lembur') return [];
        return detailPresensiKaryawan.filter(p => p.status === filterDetailStatus);
      }, [detailPresensiKaryawan, filterDetailStatus]);

      return (
        <div className="space-y-6">
          {/* Header Panel */}
          <div className="flex flex-wrap justify-between items-center gap-4 bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
            <div>
              <div className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full bg-violet-500"></span>
                <span className="text-[10px] font-bold uppercase tracking-widest text-violet-600">Cut-Off Payroll System</span>
              </div>
              <h2 className="font-display font-bold text-xl text-slate-800 mt-1">Rekapan Periode 26 - 25</h2>
              <p className="text-xs text-slate-400 font-medium">Rekapitulasi kumulatif presensi, libur, sakit, cuti, dan jam lembur per periode</p>
            </div>

            <div className="flex flex-wrap items-center gap-3">
              {/* Dropdown Periode */}
              <div className="flex items-center space-x-2 bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-2xl shadow-sm">
                <i className="fa-regular fa-calendar-days text-slate-400 text-xs"></i>
                <select 
                  value={selectedPeriodIndex} 
                  onChange={(e) => {
                    setSelectedPeriodIndex(Number(e.target.value));
                    setSelectedKaryawanDetail(null);
                  }}
                  className="bg-transparent border-0 text-xs font-bold text-slate-700 outline-none cursor-pointer"
                >
                  {periods.map((p, idx) => (
                    <option key={idx} value={idx}>{p.label}</option>
                  ))}
                </select>
              </div>

              {/* Tombol Ekspor Excel */}
              <button 
                onClick={handleExportRekapan}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-2xl text-xs font-bold shadow-md active:scale-95 transition flex items-center gap-2"
              >
                <i className="fa-solid fa-file-excel"></i>
                <span>Ekspor Excel</span>
              </button>
            </div>
          </div>

          {/* Drill-Down View jika ada karyawan yang diklik */}
          {selectedKaryawanDetail ? (
            <div className="bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm space-y-5">
              <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-slate-100">
                <div className="flex items-center space-x-3.5">
                  <div className="w-12 h-12 rounded-full bg-slate-100 border border-slate-200 overflow-hidden flex items-center justify-center font-bold text-slate-500">
                    {selectedKaryawanDetail.foto_profil ? (
                      <img src={getCleanImageUrl(selectedKaryawanDetail.foto_profil)} className="w-full h-full object-cover" />
                    ) : (
                      selectedKaryawanDetail.nama_lengkap.split(' ').map(n=>n[0]).slice(0, 2).join('')
                    )}
                  </div>
                  <div>
                    <div className="flex items-center gap-2">
                      <h3 className="font-display font-bold text-base text-slate-800">{selectedKaryawanDetail.nama_lengkap}</h3>
                      <span className="px-2.5 py-0.5 rounded-full bg-violet-50 text-violet-700 text-[10px] font-bold">
                        {selectedKaryawanDetail.divisi} {selectedKaryawanDetail.jabatan ? `• ${selectedKaryawanDetail.jabatan}` : ''}
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-400 font-medium mt-0.5">
                      Rincian presensi harian pada periode <span className="font-bold text-slate-600">{activePeriod.label}</span>
                    </p>
                  </div>
                </div>

                <button 
                  onClick={() => setSelectedKaryawanDetail(null)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-full text-xs font-bold transition flex items-center gap-1.5"
                >
                  <i className="fa-solid fa-arrow-left"></i>
                  <span>Kembali ke Rekap Semua</span>
                </button>
              </div>

              {/* Filter Status Presensi Detail */}
              <div className="flex flex-wrap gap-2 text-[10px] font-bold">
                {['Semua', 'Hadir', 'Terlambat', 'Libur', 'Sakit', 'Izin', 'Cuti', 'Lembur'].map(st => (
                  <button 
                    key={st}
                    onClick={() => setFilterDetailStatus(st)}
                    className={`px-3 py-1.5 rounded-full transition shadow-sm ${
                      filterDetailStatus === st 
                        ? 'bg-slate-900 text-white' 
                        : 'bg-slate-50 text-slate-600 border border-slate-200 hover:bg-white'
                    }`}
                  >
                    {st === 'Lembur' ? `Lembur (${detailLemburKaryawan.length})` : st}
                  </button>
                ))}
              </div>

              {/* Tabel Rincian Harian Karyawan / Lembur */}
              {filterDetailStatus === 'Lembur' ? (
                <div className="overflow-x-auto rounded-2xl border border-slate-100">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                        <th className="p-3 text-center">No</th>
                        <th className="p-3">Tanggal</th>
                        <th className="p-3 text-center">Durasi</th>
                        <th className="p-3">Keterangan Pekerjaan</th>
                        <th className="p-3 text-center">Status</th>
                        <th className="p-3 text-center">Foto Bukti</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {detailLemburKaryawan.length > 0 ? (
                        detailLemburKaryawan.map((lmb, idx) => (
                          <tr key={lmb.id} className="hover:bg-slate-50/50">
                            <td className="p-3 text-center font-mono-numbers text-slate-400 font-bold">{idx + 1}</td>
                            <td className="p-3 font-semibold text-slate-700 font-mono-numbers">{lmb.tanggal}</td>
                            <td className="p-3 text-center">
                              <span className="px-2.5 py-1 rounded-full bg-indigo-50 text-indigo-700 font-mono-numbers font-bold text-[10px]">
                                {lmb.durasi_jam} Jam
                              </span>
                            </td>
                            <td className="p-3">
                              <p className="text-slate-700 font-medium text-[11px] leading-relaxed max-w-sm">{lmb.keterangan}</p>
                            </td>
                            <td className="p-3 text-center">
                              <span className={`px-2.5 py-1 rounded-full text-[10px] font-bold ${
                                lmb.status === 'Approved' ? 'bg-emerald-50 text-emerald-600' :
                                lmb.status === 'Rejected' ? 'bg-rose-50 text-rose-600' :
                                'bg-amber-50 text-amber-600'
                              }`}>
                                {lmb.status === 'Approved' ? 'Disetujui' : lmb.status === 'Rejected' ? 'Ditolak' : 'Pending'}
                              </span>
                            </td>
                            <td className="p-3 text-center">
                              {lmb.bukti_url ? (
                                <button 
                                  onClick={() => setPreviewPhotoUrl({ url: getCleanImageUrl(lmb.bukti_url), label: `Bukti Lembur - ${selectedKaryawanDetail.nama_lengkap} (${lmb.tanggal})` })}
                                  className="inline-block relative group"
                                  title="Klik untuk memperbesar"
                                >
                                  <img 
                                    src={getCleanImageUrl(lmb.bukti_url)} 
                                    alt="Bukti Lembur" 
                                    className="w-8 h-8 rounded-lg object-cover border border-indigo-200 hover:scale-105 transition" 
                                  />
                                </button>
                              ) : (
                                <span className="text-[10px] text-slate-300">Tanpa Bukti</span>
                              )}
                            </td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan="6" className="p-8 text-center text-slate-400 font-medium">
                            Tidak ada riwayat pengajuan lembur pada periode ini.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="overflow-x-auto rounded-2xl border border-slate-100">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                        <th className="p-3 text-center">No</th>
                        <th className="p-3">Tanggal</th>
                        <th className="p-3">Tipe</th>
                        <th className="p-3">Jam</th>
                        <th className="p-3">Status</th>
                        <th className="p-3">Shift</th>
                        <th className="p-3">Lokasi / Keterangan</th>
                        <th className="p-3 text-center">Foto</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredDetailPresensi.length > 0 ? (
                        filteredDetailPresensi.map((p, idx) => (
                          <tr key={p.id} className="hover:bg-slate-50/50">
                            <td className="p-3 text-center font-mono-numbers text-slate-400 font-bold">{idx + 1}</td>
                            <td className="p-3 font-semibold text-slate-700 font-mono-numbers">{p.tanggal}</td>
                            <td className="p-3 uppercase font-bold text-[10px] text-slate-600">{p.tipe}</td>
                            <td className="p-3 font-mono-numbers font-bold text-slate-800">{p.jam}</td>
                            <td className="p-3">
                              <span className={`px-2.5 py-1 rounded-full text-[10px] font-bold ${
                                p.status === 'Hadir' ? 'bg-emerald-50 text-emerald-600' :
                                p.status === 'Terlambat' ? 'bg-amber-50 text-amber-600' :
                                p.status === 'Libur' ? 'bg-sky-50 text-sky-600' :
                                p.status === 'Izin' ? 'bg-orange-50 text-orange-600' :
                                p.status === 'Sakit' ? 'bg-amber-50 text-amber-600' :
                                p.status === 'Cuti' ? 'bg-purple-50 text-purple-600' :
                                'bg-slate-100 text-slate-600'
                              }`}>
                                {p.status}
                              </span>
                            </td>
                            <td className="p-3 text-slate-600">{p.shift_nama || '-'}</td>
                            <td className="p-3">
                              <span className="font-semibold text-slate-700 block">{p.lokasi_nama}</span>
                              <span className="text-[10px] text-slate-400">{p.keterangan || '-'}</span>
                            </td>
                            <td className="p-3 text-center">
                              {p.foto_url ? (
                                <button 
                                  onClick={() => setPreviewPhotoUrl({ url: getCleanImageUrl(p.foto_url), label: `${p.tipe.toUpperCase()} - ${selectedKaryawanDetail.nama_lengkap} (${p.tanggal})` })}
                                  className="inline-block"
                                >
                                  <img src={getCleanImageUrl(p.foto_url)} className="w-8 h-8 rounded-lg object-cover border border-slate-200 hover:scale-105 transition" />
                                </button>
                              ) : (
                                <span className="text-slate-300">-</span>
                              )}
                            </td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan="8" className="p-8 text-center text-slate-400 font-medium">
                            Tidak ada catatan presensi dengan status "{filterDetailStatus}" pada periode ini.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          ) : (
            /* Tampilan Tabel Ringkasan Periode Semua Karyawan */
            <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
              {/* Filter Pencarian Karyawan */}
              <div className="p-4 border-b border-slate-100 flex items-center justify-between gap-4">
                <div className="relative flex-1 max-w-sm">
                  <i className="fa-solid fa-magnifying-glass absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
                  <input 
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Cari nama karyawan, divisi, atau jabatan..."
                    className="w-full bg-slate-50 border border-slate-200 rounded-2xl pl-9 pr-4 py-2 text-xs font-semibold outline-none focus:border-violet-500"
                  />
                </div>

                <div className="text-[11px] font-bold text-slate-500">
                  Total: <span className="text-violet-600 font-mono-numbers">{filteredSummary.length}</span> Karyawan
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                      <th className="p-4 text-center">No</th>
                      <th className="p-4">Karyawan</th>
                      <th className="p-4 text-center">Masuk (Hadir)</th>
                      <th className="p-4 text-center">Terlambat</th>
                      <th className="p-4 text-center">Libur</th>
                      <th className="p-4 text-center">Sakit</th>
                      <th className="p-4 text-center">Izin</th>
                      <th className="p-4 text-center">Cuti</th>
                      <th className="p-4 text-center">Total Lembur</th>
                      <th className="p-4 text-center">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {filteredSummary.length > 0 ? (
                      filteredSummary.map((item, idx) => (
                        <tr 
                          key={item.karyawan.id} 
                          onClick={() => setSelectedKaryawanDetail(item.karyawan)}
                          className="hover:bg-violet-50/40 cursor-pointer transition"
                        >
                          <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{idx + 1}</td>
                          <td className="p-4">
                            <div className="font-bold text-slate-800">{item.karyawan.nama_lengkap}</div>
                            <div className="text-[10px] text-slate-400 font-medium">
                              {item.karyawan.divisi} {item.karyawan.jabatan ? `• ${item.karyawan.jabatan}` : ''}
                            </div>
                          </td>
                          <td className="p-4 text-center font-mono-numbers font-bold text-emerald-600">
                            {item.masuk} <span className="text-[10px] font-normal text-slate-400">hari</span>
                          </td>
                          <td className="p-4 text-center font-mono-numbers font-bold text-amber-600">
                            {item.terlambat}
                          </td>
                          <td className="p-4 text-center font-mono-numbers font-bold text-sky-600">
                            {item.libur}
                          </td>
                          <td className="p-4 text-center font-mono-numbers font-bold text-amber-600">
                            {item.sakit}
                          </td>
                          <td className="p-4 text-center font-mono-numbers font-bold text-orange-600">
                            {item.izin}
                          </td>
                          <td className="p-4 text-center font-mono-numbers font-bold text-purple-600">
                            {item.cuti}
                          </td>
                          <td className="p-4 text-center">
                            <span className="px-2.5 py-1 rounded-full bg-indigo-50 text-indigo-700 font-mono-numbers font-bold text-[10px]">
                              {item.totalJamLembur} Jam
                            </span>
                          </td>
                          <td className="p-4 text-center">
                            <button 
                              onClick={(e) => {
                                e.stopPropagation();
                                setSelectedKaryawanDetail(item.karyawan);
                              }}
                              className="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-violet-100 hover:text-violet-700 text-slate-600 text-[10px] font-bold transition flex items-center gap-1 mx-auto"
                            >
                              <span>Detail</span>
                              <i className="fa-solid fa-chevron-right text-[9px]"></i>
                            </button>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan="10" className="p-8 text-center text-slate-400">
                          Tidak ditemukan karyawan yang cocok dengan pencarian "{searchQuery}".
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      );
    }

    /* =========================================================
       MODUL PENGAJUAN LEMBUR ADMIN
       ========================================================= */
    function PengajuanLemburAdminView({ lemburList, setLemburList, showToast, setPreviewPhotoUrl, exportToExcel }) {
      const [filterStatus, setFilterStatus] = useState('Semua');
      const [searchQuery, setSearchQuery] = useState('');

      const stats = useMemo(() => {
        const pending = lemburList.filter(l => l.status === 'Pending').length;
        const approved = lemburList.filter(l => l.status === 'Approved').length;
        const rejected = lemburList.filter(l => l.status === 'Rejected').length;
        return { pending, approved, rejected, total: lemburList.length };
      }, [lemburList]);

      const filteredList = useMemo(() => {
        return lemburList.filter(item => {
          const matchStatus = filterStatus === 'Semua' || item.status === filterStatus;
          const matchSearch = !searchQuery.trim() || 
            item.karyawan_nama.toLowerCase().includes(searchQuery.toLowerCase()) ||
            (item.divisi && item.divisi.toLowerCase().includes(searchQuery.toLowerCase())) ||
            item.keterangan.toLowerCase().includes(searchQuery.toLowerCase());
          return matchStatus && matchSearch;
        });
      }, [lemburList, filterStatus, searchQuery]);

      const handleUpdateStatus = (id, newStatus) => {
        setLemburList(prev => prev.map(l => l.id === id ? { ...l, status: newStatus } : l));
        showToast(`Pengajuan lembur berhasil di-${newStatus === 'Approved' ? 'Setujui' : 'Tolak'}!`);
      };

      const handleDelete = async (id) => {
        if (await modernConfirm('Apakah Anda yakin ingin menghapus data pengajuan lembur ini?')) {
          setLemburList(prev => prev.filter(l => l.id !== id));
          showToast('Data pengajuan lembur dihapus.');
        }
      };

      const handleExport = () => {
        const rows = filteredList.map((item, idx) => ({
          'No': idx + 1,
          'Nama Karyawan': item.karyawan_nama,
          'Divisi': item.divisi || '-',
          'Tanggal': item.tanggal,
          'Durasi (Jam)': item.durasi_jam,
          'Keterangan': item.keterangan,
          'Status': item.status,
          'Waktu Pengajuan': new Date(item.created_at).toLocaleString('id-ID')
        }));
        exportToExcel(rows, `Data_Pengajuan_Lembur_${Date.now()}`);
      };

      return (
        <div className="space-y-6">
          {/* Header Panel */}
          <div className="flex flex-wrap justify-between items-center gap-4 bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
            <div>
              <div className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full bg-indigo-500"></span>
                <span className="text-[10px] font-bold uppercase tracking-widest text-indigo-600">Overtime Management</span>
              </div>
              <h2 className="font-display font-bold text-xl text-slate-800 mt-1">Pengajuan Lembur Karyawan</h2>
              <p className="text-xs text-slate-400 font-medium">Verifikasi, persetujuan, dan penolakan pengajuan lembur mandiri staf</p>
            </div>

            <button 
              onClick={handleExport}
              className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-full text-xs font-bold shadow-md active:scale-95 transition flex items-center gap-2"
            >
              <i className="fa-solid fa-file-excel"></i>
              <span>Unduh Excel</span>
            </button>
          </div>

          {/* 4 Kartu Metrik Interaktif */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
            <div 
              onClick={() => setFilterStatus('Semua')} 
              className={`p-4 rounded-3xl border cursor-pointer transition shadow-sm ${filterStatus === 'Semua' ? 'bg-slate-900 text-white border-slate-900' : 'bg-white text-slate-700 border-slate-100 hover:border-slate-300'}`}
            >
              <span className="text-[10px] font-bold uppercase tracking-wider block opacity-70">Total Pengajuan</span>
              <span className="font-mono-numbers text-2xl font-extrabold mt-1 block">{stats.total}</span>
            </div>

            <div 
              onClick={() => setFilterStatus('Pending')} 
              className={`p-4 rounded-3xl border cursor-pointer transition shadow-sm ${filterStatus === 'Pending' ? 'bg-amber-500 text-white border-amber-500' : 'bg-white text-slate-700 border-slate-100 hover:border-amber-200'}`}
            >
              <span className="text-[10px] font-bold uppercase tracking-wider block opacity-70">Menunggu (Pending)</span>
              <span className="font-mono-numbers text-2xl font-extrabold text-amber-500 mt-1 block group-hover:text-amber-600" style={{ color: filterStatus === 'Pending' ? 'white' : undefined }}>{stats.pending}</span>
            </div>

            <div 
              onClick={() => setFilterStatus('Approved')} 
              className={`p-4 rounded-3xl border cursor-pointer transition shadow-sm ${filterStatus === 'Approved' ? 'bg-emerald-600 text-white border-emerald-600' : 'bg-white text-slate-700 border-slate-100 hover:border-emerald-200'}`}
            >
              <span className="text-[10px] font-bold uppercase tracking-wider block opacity-70">Disetujui</span>
              <span className="font-mono-numbers text-2xl font-extrabold text-emerald-600 mt-1 block" style={{ color: filterStatus === 'Approved' ? 'white' : undefined }}>{stats.approved}</span>
            </div>

            <div 
              onClick={() => setFilterStatus('Rejected')} 
              className={`p-4 rounded-3xl border cursor-pointer transition shadow-sm ${filterStatus === 'Rejected' ? 'bg-rose-600 text-white border-rose-600' : 'bg-white text-slate-700 border-slate-100 hover:border-rose-200'}`}
            >
              <span className="text-[10px] font-bold uppercase tracking-wider block opacity-70">Ditolak</span>
              <span className="font-mono-numbers text-2xl font-extrabold text-rose-600 mt-1 block" style={{ color: filterStatus === 'Rejected' ? 'white' : undefined }}>{stats.rejected}</span>
            </div>
          </div>

          {/* Tabel Pengajuan Lembur */}
          <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between gap-4">
              <div className="relative flex-1 max-w-sm">
                <i className="fa-solid fa-magnifying-glass absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
                <input 
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Cari nama karyawan atau keterangan..."
                  className="w-full bg-slate-50 border border-slate-200 rounded-2xl pl-9 pr-4 py-2 text-xs font-semibold outline-none focus:border-indigo-500"
                />
              </div>

              <div className="text-[11px] font-bold text-slate-400">
                Menampilkan <span className="font-mono-numbers text-slate-700">{filteredList.length}</span> pengajuan
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                    <th className="p-4 text-center">No</th>
                    <th className="p-4">Karyawan</th>
                    <th className="p-4">Tanggal Lembur</th>
                    <th className="p-4">Durasi</th>
                    <th className="p-4">Keterangan Pekerjaan</th>
                    <th className="p-4 text-center">Bukti Foto</th>
                    <th className="p-4 text-center">Status</th>
                    <th className="p-4 text-center">Aksi Verifikasi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {filteredList.length > 0 ? (
                    filteredList.map((item, idx) => (
                      <tr key={item.id} className="hover:bg-slate-50/50 transition">
                        <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{idx + 1}</td>
                        <td className="p-4">
                          <div className="font-bold text-slate-800">{item.karyawan_nama}</div>
                          <div className="text-[10px] text-slate-400 font-medium">{item.divisi || 'Staff'}</div>
                        </td>
                        <td className="p-4 font-mono-numbers font-semibold text-slate-700">{item.tanggal}</td>
                        <td className="p-4">
                          <span className="px-2.5 py-1 rounded-full bg-indigo-50 text-indigo-700 font-mono-numbers font-bold text-[10px]">
                            {item.durasi_jam} Jam
                          </span>
                        </td>
                        <td className="p-4 max-w-xs">
                          <p className="text-slate-700 font-medium text-[11px] leading-relaxed line-clamp-2" title={item.keterangan}>
                            {item.keterangan}
                          </p>
                        </td>
                        <td className="p-4 text-center">
                          {item.bukti_url ? (
                            <button 
                              onClick={() => setPreviewPhotoUrl({ url: getCleanImageUrl(item.bukti_url), label: `Bukti Lembur - ${item.karyawan_nama} (${item.tanggal})` })}
                              className="inline-block relative group"
                              title="Klik untuk melihat bukti foto"
                            >
                              <img 
                                src={getCleanImageUrl(item.bukti_url)} 
                                alt="Bukti Lembur" 
                                className="w-10 h-10 rounded-xl object-cover border border-indigo-200 shadow-sm group-hover:scale-105 transition" 
                              />
                              <span className="absolute -bottom-1 -right-1 w-4 h-4 rounded-full bg-indigo-600 text-white text-[8px] flex items-center justify-center shadow">
                                <i className="fa-solid fa-magnifying-glass"></i>
                              </span>
                            </button>
                          ) : (
                            <span className="text-[10px] text-slate-300 font-medium">Tanpa Bukti</span>
                          )}
                        </td>
                        <td className="p-4 text-center">
                          <span className={`px-2.5 py-1 rounded-full text-[10px] font-bold ${
                            item.status === 'Approved' ? 'bg-emerald-50 text-emerald-600' :
                            item.status === 'Rejected' ? 'bg-rose-50 text-rose-600' :
                            'bg-amber-50 text-amber-600'
                          }`}>
                            {item.status === 'Approved' ? 'Disetujui' : item.status === 'Rejected' ? 'Ditolak' : 'Menunggu'}
                          </span>
                        </td>
                        <td className="p-4 text-center">
                          <div className="flex items-center justify-center space-x-1.5">
                            {item.status === 'Pending' ? (
                              <>
                                <button 
                                  onClick={() => handleUpdateStatus(item.id, 'Approved')}
                                  className="w-8 h-8 rounded-full bg-emerald-50 hover:bg-emerald-100 text-emerald-600 flex items-center justify-center transition active:scale-95"
                                  title="Setujui Lembur"
                                >
                                  <i className="fa-solid fa-check text-xs"></i>
                                </button>
                                <button 
                                  onClick={() => handleUpdateStatus(item.id, 'Rejected')}
                                  className="w-8 h-8 rounded-full bg-rose-50 hover:bg-rose-100 text-rose-600 flex items-center justify-center transition active:scale-95"
                                  title="Tolak Lembur"
                                >
                                  <i className="fa-solid fa-xmark text-xs"></i>
                                </button>
                              </>
                            ) : (
                              <button 
                                onClick={() => handleUpdateStatus(item.id, 'Pending')}
                                className="px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-600 text-[10px] font-bold transition"
                                title="Reset status kembali ke Menunggu"
                              >
                                Reset
                              </button>
                            )}

                            <button 
                              onClick={() => handleDelete(item.id)}
                              className="w-8 h-8 rounded-full bg-slate-50 hover:bg-rose-50 text-slate-400 hover:text-rose-600 flex items-center justify-center transition active:scale-95"
                              title="Hapus Pengajuan"
                            >
                              <i className="fa-solid fa-trash text-xs"></i>
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="8" className="p-8 text-center text-slate-400 font-medium">
                        Tidak ada data pengajuan lembur dengan status "{filterStatus}".
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      );
    }

    // Modern Minimalist Master Karyawan View
    function MasterKaryawanView({ karyawanList = [], setKaryawanList, divisiList = [], jabatanList = [], lokasiList = [], pendaftaranKaryawanAktif = false, setPendaftaranKaryawanAktif, showToast, setPreviewPhotoUrl }) {
      // Normalisasi defensif agar data Supabase yang null/bukan array tidak membuat halaman blank.
      karyawanList = Array.isArray(karyawanList) ? karyawanList : [];
      divisiList = Array.isArray(divisiList) ? divisiList : [];
      jabatanList = Array.isArray(jabatanList) ? jabatanList : [];
      lokasiList = Array.isArray(lokasiList) ? lokasiList : [];
      const [showModal, setShowModal] = useState(false);
      const [editingId, setEditingId] = useState(null);
      const [savingPendaftaran, setSavingPendaftaran] = useState(false);
      const handleTogglePendaftaran = async (enabled) => {
        if (!setPendaftaranKaryawanAktif) return;
        setSavingPendaftaran(true);
        try {
          await setPendaftaranKaryawanAktif(enabled);
          showToast(enabled ? 'Pendaftaran karyawan berhasil diaktifkan.' : 'Pendaftaran karyawan berhasil dinonaktifkan.');
        } catch (error) {
          showToast(`Pengaturan pendaftaran gagal disimpan: ${error.message || 'periksa izin tabel settings.'}`, 'error');
        } finally { setSavingPendaftaran(false); }
      };
      const [form, setForm] = useState({ 
        nama_lengkap: '', 
        nik: '',
        lokasi_absen: [],
        tgl_lahir: '', 
        username: '', 
        password: '', 
        no_wa: '', 
        email: '', 
        divisi: divisiList[0]?.nama_divisi || '', 
        jabatan: '', 
        provinsi: '',
        kota: '',
        kecamatan: '',
        desa: '',
        jalan: '',
        foto_profil: ''
      });

      // Dropdown wilayah administratif Indonesia (Provinsi > Kabupaten/Kota > Kecamatan > Desa/Kelurahan)
      const WILAYAH_API = 'https://www.emsifa.com/api-wilayah-indonesia/api';
      const [provinceOptions, setProvinceOptions] = useState([]);
      const [regencyOptions, setRegencyOptions] = useState([]);
      const [districtOptions, setDistrictOptions] = useState([]);
      const [villageOptions, setVillageOptions] = useState([]);
      const [provinceId, setProvinceId] = useState('');
      const [regencyId, setRegencyId] = useState('');
      const [districtId, setDistrictId] = useState('');
      const [villageId, setVillageId] = useState('');
      const [wilayahLoading, setWilayahLoading] = useState(false);
      const [wilayahError, setWilayahError] = useState('');

      const fetchWilayah = async (path) => {
        const response = await fetch(`${WILAYAH_API}/${path}.json`);
        if (!response.ok) throw new Error('Data wilayah Indonesia tidak dapat dimuat. Periksa koneksi internet.');
        return response.json();
      };
      const toDisplayDate = (value) => {
        if (!value) return '';
        const raw = String(value).trim();
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) return raw;
        const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
        return match ? `${match[3]}/${match[2]}/${match[1]}` : raw;
      };
      const toIsoDate = (value) => {
        const match = String(value || '').trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
        if (!match) return '';
        const [, dd, mm, yyyy] = match;
        const date = new Date(Number(yyyy), Number(mm) - 1, Number(dd));
        if (date.getFullYear() !== Number(yyyy) || date.getMonth() !== Number(mm) - 1 || date.getDate() !== Number(dd)) return '';
        return `${yyyy}-${mm}-${dd}`;
      };
      const formatDateInput = (value) => {
        const digits = String(value || '').replace(/\D/g, '').slice(0, 8);
        if (digits.length <= 2) return digits;
        if (digits.length <= 4) return `${digits.slice(0, 2)}/${digits.slice(2)}`;
        return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
      };

      useEffect(() => {
        let active = true;
        fetchWilayah('provinces').then(items => { if (active) setProvinceOptions(items); })
          .catch(err => { if (active) setWilayahError(err.message || 'Gagal memuat data wilayah.'); });
        return () => { active = false; };
      }, []);

      // Urutkan A-Z dan tandai karyawan yang datanya belum lengkap.
      // Field opsional yang tidak memengaruhi status: NIK, foto profil, dan jalan/alamat detail.
      const getAlamatKaryawan = (k) => {
        try { return typeof k?.alamat === 'string' ? JSON.parse(k.alamat || '{}') : (k?.alamat || {}); }
        catch (_) { return {}; }
      };
      const isKaryawanDataIncomplete = (k) => {
        const alamatKaryawan = getAlamatKaryawan(k);
        const requiredValues = [
          k?.nama_lengkap, k?.tgl_lahir, k?.username, k?.password,
          k?.no_wa, k?.email, k?.divisi, k?.jabatan,
          alamatKaryawan?.provinsi, alamatKaryawan?.kota,
          alamatKaryawan?.kecamatan, alamatKaryawan?.desa || alamatKaryawan?.kelurahan
        ];
        return requiredValues.some(value => value === null || value === undefined || String(value).trim() === '');
      };
      const sortedKaryawanList = [...(Array.isArray(karyawanList) ? karyawanList : [])].filter(Boolean).sort((a, b) => String(a?.nama_lengkap || '').localeCompare(String(b?.nama_lengkap || ''), 'id'));

      const availableJabatan = useMemo(() => {
        if (!form.divisi) return jabatanList;
        return jabatanList.filter(j => {
          const divisiJabatan = Array.isArray(j.divisi)
            ? j.divisi
            : String(j.divisi || '').split('|').map(value => value.trim()).filter(Boolean);
          return divisiJabatan.includes(form.divisi);
        });
      }, [form.divisi, jabatanList]);

      const openModal = async (k = null) => {
        setWilayahError('');
        setWilayahLoading(true);
        setRegencyOptions([]); setDistrictOptions([]); setVillageOptions([]);
        setProvinceId(''); setRegencyId(''); setDistrictId(''); setVillageId('');
        if (k) {
          setEditingId(k.id);
          let alm = {};
          try { alm = typeof k.alamat === 'string' ? JSON.parse(k.alamat) : (k.alamat || {}); } catch (e) { alm = { jalan: k.alamat || '' }; }
          setForm({
            nama_lengkap: k.nama_lengkap || '', nik: String(k.nik || ''), lokasi_absen: Array.isArray(k.lokasi_absen) ? (k.lokasi_absen.map(String).includes('*') && !k.lokasi_absen.map(String).some(id => id !== '*') ? ['*', ...lokasiList.map(l => String(l.id))] : k.lokasi_absen.map(String)) : lokasiList.map(l => String(l.id)), tgl_lahir: toDisplayDate(k.tgl_lahir), username: k.username || '',
            password: k.password || '', no_wa: k.no_wa || '', email: k.email || '', divisi: k.divisi || '',
            jabatan: k.jabatan || '', provinsi: alm.provinsi || '', kota: alm.kota || '',
            kecamatan: alm.kecamatan || '', desa: alm.desa || alm.kelurahan || '', jalan: alm.jalan || '',
            foto_profil: k.foto_profil || ''
          });
          setShowModal(true);
          try {
            const provinces = provinceOptions.length ? provinceOptions : await fetchWilayah('provinces');
            setProvinceOptions(provinces);
            const province = provinces.find(x => x.name.toLowerCase() === String(alm.provinsi || '').toLowerCase());
            if (province) {
              setProvinceId(province.id);
              const regencies = await fetchWilayah(`regencies/${province.id}`); setRegencyOptions(regencies);
              const regency = regencies.find(x => x.name.toLowerCase() === String(alm.kota || '').toLowerCase());
              if (regency) {
                setRegencyId(regency.id);
                const districts = await fetchWilayah(`districts/${regency.id}`); setDistrictOptions(districts);
                const district = districts.find(x => x.name.toLowerCase() === String(alm.kecamatan || '').toLowerCase());
                if (district) {
                  setDistrictId(district.id);
                  const villages = await fetchWilayah(`villages/${district.id}`); setVillageOptions(villages);
                  const villageName = alm.desa || alm.kelurahan || '';
                  const village = villages.find(x => x.name.toLowerCase() === String(villageName).toLowerCase());
                  if (village) setVillageId(village.id);
                }
              }
            }
          } catch (err) { setWilayahError(err.message || 'Gagal memuat data wilayah.'); }
        } else {
          setEditingId(null);
          setForm({ nama_lengkap: '', nik: '', lokasi_absen: lokasiList.map(l => String(l.id)), tgl_lahir: '', username: '', password: '', no_wa: '', email: '', divisi: divisiList[0]?.nama_divisi || '', jabatan: '', provinsi: '', kota: '', kecamatan: '', desa: '', jalan: '', foto_profil: '' });
          setShowModal(true);
          if (!provinceOptions.length) {
            try { setProvinceOptions(await fetchWilayah('provinces')); }
            catch (err) { setWilayahError(err.message || 'Gagal memuat data wilayah.'); }
          }
        }
        setWilayahLoading(false);
      };

      const handleProvinceChange = async (id) => {
        const selected = provinceOptions.find(x => x.id === id);
        setProvinceId(id); setRegencyId(''); setDistrictId(''); setVillageId('');
        setRegencyOptions([]); setDistrictOptions([]); setVillageOptions([]);
        setForm(prev => ({ ...prev, provinsi: selected?.name || '', kota: '', kecamatan: '', desa: '' }));
        if (!id) return;
        setWilayahLoading(true); setWilayahError('');
        try { setRegencyOptions(await fetchWilayah(`regencies/${id}`)); }
        catch (err) { setWilayahError(err.message || 'Gagal memuat kabupaten/kota.'); }
        finally { setWilayahLoading(false); }
      };
      const handleRegencyChange = async (id) => {
        const selected = regencyOptions.find(x => x.id === id);
        setRegencyId(id); setDistrictId(''); setVillageId(''); setDistrictOptions([]); setVillageOptions([]);
        setForm(prev => ({ ...prev, kota: selected?.name || '', kecamatan: '', desa: '' }));
        if (!id) return;
        setWilayahLoading(true); setWilayahError('');
        try { setDistrictOptions(await fetchWilayah(`districts/${id}`)); }
        catch (err) { setWilayahError(err.message || 'Gagal memuat kecamatan.'); }
        finally { setWilayahLoading(false); }
      };
      const handleDistrictChange = async (id) => {
        const selected = districtOptions.find(x => x.id === id);
        setDistrictId(id); setVillageId(''); setVillageOptions([]);
        setForm(prev => ({ ...prev, kecamatan: selected?.name || '', desa: '' }));
        if (!id) return;
        setWilayahLoading(true); setWilayahError('');
        try { setVillageOptions(await fetchWilayah(`villages/${id}`)); }
        catch (err) { setWilayahError(err.message || 'Gagal memuat desa/kelurahan.'); }
        finally { setWilayahLoading(false); }
      };
      const handleVillageChange = (id) => {
        const selected = villageOptions.find(x => x.id === id);
        setVillageId(id); setForm(prev => ({ ...prev, desa: selected?.name || '' }));
      };

      const handleSave = async (e) => {
        e.preventDefault();
        const noWa = String(form.no_wa || '').trim();
        if (!/^08\d{8,13}$/.test(noWa)) {
          showToast('Nomor WhatsApp wajib diawali 08 dan berisi minimal 10 digit (maksimal 15 digit).', 'error');
          return;
        }
        const nik = String(form.nik || '').trim();
        const tglLahirIso = toIsoDate(form.tgl_lahir);
        if (!tglLahirIso) {
          showToast('Tanggal lahir harus valid dengan format dd/mm/yyyy, contoh 31/12/1998.', 'error');
          return;
        }
        if (!form.provinsi || !form.kota || !form.kecamatan || !form.desa) {
          showToast('Pilih provinsi, kabupaten/kota, kecamatan, dan desa/kelurahan.', 'error');
          return;
        }
        try {
          let fotoProfilUrl = form.foto_profil || '';
          if (fotoProfilUrl.startsWith('data:image')) {
            fotoProfilUrl = await uploadToGoogleDriveViaGAS(
              fotoProfilUrl,
              `profil_${editingId || ('kry-' + Date.now())}_${Date.now()}.jpg`,
              'profil'
            );
            if (!fotoProfilUrl || fotoProfilUrl.startsWith('data:image')) {
              throw new Error('Foto profil gagal diunggah ke Google Drive. Periksa konfigurasi Google Apps Script.');
            }
          }
          const alamat = JSON.stringify({ provinsi: form.provinsi, kota: form.kota, kecamatan: form.kecamatan, desa: form.desa, jalan: form.jalan });
          const cleanForm = { ...form, nik: nik || null, lokasi_absen: (form.lokasi_absen || []).map(String), tgl_lahir: tglLahirIso, no_wa: noWa, foto_profil: fotoProfilUrl };
          if(editingId) {
            setKaryawanList(karyawanList.map(k => k.id === editingId ? { ...k, ...cleanForm, alamat } : k));
          } else {
            setKaryawanList([...karyawanList, {
              id: 'kry-' + Date.now(),
              ...cleanForm,
              alamat,
              shifts: ['Pagi (08:00 - 17:00)'],
              sisa_libur: 4,
              sisa_cuti: 0
            }]);
          }
          setShowModal(false);
          showToast(editingId ? 'Data karyawan diperbarui!' : 'Karyawan baru ditambahkan!');
        } catch (err) {
          console.error('Gagal menyimpan karyawan/foto:', err);
          showToast(`Gagal menyimpan karyawan: ${err.message || 'kesalahan tidak diketahui'}`, 'error');
        }
      };

      return (
        <div className="space-y-6">
          <div className="flex justify-between items-center bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
            <div>
              <h2 className="font-display font-bold text-xl text-slate-800">Master Karyawan</h2>
              <p className="text-xs text-slate-400 font-medium mt-1">Kelola data pegawai, foto profil & akun akses</p>
            </div>
            <button onClick={() => openModal()} className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full text-xs font-bold shadow-md active:scale-95 transition">
              <i className="fa-solid fa-plus mr-2"></i> Tambah Karyawan
            </button>
          </div>
          <div className="bg-white p-5 rounded-[2rem] border border-slate-100 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-start gap-3">
              <div className={`w-11 h-11 rounded-2xl flex items-center justify-center shrink-0 ${pendaftaranKaryawanAktif ? 'bg-emerald-50 text-emerald-600' : 'bg-amber-50 text-amber-600'}`}><i className={`fa-solid ${pendaftaranKaryawanAktif ? 'fa-user-check' : 'fa-user-lock'} text-lg`}></i></div>
              <div>
                <h3 className="font-display font-bold text-base text-slate-800">Pendaftaran Karyawan Mandiri</h3>
                <p className="text-xs text-slate-500 mt-1">Atur apakah karyawan dapat membuka formulir pendaftaran dari halaman login.</p>
                <p className={`text-[11px] font-bold mt-2 ${pendaftaranKaryawanAktif ? 'text-emerald-600' : 'text-amber-600'}`}>{pendaftaranKaryawanAktif ? 'Status: AKTIF — formulir dapat dibuka' : 'Status: NONAKTIF — karyawan diminta menghubungi HRD'}</p>
              </div>
            </div>
            <button type="button" disabled={savingPendaftaran} onClick={() => handleTogglePendaftaran(!pendaftaranKaryawanAktif)} className={`px-5 py-3 rounded-2xl text-xs font-bold transition disabled:opacity-50 whitespace-nowrap ${pendaftaranKaryawanAktif ? 'bg-rose-50 text-rose-700 border border-rose-200 hover:bg-rose-100' : 'bg-violet-600 text-white hover:bg-violet-700'}`}>
              <i className={`fa-solid ${savingPendaftaran ? 'fa-spinner fa-spin' : pendaftaranKaryawanAktif ? 'fa-toggle-on' : 'fa-toggle-off'} mr-2`}></i>{savingPendaftaran ? 'Menyimpan...' : pendaftaranKaryawanAktif ? 'Nonaktifkan Pendaftaran' : 'Aktifkan Pendaftaran'}
            </button>
          </div>
          <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                  <th className="p-4 text-center">No</th>
                  <th className="p-4 text-center">Foto Profil</th>
                  <th className="p-4">Divisi</th>
                  <th className="p-4">Username</th>
                  <th className="p-4">Divisi & Jabatan</th>
                  <th className="p-4 text-center">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {sortedKaryawanList.map((k, i) => (
                  <tr key={k.id} title={isKaryawanDataIncomplete(k) ? 'Data karyawan belum lengkap. Lengkapi data wajib melalui tombol Edit.' : 'Data karyawan lengkap'} className={`transition ${isKaryawanDataIncomplete(k) ? 'bg-rose-500/[0.09] backdrop-blur-sm border-l-4 border-rose-400 hover:bg-rose-500/[0.15]' : 'hover:bg-slate-50/50'}`}>
                    <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{i + 1}</td>
                    
                    {/* Kolom Preview Foto Profil di antara Nomor dan Nama Lengkap */}
                    <td className="p-4 text-center">
                      <div className="flex items-center justify-center">
                        {k.foto_profil ? (
                          <button
                            type="button"
                            onClick={() => setPreviewPhotoUrl({ url: getCleanImageUrl(k.foto_profil), label: `Foto Profil - ${k.nama_lengkap || 'Nama belum diisi'}` })}
                            className="relative group inline-block"
                            title="Klik untuk memperbesar foto"
                          >
                            <img 
                              src={getCleanImageUrl(k.foto_profil)} 
                              alt={k.nama_lengkap || 'Nama belum diisi'} 
                              className="w-10 h-10 rounded-full object-cover border-2 border-violet-200 shadow-sm group-hover:scale-105 group-hover:border-violet-500 transition"
                              onError={(e) => {
                                e.target.onerror = null;
                                e.target.parentElement.innerHTML = `<div class="w-10 h-10 rounded-full bg-violet-100 text-violet-700 font-bold text-xs flex items-center justify-center border border-violet-200">${String(k.nama_lengkap || '?').split(' ').filter(Boolean).map(n=>n[0]).slice(0, 2).join('')}</div>`;
                              }}
                            />
                            <span className="absolute -bottom-1 -right-1 w-4 h-4 rounded-full bg-violet-600 text-white text-[8px] flex items-center justify-center shadow">
                              <i className="fa-solid fa-magnifying-glass"></i>
                            </span>
                          </button>
                        ) : (
                          <div className="w-10 h-10 rounded-full bg-slate-100 text-slate-500 font-bold text-xs flex items-center justify-center border border-slate-200 shadow-inner" title="Belum unggah foto profil">
                            {String(k.nama_lengkap || '?').split(' ').filter(Boolean).map(n => n[0]).slice(0, 2).join('')}
                          </div>
                        )}
                      </div>
                    </td>

                    <td className="p-4">
                      <div className="font-bold text-slate-800">{k.nama_lengkap || 'Nama belum diisi'}</div>
                      <div className="text-[10px] text-slate-400 font-medium">{k.email || k.no_wa}</div>
                    </td>
                    <td className="p-4 font-mono-numbers text-slate-500">{k.username}</td>
                    <td className="p-4">
                      <div className="flex flex-col gap-1 items-start">
                        <span className="px-2.5 py-0.5 rounded-full bg-slate-100 text-slate-600 text-[10px] font-bold">{k.divisi}</span>
                        {k.jabatan && <span className="px-2.5 py-0.5 rounded-full bg-indigo-50 text-indigo-600 text-[9px] font-bold">{k.jabatan}</span>}
                      </div>
                    </td>
                    <td className="p-4 text-center">
                      <button onClick={()=>openModal(k)} className="w-8 h-8 rounded-full bg-slate-50 text-indigo-500 hover:bg-indigo-50 mx-1" title="Edit Data Karyawan"><i className="fa-solid fa-pen-to-square"></i></button>
                      <button onClick={async()=>{if(!await modernConfirm(`Yakin ingin menghapus data karyawan ${k.nama_lengkap || k.username || ''}? Tindakan ini akan menghapus data karyawan dari master.`, 'Hapus karyawan')) return; setKaryawanList((Array.isArray(karyawanList) ? karyawanList : []).filter(x=>String(x.id)!==String(k.id))); showToast('Permintaan penghapusan karyawan diproses.');}} className="w-8 h-8 rounded-full bg-slate-50 text-rose-500 hover:bg-rose-50 mx-1" title="Hapus Karyawan"><i className="fa-solid fa-trash"></i></button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {/* Form Modal Karyawan */}
          {showModal && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-md w-full p-6 shadow-2xl border border-slate-100">
                 <h3 className="font-display font-bold text-lg text-slate-800 mb-4">{editingId ? 'Edit Karyawan' : 'Tambah Karyawan'}</h3>
                 <form onSubmit={handleSave} className="space-y-3">
                   <div className="max-h-[65vh] overflow-y-auto space-y-3 pr-2">
                     {/* Preview & Ganti Foto Profil di Form */}
                     <div className="flex items-center space-x-3 bg-slate-50 border border-slate-200 rounded-2xl p-3">
                       <div className="w-12 h-12 rounded-full bg-white border border-slate-200 overflow-hidden flex items-center justify-center shrink-0 shadow-sm">
                         {form.foto_profil ? (
                           <img src={getCleanImageUrl(form.foto_profil)} alt="Preview" className="w-full h-full object-cover" />
                         ) : (
                           <i className="fa-solid fa-user text-slate-300 text-lg"></i>
                         )}
                       </div>
                       <div className="flex-1 min-w-0">
                         <label className="text-[10px] font-bold text-slate-500 block mb-1">Unggah / Ganti Foto Profil</label>
                         <input 
                           type="file" 
                           accept="image/*" 
                           onChange={(e) => {
                             const reader = new FileReader();
                             reader.onloadend = () => setForm({ ...form, foto_profil: reader.result });
                             if (e.target.files[0]) reader.readAsDataURL(e.target.files[0]);
                           }} 
                           className="w-full text-[11px] text-slate-500 file:mr-2 file:py-1 file:px-3 file:rounded-full file:border-0 file:text-[10px] file:font-bold file:bg-violet-100 file:text-violet-700 cursor-pointer" 
                         />
                       </div>
                     </div>

                     <input type="text" placeholder="Nama Lengkap" value={form.nama_lengkap} onChange={e=>setForm({...form, nama_lengkap:e.target.value})} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none" required />
                     <div>
                       <label className="text-[10px] font-bold text-slate-400 ml-1 block mb-1">Nomor Induk Karyawan</label>
                       <input type="text" inputMode="numeric" autoComplete="off" placeholder="Nomor induk karyawan (opsional)" value={form.nik || ''} onChange={e=>setForm({...form, nik:e.target.value.slice(0,40)})} maxLength={40} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none font-mono-numbers" />
                       <p className="text-[10px] text-slate-400 mt-1 ml-1">Nomor induk karyawan bersifat opsional; panjangnya bebas.</p>
                     </div>
                     
                     <div className="grid grid-cols-2 gap-2">
                       <div>
                         <label className="text-[10px] font-bold text-slate-400 ml-1 block mb-1">Tanggal Lahir</label>
                         <input type="text" inputMode="numeric" placeholder="dd/mm/yyyy" maxLength={10} value={form.tgl_lahir} onChange={e=>setForm({...form, tgl_lahir:formatDateInput(e.target.value)})} pattern="\d{2}/\d{2}/\d{4}" title="Gunakan format dd/mm/yyyy, contoh 31/12/1998" className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none font-mono-numbers" required />
                       </div>
                       <div>
                         <label className="text-[10px] font-bold text-slate-400 ml-1 block mb-1">No. WhatsApp</label>
                         <input type="tel" inputMode="numeric" placeholder="08xxxxxxxxxx" value={form.no_wa} onChange={e=>setForm({...form, no_wa:e.target.value.replace(/\D/g, "").slice(0, 15)})} pattern="08[0-9]{8,13}" minLength={10} maxLength={15} title="Nomor WhatsApp wajib diawali 08 dan minimal 10 digit" className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none font-mono-numbers" required />
                       </div>
                     </div>

                     <div className="grid grid-cols-2 gap-2">
                       <input type="text" placeholder="Username" value={form.username} onChange={e=>setForm({...form, username:e.target.value})} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none" required />
                       <input type="password" placeholder="Password Login" value={form.password} onChange={e=>setForm({...form, password:e.target.value})} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none" required />
                     </div>
                     <input type="email" placeholder="Alamat Email" value={form.email} onChange={e=>setForm({...form, email:e.target.value})} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none" required />

                     <div className="bg-indigo-50/70 border border-indigo-100 rounded-2xl p-3 space-y-3">
                       <div className="flex items-center justify-between gap-3">
                         <div><label className="text-[11px] font-extrabold text-indigo-800 block">Master Lokasi Absen</label><span className="text-[9px] text-indigo-500">Atur pembatasan lokasi per karyawan</span></div>
                         <label className="flex items-center gap-2 cursor-pointer shrink-0"><span className="text-[10px] font-bold text-slate-700">Unlock Lokasi</span><input type="checkbox" checked={(form.lokasi_absen || []).map(String).includes('*')} onChange={e=>setForm(prev=>{const ids=(prev.lokasi_absen||[]).map(String).filter(id=>id!=='*'); return {...prev, lokasi_absen:e.target.checked ? [...new Set(['*', ...ids])] : (ids.length ? ids : lokasiList.map(l=>String(l.id)))}})} className="h-4 w-4 accent-indigo-600" /></label>
                       </div>
                       {(form.lokasi_absen || []).map(String).includes('*') && <div className="rounded-xl border border-emerald-200 bg-emerald-50 p-3 text-[10px] leading-relaxed text-emerald-800"><b>Unlock Lokasi aktif.</b> Karyawan tetap wajib memilih salah satu lokasi yang dicentang dan peta/GPS tetap tampil, tetapi boleh melanjutkan presensi dari luar radius kantor. Lokasi yang dicentang di bawah tetap menjadi pilihan karyawan.</div>}
                       {lokasiList.length ? (
                         <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                           {[...lokasiList].sort((a,b)=>String(a.nama_lokasi||'').localeCompare(String(b.nama_lokasi||''),'id')).map(loc => {
                             const checked = (form.lokasi_absen || []).map(String).includes(String(loc.id));
                             return <label key={loc.id} className={`flex items-start gap-2 p-2 rounded-xl border cursor-pointer ${checked ? 'bg-white border-indigo-300' : 'bg-white/60 border-slate-200'}`}>
                               <input type="checkbox" checked={checked} onChange={e=>setForm(prev=>({...prev, lokasi_absen: e.target.checked ? [...new Set([...(prev.lokasi_absen||[]).map(String), String(loc.id)])] : (prev.lokasi_absen||[]).map(String).filter(id=>id!==String(loc.id))}))} className="mt-0.5 accent-indigo-600" />
                               <span className="text-[11px] font-bold text-slate-700">{loc.nama_lokasi}<span className="block text-[9px] font-medium text-slate-400">Radius {loc.radius_meter} m</span></span>
                             </label>;
                           })}
                         </div>
                       ) : <p className="text-[10px] text-amber-600">Belum ada lokasi. Tambahkan lokasi terlebih dahulu di menu Master Lokasi.</p>}
                       <p className="text-[9px] text-slate-500">Lokasi yang dicentang menjadi pilihan lokasi absen karyawan ini. Unlock Lokasi hanya menonaktifkan kewajiban berada dalam radius.</p>
                     </div>

                     <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                       <div>
                         <label className="text-[10px] font-bold text-slate-400 block mb-1">Penempatan Divisi</label>
                         <select value={form.divisi} onChange={e=>setForm({...form, divisi:e.target.value, jabatan: ''})} className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold text-slate-700 outline-none">
                           {divisiList.map(d => (
                             <option key={d.id} value={d.nama_divisi}>{d.nama_divisi}</option>
                           ))}
                         </select>
                       </div>

                       <div>
                         <label className="text-[10px] font-bold text-slate-400 block mb-1">Posisi / Jabatan</label>
                         <select value={form.jabatan} onChange={e=>setForm({...form, jabatan:e.target.value})} className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold text-slate-700 outline-none">
                           <option value="">Pilih Posisi / Jabatan...</option>
                           {availableJabatan.map(j => (
                             <option key={j.id} value={j.nama_jabatan}>{j.nama_jabatan}</option>
                           ))}
                         </select>
                       </div>
                       
                       {/* Indikator Shift Terhubung Real-Time */}
                       {form.divisi && divisiList.find(d => d.nama_divisi === form.divisi) && (
                         <div className="bg-violet-50 p-2.5 rounded-xl border border-violet-100 flex items-start space-x-2">
                           <i className="fa-solid fa-clock text-violet-500 mt-0.5 text-xs"></i>
                           <div>
                             <span className="text-[10px] font-bold text-violet-700 block leading-tight">Shift Terhubung:</span>
                             <span className="text-[10px] font-medium text-violet-600">
                               {(Array.isArray(divisiList.find(d => d.nama_divisi === form.divisi)?.shifts) ? divisiList.find(d => d.nama_divisi === form.divisi).shifts : []).join(', ') || 'Belum ada shift terkait'}
                             </span>
                           </div>
                         </div>
                       )}
                     </div>

                     <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                       <label className="text-[10px] font-bold text-slate-400 block">Alamat Indonesia</label>
                       {wilayahError && <p className="text-[10px] text-rose-600 font-semibold">{wilayahError}</p>}
                       {wilayahLoading && <p className="text-[10px] text-violet-600 font-semibold"><i className="fa-solid fa-spinner fa-spin mr-1"></i> Memuat data wilayah...</p>}
                       <select value={provinceId} onChange={e=>handleProvinceChange(e.target.value)} className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold outline-none" required>
                         <option value="">Pilih Provinsi...</option>
                         {provinceOptions.map(x => <option key={x.id} value={x.id}>{x.name}</option>)}
                       </select>
                       <select value={regencyId} onChange={e=>handleRegencyChange(e.target.value)} disabled={!provinceId || !regencyOptions.length} className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold outline-none disabled:opacity-50" required>
                         <option value="">Pilih Kabupaten/Kota...</option>
                         {regencyOptions.map(x => <option key={x.id} value={x.id}>{x.name}</option>)}
                       </select>
                       <select value={districtId} onChange={e=>handleDistrictChange(e.target.value)} disabled={!regencyId || !districtOptions.length} className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold outline-none disabled:opacity-50" required>
                         <option value="">Pilih Kecamatan...</option>
                         {districtOptions.map(x => <option key={x.id} value={x.id}>{x.name}</option>)}
                       </select>
                       <select value={villageId} onChange={e=>handleVillageChange(e.target.value)} disabled={!districtId || !villageOptions.length} className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold outline-none disabled:opacity-50" required>
                         <option value="">Pilih Desa/Kelurahan...</option>
                         {villageOptions.map(x => <option key={x.id} value={x.id}>{x.name}</option>)}
                       </select>
                       <input type="text" placeholder="Detail jalan, nomor rumah, RT/RW (opsional)" value={form.jalan} onChange={e=>setForm({...form, jalan:e.target.value})} className="w-full bg-white border border-slate-200 text-xs rounded-xl p-2 font-bold outline-none" />
                     </div>
                   </div>

                   <div className="flex gap-2 pt-3 border-t border-slate-100">
                     <button type="button" onClick={()=>setShowModal(false)} className="w-1/3 py-3.5 bg-slate-100 hover:bg-slate-200 transition text-slate-600 rounded-full font-bold text-xs">Batal</button>
                     <button type="submit" className="w-2/3 py-3.5 bg-slate-900 hover:bg-slate-800 transition text-white rounded-full font-bold text-xs shadow-md">Simpan Data</button>
                   </div>
                 </form>
              </div>
            </div>
          )}
        </div>
      );
    }
    
    // ==========================================
    //  MASTER LOKASI, SHIFT, DIVISI VIEWS
    // ==========================================
    
    function MasterLokasiView({ lokasiList, setLokasiList, showToast }) {
      const [showModal, setShowModal] = useState(false);
      const [editingId, setEditingId] = useState(null);
      const [form, setForm] = useState({ nama_lokasi: '', latitude: '', longitude: '', radius_meter: 100 });
      const mapContainerRef = useRef(null);
      const mapRef = useRef(null);
      const markerRef = useRef(null);
      const circleRef = useRef(null);

      // Urutkan A-Z
      const sortedList = [...lokasiList].sort((a, b) => a.nama_lokasi.localeCompare(b.nama_lokasi));

      const openModal = (item = null) => {
        if(item) {
          setEditingId(item.id); setForm({ nama_lokasi: item.nama_lokasi, latitude: item.latitude, longitude: item.longitude, radius_meter: item.radius_meter });
        } else {
          setEditingId(null); setForm({ nama_lokasi: '', latitude: '', longitude: '', radius_meter: 100 });
        }
        setShowModal(true);
      };

      const updateCoordinates = (lat, lng) => {
        const latitude = Number(lat);
        const longitude = Number(lng);
        if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) return;
        setForm(prev => ({ ...prev, latitude: latitude.toFixed(6), longitude: longitude.toFixed(6) }));
      };

      // Peta OpenStreetMap; klik peta untuk memindahkan pin.
      useEffect(() => {
        if (!showModal || !mapContainerRef.current || !window.L) return;
        const hasCoords = form.latitude !== '' && form.longitude !== '' && Number.isFinite(Number(form.latitude)) && Number.isFinite(Number(form.longitude)) && !(Number(form.latitude) === 0 && Number(form.longitude) === 0);
        const initialCenter = hasCoords ? [Number(form.latitude), Number(form.longitude)] : [-7.797, 110.370];
        const map = window.L.map(mapContainerRef.current, { zoomControl: true }).setView(initialCenter, hasCoords ? 16 : 5);
        mapRef.current = map;
        window.L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
          maxZoom: 19,
          attribution: '&copy; <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noreferrer">OpenStreetMap</a>'
        }).addTo(map);

        const marker = window.L.marker(initialCenter, { draggable: true }).addTo(map);
        markerRef.current = marker;
        const circle = window.L.circle(initialCenter, {
          radius: Math.max(1, Number(form.radius_meter) || 100),
          color: '#ec4899',
          fillColor: '#ec4899',
          fillOpacity: 0.12,
          weight: 2
        }).addTo(map);
        circleRef.current = circle;

        marker.on('dragend', () => {
          const pos = marker.getLatLng();
          updateCoordinates(pos.lat, pos.lng);
        });
        map.on('click', (event) => {
          marker.setLatLng(event.latlng);
          circle.setLatLng(event.latlng);
          updateCoordinates(event.latlng.lat, event.latlng.lng);
        });

        // Pastikan ukuran peta benar setelah modal tampil.
        const resizeTimer = setTimeout(() => map.invalidateSize(), 120);
        return () => {
          clearTimeout(resizeTimer);
          map.remove();
          if (mapRef.current === map) mapRef.current = null;
          markerRef.current = null;
          circleRef.current = null;
        };
      }, [showModal, editingId]);

      // Sinkronkan pin dan lingkaran radius jika koordinat/radius diketik manual.
      useEffect(() => {
        if (!showModal || !mapRef.current || !markerRef.current || !circleRef.current) return;
        const lat = Number(form.latitude);
        const lng = Number(form.longitude);
        if (form.latitude !== '' && form.longitude !== '' && Number.isFinite(lat) && Number.isFinite(lng) && lat >= -90 && lat <= 90 && lng >= -180 && lng <= 180) {
          const point = [lat, lng];
          markerRef.current.setLatLng(point);
          circleRef.current.setLatLng(point);
          circleRef.current.setRadius(Math.max(1, Number(form.radius_meter) || 100));
        }
      }, [form.latitude, form.longitude, form.radius_meter, showModal]);

      const useDeviceLocation = () => {
        if (!navigator.geolocation) {
          showToast('Perangkat ini tidak mendukung akses lokasi.');
          return;
        }
        navigator.geolocation.getCurrentPosition(
          (position) => {
            const { latitude, longitude } = position.coords;
            updateCoordinates(latitude, longitude);
            if (mapRef.current) mapRef.current.setView([latitude, longitude], 17);
            if (markerRef.current) markerRef.current.setLatLng([latitude, longitude]);
            if (circleRef.current) circleRef.current.setLatLng([latitude, longitude]);
            showToast('Lokasi perangkat berhasil diambil.');
          },
          (error) => {
            const message = error.code === 1 ? 'Izin lokasi ditolak. Izinkan akses lokasi di browser.' : error.code === 2 ? 'Lokasi perangkat tidak tersedia.' : 'Permintaan lokasi terlalu lama. Coba lagi.';
            showToast(message);
          },
          { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 }
        );
      };

      const handleSave = (e) => {
        e.preventDefault();
        const latitude = Number(form.latitude);
        const longitude = Number(form.longitude);
        const radius = Number.parseInt(form.radius_meter, 10);
        if (form.latitude === '' || form.longitude === '' || !Number.isFinite(latitude) || !Number.isFinite(longitude) || latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) {
          showToast('Tentukan latitude dan longitude yang valid melalui peta atau isi manual.');
          return;
        }
        if (!Number.isFinite(radius) || radius <= 0) {
          showToast('Radius harus lebih besar dari 0 meter.');
          return;
        }
        const payload = { ...form, latitude, longitude, radius_meter: radius };
        if(editingId) setLokasiList(lokasiList.map(item => item.id === editingId ? { ...item, ...payload } : item));
        else setLokasiList([...lokasiList, { id: 'lok-'+Date.now(), ...payload }]);
        setShowModal(false); showToast(editingId ? 'Lokasi diperbarui!' : 'Lokasi ditambahkan!');
      };

      return (
        <div className="space-y-6">
          <div className="flex justify-between items-center bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
            <div>
              <h2 className="font-display font-bold text-xl text-slate-800">Master Lokasi</h2>
              <p className="text-xs text-slate-400 font-medium mt-1">Kelola geofence presensi</p>
            </div>
            <button onClick={() => openModal()} className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full text-xs font-bold shadow-md active:scale-95 transition"><i className="fa-solid fa-plus mr-2"></i> Tambah Lokasi</button>
          </div>
          <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                  <th className="p-4 text-center">No</th><th className="p-4">Nama Lokasi</th><th className="p-4">Koordinat</th><th className="p-4">Radius (m)</th><th className="p-4 text-center">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {sortedList.map((item, i) => (
                  <tr key={item.id} className="hover:bg-slate-50/50">
                    <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{i+1}</td>
                    <td className="p-4 font-bold text-slate-800">{item.nama_lokasi}</td>
                    <td className="p-4 font-mono-numbers text-slate-500">{item.latitude}, {item.longitude}</td>
                    <td className="p-4 font-mono-numbers text-slate-500">{item.radius_meter}m</td>
                    <td className="p-4 text-center">
                      <button onClick={()=>openModal(item)} className="w-8 h-8 rounded-full bg-slate-50 text-indigo-500 hover:bg-indigo-50 mx-1"><i className="fa-solid fa-pen-to-square"></i></button>
                      <button onClick={()=>{setLokasiList(lokasiList.filter(x=>x.id!==item.id)); showToast('Dihapus')}} className="w-8 h-8 rounded-full bg-slate-50 text-rose-500 hover:bg-rose-50 mx-1"><i className="fa-solid fa-trash"></i></button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {showModal && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-3 sm:p-4">
              <div className="bg-white rounded-[2rem] max-w-xl w-full max-h-[94vh] overflow-y-auto p-5 sm:p-6 shadow-2xl border border-slate-100">
                 <h3 className="font-display font-bold text-lg text-slate-800 mb-1">{editingId ? 'Edit Lokasi' : 'Tambah Lokasi'}</h3>
                 <p className="text-xs text-slate-500 mb-4">Klik peta atau geser pin untuk menentukan titik lokasi.</p>
                 <form onSubmit={handleSave} className="space-y-3">
                   <input type="text" placeholder="Nama Lokasi (Gudang/Kantor)" value={form.nama_lokasi} onChange={e=>setForm({...form, nama_lokasi:e.target.value})} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none" required />
                   <div className="flex flex-wrap items-center justify-between gap-2">
                     <span className="text-[11px] font-bold text-slate-600"><i className="fa-solid fa-map-location-dot mr-1 text-pink-500"></i> Pilih titik di peta</span>
                     <button type="button" onClick={useDeviceLocation} className="px-3 py-2 rounded-full bg-pink-50 hover:bg-pink-100 text-pink-700 text-[11px] font-bold"><i className="fa-solid fa-location-crosshairs mr-1"></i> Gunakan lokasi perangkat</button>
                   </div>
                   <div ref={mapContainerRef} className="w-full h-64 sm:h-72 rounded-2xl overflow-hidden border border-slate-200 bg-slate-100" aria-label="Peta pemilih lokasi"></div>
                   <p className="text-[10px] leading-relaxed text-slate-400">Peta menggunakan OpenStreetMap. Jika memilih lokasi perangkat, browser akan meminta izin akses lokasi dan umumnya memerlukan koneksi HTTPS.</p>
                   <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                     <div>
                       <label className="block text-[10px] font-bold text-slate-500 mb-1">Latitude</label>
                       <input type="number" step="any" placeholder="Latitude (-6.xxx)" value={form.latitude} onChange={e=>setForm({...form, latitude:e.target.value})} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none font-mono-numbers" required />
                     </div>
                     <div>
                       <label className="block text-[10px] font-bold text-slate-500 mb-1">Longitude</label>
                       <input type="number" step="any" placeholder="Longitude (107.xxx)" value={form.longitude} onChange={e=>setForm({...form, longitude:e.target.value})} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none font-mono-numbers" required />
                     </div>
                   </div>
                   <div>
                     <label className="block text-[10px] font-bold text-slate-500 mb-1">Radius (meter)</label>
                     <input type="number" min="1" step="1" placeholder="Radius (Meter)" value={form.radius_meter} onChange={e=>setForm({...form, radius_meter:e.target.value})} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none font-mono-numbers" required />
                   </div>
                   <div className="flex gap-2 pt-2">
                     <button type="button" onClick={()=>setShowModal(false)} className="w-1/3 py-3.5 bg-slate-100 text-slate-600 rounded-full font-bold text-xs">Batal</button>
                     <button type="submit" className="w-2/3 py-3.5 bg-slate-900 text-white rounded-full font-bold text-xs shadow-md">Simpan</button>
                   </div>
                 </form>
              </div>
            </div>
          )}
        </div>
      );
    }

    const formatShiftTimeDisplay = (value) => {
      const match = String(value ?? '').match(/^(\d{1,2}:\d{2})/);
      return match ? match[1].padStart(5, '0') : String(value ?? '');
    };

    function MasterShiftView({ shiftList, setShiftList, divisiList, setDivisiList, toleransiMenit, setToleransiMenit, showToast }) {
      const [showModal, setShowModal] = useState(false);
      const [editingId, setEditingId] = useState(null);
      const [form, setForm] = useState({ nama_shift: '', jam_masuk: '08:00', jam_pulang: '17:00' });
      const [tempToleransi, setTempToleransi] = useState(toleransiMenit);

      // Urutkan Master Shift berdasarkan jam masuk paling pagi, lalu nama shift.
      const sortedList = [...shiftList].sort((a, b) => {
        const toMinutes = (value) => {
          const match = String(value || '').match(/^(\d{1,2}):(\d{2})/);
          return match ? Number(match[1]) * 60 + Number(match[2]) : Number.MAX_SAFE_INTEGER;
        };
        return toMinutes(a.jam_masuk) - toMinutes(b.jam_masuk) ||
          String(a.nama_shift || '').localeCompare(String(b.nama_shift || ''), 'id');
      });

      const openModal = (item = null) => {
        if(item) {
          setEditingId(item.id);
          setForm({ nama_shift: item.nama_shift, jam_masuk: formatShiftTimeDisplay(item.jam_masuk), jam_pulang: formatShiftTimeDisplay(item.jam_pulang) });
        } else {
          setEditingId(null);
          setForm({ nama_shift: '', jam_masuk: '08:00', jam_pulang: '17:00' });
        }
        setShowModal(true);
      };

      // Format otomatis ke HH:MM saat angka menit mulai dimasukkan.
      // Input boleh dikosongkan sepenuhnya ketika mengedit; validasi dilakukan saat Simpan.
      const formatShiftTimeInput = (value) => {
        const digits = String(value || '').replace(/\D/g, '').slice(0, 4);
        if (digits.length <= 2) return digits;
        return digits.slice(0, 2) + ':' + digits.slice(2);
      };

      const isValidShiftTime = (value) => {
        const match = String(value || '').match(/^(\d{2}):(\d{2})$/);
        if (!match) return false;
        const hours = Number(match[1]);
        const minutes = Number(match[2]);
        return hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59;
      };

      // Estimasi jam kerja dihitung dari rentang jam masuk hingga jam pulang.
      // Waktu istirahat tetap masuk hitungan; sisa menit diabaikan, bukan dibulatkan ke atas.
      // Jika jam pulang lebih awal dari jam masuk, dianggap melewati tengah malam.
      const getShiftWorkHours = (jamMasuk, jamPulang) => {
        const toMinutes = (value) => {
          const match = String(value || '').match(/^(\d{1,2}):(\d{2})/);
          return match ? Number(match[1]) * 60 + Number(match[2]) : null;
        };
        const masuk = toMinutes(jamMasuk);
        const pulang = toMinutes(jamPulang);
        if (masuk === null || pulang === null) return '—';
        let rentangMenit = pulang - masuk;
        if (rentangMenit < 0) rentangMenit += 24 * 60;
        return `${Math.floor(rentangMenit / 60)} jam`;
      };

      const handleSave = (e) => {
        e.preventDefault();
        if (!String(form.nama_shift || '').trim()) {
          showToast('Nama shift wajib diisi!', 'error');
          return;
        }
        if (!isValidShiftTime(form.jam_masuk) || !isValidShiftTime(form.jam_pulang)) {
          showToast('Isi jam masuk dan jam pulang dengan format HH:MM (24 jam), contoh 08:30.', 'error');
          return;
        }
        const cleanForm = {
          ...form,
          nama_shift: String(form.nama_shift).trim(),
          jam_masuk: form.jam_masuk,
          jam_pulang: form.jam_pulang
        };
        if(editingId) {
          const previousShift = shiftList.find(item => String(item.id) === String(editingId));
          const previousName = String(previousShift?.nama_shift || '').trim();
          const nextName = cleanForm.nama_shift;
          setShiftList(shiftList.map(item => String(item.id) === String(editingId) ? { ...item, ...cleanForm } : item));
          // Jika nama shift diganti, perbarui semua referensi divisi agar tidak menyimpan nama lama.
          if (previousName && previousName !== nextName) {
            setDivisiList((Array.isArray(divisiList) ? divisiList : []).map(divisi => ({
              ...divisi,
              shifts: (Array.isArray(divisi.shifts) ? divisi.shifts : []).map(name => String(name).trim() === previousName ? nextName : name).filter((name, index, arr) => arr.indexOf(name) === index)
            })));
          }
        } else {
          setShiftList([...shiftList, { id: 'sh-' + Date.now(), ...cleanForm }]);
        }
        setShowModal(false);
        showToast(editingId ? 'Shift diperbarui!' : 'Shift ditambahkan!');
      };

      useEffect(() => {
        setTempToleransi(toleransiMenit);
      }, [toleransiMenit]);

      const handleSaveToleransi = async () => {
        const value = Number(tempToleransi);
        if (!Number.isFinite(value) || value < 0 || String(tempToleransi).trim() === '') {
          showToast('Masukkan toleransi dalam menit yang valid.', 'error');
          return;
        }
        try {
          await setToleransiMenit(value);
          showToast(`Toleransi berhasil disimpan: ${value} menit.`);
        } catch (error) {
          showToast(`Toleransi gagal disimpan: ${error.message || 'periksa koneksi dan izin tabel settings.'}`, 'error');
        }
      };


      return (
        <div className="space-y-6">
          <div className="flex flex-wrap justify-between items-center gap-4 bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
            <div>
              <h2 className="font-display font-bold text-xl text-slate-800">Master Shift</h2>
              <p className="text-xs text-slate-400 font-medium mt-1">Kelola jam kerja & toleransi keterlambatan</p>
            </div>
            <div className="flex items-center gap-3">
              <div className="flex items-center bg-slate-50 border border-slate-200 rounded-full px-3 py-1.5">
                <span className="text-[11px] font-bold text-slate-500 mr-2">Toleransi:</span>
                <input 
                  type="number" 
                  min="0" 
                  max="60" 
                  value={tempToleransi} 
                  onChange={(e) => setTempToleransi(e.target.value)}
                  className="w-12 bg-transparent text-xs font-bold text-slate-800 outline-none text-center font-mono-numbers"
                />
                <span className="text-[10px] text-slate-400 mr-2">mnt</span>
                <button 
                  onClick={handleSaveToleransi}
                  className="w-6 h-6 rounded-full bg-violet-600 text-white flex items-center justify-center text-[10px] hover:bg-violet-700"
                  title="Simpan Toleransi"
                >
                  <i className="fa-solid fa-check"></i>
                </button>
              </div>
              <button 
                onClick={() => openModal()} 
                className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full text-xs font-bold shadow-md active:scale-95 transition"
              >
                <i className="fa-solid fa-plus mr-2"></i> Tambah Shift
              </button>
            </div>
          </div>

          <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                  <th className="p-4 text-center">No</th>
                  <th className="p-4">Nama Shift</th>
                  <th className="p-4">Jam Masuk</th>
                  <th className="p-4">Jam Pulang</th>
                  <th className="p-4 text-center">Jam Kerja</th>
                  <th className="p-4 text-center">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {sortedList.map((item, i) => (
                  <tr key={item.id} className="hover:bg-slate-50/50">
                    <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{i+1}</td>
                    <td className="p-4 font-bold text-slate-800">{item.nama_shift}</td>
                    <td className="p-4 font-mono-numbers text-emerald-600 font-bold">{formatShiftTimeDisplay(item.jam_masuk)} WIB</td>
                    <td className="p-4 font-mono-numbers text-rose-600 font-bold">{formatShiftTimeDisplay(item.jam_pulang)} WIB</td>
                    <td className="p-4 text-center">
                      <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-violet-50 text-violet-700 font-bold whitespace-nowrap">
                        <i className="fa-solid fa-hourglass-half text-[10px]"></i>
                        {getShiftWorkHours(item.jam_masuk, item.jam_pulang)}
                      </span>
                    </td>
                    <td className="p-4 text-center">
                      <button onClick={()=>openModal(item)} className="w-8 h-8 rounded-full bg-slate-50 text-indigo-500 hover:bg-indigo-50 mx-1">
                        <i className="fa-solid fa-pen-to-square"></i>
                      </button>
                      <button onClick={()=>{setShiftList(shiftList.filter(x=>String(x.id)!==String(item.id))); setDivisiList((Array.isArray(divisiList) ? divisiList : []).map(divisi=>({...divisi, shifts:(Array.isArray(divisi.shifts) ? divisi.shifts : []).filter(name=>String(name).trim()!==String(item.nama_shift).trim())}))); showToast('Shift dihapus dan referensi shift pada Master Divisi dibersihkan')}} className="w-8 h-8 rounded-full bg-slate-50 text-rose-500 hover:bg-rose-50 mx-1">
                        <i className="fa-solid fa-trash"></i>
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="text-[11px] text-slate-400 px-2 -mt-3">
            Keterangan jam kerja dihitung otomatis dari jam masuk hingga jam pulang, termasuk waktu istirahat. Sisa menit tidak dibulatkan ke atas (contoh 07:45–16:00 = 8 jam).
          </p>

          {showModal && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-sm w-full p-6 shadow-2xl border border-slate-100">
                <h3 className="font-display font-bold text-lg text-slate-800 mb-4">{editingId ? 'Edit Shift' : 'Tambah Shift'}</h3>
                <form onSubmit={handleSave} className="space-y-3">
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 block mb-1">Nama Shift</label>
                    <input 
                      type="text" 
                      placeholder="Contoh: Pagi (08:00 - 17:00)" 
                      value={form.nama_shift} 
                      onChange={e=>setForm({...form, nama_shift:e.target.value})} 
                      className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none" 
                      required 
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="text-[10px] font-bold text-slate-400 block mb-1">Jam Masuk</label>
                      <input 
                        type="text" inputMode="numeric" maxLength={5} placeholder="HH:MM" 
                        value={form.jam_masuk} 
                        onChange={e=>setForm({...form, jam_masuk:formatShiftTimeInput(e.target.value)})} 
                        className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none font-mono-numbers" 
                        required 
                      />
                    </div>
                    <div>
                      <label className="text-[10px] font-bold text-slate-400 block mb-1">Jam Pulang</label>
                      <input 
                        type="text" inputMode="numeric" maxLength={5} placeholder="HH:MM" 
                        value={form.jam_pulang} 
                        onChange={e=>setForm({...form, jam_pulang:formatShiftTimeInput(e.target.value)})} 
                        className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none font-mono-numbers" 
                        required 
                      />
                    </div>
                  </div>
                  <div className="flex gap-2 pt-2">
                    <button type="button" onClick={()=>setShowModal(false)} className="w-1/3 py-3.5 bg-slate-100 text-slate-600 rounded-full font-bold text-xs">Batal</button>
                    <button type="submit" className="w-2/3 py-3.5 bg-slate-900 text-white rounded-full font-bold text-xs shadow-md">Simpan</button>
                  </div>
                </form>
              </div>
            </div>
          )}
        </div>
      );
    }

    function MasterDivisiView({ divisiList, setDivisiList, shiftList, showToast }) {
      const [showModal, setShowModal] = useState(false);
      const [editingId, setEditingId] = useState(null);
      const [form, setForm] = useState({ nama_divisi: '', shifts: [], allow_cuti: true });

      const sortedList = [...divisiList].sort((a, b) => a.nama_divisi.localeCompare(b.nama_divisi, 'id'));
      const sortedShiftList = [...(Array.isArray(shiftList) ? shiftList : [])].sort((a, b) => {
        const toMinutes = (value) => {
          const match = String(value || '').match(/^(\d{1,2}):(\d{2})/);
          return match ? Number(match[1]) * 60 + Number(match[2]) : Number.MAX_SAFE_INTEGER;
        };
        return toMinutes(a.jam_masuk) - toMinutes(b.jam_masuk) || String(a.nama_shift || '').localeCompare(String(b.nama_shift || ''), 'id');
      });
      const validShiftNames = new Set(sortedShiftList.map(s => String(s.nama_shift || '').trim()));

      const openModal = (item = null) => {
        if(item) {
          setEditingId(item.id);
          setForm({ 
            nama_divisi: item.nama_divisi, 
            shifts: (item.shifts || []).filter(name => (Array.isArray(shiftList) ? shiftList : []).some(s => String(s.nama_shift || '').trim() === String(name).trim())), 
            allow_cuti: item.allow_cuti !== false 
          });
        } else {
          setEditingId(null);
          setForm({ 
            nama_divisi: '', 
            shifts: sortedShiftList[0] ? [sortedShiftList[0].nama_shift] : [], 
            allow_cuti: true 
          });
        }
        setShowModal(true);
      };

      const toggleShiftSelection = (shiftName) => {
        const current = [...form.shifts];
        if (current.includes(shiftName)) {
          setForm({ ...form, shifts: current.filter(s => s !== shiftName) });
        } else {
          setForm({ ...form, shifts: [...current, shiftName] });
        }
      };

      const handleSave = (e) => {
        e.preventDefault();
        if(editingId) {
          setDivisiList(divisiList.map(item => item.id === editingId ? { ...item, ...form } : item));
        } else {
          setDivisiList([...divisiList, { id: 'div-' + Date.now(), ...form }]);
        }
        setShowModal(false);
        showToast(editingId ? 'Divisi diperbarui!' : 'Divisi ditambahkan!');
      };

      return (
        <div className="space-y-6">
          <div className="flex justify-between items-center bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
            <div>
              <h2 className="font-display font-bold text-xl text-slate-800">Master Divisi</h2>
              <p className="text-xs text-slate-400 font-medium mt-1">Kelola divisi, hak akses cuti & penugasan shift</p>
            </div>
            <button 
              onClick={() => openModal()} 
              className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full text-xs font-bold shadow-md active:scale-95 transition"
            >
              <i className="fa-solid fa-plus mr-2"></i> Tambah Divisi
            </button>
          </div>

          <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                  <th className="p-4 text-center">No</th>
                  <th className="p-4">Nama Divisi</th>
                  <th className="p-4">Shift Terhubung</th>
                  <th className="p-4 text-center">Akses Cuti</th>
                  <th className="p-4 text-center">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {sortedList.map((item, i) => (
                  <tr key={item.id} className="hover:bg-slate-50/50">
                    <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{i+1}</td>
                    <td className="p-4 font-bold text-slate-800">{item.nama_divisi}</td>
                    <td className="p-4">
                      <div className="flex flex-wrap gap-1">
                        {item.shifts && item.shifts.filter(s => validShiftNames.has(String(s).trim())).length > 0 ? (
                          item.shifts.filter(s => validShiftNames.has(String(s).trim())).map((s, idx) => {
                            const shift = sortedShiftList.find(option => String(option.nama_shift || '').trim() === String(s).trim());
                            return <span key={`${s}-${idx}`} className="px-2 py-0.5 rounded-full bg-violet-50 text-violet-700 text-[9px] font-bold">
                              {s}{shift ? ` (${formatShiftTimeDisplay(shift.jam_masuk)})` : ''}
                            </span>;
                          })
                        ) : (
                          <span className="text-slate-400 text-[10px]">-</span>
                        )}
                      </div>
                    </td>
                    <td className="p-4 text-center">
                      <span className={`px-2.5 py-1 rounded-full text-[10px] font-bold ${item.allow_cuti ? 'bg-emerald-50 text-emerald-600' : 'bg-rose-50 text-rose-600'}`}>
                        {item.allow_cuti ? 'Diizinkan' : 'Terkunci'}
                      </span>
                    </td>
                    <td className="p-4 text-center">
                      <button onClick={()=>openModal(item)} className="w-8 h-8 rounded-full bg-slate-50 text-indigo-500 hover:bg-indigo-50 mx-1">
                        <i className="fa-solid fa-pen-to-square"></i>
                      </button>
                      <button onClick={()=>{setDivisiList(divisiList.filter(x=>x.id!==item.id)); showToast('Divisi dihapus')}} className="w-8 h-8 rounded-full bg-slate-50 text-rose-500 hover:bg-rose-50 mx-1">
                        <i className="fa-solid fa-trash"></i>
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {showModal && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-sm w-full p-6 shadow-2xl border border-slate-100">
                <h3 className="font-display font-bold text-lg text-slate-800 mb-4">{editingId ? 'Edit Divisi' : 'Tambah Divisi'}</h3>
                <form onSubmit={handleSave} className="space-y-4">
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 block mb-1">Nama Divisi</label>
                    <input 
                      type="text" 
                      placeholder="Contoh: Operasional Toko" 
                      value={form.nama_divisi} 
                      onChange={e=>setForm({...form, nama_divisi:e.target.value})} 
                      className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none" 
                      required 
                    />
                  </div>

                  <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3">
                    <label className="text-[10px] font-bold text-slate-400 block mb-2">Shift yang Berlaku</label>
                    <div className="space-y-1.5 max-h-36 overflow-y-auto pr-1">
                      {sortedShiftList.map(s => (
                        <label key={s.id} className="flex items-center space-x-2 text-xs font-semibold text-slate-700 cursor-pointer p-1.5 rounded-xl hover:bg-white transition">
                          <input 
                            type="checkbox" 
                            checked={form.shifts.includes(s.nama_shift)} 
                            onChange={() => toggleShiftSelection(s.nama_shift)}
                            className="rounded text-violet-600 focus:ring-violet-500" 
                          />
                          <span className="text-[11px] flex-1">{s.nama_shift}</span>
                          <span className="text-[10px] font-mono-numbers text-slate-400">{formatShiftTimeDisplay(s.jam_masuk)}–{formatShiftTimeDisplay(s.jam_pulang)}</span>
                        </label>
                      ))}
                    </div>
                  </div>

                  <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3 flex items-center justify-between">
                    <div>
                      <span className="text-xs font-bold text-slate-700 block">Akses Pengajuan Cuti</span>
                      <span className="text-[10px] text-slate-400">Izinkan staf mengajukan cuti mandiri</span>
                    </div>
                    <input 
                      type="checkbox" 
                      checked={form.allow_cuti} 
                      onChange={e => setForm({...form, allow_cuti: e.target.checked})}
                      className="w-4 h-4 text-violet-600 rounded focus:ring-violet-500 cursor-pointer" 
                    />
                  </div>

                  <div className="flex gap-2 pt-2">
                    <button type="button" onClick={()=>setShowModal(false)} className="w-1/3 py-3.5 bg-slate-100 text-slate-600 rounded-full font-bold text-xs">Batal</button>
                    <button type="submit" className="w-2/3 py-3.5 bg-slate-900 text-white rounded-full font-bold text-xs shadow-md">Simpan</button>
                  </div>
                </form>
              </div>
            </div>
          )}
        </div>
      );
    }

    function MasterJabatanView({ jabatanList, setJabatanList, divisiList, showToast }) {
      const [showModal, setShowModal] = useState(false);
      const [editingId, setEditingId] = useState(null);
      const [form, setForm] = useState({ nama_jabatan: '', divisi: divisiList[0]?.nama_divisi ? [divisiList[0].nama_divisi] : [] });

      const sortedList = [...jabatanList].sort((a, b) => (a.nama_jabatan || '').localeCompare(b.nama_jabatan || '', 'id'));
      const sortedDivisiList = [...(Array.isArray(divisiList) ? divisiList : [])].sort((a, b) => String(a.nama_divisi || '').localeCompare(String(b.nama_divisi || ''), 'id'));
      const getDivisions = (value) => Array.isArray(value)
        ? value.filter(Boolean)
        : String(value || '').split('|').map(v => v.trim()).filter(Boolean);

      const openModal = (item = null) => {
        if (item) {
          setEditingId(item.id);
          setForm({ nama_jabatan: item.nama_jabatan || '', divisi: getDivisions(item.divisi) });
        } else {
          setEditingId(null);
          setForm({ nama_jabatan: '', divisi: divisiList[0]?.nama_divisi ? [divisiList[0].nama_divisi] : [] });
        }
        setShowModal(true);
      };

      const toggleDivision = (division, checked) => {
        setForm(prev => ({
          ...prev,
          divisi: checked
            ? [...new Set([...getDivisions(prev.divisi), division])]
            : getDivisions(prev.divisi).filter(value => value !== division)
        }));
      };

      const handleSave = (e) => {
        e.preventDefault();
        const namaJabatan = String(form.nama_jabatan || '').trim();
        const divisiTerpilih = getDivisions(form.divisi);
        if (!namaJabatan) {
          showToast('Nama jabatan wajib diisi.', 'error');
          return;
        }
        if (!divisiTerpilih.length) {
          showToast('Pilih minimal satu divisi.', 'error');
          return;
        }
        // Simpan daftar divisi pada kolom `divisi` yang sudah ada, dipisahkan dengan `|`.
        // Ini menghindari kebutuhan menambah kolom/tabel baru di Supabase.
        const payload = { nama_jabatan: namaJabatan, divisi: divisiTerpilih.join('|') };
        if (editingId) {
          setJabatanList(jabatanList.map(item => item.id === editingId ? { ...item, ...payload } : item));
        } else {
          setJabatanList([...jabatanList, { id: 'jab-' + Date.now(), ...payload }]);
        }
        setShowModal(false);
        showToast(editingId ? 'Jabatan diperbarui!' : 'Jabatan ditambahkan!');
      };

      return (
        <div className="space-y-6">
          <div className="flex justify-between items-center bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
            <div>
              <h2 className="font-display font-bold text-xl text-slate-800">Master Jabatan</h2>
              <p className="text-xs text-slate-400 font-medium mt-1">Kelola jabatan dan hubungkan ke satu atau beberapa divisi</p>
            </div>
            <button onClick={() => openModal()} className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full text-xs font-bold shadow-md active:scale-95 transition">
              <i className="fa-solid fa-plus mr-2"></i> Tambah Jabatan
            </button>
          </div>

          <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                  <th className="p-4 text-center">No</th>
                  <th className="p-4">Nama Jabatan</th>
                  <th className="p-4">Divisi Terkait</th>
                  <th className="p-4 text-center">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {sortedList.length > 0 ? sortedList.map((item, i) => (
                  <tr key={item.id} className="hover:bg-slate-50/50">
                    <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{i+1}</td>
                    <td className="p-4 font-bold text-slate-800">{item.nama_jabatan}</td>
                    <td className="p-4">
                      <div className="flex flex-wrap gap-1.5">
                        {getDivisions(item.divisi).map(division => (
                          <span key={division} className="px-2.5 py-1 rounded-full bg-violet-50 text-violet-700 text-[10px] font-bold">{division}</span>
                        ))}
                      </div>
                    </td>
                    <td className="p-4 text-center">
                      <button onClick={() => openModal(item)} className="w-8 h-8 rounded-full bg-slate-50 text-indigo-500 hover:bg-indigo-50 mx-1" title="Edit jabatan">
                        <i className="fa-solid fa-pen-to-square"></i>
                      </button>
                      <button onClick={() => { setJabatanList(jabatanList.filter(x => x.id !== item.id)); showToast('Jabatan dihapus'); }} className="w-8 h-8 rounded-full bg-slate-50 text-rose-500 hover:bg-rose-50 mx-1" title="Hapus jabatan">
                        <i className="fa-solid fa-trash"></i>
                      </button>
                    </td>
                  </tr>
                )) : (
                  <tr><td colSpan="4" className="p-8 text-center text-slate-400">Belum ada data jabatan. Klik "Tambah Jabatan" di atas.</td></tr>
                )}
              </tbody>
            </table>
          </div>

          {showModal && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-md w-full p-6 shadow-2xl border border-slate-100 max-h-[90vh] overflow-y-auto">
                <h3 className="font-display font-bold text-lg text-slate-800 mb-4">{editingId ? 'Edit Jabatan' : 'Tambah Jabatan'}</h3>
                <form onSubmit={handleSave} className="space-y-4">
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 block mb-1.5">Nama Posisi / Jabatan</label>
                    <input type="text" placeholder="Contoh: Supervisor / Staf Gudang" value={form.nama_jabatan} onChange={e => setForm({ ...form, nama_jabatan: e.target.value })} className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none focus:border-violet-400" required />
                  </div>

                  <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3">
                    <div className="flex items-center justify-between gap-2 mb-2">
                      <label className="text-[10px] font-bold text-slate-500">Hubungkan ke Divisi</label>
                      <span className="text-[10px] font-extrabold text-violet-600">{getDivisions(form.divisi).length} dipilih</span>
                    </div>
                    <p className="text-[10px] text-slate-400 mb-3">Pilih satu atau beberapa divisi untuk jabatan ini.</p>
                    <div className="space-y-1.5 max-h-56 overflow-y-auto">
                      {sortedDivisiList.map(d => {
                        const checked = getDivisions(form.divisi).includes(d.nama_divisi);
                        return (
                          <label key={d.id || d.nama_divisi} className={`flex items-center gap-3 rounded-xl p-3 cursor-pointer border transition ${checked ? 'bg-violet-50 border-violet-200' : 'bg-white border-slate-100 hover:border-slate-200'}`}>
                            <input type="checkbox" checked={checked} onChange={e => toggleDivision(d.nama_divisi, e.target.checked)} className="w-4 h-4 accent-violet-600" />
                            <span className="text-xs font-bold text-slate-700 flex-1">{d.nama_divisi}</span>
                            {checked && <i className="fa-solid fa-check text-violet-600 text-xs"></i>}
                          </label>
                        );
                      })}
                      {divisiList.length === 0 && <p className="text-xs text-slate-400 p-3 text-center">Belum ada divisi. Tambahkan divisi terlebih dahulu.</p>}
                    </div>
                  </div>

                  <div className="flex gap-2 pt-1">
                    <button type="button" onClick={() => setShowModal(false)} className="w-1/3 py-3.5 bg-slate-100 text-slate-600 rounded-full font-bold text-xs">Batal</button>
                    <button type="submit" className="w-2/3 py-3.5 bg-slate-900 text-white rounded-full font-bold text-xs shadow-md">Simpan Jabatan</button>
                  </div>
                </form>
              </div>
            </div>
          )}
        </div>
      );
    }

    function MasterAdminView({ adminList, setAdminList, showToast }) {
      const [showModal, setShowModal] = useState(false);
      const [editingId, setEditingId] = useState(null);
      const [form, setForm] = useState({
        nama_lengkap: '',
        username: '',
        password: '',
        email: '',
        role: encodeAdminMenuAccess(ADMIN_MENU_ALL),
        menuAccess: [...ADMIN_MENU_ALL]
      });

      const openModal = (item = null) => {
        if (item) {
          setEditingId(item.id);
          setForm({
            nama_lengkap: item.nama_lengkap,
            username: item.username,
            password: item.password || '',
            email: item.email || '',
            role: item.role || encodeAdminMenuAccess(ADMIN_MENU_ALL),
            menuAccess: decodeAdminMenuAccess(item.role)
          });
        } else {
          setEditingId(null);
          setForm({
            nama_lengkap: '',
            username: '',
            password: '',
            email: '',
            role: encodeAdminMenuAccess(ADMIN_MENU_ALL),
            menuAccess: [...ADMIN_MENU_ALL]
          });
        }
        setShowModal(true);
      };

      const handleSave = (e) => {
        e.preventDefault();
        if (!form.nama_lengkap.trim() || !form.username.trim() || !form.password.trim()) {
          showToast('Nama, username, dan password wajib diisi!', 'error');
          return;
        }

        const payload = {
          nama_lengkap: form.nama_lengkap.trim(),
          username: form.username.trim(),
          password: form.password,
          email: form.email.trim(),
          // Simpan daftar hak akses di kolom role yang sudah ada agar tidak perlu mengubah skema Supabase.
          role: encodeAdminMenuAccess(form.menuAccess)
        };
        if (!payload.menuAccess && form.menuAccess.length === 0) {
          showToast('Pilih minimal satu menu untuk akun administrator.', 'error');
          return;
        }
        if (form.menuAccess.length === 0) {
          showToast('Pilih minimal satu menu untuk akun administrator.', 'error');
          return;
        }
        if (editingId) {
          setAdminList(prev => prev.map(a => a.id === editingId ? { ...a, ...payload } : a));
          showToast('Data administrator dan hak akses menu berhasil diperbarui!');
        } else {
          const newAdmin = { id: 'adm-' + Date.now(), ...payload };
          setAdminList([...adminList, newAdmin]);
          showToast('Admin baru berhasil ditambahkan!');
        }
        setShowModal(false);
      };

      const handleDelete = async (id, nama) => {
        if (adminList.length <= 1) {
          showToast('Tidak dapat menghapus admin terakhir!', 'error');
          return;
        }
        if (await modernConfirm(`Hapus akun admin "${nama}"?`)) {
          setAdminList(prev => prev.filter(a => a.id !== id));
          showToast('Akun admin berhasil dihapus.');
        }
      };

      return (
        <div className="space-y-6">
          <div className="flex justify-between items-center bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
            <div>
              <div className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full bg-violet-600"></span>
                <span className="text-[10px] font-bold uppercase tracking-widest text-violet-600">Access Management</span>
              </div>
              <h2 className="font-display font-bold text-xl text-slate-800 mt-1">Master User</h2>
              <p className="text-xs text-slate-400 font-medium">Kelola akun administrator dan pilih menu panel yang dapat diakses</p>
            </div>
            <button 
              onClick={() => openModal()} 
              className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full text-xs font-bold shadow-md active:scale-95 transition flex items-center gap-2"
            >
              <i className="fa-solid fa-plus text-xs"></i>
              <span>Tambah Admin</span>
            </button>
          </div>

          <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                  <th className="p-4 text-center">No</th>
                  <th className="p-4">Divisi</th>
                  <th className="p-4">Username</th>
                  <th className="p-4">Email</th>
                  <th className="p-4">Hak Akses Menu</th>
                  <th className="p-4 text-center">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {adminList.map((a, i) => (
                  <tr key={a.id} className="hover:bg-slate-50/50">
                    <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{i + 1}</td>
                    <td className="p-4">
                      <div className="font-bold text-slate-800">{a.nama_lengkap}</div>
                    </td>
                    <td className="p-4 font-mono-numbers text-slate-600 font-bold">{a.username}</td>
                    <td className="p-4 text-slate-500">{a.email || '-'}</td>
                    <td className="p-4">
                      <div className="flex flex-wrap gap-1 max-w-xs">
                        {getAdminMenuLabels(a.role).map(label => <span key={label} className="px-2 py-1 rounded-full text-[10px] font-bold bg-indigo-50 text-indigo-700 border border-indigo-100">{label}</span>)}
                      </div>
                    </td>
                    <td className="p-4 text-center">
                      <div className="flex items-center justify-center space-x-1">
                        <button 
                          onClick={() => openModal(a)} 
                          className="w-8 h-8 rounded-full bg-slate-50 text-indigo-500 hover:bg-indigo-50 transition active:scale-95 flex items-center justify-center"
                          title="Edit Admin"
                        >
                          <i className="fa-solid fa-pen-to-square text-xs"></i>
                        </button>
                        <button 
                          onClick={() => handleDelete(a.id, a.nama_lengkap)} 
                          className="w-8 h-8 rounded-full bg-slate-50 text-rose-500 hover:bg-rose-50 transition active:scale-95 flex items-center justify-center"
                          title="Hapus Admin"
                        >
                          <i className="fa-solid fa-trash text-xs"></i>
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {showModal && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-sm w-full p-6 shadow-2xl border border-slate-100">
                <div className="flex justify-between items-center mb-4">
                  <h3 className="font-display font-bold text-lg text-slate-800">
                    {editingId ? 'Edit Akun Admin' : 'Tambah Admin Baru'}
                  </h3>
                  <button onClick={() => setShowModal(false)} className="w-8 h-8 rounded-full bg-slate-50 text-slate-400 hover:text-slate-600">
                    <i className="fa-solid fa-xmark"></i>
                  </button>
                </div>
                <form onSubmit={handleSave} className="space-y-3">
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 block mb-1">Divisi</label>
                    <input 
                      type="text" 
                      placeholder="Contoh: HRD" 
                      value={form.nama_lengkap} 
                      onChange={e => setForm({ ...form, nama_lengkap: e.target.value })} 
                      className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none focus:border-violet-500" 
                      required 
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 block mb-1">Username Login</label>
                    <input 
                      type="text" 
                      placeholder="username_admin" 
                      value={form.username} 
                      onChange={e => setForm({ ...form, username: e.target.value })} 
                      className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none font-mono-numbers focus:border-violet-500" 
                      required 
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 block mb-1">Password</label>
                    <input 
                      type="password" 
                      placeholder="Password login..." 
                      value={form.password} 
                      onChange={e => setForm({ ...form, password: e.target.value })} 
                      className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none focus:border-violet-500" 
                      required 
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 block mb-1">Email</label>
                    <input 
                      type="email" 
                      placeholder="admin@suryatokoku.com" 
                      value={form.email} 
                      onChange={e => setForm({ ...form, email: e.target.value })} 
                      className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold outline-none focus:border-violet-500" 
                      required 
                    />
                  </div>
                  <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                    <div className="flex items-center justify-between gap-2 mb-1">
                      <label className="text-xs font-extrabold text-slate-700">Hak Akses Menu Panel Admin</label>
                      <span className="text-[10px] font-extrabold text-violet-600">{form.menuAccess.length} dipilih</span>
                    </div>
                    <p className="text-[10px] text-slate-400 mb-3">Centang menu yang boleh dilihat akun ini setelah login. Pilihan bisa lebih dari satu.</p>
                    <div className="space-y-2">
                      {ADMIN_MENU_OPTIONS.map(item => {
                        const checked = form.menuAccess.includes(item.id);
                        return <label key={item.id} className={`flex items-center gap-3 rounded-xl border p-3 cursor-pointer transition ${checked ? 'bg-violet-50 border-violet-200' : 'bg-white border-slate-100 hover:border-slate-200'}`}>
                          <input type="checkbox" checked={checked} onChange={e => setForm(prev => ({ ...prev, menuAccess: e.target.checked ? [...prev.menuAccess, item.id] : prev.menuAccess.filter(id => id !== item.id) }))} className="w-4 h-4 accent-violet-600" />
                          <span className="text-xs font-bold text-slate-700 flex-1">{item.label}</span>
                          {checked && <i className="fa-solid fa-check text-violet-600 text-xs"></i>}
                        </label>;
                      })}
                    </div>
                  </div>
                  <div className="flex gap-2 pt-2">
                    <button 
                      type="button" 
                      onClick={() => setShowModal(false)} 
                      className="w-1/3 py-3.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-full font-bold text-xs"
                    >
                      Batal
                    </button>
                    <button 
                      type="submit" 
                      className="w-2/3 py-3.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full font-bold text-xs shadow-md active:scale-95 transition"
                    >
                      Simpan Admin
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}
        </div>
      );
    }

    /* =========================================================
       MODUL MASTER PENGUMUMAN HRD
       ========================================================= */
    function MasterPengumumanView({ pengumumanList, setPengumumanList, showToast, currentUser }) {
      const [searchQuery, setSearchQuery] = useState('');
      const [showModal, setShowModal] = useState(false);
      const [editingId, setEditingId] = useState(null);
      const [previewItem, setPreviewItem] = useState(null);
      const currentDivision = String(currentUser?.divisi || currentUser?.nama_lengkap || '').trim() || 'HRD';
      const [form, setForm] = useState({ judul: '', isi: '', author: currentDivision });

      const filteredList = useMemo(() => {
        return pengumumanList.filter(item => {
          const matchSearch = !searchQuery.trim() ||
            String(item.judul || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
            String(item.isi || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
            String(item.author || '').toLowerCase().includes(searchQuery.toLowerCase());
          return matchSearch;
        });
      }, [pengumumanList, searchQuery]);

      const openModal = (item = null) => {
        if (item) {
          setEditingId(item.id);
          setForm({
            judul: item.judul || '',
            isi: item.isi || '',
            author: item.author || currentDivision
          });
        } else {
          setEditingId(null);
          setForm({ judul: '', isi: '', author: currentDivision });
        }
        setShowModal(true);
      };

      const handleSave = (e) => {
        e.preventDefault();
        if (!form.judul.trim() || !form.isi.trim()) {
          showToast('Judul dan isi pengumuman wajib diisi!', 'error');
          return;
        }

        if (editingId) {
          setPengumumanList(prev => prev.map(p => p.id === editingId ? { ...p, judul: form.judul.trim(), isi: form.isi.trim(), author: p.author || form.author } : p));
          showToast('Pengumuman berhasil diperbarui!');
        } else {
          const newItem = {
            id: 'pgm-' + Date.now(),
            judul: form.judul.trim(),
            isi: form.isi.trim(),
            author: currentDivision,
            kategori: 'Informasi Umum', // nilai kompatibilitas internal; kategori tidak ditampilkan di antarmuka
            created_at: new Date().toISOString()
          };
          setPengumumanList([newItem, ...pengumumanList]);
          showToast('Pengumuman baru berhasil diterbitkan!');
        }
        setShowModal(false);
      };

      const handleDelete = async (id, judul) => {
        if (await modernConfirm(`Hapus pengumuman "${judul}"?`)) {
          setPengumumanList(prev => prev.filter(p => p.id !== id));
          showToast('Pengumuman berhasil dihapus.');
        }
      };


      return (
        <div className="space-y-6">
          {/* Header Panel */}
          <div className="flex flex-wrap justify-between items-center gap-4 bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm">
            <div>
              <div className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full bg-fuchsia-500"></span>
                <span className="text-[10px] font-bold uppercase tracking-widest text-fuchsia-600">Broadcast & Info Center</span>
              </div>
              <h2 className="font-display font-bold text-xl text-slate-800 mt-1">Master Pengumuman HR</h2>
              <p className="text-xs text-slate-400 font-medium">Publikasikan informasi resmi, edaran, dan SOP kerja kepada seluruh staf</p>
            </div>

            <button 
              onClick={() => openModal()}
              className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full text-xs font-bold shadow-md active:scale-95 transition flex items-center gap-2"
            >
              <i className="fa-solid fa-plus text-xs"></i>
              <span>Buat Pengumuman Baru</span>
            </button>
          </div>

          {/* Filter Bar & Search */}
          <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm p-4 space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="relative flex-1 max-w-sm">
                <i className="fa-solid fa-magnifying-glass absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
                <input 
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Cari judul atau isi pengumuman..."
                  className="w-full bg-slate-50 border border-slate-200 rounded-2xl pl-9 pr-4 py-2 text-xs font-semibold outline-none focus:border-fuchsia-500"
                />
              </div>

              <div className="text-[11px] font-bold text-slate-400">
                Total: <span className="font-mono-numbers text-slate-700 font-bold">{filteredList.length}</span> Pengumuman
              </div>
            </div>

          </div>

          {/* Tabel Pengumuman */}
          <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 text-slate-400 uppercase font-bold border-b border-slate-100 text-[10px] tracking-wider">
                    <th className="p-4 text-center">No</th>
                    <th className="p-4">Judul Pengumuman</th>
                    <th className="p-4">Ringkasan Pesan</th>
                    <th className="p-4">Penulis</th>
                    <th className="p-4">Tanggal Rilis</th>
                    <th className="p-4 text-center">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {filteredList.length > 0 ? (
                    filteredList.map((item, idx) => (
                      <tr key={item.id} className="hover:bg-fuchsia-50/20 transition">
                        <td className="p-4 text-center font-mono-numbers text-slate-400 font-bold">{idx + 1}</td>
                        <td className="p-4 max-w-xs">
                          <span 
                            onClick={() => setPreviewItem(item)}
                            className="font-bold text-slate-800 hover:text-fuchsia-600 cursor-pointer block truncate"
                            title="Klik untuk membaca selengkapnya"
                          >
                            {item.judul}
                          </span>
                        </td>
                        <td className="p-4 max-w-sm">
                          <p className="text-slate-500 font-medium text-[11px] truncate leading-relaxed">
                            {item.isi}
                          </p>
                        </td>
                        <td className="p-4 font-semibold text-slate-700 text-[11px]">
                          {item.author || 'HRD'}
                        </td>
                        <td className="p-4 text-[11px] text-slate-400 font-mono-numbers">
                          {new Date(item.created_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}
                        </td>
                        <td className="p-4 text-center">
                          <div className="flex items-center justify-center space-x-1">
                            <button 
                              onClick={() => setPreviewItem(item)}
                              className="w-8 h-8 rounded-full bg-slate-50 text-slate-600 hover:bg-slate-100 flex items-center justify-center transition active:scale-95"
                              title="Baca Lengkap"
                            >
                              <i className="fa-regular fa-eye text-xs"></i>
                            </button>
                            <button 
                              onClick={() => openModal(item)}
                              className="w-8 h-8 rounded-full bg-slate-50 text-indigo-600 hover:bg-indigo-50 flex items-center justify-center transition active:scale-95"
                              title="Edit Pengumuman"
                            >
                              <i className="fa-solid fa-pen-to-square text-xs"></i>
                            </button>
                            <button 
                              onClick={() => handleDelete(item.id, item.judul)}
                              className="w-8 h-8 rounded-full bg-slate-50 text-rose-500 hover:bg-rose-50 flex items-center justify-center transition active:scale-95"
                              title="Hapus Pengumuman"
                            >
                              <i className="fa-solid fa-trash text-xs"></i>
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="6" className="p-8 text-center text-slate-400 font-medium">
                        Tidak ditemukan pengumuman dengan kata kunci tersebut.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* Modal Form Tambah / Edit */}
          {showModal && (
            <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-[2rem] max-w-lg w-full p-6 shadow-2xl border border-slate-100">
                <div className="flex justify-between items-center mb-4">
                  <div>
                    <span className="text-[10px] font-bold uppercase text-slate-400">Editor Pengumuman</span>
                    <h3 className="font-display font-bold text-lg text-slate-800">{editingId ? 'Edit Pengumuman' : 'Buat Pengumuman Baru'}</h3>
                  </div>
                  <button onClick={() => setShowModal(false)} className="w-8 h-8 rounded-full bg-slate-50 text-slate-400 hover:text-slate-600">
                    <i className="fa-solid fa-xmark"></i>
                  </button>
                </div>

                <form onSubmit={handleSave} className="space-y-4">
                  <div>
                    <label className="text-[10px] font-bold text-slate-500 block mb-1">Judul Pengumuman</label>
                    <input 
                      type="text" 
                      placeholder="Contoh: Pemberitahuan Libur Operasional"
                      value={form.judul}
                      onChange={(e) => setForm({ ...form, judul: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3 font-bold text-slate-800 outline-none focus:border-fuchsia-500"
                      required
                    />
                  </div>

                  <div>
                    <label className="text-[10px] font-bold text-slate-500 block mb-1">Penerbit / Penulis</label>
                    <input
                      type="text"
                      value={editingId ? (form.author || currentDivision) : currentDivision}
                      readOnly
                      aria-readonly="true"
                      className="w-full bg-slate-100 border border-slate-200 text-xs rounded-2xl p-3 font-bold text-slate-500 outline-none cursor-not-allowed"
                    />
                    <p className="text-[10px] text-slate-400 mt-1">Terisi otomatis berdasarkan Divisi akun administrator yang login dan tidak dapat diedit.</p>
                  </div>

                  <div>
                    <label className="text-[10px] font-bold text-slate-500 block mb-1">Isi Lengkap Pengumuman</label>
                    <textarea 
                      rows="6"
                      placeholder="Tuliskan detail informasi, petunjuk, instruksi operasional, atau ketentuan lengkap..."
                      value={form.isi}
                      onChange={(e) => setForm({ ...form, isi: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 text-xs rounded-2xl p-3.5 text-slate-800 outline-none leading-relaxed focus:border-fuchsia-500"
                      required
                    ></textarea>
                  </div>

                  <div className="flex gap-2 pt-2 border-t border-slate-100">
                    <button 
                      type="button" 
                      onClick={() => setShowModal(false)}
                      className="w-1/3 py-3.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-full font-bold text-xs"
                    >
                      Batal
                    </button>
                    <button 
                      type="submit"
                      className="w-2/3 py-3.5 bg-slate-900 hover:bg-slate-800 text-white rounded-full font-bold text-xs shadow-md active:scale-95 transition"
                    >
                      {editingId ? 'Simpan Perubahan' : 'Terbitkan Pengumuman'}
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}

          {/* Modal Pratinjau Pengumuman */}
          {previewItem && (
            <div 
              onClick={() => setPreviewItem(null)}
              className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4"
            >
              <div 
                onClick={e => e.stopPropagation()}
                className="bg-white rounded-[2rem] max-w-md w-full p-6 shadow-2xl border border-slate-100"
              >
                <div className="flex items-center justify-end mb-3">
                  <button 
                    onClick={() => setPreviewItem(null)}
                    className="w-8 h-8 rounded-full bg-slate-100 text-slate-400 hover:text-slate-600 flex items-center justify-center"
                  >
                    <i className="fa-solid fa-xmark"></i>
                  </button>
                </div>

                <h3 className="font-display font-bold text-base text-slate-800 mb-2 leading-snug">{previewItem.judul}</h3>
                
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 my-4 max-h-60 overflow-y-auto">
                  <p className="text-xs text-slate-600 leading-relaxed whitespace-pre-line">{previewItem.isi}</p>
                </div>

                <div className="flex justify-between items-center text-[10px] font-bold text-slate-400 pt-2 border-t border-slate-100">
                  <span className="flex items-center gap-1.5 text-slate-600">
                    <i className="fa-solid fa-user-pen text-violet-500"></i>
                    {previewItem.author}
                  </span>
                  <span className="font-mono-numbers">
                    {new Date(previewItem.created_at).toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}
                  </span>
                </div>
              </div>
            </div>
          )}
        </div>
      );
    }

    const rootElement = document.getElementById('root');
    ReactDOM.createRoot(rootElement).render(<App />);
  
