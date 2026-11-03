# Dashboard UBK SMadJiwa — GitHub Pages + GAS Backend

## Fail
- `index.html`: frontend untuk GitHub Pages (tema maroon/purple, mobile-first).
- `Code.gs`: backend GAS asal dengan endpoint bridge tambahan.

## Setup
1. Dalam projek Apps Script sedia ada, tambah/kemas kini `doGet(e)` dan `getGitHubBridgeHtml_()` daripada fail `Code.gs` pakej ini. Kekalkan semua fungsi sedia ada.
2. Deploy Apps Script sebagai Web App: Execute as **Me**, access **Anyone** (atau pilihan organisasi yang sepadan dengan akaun pengguna). Simpan URL berakhir `/exec`.
3. Buka `index.html`, cari `PASTE_YOUR_GAS_WEB_APP_EXEC_URL_HERE`, ganti dengan URL Web App `/exec`. Jangan masukkan URL `/dev`.
4. Muat naik `index.html` ke repository GitHub, pada root sebagai `index.html`. Aktifkan Settings → Pages → Deploy from branch → `main` / root.
5. Buka URL GitHub Pages. Login akan dihantar ke backend GAS melalui iframe bridge.

## Nota keselamatan & operasi
- Jangan letak Spreadsheet ID, password, atau token tetap dalam frontend.
- Fungsi server masih perlu mengesahkan token dan kebenaran modul seperti kod asal.
- Bridge menerima panggilan daripada origin `https://<username>.github.io` sahaja. Jika guna custom domain, tambah domain tersebut pada regex `TRUSTED` dalam `getGitHubBridgeHtml_()`.
- Google Apps Script masih backend; banner Apps Script tidak muncul pada halaman GitHub Pages kerana GAS hanya dimuatkan sebagai iframe 1px untuk RPC.
- Ini memerlukan deploy semula GAS selepas menambah bridge. `Code.gs` dalam pakej mengandungi kod asal yang diberikan bersama penambahan `doGet(e)` bridge.
