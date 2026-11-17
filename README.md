# Portal Latihan PKD Kuala Selangor

Google Apps Script: `Code.gs` dan `Index.html`.

## Dashboard Portal / HRMIS / Banding

Kemas kini 17 November 2026 menambah pilihan sumber untuk kedua-dua graf. Jata, ikon, tajuk, wording menu, borang dan aliran sedia ada dikekalkan.

- **Portal:** pengiraan sedia ada berdasarkan Tarikh Mula dan jumlah staf organisasi bagi tahun dipilih.
- **HRMIS:** jumlah tahunan mengikut IC/tahun daripada tab `HRMIS_JAM` yang sudah diimport dan disahkan oleh admin. Pembahagi ialah bilangan staf dalam data HRMIS tahun itu, bukan tetapan 551. Tiada sambungan automatik ke sistem HRMIS.
- **Banding:** kedua-dua sumber dipaparkan bersebelahan pada desktop dan bertingkat pada telefon. Skop staf yang berbeza ditandakan; peratus menggunakan pembahagi masing-masing.
- Graf jawatan HRMIS menggunakan label jawatan dalam laporan dan bilangan staf setiap label sebagai pembahagi. Ia tidak memadankan skim/gred secara tekaan dengan kategori perjawatan portal.
- 0 jam HRMIS hanya staf yang dilaporkan tepat 0 minit. Staf yang tiada dalam laporan tidak ditambah sebagai 0 jam. Jika ada jumlah positif kurang 1 jam, ia dipaparkan sebagai kategori tambahan supaya pecahan lengkap.
- Jika import beberapa tarikh, semua jumlah terkini yang disimpan bagi IC/tahun dikekalkan dan julat tarikh dipaparkan dengan amaran. Laporan lama yang tiada dalam import baharu tidak dipadam secara automatik.
- Nama jawatan boleh membalut baris pada telefon. Legend doughnut di kanan.

## Pasang dengan URL portal sedia ada

1. Ganti kandungan kedua-dua fail dalam projek Apps Script yang betul.
2. **Save → Deploy → Manage deployments**.
3. Pilih deployment yang **Web app URL** sepadan dengan URL lama yang digunakan staf.
4. **Edit → Version: New version → Deploy**.
5. Buka semula URL lama. Gunakan **Refresh** untuk membaca semula data.

Tidak perlu mengulang persediaan tab tahunan, menetapkan semula `HRMIS_ADMIN_PASSWORD`, atau mengimport semula laporan yang sudah disimpan dalam `HRMIS_JAM`. Tiada perubahan manifest/scopes diperlukan. Perubahan di GitHub tidak deploy Apps Script secara automatik.

Kod lama kekal dalam sejarah commit GitHub. Kod ini tidak menjalankan rollback atau memadam laporan apabila tab dashboard ditukar.

## Semakan

101 semakan sintaks, pengiraan dan aliran UI melalui simulasi Spreadsheet/DOM/Chart lulus: sempadan minit, tahun berbeza, import separa, skop 578 berbanding 551, tarikh laporan bercampur, data tidak sah, kebebasan ralat setiap sumber, tab dan respons lewat. Semakan juga membandingkan hero/menu dan modul simpan/import/status/sijil dengan kod asal.

Tiada database atau deployment sebenar diubah oleh proses kemas kini kod ini. Ujian simulasi tidak menggantikan semakan selepas deploy dalam GAS.
