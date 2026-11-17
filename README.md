# Portal Latihan PKD Kuala Selangor

Google Apps Script: `Code.gs` dan `Index.html`.

## Dashboard Portal / HRMIS / Banding

Kemas kini 17 November 2026 menambah pilihan sumber untuk kedua-dua graf. Jata, ikon, tajuk, wording menu, borang dan aliran sedia ada dikekalkan.

- **Portal:** pengiraan sedia ada berdasarkan Tarikh Mula dan jumlah staf organisasi bagi tahun dipilih.
- **HRMIS:** jumlah tahunan mengikut IC/tahun daripada tab `HRMIS_JAM` yang sudah diimport dan disahkan oleh admin. Pembahagi ialah bilangan staf dalam data HRMIS tahun itu, bukan tetapan 551. Tiada sambungan automatik ke sistem HRMIS.
- **Banding:** kedua-dua sumber dipaparkan bersebelahan pada desktop dan bertingkat pada telefon. Satu nota kecil menyatakan bilangan staf dan tarikh HRMIS; peratus menggunakan pembahagi masing-masing.
- Graf jawatan HRMIS menggunakan label jawatan dalam laporan dan bilangan staf setiap label sebagai pembahagi. Ia tidak memadankan skim/gred secara tekaan dengan kategori perjawatan portal.
- 0 jam HRMIS hanya staf yang dilaporkan tepat 0 minit. Staf yang tiada dalam laporan tidak ditambah sebagai 0 jam. Jika ada jumlah positif kurang 1 jam, ia dipaparkan sebagai kategori tambahan supaya pecahan lengkap.
- Nota kecil di atas dashboard menunjukkan jumlah staf Portal/HRMIS dan tarikh laporan. Rekod lama yang masih bercampur tarikh ditunjukkan sebagai julat tarikh sekali sahaja; amaran tarikh tidak diulang pada kedua-dua graf.
- Nama jawatan boleh membalut baris pada telefon. Legend doughnut di kanan.

## Import HRMIS menggantikan seluruh tahun dipilih

- Selepas admin mengesahkan pratonton, semua rekod HRMIS bagi tahun dipilih digantikan dengan gabungan fail/pautan dalam import itu sahaja. Staf lama yang tiada dalam import dikeluarkan daripada `HRMIS_JAM` tahun tersebut; ia tidak dikira sebagai 0 jam.
- Gunakan laporan keseluruhan. Jika laporan cukup dan kurang 40 jam berasingan, pilih kedua-duanya dalam satu import. Import fail kedua secara berasingan akan menggantikan fail pertama bagi tahun yang sama.
- Pratonton menyenaraikan rekod yang akan dikeluarkan serta bilangan staf lama/baharu. Checkbox pengesahan menerangkan penggantian penuh tahun dan bilangan rekod yang dikeluarkan.
- Tahun HRMIS lain, latihan Portal, sijil, CPD dan status pengakuan HRMIS setiap kursus dikekalkan. Pratonton sendiri tidak menulis database.
- Penggantian boleh membetulkan laporan lama yang tersalah bertarikh lebih lewat. Tarikh dipilih yang lebih awal diberitahu dalam pratonton tetapi dibenarkan selepas pengesahan. Tempoh laporan mesti tetap sepadan dengan tahun/tarikh dipilih.
- Pratonton versi lama tidak boleh mengesahkan penggantian. Semakan sesi admin, hash database, kunci skrip dan resit pengesahan berulang dikekalkan.
- Kandungan pengganti dibina dahulu, kemudian satu `setValues` menulis laporan dan mengosongkan baki baris lama. Tiada `clear`/`delete` data sebelum penulisan pengganti.

## Pasang dengan URL portal sedia ada

1. Ganti kandungan kedua-dua fail dalam projek Apps Script yang betul.
2. **Save → Deploy → Manage deployments**.
3. Pilih deployment yang **Web app URL** sepadan dengan URL lama yang digunakan staf.
4. **Edit → Version: New version → Deploy**.
5. Buka semula URL lama. **Admin → upload laporan keseluruhan terkini → Semak import → semak senarai dikeluarkan → tandakan pengesahan → Sahkan & kemas kini HRMIS_JAM**.

Tidak perlu mengulang persediaan tab tahunan atau menetapkan semula `HRMIS_ADMIN_PASSWORD`. **Untuk membersihkan data lama yang sudah bercampur, import semula laporan terkini selepas deploy. Memasang kod sahaja tidak mengubah data.** Tiada perubahan manifest/scopes diperlukan. Perubahan di GitHub tidak deploy Apps Script secara automatik.

Kod lama kekal dalam sejarah commit GitHub. Kod ini tidak menjalankan rollback atau memadam laporan apabila tab dashboard ditukar.

## Semakan

116 semakan simulasi semasa lulus: penggantian tahun penuh, pengosongan baki baris, tahun lain dikekalkan, deduplikasi dua fail, pembetulan tarikh, input kosong/tidak sah, hash database berubah, sesi admin, protokol lama, pengesahan berulang, kegagalan penulisan, satu nota dashboard dan checkbox persetujuan. Hero, ikon/wording menu serta modul simpan/sijil/status staf dibandingkan dengan versi sebelum perubahan ini.

Tiada database atau deployment sebenar diubah oleh proses kemas kini kod ini. Ujian simulasi tidak menggantikan semakan selepas deploy dalam GAS.

