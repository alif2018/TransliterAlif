# 📝 Transliterasi Alif - Word Add-in

**Aplikasi Transliterasi Arab - Indonesia SKB Nomor 158 Tahun 1987 dan Nomor 0543b/U/1987**

Transliterasi Alif adalah sebuah *Office Add-in* (ekstensi) untuk Microsoft Word yang dirancang untuk mempermudah pengguna dalam menyisipkan karakter khusus transliterasi Arab-Latin. Add-in ini sangat membantu penulis, akademisi, santri, dan mahasiswa yang sering menggunakan pedoman transliterasi sesuai dengan **Surat Keputusan Bersama (SKB) Menteri Agama dan Menteri Pendidikan dan Kebudayaan Republik Indonesia**.

---

## ✨ Fitur Utama

* **Sesuai Standar SKB:** Menyediakan karakter khusus seperti *macron* (Ā, Ī, Ū), titik di bawah (Ḍ, Ṣ, Ṭ), dan *half-ring modifier* yang tepat untuk Ain (ʿ) dan Hamzah (ʾ).
* **Antarmuka Bersih (Clean UI):** Tampilan panel samping (Task Pane) yang rapi dengan tata letak 6 kolom menggunakan font *Serif* klasik yang mudah dibaca.
* **Tombol Aksi Cepat:** Dilengkapi tombol fungsional untuk **Spasi**, **Enter** (baris baru), **Backspace**, dan **Del** (menghapus teks yang disorot).
* **Mode Preview:** Dapat dibuka langsung di browser web sebagai mode pratinjau (tanpa merusak fungsi), dan akan berfungsi penuh saat dipasang di Microsoft Word.
* **Panduan Terintegrasi:** Tombol info (ⓘ) di pojok kanan atas untuk mengunduh file manifest dan melihat tutorial instalasi secara langsung.

---

## 🚀 Cara Instalasi (Untuk Pengguna)

Add-in ini di-hosting secara publik melalui Vercel (`https://transliteralif.vercel.app/`). Untuk memasangnya di Microsoft Word:

1. Buka [https://transliteralif.vercel.app/](https://transliteralif.vercel.app/) di browser Anda.
2. Klik ikon **Info (ⓘ)** di pojok kanan atas.
3. Klik tombol **Download Transliterasi Alif (XML)** untuk mengunduh file `transliteralif.xml`.
4. Buka aplikasi **Microsoft Word**.
5. Pergi ke tab **Home** > klik **Add-ins** > pilih **More Add-ins** (atau *Get Add-ins*).
6. Pilih tab **MY ADD-INS**.
7. Di bagian bawah, klik **Upload My Add-in**.
8. Pilih file **`transliteralif.xml`** yang sudah Anda unduh tadi, lalu klik **OK**.
9. Add-in siap digunakan dan akan muncul di panel kanan Word Anda!

---

## 🛠️ Detail Pengembangan (Untuk Developer)

Aplikasi ini dibangun menggunakan *stack* web standar tanpa *framework* berat, sehingga sangat ringan dan mudah dikembangkan.

* **Frontend:** HTML5, CSS3 (CSS Grid & Flexbox), JavaScript (ES6).
* **Library:** [Office.js](https://appsforoffice.microsoft.com/lib/1/hosted/office.js) (Microsoft Office Add-ins API), jQuery 3.6.0.
* **Hosting:** [Vercel](https://vercel.com) (Otomatis HTTPS).
* **Repository:** GitHub.

### Struktur File:
```text
├── index.html           # Antarmuka utama Add-in (UI)
├── home.css             # Styling layout dan tombol
├── home.js              # Logika Office.js dan pendeteksi klik tombol
├── transliteralif.xml   # File manifest (KTP Add-in)
└── assets/              # Ikon Add-in (16x16, 32x32, 80x80)

```
---
## 📚 Referensi

* Surat Keputusan Bersama Menteri Agama dan Menteri Pendidikan dan Kebudayaan Republik Indonesia:
*   Nomor: 158 Tahun 1987
*   Nomor: 0543b/U/1987
*   Tentang: Pedoman Transliterasi Arab-Latin.

---

Dibuat untuk mempermudah penulisan karya ilmiah dan teks keagamaan dengan standar ejaan yang tepat.
