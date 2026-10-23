'use strict';

// Variabel untuk mengecek apakah kita sedang berjalan di dalam Word atau di Browser biasa
let isRunningInWord = false;

(function () {
    Office.onReady(function (info) {
        $(document).ready(function () {
            
            // 1. Cek Lingkungan (Word vs HTML Editor)
            if (info.host === Office.HostType.Word && Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                isRunningInWord = true;
                $('#supportedVersion').html('Berjalan di Microsoft Word (Word 2016 atau lebih baru).');
            } else {
                // Jika dibuka di HTML Editor / Browser biasa
                $('#supportedVersion').html('Mode Preview: Buka di Microsoft Word agar tombol berfungsi menyisipkan teks.');
            }

            // 2. Daftarkan semua event klik tombol
            bindButtonEvents();
        });
    });

    // ==========================================
    // FUNGSI UNTUK MENDAFTARKAN TOMBOL
    // ==========================================
    function bindButtonEvents() {
        // Pemetaan ID Tombol ke Karakter (Jauh lebih ringkas daripada 24 fungsi terpisah)
        const charMap = {
            'a1': 'Ā', 'a2': 'ā', 
            'i1': 'Ī', 'i2': 'ī', 
            'u1': 'Ū', 'u2': 'ū',
            'd1': 'Ḍ', 'd2': 'ḍ', 
            'g1': 'Ġ', 'g2': 'ġ', 
            'h1': 'Ḥ', 'h2': 'ḥ',
            's1': 'Ṡ', 's2': 'ṡ', 
            's3': 'Ṣ', 's4': 'ṣ', 
            't1': 'Ṭ', 't2': 'ṭ',
            'z1': 'Ż', 'z2': 'ż', 
            'z3': 'Ẓ', 'z4': 'ẓ', 
            'ain': 'ʿ', 'hamzah': 'ʾ'
        };

        // Looping untuk mendaftarkan klik pada tombol karakter
        for (const [id, char] of Object.entries(charMap)) {
            $('#' + id).click(function () {
                insertCharacter(char);
            });
        }

        // Daftarkan klik pada tombol aksi
        $('#btn-space').click(function () { insertCharacter(' '); });
        $('#btn-enter').click(insertEnter);
        $('#btn-backspace').click(deleteSelection);
        $('#btn-del').click(deleteSelection);
    }

    // ==========================================
    // FUNGSI UTAMA (Menyisipkan ke Word / Preview)
    // ==========================================

    // Fungsi untuk menyisipkan karakter biasa & spasi
    function insertCharacter(char) {
        if (!isRunningInWord) {
            // Simulasi jika dijalankan di HTML Editor
            alert('Mode Preview: Menyisipkan karakter "' + char + '"\n\n(Buka di Microsoft Word untuk fungsi asli)');
            return;
        }

        Word.run(function (context) {
            const range = context.document.getSelection();
            range.insertText(char, Word.InsertLocation.replace);
            return context.sync();
        }).catch(handleError);
    }

    // Fungsi untuk tombol Enter (Baris Baru)
    function insertEnter() {
        if (!isRunningInWord) {
            alert('Mode Preview: Menyisipkan Enter / Baris Baru');
            return;
        }

        Word.run(function (context) {
            const range = context.document.getSelection();
            range.insertParagraph(); // Lebih tepat untuk baris baru di Word
            return context.sync();
        }).catch(handleError);
    }

    // Fungsi untuk tombol Backspace & Del (Menghapus teks yang disorot)
    function deleteSelection() {
        if (!isRunningInWord) {
            alert('Mode Preview: Menghapus teks yang dipilih');
            return;
        }

        Word.run(function (context) {
            const range = context.document.getSelection();
            range.clear(); // Menghapus teks yang sedang disorot/diblok
            return context.sync();
        }).catch(handleError);
    }

    // ==========================================
    // FUNGSI BANTUAN UNTUK ERROR
    // ==========================================
    function handleError(error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    }

})();