'use strict';

// Variabel untuk mengecek apakah kita sedang berjalan di dalam Word atau di Browser biasa
let isRunningInWord = false;

(function () {
    Office.onReady(function (info) {
        $(document).ready(function () {
            
            // 1. Cek Lingkungan (Word vs Browser/HTML Editor)
            if (info.host === Office.HostType.Word && Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                isRunningInWord = true;
                $('#supportedVersion').html('Berjalan di Microsoft Word.');
                
                // Sembunyikan tombol download di dalam Modal karena pengguna sudah berada di Word
                $('#download-btn').hide(); 
            } else {
                $('#supportedVersion').html('Mode Web: Klik ikon (ⓘ) di pojok kanan atas untuk panduan install.');
                
                // Tampilkan tombol download manifest
                $('#download-btn').show();
            }

            // 2. Daftarkan semua event klik tombol transliterasi dan aksi
            bindButtonEvents();

            // 3. Logika Modal (Pop-up) Tutorial
            // Buka modal saat tombol info diklik
            $('#info-btn').click(function() {
                $('#tutorial-modal').fadeIn(200);
            });

            // Tutup modal saat tombol silang (X) diklik
            $('.close-btn').click(function() {
                $('#tutorial-modal').fadeOut(200);
            });

            // Tutup modal jika area gelap di luar kotak (overlay) diklik
            $(window).click(function(event) {
                if ($(event.target).is('#tutorial-modal')) {
                    $('#tutorial-modal').fadeOut(200);
                }
            });

        });
    });

    // ==========================================
    // FUNGSI UNTUK MENDAFTARKAN EVENT KLIK TOMBOL
    // ==========================================
    function bindButtonEvents() {
        // Pemetaan ID Tombol ke Karakter (Ringkas & Mudah dikelola)
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

        // Looping untuk mendaftarkan klik pada semua tombol karakter
        for (const [id, char] of Object.entries(charMap)) {
            $('#' + id).click(function () {
                insertCharacter(char);
            });
        }

        // Daftarkan klik pada tombol aksi bawah
        $('#btn-space').click(function () { insertCharacter(' '); });
        $('#btn-enter').click(insertEnter);
        $('#btn-backspace').click(deleteSelection);
        $('#btn-del').click(deleteSelection);
    }

    // ==========================================
    // FUNGSI UTAMA (Menyisipkan ke Word / Mode Preview)
    // ==========================================

    // Menyisipkan karakter biasa & spasi
    function insertCharacter(char) {
        if (!isRunningInWord) {
            alert('Mode Preview: Menyisipkan karakter "' + char + '"\n\n(Buka di Microsoft Word untuk fungsi asli)');
            return;
        }

        Word.run(function (context) {
            const range = context.document.getSelection();
            // InsertLocation.replace akan menimpa teks yang sedang disorot (jika ada)
            range.insertText(char, Word.InsertLocation.replace); 
            return context.sync();
        }).catch(handleError);
    }

    // Menyisipkan Enter (Baris Baru)
    function insertEnter() {
        if (!isRunningInWord) {
            alert('Mode Preview: Menyisipkan Enter / Baris Baru');
            return;
        }

        Word.run(function (context) {
            const range = context.document.getSelection();
            // insertParagraph adalah cara paling tepat membuat baris baru di Word
            range.insertParagraph(); 
            return context.sync();
        }).catch(handleError);
    }

    // Menghapus teks (Backspace & Del)
    function deleteSelection() {
        if (!isRunningInWord) {
            alert('Mode Preview: Menghapus teks yang dipilih');
            return;
        }

        Word.run(function (context) {
            const range = context.document.getSelection();
            // Menghapus teks yang sedang disorot/diblok oleh kursor
            range.clear(); 
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
