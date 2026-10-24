'use strict';

let isRunningInWord = false;

(function () {
    Office.onReady(function (info) {
        $(document).ready(function () {
            
            // 1. Cek Lingkungan (Word vs Browser)
            if (info.host === Office.HostType.Word && Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                isRunningInWord = true;
                $('#supportedVersion').html('Berjalan di Microsoft Word.');
                
                // Sembunyikan tombol Panduan & Download di dalam Word
                $('#info-btn').hide(); 
                $('#download-btn').hide(); 
            } else {
                $('#supportedVersion').html('Mode Web: Klik ikon (ⓘ) untuk panduan install.');
                $('#download-btn').show();
            }

            // 2. Daftarkan Event Tombol
            bindButtonEvents();

            // 3. Logika Modal Tutorial
            $('#info-btn').click(function() {
                $('#tutorial-modal').fadeIn(200);
            });

            $('.close-btn').click(function() {
                $('#tutorial-modal').fadeOut(200);
            });

            $(window).click(function(event) {
                if ($(event.target).is('#tutorial-modal')) {
                    $('#tutorial-modal').fadeOut(200);
                }
            });

        });
    });

    // ==========================================
    // FUNGSI MENDAFTARKAN TOMBOL
    // ==========================================
    function bindButtonEvents() {
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

        for (const [id, char] of Object.entries(charMap)) {
            $('#' + id).click(function () {
                insertCharacter(char);
            });
        }

        $('#btn-space').click(function () { insertCharacter(' '); });
        $('#btn-enter').click(insertEnter);
        $('#btn-backspace').click(deleteSelection);
        $('#btn-del').click(deleteSelection);
    }

    // ==========================================
    // FUNGSI UTAMA
    // ==========================================

    function insertCharacter(char) {
        if (!isRunningInWord) {
            alert('Mode Preview: Menyisipkan karakter "' + char + '"');
            return;
        }

        Word.run(function (context) {
            const range = context.document.getSelection();
            
            // Sisipkan teks
            range.insertText(char, Word.InsertLocation.replace);
            
            // KUNCI PERBAIKAN: Paksa kursor pindah ke ujung kanan setelah mengetik
            range.select(Word.InsertLocation.end);
            
            return context.sync();
        }).catch(handleError);
    }

    function insertEnter() {
        if (!isRunningInWord) {
            alert('Mode Preview: Menyisipkan Enter');
            return;
        }

        Word.run(function (context) {
            const range = context.document.getSelection();
            range.insertParagraph(); 
            return context.sync();
        }).catch(handleError);
    }

    function deleteSelection() {
        if (!isRunningInWord) {
            alert('Mode Preview: Menghapus teks');
            return;
        }

        Word.run(function (context) {
            const range = context.document.getSelection();
            range.clear(); 
            return context.sync();
        }).catch(handleError);
    }

    function handleError(error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    }

})();
