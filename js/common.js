/**
 * 依頼事項 期限管理システム - 共通JavaScript
 */

// 削除確認
function confirmDelete(message) {
    return confirm(message || '本当に削除しますか？');
}

// フォーム送信前の確認
function confirmSubmit(message) {
    return confirm(message || '登録してよろしいですか？');
}

// 入力チェック：必須項目
function validateRequired(form) {
    var requiredFields = form.querySelectorAll('[required]');
    var isValid = true;

    requiredFields.forEach(function(field) {
        if (!field.value.trim()) {
            field.classList.add('error');
            isValid = false;
        } else {
            field.classList.remove('error');
        }
    });

    if (!isValid) {
        alert('必須項目を入力してください。');
    }

    return isValid;
}

// 日付の妥当性チェック
function validateDateRange(startDate, endDate) {
    if (startDate && endDate) {
        var start = new Date(startDate);
        var end = new Date(endDate);
        if (start > end) {
            alert('終了日は開始日以降の日付を指定してください。');
            return false;
        }
    }
    return true;
}

// ファイルアップロードのドラッグ＆ドロップ対応
function initFileUpload() {
    var dropZone = document.querySelector('.file-upload');
    if (!dropZone) return;

    var fileInput = dropZone.querySelector('input[type="file"]');

    // クリックでファイル選択
    dropZone.addEventListener('click', function() {
        fileInput.click();
    });

    // ドラッグオーバー
    dropZone.addEventListener('dragover', function(e) {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    // ドラッグリーブ
    dropZone.addEventListener('dragleave', function(e) {
        e.preventDefault();
        dropZone.classList.remove('dragover');
    });

    // ドロップ
    dropZone.addEventListener('drop', function(e) {
        e.preventDefault();
        dropZone.classList.remove('dragover');

        var files = e.dataTransfer.files;
        if (files.length > 0) {
            fileInput.files = files;
            updateFileName(fileInput);
        }
    });

    // ファイル選択時
    fileInput.addEventListener('change', function() {
        updateFileName(this);
    });
}

// 選択されたファイル名を表示
function updateFileName(input) {
    var dropZone = input.closest('.file-upload');
    var fileNameSpan = dropZone.querySelector('.file-name');

    if (input.files.length > 0) {
        fileNameSpan.textContent = input.files[0].name;
    } else {
        fileNameSpan.textContent = 'ファイルを選択またはドラッグ＆ドロップ';
    }
}

// テーブルのソート
function sortTable(tableId, columnIndex) {
    var table = document.getElementById(tableId);
    var tbody = table.querySelector('tbody');
    var rows = Array.from(tbody.querySelectorAll('tr'));

    var currentSort = table.getAttribute('data-sort-column');
    var currentDir = table.getAttribute('data-sort-dir') || 'asc';
    var newDir = (currentSort == columnIndex && currentDir === 'asc') ? 'desc' : 'asc';

    rows.sort(function(a, b) {
        var aVal = a.cells[columnIndex].textContent.trim();
        var bVal = b.cells[columnIndex].textContent.trim();

        // 数値判定
        if (!isNaN(aVal) && !isNaN(bVal)) {
            return newDir === 'asc' ? aVal - bVal : bVal - aVal;
        }

        // 日付判定
        var aDate = Date.parse(aVal);
        var bDate = Date.parse(bVal);
        if (!isNaN(aDate) && !isNaN(bDate)) {
            return newDir === 'asc' ? aDate - bDate : bDate - aDate;
        }

        // 文字列比較
        return newDir === 'asc'
            ? aVal.localeCompare(bVal, 'ja')
            : bVal.localeCompare(aVal, 'ja');
    });

    // 行を並び替え
    rows.forEach(function(row) {
        tbody.appendChild(row);
    });

    table.setAttribute('data-sort-column', columnIndex);
    table.setAttribute('data-sort-dir', newDir);
}

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', function() {
    initFileUpload();
});
