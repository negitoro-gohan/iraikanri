// ============================================
// 社員検索モーダル・社員コードルックアップ共通スクリプト
// ============================================

// --------------------------------------------
// 社員コードから社員名を検索して表示する（既存lookupEmployee共通化）
// --------------------------------------------
function lookupEmployee(type) {
    var codeInput, nameDisplay;

    if (type === 'requester') {
        codeInput = document.getElementById('requester_code');
        nameDisplay = document.getElementById('requester_name_display');
    } else {
        codeInput = document.getElementById('assignee_code');
        nameDisplay = document.getElementById('assignee_name_display');
    }

    var code = codeInput.value.trim();
    if (code === '') {
        nameDisplay.textContent = '';
        nameDisplay.className = 'employee-name-display';
        return;
    }

    // AJAX呼び出し
    var xhr = new XMLHttpRequest();
    xhr.open('GET', 'api_employee_lookup.asp?code=' + encodeURIComponent(code), true);
    xhr.onreadystatechange = function() {
        if (xhr.readyState === 4 && xhr.status === 200) {
            try {
                var result = JSON.parse(xhr.responseText);
                if (result.found) {
                    nameDisplay.textContent = result.email
                        ? result.name + '（' + result.email + '）'
                        : result.name;
                    nameDisplay.className = 'employee-name-display employee-name-found';
                } else {
                    nameDisplay.textContent = '該当なし';
                    nameDisplay.className = 'employee-name-display employee-name-notfound';
                }
            } catch (e) {
                nameDisplay.textContent = '検索エラー';
                nameDisplay.className = 'employee-name-display employee-name-notfound';
            }
        }
    };
    xhr.send();
}

// --------------------------------------------
// ページ読み込み時、入力済みの社員コードがあれば検索実行
// --------------------------------------------
document.addEventListener('DOMContentLoaded', function() {
    if (document.getElementById('requester_code') &&
        document.getElementById('requester_code').value.trim() !== '') {
        lookupEmployee('requester');
    }
    if (document.getElementById('assignee_code') &&
        document.getElementById('assignee_code').value.trim() !== '') {
        lookupEmployee('assignee');
    }
});

// --------------------------------------------
// モーダルHTML動的生成（初回のみ）
// --------------------------------------------
var employeeSearchModalCreated = false;

function ensureEmployeeSearchModal() {
    if (employeeSearchModalCreated) return;
    employeeSearchModalCreated = true;

    var overlay = document.createElement('div');
    overlay.id = 'empSearchOverlay';
    overlay.className = 'emp-search-overlay';
    // インラインスタイルで確実にモーダルオーバーレイとして表示
    overlay.style.cssText = 'display:none;position:fixed;top:0;left:0;width:100%;height:100%;background-color:rgba(0,0,0,0.5);z-index:1000;justify-content:center;align-items:center;';
    overlay.onclick = function(e) {
        if (e.target === overlay) closeEmployeeSearch();
    };

    var modal = document.createElement('div');
    modal.className = 'emp-search-modal';
    modal.style.cssText = 'background-color:#fff;border-radius:8px;box-shadow:0 4px 20px rgba(0,0,0,0.3);width:600px;max-width:90vw;max-height:80vh;display:flex;flex-direction:column;';

    modal.innerHTML =
        '<div style="display:flex;justify-content:space-between;align-items:center;padding:16px 20px;border-bottom:1px solid #eee;">' +
            '<h3 style="margin:0;font-size:18px;">社員検索</h3>' +
            '<button type="button" style="background:none;border:none;font-size:24px;cursor:pointer;color:#999;padding:0 5px;line-height:1;" onclick="closeEmployeeSearch()">&times;</button>' +
        '</div>' +
        '<div style="padding:20px;overflow-y:auto;flex:1;">' +
            '<div style="display:flex;gap:10px;margin-bottom:15px;">' +
                '<input type="text" id="empSearchKeyword" class="form-control" placeholder="社員コードまたは社員名を入力" style="flex:1;">' +
                '<button type="button" class="btn btn-primary" onclick="executeEmployeeSearch()">検索</button>' +
            '</div>' +
            '<div id="empSearchResult" style="min-height:100px;max-height:400px;overflow-y:auto;">' +
                '<p style="color:#999;text-align:center;padding:30px 0;">キーワードを入力して検索してください。</p>' +
            '</div>' +
        '</div>';

    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    // Enterキーで検索実行
    document.getElementById('empSearchKeyword').addEventListener('keydown', function(e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            executeEmployeeSearch();
        }
    });
}

// --------------------------------------------
// モーダルを開く
// targetField: 'requester' or 'assignee'
// --------------------------------------------
var currentTargetField = '';

function openEmployeeSearch(targetField) {
    currentTargetField = targetField;
    ensureEmployeeSearchModal();

    // 検索欄をリセット
    document.getElementById('empSearchKeyword').value = '';

    // モーダル表示
    document.getElementById('empSearchOverlay').style.display = 'flex';

    // 検索欄にフォーカスしてから全件検索を実行
    setTimeout(function() {
        document.getElementById('empSearchKeyword').focus();
        executeEmployeeSearch(); // 空キーワードで即時全件表示
    }, 100);
}

// --------------------------------------------
// モーダルを閉じる
// --------------------------------------------
function closeEmployeeSearch() {
    document.getElementById('empSearchOverlay').style.display = 'none';
}

// --------------------------------------------
// AJAX検索実行
// --------------------------------------------
function executeEmployeeSearch() {
    var keyword = document.getElementById('empSearchKeyword').value.trim();
    var resultDiv = document.getElementById('empSearchResult');

    // 空キーワードは全件検索（APIが全件返す）
    resultDiv.innerHTML = '<p style="color:#999;text-align:center;padding:30px 0;">検索中...</p>';

    var xhr = new XMLHttpRequest();
    xhr.open('GET', 'employee_search.asp?keyword=' + encodeURIComponent(keyword), true);
    xhr.onreadystatechange = function() {
        if (xhr.readyState === 4) {
            if (xhr.status === 200) {
                try {
                    var employees = JSON.parse(xhr.responseText);
                    renderEmployeeSearchResult(employees);
                } catch (e) {
                    resultDiv.innerHTML = '<p class="emp-search-hint" style="color:#c62828;">検索エラーが発生しました。</p>';
                }
            } else {
                resultDiv.innerHTML = '<p class="emp-search-hint" style="color:#c62828;">検索エラーが発生しました。</p>';
            }
        }
    };
    xhr.send();
}

// --------------------------------------------
// 検索結果テーブルを描画
// --------------------------------------------
function renderEmployeeSearchResult(employees) {
    var resultDiv = document.getElementById('empSearchResult');

    if (employees.length === 0) {
        resultDiv.innerHTML = '<p style="color:#999;text-align:center;padding:30px 0;">該当する社員が見つかりません。</p>';
        return;
    }

    var html = '<table class="data-table emp-search-table">' +
        '<thead><tr>' +
            '<th>社員コード</th>' +
            '<th>社員名</th>' +
            '<th>メール</th>' +
        '</tr></thead><tbody>';

    for (var i = 0; i < employees.length; i++) {
        var emp = employees[i];
        html += '<tr class="emp-search-row" onclick="selectEmployee(\'' +
            escapeHtmlAttr(emp.employee_code) + '\',\'' +
            escapeHtmlAttr(emp.employee_name) + '\',\'' +
            escapeHtmlAttr(emp.email || '') + '\')">' +
            '<td>' + escapeHtml(emp.employee_code) + '</td>' +
            '<td>' + escapeHtml(emp.employee_name) + '</td>' +
            '<td>' + escapeHtml(emp.email || '') + '</td>' +
            '</tr>';
    }

    html += '</tbody></table>';
    resultDiv.innerHTML = html;
}

// --------------------------------------------
// 社員を選択 → テキストボックスにセット＋モーダルを閉じる
// --------------------------------------------
function selectEmployee(code, name, email) {
    var codeInput, nameDisplay;

    if (currentTargetField === 'requester') {
        codeInput = document.getElementById('requester_code');
        nameDisplay = document.getElementById('requester_name_display');
    } else {
        codeInput = document.getElementById('assignee_code');
        nameDisplay = document.getElementById('assignee_name_display');
    }

    // 社員コードをセット
    codeInput.value = code;

    // 社員名とメールアドレスを表示
    nameDisplay.textContent = email ? name + '（' + email + '）' : name;
    nameDisplay.className = 'employee-name-display employee-name-found';

    // モーダルを閉じる
    closeEmployeeSearch();
}

// --------------------------------------------
// HTMLエスケープユーティリティ
// --------------------------------------------
function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, '&amp;')
              .replace(/</g, '&lt;')
              .replace(/>/g, '&gt;')
              .replace(/"/g, '&quot;');
}

function escapeHtmlAttr(str) {
    if (!str) return '';
    return str.replace(/\\/g, '\\\\')
              .replace(/'/g, "\\'")
              .replace(/"/g, '\\"');
}
