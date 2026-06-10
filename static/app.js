// 全フォーム共通のユーティリティ

// 月・日のセレクトボックスに選択肢を生成
function fillMonthDay() {
  document.querySelectorAll('.month-sel').forEach(s => { for (let m = 1; m <= 12; m++) s.appendChild(new Option(m, m)); });
  document.querySelectorAll('.day-sel').forEach(s => { for (let d = 1; d <= 31; d++) s.appendChild(new Option(d, d)); });
}

// 和暦パーツを「令和8年6月15日」の形式に結合
function combineDate(eraId, yearId, monthId, dayId) {
  const era = document.getElementById(eraId).value,
        year = document.getElementById(yearId).value,
        month = document.getElementById(monthId).value,
        day = document.getElementById(dayId).value;
  return (year && month && day) ? `${era}${year}年${month}月${day}日` : '';
}

// 年・月・日の入力欄に今日の日付（令和）を初期セット
function setToday(yearId, monthId, dayId) {
  const today = new Date();
  document.getElementById(yearId).value = today.getFullYear() - 2018;
  document.getElementById(monthId).value = today.getMonth() + 1;
  document.getElementById(dayId).value = today.getDate();
}

// 郵便番号 → 住所検索（zipcloud API）
async function searchZip(zipInputId, addressInputId) {
  const zipEl = document.getElementById(zipInputId),
        msgEl = document.getElementById(zipInputId + '_msg'),
        addrEl = document.getElementById(addressInputId);
  const zip = zipEl.value.replace(/-/g, '').replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  if (zip.length !== 7 || !/^\d{7}$/.test(zip)) {
    msgEl.textContent = '7桁の数字で入力してください'; msgEl.className = 'zip-msg err'; return;
  }
  msgEl.textContent = '検索中...'; msgEl.className = 'zip-msg loading';
  try {
    const res = await fetch(`https://zipcloud.ibsnet.co.jp/api/search?zipcode=${zip}`);
    const data = await res.json();
    if (data.results && data.results.length > 0) {
      const r = data.results[0];
      addrEl.value = r.address1 + r.address2 + r.address3;
      msgEl.textContent = '住所が見つかりました。番地・部屋番号を追記してください。'; msgEl.className = 'zip-msg ok';
    } else {
      msgEl.textContent = '該当する住所が見つかりませんでした。手動で入力してください。'; msgEl.className = 'zip-msg err';
    }
  } catch {
    msgEl.textContent = '検索に失敗しました。手動で入力してください。'; msgEl.className = 'zip-msg err';
  }
}

// 郵便番号欄に自動検索とEnterキー検索をまとめて設定する
// 例： bindZip({ zip1: 'jusho_input', zip2: 'mae_jusho_input' })
function bindZip(map) {
  for (const [zipId, addrId] of Object.entries(map)) {
    const el = document.getElementById(zipId);
    if (!el) continue;
    el.addEventListener('input', e => { if (e.target.value.replace(/-/g, '').length === 7) searchZip(zipId, addrId); });
    el.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); searchZip(zipId, addrId); } });
  }
}

// 法人格（前株・後株）＋勤務先名（定住・出産フォームで使用）
let currentPos = 'mae';

function setPos(pos) {
  currentPos = pos;
  document.getElementById('lbl_mae').classList.toggle('active', pos === 'mae');
  document.getElementById('lbl_ato').classList.toggle('active', pos === 'ato');
  updateCompanyPreview();
}

function buildCompanyName() {
  const type = document.getElementById('hojin_type').value;
  const name = document.getElementById('kinmusaki_name').value.trim();
  if (!type) return name;
  return currentPos === 'mae' ? type + name : name + type;
}

function updateCompanyPreview() {
  const type = document.getElementById('hojin_type').value;
  const name = document.getElementById('kinmusaki_name').value.trim();
  const preview = document.getElementById('company_preview');
  const toggle = document.getElementById('pos_toggle');
  if (toggle) {
    toggle.style.opacity = type ? '1' : '0.35';
    toggle.style.pointerEvents = type ? 'auto' : 'none';
  }
  if (!name && !type) {
    preview.innerHTML = '<span>法人格と会社名を入力するとプレビューが表示されます</span>';
    return;
  }
  preview.textContent = '▶ ' + buildCompanyName();
}
