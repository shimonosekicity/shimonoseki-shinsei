# -*- coding: utf-8 -*-
"""全ページ表示と全申請書の生成をFlaskテストクライアントで検証する。"""
import io
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.stdout.reconfigure(encoding="utf-8")

from app import app, PAGES

client = app.test_client()

# ── ページ表示 ────────────────────────────────────────────
for url in PAGES:
    res = client.get(url)
    assert res.status_code == 200, f"{url}: {res.status_code}"
print("pages ok:", ", ".join(PAGES))

# ── 申請書生成 ────────────────────────────────────────────
KOUMINKAN_BASE = {
    "shinsei_date": "令和8年6月10日", "kouminkan_name": "彦島",
    "jusho": "下関市彦島1-1", "dantai_name": "テスト会", "daihyo_name": "会長 山田太郎",
    "denwa": "083-000-0000", "shiyou_mokuteki": "料理教室", "shiyou_ninzu": "20",
    "date_1": "令和8年6月15日", "room_1": "会議室", "time_1": "午前・延長", "biko_1": "定例会",
    "date_2": "令和8年6月22日", "room_2": "和室", "time_2": "夜間", "biko_2": "",
    "houki_kakunin": "1",
}

CASES = {
    "teijuu": {
        "shinsei_date": "令和8年6月10日", "shimei": "山田太郎", "jusho": "下関市1-1",
        "denwa": "083-000-0000", "seibetsu": "男", "seinengappi": "昭和60年1月1日",
        "nenrei": "41", "kinmusaki": "株式会社テスト", "shokushu": "営業",
        "tennyu_date": "令和7年4月1日", "teijuu_date": "令和7年4月1日",
        "tennyu_mae_jusho": "福岡市1-1",
    },
    "shussan": {
        "shinsei_date": "令和8年6月10日", "shimei": "山田花子", "jusho": "下関市1-1",
        "denwa": "083-000-0000", "shussei_date": "令和8年1月1日", "dai_ko": "1",
        "shussei_furigana": "ヤマダ ハナコ", "shussei_shimei": "山田 花子",
        "yoikusha_shimei": "山田太郎", "genjusho": "下関市1-1", "kinmusaki": "株式会社テスト",
    },
    "shogai": {
        "todoke_nen": "8", "todoke_tsuki": "6", "todoke_hi": "10",
        "shozaichi": "下関市1-1", "meisho": "社会福祉法人テスト会", "daihyo": "理事長 山田",
        "tel": "083-000-0000", "jigyo_shurui": "障害福祉サービス事業",
        "jigyo_naiyou": "居宅介護\n対象：身体障害者", "jigyo_kuiki": "下関市全域",
        "keieisha_shimei": "社会福祉法人テスト会", "keieisha_jusho": "下関市1-1",
        "shisetsu_meisho": "テストホーム", "shisetsu_shurui": "グループホーム",
        "shisetsu_shozaichi": "下関市2-2", "nyusho_teiin": "10",
        "kaishi_nen": "8", "kaishi_tsuki": "7", "kaishi_hi": "1",
    },
    "noshin": {
        "moushide_date": "令和8年6月10日", "shinshutsu_zip": "7500031",
        "shinshutsu_jusho": "下関市1-1", "shinshutsu_tel": "083-000-0000",
        "shinshutsu_shimei": "山田太郎", "oaza": "彦島", "aza": "江の浦",
        "chiban": "123-4", "chimoku": "田", "menseki": "500", "shoyu": "下関市1-1 山田太郎",
        "henko_riyu": "自宅建築のため\n早急に必要", "sentei_riyu": "代替地なし",
        "nochi_jokyo": "田：500kg", "koko_toshi": "無", "osui_shori": "合併浄化槽",
        "nicchou": "影響なし", "sonota_eikyou": "なし",
        "shutai_shimei": "山田太郎", "shutai_jusho": "下関市1-1", "shutai_kankei": "地権者本人",
        "kuiki_nai_menseki": "500", "kuiki_gai_menseki": "100",
        "shisetsu_menseki": "木造2階建80㎡", "jigyo_hi": "15,000,000",
        "ta_genzai": "1000", "ta_ato": "500", "hata_genzai": "300", "hata_ato": "300",
        "senkyo_kenbyo": "専業",
    },
    "suido": {
        "todoke_date": "令和8年6月10日", "shutsu_jusho": "下関市1-1", "shutsu_shimei": "山田太郎",
        "settichi": "竹崎町1-1", "settichi_tatemono": "テストマンション101",
        "jyoto_date": "令和8年5月1日", "mae_jusho": "下関市1-1", "mae_shimei": "山田太郎",
        "ato_jusho": "下関市2-2", "ato_shimei": "山田花子", "jyoto_riyu": "売買による所有権移転",
    },
    "kouminkan_shiyo": KOUMINKAN_BASE,
    "kouminkan_genmen": KOUMINKAN_BASE,
    "kouminkan_chushi": KOUMINKAN_BASE,
}

# 生成されたdocxに含まれているべき文字列
EXPECT = {
    "teijuu": ["山田太郎", "株式会社テスト", "福岡市1-1"],
    "shussan": ["ヤマダ ハナコ", "山田 花子"],
    "shogai": ["居宅介護", "下関市全域", "テストホーム", "グループホーム", "10人", "社会福祉法人テスト会"],
    "noshin": ["自宅建築のため", "代替地なし", "合併浄化槽", "山田太郎", "地権者本人",
               "15,000,000", "600", "△500", "専業"],
    "suido": ["山田太郎", "山田花子", "売買による所有権移転", "テストマンション101"],
    "kouminkan_shiyo": ["令和8年6月15日", "会議室", "定例会", "和室", "☑午前", "☑夜間", "料理教室"],
    "kouminkan_genmen": ["令和8年6月15日", "会議室", "テスト会", "料理教室"],
    "kouminkan_chushi": ["令和8年6月15日", "会議室", "和室"],
}


def docx_text(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    return re.sub(r"<[^>]+>", "", xml)


failures = []
for form_id, payload in CASES.items():
    res = client.post(f"/generate/{form_id}", data=payload)
    if res.status_code != 200:
        failures.append(f"{form_id}: HTTP {res.status_code}")
        continue
    text = docx_text(res.data)
    leftover = re.findall(r"\{\{[^}]*\}?\}?", text)
    if leftover:
        failures.append(f"{form_id}: unrendered placeholders {leftover[:5]}")
    for expected in EXPECT[form_id]:
        if expected not in text:
            failures.append(f"{form_id}: missing {expected!r}")
    print(f"generate ok: {form_id} ({len(res.data)} bytes)")

# ── 旧URL互換 ────────────────────────────────────────────
res = client.post("/generate", data=CASES["teijuu"])
assert res.status_code == 200, f"legacy /generate: {res.status_code}"
print("legacy ok: /generate")

# ── PDF（LibreOffice無し環境では503） ─────────────────────
res = client.post("/generate/teijuu", data={**CASES["teijuu"], "fmt": "pdf"})
print(f"pdf request: HTTP {res.status_code} "
      f"({'PDF生成' if res.status_code == 200 else res.get_data(as_text=True).strip()})")

if failures:
    print("\nFAILURES:")
    for f in failures:
        print(" -", f)
    sys.exit(1)
print("\nALL OK")
