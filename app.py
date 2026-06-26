import datetime
import io
import os
import re
import shutil
import subprocess
import tempfile

from docxtpl import DocxTemplate, Listing
from flask import Flask, abort, render_template, request, send_file

app = Flask(__name__)
BASE_DIR = os.path.dirname(__file__)

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# ──────────────────────────────────────────────────────────────
# 入力フォームページ
# ──────────────────────────────────────────────────────────────
PAGES = {
    "/":                    "top.html",
    "/teijuu":              "teijuu.html",
    "/shussan":             "shussan.html",
    "/shogai":              "shogai.html",
    "/noshin":              "noshin.html",
    "/suido":               "suido.html",
    "/kouminkan":           "kouminkan.html",
    "/kurashi_support":     "kurashi_support.html",
    "/ijuu_shien":          "ijuu_shien.html",
    "/kekkon_shinseikatsu": "kekkon_shinseikatsu.html",
    "/jutaku_konyuu":       "jutaku_konyuu.html",
    "/telework_ijuu":       "telework_ijuu.html",
    "/yytan_kotsuhi":       "yytan_kotsuhi.html",
}

for url, page in PAGES.items():
    app.add_url_rule(url, endpoint=f"page_{page}",
                     view_func=(lambda page=page: render_template(page)))


# ──────────────────────────────────────────────────────────────
# 和暦ユーティリティ
# ──────────────────────────────────────────────────────────────
ERA_BASE = {"令和": 2018, "平成": 1988, "昭和": 1925, "大正": 1911}
YOUBI = "月火水木金土日"


def wareki_youbi(wareki):
    """「令和8年6月15日」のような和暦文字列から曜日を返す。解釈できなければ空文字。"""
    m = re.match(r"(令和|平成|昭和|大正)(\d+)年(\d+)月(\d+)日", wareki.strip())
    if not m:
        return ""
    era, y, mo, d = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4))
    try:
        date = datetime.date(ERA_BASE[era] + y, mo, d)
    except ValueError:
        return ""
    return YOUBI[date.weekday()]


def to_number(value):
    try:
        return float(value.replace(",", "").strip())
    except (ValueError, AttributeError):
        return None


def fmt_number(num):
    return str(int(num)) if num == int(num) else str(num)


# ──────────────────────────────────────────────────────────────
# フォームごとの追加コンテキスト（チェックボックス・計算項目など）
# ──────────────────────────────────────────────────────────────
CK_KEYS = ("gozen", "encho1", "gogo", "encho2", "yakan")
TIME_CHECKS = {
    "午前":            {"gozen"},
    "午前・延長":      {"gozen", "encho1"},
    "午後":            {"gogo"},
    "午後・延長":      {"gogo", "encho2"},
    "夜間":            {"yakan"},
    "午前・午後":      {"gozen", "encho1", "gogo"},
    "午前・午後・夜間": {"gozen", "encho1", "gogo", "encho2", "yakan"},
}


def kouminkan_extra(form, ctx):
    # 入力済みの日程行だけを集めて1行目から詰める
    rows = []
    for n in range(1, 6):
        row = {k: form.get(f"{k}_{n}", "").strip() for k in ("date", "room", "time", "biko")}
        if any(row.values()):
            rows.append(row)
    for n in range(1, 6):
        row = rows[n - 1] if n <= len(rows) else {"date": "", "room": "", "time": "", "biko": ""}
        ctx[f"date_{n}"] = row["date"]
        ctx[f"room_{n}"] = row["room"]
        ctx[f"biko_{n}"] = row["biko"]
        ctx[f"youbi_{n}"] = wareki_youbi(row["date"])
        checked = TIME_CHECKS.get(row["time"], set())
        for key in CK_KEYS:
            ctx[f"ck_{n}_{key}"] = "☑" if key in checked else "□"
    ctx["houki_check"] = "☑" if form.get("houki_kakunin") else "□"


def noshin_extra(form, ctx):
    nai = to_number(form.get("kuiki_nai_menseki", ""))
    gai = to_number(form.get("kuiki_gai_menseki", ""))
    ctx["kuiki_kei"] = fmt_number((nai or 0) + (gai or 0)) if (nai is not None or gai is not None) else ""
    for field in ("ta", "hata"):
        genzai = to_number(form.get(f"{field}_genzai", ""))
        ato = to_number(form.get(f"{field}_ato", ""))
        if genzai is not None and ato is not None:
            diff = ato - genzai
            ctx[f"{field}_zougen"] = f"△{fmt_number(abs(diff))}" if diff < 0 else fmt_number(diff)
        else:
            ctx[f"{field}_zougen"] = ""


def shogai_extra(form, ctx):
    teiin = form.get("nyusho_teiin", "").strip()
    if teiin and not teiin.endswith("人"):
        ctx["nyusho_teiin"] = teiin + "人"


# ──────────────────────────────────────────────────────────────
# 申請書定義：新しい様式はここに1件追加するだけ
#   template   : docxテンプレートのファイル名
#   filename   : ダウンロードファイル名の接頭辞
#   name_fields: ファイル名に使う申請者名（先に見つかった非空値を使用）
#   extra      : 追加コンテキストを組み立てる関数（任意）
# ──────────────────────────────────────────────────────────────
FORMS = {
    "teijuu": {
        "template": "teijuu_template.docx",
        "filename": "定住奨励金支給申請書",
        "name_fields": ["shimei"],
    },
    "shussan": {
        "template": "shussan_template.docx",
        "filename": "出産祝い金支給申請書",
        "name_fields": ["shimei"],
    },
    "shogai": {
        "template": "shogai_jigyo_template.docx",
        "filename": "障害福祉サービス事業等開始届",
        "name_fields": ["meisho"],
        "extra": shogai_extra,
    },
    "noshin": {
        "template": "noshin_moushide_template.docx",
        "filename": "農業振興地域整備計画変更申出書",
        "name_fields": ["shinshutsu_shimei"],
        "extra": noshin_extra,
    },
    "suido": {
        "template": "suido_syoyusha_template.docx",
        "filename": "排水設備所有者変更届",
        "name_fields": ["ato_shimei"],
    },
    "kouminkan_shiyo": {
        "template": "kouminkan_shiyo_template.docx",
        "filename": "公民館使用許可申請書",
        "name_fields": ["dantai_name", "daihyo_name"],
        "extra": kouminkan_extra,
    },
    "kouminkan_genmen": {
        "template": "kouminkan_genmen_template.docx",
        "filename": "公民館使用料減免申請書",
        "name_fields": ["dantai_name", "daihyo_name"],
        "extra": kouminkan_extra,
    },
    "kouminkan_chushi": {
        "template": "kouminkan_chushi_template.docx",
        "filename": "公民館使用中止届",
        "name_fields": ["dantai_name", "daihyo_name"],
        "extra": kouminkan_extra,
    },
    "kurashi_support": {
        "template": "kurashi_support_template.docx",
        "filename": "暮らしサポート補助金交付申請書",
        "name_fields": ["shimei"],
    },
    "ijuu_shien": {
        "template": "ijuu_shien_template.docx",
        "filename": "移住支援金支給申請書",
        "name_fields": ["shimei"],
    },
    "kekkon_shinseikatsu": {
        "template": "kekkon_shinseikatsu_template.docx",
        "filename": "結婚新生活支援補助金交付申請書",
        "name_fields": ["otto_shimei", "tsuma_shimei"],
    },
    "jutaku_konyuu": {
        "template": "jutaku_konyuu_template.docx",
        "filename": "移住者向け住宅購入支援事業補助金交付申請書",
        "name_fields": ["shimei"],
    },
    "telework_ijuu": {
        "template": "telework_ijuu_template.docx",
        "filename": "やまぐち創生テレワーク移住支援事業補助金交付申請書",
        "name_fields": ["shimei"],
    },
    "yytan_kotsuhi": {
        "template": "yytan_kotsuhi_template.docx",
        "filename": "YYターン支援交通費補助金申請書",
        "name_fields": ["shimei"],
    },
}


# ──────────────────────────────────────────────────────────────
# PDF変換（LibreOffice）
# ──────────────────────────────────────────────────────────────
def find_soffice():
    for cand in ("soffice", "libreoffice"):
        path = shutil.which(cand)
        if path:
            return path
    for path in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(path):
            return path
    return None


def docx_to_pdf(docx_bytes):
    soffice = find_soffice()
    if soffice is None:
        raise RuntimeError("LibreOffice (soffice) が見つかりません")
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "output.docx")
        with open(src, "wb") as f:
            f.write(docx_bytes)
        profile = "file:///" + os.path.join(tmp, "lo").replace(os.sep, "/")
        subprocess.run(
            [soffice, "--headless", "--norestore", f"-env:UserInstallation={profile}",
             "--convert-to", "pdf", "--outdir", tmp, src],
            check=True, timeout=120,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        with open(os.path.join(tmp, "output.pdf"), "rb") as f:
            return f.read()


@app.context_processor
def inject_pdf_enabled():
    return {"pdf_enabled": find_soffice() is not None}


# ──────────────────────────────────────────────────────────────
# 申請書生成
# ──────────────────────────────────────────────────────────────
def build_context(cfg, form):
    ctx = {k: v for k, v in form.items() if k != "fmt"}
    if "extra" in cfg:
        cfg["extra"](form, ctx)
    # 複数行入力はWord内でも改行されるように Listing でラップ
    def wrap(value):
        if isinstance(value, str):
            value = value.replace("\r\n", "\n").replace("\r", "\n")
            if "\n" in value:
                return Listing(value)
        return value

    return {k: wrap(v) for k, v in ctx.items()}


def safe_name(cfg, form):
    for field in cfg["name_fields"]:
        value = form.get(field, "").strip()
        if value:
            return re.sub(r'[\\/:*?"<>|\s　]+', "_", value)
    return "申請者"


@app.route("/generate/<form_id>", methods=["POST"])
def generate(form_id):
    cfg = FORMS.get(form_id)
    if cfg is None:
        abort(404)

    doc = DocxTemplate(os.path.join(BASE_DIR, cfg["template"]))
    doc.render(build_context(cfg, request.form))
    buffer = io.BytesIO()
    doc.save(buffer)

    filename = f"{cfg['filename']}_{safe_name(cfg, request.form)}"
    if request.form.get("fmt") == "pdf":
        try:
            pdf_bytes = docx_to_pdf(buffer.getvalue())
        except (RuntimeError, subprocess.SubprocessError):
            return ("PDF変換は現在利用できません。Word形式をご利用ください。", 503)
        return send_file(io.BytesIO(pdf_bytes), as_attachment=True,
                         download_name=f"{filename}.pdf", mimetype="application/pdf")

    buffer.seek(0)
    return send_file(buffer, as_attachment=True,
                     download_name=f"{filename}.docx", mimetype=DOCX_MIME)


# 旧URL（ブックマーク・キャッシュ済みページ用）
LEGACY_ENDPOINTS = {
    "/generate":                   "teijuu",
    "/generate_shussan":           "shussan",
    "/generate_shogai_jigyo":      "shogai",
    "/generate_noshin_moushide":   "noshin",
    "/generate_suido_syoyusha":    "suido",
    "/generate_kouminkan_shiyo":   "kouminkan_shiyo",
    "/generate_kouminkan_genmen":  "kouminkan_genmen",
    "/generate_kouminkan_chushi":  "kouminkan_chushi",
}

for url, form_id in LEGACY_ENDPOINTS.items():
    app.add_url_rule(url, endpoint=f"legacy_{form_id}", methods=["POST"],
                     view_func=(lambda form_id=form_id: generate(form_id)))


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
