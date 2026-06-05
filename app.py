import io
import os
from docxtpl import DocxTemplate
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

TEIJUU_TEMPLATE   = os.path.join(os.path.dirname(__file__), "teijuu_template.docx")
SHUSSAN_TEMPLATE  = os.path.join(os.path.dirname(__file__), "shussan_template.docx")
SUIDO_SYOYUSHA    = os.path.join(os.path.dirname(__file__), "suido_syoyusha_template.docx")
NOSHIN_MOUSHIDE   = os.path.join(os.path.dirname(__file__), "noshin_moushide_template.docx")
SHOGAI_JIGYO      = os.path.join(os.path.dirname(__file__), "shogai_jigyo_template.docx")
KOUMINKAN_SHIYO   = os.path.join(os.path.dirname(__file__), "kouminkan_shiyo_template.docx")
KOUMINKAN_GENMEN  = os.path.join(os.path.dirname(__file__), "kouminkan_genmen_template.docx")
KOUMINKAN_CHUSHI  = os.path.join(os.path.dirname(__file__), "kouminkan_chushi_template.docx")


@app.route("/", methods=["GET"])
def top():
    return render_template("top.html")


@app.route("/teijuu", methods=["GET"])
def teijuu():
    return render_template("index.html")


@app.route("/shussan", methods=["GET"])
def shussan():
    return render_template("shussan.html")


@app.route("/generate", methods=["POST"])
def generate():
    context = {
        "shinsei_date":     request.form.get("shinsei_date", ""),
        "shimei":           request.form.get("shimei", ""),
        "jusho":            request.form.get("jusho", ""),
        "denwa":            request.form.get("denwa", ""),
        "seibetsu":         request.form.get("seibetsu", ""),
        "seinengappi":      request.form.get("seinengappi", ""),
        "nenrei":           request.form.get("nenrei", ""),
        "kinmusaki":        request.form.get("kinmusaki", ""),
        "shokushu":         request.form.get("shokushu", ""),
        "tennyu_date":      request.form.get("tennyu_date", ""),
        "teijuu_date":      request.form.get("teijuu_date", ""),
        "tennyu_mae_jusho": request.form.get("tennyu_mae_jusho", ""),
    }
    doc = DocxTemplate(TEIJUU_TEMPLATE)
    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    safe_name = context["shimei"].replace(" ", "_").replace("　", "_")
    return send_file(buffer, as_attachment=True,
                     download_name=f"定住奨励金支給申請書_{safe_name}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/generate_shussan", methods=["POST"])
def generate_shussan():
    context = {
        "shinsei_date":    request.form.get("shinsei_date", ""),
        "shimei":          request.form.get("shimei", ""),
        "jusho":           request.form.get("jusho", ""),
        "denwa":           request.form.get("denwa", ""),
        "shussei_date":    request.form.get("shussei_date", ""),
        "dai_ko":          request.form.get("dai_ko", ""),
        "shussei_furigana": request.form.get("shussei_furigana", ""),
        "shussei_shimei":  request.form.get("shussei_shimei", ""),
        "yoikusha_shimei": request.form.get("yoikusha_shimei", ""),
        "genjusho":        request.form.get("genjusho", ""),
        "kinmusaki":       request.form.get("kinmusaki", ""),
    }
    doc = DocxTemplate(SHUSSAN_TEMPLATE)
    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    safe_name = context["shimei"].replace(" ", "_").replace("　", "_")
    return send_file(buffer, as_attachment=True,
                     download_name=f"出産祝い金支給申請書_{safe_name}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/shogai", methods=["GET"])
def shogai():
    return render_template("shogai.html")


@app.route("/generate_shogai_jigyo", methods=["POST"])
def generate_shogai_jigyo():
    context = {
        "todoke_nen":   request.form.get("todoke_nen", ""),
        "todoke_tsuki": request.form.get("todoke_tsuki", ""),
        "todoke_hi":    request.form.get("todoke_hi", ""),
        "shozaichi":    request.form.get("shozaichi", ""),
        "meisho":       request.form.get("meisho", ""),
        "daihyo":       request.form.get("daihyo", ""),
        "tel":          request.form.get("tel", ""),
        "kaishi_nen":        request.form.get("kaishi_nen", ""),
        "kaishi_tsuki":      request.form.get("kaishi_tsuki", ""),
        "kaishi_hi":         request.form.get("kaishi_hi", ""),
        "jigyo_shurui":      request.form.get("jigyo_shurui", ""),
        "jigyo_naiyou":      request.form.get("jigyo_naiyou", ""),
        "keieisha_shimei":   request.form.get("keieisha_shimei", ""),
        "keieisha_jusho":    request.form.get("keieisha_jusho", ""),
        "shisetsu_meisho":   request.form.get("shisetsu_meisho", ""),
        "shisetsu_shurui":   request.form.get("shisetsu_shurui", ""),
        "shisetsu_shozaichi": request.form.get("shisetsu_shozaichi", ""),
        "nyusho_teiin":      request.form.get("nyusho_teiin", ""),
    }
    doc = DocxTemplate(SHOGAI_JIGYO)
    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    safe = context["meisho"].replace(" ", "_").replace("　", "_") or "届出者"
    return send_file(buffer, as_attachment=True,
                     download_name=f"障害福祉サービス事業等開始届_{safe}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/noshin", methods=["GET"])
def noshin():
    return render_template("noshin.html")


@app.route("/generate_noshin_moushide", methods=["POST"])
def generate_noshin_moushide():
    context = {
        "moushide_date":    request.form.get("moushide_date", ""),
        "shinshutsu_zip":   request.form.get("shinshutsu_zip", ""),
        "shinshutsu_jusho": request.form.get("shinshutsu_jusho", ""),
        "shinshutsu_tel":   request.form.get("shinshutsu_tel", ""),
        "shinshutsu_shimei": request.form.get("shinshutsu_shimei", ""),
        "daiko_zip":        request.form.get("daiko_zip", ""),
        "daiko_jusho":      request.form.get("daiko_jusho", ""),
        "daiko_tel":        request.form.get("daiko_tel", ""),
        "daiko_shimei":     request.form.get("daiko_shimei", ""),
        "oaza":             request.form.get("oaza", ""),
        "aza":              request.form.get("aza", ""),
        "chiban":           request.form.get("chiban", ""),
        "chimoku":          request.form.get("chimoku", ""),
        "menseki":          request.form.get("menseki", ""),
        "shoyu":            request.form.get("shoyu", ""),
        "henko_riyu":       request.form.get("henko_riyu", ""),
        # セクション3〜7
        "sentei_riyu":      request.form.get("sentei_riyu", ""),
        "nochi_jokyo":      request.form.get("nochi_jokyo", ""),
        "koko_toshi":       request.form.get("koko_toshi", ""),
        "osui_shori":       request.form.get("osui_shori", ""),
        "nicchou":          request.form.get("nicchou", ""),
        "sonota_eikyou":    request.form.get("sonota_eikyou", ""),
        "jigyo_shutai":     request.form.get("jigyo_shutai", ""),
        "kuiki_nai_menseki": request.form.get("kuiki_nai_menseki", ""),
        "kuiki_gai_menseki": request.form.get("kuiki_gai_menseki", ""),
        "shisetsu_menseki": request.form.get("shisetsu_menseki", ""),
        "jigyo_hi":         request.form.get("jigyo_hi", ""),
        "ta_genzai":        request.form.get("ta_genzai", ""),
        "ta_ato":           request.form.get("ta_ato", ""),
        "hata_genzai":      request.form.get("hata_genzai", ""),
        "hata_ato":         request.form.get("hata_ato", ""),
        "senkyo_kenbyo":    request.form.get("senkyo_kenbyo", ""),
    }
    doc = DocxTemplate(NOSHIN_MOUSHIDE)
    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    safe = context["shinshutsu_shimei"].replace(" ", "_").replace("　", "_") or "申請者"
    return send_file(buffer, as_attachment=True,
                     download_name=f"農業振興地域整備計画変更申出書_{safe}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/suido", methods=["GET"])
def suido():
    return render_template("suido.html")


@app.route("/generate_suido_syoyusha", methods=["POST"])
def generate_suido_syoyusha():
    context = {
        "todoke_date":        request.form.get("todoke_date", ""),
        "shutsu_jusho":       request.form.get("shutsu_jusho", ""),
        "shutsu_shimei":      request.form.get("shutsu_shimei", ""),
        "settichi":           request.form.get("settichi", ""),
        "settichi_tatemono":  request.form.get("settichi_tatemono", ""),
        "jyoto_date":         request.form.get("jyoto_date", ""),
        "mae_jusho":          request.form.get("mae_jusho", ""),
        "mae_shimei":         request.form.get("mae_shimei", ""),
        "ato_jusho":          request.form.get("ato_jusho", ""),
        "ato_shimei":         request.form.get("ato_shimei", ""),
        "jyoto_riyu":         request.form.get("jyoto_riyu", ""),
    }
    doc = DocxTemplate(SUIDO_SYOYUSHA)
    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    safe = context["ato_shimei"].replace(" ", "_").replace("　", "_") or "申請者"
    return send_file(buffer, as_attachment=True,
                     download_name=f"排水設備所有者変更届_{safe}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/kouminkan", methods=["GET"])
def kouminkan():
    return render_template("kouminkan.html")


def _kouminkan_context():
    ctx = {
        "shinsei_date":   request.form.get("shinsei_date", ""),
        "kouminkan_name": request.form.get("kouminkan_name", ""),
        "jusho":          request.form.get("jusho", ""),
        "dantai_name":    request.form.get("dantai_name", ""),
        "daihyo_name":    request.form.get("daihyo_name", ""),
        "denwa":          request.form.get("denwa", ""),
        "shiyou_mokuteki": request.form.get("shiyou_mokuteki", ""),
        "shiyou_ninzu":   request.form.get("shiyou_ninzu", ""),
    }
    # 使用日程（5行分）
    for i in range(1, 6):
        ctx[f"date_{i}"]  = request.form.get(f"date_{i}", "")
        ctx[f"room_{i}"]  = request.form.get(f"room_{i}", "")
        ctx[f"time_{i}"]  = request.form.get(f"time_{i}", "")
        ctx[f"biko_{i}"]  = request.form.get(f"biko_{i}", "")
    return ctx


@app.route("/generate_kouminkan_shiyo", methods=["POST"])
def generate_kouminkan_shiyo():
    context = _kouminkan_context()
    doc = DocxTemplate(KOUMINKAN_SHIYO)
    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    safe = context["dantai_name"].replace(" ", "_").replace("　", "_") or "申請者"
    return send_file(buffer, as_attachment=True,
                     download_name=f"公民館使用許可申請書_{safe}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/generate_kouminkan_genmen", methods=["POST"])
def generate_kouminkan_genmen():
    context = _kouminkan_context()
    doc = DocxTemplate(KOUMINKAN_GENMEN)
    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    safe = context["dantai_name"].replace(" ", "_").replace("　", "_") or "申請者"
    return send_file(buffer, as_attachment=True,
                     download_name=f"公民館使用料減免申請書_{safe}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/generate_kouminkan_chushi", methods=["POST"])
def generate_kouminkan_chushi():
    ctx = {
        "shinsei_date":   request.form.get("shinsei_date", ""),
        "kouminkan_name": request.form.get("kouminkan_name", ""),
        "jusho":          request.form.get("jusho", ""),
        "dantai_name":    request.form.get("dantai_name", ""),
        "daihyo_name":    request.form.get("daihyo_name", ""),
        "denwa":          request.form.get("denwa", ""),
    }
    doc = DocxTemplate(KOUMINKAN_CHUSHI)
    doc.render(ctx)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    safe = ctx["dantai_name"].replace(" ", "_").replace("　", "_") or "申請者"
    return send_file(buffer, as_attachment=True,
                     download_name=f"公民館使用中止届_{safe}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
