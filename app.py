import io
import os
from docxtpl import DocxTemplate
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

TEIJUU_TEMPLATE  = os.path.join(os.path.dirname(__file__), "teijuu_template.docx")
SHUSSAN_TEMPLATE = os.path.join(os.path.dirname(__file__), "shussan_template.docx")


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


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
