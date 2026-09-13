import os
import requests
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)

# Saytingiz qaysi domendan so'rov yuborishiga ruxsat berish.
# Ehtiyot chorasi sifatida hozircha "*" (hammaga ochiq) qilib qo'ydik.
# Xohlasangiz keyinroq faqat o'z GitHub Pages domeningizga cheklashingiz mumkin:
# CORS(app, resources={r"/api/*": {"origins": "https://gabimarutopp-one.github.io"}})
CORS(app)

HEMIS_LOGIN_PATH = "/auth/login"
HEMIS_GPA_PATH = "/education/gpa-list"
HEMIS_SUBJECTS_PATH = "/education/subject-list"
REQUEST_TIMEOUT = 15


def normalize_domain(text: str) -> str:
    """'student.tuit.uz' yoki 'tuit' -> 'https://student.tuit.uz/rest/v1'"""
    text = (text or "").strip().rstrip("/").lower()
    if not text.startswith("http://") and not text.startswith("https://") and "." not in text:
        text = f"student.{text}.uz"
    if not text.startswith("http://") and not text.startswith("https://"):
        text = "https://" + text
    if not text.endswith("/rest/v1"):
        text = text + "/rest/v1"
    return text


@app.route("/")
def health():
    return jsonify({"status": "ok"})


@app.route("/api/login", methods=["POST"])
def login():
    body = request.get_json(silent=True) or {}
    domain_raw = body.get("domain", "")
    login_val = body.get("login", "")
    password = body.get("password", "")

    if not domain_raw or not login_val or not password:
        return jsonify({"ok": False, "error": "missing_fields"}), 400

    base_url = normalize_domain(domain_raw)

    try:
        resp = requests.post(
            base_url + HEMIS_LOGIN_PATH,
            json={"login": login_val, "password": password},
            timeout=REQUEST_TIMEOUT,
        )
    except requests.exceptions.ConnectionError:
        return jsonify({"ok": False, "error": "dns"}), 400
    except Exception:
        return jsonify({"ok": False, "error": "error"}), 500

    if resp.status_code != 200:
        return jsonify({"ok": False, "error": "auth"}), 401

    try:
        data = resp.json()
    except Exception:
        return jsonify({"ok": False, "error": "auth"}), 401

    inner = data.get("data") if isinstance(data, dict) else None
    token = (
        (inner or {}).get("token") if isinstance(inner, dict) else None
    ) or data.get("token") or data.get("access_token")

    if not token:
        return jsonify({"ok": False, "error": "auth"}), 401

    return jsonify({"ok": True, "token": token, "domain": base_url})


@app.route("/api/data", methods=["GET"])
def data():
    base_url = request.args.get("domain", "")
    token = request.args.get("token", "")

    if not base_url or not token:
        return jsonify({"ok": False, "error": "missing_fields"}), 400

    headers = {"Authorization": f"Bearer {token}"}

    result = {"ok": True, "gpa": [], "subjects": []}

    try:
        gpa_resp = requests.get(
            base_url + HEMIS_GPA_PATH, headers=headers, timeout=REQUEST_TIMEOUT
        )
        if gpa_resp.status_code == 401:
            return jsonify({"ok": False, "error": "token_expired"}), 401
        if gpa_resp.status_code == 200:
            gpa_json = gpa_resp.json()
            gpa_list = (gpa_json or {}).get("data", []) if isinstance(gpa_json, dict) else []
            for item in gpa_list:
                level = (
                    item.get("level", {}).get("name")
                    if isinstance(item.get("level"), dict)
                    else item.get("level", "")
                )
                gpa_val = item.get("gpa", item.get("avg_gpa", "—"))
                result["gpa"].append({"level": level, "gpa": gpa_val})
    except Exception:
        pass

    try:
        sub_resp = requests.get(
            base_url + HEMIS_SUBJECTS_PATH, headers=headers, timeout=REQUEST_TIMEOUT
        )
        if sub_resp.status_code == 200:
            sub_json = sub_resp.json()
            subjects = (sub_json or {}).get("data", []) if isinstance(sub_json, dict) else []
            for s in subjects[:30]:
                name = (
                    s.get("subject", {}).get("name")
                    if isinstance(s.get("subject"), dict)
                    else s.get("subject", "")
                )
                grade = s.get("grade", s.get("total_ball", "—"))
                result["subjects"].append({"name": name, "grade": grade})
    except Exception:
        pass

    return jsonify(result)


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
