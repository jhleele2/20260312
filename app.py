# -*- coding: utf-8 -*-
"""
재고 파악 → 부족 시 담당 기업에 발주서 이메일 발송 웹 시스템.
데이터: 엑셀 파일(domino_inventory_training.xlsx 구조) 업로드 또는 기본 파일 사용.
"""
import os
from pathlib import Path
from time import time

_base_dir = Path(__file__).resolve().parent
_env_file = _base_dir / ".env"


def _load_env():
    """ .env 파일을 읽어 os.environ에 반영 (dotenv 미설치/실패 시 대비) """
    if not _env_file.exists():
        return
    try:
        with open(_env_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, _, value = line.partition("=")
                    key = key.strip()
                    value = value.strip().strip('"').strip("'")
                    if key:
                        os.environ[key] = value
    except Exception:
        pass


try:
    from dotenv import load_dotenv
    load_dotenv(_env_file)
except ImportError:
    pass
_load_env()

from flask import Flask, request, render_template, jsonify, redirect, url_for, session
from werkzeug.utils import secure_filename

from inventory_loader import load_all, get_orders_by_supplier, update_inventory_item
from email_sender import fill_template, send_order_email, DEFAULT_SENDER_EMAIL, DEFAULT_STORE_NAME, DEFAULT_INTERNAL_OWNER

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "inventory-dev-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10MB
# Vercel: 쓰기 가능한 경로 사용
if os.environ.get("VERCEL") == "1":
    app.config["UPLOAD_FOLDER"] = Path("/tmp/uploads")
else:
    app.config["UPLOAD_FOLDER"] = _base_dir / "uploads"
app.config["DEFAULT_EXCEL"] = _base_dir / "domino_inventory_training.xlsx"
app.config["UPLOAD_FOLDER"].mkdir(parents=True, exist_ok=True)

TEAM_PASSWORD = os.environ.get("TEAM_PASSWORD", "1234")

ALLOWED_EXTENSIONS = {"xlsx", "xls"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[-1].lower() in ALLOWED_EXTENSIONS


def auth_required():
    if session.get("auth") is True:
        return None
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if (request.form.get("password") or "").strip() == TEAM_PASSWORD:
            session["auth"] = True
            return redirect(url_for("index"))
        return render_template("login.html", error="비밀번호가 올바르지 않습니다.")
    if session.get("auth") is True:
        return redirect(url_for("index"))
    return render_template("login.html", error=None)


@app.route("/logout")
def logout():
    session.pop("auth", None)
    return redirect(url_for("login"))


@app.before_request
def require_auth():
    if request.endpoint in ("login", "static") or request.path.startswith("/static"):
        return None
    return auth_required()


@app.route("/")
def index():
    # 업로드한 파일은 세션에 저장된 파일명으로 경로 구성 (URL 경로 깨짐 방지)
    excel_path = str(app.config["DEFAULT_EXCEL"])
    if session.get("uploaded_file"):
        candidate = Path(app.config["UPLOAD_FOLDER"]) / session["uploaded_file"]
        if candidate.exists():
            excel_path = str(candidate)
    # 쿼리 파라미터로 파일명만 넘긴 경우 (file=업로드파일명)
    file_param = request.args.get("file")
    if file_param and ".." not in file_param:
        candidate = Path(app.config["UPLOAD_FOLDER"]) / file_param
        if candidate.exists():
            excel_path = str(candidate)
            session["uploaded_file"] = file_param
    if not Path(excel_path).exists():
        excel_path = str(app.config["DEFAULT_EXCEL"])
        session.pop("uploaded_file", None)
    data = load_all(excel_path)
    if data.get("error"):
        return render_template(
            "index.html", error=data["error"], orders=[], inventory=[], summary=None, read_only_deploy=os.environ.get("VERCEL") == "1"
        )
    inventory = data["inventory"]
    orders = get_orders_by_supplier(inventory)
    need_count = sum(1 for i in inventory if i.get("order_quantity", 0) > 0)
    total_items = len(inventory)
    total_order_qty = sum(i.get("order_quantity", 0) for i in inventory)
    summary = {
        "total_items": total_items,
        "need_order_count": need_count,
        "total_order_quantity": total_order_qty,
        "supplier_count": len(orders),
    }
    return render_template(
        "index.html",
        error=None,
        inventory=inventory,
        orders=orders,
        summary=summary,
        excel_path=excel_path,
        email_template=data.get("email_template", {}),
        read_only_deploy=os.environ.get("VERCEL") == "1",
    )


@app.route("/upload", methods=["POST"])
def upload():
    if "excel" not in request.files:
        return redirect(url_for("index"))
    f = request.files["excel"]
    if f.filename == "" or not allowed_file(f.filename):
        return redirect(url_for("index"))
    upload_dir = Path(app.config["UPLOAD_FOLDER"])
    upload_dir.mkdir(parents=True, exist_ok=True)
    ext = (f.filename or "").rsplit(".", 1)[-1].lower() if "." in (f.filename or "") else "xlsx"
    if ext not in ALLOWED_EXTENSIONS:
        ext = "xlsx"
    safe_name = secure_filename(f.filename) or f"upload_{int(time())}.{ext}"
    if not safe_name.strip():
        safe_name = f"upload_{int(time())}.{ext}"
    path = upload_dir / safe_name
    try:
        f.save(str(path))
    except (PermissionError, OSError) as e:
        return render_template(
            "index.html",
            error=f"파일 저장 실패(권한 오류): {e}. 'uploads' 폴더 쓰기 권한을 확인하세요.",
            orders=[],
            inventory=[],
            summary=None,
        )
    session["uploaded_file"] = safe_name
    return redirect(url_for("index", file=safe_name))


@app.route("/api/inventory/update", methods=["POST"])
def api_inventory_update():
    """재고 품목 한 건 수정. 엑셀 파일에 반영."""
    data = request.get_json() or {}
    excel_path = data.get("excel_path") or str(app.config["DEFAULT_EXCEL"])
    item_code = data.get("item_code")
    if not item_code:
        return jsonify({"ok": False, "message": "품목코드가 필요합니다."})
    if not Path(excel_path).exists():
        return jsonify({"ok": False, "message": "엑셀 파일을 찾을 수 없습니다."})
    try:
        current_stock = data.get("current_stock")
        safety_stock = data.get("safety_stock")
        moq = data.get("moq")
        if current_stock is not None:
            current_stock = int(current_stock)
        if safety_stock is not None:
            safety_stock = int(safety_stock)
        if moq is not None:
            moq = int(moq)
    except (TypeError, ValueError):
        return jsonify({"ok": False, "message": "현재고/안전재고/MOQ는 숫자로 입력하세요."})
    ok, msg = update_inventory_item(
        excel_path,
        str(item_code).strip(),
        current_stock=current_stock,
        safety_stock=safety_stock,
        moq=moq,
        name=data.get("name"),
        spec=data.get("spec"),
        unit=data.get("unit"),
        supplier=data.get("supplier"),
    )
    return jsonify({"ok": ok, "message": msg})


@app.route("/api/send-orders", methods=["POST"])
def api_send_orders():
    """선택한 공급업체에 대해 발주서 이메일 발송."""
    excel_path = request.json.get("excel_path") or str(app.config["DEFAULT_EXCEL"])
    supplier_indices = request.json.get("supplier_indices")  # 보낼 공급업체 인덱스 목록 (전부면 생략 가능)
    store_name = request.json.get("store_name") or DEFAULT_STORE_NAME
    internal_owner = request.json.get("internal_owner") or DEFAULT_INTERNAL_OWNER

    sender_email = os.environ.get("SMTP_USER", DEFAULT_SENDER_EMAIL)
    sender_password = os.environ.get("SMTP_PASSWORD", "")
    smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))

    data = load_all(excel_path)
    if data.get("error"):
        return jsonify({"ok": False, "message": data["error"]})
    inventory = data["inventory"]
    orders = get_orders_by_supplier(inventory)
    # Suppliers 시트에서 공급업체명 → 이메일 보정 (Inventory에 이메일이 비어 있을 때 사용)
    supplier_email_map = {}
    for s in data.get("suppliers", []):
        name = (s.get("name") or "").strip()
        email = (s.get("email") or "").strip()
        if name and email and "@" in email:
            supplier_email_map[name] = email
    for order in orders:
        if not order.get("email") or "@" not in str(order.get("email", "")):
            fallback = supplier_email_map.get((order.get("supplier_name") or "").strip())
            if fallback:
                order["email"] = fallback
    tpl = data.get("email_template", {})
    subject_tpl = tpl.get("subject") or "[발주요청] {{STORE_NAME}} / {{SUPPLIER_NAME}} / {{ORDER_DATE}}"
    body_tpl = tpl.get("body") or "안녕하세요 {{SUPPLIER_NAME}} 담당자님.\n\n{{STORE_NAME}}입니다.\n아래 품목 발주 요청 드립니다.\n\n{{ITEM_LIST}}\n\n확인 부탁드립니다.\n{{INTERNAL_OWNER}}"

    results = []
    for idx, order in enumerate(orders):
        if supplier_indices is not None and idx not in supplier_indices:
            continue
        subject, body = fill_template(
            subject_tpl, body_tpl,
            order["supplier_name"], order["items"],
            store_name=store_name, internal_owner=internal_owner,
        )
        ok, msg = send_order_email(
            order["email"], subject, body,
            sender_email=sender_email, sender_password=sender_password,
            smtp_host=smtp_host, smtp_port=smtp_port,
            bcc=os.environ.get("SMTP_BCC") or sender_email,
        )
        results.append({
            "supplier": order["supplier_name"],
            "email": order["email"],
            "ok": ok,
            "message": msg,
        })
    return jsonify({"ok": True, "results": results})


if __name__ == "__main__":
    # 코드·.env 수정 시 서버 자동 재시작 (use_reloader=True)
    _extra = [str(_env_file)] if _env_file.exists() else []
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True,
        use_reloader=True,
        extra_files=_extra,
    )
