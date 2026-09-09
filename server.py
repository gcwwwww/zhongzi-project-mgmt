# -*- coding: utf-8 -*-
"""
中咨养护检测项目与费用管理 - 共享数据服务器
零依赖，使用 Python 标准库。
- 静态文件服务（site/ 目录）
- GET /api/data  读取共享数据
- POST /api/data 保存共享数据
绑定 0.0.0.0:8765，局域网内多人可同时访问。
"""
import os, json, threading
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse

SITE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "site")
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")
LOCK = threading.Lock()

class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=SITE_DIR, **kwargs)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/data":
            self._serve_data()
        elif parsed.path == "/api/reset":
            self._reset_data()
        else:
            super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/data":
            self._save_data()
        elif parsed.path == "/api/reset":
            self._reset_data()
        else:
            self.send_error(404, "Not Found")

    def _reset_data(self):
        """删除 data.json，重新从 seed.js 生成"""
        try:
            with LOCK:
                if os.path.exists(DATA_FILE):
                    os.remove(DATA_FILE)
            self._init_from_seed()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(b'{"ok":true,"msg":"reset done"}')
        except Exception as e:
            self._json_err(500, str(e))

    def _init_from_seed(self):
        """从 seed.js 初始化 data.json（调用模块级静态方法）"""
        _init_from_seed_static(SITE_DIR, DATA_FILE)

    def _serve_data(self):
        try:
            with LOCK:
                if os.path.exists(DATA_FILE):
                    with open(DATA_FILE, "r", encoding="utf-8") as f:
                        raw = f.read()
                else:
                    raw = "{}"
            body = raw.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(body)
        except Exception as e:
            self._json_err(500, str(e))

    def _save_data(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length).decode("utf-8")
            # 验证是合法 JSON
            json.loads(raw)
            with LOCK:
                with open(DATA_FILE, "w", encoding="utf-8") as f:
                    f.write(raw)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')
        except json.JSONDecodeError as e:
            self._json_err(400, "Invalid JSON: " + str(e))
        except Exception as e:
            self._json_err(500, str(e))

    def _json_err(self, code, msg):
        body = json.dumps({"error": msg}, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def end_headers(self):
        # 防缓存：HTML/JS/CSS 始终获取最新
        self.send_header("Cache-Control", "no-cache, must-revalidate")
        super().end_headers()

    def log_message(self, fmt, *args):
        # 简化日志
        import datetime
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        print("[%s] %s - %s" % (ts, self.address_string(), fmt % args))

def main():
    # Render 等云平台通过 PORT 环境变量注入端口，默认 8765
    port = int(os.environ.get("PORT", "8765"))
    host = "0.0.0.0"
    # 初始化 data.json（若不存在则从 seed 生成）
    if not os.path.exists(DATA_FILE):
        Handler()._init_from_seed() if False else _init_from_seed_static(SITE_DIR, DATA_FILE)
    print("=" * 60)
    print("  中咨养护检测项目与费用管理系统")
    print("  服务器已启动，局域网内多人可访问：")
    print("  本机:   http://127.0.0.1:%d/" % port)
    print("  局域网: http://<本机IP>:%d/" % port)
    print("  数据文件: %s" % DATA_FILE)
    print("  重置数据: 访问 http://127.0.0.1:%d/api/reset" % port)
    print("  按 Ctrl+C 停止服务")
    print("=" * 60)
    server = HTTPServer((host, port), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务器已停止")
        server.server_close()


def _init_from_seed_static(site_dir, data_file):
    """模块级静态方法，从 seed.js 初始化 data.json"""
    seed_path = os.path.join(site_dir, "js", "seed.js")
    if not os.path.exists(seed_path):
        return
    with open(seed_path, "r", encoding="utf-8") as f:
        txt = f.read()
    start = txt.find("{")
    if start < 0:
        return
    end = txt.rfind("}")
    if end <= start:
        return
    seed = json.loads(txt[start:end+1])
    if "achievements" not in seed:
        seed["achievements"] = []
    for p in seed.get("projects", []):
        if "id" not in p:
            p["id"] = "p" + str(abs(hash(p["name"])) % 10000)
        if "monthlyReports" not in p:
            p["monthlyReports"] = {}
        if "targets" not in p:
            p["targets"] = {}
    personnel = []
    for pe in seed.get("people", []):
        personnel.append({
            "id": "pe" + str(abs(hash(pe["name"])) % 10000),
            "name": pe["name"],
            "responsibility": pe.get("responsibility", ""),
            "projectId": None,
            "attendance": {seed.get("currentMonth", ""): {"days": pe.get("days", {})}}
        })
    out = {
        "settings": {
            "companyName": seed.get("companyName", ""),
            "currentMonth": seed.get("currentMonth", ""),
            "dailyRate": seed.get("dailyRate", 300),
            "signatures": {}
        },
        "projects": seed.get("projects", []),
        "personnel": personnel,
        "achievements": seed.get("achievements", [])
    }
    with open(data_file, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()
