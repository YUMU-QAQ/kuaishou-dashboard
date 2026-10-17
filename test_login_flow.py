"""登入流程自检：未登录跳转/接口 401、错误密码拒绝、正确密码发 cookie、退出后失效。"""
import os, sys
from pathlib import Path

PROJ = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJ))
os.environ["AUTH_USER"] = "admin"
os.environ["AUTH_PASSWORD"] = "testpw123"

import importlib
app_mod = importlib.import_module("app")
from fastapi.testclient import TestClient

c = TestClient(app_mod.app, follow_redirects=False)

# 未登录
r = c.get("/")
assert r.status_code == 303 and r.headers["location"] == "/login", r.status_code
r = c.get("/api/data")
assert r.status_code == 401, r.status_code
r = c.get("/login")
assert r.status_code == 200 and "青野快手数据看板" in r.text, r.status_code

# 密码错
r = c.post("/api/login", json={"username": "admin", "password": "bad"})
assert r.status_code == 401, r.status_code
# 用户名错
r = c.post("/api/login", json={"username": "root", "password": "testpw123"})
assert r.status_code == 401, r.status_code

# 密码对
r = c.post("/api/login", json={"username": "admin", "password": "testpw123"})
assert r.status_code == 200, r.status_code
assert c.cookies.get("qy_session"), "未下发 cookie"
r = c.get("/")
assert r.status_code == 200 and "数据看板" in r.text, r.status_code
r = c.get("/api/accounts")
assert r.status_code == 200, r.status_code

# 退出
r = c.post("/api/logout")
assert r.status_code == 200, r.status_code
r = c.get("/")
assert r.status_code == 303, r.status_code

# 过期 token 失效
app_mod.SESSIONS["stale"] = 1.0
r = c.get("/", cookies={"qy_session": "stale"})
assert r.status_code == 303, r.status_code

print("登录流程 OK")
