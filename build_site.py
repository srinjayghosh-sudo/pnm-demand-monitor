"""Builds the GitHub Pages site after the daily update.
Reads the latest dashboard data, run status and Excel from TiDB, encrypts them with the DASHBOARD_KEY
secret (AES-256-GCM, key derived with PBKDF2-SHA256), and writes site/ for GitHub Pages.
Without the key, the published files are unreadable."""
import base64, gzip, json, os, shutil, sys
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
import pnm_daily as P

ITER = 250_000

def seal(raw: bytes, key: str) -> dict:
    salt, iv = os.urandom(16), os.urandom(12)
    k = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(key.encode())
    ct = AESGCM(k).encrypt(iv, raw, None)
    b = lambda x: base64.b64encode(x).decode()
    return {"v": 1, "iter": ITER, "salt": b(salt), "iv": b(iv), "ct": b(ct)}

def parts(con, name):
    df = P.q(con, "SELECT data, updated_at FROM dashboard_files WHERE name = %s ORDER BY part", (name,))
    return (b"".join(df.data), str(df.updated_at.iloc[0])) if len(df) else (None, None)

def main():
    key = os.environ.get("DASHBOARD_KEY", "").strip()
    if len(key) < 12:
        sys.exit("DASHBOARD_KEY secret is missing or shorter than 12 characters.")
    con = P.db_connect(P.CONFIG)
    gz, _ = parts(con, "payload_gz")
    data = json.loads(gzip.decompress(gz)) if gz else {"empty": True, "meta": {}}
    ok = P.q(con, "SELECT run_at, detail FROM run_log WHERE step = 'refresh-ok' ORDER BY id DESC LIMIT 1")
    er = P.q(con, "SELECT run_at, detail FROM run_log WHERE step LIKE '%%error%%' ORDER BY id DESC LIMIT 1")
    last = ok.iloc[0].to_dict() if len(ok) else None
    lerr = er.iloc[0].to_dict() if len(er) and (last is None or er.run_at.iloc[0] > last["run_at"]) else None
    repo = os.environ.get("GITHUB_SERVER_URL", "https://github.com") + "/" + os.environ.get("GITHUB_REPOSITORY", "")
    data["status"] = {"last_refresh": last, "error": f"{lerr['run_at']}: {lerr['detail']}" if lerr else "",
                      "actions_url": repo + "/actions"}
    os.makedirs("site", exist_ok=True)
    shutil.copy("dashboard_pages.html", "site/index.html")
    open("site/.nojekyll", "w").close()
    raw = gzip.compress(json.dumps(data, default=str).encode())
    json.dump(seal(raw, key), open("site/data.enc.json", "w"))
    xl, upd = parts(con, "xlsx")
    if xl:
        json.dump({**seal(xl, key), "name": f"pnm_demand_{(upd or '')[:10]}.xlsx"}, open("site/excel.enc.json", "w"))
    print(f"site built: data {len(raw)//1024} KB, excel {len(xl or b'')//1024} KB (both encrypted)")

if __name__ == "__main__":
    main()
