"""
page_crypto.py — password-protect the published dashboard (GitHub Pages).

The HTML report is encrypted with AES-256-GCM using a key derived from a password
(PBKDF2-HMAC-SHA256, 600,000 iterations, random salt). The published page is a small unlock
screen that holds only the encrypted data; the browser decrypts it locally (WebCrypto) after the
password is typed. summary.json is encrypted the same way (it contains the account numbers).

The password comes from the PAGE_PASSWORD environment variable (a GitHub secret) — never from
the repository. CRYPTO_PASSWORD, if set, takes precedence; the workflow uses it (with the Telegram
bot token secret) to encrypt the summary handed from the strategy job to the Telegram job.

Usage:
  python page_crypto.py page  <report.html>   <public/index.html>
  python page_crypto.py enc   <summary.json>  <public/summary.json.enc>
  python page_crypto.py dec   <summary.json.enc> <summary.json>
"""
import base64
import json
import os
import sys

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ITERATIONS = 600_000


def _key(password, salt, iterations=ITERATIONS):
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt,
                      iterations=iterations).derive(password.encode("utf-8"))


def encrypt_bytes(data, password):
    salt, iv = os.urandom(16), os.urandom(12)
    ct = AESGCM(_key(password, salt)).encrypt(iv, data, None)
    b64 = lambda b: base64.b64encode(b).decode("ascii")
    return {"v": 1, "kdf": "PBKDF2-SHA256", "iter": ITERATIONS, "salt": b64(salt), "iv": b64(iv), "data": b64(ct)}


def decrypt_bytes(box, password):
    d = lambda k: base64.b64decode(box[k])
    return AESGCM(_key(password, d("salt"), box["iter"])).decrypt(d("iv"), d("data"), None)


LOADER = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>ETF Momentum Live Dashboard</title>
<style>
 :root { --bg:#0D1B2A; --panel:#1B2A3A; --txt:#F0F0F0; --muted:#9AA8B6; --acc:#F5A623; --bad:#E74C3C; }
 body { background:var(--bg); color:var(--txt); font-family:Segoe UI,Roboto,Arial,sans-serif; margin:0;
        min-height:100vh; display:flex; align-items:center; justify-content:center; padding:16px; box-sizing:border-box; }
 .box { background:var(--panel); border-radius:12px; padding:24px; width:100%; max-width:360px; }
 h1 { font-size:18px; margin:0 0 6px; } p { color:var(--muted); font-size:13px; margin:0 0 16px; }
 input[type=password] { width:100%; box-sizing:border-box; padding:12px; border-radius:8px; border:1px solid #34495E;
        background:#0F2233; color:var(--txt); font-size:16px; }
 label { display:flex; gap:8px; align-items:center; color:var(--muted); font-size:13px; margin:12px 0; }
 button { width:100%; padding:12px; border:0; border-radius:8px; background:var(--acc); color:#2B1D02;
        font-weight:700; font-size:15px; cursor:pointer; }
 button:disabled { opacity:.6; } .err { color:var(--bad); font-size:13px; min-height:18px; margin-top:10px; }
</style></head><body>
<form class="box" id="f" autocomplete="on">
 <h1>🔒 ETF Momentum — Live Dashboard</h1>
 <p>This page is password-protected. It is decrypted on your device; the password is never sent anywhere.</p>
 <input type="text" name="username" value="etf-dashboard" autocomplete="username" hidden>
 <input type="password" id="pw" placeholder="Password" autocomplete="current-password" required autofocus>
 <label><input type="checkbox" id="rem"> Remember on this device</label>
 <button id="go" type="submit">Unlock</button>
 <div class="err" id="err"></div>
</form>
<script id="payload" type="application/json">__PAYLOAD__</script>
<script>
(function () {
  var box = JSON.parse(document.getElementById('payload').textContent);
  var KEY = 'etfDashPw';
  function b64(s) { var b = atob(s), a = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; }
  async function unlock(pw) {
    var base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
    var key = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt: b64(box.salt), iterations: box.iter, hash: 'SHA-256' },
                                            base, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
    var plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b64(box.iv) }, key, b64(box.data));
    return new TextDecoder().decode(plain);
  }
  async function attempt(pw, remember) {
    var btn = document.getElementById('go'), err = document.getElementById('err');
    btn.disabled = true; btn.textContent = 'Unlocking…'; err.textContent = '';
    try {
      var html = await unlock(pw);
      try { if (remember) localStorage.setItem(KEY, pw); } catch (e) {}
      document.open(); document.write(html); document.close();
    } catch (e) {
      try { localStorage.removeItem(KEY); } catch (e2) {}
      btn.disabled = false; btn.textContent = 'Unlock';
      err.textContent = window.isSecureContext === false ? 'Open this page over https.' : 'Wrong password — try again.';
    }
  }
  document.getElementById('f').addEventListener('submit', function (ev) {
    ev.preventDefault();
    attempt(document.getElementById('pw').value, document.getElementById('rem').checked);
  });
  var saved = null; try { saved = localStorage.getItem(KEY); } catch (e) {}
  if (saved) { document.getElementById('rem').checked = true; attempt(saved, true); }
})();
</script>
</body></html>
"""


def _password():
    """CRYPTO_PASSWORD (used for the private hand-off between workflow jobs) or PAGE_PASSWORD."""
    pw = os.environ.get("CRYPTO_PASSWORD") or os.environ.get("PAGE_PASSWORD", "")
    if not pw:
        sys.exit("No password: set PAGE_PASSWORD (or CRYPTO_PASSWORD)")
    return pw


def main(argv):
    if len(argv) != 4 or argv[1] not in ("page", "enc", "dec"):
        print(__doc__)
        return 2
    mode, src, dst = argv[1:]
    pw = _password()
    if mode == "page":
        box = encrypt_bytes(open(src, "rb").read(), pw)
        payload = json.dumps(box).replace("</", "<\\/")
        with open(dst, "w", encoding="utf-8") as f:
            f.write(LOADER.replace("__PAYLOAD__", payload))
    elif mode == "enc":
        with open(dst, "w", encoding="utf-8") as f:
            json.dump(encrypt_bytes(open(src, "rb").read(), pw), f)
    else:
        with open(dst, "wb") as f:
            f.write(decrypt_bytes(json.load(open(src, encoding="utf-8")), pw))
    print(f"{mode}: {src} -> {dst}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
