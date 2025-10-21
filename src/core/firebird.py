from __future__ import annotations
import os, sys, pathlib, logging
import fdb

log = logging.getLogger(__name__)

def _carregar_fbclient() -> str:
    candidatos, add_dirs = [], set()
    env = [os.getenv("FIREBIRD_CLIENT_PATH", ""), os.getenv("FBCLIENT_PATH", "")]
    candidatos += [p for p in env if p]
    add_dirs.update([os.path.dirname(p) for p in candidatos if p])

    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        exe_dir = os.path.dirname(sys.executable)
        candidatos += [
            os.path.join(base, "fbclient.dll"),
            os.path.join(exe_dir, "fbclient.dll"),
            os.path.join(base, "bin", "fbclient.dll"),
        ]
        add_dirs.update([base, exe_dir, os.path.join(base, "plugins"), os.path.join(base, "bin")])
    else:
        comuns = [
            r"C:\Program Files\Firebird\Firebird_5_0\bin\fbclient.dll",
            r"C:\Program Files\Firebird\Firebird_4_0\bin\fbclient.dll",
            r"C:\Program Files (x86)\Firebird\Firebird_5_0\bin\fbclient.dll",
            r"C:\Program Files (x86)\Firebird\Firebird_4_0\bin\fbclient.dll",
        ]
        this_dir = pathlib.Path(__file__).resolve().parent
        cwd = pathlib.Path.cwd()
        locais = [
            str(this_dir / "fbclient.dll"),
            str(cwd / "fbclient.dll"),
            str(this_dir / "bin" / "fbclient.dll"),
            str(cwd / "bin" / "fbclient.dll"),
        ]
        candidatos += comuns + locais
        add_dirs.update([os.path.dirname(p) for p in comuns] + [str(this_dir), str(cwd), str(this_dir / "bin"), str(cwd / "bin")])

    candidatos.append("fbclient.dll")
    for d in list(add_dirs):
        try:
            if d and os.path.isdir(d):
                os.add_dll_directory(d)
        except Exception:
            pass

    for p in candidatos:
        try:
            if os.path.isabs(p):
                if os.path.exists(p):
                    fdb.load_api(p); return p
            else:
                fdb.load_api(p); return p
        except Exception:
            continue

    raise RuntimeError(
        "fbclient.dll não encontrada. Instale o Firebird Client (mesma arquitetura do Python) "
        "ou defina FIREBIRD_CLIENT_PATH com o caminho completo."
    )

def conectar():
    _carregar_fbclient()
    return fdb.connect(
        host=os.getenv("FB_HOST", "localhost"),
        database=os.getenv("FB_DB"),
        user=os.getenv("FB_USER"),
        password=os.getenv("FB_PASSWORD"),
        port=int(os.getenv("FB_PORT", "3050")),
        charset=os.getenv("FB_CHARSET", "ISO8859_1"),
    )
