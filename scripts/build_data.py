"""Lee los Excel de /datos y genera docs/data.json para la web."""
import glob, json, re, unicodedata, datetime as dt
from openpyxl import load_workbook


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", s.lower())

def num(v):
    if v in (None, ""): return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).replace("€", "").replace(" ", "").strip()
    return float(s.replace(".", "").replace(",", ".")) if s else 0.0

def fecha(v):
    if isinstance(v, (dt.datetime, dt.date)): return v.strftime("%Y-%m-%d")
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", str(v or ""))
    return f"{m[3]}-{int(m[2]):02d}-{int(m[1]):02d}" if m else ""

def filas(ws):
    it = ws.iter_rows(values_only=True)
    cab = [norm(c) for c in next(it, [])]
    for r in it:
        if any(c not in (None, "") for c in r):
            yield dict(zip(cab, r))

mov, cod, pres = [], {}, []
for f in sorted(glob.glob("datos/*.xlsx")):
    if f.endswith("plantilla.xlsx"): continue
    wb = load_workbook(f, data_only=True)
    for ws in wb.worksheets:
        n = norm(ws.title)
        if n == "movimientos":
            mov += list(filas(ws))
        elif n == "codigos":
            for r in filas(ws): cod[str(r.get("codigo"))] = r.get("nombrecorto")
        elif n == "presupuesto":
            pres += list(filas(ws))

s = lambda v: str(v).strip() if v not in (None, "") else ""
movimientos = []
for r in mov:
    ronda = s(r.get("ronda"))
    if not ronda: continue
    o = s(r.get("origendestino"))
    movimientos.append({"ronda": ronda, "n": s(r.get("n")), "fecha": fecha(r.get("fecha")),
         "concepto": s(r.get("concepto")), "partida": s(r.get("partida")),
         "subpartida": s(r.get("subpartida")), "importe": abs(num(r.get("importe"))),
         "tipo": s(r.get("tipo")), "saldo": num(r.get("saldo")),
         "seccion": s(r.get("seccion")) or "Grupo", "doc": s(r.get("doc")),
         "origen": s(r.get("origen")), "ndoc": s(r.get("ndoc"))})  # las observaciones no se publican
movimientos.sort(key=lambda m: (m["fecha"], m["n"]))

presupuesto = [{
    "ronda": s(r.get("ronda")), "agrupacion": s(r.get("agrupacion")), "tipo": s(r.get("tipo")) or "Gasto",
    "partida": s(r.get("partida")), "subpartida": s(r.get("subpartida")),
    "calendario": s(r.get("calendario")) or "Anual", "anual": num(r.get("anual"))} for r in pres]

json.dump({"actualizado": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
           "movimientos": movimientos, "presupuesto": presupuesto},
          open("docs/data.json", "w", encoding="utf-8"), ensure_ascii=False)
print(len(movimientos), "movimientos,", len(presupuesto), "líneas de presupuesto")
