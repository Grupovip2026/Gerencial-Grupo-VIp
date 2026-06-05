#!/usr/bin/env python3
"""
gerar_dados.py
Lê Resultado_Gerencial_2026.xlsx e gera docs/data.json
para ser consumido pelo index.html no GitHub Pages.
"""

import json
import sys
from pathlib import Path
from datetime import datetime, date

try:
    from openpyxl import load_workbook
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "--quiet"])
    from openpyxl import load_workbook

XLSX_PATH = Path("Resultado_Gerencial_2026.xlsx")
OUT_PATH  = Path("docs/data.json")

MESES_NOMES = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
               7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}

def safe(v):
    if v is None:
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0

def to_mes_key(v):
    """Converte datetime/date/serial para (ano, mes, nome)."""
    if isinstance(v, (datetime, date)):
        return v.year, v.month, MESES_NOMES.get(v.month, str(v.month))
    try:
        f = float(v)
        # Excel serial
        from datetime import timedelta
        origin = datetime(1899, 12, 30)
        d = origin + timedelta(days=int(f))
        return d.year, d.month, MESES_NOMES.get(d.month, str(d.month))
    except:
        return None

def parse_aba(ws):
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 3:
        return []

    # Row 0: empresa | Previsto | Realizado | Previsto | Realizado ...
    # Row 1: None    | datetime | datetime  | datetime | datetime  ...
    # Row 2+: labels com dados

    pr_row   = rows[0]   # "Previsto" / "Realizado"
    date_row = rows[1]   # datas

    # Mapeia coluna -> {mes, ano, tipo, serial}
    col_map = {}
    serial = 0
    for ci in range(2, len(date_row)):
        info = to_mes_key(date_row[ci])
        if not info:
            continue
        ano, mes_num, mes_nome = info
        tipo_raw = str(pr_row[ci]).strip().lower() if pr_row[ci] else ""
        tipo = "previsto" if "prev" in tipo_raw else "realizado"
        serial += 1
        col_map[ci] = {"mes": mes_nome, "ano": ano, "mes_num": mes_num,
                       "tipo": tipo, "order": ano*100+mes_num}

    # Mapeia label -> índice de linha
    LABEL_MAP = {
        "receita bruta":              "rBruta",
        "(-) despesas do negócio":    "despTotal",
        "resultado":                  "resultado",
        "resultado ":                 "resultado",
        "despesas de pessoal":        "dPessoal",
        "despesas de materia prima":  "dMateria",
        "impostos":                   "dImpostos",
        "taxas plataforma":           "dTaxas",
        "cancelamentos e remmbolso":  "dCancel",
        "despesas de marketing":      "dMarketing",
        "financiamentos":             "dFinanc",
        "financiamentos ":            "dFinanc",
        # canais
        "vendas mercado livre  - vip mx": "mlVip",
        "vendas amazon - vip mx":         "amazonVip",
        "vendas shopee  - vip mx":        "shopeeVip",
        "vendas loja fisica":             "lojaFisica",
        "vendas mercado livre - vidal":   "mlVidal",
        "vendas shopee - vidal":          "shopeeVidal",
        "vendas mercado livre":           "ml",
        "vendas shopee":                  "shopee",
        "vendas amazon":                  "amazon",
    }
    row_data = {}
    for row in rows[2:]:
        lbl = str(row[0]).strip().lower() if row[0] else ""
        if lbl in LABEL_MAP:
            row_data[LABEL_MAP[lbl]] = row

    # Agrupa por (ano, mes_num)
    meses_data = {}
    for ci, info in col_map.items():
        key = (info["ano"], info["mes_num"])
        if key not in meses_data:
            meses_data[key] = {
                "mes": info["mes"], "ano": info["ano"],
                "mes_num": info["mes_num"], "order": info["order"],
                "rPrev":0,"rReal":0,"dPrev":0,"dReal":0,
                "resPrev":0,"resReal":0,
                "dPessoalPrev":0,"dPessoalReal":0,
                "dMateriaPrev":0,"dMateriaReal":0,
                "dImpostosPrev":0,"dImpostosReal":0,
                "dTaxasPrev":0,"dTaxasReal":0,
                "dCancelPrev":0,"dCancelReal":0,
                "dMarketingPrev":0,"dMarketingReal":0,
                "dFinancPrev":0,"dFinancReal":0,
                "canais":{},
            }
        tipo = info["tipo"]
        suf = "Prev" if tipo == "previsto" else "Real"

        def put(fname, rkey):
            if rkey in row_data:
                meses_data[key][fname+suf] += safe(row_data[rkey][ci])

        put("r",        "rBruta")
        put("d",        "despTotal")
        put("res",      "resultado")
        put("dPessoal", "dPessoal")
        put("dMateria", "dMateria")
        put("dImpostos","dImpostos")
        put("dTaxas",   "dTaxas")
        put("dCancel",  "dCancel")
        put("dMarketing","dMarketing")
        put("dFinanc",  "dFinanc")

        CANAL_NAMES = {
            "mlVip":"Mercado Livre","ml":"Mercado Livre","mlVidal":"Mercado Livre",
            "shopeeVip":"Shopee","shopee":"Shopee","shopeeVidal":"Shopee",
            "amazonVip":"Amazon","amazon":"Amazon",
            "lojaFisica":"Loja Física",
        }
        for ck, cname in CANAL_NAMES.items():
            if ck in row_data:
                v = safe(row_data[ck][ci])
                if v == 0:
                    continue
                if cname not in meses_data[key]["canais"]:
                    meses_data[key]["canais"][cname] = {"prev":0,"real":0}
                if tipo == "previsto":
                    meses_data[key]["canais"][cname]["prev"] += v
                else:
                    meses_data[key]["canais"][cname]["real"] += v

    result = sorted(meses_data.values(), key=lambda x: x["order"])
    result = [m for m in result if m["rReal"] != 0 or m["rPrev"] != 0]
    return result


def build_json():
    if not XLSX_PATH.exists():
        print(f"ERRO: {XLSX_PATH} não encontrado.")
        sys.exit(1)

    wb = load_workbook(XLSX_PATH, read_only=True, data_only=True)
    print(f"Abas: {wb.sheetnames}")

    colors = {
        "VIP":   {"color":"#00C9A7","accent":"#00856E","label":"VIP MX"},
        "VIDAL": {"color":"#6C63FF","accent":"#4B44CC","label":"VIDAL"},
        "V3":    {"color":"#FF6B6B","accent":"#CC4444","label":"V3"},
        "GRUPO": {"color":"#F5A623","accent":"#C47D0E","label":"Grupo VIP"},
    }

    aba_map = {
        "VIP":   "PLANO ORÇAMENTARIO - VIP",
        "VIDAL": "PLANO ORÇAMENTARIO - VIDAL",
        "V3":    "PLANO ORÇAMENTARIO - V3",
    }

    data = {}
    for key, aba in aba_map.items():
        if aba not in wb.sheetnames:
            print(f"Aba '{aba}' não encontrada.")
            continue
        months = parse_aba(wb[aba])
        data[key] = {**colors[key], "months": months}
        print(f"  {key}: {len(months)} meses")

    # GRUPO = soma
    all_keys = {}
    for emp in ["VIP","VIDAL","V3"]:
        if emp not in data:
            continue
        for m in data[emp]["months"]:
            k = m["order"]
            if k not in all_keys:
                import copy
                all_keys[k] = copy.deepcopy(m)
                all_keys[k]["canais"] = {}
            else:
                for f in ["rPrev","rReal","dPrev","dReal","resPrev","resReal",
                          "dPessoalPrev","dPessoalReal","dMateriaPrev","dMateriaReal",
                          "dImpostosPrev","dImpostosReal","dTaxasPrev","dTaxasReal",
                          "dCancelPrev","dCancelReal","dMarketingPrev","dMarketingReal",
                          "dFinancPrev","dFinancReal"]:
                    all_keys[k][f] = all_keys[k].get(f,0) + m.get(f,0)
            for cname, cv in m.get("canais",{}).items():
                if cname not in all_keys[k]["canais"]:
                    all_keys[k]["canais"][cname] = {"prev":0,"real":0}
                all_keys[k]["canais"][cname]["prev"] += cv.get("prev",0)
                all_keys[k]["canais"][cname]["real"] += cv.get("real",0)

    data["GRUPO"] = {**colors["GRUPO"],
                     "months": sorted(all_keys.values(), key=lambda x: x["order"])}
    print(f"  GRUPO: {len(data['GRUPO']['months'])} meses")

    output = {
        "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "fonte": XLSX_PATH.name,
        "empresas": data,
    }

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"✅ {OUT_PATH} gerado ({OUT_PATH.stat().st_size} bytes)")

if __name__ == "__main__":
    build_json()
