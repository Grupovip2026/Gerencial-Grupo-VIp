#!/usr/bin/env python3
import json, sys, copy
from pathlib import Path
from datetime import datetime, date, timedelta
try:
    from openpyxl import load_workbook
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "--quiet"])
    from openpyxl import load_workbook

XLSX_PATH = Path("Resultado_Gerencial_2026.xlsx")
OUT_PATH  = Path("docs/data.json")
MESES_NOMES = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
MESES_PARSE = {"jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,"jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12}

def safe(v):
    if v is None: return 0.0
    try: return float(v)
    except: return 0.0

def to_mes_key(v):
    if isinstance(v, (datetime, date)):
        return v.year, v.month, MESES_NOMES.get(v.month, str(v.month))
    try:
        d = datetime(1899,12,30) + timedelta(days=int(float(v)))
        return d.year, d.month, MESES_NOMES.get(d.month, str(d.month))
    except: return None

def parse_mes_str(v):
    if not v: return None
    s = str(v).strip().lower()
    for nome, num in MESES_PARSE.items():
        if s.startswith(nome):
            partes = s.replace(nome,"").strip("/").strip()
            try:
                ano = int(partes)
                if ano < 100: ano += 2000
                return ano, num, MESES_NOMES[num]
            except: pass
    return None

def parse_aba(ws):
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 3: return []
    pr_row, date_row = rows[0], rows[1]
    col_map = {}
    for ci in range(2, len(date_row)):
        info = to_mes_key(date_row[ci])
        if not info: continue
        ano, mes_num, mes_nome = info
        tipo_raw = str(pr_row[ci]).strip().lower() if pr_row[ci] else ""
        tipo = "previsto" if "prev" in tipo_raw else "realizado"
        col_map[ci] = {"mes":mes_nome,"ano":ano,"mes_num":mes_num,"tipo":tipo,"order":ano*100+mes_num}
    LABEL_MAP = {
        "receita bruta":"rBruta","(-) despesas do negocio":"despTotal",
        "(-) despesas do negócio":"despTotal","resultado":"resultado","resultado ":"resultado",
        "despesas de pessoal":"dPessoal","despesas de materia prima":"dMateria",
        "impostos":"dImpostos","taxas plataforma":"dTaxas",
        "cancelamentos e remmbolso":"dCancel","despesas de marketing":"dMarketing",
        "financiamentos":"dFinanc","financiamentos ":"dFinanc",
        "vendas mercado livre  - vip mx":"mlVip","vendas amazon - vip mx":"amazonVip",
        "vendas shopee  - vip mx":"shopeeVip","vendas loja fisica":"lojaFisica",
        "vendas mercado livre - vidal":"mlVidal","vendas shopee - vidal":"shopeeVidal",
        "vendas mercado livre":"ml","vendas shopee":"shopee","vendas amazon":"amazon",
    }
    row_data = {}
    for row in rows[2:]:
        lbl = str(row[0]).strip().lower() if row[0] else ""
        if lbl in LABEL_MAP: row_data[LABEL_MAP[lbl]] = row
    meses_data = {}
    for ci, info in col_map.items():
        key = (info["ano"], info["mes_num"])
        if key not in meses_data:
            meses_data[key] = {"mes":info["mes"],"ano":info["ano"],"mes_num":info["mes_num"],"order":info["order"],
                "rPrev":0,"rReal":0,"dPrev":0,"dReal":0,"resPrev":0,"resReal":0,
                "dPessoalPrev":0,"dPessoalReal":0,"dMateriaPrev":0,"dMateriaReal":0,
                "dImpostosPrev":0,"dImpostosReal":0,"dTaxasPrev":0,"dTaxasReal":0,
                "dCancelPrev":0,"dCancelReal":0,"dMarketingPrev":0,"dMarketingReal":0,
                "dFinancPrev":0,"dFinancReal":0,"canais":{}}
        suf = "Prev" if info["tipo"]=="previsto" else "Real"
        for fname,rkey in [("r","rBruta"),("d","despTotal"),("res","resultado"),
            ("dPessoal","dPessoal"),("dMateria","dMateria"),("dImpostos","dImpostos"),
            ("dTaxas","dTaxas"),("dCancel","dCancel"),("dMarketing","dMarketing"),("dFinanc","dFinanc")]:
            if rkey in row_data: meses_data[key][fname+suf] += safe(row_data[rkey][ci])
        for ck,cname in [("mlVip","Mercado Livre"),("ml","Mercado Livre"),("mlVidal","Mercado Livre"),
            ("shopeeVip","Shopee"),("shopee","Shopee"),("shopeeVidal","Shopee"),
            ("amazonVip","Amazon"),("amazon","Amazon"),("lojaFisica","Loja Fisica")]:
            if ck in row_data:
                v = safe(row_data[ck][ci])
                if v == 0: continue
                if cname not in meses_data[key]["canais"]: meses_data[key]["canais"][cname]={"prev":0,"real":0}
                if info["tipo"]=="previsto": meses_data[key]["canais"][cname]["prev"]+=v
                else: meses_data[key]["canais"][cname]["real"]+=v
    result = sorted(meses_data.values(), key=lambda x: x["order"])
    meses_com_real = [m for m in result if m["rReal"]>0]
    if meses_com_real:
        ultimo = meses_com_real[-1]["order"]
        return [m for m in result if m["order"]<=ultimo]
    return []

def parse_divida_fornecedores(wb):
    aba_nome = next((s for s in wb.sheetnames if "divida" in s.lower() or "dívida" in s.lower()), None)
    if not aba_nome:
        print("  [Divida Fornecedores] aba NAO encontrada! Abas:", wb.sheetnames)
        return []
    print(f"  [Divida Fornecedores] lendo aba '{aba_nome}'...")
    ws = wb[aba_nome]
    rows = list(ws.iter_rows(values_only=True))
    COL = {
        "pagas_vip_mes":1,   "pagas_vip_val":2,
        "pagas_vidal_mes":4, "pagas_vidal_val":5,
        "pagar_vip_mes":9,   "pagar_vip_val":10,
        "pagar_vidal_mes":12,"pagar_vidal_val":13,
    }
    result = {}
    def get(row, idx):
        return row[idx] if len(row) > idx else None
    for row in rows[6:]:
        blocos = [
            ("pagas_vip",   "pagas_vip_mes",   "pagas_vip_val"),
            ("pagas_vidal", "pagas_vidal_mes",  "pagas_vidal_val"),
            ("pagar_vip",   "pagar_vip_mes",    "pagar_vip_val"),
            ("pagar_vidal", "pagar_vidal_mes",  "pagar_vidal_val"),
        ]
        for bloco, col_mes, col_val in blocos:
            v_mes = get(row, COL[col_mes])
            v_val = get(row, COL[col_val])
            if not v_mes: continue
            info = parse_mes_str(v_mes)
            if not info: continue
            ano, mes_num, mes_nome = info
            order = ano * 100 + mes_num
            if order not in result:
                result[order] = {
                    "mes": mes_nome, "mes_num": mes_num, "ano": ano, "order": order,
                    "pagas_vip": 0.0, "pagas_vidal": 0.0,
                    "pagar_vip": 0.0, "pagar_vidal": 0.0,
                }
            valor = safe(v_val)
            if valor != 0:
                result[order][bloco] += valor
    lista = []
    for m in sorted(result.values(), key=lambda x: x["order"]):
        m["pagas_grupo"] = m["pagas_vip"] + m["pagas_vidal"]
        m["pagar_grupo"] = m["pagar_vip"] + m["pagar_vidal"]
        m["total_vip"]   = m["pagas_vip"] + m["pagar_vip"]
        m["total_vidal"] = m["pagas_vidal"] + m["pagar_vidal"]
        m["total_grupo"] = m["pagas_grupo"] + m["pagar_grupo"]
        lista.append(m)
    print(f"  [Divida Fornecedores] {len(lista)} meses: {[m['mes'] for m in lista]}")
    return lista

def build_json():
    if not XLSX_PATH.exists(): print(f"ERRO: {XLSX_PATH} nao encontrado."); sys.exit(1)
    wb = load_workbook(XLSX_PATH, read_only=True, data_only=True)
    colors = {"VIP":{"color":"#00C9A7","accent":"#00856E","label":"VIP MX"},
              "VIDAL":{"color":"#6C63FF","accent":"#4B44CC","label":"VIDAL"},
              "V3":{"color":"#FF6B6B","accent":"#CC4444","label":"V3"},
              "GRUPO":{"color":"#F5A623","accent":"#C47D0E","label":"Grupo VIP"}}
    aba_map = {"VIP":"PLANO ORCAMENTARIO - VIP","VIDAL":"PLANO ORCAMENTARIO - VIDAL","V3":"PLANO ORCAMENTARIO - V3"}
    data = {}
    for key, aba in aba_map.items():
        real_aba = next((s for s in wb.sheetnames if s.upper().replace("Ç","C").replace("Ã","A") == aba), None)
        if not real_aba: continue
        months = parse_aba(wb[real_aba])
        data[key] = {**colors[key], "months": months}
        print(f"  {key}: {len(months)} meses")
    all_keys = {}
    for emp in ["VIP","VIDAL","V3"]:
        if emp not in data: continue
        for m in data[emp]["months"]:
            k = m["order"]
            if k not in all_keys: all_keys[k] = copy.deepcopy(m); all_keys[k]["canais"]={}
            else:
                for f in ["rPrev","rReal","dPrev","dReal","resPrev","resReal",
                    "dPessoalPrev","dPessoalReal","dMateriaPrev","dMateriaReal",
                    "dImpostosPrev","dImpostosReal","dTaxasPrev","dTaxasReal",
                    "dCancelPrev","dCancelReal","dMarketingPrev","dMarketingReal","dFinancPrev","dFinancReal"]:
                    all_keys[k][f] = all_keys[k].get(f,0)+m.get(f,0)
            for cname,cv in m.get("canais",{}).items():
                if cname not in all_keys[k]["canais"]: all_keys[k]["canais"][cname]={"prev":0,"real":0}
                all_keys[k]["canais"][cname]["prev"]+=cv.get("prev",0)
                all_keys[k]["canais"][cname]["real"]+=cv.get("real",0)
    data["GRUPO"] = {**colors["GRUPO"],"months":sorted(all_keys.values(),key=lambda x:x["order"])}
    financeiro = parse_divida_fornecedores(wb)
    print(f"  Financeiro: {len(financeiro)} meses")
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    out = {"gerado_em":datetime.now().strftime("%d/%m/%Y %H:%M"),"fonte":XLSX_PATH.name,
           "empresas":data,"financeiro":financeiro}
    with open(OUT_PATH,"w",encoding="utf-8") as f: json.dump(out,f,ensure_ascii=False,indent=2)
    print(f"OK {OUT_PATH} gerado ({OUT_PATH.stat().st_size} bytes)")

if __name__ == "__main__": build_json()
