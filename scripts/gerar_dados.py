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
MESES_PT = {1:"JANEIRO",2:"FEVEREIRO",3:"MARCO",4:"ABRIL",5:"MAIO",6:"JUNHO",7:"JULHO",8:"AGOSTO",9:"SETEMBRO",10:"OUTUBRO",11:"NOVEMBRO",12:"DEZEMBRO"}

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
    if v is None: return None
    if isinstance(v, (datetime, date)):
        return v.year, v.month, MESES_NOMES.get(v.month, str(v.month))
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
    aba_nome = next((s for s in wb.sheetnames if 'fornecedores' in s.lower()), None)
    if not aba_nome:
        print("  [Divida Fornecedores] aba NAO encontrada!")
        return []
    print(f"  [Divida Fornecedores] lendo aba '{aba_nome}'...")
    ws = wb[aba_nome]
    rows = list(ws.iter_rows(values_only=True))
    COL = {
        "pagas_vip_mes":1,   "pagas_vip_val":2,
        "pagas_vidal_mes":5, "pagas_vidal_val":6,
        "pagar_vip_mes":10,  "pagar_vip_val":11,
        "pagar_vidal_mes":14,"pagar_vidal_val":15,
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

def parse_dre(wb):
    LABEL_MAP = {
        "receita bruta de vendas": "recBruta",
        "(-) devoluções e cancelamentos": "devol",
        "(-) impostos sobre vendas": "impostos",
        "receita líquida": "recLiq",
        "(-) custo de mercadoria vendida (cmv)": "cmv",
        "= lucro bruto": "lucroBruto",
        "(-) despesas de pessoal": "dPessoal",
        "(-) despesas de marketing": "dMarketing",
        "(-) taxas de plataforma": "dTaxas",
        "(-) despesas financeiras / financiamentos": "dFinanc",
        "(-) outras despesas operacionais": "dOutras",
        "ebitda": "ebitda",
        "(-) depreciação / amortização": "deprec",
        "ebit": "ebit",
        "lucro líquido": "lucroLiq",
    }
    CAMPOS = ["recBruta","devol","impostos","recLiq","cmv","lucroBruto",
              "dPessoal","dMarketing","dTaxas","dFinanc","dOutras",
              "ebitda","deprec","ebit","lucroLiq"]
    def ler_aba(ws):
        rows = list(ws.iter_rows(values_only=True))
        date_row = rows[3] if len(rows) > 3 else []
        col_mes = {}
        for ci in range(5, len(date_row)):
            info = to_mes_key(date_row[ci])
            if info:
                ano, mes_num, mes_nome = info
                col_mes[ci] = {"order": ano*100+mes_num, "mes": mes_nome, "mes_num": mes_num, "ano": ano}
        meses = {}
        for ci, info in col_mes.items():
            k = info["order"]
            if k not in meses:
                meses[k] = {"mes": info["mes"], "mes_num": info["mes_num"], "ano": info["ano"], "order": k}
                for c in CAMPOS: meses[k][c] = 0.0
        i = 0
        while i < len(rows):
            row = rows[i]
            lbl_raw = row[5] if len(row) > 5 else None
            if lbl_raw and isinstance(lbl_raw, str):
                lbl = lbl_raw.strip().lower()
                if lbl in LABEL_MAP:
                    campo = LABEL_MAP[lbl]
                    if i+1 < len(rows):
                        val_row = rows[i+1]
                        for ci, info in col_mes.items():
                            k = info["order"]
                            v = val_row[ci] if len(val_row) > ci else None
                            meses[k][campo] = safe(v)
                    i += 2
                    continue
            i += 1
        return sorted(meses.values(), key=lambda x: x["order"])
    resultado = {}
    for emp, aba_nome in [("VIP", "DRE - VIP"), ("VIDAL", "DRE - VIDAL")]:
        real = next((s for s in wb.sheetnames if s.strip() == aba_nome), None)
        if not real:
            print(f"  [DRE] aba '{aba_nome}' NAO encontrada!")
            resultado[emp] = []
            continue
        resultado[emp] = ler_aba(wb[real])
        print(f"  [DRE] {emp}: {len(resultado[emp])} meses")
    grupo = {}
    for emp in ["VIP", "VIDAL"]:
        for m in resultado.get(emp, []):
            k = m["order"]
            if k not in grupo:
                grupo[k] = {"mes": m["mes"], "mes_num": m["mes_num"], "ano": m["ano"], "order": k}
                for c in CAMPOS: grupo[k][c] = 0.0
            for c in CAMPOS:
                grupo[k][c] += m.get(c, 0.0)
    resultado["GRUPO"] = sorted(grupo.values(), key=lambda x: x["order"])
    print(f"  [DRE] GRUPO: {len(resultado['GRUPO'])} meses consolidados")
    return resultado

def parse_custos(wb):
    aba = next((s for s in wb.sheetnames if s.strip().lower() == 'custos'), None)
    if not aba:
        print('  [Custos] aba NAO encontrada!')
        return {'fixo': [], 'variavel': [], 'total_fixo': 0, 'total_variavel': 0}
    ws = wb[aba]
    rows = list(ws.iter_rows(values_only=True))
    fixo, variavel = [], []
    for row in rows[4:]:
        origem_f = str(row[1]).strip() if len(row) > 1 and row[1] else None
        valor_f = row[2] if len(row) > 2 else None
        if origem_f and origem_f.upper().strip() not in ('ORIGEM', 'VALOR ESTIMADO') and valor_f:
            try:
                v = float(valor_f)
                if v != 0: fixo.append({'origem': origem_f, 'valor': v})
            except: pass
        origem_v = str(row[5]).strip() if len(row) > 5 and row[5] else None
        valor_v = row[6] if len(row) > 6 else None
        if origem_v and origem_v.upper().strip() not in ('ORIGEM', 'VALOR ESTIMADO') and valor_v:
            try:
                v = float(valor_v)
                if v != 0: variavel.append({'origem': origem_v, 'valor': v})
            except: pass
    total_fixo = sum(i['valor'] for i in fixo)
    total_variavel = sum(i['valor'] for i in variavel)
    print(f'  [Custos] fixo: {len(fixo)} itens | variavel: {len(variavel)} itens')
    return {'fixo': fixo, 'variavel': variavel, 'total_fixo': total_fixo, 'total_variavel': total_variavel}

def parse_fluxo_caixa(wb):
    def parse_aba_fluxo(ws):
        rows = list(ws.iter_rows(values_only=True))
        meses = {}
        dias = []
        mes_atual = None
        saldo_inicial_set = False
        for row in rows:
            b = row[1] if len(row) > 1 else None
            # Detecta cabecalho de mes
            if isinstance(b, str) and b.strip() and b.strip().upper() not in ('DATA','TOTAL'):
                b_norm = b.upper().replace('Ç','C').replace('Ã','A').replace('É','E').replace('Á','A')
                for num, nome_pt in MESES_PT.items():
                    if nome_pt in b_norm or MESES_NOMES[num].upper() in b_norm:
                        ano = 2025 if '25' in b else 2026
                        mes_atual = {
                            'mes': MESES_NOMES[num], 'mes_num': num, 'ano': ano,
                            'order': ano*100+num,
                            'entradas': 0.0, 'resgates': 0.0, 'aportes': 0.0,
                            'despesas': 0.0, 'aplicacoes': 0.0, 'shareholder': 0.0,
                            'saldo_inicial': 0.0, 'saldo_final': 0.0
                        }
                        saldo_inicial_set = False
                        break
                continue
            if mes_atual is None: continue
            if isinstance(b, str) and b.strip().upper() == 'DATA': continue
            if len(row) > 2 and isinstance(row[2], str): continue
            # Linha saldo inicial
            if b is None and not saldo_inicial_set:
                v = row[8] if len(row) > 8 else None
                if v is not None and not isinstance(v, str):
                    mes_atual['saldo_inicial'] = safe(v)
                    saldo_inicial_set = True
                continue
            # Linha TOTAL
            if isinstance(b, str) and 'TOTAL' in b.upper():
                mes_atual['entradas']    = safe(row[2] if len(row)>2 else None)
                mes_atual['resgates']    = safe(row[3] if len(row)>3 else None)
                mes_atual['aportes']     = safe(row[4] if len(row)>4 else None)
                mes_atual['despesas']    = safe(row[5] if len(row)>5 else None)
                mes_atual['aplicacoes']  = safe(row[6] if len(row)>6 else None)
                mes_atual['shareholder'] = safe(row[7] if len(row)>7 else None)
                meses[mes_atual['order']] = mes_atual
                mes_atual = None
                continue
            # Linha diaria
            if isinstance(b, (datetime, date)):
                saldo = safe(row[8] if len(row) > 8 else None)
                if mes_atual:
                    mes_atual['saldo_final'] = saldo
                dia_str = b.strftime('%d/%m') if isinstance(b, (datetime, date)) else str(b)
                mes_num = b.month if isinstance(b, datetime) else b.month
                mes_nome = MESES_NOMES.get(mes_num, '')
                ano = b.year if isinstance(b, datetime) else b.year
                dias.append({
                    'data': dia_str,
                    'dia': b.day if isinstance(b, (datetime, date)) else 0,
                    'mes': mes_nome,
                    'mes_num': mes_num,
                    'ano': ano,
                    'order_mes': ano*100+mes_num,
                    'order': ano*10000+mes_num*100+(b.day if isinstance(b, (datetime, date)) else 0),
                    'entradas': safe(row[2] if len(row)>2 else None),
                    'resgates': safe(row[3] if len(row)>3 else None),
                    'aportes': safe(row[4] if len(row)>4 else None),
                    'despesas': safe(row[5] if len(row)>5 else None),
                    'aplicacoes': safe(row[6] if len(row)>6 else None),
                    'shareholder': safe(row[7] if len(row)>7 else None),
                    'saldo': saldo,
                })
        meses_lista = sorted(meses.values(), key=lambda x: x['order'])
        dias_lista = sorted(dias, key=lambda x: x['order'])
        # Filtra apenas dias com dados reais
        dias_lista = [d for d in dias_lista if d['saldo'] != 0 or d['despesas'] != 0 or d['entradas'] != 0]
        return meses_lista, dias_lista

    resultado = {}
    for emp, aba_nome in [('VIP','Fluxo De caixa - VIP'),('VIDAL','Fluxo De caixa - VIDAL')]:
        real = next((s for s in wb.sheetnames if s.strip() == aba_nome), None)
        if not real:
            print(f'  [Fluxo] aba {aba_nome} NAO encontrada!')
            resultado[emp] = {'meses': [], 'dias': []}
            continue
        meses, dias = parse_aba_fluxo(wb[real])
        meses = [m for m in meses if m['saldo_final'] != 0 or m['despesas'] != 0 or m['entradas'] != 0]
        resultado[emp] = {'meses': meses, 'dias': dias}
        print(f'  [Fluxo] {emp}: {len(meses)} meses, {len(dias)} dias')

    # Consolida GRUPO
    grupo_meses = {}
    grupo_dias = {}
    for emp in ['VIP','VIDAL']:
        for m in resultado.get(emp, {}).get('meses', []):
            k = m['order']
            if k not in grupo_meses:
                grupo_meses[k] = {'mes':m['mes'],'mes_num':m['mes_num'],'ano':m['ano'],'order':k,
                    'entradas':0.0,'resgates':0.0,'aportes':0.0,'despesas':0.0,
                    'aplicacoes':0.0,'shareholder':0.0,'saldo_inicial':0.0,'saldo_final':0.0}
            for f in ['entradas','resgates','aportes','despesas','aplicacoes','shareholder','saldo_final']:
                grupo_meses[k][f] += m.get(f, 0.0)
        for d in resultado.get(emp, {}).get('dias', []):
            k = d['order']
            if k not in grupo_dias:
                grupo_dias[k] = {'data':d['data'],'dia':d['dia'],'mes':d['mes'],'mes_num':d['mes_num'],
                    'ano':d['ano'],'order_mes':d['order_mes'],'order':k,
                    'entradas':0.0,'resgates':0.0,'aportes':0.0,'despesas':0.0,
                    'aplicacoes':0.0,'shareholder':0.0,'saldo':0.0}
            for f in ['entradas','resgates','aportes','despesas','aplicacoes','shareholder','saldo']:
                grupo_dias[k][f] += d.get(f, 0.0)

    resultado['GRUPO'] = {
        'meses': sorted(grupo_meses.values(), key=lambda x: x['order']),
        'dias': sorted(grupo_dias.values(), key=lambda x: x['order'])
    }
    print(f'  [Fluxo] GRUPO: {len(resultado["GRUPO"]["meses"])} meses, {len(resultado["GRUPO"]["dias"])} dias consolidados')
    return resultado

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
    dre = parse_dre(wb)
    custos = parse_custos(wb)
    fluxo = parse_fluxo_caixa(wb)
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    out = {"gerado_em":datetime.now().strftime("%d/%m/%Y %H:%M"),"fonte":XLSX_PATH.name,
           "empresas":data,"financeiro":financeiro,"dre":dre,"custos":custos,"fluxo":fluxo}
    with open(OUT_PATH,"w",encoding="utf-8") as f: json.dump(out,f,ensure_ascii=False,indent=2)
    print(f"OK {OUT_PATH} gerado ({OUT_PATH.stat().st_size} bytes)")

if __name__ == "__main__": build_json()