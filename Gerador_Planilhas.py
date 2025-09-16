# -*- coding: utf-8 -*-
"""
Gerador_Planilhas.py — multi-temas com estilos
"""

import sys, random
from typing import Dict, Any, List, Tuple, Optional
from datetime import datetime, timedelta
import numpy as np
import pandas as pd

# ========= util de prompt =========
def prompt_menu(title: str, options: List[str], default_idx: Optional[int]=None) -> int:
    print(f"\n{title}")
    for i,opt in enumerate(options,1):
        print(f"  {i}) {opt}")
    while True:
        resp = input(f"Escolha [1-{len(options)}]{f' (padrão {default_idx+1})' if default_idx is not None else ''}: ").strip()
        if not resp and default_idx is not None: return default_idx
        if resp.isdigit():
            idx=int(resp)-1
            if 0<=idx<len(options): return idx
        print("Entrada inválida.")

def prompt_int(msg: str, default: Optional[int]=None, min_val: Optional[int]=None, max_val: Optional[int]=None) -> int:
    while True:
        s = input(f"{msg}{f' (padrão {default})' if default is not None else ''}: ").strip()
        if not s and default is not None: return default
        try:
            v=int(s)
            if (min_val is not None and v<min_val) or (max_val is not None and v>max_val):
                print("Fora do intervalo."); continue
            return v
        except: print("Número inválido.")

def parse_ranges_to_indices(expr: str, total: int) -> List[int]:
    if not expr: return list(range(total))
    parts=[p.strip() for p in expr.split(",") if p.strip()]
    idxs=set()
    for p in parts:
        if "-" in p:
            a,b=p.split("-",1)
            if a.isdigit() and b.isdigit():
                a,b=int(a),int(b)
                for k in range(min(a,b),max(a,b)+1):
                    if 1<=k<=total: idxs.add(k-1)
        elif p.isdigit():
            k=int(p); 
            if 1<=k<=total: idxs.add(k-1)
    return sorted(idxs)

# ========= estilos =========
ESTILOS = {
    "Azul":   {"header_bg":"#E8F1FF","kpi_bg":"#DCEBFF","zebra":"#F7FAFF","neg":"#FCE8E6","pos":"#E6F4EA","scale_min":"#E8F1FF","scale_max":"#2B6CB0"},
    "Verde":  {"header_bg":"#E9F7EF","kpi_bg":"#D4EFDF","zebra":"#F5FBF7","neg":"#FDEDEC","pos":"#E8F8F5","scale_min":"#E9F7EF","scale_max":"#1E8449"},
    "Cinza":  {"header_bg":"#F0F0F0","kpi_bg":"#E6E6E6","zebra":"#FAFAFA","neg":"#FDEDEC","pos":"#EBF5FB","scale_min":"#F0F0F0","scale_max":"#5D6D7E"},
    "Laranja":{"header_bg":"#FFF1E6","kpi_bg":"#FFE0CC","zebra":"#FFF9F3","neg":"#FDECEA","pos":"#FFF7E6","scale_min":"#FFF1E6","scale_max":"#D35400"},
}

# ========= motor xlsx =========
def _excel_cell_to_tuple(cell_ref: str) -> Tuple[int,int]:
    col=0; row=0
    for c in cell_ref:
        if c.isalpha(): col=col*26 + (ord(c.upper())-ord('A'))+1
        else: row=row*10 + int(c)
    return (row-1, col-1)

def _apply_common_formats(workbook, estilo_key: str):
    pal = ESTILOS[estilo_key]
    fmt = {
        "text":   workbook.add_format({'text_wrap':False}),
        "int":    workbook.add_format({'num_format':'#,##0'}),
        "float":  workbook.add_format({'num_format':'#,##0.00'}),
        "money":  workbook.add_format({'num_format':'R$ #,##0.00'}),
        "date":   workbook.add_format({'num_format':'yyyy-mm-dd'}),
        "header": workbook.add_format({'bold':True,'bg_color':pal["header_bg"],'border':1}),
        "kpi_lbl":workbook.add_format({'bold':True}),
        "kpi_val":workbook.add_format({'num_format':'#,##0.00','bold':True,'bg_color':pal["kpi_bg"],'border':1}),
    }
    return fmt, pal

def gerar_planilha(spec: Dict[str, Any], output_path: str, estilo_key: str="Azul") -> None:
    sheets_spec = spec.get('sheets', [])
    pivots_spec = spec.get('pivots', [])
    kpis_spec   = spec.get('kpis', [])
    dashboard_name = spec.get('dashboard_name', 'Dashboard')
    with pd.ExcelWriter(output_path, engine='xlsxwriter', datetime_format='yyyy-mm-dd', date_format='yyyy-mm-dd') as writer:
        workbook = writer.book
        fmt, pal = _apply_common_formats(workbook, estilo_key)
        name_to_df = {}

        # abas
        for sh in sheets_spec:
            name = sh['name']; data = sh.get('data', pd.DataFrame())
            if isinstance(data, list): data = pd.DataFrame(data)
            df = data.copy(); name_to_df[name]=df
            df.to_excel(writer, sheet_name=name, index=False, startrow=1)
            ws = writer.sheets[name]

            for i, colname in enumerate(df.columns):
                ws.write(0, i, colname, fmt["header"])

            for col in sh.get('columns', []):
                colname=col.get('name'); width=col.get('width',15); f=col.get('fmt','text')
                if colname in df.columns:
                    ci=df.columns.get_loc(colname)
                    base = fmt["text"] if f=="text" else fmt.get("money" if f=="currency" else f, fmt["text"])
                    ws.set_column(ci, ci, width, base)

            if sh.get('autofilter', True) and not df.empty:
                ws.autofilter(0,0, df.shape[0], df.shape[1]-1)
            if sh.get('freeze'): ws.freeze_panes(*_excel_cell_to_tuple(sh['freeze']))

            if not df.empty:
                ws.conditional_format(1,0, df.shape[0], df.shape[1]-1, {
                    'type':'formula','criteria':'=MOD(ROW(),2)=0','format':workbook.add_format({'bg_color':pal["zebra"]})
                })

            if "validade" in df.columns:
                ci_val = df.columns.get_loc("validade")
                ws.conditional_format(1,0, df.shape[0], df.shape[1]-1, {
                    'type':'formula','criteria': f'=INDIRECT(ADDRESS(ROW(),{ci_val+1}))<=TODAY()+7',
                    'format':workbook.add_format({'bg_color':pal["neg"]})
                })
            if {"vencimento","pago"}.issubset(df.columns):
                ci_v=df.columns.get_loc("vencimento"); ci_p=df.columns.get_loc("pago")
                ws.conditional_format(1,0, df.shape[0], df.shape[1]-1, {
                    'type':'formula','criteria': f'=AND(TODAY()>INDIRECT(ADDRESS(ROW(),{ci_v+1})),INDIRECT(ADDRESS(ROW(),{ci_p+1}))=FALSE())',
                    'format':workbook.add_format({'bg_color':pal["neg"]})
                })

        # KPIs
        if kpis_spec:
            if dashboard_name not in writer.sheets:
                pd.DataFrame().to_excel(writer, sheet_name=dashboard_name, index=False)
            ws = writer.sheets[dashboard_name]
            ws.write(0,0,"KPIs", fmt["header"]); r=2
            for k in kpis_spec:
                ws.write(r,0,k.get("label","KPI"), fmt["kpi_lbl"])
                if "formula" in k:
                    ws.write_formula(r,1,k["formula"], fmt["kpi_val"])
                else:
                    val=k.get("value",""); f=k.get("fmt","text")
                    cellfmt = fmt["kpi_val"] if f in ("float","int","currency") else fmt["text"]
                    ws.write(r,1,val, cellfmt)
                r+=1

        # pivôs
        for pv in pivots_spec:
            name=pv['name']; src_sheet=pv['data_sheet']
            if src_sheet not in name_to_df: continue
            src=name_to_df[src_sheet]
            if src.empty:
                pd.DataFrame().to_excel(writer, sheet_name=name, index=False); continue
            pvt=pd.pivot_table(
                src,
                index=pv.get('index',[]),
                columns=pv.get('columns',[]),
                values=list(pv.get('values', {'valor':'sum'}).keys()),
                aggfunc=pv.get('values', {'valor':'sum'}),
                fill_value=pv.get('fill_value',0)
            )
            if isinstance(pvt.columns, pd.MultiIndex):
                pvt.columns=[' | '.join(map(str,c)).strip() for c in pvt.columns.values]
            pvt=pvt.reset_index()
            rnd=pv.get('round')
            if isinstance(rnd,int):
                nums=pvt.select_dtypes(include=[np.number]).columns; pvt[nums]=pvt[nums].round(rnd)
            pvt.to_excel(writer, sheet_name=name, index=False, startrow=1)
            ws=writer.sheets[name]
            for i,colname in enumerate(pvt.columns):
                ws.write(0,i,colname, fmt["header"]); ws.set_column(i,i,max(12,len(str(colname))+2))
            ch=pv.get('chart')
            if ch and not pvt.empty:
                chart=workbook.add_chart({'type': ch.get('type','column')})
                for col_idx in range(1,pvt.shape[1]):
                    chart.add_series({'name':[name,0,col_idx],'categories':[name,1,0,pvt.shape[0],0],'values':[name,1,col_idx,pvt.shape[0],col_idx]})
                chart.set_title({'name': ch.get('title',name)})
                chart.set_y_axis({'name': ch.get('y_title','')})
                ws.insert_chart('B8', chart, {'x_scale':1.2,'y_scale':1.2})

# ========= faker / bases =========
UFs = ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"]

_FAKE=None; _FAKER_OK=False; _COMMERCE=False
try:
    from faker import Faker
    _FAKE = Faker("pt_BR"); _FAKER_OK=True
    try:
        from faker_commerce import Provider as CommerceProvider
        _FAKE.add_provider(CommerceProvider); _COMMERCE=True
    except Exception:
        _COMMERCE=False
except Exception:
    _FAKER_OK=False

MARCAS_BR = ["Aurora","Predilecta","Nestlé","Camil","Ypê","Itambé","Seara","Qualitá","Heinz","Coca-Cola","Ambev","Vitao","Italac","Piracanjuba","Piraquê","Tio João","União","Colgate","Oral-B","Tramontina","Vigor","Sadia","Perdigão","Bauducco","Santa Helena","Fini","Bombril","Brilux","Scotch-Brite"]

def _fake_estado_sigla():
    if _FAKER_OK and hasattr(_FAKE, "estado_sigla"): return _FAKE.estado_sigla()
    return random.choice(UFs)

def _rand_date(days_back=365):
    end = datetime.now(); start = end - timedelta(days=days_back)
    return start + timedelta(seconds=random.randint(0,int((end-start).total_seconds())))

def _escolha_ponderada(opcoes):
    itens, pesos = zip(*opcoes); tot=sum(pesos); r=random.uniform(0,tot); acc=0
    for it,p in zip(itens,pesos):
        acc+=p
        if r<=acc: return it
    return itens[-1]

def _doc_fakes():
    return {"cnpj": f"{random.randint(10,99)}.{random.randint(100,999)}.{random.randint(100,999)}/0001-{random.randint(10,99)}",
            "cpf":  f"{random.randint(100,999)}.{random.randint(100,999)}.{random.randint(100,999)}-{random.randint(10,99)}",
            "ie":   f"{random.randint(1000000,9999999)}"}

# ========= EAN =========
def _ean13_checksum(num12: str) -> int:
    s = sum((3 if i%2 else 1)*int(d) for i,d in enumerate(num12[::-1]))
    return (10 - (s % 10)) % 10
def gerar_ean13(prefix: str="789") -> str:
    base_len = 12 - len(prefix); middle = "".join(str(random.randint(0,9)) for _ in range(base_len))
    num12 = prefix + middle; dv = _ean13_checksum(num12); return num12 + str(dv)

# ========= produtos PT-BR para varejo =========
_UNIDADES = [
    ("g", 140, 0.5), ("g", 200, 0.7), ("g", 500, 0.9), ("g", 1000, 1.0),
    ("ml", 300, 0.8), ("ml", 500, 1.0), ("L", 1, 1.2), ("L", 2, 1.9),
    ("un", 1, 1.0), ("un", 4, 3.6), ("un", 6, 5.2), ("un", 12, 10.0)
]
_BASE_PRECO_PT = {
    "Mercearia": (5.90, 29.90), "Bebidas": (4.90, 39.90), "Higiene & Beleza": (7.90, 49.90),
    "Limpeza": (5.90, 29.90), "Frios & Laticínios": (7.90, 59.90), "Açougue": (14.90, 79.90),
    "Padaria & Confeitaria": (4.90, 24.90), "Pets": (9.90, 49.90), "Utilidades": (9.90, 69.90),
}
CAT_PT = {
    "Mercearia": ["Arroz","Feijão Carioca","Feijão Preto","Macarrão Spaghetti","Macarrão Parafuso","Molho de Tomate","Azeite de Oliva","Açúcar Refinado","Farinha de Trigo","Café Torrado e Moído","Atum em Óleo","Sardinha em Óleo","Azeitona Verde","Milho Verde","Ervilha em Conserva","Biscoito Recheado Chocolate","Biscoito Cream Cracker","Achocolatado em Pó","Granola Tradicional","Aveia em Flocos","Leite em Pó","Leite Condensado","Creme de Leite"],
    "Bebidas": ["Água Mineral","Refrigerante Cola","Refrigerante Guaraná","Suco de Uva","Suco de Laranja","Cerveja Pilsen","Cerveja IPA","Vinho Tinto Seco","Chá Gelado","Água de Coco","Energético"],
    "Higiene & Beleza": ["Sabonete","Shampoo","Condicionador","Desodorante Aerosol","Creme Dental","Escova Dental","Fio Dental","Enxaguante Bucal","Lenço Umedecido"],
    "Limpeza": ["Detergente Líquido Neutro","Desinfetante","Amaciante de Roupas","Sabão em Pó","Limpador Multiuso","Água Sanitária","Esponja Multiuso","Lustra Móveis","Saco de Lixo"],
    "Frios & Laticínios": ["Queijo Mussarela Fatiado","Queijo Prato","Queijo Parmesão Ralado","Presunto Cozido","Peito de Peru","Iogurte Natural","Iogurte Grego","Requeijão Cremoso","Manteiga com Sal","Ricota Fresca","Cream Cheese"],
    "Açougue": ["Frango Congelado","Coxa e Sobrecoxa","Peito de Frango","Carne Moída","Bife de Alcatra","Carne Suína em Cubos","Linguiça Toscana","Carne para Panela"],
    "Padaria & Confeitaria": ["Pão Francês","Pão de Forma","Bolo de Chocolate","Bolo de Cenoura","Pão de Queijo","Croissant","Biscoito Amanteigado"],
    "Pets": ["Ração Cães Adultos","Ração Gatos Adultos","Petisco para Cães","Areia Sanitária"],
    "Utilidades": ["Pano de Prato","Esponja de Aço","Vassoura","Rodo","Balde Plástico"],
}
ADJETIVOS = ["Premium","Tradicional","Integral","Zero Açúcar","Zero Lactose","Light","Orgânico","Clássico","Caseiro","Intenso","Extra Forte","Sabor Chocolate","Sabor Morango","Sabor Baunilha"]

def _preco_realista_pt(familia: str, unidade: Tuple[str, float, float]) -> float:
    low, high = _BASE_PRECO_PT.get(familia, (7.90, 49.90))
    base = random.uniform(low, high)
    _, _, fator = unidade
    brand_bump = random.choice([0.95, 1.0, 1.05, 1.1])
    preco = base * fator * brand_bump
    cents = random.choice([0.90, 0.99, 0.79, 0.49, 0.19])
    return float(int(preco)) + cents

def produto_pt_br() -> Dict[str, Any]:
    familia = random.choice(list(CAT_PT.keys()))
    base = random.choice(CAT_PT[familia])
    adic = " " + random.choice(ADJETIVOS) if random.random()<0.35 else ""
    unidade = random.choice(_UNIDADES)
    unidade_str = f"{unidade[1]}{unidade[0]}"
    nome = f"{base}{adic} {unidade_str}"
    marca = random.choice(MARCAS_BR + ["Genérico","Local","Premium","Eco"])
    ean = gerar_ean13("789")
    sku = f"{familia[:2].upper()}-{random.randint(10000,99999)}"
    preco_base = round(_preco_realista_pt(familia, unidade), 2)
    return {"sku": sku, "ean13": ean, "produto": nome, "categoria": familia, "marca": marca, "unidade": unidade_str, "preco_base": preco_base}

# ========= clientes =========
def _cliente():
    if _FAKER_OK:
        nome=_FAKE.name(); empresa=_FAKE.company(); cidade=_FAKE.city(); uf=_fake_estado_sigla(); cep=_FAKE.postcode()
    else:
        nome=f"Cliente {random.randint(1000,9999)}"; empresa=f"Empresa {random.randint(100,999)} Ltda"; cidade=f"Cidade {random.randint(1,200)}"; uf=random.choice(UFs); cep=f"{random.randint(10000,99999)}-{random.randint(100,999)}"
    seg=_escolha_ponderada([("Varejo",0.5),("Atacado",0.3),("E-commerce",0.2)])
    return {"cliente_nome":nome,"empresa":empresa,"cidade":cidade,"uf":uf,"cep":cep,"segmento":seg, **_doc_fakes()}

# ========= datasets originais (resumo) =========
def dataset_market(n=1000):
    clientes=[_cliente() for _ in range(max(120,int(n*0.18)))]
    produtos=[produto_pt_br() for _ in range(260)]
    rows=[]
    for _ in range(n):
        cli=random.choice(clientes); prod=random.choice(produtos); d=_rand_date(365)
        quantidade=max(1,int(round(abs(random.gauss(3.0,1.4)))))
        preco_unit=round(prod["preco_base"]*_escolha_ponderada([(0.95,0.6),(1.0,1.6),(1.05,0.7)]),2)
        desconto=round(_escolha_ponderada([(0.00,3.0),(0.03,0.8),(0.05,0.6),(0.10,0.25),(0.15,0.1)]),2)
        receita=round(quantidade*preco_unit*(1-desconto),2)
        pagamento=_escolha_ponderada([("Pix",0.5),("Crédito",0.3),("Débito",0.15),("Boleto",0.05)])
        rows.append({"data":d.date(),"cliente":cli["cliente_nome"],"empresa":cli["empresa"],"uf":cli["uf"],"cidade":cli["cidade"],"segmento":cli["segmento"],
                     "sku":prod["sku"],"ean13":prod["ean13"],"produto":prod["produto"],"categoria":prod["categoria"],"marca":prod["marca"],"unidade":prod["unidade"],
                     "quantidade":quantidade,"preco_unit":preco_unit,"desconto":desconto,"receita":receita,"pagamento":pagamento})
    return {"dados":pd.DataFrame(rows), "clientes":pd.DataFrame(clientes).drop_duplicates(subset=["empresa"]).reset_index(drop=True), "produtos":pd.DataFrame(produtos)}

def dataset_financeira(n=1000):
    BANCOS=["Banco do Brasil","Caixa","Bradesco","Itaú","Santander","Sicredi","Sicoob","BTG Pactual","Inter","Nubank","Safra"]
    clientes=[_cliente() for _ in range(max(90,int(n*0.14)))]
    rows=[]
    for _ in range(n):
        cli=random.choice(clientes); emissao=_rand_date(365); prazo=_escolha_ponderada([(15,0.65),(30,1.6),(45,0.8),(60,0.5),(90,0.2)])
        venc=emissao+timedelta(days=prazo)
        valor=round(_escolha_ponderada([(120,0.5),(250,1.2),(520,1.5),(990,1.3),(1800,0.9),(3500,0.35)]),2)
        atrasodias=max(0,int(abs(random.gauss(1.8,3.8))))
        pago=random.random()<0.88; data_pag=(venc+timedelta(days=atrasodias)) if pago else None
        multa=round(0.02*valor if (data_pag and data_pag>venc) else 0,2)
        juros=round(0.00033*valor*max(0,((data_pag or datetime.now())-venc).days),2) if (data_pag or datetime.now())>venc else 0.0
        desconto=round(_escolha_ponderada([(0,3.0),(0.02*valor,0.5),(0.05*valor,0.2)]),2) if pago and random.random()<0.1 else 0.0
        liquido=round((valor+multa+juros)-desconto,2) if pago else 0.0
        rows.append({"emissao":emissao.date(),"vencimento":venc.date(),"empresa":cli["empresa"],"cnpj":cli["cnpj"],"cidade":cli["cidade"],"uf":cli["uf"],
                     "banco":random.choice(BANCOS),"nosso_numero":f"{random.randint(10_000_000_000,99_999_999_999)}",
                     "valor_face":valor,"multa":multa,"juros":juros,"desconto":desconto,"pago":pago,"data_pagamento":data_pag.date() if data_pag else None,"valor_liquido":liquido})
    return {"titulos":pd.DataFrame(rows),"sacados":pd.DataFrame(clientes).drop_duplicates(subset=["empresa"]).reset_index(drop=True)}

def dataset_logistica(n=1000):
    TRANSPORTADORAS=["Rapidão Norte","TransLog BR","ViaCargo","Azul Cargo","Correios","JadLog","Total Express","Sequoia","Loggi","Braspress","DDL Express"]
    clientes=[_cliente() for _ in range(max(80,int(n*0.12)))]
    rows=[]
    for _ in range(n):
        cli=random.choice(clientes); coleta=_rand_date(365); dias=max(1,int(abs(random.gauss(3.6,1.5))))
        prev=coleta+timedelta(days=dias)
        modal=_escolha_ponderada([("Rodoviário",2.6),("Aéreo",0.6),("Ferroviário",0.4),("Hidroviário",0.3)])
        peso=round(max(0.2, random.gauss(16,9)),2); volume=round(max(0.01, random.gauss(0.14,0.08)),3); distancia=max(10,int(abs(random.gauss(520,240))))
        base={"Rodoviário":2.1,"Aéreo":4.2,"Ferroviário":1.9,"Hidroviário":1.6}[modal]; frete=round(base*peso + 0.28*distancia + 12,2)
        entregue=random.random()<0.95; atraso=max(0,int(abs(random.gauss(0.4,1.0))))
        entrega=(prev+timedelta(days=atraso)) if entregue else None
        rows.append({"pedido":f"PED{random.randint(100000,999999)}","cliente":cli["empresa"],"origem_uf":random.choice(UFs),"destino_uf":cli["uf"],"modal":modal,
                     "coleta":coleta.date(),"previsao_entrega":prev.date(),"entrega":entrega.date() if entrega else None,"transportadora":random.choice(TRANSPORTADORAS),
                     "peso_kg":peso,"volume_m3":volume,"distancia_km":distancia,"frete":frete,"entregue":entregue})
    return {"embarques":pd.DataFrame(rows),"clientes":pd.DataFrame(clientes).drop_duplicates(subset=["empresa"]).reset_index(drop=True)}

def dataset_agro(n=1000):
    CULTURAS=["Soja","Milho","Cana-de-Açúcar","Café","Algodão","Arroz","Feijão","Trigo","Laranja","Uva"]
    INSUMOS=["Fertilizante NPK","Calcário","Herbicida","Inseticida","Fungicida","Sementes Certificadas","Adubo Orgânico","Micronutrientes","Regulador de Crescimento"]
    produtores=[]
    for _ in range(max(60,int(n*0.1))):
        if _FAKER_OK: nome=_FAKE.name(); cidade=_FAKE.city(); uf=_fake_estado_sigla()
        else: nome=f"Produtor {random.randint(1000,9999)}"; cidade=f"Cidade {random.randint(1,200)}"; uf=random.choice(UFs)
        produtores.append({"produtor":nome,"cidade":cidade,"uf":uf, **_doc_fakes()})
    talhoes=[f"T{random.randint(1,80)}" for _ in range(160)]
    items=[{"sku":f"AG-{random.randint(1000,9999)}","item":random.choice(INSUMOS),"cultura":random.choice(CULTURAS),"preco_base":round(_escolha_ponderada([(90,0.6),(120,1.0),(260,1.4),(480,0.9),(950,0.4)]),2)} for _ in range(90)]
    col=[]; ins=[]
    for _ in range(n):
        prod=random.choice(produtores); talhao=random.choice(talhoes); cultura=random.choice(CULTURAS)
        area=round(max(1.0, random.gauss(48,22)),1); plantio=_rand_date(300); colheita=plantio+timedelta(days=_escolha_ponderada([(110,0.6),(130,1.2),(150,0.9)]))
        produtividade=round(max(0.8, random.gauss(3.2,0.8)),2); producao=round(produtividade*area,2)
        preco_t=round(_escolha_ponderada([(850,0.5),(1000,1.1),(1200,1.2),(1400,0.8)]),2); receita=round(producao*preco_t,2)
        col.append({"produtor":prod["produtor"],"uf":prod["uf"],"talhao":talhao,"cultura":cultura,"area_ha":area,"plantio":plantio.date(),"colheita":colheita.date(),
                    "produtividade_t_ha":produtividade,"producao_t":producao,"preco_t":preco_t,"receita":receita})
        if random.random()<0.75:
            it=random.choice(items); qtd=max(1,int(abs(random.gauss(8,4))))
            custo=round(it["preco_base"]*qtd*_escolha_ponderada([(0.95,0.5),(1.0,1.2),(1.05,0.6)]),2)
            ins.append({"produtor":prod["produtor"],"talhao":talhao,"cultura":cultura,"item":it["item"],"sku":it["sku"],"qtd":qtd,"custo_total":custo})
    return {"colheita":pd.DataFrame(col),"insumos":pd.DataFrame(ins),"produtores":pd.DataFrame(produtores).drop_duplicates(subset=["produtor"]).reset_index(drop=True),"catalogo":pd.DataFrame(items)}

def dataset_supermercado(n=1000):
    base = dataset_market(n)
    df = base["dados"].copy()
    lojas = [f"Loja {i:02d}" for i in range(1,16)]; gondolas = [f"G{i:02d}" for i in range(1,31)]
    df["loja"]=[random.choice(lojas) for _ in range(n)]
    df["gondola"]=[random.choice(gondolas) for _ in range(n)]
    df["lote"]=[f"L{random.randint(10000,99999)}" for _ in range(n)]
    df["validade"]=[datetime.now().date()+timedelta(days=max(1,int(abs(random.gauss(35,25))))) for _ in range(n)]
    base["dados"]=df; return base

def dataset_estoque(n=1000):
    produtos=[produto_pt_br() for _ in range(240)]
    mov=[]
    for _ in range(n):
        prod=random.choice(produtos); d=_rand_date(180).date()
        tipo=_escolha_ponderada([("Entrada",0.9),("Saída",1.4)])
        qtd=max(1,int(abs(random.gauss(8,6))))
        custo_unit = round(prod["preco_base"]*_escolha_ponderada([(0.85,1.0),(0.9,1.2),(0.95,0.8)]),2)
        valor=round(qtd*custo_unit,2)
        mov.append({"data":d,"sku":prod["sku"],"ean13":prod["ean13"],"produto":prod["produto"],"categoria":prod["categoria"],"tipo":tipo,"qtd":qtd,"custo_unit":custo_unit,"valor":valor,"almox":f"AX-{random.randint(1,5)}"})
    df = pd.DataFrame(mov)
    pos = df.groupby(["sku","produto","categoria","ean13"], as_index=False).apply(
        lambda g: pd.Series({"saldo": int(g.apply(lambda r: r["qtd"]*(1 if r["tipo"]=="Entrada" else -1), axis=1).sum()),
                             "valor_mov": round(g["valor"].sum(),2)})
    ).reset_index(drop=True)
    return {"mov": df, "posicao": pos}

# ========= NOVOS DATASETS =========
def dataset_saude(n=1000):
    especialidades=["Clínico Geral","Cardiologia","Ortopedia","Dermatologia","Pediatria","Ginecologia","Oftalmologia"]
    convs=["Particular","Unimed","Amil","Bradesco Saúde","SulAmérica","Hapvida","IPASGO"]
    prof=[f"Dr(a). {_FAKE.last_name() if _FAKER_OK else random.randint(1000,9999)}" for _ in range(30)]
    consult=[]; exames=[]
    for _ in range(n):
        dt=_rand_date(365).date()
        esp=random.choice(especialidades); conv=random.choice(convs)
        pago=random.random()<0.85
        valor=round(_escolha_ponderada([(80,0.6),(120,1.2),(180,1.0),(250,0.6),(320,0.3)]),2)
        retorno = dt + timedelta(days=random.choice([7,15,30,0]))
        consult.append({
            "data":dt,"paciente":(_FAKE.name() if _FAKER_OK else f"Paciente {random.randint(1000,9999)}"),
            "cpf":f"{random.randint(100,999)}.{random.randint(100,999)}.{random.randint(100,999)}-{random.randint(10,99)}",
            "especialidade":esp,"profissional":random.choice(prof),"procedimento":"Consulta",
            "convenio":conv,"valor":valor,"pago":pago,"retorno_previsto":(retorno if random.random()<0.4 else None)
        })
        if random.random()<0.5:
            tipo=random.choice(["Hemograma","Raio-X Tórax","US Abdômen","Colesterol","Glicemia","Eletrocardiograma"])
            v=round(_escolha_ponderada([(30,0.7),(55,1.0),(90,0.8),(140,0.4)]),2)
            exames.append({"data":dt,"paciente":consult[-1]["paciente"],"tipo_exame":tipo,"resultado":"Aguardando" if random.random()<0.5 else "Normal","valor":v,"pago":random.random()<0.8})
    return {"consultas":pd.DataFrame(consult),"exames":pd.DataFrame(exames)}

def dataset_educacao(n=1000):
    turmas=[f"{random.choice(['1A','2B','3C','4D','5E'])}-{random.randint(2023,2025)}" for _ in range(20)]
    disciplinas=["Português","Matemática","História","Geografia","Ciências","Inglês","Artes","Educação Física"]
    alunos=[(_FAKE.name() if _FAKER_OK else f"Aluno {i}") for i in range(max(80,int(n*0.25)))]
    matriculas=[]; avals=[]
    for a in alunos:
        turma=random.choice(turmas)
        matriculas.append({"aluno":a,"turma":turma,"situacao":_escolha_ponderada([("Ativo",2.0),("Trancado",0.2),("Evadido",0.1)])})
    for _ in range(n):
        a=random.choice(alunos); disc=random.choice(disciplinas)
        data=_rand_date(200).date()
        nota=round(min(10,max(0,random.gauss(7.2,1.8))),1)
        freq=round(min(100,max(40,random.gauss(88,8))),1)
        avals.append({"data":data,"aluno":a,"turma":random.choice(turmas),"disciplina":disc,"avaliacao":random.choice(["P1","P2","Trabalho","Prova Final"]), "nota":nota,"frequencia_pct":freq})
    return {"matriculas":pd.DataFrame(matriculas),"avaliacoes":pd.DataFrame(avals)}

def dataset_televisao(n=1000):
    emis=["Globo","SBT","Record","Band","RedeTV!","Cultura"]
    progs=["Jornal da Noite","Novela das 9","Reality Show","Talk Show","Esporte Total","Filme"]
    cats=["Alimentos","Bebidas","Eletro","Varejo","Serviços","Automotivo","Apps"]
    aud=[]; com=[]
    for _ in range(n):
        dt=_rand_date(90)
        emissora=random.choice(emis); programa=random.choice(progs)
        dur=min(180,max(20,int(abs(random.gauss(60,25)))))
        pontos=round(max(0.2, random.gauss(8.0 if emissora=="Globo" else 3.0, 2.0)),2)
        share=round(min(60,max(1, random.gauss(24 if emissora=="Globo" else 10,6))),2)
        aud.append({"data_hora":dt,"emissora":emissora,"programa":programa,"duracao_min":dur,"audiencia_pontos":pontos,"share_pct":share})
        if random.random()<0.6:
            com.append({"data_hora":dt,"emissora":emissora,"programa":programa,"anunciante":f"{random.choice(cats)} {random.randint(1,99)}","categoria":random.choice(cats),
                        "preco_30s":round(_escolha_ponderada([(8000,0.5),(15000,0.9),(30000,0.6),(60000,0.2)]),2)})
    return {"audiencia":pd.DataFrame(aud),"comerciais":pd.DataFrame(com)}

def dataset_informatica(n=1000):
    categorias=["Acesso","Email","Impressora","Rede","Hardware","Software","Backup","Segurança"]
    prioridade=["Baixa","Média","Alta","Crítica"]
    status_list=["Aberto","Em Andamento","Aguardando Usuário","Resolvido","Cancelado"]
    usuarios=[(_FAKE.name() if _FAKER_OK else f"Usuário {i}") for i in range(200)]
    tickets=[]
    for _ in range(n):
        ab=_rand_date(180); sla=max(2,int(abs(random.gauss(16,8))))
        st=_escolha_ponderada([("Aberto",0.6),("Em Andamento",0.8),("Aguardando Usuário",0.4),("Resolvido",1.6),("Cancelado",0.1)])
        fech=None
        if st in ("Resolvido","Cancelado"):
            fech=ab+timedelta(hours=max(1,int(abs(random.gauss(sla*0.8, sla*0.4)))))
        tickets.append({
            "ticket":f"INC{random.randint(100000,999999)}","abertura":ab,"solicitante":random.choice(usuarios),
            "categoria":random.choice(categorias),"prioridade":random.choice(prioridade),"sla_h":sla,
            "fechamento":fech,"status":st,"tempo_atendimento_h":(None if fech is None else round((fech-ab).total_seconds()/3600,1)),
            "satisfacao": (None if st!="Resolvido" else random.randint(3,5))
        })
    # ativos
    marcasHW=["Dell","HP","Lenovo","Acer","Apple","Samsung","Asus"]
    ativos=[{"patrimonio":f"PAT{random.randint(10000,99999)}","tipo":random.choice(["Notebook","Desktop","Impressora","Monitor","Roteador"]),
             "marca":random.choice(marcasHW),"usuario":random.choice(usuarios),
             "aquisicao":_rand_date(1200).date(),"garantia_fim":(datetime.now().date()+timedelta(days=random.randint(30,900)))} for _ in range(max(80,int(n*0.2)))]
    return {"tickets":pd.DataFrame(tickets),"ativos":pd.DataFrame(ativos)}

def dataset_odontologia(n=800):
    procs=["Profilaxia","Restauração","Canal","Extração","Clareamento","Implante","Consulta"]
    dentistas=[f"Dr(a). {_FAKE.last_name() if _FAKER_OK else random.randint(1000,9999)}" for _ in range(18)]
    convs=["Particular","OdontoPrev","Amil Dental","Bradesco Dental","SulAmérica Odonto"]
    dentes=[f"{arc}-{num}" for arc in ["Sup","Inf"] for num in range(11,49)]
    linhas=[]
    for _ in range(n):
        dt=_rand_date(365).date()
        pac=_FAKE.name() if _FAKER_OK else f"Paciente {random.randint(1000,9999)}"
        proc=random.choice(procs); dente=random.choice(dentes) if proc in ("Restauração","Canal","Extração","Implante") else None
        valor=round(_escolha_ponderada([(120,0.8),(250,1.0),(450,0.8),(900,0.4),(1800,0.2)]),2)
        pago=random.random()<0.85
        linhas.append({"data":dt,"paciente":pac,"dentista":random.choice(dentistas),"procedimento":proc,"dente":dente,"convenio":random.choice(convs),"valor":valor,"pago":pago})
    return {"atendimentos":pd.DataFrame(linhas)}

def dataset_restaurante(n=1200):
    garcons=[f"Garçom {i:02d}" for i in range(1,25)]
    mesas=[f"M{i:02d}" for i in range(1,40)]
    cat=["Prato","Bebida","Sobremesa"]
    itens_menu={"Prato":["PF Bife","PF Frango","Lasanha","Parmegiana","Feijoada","Strogonoff"],
                "Bebida":["Refrigerante Lata","Suco 300ml","Água 500ml","Cerveja 600ml","Caipirinha"],
                "Sobremesa":["Pudim","Mousse","Petit Gateau","Sorvete 2 bolas"]}
    linhas=[]
    for _ in range(n):
        dt=_rand_date(120).date()
        mesa=random.choice(mesas); gar=random.choice(garcons)
        c=random.choice(cat); item=random.choice(itens_menu[c])
        qtd=max(1,int(abs(random.gauss(1.4,0.9))))
        preco=round(_escolha_ponderada([(8,0.3),(12,0.6),(18,1.0),(28,0.9),(39,0.5)]),2) if c!="Bebida" else round(_escolha_ponderada([(4,0.5),(7,1.0),(10,0.8),(15,0.5)]),2)
        total=round(preco*qtd,2)
        linhas.append({"data":dt,"mesa":mesa,"garcom":gar,"categoria":c,"item":item,"quantidade":qtd,"preco_unit":preco,"total":total,"pagamento":random.choice(["Pix","Crédito","Débito","Dinheiro"])})
    return {"pedidos":pd.DataFrame(linhas)}

def dataset_construcao(n=800):
    etapas=["Projeto","Fundação","Estrutura","Alvenaria","Instalações","Acabamento","Entrega"]
    obras=[f"Obra {i:03d}" for i in range(1,60)]
    clientes=[_FAKE.company() if _FAKER_OK else f"Cliente {i}" for i in range(60)]
    registros=[]; compras=[]
    for _ in range(n):
        obra=random.choice(obras); cli=random.choice(clientes); inicio=_rand_date(540).date()
        prev_fim=inicio+timedelta(days=random.randint(90,420))
        etapa=random.choice(etapas)
        prog=round(min(100,max(0,random.gauss(45,30))),1)
        orcado=round(_escolha_ponderada([(50000,0.6),(120000,1.0),(280000,0.9),(550000,0.5),(900000,0.3)]),2)
        real=round(orcado*_escolha_ponderada([(0.85,0.5),(0.95,1.2),(1.05,1.0),(1.15,0.6)]),2)
        fim=None if random.random()<0.7 else (prev_fim + timedelta(days=int(abs(random.gauss(10,20)))))
        registros.append({"obra":obra,"cliente":cli,"cidade":_FAKE.city() if _FAKER_OK else f"Cidade {random.randint(1,200)}","data_inicio":inicio,"data_prev_fim":prev_fim,
                          "data_fim":fim,"etapa":etapa,"progresso_pct":prog,"custo_orcado":orcado,"custo_real":real})
        if random.random()<0.8:
            mat=random.choice(["Cimento","Areia","Brita","Tijolo","Aço","Piso","Revestimento","Tinta","Cano PVC"])
            qtd=max(1,int(abs(random.gauss(50,40))))
            compras.append({"obra":obra,"material":mat,"unidade":random.choice(["saco","m³","kg","un","m²"]), "qtd":qtd, "custo_total":round(_escolha_ponderada([(300,0.8),(1200,1.0),(3800,0.6),(7200,0.3)]),2)})
    return {"obras":pd.DataFrame(registros),"compras":pd.DataFrame(compras)}

# ========= CAMPOS por tema =========
CAMPOS_TEMA = {
    "Market": ["data","cliente","empresa","uf","cidade","segmento","sku","ean13","produto","categoria","marca","unidade","quantidade","preco_unit","desconto","receita","pagamento"],
    "Financeira": ["emissao","vencimento","empresa","cnpj","cidade","uf","banco","nosso_numero","valor_face","multa","juros","desconto","pago","data_pagamento","valor_liquido"],
    "Logística": ["pedido","cliente","origem_uf","destino_uf","modal","coleta","previsao_entrega","entrega","transportadora","peso_kg","volume_m3","distancia_km","frete","entregue"],
    "Agro": ["produtor","uf","talhao","cultura","area_ha","plantio","colheita","produtividade_t_ha","producao_t","preco_t","receita"],
    "Supermercado": ["data","loja","gondola","lote","validade","sku","ean13","produto","categoria","marca","unidade","quantidade","preco_unit","desconto","receita","pagamento"],
    "Estoque": ["data","almox","sku","ean13","produto","categoria","tipo","qtd","custo_unit","valor"],

    # novos
    "Saúde": ["data","paciente","cpf","especialidade","profissional","procedimento","convenio","valor","pago","retorno_previsto"],
    "Educação": ["data","aluno","turma","disciplina","avaliacao","nota","frequencia_pct"],
    "Televisão": ["data_hora","emissora","programa","duracao_min","audiencia_pontos","share_pct"],
    "Informática": ["ticket","abertura","solicitante","categoria","prioridade","sla_h","fechamento","status","tempo_atendimento_h","satisfacao"],
    "Odontologia": ["data","paciente","dentista","procedimento","dente","convenio","valor","pago"],
    "Restaurante": ["data","mesa","garcom","categoria","item","quantidade","preco_unit","total","pagamento"],
    "Construção": ["obra","cliente","cidade","data_inicio","data_prev_fim","data_fim","etapa","progresso_pct","custo_orcado","custo_real"],
}

PERFIL_IDX = {
    "Market":{"basico":["data","empresa","produto","categoria","unidade","quantidade","preco_unit","receita"],"completo":CAMPOS_TEMA["Market"]},
    "Financeira":{"basico":["emissao","vencimento","empresa","valor_face","pago","valor_liquido"],"completo":CAMPOS_TEMA["Financeira"]},
    "Logística":{"basico":["pedido","cliente","destino_uf","modal","coleta","previsao_entrega","frete","entregue"],"completo":CAMPOS_TEMA["Logística"]},
    "Agro":{"basico":["produtor","cultura","area_ha","plantio","colheita","producao_t","receita"],"completo":CAMPOS_TEMA["Agro"]},
    "Supermercado":{"basico":["data","loja","produto","categoria","quantidade","preco_unit","receita","validade"],"completo":CAMPOS_TEMA["Supermercado"]},
    "Estoque":{"basico":["data","almox","sku","produto","tipo","qtd","valor"],"completo":CAMPOS_TEMA["Estoque"]},

    "Saúde":{"basico":["data","paciente","especialidade","procedimento","valor","pago"],"completo":CAMPOS_TEMA["Saúde"]},
    "Educação":{"basico":["data","aluno","disciplina","avaliacao","nota"],"completo":CAMPOS_TEMA["Educação"]},
    "Televisão":{"basico":["data_hora","emissora","programa","audiencia_pontos","share_pct"],"completo":CAMPOS_TEMA["Televisão"]},
    "Informática":{"basico":["ticket","abertura","categoria","prioridade","status","tempo_atendimento_h"],"completo":CAMPOS_TEMA["Informática"]},
    "Odontologia":{"basico":["data","paciente","procedimento","valor","pago"],"completo":CAMPOS_TEMA["Odontologia"]},
    "Restaurante":{"basico":["data","mesa","item","quantidade","preco_unit","total"],"completo":CAMPOS_TEMA["Restaurante"]},
    "Construção":{"basico":["obra","etapa","progresso_pct","custo_real"],"completo":CAMPOS_TEMA["Construção"]},
}

def _col_def(name: str) -> Dict[str, Any]:
    if name in ("data","emissao","vencimento","coleta","previsao_entrega","entrega","plantio","colheita","data_pagamento","validade","abertura","fechamento","data_hora","data_inicio","data_prev_fim","data_fim","retorno_previsto"): return {"name":name,"fmt":"date","width":14}
    if name in ("quantidade","qtd","sla_h","duracao_min","satisfacao"): return {"name":name,"fmt":"int","width":12}
    if name in ("preco_unit","valor_face","multa","juros","desconto","valor_liquido","frete","preco_t","receita","custo_total","total","preco_kg","cambio","valor","valor_beneficios","salario","descontos","liquido","preco_unit_moeda","total_moeda","total_brl","preco_30s","custo_orcado","custo_real"): return {"name":name,"fmt":"currency","width":13}
    if name in ("peso_kg","volume_m3","area_ha","produtividade_t_ha","producao_t","mao_obra_horas","tempo_atendimento_h","nota","frequencia_pct","audiencia_pontos","share_pct","progresso_pct"): return {"name":name,"fmt":"float","width":13}
    return {"name":name,"fmt":"text","width":max(10,min(26,len(name)+6))}

# ========= builders de planilha por tema =========
def build_spec_from_bundle(tema: str, bundle: Dict[str,pd.DataFrame], campos: List[str]) -> Dict[str,Any]:
    base={"workbook":{"title":f"Relatório {tema}","author":"Gerador Interativo","created_at":datetime.now()},
          "dashboard_name":"Dashboard"}

    # ----- temas originais -----
    if tema=="Market":
        df=bundle["dados"][campos].copy()
        cli=bundle["clientes"][["empresa","cnpj","cidade","uf","segmento"]]
        prod=bundle["produtos"][["sku","ean13","produto","categoria","marca","unidade","preco_base"]]
        return {**base,"sheets":[
            {"name":"Vendas","data":df,"columns":[_col_def(c) for c in df.columns],"freeze":"B2","autofilter":True},
            {"name":"Clientes","data":cli,"columns":[_col_def(c) for c in cli.columns],"freeze":"A2","autofilter":True},
            {"name":"Produtos","data":prod,"columns":[_col_def(c) for c in prod.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Receita Total","value":float(bundle["dados"]["receita"].sum()),"fmt":"currency"},
            {"label":"Itens Vendidos","value":int(bundle["dados"]["quantidade"].sum()),"fmt":"int"},
            {"label":"Ticket Médio","value":float(bundle["dados"]["receita"].sum()/max(1,bundle["dados"]["quantidade"].sum())),"fmt":"float"},
        ],"pivots":[
            {"name":"Receita por Categoria","data_sheet":"Vendas","index":["categoria"],"columns":[],"values":{"receita":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Receita por Categoria","y_title":"R$"}},
        ]}

    if tema=="Financeira":
        df=bundle["titulos"][campos].copy(); sac=bundle["sacados"][["empresa","cnpj","cidade","uf","segmento"]]
        return {**base,"sheets":[
            {"name":"Títulos","data":df,"columns":[_col_def(c) for c in df.columns],"freeze":"A2","autofilter":True},
            {"name":"Sacados","data":sac,"columns":[_col_def(c) for c in sac.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Carteira (face)","value":float(bundle["titulos"]["valor_face"].sum()),"fmt":"currency"},
            {"label":"Recebido (líquido)","value":float(bundle["titulos"]["valor_liquido"].sum()),"fmt":"currency"},
            {"label":"% Pago","value":float(bundle["titulos"]["pago"].mean()*100),"fmt":"float"},
        ],"pivots":[
            {"name":"Carteira por UF","data_sheet":"Títulos","index":["uf"],"columns":[],"values":{"valor_face":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Carteira por UF","y_title":"R$"}},
        ]}

    if tema=="Logística":
        df=bundle["embarques"][campos].copy()
        sheets=[{"name":"Embarques","data":df,"columns":[_col_def(c) for c in df.columns],"freeze":"B2","autofilter":True}]
        if "clientes" in bundle:
            cli=bundle["clientes"][["empresa","cnpj","cidade","uf"]]
            sheets.append({"name":"Clientes","data":cli,"columns":[_col_def(c) for c in cli.columns],"freeze":"A2","autofilter":True})
        return {**base,"sheets":sheets,"kpis":[
            {"label":"Frete Total","value":float(bundle["embarques"]["frete"].sum()),"fmt":"currency"},
            {"label":"Peso Total (kg)","value":float(bundle["embarques"]["peso_kg"].sum()),"fmt":"float"},
            {"label":"% Entregue","value":float(bundle["embarques"]["entregue"].mean()*100),"fmt":"float"},
        ],"pivots":[
            {"name":"Frete por Modal","data_sheet":"Embarques","index":["modal"],"columns":[],"values":{"frete":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Frete por Modal","y_title":"R$"}},
        ]}

    if tema=="Agro":
        df=bundle["colheita"][campos].copy(); ins=bundle["insumos"][["produtor","talhao","cultura","item","sku","qtd","custo_total"]]
        prods=bundle["produtores"][["produtor","cnpj","cpf","cidade","uf"]]; cat=bundle["catalogo"][["sku","item","cultura","preco_base"]]
        return {**base,"sheets":[
            {"name":"Colheita","data":df,"columns":[_col_def(c) for c in df.columns],"freeze":"A2","autofilter":True},
            {"name":"Insumos","data":ins,"columns":[_col_def(c) for c in ins.columns],"freeze":"A2","autofilter":True},
            {"name":"Produtores","data":prods,"columns":[_col_def(c) for c in prods.columns],"freeze":"A2","autofilter":True},
            {"name":"Catálogo","data":cat,"columns":[_col_def(c) for c in cat.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Receita Total","value":float(bundle["colheita"]["receita"].sum()),"fmt":"currency"},
            {"label":"Área Total (ha)","value":float(bundle["colheita"]["area_ha"].sum()),"fmt":"float"},
            {"label":"Produtividade Média (t/ha)","value":float(bundle["colheita"]["produtividade_t_ha"].mean()),"fmt":"float"},
        ],"pivots":[
            {"name":"Receita por Cultura","data_sheet":"Colheita","index":["cultura"],"columns":[],"values":{"receita":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Receita por Cultura","y_title":"R$"}},
        ]}

    if tema=="Supermercado":
        df=bundle["dados"][campos].copy()
        return {**base,"sheets":[
            {"name":"Vendas Super","data":df,"columns":[_col_def(c) for c in df.columns],"freeze":"B2","autofilter":True},
        ],"kpis":[
            {"label":"Receita (Super)","value":float(df.get('receita',pd.Series(dtype=float)).sum()),"fmt":"currency"},
            {"label":"Itens","value":int(df.get('quantidade',pd.Series(dtype=float)).sum()),"fmt":"int"},
        ],"pivots":[
            {"name":"Itens por Loja","data_sheet":"Vendas Super","index":["loja"],"columns":[],"values":{"quantidade":"sum"},"fill_value":0,"round":0,"chart":{"type":"column","title":"Itens por Loja","y_title":"Unid"}},
        ]}

    if tema=="Estoque":
        mov=bundle["mov"][campos].copy(); pos=bundle["posicao"]
        return {**base,"sheets":[
            {"name":"Movimentações","data":mov,"columns":[_col_def(c) for c in mov.columns],"freeze":"A2","autofilter":True},
            {"name":"Posição","data":pos,"columns":[_col_def(c) for c in pos.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Saldo Total (itens)","value":int(pos["saldo"].sum()),"fmt":"int"},
            {"label":"Valor Movimentado","value":float(mov["valor"].sum()),"fmt":"currency"},
        ],"pivots":[
            {"name":"Saldo por Categoria","data_sheet":"Posição","index":["categoria"],"columns":[],"values":{"saldo":"sum"},"fill_value":0,"round":0,"chart":{"type":"column","title":"Saldo por Categoria","y_title":"Unid"}},
        ]}

    # ----- novos temas -----
    if tema=="Saúde":
        cons=bundle["consultas"][campos].copy()
        exams=bundle["exames"] if "exames" in bundle else pd.DataFrame()
        sheets=[{"name":"Consultas","data":cons,"columns":[_col_def(c) for c in cons.columns],"freeze":"A2","autofilter":True}]
        if not exams.empty:
            sheets.append({"name":"Exames","data":exams,"columns":[_col_def(c) for c in exams.columns],"freeze":"A2","autofilter":True})
        return {**base,"sheets":sheets,"kpis":[
            {"label":"Faturamento Consultas","value":float(cons.get("valor",pd.Series(dtype=float)).sum()),"fmt":"currency"},
            {"label":"% Pago","value":float(cons.get("pago",pd.Series(dtype=bool)).mean()*100 if "pago" in cons else 0.0),"fmt":"float"},
        ],"pivots":[
            {"name":"Valor por Especialidade","data_sheet":"Consultas","index":["especialidade"],"columns":[],"values":{"valor":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Valor por Especialidade","y_title":"R$"}},
        ]}

    if tema=="Educação":
        aval=bundle["avaliacoes"][campos].copy()
        mats=bundle["matriculas"]
        return {**base,"sheets":[
            {"name":"Avaliações","data":aval,"columns":[_col_def(c) for c in aval.columns],"freeze":"A2","autofilter":True},
            {"name":"Matrículas","data":mats,"columns":[_col_def(c) for c in mats.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Média Geral","value":float(aval["nota"].mean()),"fmt":"float"},
            {"label":"Presença Média (%)","value":float(aval["frequencia_pct"].mean()),"fmt":"float"},
        ],"pivots":[
            {"name":"Média por Disciplina","data_sheet":"Avaliações","index":["disciplina"],"columns":[],"values":{"nota":"mean"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Média por Disciplina","y_title":"Nota"}},
        ]}

    if tema=="Televisão":
        aud=bundle["audiencia"][campos].copy(); com=bundle["comerciais"]
        return {**base,"sheets":[
            {"name":"Audiência","data":aud,"columns":[_col_def(c) for c in aud.columns],"freeze":"A2","autofilter":True},
            {"name":"Comerciais","data":com,"columns":[_col_def(c) for c in com.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Pontos Médios","value":float(aud["audiencia_pontos"].mean()),"fmt":"float"},
            {"label":"Share Médio (%)","value":float(aud["share_pct"].mean()),"fmt":"float"},
        ],"pivots":[
            {"name":"Audiência por Emissora","data_sheet":"Audiência","index":["emissora"],"columns":[],"values":{"audiencia_pontos":"mean"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Pontos médios por Emissora","y_title":"Pontos"}},
        ]}

    if tema=="Informática":
        tk=bundle["tickets"][campos].copy(); at=bundle["ativos"]
        return {**base,"sheets":[
            {"name":"Tickets","data":tk,"columns":[_col_def(c) for c in tk.columns],"freeze":"B2","autofilter":True},
            {"name":"Ativos","data":at,"columns":[_col_def(c) for c in at.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"% Resolvidos","value":float(tk["status"].eq("Resolvido").mean()*100),"fmt":"float"},
            {"label":"Satisfação Média","value":float(pd.to_numeric(tk["satisfacao"], errors="coerce").mean()),"fmt":"float"},
            {"label":"TMA (h)","value":float(pd.to_numeric(tk["tempo_atendimento_h"], errors="coerce").mean()),"fmt":"float"},
        ],"pivots":[
            {"name":"Tickets por Categoria","data_sheet":"Tickets","index":["categoria"],"columns":[],"values":{"ticket":"count"},"fill_value":0,"round":0,"chart":{"type":"column","title":"Tickets por Categoria","y_title":"Qtde"}},
        ]}

    if tema=="Odontologia":
        at=bundle["atendimentos"][campos].copy()
        return {**base,"sheets":[
            {"name":"Atendimentos","data":at,"columns":[_col_def(c) for c in at.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Faturamento Odonto","value":float(at["valor"].sum()),"fmt":"currency"},
            {"label":"% Pago","value":float(at["pago"].mean()*100),"fmt":"float"},
        ],"pivots":[
            {"name":"Valor por Procedimento","data_sheet":"Atendimentos","index":["procedimento"],"columns":[],"values":{"valor":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Valor por Procedimento","y_title":"R$"}},
        ]}

    if tema=="Restaurante":
        pdv=bundle["pedidos"][campos].copy()
        return {**base,"sheets":[
            {"name":"Pedidos","data":pdv,"columns":[_col_def(c) for c in pdv.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Faturamento","value":float(pdv["total"].sum()),"fmt":"currency"},
            {"label":"Ticket Médio","value":float(pdv["total"].mean()),"fmt":"currency"},
            {"label":"Itens Vendidos","value":int(pdv["quantidade"].sum()),"fmt":"int"},
        ],"pivots":[
            {"name":"Vendas por Categoria","data_sheet":"Pedidos","index":["categoria"],"columns":[],"values":{"total":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Vendas por Categoria","y_title":"R$"}},
        ]}

    if tema=="Construção":
        ob=bundle["obras"][campos].copy(); comp=bundle["compras"]
        return {**base,"sheets":[
            {"name":"Obras","data":ob,"columns":[_col_def(c) for c in ob.columns],"freeze":"A2","autofilter":True},
            {"name":"Compras","data":comp,"columns":[_col_def(c) for c in comp.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Desvio Orçamentário (R$)","value":float((ob["custo_real"]-ob["custo_orcado"]).sum()),"fmt":"currency"},
            {"label":"% Conclusão Média","value":float(ob["progresso_pct"].mean()),"fmt":"float"},
        ],"pivots":[
            {"name":"Custo por Etapa","data_sheet":"Obras","index":["etapa"],"columns":[],"values":{"custo_real":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Custo por Etapa","y_title":"R$"}},
        ]}

    raise ValueError("Tema não suportado")

# ========= API =========
_TEMAS = {
    "Market": dataset_market,
    "Financeira": dataset_financeira,
    "Logística": dataset_logistica,
    "Agro": dataset_agro,
    "Supermercado": dataset_supermercado,
    "Estoque": dataset_estoque,
    # novos
    "Saúde": dataset_saude,
    "Educação": dataset_educacao,
    "Televisão": dataset_televisao,
    "Informática": dataset_informatica,
    "Odontologia": dataset_odontologia,
    "Restaurante": dataset_restaurante,
    "Construção": dataset_construcao,
}

ALIASES = {
    "market":"Market","mercado":"Market",
    "financeira":"Financeira",
    "logistica":"Logística","logística":"Logística",
    "agro":"Agro",
    "supermercado":"Supermercado","super-mercado":"Supermercado",
    "estoque":"Estoque",

    "saude":"Saúde","saúde":"Saúde","clinica":"Saúde","clínica":"Saúde",
    "educacao":"Educação","educação":"Educação","escola":"Educação",
    "televisao":"Televisão","televisão":"Televisão","tv":"Televisão",
    "informatica":"Informática","informática":"Informática","helpdesk":"Informática","ti":"Informática",
    "odontologia":"Odontologia","odonto":"Odontologia",
    "restaurante":"Restaurante","bar":"Restaurante","lanchonete":"Restaurante",
    "construcao":"Construção","construção":"Construção","obra":"Construção",
}

PERFIS = ["Básico","Completo","Personalizado"]

def listar_temas()->List[str]: return list(_TEMAS.keys())

def gerar_excel_tema(tema: str, n_linhas: int, campos: List[str], output_path: str, estilo="Azul") -> str:
    if tema not in _TEMAS: raise ValueError(f"Tema inválido. Opções: {listar_temas()}")
    bundle=_TEMAS[tema](n_linhas)
    spec=build_spec_from_bundle(tema, bundle, campos)
    gerar_planilha(spec, output_path, estilo_key=estilo)
    return output_path

# ========= seleção / CLI =========
def normaliza_tema(v: str)->str:
    key=v.strip().lower()
    if key in ALIASES: return ALIASES[key]
    for k in _TEMAS.keys():
        if key==k.lower(): return k
    raise ValueError(f"Tema inválido: {v}. Opções: {', '.join(listar_temas())}")

def resolve_campos_por_perfil(tema: str, perfil: str, expr: Optional[str]=None)->List[str]:
    if perfil.lower().startswith("b"): return list(PERFIL_IDX[tema]["basico"])
    if perfil.lower().startswith("c"): return list(PERFIL_IDX[tema]["completo"])
    todos=CAMPOS_TEMA[tema]
    if expr is None:
        print("\nSelecione campos (números, vírgulas, intervalos com '-').")
        for i,c in enumerate(todos,1): print(f"  {i}) {c}")
        expr=input("Ex.: 1-5,8,10 (ENTER=todos): ").strip()
    idxs=parse_ranges_to_indices(expr, len(todos))
    if not idxs: idxs=list(range(len(todos)))
    return [todos[i] for i in idxs]

def modo_interativo():
    temas=listar_temas(); estilos=list(ESTILOS.keys())
    tema=temas[prompt_menu("Tema", temas, 0)]
    linhas=prompt_int("Quantidade de linhas", 1000, 1)
    estilo=estilos[prompt_menu("Estilo", estilos, 0)]
    perfil=PERFIS[prompt_menu("Perfil de saída", PERFIS, 0)]
    campos = resolve_campos_por_perfil(tema, perfil)
    saida=input("Arquivo de saída (padrão: saida.xlsx): ").strip() or "saida.xlsx"
    print("\nGerando...")
    caminho=gerar_excel_tema(tema, linhas, campos, saida, estilo=estilo)
    print(f"✅ Planilha gerada: {caminho}")

def modo_argparse():
    import argparse
    p=argparse.ArgumentParser(description="Gerador XLSX multi-temas (PT-BR), com estilos e campos personalizáveis")
    p.add_argument("--tema", default="Market")
    p.add_argument("--linhas", type=int, default=1000)
    p.add_argument("--saida", default="saida.xlsx")
    p.add_argument("--perfil", default="basico", choices=["basico","completo","personalizado","básico","completo","personalizado"])
    p.add_argument("--campos", default=None, help="Para perfil personalizado. Ex.: '1-5,8,10'")
    p.add_argument("--estilo", default="Azul", choices=list(ESTILOS.keys()))
    p.add_argument("--nao_interativo", action="store_true")
    args=p.parse_args()

    tema=normaliza_tema(args.tema)
    perfil="Básico" if args.perfil.startswith("b") else "Completo" if args.perfil.startswith("c") else "Personalizado"
    campos = resolve_campos_por_perfil(tema, perfil, expr=args.campos if (args.nao_interativo or perfil=="Personalizado") else None)
    caminho=gerar_excel_tema(tema, args.linhas, campos, args.saida, estilo=args.estilo)
    print(f"✅ Planilha gerada: {caminho}")

if __name__=="__main__":
    if len(sys.argv)==1: modo_interativo()
    else: modo_argparse()
