"""
Gerador_Planilhas.py — multi-temas
Temas:
Market, Financeira, Logística, Agro, Supermercado, RH, Telemarketing, Oficina,
Hortifruti, Exportação, Estoque (e alias Mercado).
"""

import sys, random
from typing import Dict, Any, List, Tuple, Optional
from datetime import datetime, timedelta
import numpy as np
import pandas as pd

# ================== utils de prompt ==================
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

# ================== estilos ==================
ESTILOS = {
    "Azul":   {"header_bg":"#E8F1FF","kpi_bg":"#DCEBFF","zebra":"#F7FAFF","neg":"#FCE8E6","pos":"#E6F4EA","scale_min":"#E8F1FF","scale_max":"#2B6CB0"},
    "Verde":  {"header_bg":"#E9F7EF","kpi_bg":"#D4EFDF","zebra":"#F5FBF7","neg":"#FDEDEC","pos":"#E8F8F5","scale_min":"#E9F7EF","scale_max":"#1E8449"},
    "Cinza":  {"header_bg":"#F0F0F0","kpi_bg":"#E6E6E6","zebra":"#FAFAFA","neg":"#FDEDEC","pos":"#EBF5FB","scale_min":"#F0F0F0","scale_max":"#5D6D7E"},
    "Laranja":{"header_bg":"#FFF1E6","kpi_bg":"#FFE0CC","zebra":"#FFF9F3","neg":"#FDECEA","pos":"#FFF7E6","scale_min":"#FFF1E6","scale_max":"#D35400"},
}

# ================== motor xlsx ==================
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

        # abas de dados
        for sh in sheets_spec:
            name = sh['name']; data = sh.get('data', pd.DataFrame())
            if isinstance(data, list): data = pd.DataFrame(data)
            df = data.copy(); name_to_df[name]=df
            df.to_excel(writer, sheet_name=name, index=False, startrow=1)
            ws = writer.sheets[name]

            # header
            for i, colname in enumerate(df.columns):
                ws.write(0, i, colname, fmt["header"])

            # larguras/formatos
            for col in sh.get('columns', []):
                colname=col.get('name'); width=col.get('width',15); f=col.get('fmt','text')
                if colname in df.columns:
                    ci=df.columns.get_loc(colname)
                    base = fmt["text"] if f=="text" else fmt.get(f, fmt["text"])
                    ws.set_column(ci, ci, width, base)

            if sh.get('autofilter', True) and not df.empty:
                ws.autofilter(0,0, df.shape[0], df.shape[1]-1)
            if sh.get('freeze'): ws.freeze_panes(*_excel_cell_to_tuple(sh['freeze']))

            # zebra
            if not df.empty:
                ws.conditional_format(1,0, df.shape[0], df.shape[1]-1, {
                    'type':'formula','criteria':'=MOD(ROW(),2)=0','format':workbook.add_format({'bg_color':pal["zebra"]})
                })
            # barras de dados
            numeric_cols = [c for c in ["quantidade","receita","valor_face","valor_liquido","frete","peso_kg","distancia_km","custo_total","producao_t","salario","duracao_s","mao_obra_horas","qtd","valor_total","valor"] if c in df.columns]
            for c in numeric_cols:
                ci=df.columns.get_loc(c)
                ws.conditional_format(1,ci, df.shape[0], ci, {'type':'data_bar'})

            # destaques contexto
            if {"vencimento","pago"}.issubset(set(df.columns)):
                ci_v=df.columns.get_loc("vencimento"); ci_p=df.columns.get_loc("pago")
                ws.conditional_format(1,0, df.shape[0], df.shape[1]-1, {
                    'type':'formula',
                    'criteria': f'=AND(TODAY()>INDIRECT(ADDRESS(ROW(),{ci_v+1})),INDIRECT(ADDRESS(ROW(),{ci_p+1}))=FALSE())',
                    'format':workbook.add_format({'bg_color':pal["neg"]})
                })
            if {"previsao_entrega","entrega"}.issubset(set(df.columns)):
                ci_prev=df.columns.get_loc("previsao_entrega"); ci_ent=df.columns.get_loc("entrega")
                ws.conditional_format(1,0, df.shape[0], df.shape[1]-1, {
                    'type':'formula',
                    'criteria': f'=AND(NOT(ISBLANK(INDIRECT(ADDRESS(ROW(),{ci_ent+1})))),INDIRECT(ADDRESS(ROW(),{ci_ent+1}))>INDIRECT(ADDRESS(ROW(),{ci_prev+1})))',
                    'format':workbook.add_format({'bg_color':pal["neg"]})
                })
            if "validade" in df.columns:
                ci_val = df.columns.get_loc("validade")
                ws.conditional_format(1,0, df.shape[0], df.shape[1]-1, {
                    'type':'formula',
                    'criteria': f'=INDIRECT(ADDRESS(ROW(),{ci_val+1}))<=TODAY()+7',
                    'format':workbook.add_format({'bg_color':pal["neg"]})
                })

        # dashboard KPIs
        if kpis_spec:
            if dashboard_name not in writer.sheets:
                pd.DataFrame().to_excel(writer, sheet_name=dashboard_name, index=False)
            ws = writer.sheets[dashboard_name]
            ws.write(0,0,"KPIs", fmt["header"])
            r=2
            for k in kpis_spec:
                ws.write(r,0,k.get("label","KPI"), fmt["kpi_lbl"])
                if "formula" in k:
                    ws.write_formula(r,1,k["formula"], fmt["kpi_val"])
                else:
                    val=k.get("value",""); f=k.get("fmt","text")
                    cellfmt = fmt["kpi_val"] if f in ("float","int","currency") else fmt["text"]
                    ws.write(r,1,val, cellfmt)
                r+=1

        # pivôs + gráficos
        for pv in pivots_spec:
            name=pv['name']; src_sheet=pv['data_sheet']
            if src_sheet not in name_to_df: continue
            src=name_to_df[src_sheet]
            if src.empty:
                pd.DataFrame().to_excel(writer, sheet_name=name, index=False); continue
            pvt=pd.pivot_table(src, index=pv.get('index',[]), columns=pv.get('columns',[]),
                               values=list(pv.get('values', {'valor':'sum'}).keys()),
                               aggfunc=pv.get('values', {'valor':'sum'}), fill_value=pv.get('fill_value',0))
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

# ================== dados realistas (faker) ==================
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

# ---------- EAN-13 ----------
def _ean13_checksum(num12: str) -> int:
    s = sum((3 if i%2 else 1)*int(d) for i,d in enumerate(num12[::-1]))
    return (10 - (s % 10)) % 10
def gerar_ean13(prefix: str="789") -> str:
    base_len = 12 - len(prefix)
    middle = "".join(str(random.randint(0,9)) for _ in range(base_len))
    num12 = prefix + middle
    dv = _ean13_checksum(num12)
    return num12 + str(dv)

# ---------- Unidades & preços ----------
_UNIDADES = [
    ("g", 140, 0.5), ("g", 200, 0.7), ("g", 500, 0.9), ("g", 1000, 1.0),
    ("ml", 300, 0.8), ("ml", 500, 1.0), ("L", 1, 1.2), ("L", 2, 1.9),
    ("un", 1, 1.0), ("un", 4, 3.6), ("un", 6, 5.2), ("un", 12, 10.0)
]
# Famílias PT-BR com faixas de preço
_BASE_PRECO_PT = {
    "Mercearia": (5.90, 29.90),
    "Bebidas": (4.90, 39.90),
    "Higiene & Beleza": (7.90, 49.90),
    "Limpeza": (5.90, 29.90),
    "Frios & Laticínios": (7.90, 59.90),
    "Açougue": (14.90, 79.90),
    "Padaria & Confeitaria": (4.90, 24.90),
    "Pets": (9.90, 49.90),
    "Utilidades": (9.90, 69.90),
}

# Catálogo PT-BR (bases de nome)
CAT_PT = {
    "Mercearia": ["Arroz", "Feijão Carioca", "Feijão Preto", "Macarrão Spaghetti", "Macarrão Parafuso", "Molho de Tomate", "Azeite de Oliva", "Açúcar Refinado", "Farinha de Trigo", "Café Torrado e Moído", "Atum em Óleo", "Sardinha em Óleo", "Azeitona Verde", "Milho Verde", "Ervilha em Conserva", "Biscoito Recheado Chocolate", "Biscoito Cream Cracker", "Achocolatado em Pó", "Granola Tradicional", "Aveia em Flocos", "Leite em Pó", "Leite Condensado", "Creme de Leite"],
    "Bebidas": ["Água Mineral", "Refrigerante Cola", "Refrigerante Guaraná", "Suco de Uva", "Suco de Laranja", "Cerveja Pilsen", "Cerveja IPA", "Vinho Tinto Seco", "Chá Gelado", "Água de Coco", "Energético"],
    "Higiene & Beleza": ["Sabonete", "Shampoo", "Condicionador", "Desodorante Aerosol", "Creme Dental", "Escova Dental", "Fio Dental", "Enxaguante Bucal", "Lenço Umedecido"],
    "Limpeza": ["Detergente Líquido Neutro", "Desinfetante", "Amaciante de Roupas", "Sabão em Pó", "Limpador Multiuso", "Água Sanitária", "Esponja Multiuso", "Lustra Móveis", "Saco de Lixo"],
    "Frios & Laticínios": ["Queijo Mussarela Fatiado", "Queijo Prato", "Queijo Parmesão Ralado", "Presunto Cozido", "Peito de Peru", "Iogurte Natural", "Iogurte Grego", "Requeijão Cremoso", "Manteiga com Sal", "Ricota Fresca", "Cream Cheese"],
    "Açougue": ["Frango Congelado", "Coxa e Sobrecoxa", "Peito de Frango", "Carne Moída", "Bife de Alcatra", "Carne Suína em Cubos", "Linguiça Toscana", "Carne para Panela"],
    "Padaria & Confeitaria": ["Pão Francês", "Pão de Forma", "Bolo de Chocolate", "Bolo de Cenoura", "Pão de Queijo", "Croissant", "Biscoito Amanteigado"],
    "Pets": ["Ração Cães Adultos", "Ração Gatos Adultos", "Petisco para Cães", "Areia Sanitária"],
    "Utilidades": ["Pano de Prato", "Esponja de Aço", "Vassoura", "Rodo", "Balde Plástico"],
}
ADJETIVOS = ["Premium", "Tradicional", "Integral", "Zero Açúcar", "Zero Lactose", "Light", "Orgânico", "Clássico", "Caseiro", "Intenso", "Extra Forte", "Sabor Chocolate", "Sabor Morango", "Sabor Baunilha"]

def _preco_realista_pt(familia: str, unidade: Tuple[str, float, float]) -> float:
    low, high = _BASE_PRECO_PT.get(familia, (7.90, 49.90))
    base = random.uniform(low, high)
    _, _, fator = unidade
    brand_bump = random.choice([0.95, 1.0, 1.05, 1.1])
    preco = base * fator * brand_bump
    cents = random.choice([0.90, 0.99, 0.79, 0.49, 0.19])
    return float(int(preco)) + cents

# ---------- Clientes ----------
def _cliente():
    if _FAKER_OK:
        nome=_FAKE.name(); empresa=_FAKE.company(); cidade=_FAKE.city(); uf=_fake_estado_sigla(); cep=_FAKE.postcode()
    else:
        nome=f"Cliente {random.randint(1000,9999)}"; empresa=f"Empresa {random.randint(100,999)} Ltda"; cidade=f"Cidade {random.randint(1,200)}"; uf=random.choice(UFs); cep=f"{random.randint(10000,99999)}-{random.randint(100,999)}"
    seg=_escolha_ponderada([("Varejo",0.5),("Atacado",0.3),("E-commerce",0.2)])
    return {"cliente_nome":nome,"empresa":empresa,"cidade":cidade,"uf":uf,"cep":cep,"segmento":seg, **_doc_fakes()}

# ---------- Produtos PT-BR ----------
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

# (Opcional) AUTO: usa faker_commerce se disponível, traduz categorias e PT-ifica nomes quando vierem em EN
TRAD_CATEG = {
    "Food & Beverage":"Alimentos & Bebidas","Grocery":"Mercearia","Health & Beauty":"Higiene & Beleza",
    "Home & Garden":"Casa & Jardim","Sports & Outdoors":"Esporte & Lazer","Electronics":"Eletrônicos",
    "Kids & Baby":"Infantil","Office Supplies":"Papelaria","Pet Supplies":"Pets","Toys & Games":"Brinquedos & Jogos"
}
def produto_auto_pt() -> Dict[str, Any]:
    if _COMMERCE:
        cat_en = _FAKE.ecommerce_category()
        nome_en = _FAKE.ecommerce_name()
        familia = TRAD_CATEG.get(cat_en, cat_en)  # tenta traduzir
        # "pt-ificar" nome: adiciona unidade e remove termos muito técnicos
        unidade = random.choice(_UNIDADES)
        unidade_str = f"{unidade[1]}{unidade[0]}"
        nome = f"{nome_en} {unidade_str}"
        marca = random.choice(MARCAS_BR + ["Genérico","Local","Premium"])
        ean = gerar_ean13("789")
        sku = f"{familia[:2].upper()}-{random.randint(10000,99999)}"
        preco_base = round(_preco_realista_pt(familia if familia in _BASE_PRECO_PT else "Mercearia", unidade),2)
        return {"sku": sku, "ean13": ean, "produto": nome, "categoria": familia, "marca": marca, "unidade": unidade_str, "preco_base": preco_base}
    # fallback
    return produto_pt_br()

# ================== DATASETS DOS TEMAS ==================
def dataset_market(n=1000, idioma="pt"):
    clientes=[_cliente() for _ in range(max(120,int(n*0.18)))]
    # produtos em PT
    prod_fn = produto_pt_br if idioma=="pt" else produto_auto_pt
    produtos=[prod_fn() for _ in range(260)]
    rows=[]
    for _ in range(n):
        cli=random.choice(clientes); prod=random.choice(produtos); d=_rand_date(365)
        quantidade=max(1,int(round(abs(random.gauss(3.0,1.4)))))
        preco_unit=round(prod["preco_base"]*_escolha_ponderada([(0.95,0.6),(1.0,1.6),(1.05,0.7)]),2)
        desconto=round(_escolha_ponderada([(0.00,3.0),(0.03,0.8),(0.05,0.6),(0.10,0.25),(0.15,0.1)]),2)
        receita=round(quantidade*preco_unit*(1-desconto),2)
        pagamento=_escolha_ponderada([("Pix",0.5),("Crédito",0.3),("Débito",0.15),("Boleto",0.05)])
        rows.append({
            "data":d.date(),"cliente":cli["cliente_nome"],"empresa":cli["empresa"],"uf":cli["uf"],"cidade":cli["cidade"],"segmento":cli["segmento"],
            "sku":prod["sku"],"ean13":prod["ean13"],"produto":prod["produto"],"categoria":prod["categoria"],"marca":prod["marca"],"unidade":prod["unidade"],
            "quantidade":quantidade,"preco_unit":preco_unit,"desconto":desconto,"receita":receita,"pagamento":pagamento
        })
    return {"dados":pd.DataFrame(rows), "clientes":pd.DataFrame(clientes).drop_duplicates(subset=["empresa"]).reset_index(drop=True), "produtos":pd.DataFrame(produtos)}

def dataset_financeira(n=1000):
    BANCOS = ["Banco do Brasil","Caixa","Bradesco","Itaú","Santander","Sicredi","Sicoob","BTG Pactual","Inter","Nubank","Safra"]
    clientes=[_cliente() for _ in range(max(90,int(n*0.14)))]
    rows=[]
    for _ in range(n):
        cli=random.choice(clientes); emissao=_rand_date(365)
        prazo=_escolha_ponderada([(15,0.65),(30,1.6),(45,0.8),(60,0.5),(90,0.2)])
        venc=emissao+timedelta(days=prazo)
        valor=round(_escolha_ponderada([(120,0.5),(250,1.2),(520,1.5),(990,1.3),(1800,0.9),(3500,0.35)]),2)
        atrasodias=max(0,int(abs(random.gauss(1.8,3.8))))
        pago=random.random()<0.88
        data_pag=(venc+timedelta(days=atrasodias)) if pago else None
        multa=round(0.02*valor if (data_pag and data_pag>venc) else 0,2)
        juros=round(0.00033*valor*max(0,((data_pag or datetime.now())-venc).days),2) if (data_pag or datetime.now())>venc else 0.0
        desconto=round(_escolha_ponderada([(0,3.0),(0.02*valor,0.5),(0.05*valor,0.2)]),2) if pago and random.random()<0.1 else 0.0
        liquido=round((valor+multa+juros)-desconto,2) if pago else 0.0
        rows.append({
            "emissao":emissao.date(),"vencimento":venc.date(),"empresa":cli["empresa"],"cnpj":cli["cnpj"],"cidade":cli["cidade"],"uf":cli["uf"],
            "banco":random.choice(BANCOS),"nosso_numero":f"{random.randint(10_000_000_000,99_999_999_999)}",
            "valor_face":valor,"multa":multa,"juros":juros,"desconto":desconto,"pago":pago,"data_pagamento":data_pag.date() if data_pag else None,"valor_liquido":liquido
        })
    return {"titulos":pd.DataFrame(rows),"sacados":pd.DataFrame(clientes).drop_duplicates(subset=["empresa"]).reset_index(drop=True)}

def dataset_logistica(n=1000):
    TRANSPORTADORAS = ["Rapidão Norte","TransLog BR","ViaCargo","Azul Cargo","Correios","JadLog","Total Express","Sequoia","Loggi","Braspress","DDL Express"]
    clientes=[_cliente() for _ in range(max(80,int(n*0.12)))]
    rows=[]
    for _ in range(n):
        cli=random.choice(clientes); coleta=_rand_date(365)
        dias=max(1,int(abs(random.gauss(3.6,1.5))))
        prev=coleta+timedelta(days=dias)
        modal=_escolha_ponderada([("Rodoviário",2.6),("Aéreo",0.6),("Ferroviário",0.4),("Hidroviário",0.3)])
        peso=round(max(0.2, random.gauss(16,9)),2)
        volume=round(max(0.01, random.gauss(0.14,0.08)),3)
        distancia=max(10,int(abs(random.gauss(520,240))))
        base={"Rodoviário":2.1,"Aéreo":4.2,"Ferroviário":1.9,"Hidroviário":1.6}[modal]
        frete=round(base*peso + 0.28*distancia + 12,2)
        entregue=random.random()<0.95
        atraso=max(0,int(abs(random.gauss(0.4,1.0))))
        entrega=(prev+timedelta(days=atraso)) if entregue else None
        rows.append({
            "pedido":f"PED{random.randint(100000,999999)}","cliente":cli["empresa"],
            "origem_uf":random.choice(UFs),"destino_uf":cli["uf"],"modal":modal,
            "coleta":coleta.date(),"previsao_entrega":prev.date(),"entrega":entrega.date() if entrega else None,
            "transportadora":random.choice(TRANSPORTADORAS),"peso_kg":peso,"volume_m3":volume,"distancia_km":distancia,"frete":frete,"entregue":entregue
        })
    return {"embarques":pd.DataFrame(rows),"clientes":pd.DataFrame(clientes).drop_duplicates(subset=["empresa"]).reset_index(drop=True)}

def dataset_agro(n=1000):
    CULTURAS_AGRO = ["Soja","Milho","Cana-de-Açúcar","Café","Algodão","Arroz","Feijão","Trigo","Laranja","Uva"]
    INSUMOS_AGRO = ["Fertilizante NPK","Calcário","Herbicida","Inseticida","Fungicida","Sementes Certificadas","Adubo Orgânico","Micronutrientes","Regulador de Crescimento"]
    produtores=[]
    for _ in range(max(60,int(n*0.1))):
        if _FAKER_OK:
            nome=_FAKE.name(); cidade=_FAKE.city(); uf=_fake_estado_sigla()
        else:
            nome=f"Produtor {random.randint(1000,9999)}"; cidade=f"Cidade {random.randint(1,200)}"; uf=random.choice(UFs)
        produtores.append({"produtor":nome,"cidade":cidade,"uf":uf, **_doc_fakes()})
    talhoes=[f"T{random.randint(1,80)}" for _ in range(160)]
    items=[{"sku":f"AG-{random.randint(1000,9999)}","item":random.choice(INSUMOS_AGRO),"cultura":random.choice(CULTURAS_AGRO),"preco_base":round(_escolha_ponderada([(90,0.6),(120,1.0),(260,1.4),(480,0.9),(950,0.4)]),2)} for _ in range(90)]
    col=[]; ins=[]
    for _ in range(n):
        prod=random.choice(produtores); talhao=random.choice(talhoes); cultura=random.choice(CULTURAS_AGRO)
        area=round(max(1.0, random.gauss(48,22)),1)
        plantio=_rand_date(300); colheita=plantio+timedelta(days=_escolha_ponderada([(110,0.6),(130,1.2),(150,0.9)]))
        produtividade=round(max(0.8, random.gauss(3.2,0.8)),2)
        producao=round(produtividade*area,2)
        preco_t=round(_escolha_ponderada([(850,0.5),(1000,1.1),(1200,1.2),(1400,0.8)]),2)
        receita=round(producao*preco_t,2)
        col.append({"produtor":prod["produtor"],"uf":prod["uf"],"talhao":talhao,"cultura":cultura,"area_ha":area,
                    "plantio":plantio.date(),"colheita":colheita.date(),"produtividade_t_ha":produtividade,"producao_t":producao,
                    "preco_t":preco_t,"receita":receita})
        if random.random()<0.75:
            it=random.choice(items); qtd=max(1,int(abs(random.gauss(8,4))))
            custo=round(it["preco_base"]*qtd*_escolha_ponderada([(0.95,0.5),(1.0,1.2),(1.05,0.6)]),2)
            ins.append({"produtor":prod["produtor"],"talhao":talhao,"cultura":cultura,"item":it["item"],"sku":it["sku"],"qtd":qtd,"custo_total":custo})
    return {"colheita":pd.DataFrame(col),"insumos":pd.DataFrame(ins),"produtores":pd.DataFrame(produtores).drop_duplicates(subset=["produtor"]).reset_index(drop=True),"catalogo":pd.DataFrame(items)}

def dataset_supermercado(n=1000, idioma="pt"):
    base = dataset_market(n, idioma=idioma)  # já PT
    df = base["dados"].copy()
    lojas = [f"Loja {i:02d}" for i in range(1,16)]
    gondolas = [f"G{i:02d}" for i in range(1,31)]
    lotes = [f"L{random.randint(10000,99999)}" for _ in range(n)]
    validade = [datetime.now().date() + timedelta(days=max(1,int(abs(random.gauss(35,25))))) for _ in range(n)]
    df["loja"] = [random.choice(lojas) for _ in range(n)]
    df["gondola"] = [random.choice(gondolas) for _ in range(n)]
    df["lote"] = lotes
    df["validade"] = validade
    base["dados"] = df
    return base

def dataset_rh(n=500):
    if _FAKER_OK:
        def nome(): return _FAKE.name()
        def cidade(): return _FAKE.city()
    else:
        def nome(): return f"Colab {random.randint(1000,9999)}"
        def cidade(): return f"Cidade {random.randint(1,200)}"
    departamentos = ["Vendas","Marketing","Operações","Financeiro","RH","TI","Atendimento","Logística","Jurídico"]
    cargos = ["Assistente","Analista Jr","Analista Pl","Analista Sr","Coordenador","Gerente","Diretor"]
    beneficios = ["VR","VA","VT","Plano Saúde","Odonto","Gympass","Bônus"]
    colabs = []
    for _ in range(max(50,int(n*0.3))):
        sal = round(_escolha_ponderada([(1800,0.4),(2500,1.0),(3500,1.2),(5200,0.8),(7800,0.5),(12000,0.2)]),2)
        adm = _rand_date(900).date()
        dep = random.choice(departamentos); cargo = random.choice(cargos)
        cidade_ = cidade(); uf = _fake_estado_sigla()
        ativo = random.random() < 0.9
        demissao = None if ativo else (adm + timedelta(days=max(30,int(abs(random.gauss(240,180)))))).strftime("%Y-%m-%d")
        colabs.append({"nome":nome(),"cpf":f"{random.randint(100,999)}.{random.randint(100,999)}.{random.randint(100,999)}-{random.randint(10,99)}",
                       "departamento":dep,"cargo":cargo,"salario":sal,"admissao":adm,
                       "demissao":demissao,"cidade":cidade_,"uf":uf,"ativo":ativo})
    eventos=[]
    for _ in range(n):
        c = random.choice(colabs); ref = _rand_date(365).date().replace(day=1)
        bruto = c["salario"]
        desc = round(bruto*_escolha_ponderada([(0.08,1.4),(0.09,1.0),(0.11,0.7)]),2)
        ben = random.sample(beneficios, k=random.randint(1,3))
        ben_val = round(_escolha_ponderada([(300,1.2),(450,1.0),(650,0.7)]),2)
        liquido = round(bruto - desc + ben_val,2)
        eventos.append({"competencia":ref,"colaborador":c["nome"],"departamento":c["departamento"],"cargo":c["cargo"],
                        "salario":bruto,"descontos":desc,"beneficios":", ".join(ben),"valor_beneficios":ben_val,"liquido":liquido})
    return {"colaboradores":pd.DataFrame(colabs), "folha":pd.DataFrame(eventos)}

def dataset_telemarketing(n=1000):
    campanhas = ["Campanha A","Campanha B","Campanha C","Campanha D"]
    operadores = [f"Op-{i:03d}" for i in range(1,80)]
    motivos = ["Sem interesse","Ligação caída","Número incorreto","Proposta enviada","Venda concluída","Agendar retorno"]
    rows=[]
    for _ in range(n):
        data=_rand_date(120)
        dur=int(abs(random.gauss(180,140)))
        resultado=_escolha_ponderada([("Sem contato",1.2),("Contato",1.0),("Venda",0.35),("Follow-up",0.6)])
        valor=0.0
        if resultado=="Venda":
            valor = round(_escolha_ponderada([(79,0.6),(149,1.0),(249,0.6),(399,0.3)]),2)
        rows.append({"data":data.date(),"campanha":random.choice(campanhas),"operador":random.choice(operadores),"resultado":resultado,"duracao_s":dur,"motivo":random.choice(motivos),"valor_venda":valor})
    return {"chamadas":pd.DataFrame(rows)}

def dataset_oficina(n=800):
    marcas=["VW","GM","Fiat","Hyundai","Toyota","Honda","Renault","Peugeot","Citröen","Nissan"]
    servicos=["Troca de óleo","Revisão","Alinhamento","Balanceamento","Freios","Embreagem","Suspensão","Diagnóstico eletrônico","Ar Condicionado","Elétrica"]
    pecas=["Filtro de óleo","Pastilha de freio","Correia dentada","Amortecedor","Velas","Bateria","Filtro de ar","Filtro de combustível","Disco de freio","Kit embreagem"]
    status_list=["Aberta","Em execução","Aguardando peça","Finalizada","Cancelada"]
    rows=[]
    for _ in range(n):
        abertura=_rand_date(180)
        sla_h = max(2,int(abs(random.gauss(14,10))))
        prev_fim = abertura + timedelta(hours=sla_h)
        st = _escolha_ponderada([("Aberta",0.4),("Em execução",0.6),("Aguardando peça",0.3),("Finalizada",1.2),("Cancelada",0.1)])
        fim = None if st in ("Aberta","Em execução","Aguardando peça") else (abertura + timedelta(hours=max(sla_h-2,int(abs(random.gauss(10,7))))))
        mao_obra = round(max(0.5, abs(random.gauss(3.0,2.0))),1)
        peca = random.choice(pecas); serv = random.choice(servicos)
        valor_pecas = round(_escolha_ponderada([(59,0.8),(129,1.2),(249,0.9),(399,0.6)]),2)
        valor_mao = round(mao_obra* _escolha_ponderada([(80,1.0),(95,1.0),(120,0.7)]),2)
        total = valor_pecas + valor_mao
        rows.append({"os": f"OS{random.randint(100000,999999)}","placa": f"{random.choice(['ABC','DEF','GHI'])}-{random.randint(1000,9999)}","marca": random.choice(marcas),"servico": serv,"peca_principal": peca,"abertura": abertura, "sla_previsto_h": sla_h, "prev_fim": prev_fim, "fim": fim,"mao_obra_horas": mao_obra, "valor_pecas": valor_pecas, "valor_mao_obra": valor_mao, "valor_total": total,"status": st})
    return {"ordens":pd.DataFrame(rows)}

def dataset_hortifruti(n=1000):
    itens = ["Banana","Maçã","Laranja","Tomate","Alface","Cenoura","Batata","Cebola","Abacaxi","Uva","Manga","Pimentão","Abobrinha","Morango","Mamão"]
    origens = ["Local","Regional","Importado"]
    rows=[]
    for _ in range(n):
        item=random.choice(itens); origem=random.choice(origens)
        colheita=_rand_date(20).date()
        validade=colheita + timedelta(days=random.randint(3,12))
        preco_kg=round(_escolha_ponderada([(3.99,0.6),(5.99,1.2),(7.99,1.0),(9.99,0.7),(12.99,0.4)]),2)
        peso_kg=round(max(0.2, abs(random.gauss(1.2,0.6))),2)
        valor=round(preco_kg*peso_kg,2)
        rows.append({"data":_rand_date(30).date(),"item":item,"origem":origem,"colheita":colheita,"validade":validade,"peso_kg":peso_kg,"preco_kg":preco_kg,"valor_total":valor,"loja":f"Loja {random.randint(1,20):02d}","gondola":f"HF{random.randint(1,20):02d}"})
    return {"movimento":pd.DataFrame(rows)}

def dataset_exportacao(n=600):
    incoterms=["EXW","FOB","CIF","DAP","DDP"]; moedas=["USD","EUR"]
    ncm_samples=["1001.10.10","0901.11.10","1701.13.00","2710.12.49","8481.30.10","9403.20.00","3004.90.46"]
    portos_origem=["Santos","Paranaguá","Itajaí","Rio Grande","Manaus"]; portos_dest=["Rotterdam","Hamburg","Antwerp","Miami","Shanghai","Valencia","Felixstowe"]
    rows=[]
    for _ in range(n):
        emissao=_rand_date(365)
        inc=random.choice(incoterms); moeda=random.choice(moedas)
        fx = round(_escolha_ponderada([(4.70,0.2),(5.00,0.6),(5.30,0.9),(5.60,0.5)]),2) if moeda=="USD" else round(_escolha_ponderada([(5.20,0.4),(5.40,0.8),(5.70,0.6)]),2)
        qtde = max(1,int(abs(random.gauss(18,9))))
        preco_unit = round(_escolha_ponderada([(12,0.6),(19,1.0),(29,0.8),(39,0.5),(59,0.3)]),2)
        total_moeda = round(qtde*preco_unit,2); total_brl = round(total_moeda*fx,2)
        rows.append({"data": emissao.date(),"cliente": f"Buyer {random.randint(1000,9999)}","incoterm": inc,"ncm": random.choice(ncm_samples),"porto_origem": random.choice(portos_origem), "porto_destino": random.choice(portos_dest),"moeda": moeda,"cambio": fx, "produto": f"Item-{random.randint(100,999)}","quantidade": qtde,"preco_unit_moeda": preco_unit, "total_moeda": total_moeda, "total_brl": total_brl})
    return {"embarques": pd.DataFrame(rows)}

def dataset_estoque(n=1000, idioma="pt"):
    prod_fn = produto_pt_br if idioma=="pt" else produto_auto_pt
    produtos=[prod_fn() for _ in range(240)]
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
        lambda g: pd.Series({"saldo": int(g.apply(lambda r: r["qtd"]*(1 if r["tipo"]=="Entrada" else -1), axis=1).sum()),"valor_mov": round(g["valor"].sum(),2)})
    ).reset_index(drop=True)
    return {"mov": df, "posicao": pos}

# ================== CAMPOS & SPECS ==================
CAMPOS_TEMA = {
    "Market":      ["data","cliente","empresa","uf","cidade","segmento","sku","ean13","produto","categoria","marca","unidade","quantidade","preco_unit","desconto","receita","pagamento"],
    "Financeira":  ["emissao","vencimento","empresa","cnpj","cidade","uf","banco","nosso_numero","valor_face","multa","juros","desconto","pago","data_pagamento","valor_liquido"],
    "Logística":   ["pedido","cliente","origem_uf","destino_uf","modal","coleta","previsao_entrega","entrega","transportadora","peso_kg","volume_m3","distancia_km","frete","entregue"],
    "Agro":        ["produtor","uf","talhao","cultura","area_ha","plantio","colheita","produtividade_t_ha","producao_t","preco_t","receita"],
    "Supermercado":["data","loja","gondola","lote","validade","sku","ean13","produto","categoria","marca","unidade","quantidade","preco_unit","desconto","receita","pagamento"],
    "RH":          ["nome","cpf","departamento","cargo","salario","admissao","demissao","cidade","uf","ativo"],
    "Telemarketing":["data","campanha","operador","resultado","duracao_s","motivo","valor_venda"],
    "Oficina":     ["os","placa","marca","servico","peca_principal","abertura","sla_previsto_h","prev_fim","fim","mao_obra_horas","valor_pecas","valor_mao_obra","valor_total","status"],
    "Hortifruti":  ["data","loja","gondola","item","origem","colheita","validade","peso_kg","preco_kg","valor_total"],
    "Exportação":  ["data","cliente","incoterm","ncm","porto_origem","porto_destino","moeda","cambio","produto","quantidade","preco_unit_moeda","total_moeda","total_brl"],
    "Estoque":     ["data","almox","sku","ean13","produto","categoria","tipo","qtd","custo_unit","valor"],
}
PERFIL_IDX = {
    "Market":{"basico":["data","empresa","produto","categoria","unidade","quantidade","preco_unit","receita"],"completo":CAMPOS_TEMA["Market"]},
    "Financeira":{"basico":["emissao","vencimento","empresa","valor_face","pago","valor_liquido"],"completo":CAMPOS_TEMA["Financeira"]},
    "Logística":{"basico":["pedido","cliente","destino_uf","modal","coleta","previsao_entrega","frete","entregue"],"completo":CAMPOS_TEMA["Logística"]},
    "Agro":{"basico":["produtor","cultura","area_ha","plantio","colheita","producao_t","receita"],"completo":CAMPOS_TEMA["Agro"]},
    "Supermercado":{"basico":["data","loja","produto","categoria","quantidade","preco_unit","receita","validade"],"completo":CAMPOS_TEMA["Supermercado"]},
    "RH":{"basico":["nome","departamento","cargo","salario","admissao","ativo"],"completo":CAMPOS_TEMA["RH"]},
    "Telemarketing":{"basico":["data","campanha","operador","resultado","duracao_s","valor_venda"],"completo":CAMPOS_TEMA["Telemarketing"]},
    "Oficina":{"basico":["os","placa","servico","abertura","valor_total","status"],"completo":CAMPOS_TEMA["Oficina"]},
    "Hortifruti":{"basico":["data","loja","item","peso_kg","preco_kg","valor_total","validade"],"completo":CAMPOS_TEMA["Hortifruti"]},
    "Exportação":{"basico":["data","cliente","incoterm","moeda","total_moeda","total_brl"],"completo":CAMPOS_TEMA["Exportação"]},
    "Estoque":{"basico":["data","almox","sku","produto","tipo","qtd","valor"],"completo":CAMPOS_TEMA["Estoque"]},
}

def _col_def(name: str) -> Dict[str, Any]:
    if name in ("data","emissao","vencimento","coleta","previsao_entrega","entrega","plantio","colheita","data_pagamento","validade","abertura","prev_fim","fim","competencia","admissao","demissao"): return {"name":name,"fmt":"date","width":12}
    if name in ("quantidade","distancia_km","qtd","sla_previsto_h","duracao_s"): return {"name":name,"fmt":"int","width":12}
    if name in ("preco_unit","valor_face","multa","juros","desconto","valor_liquido","frete","preco_t","receita","custo_total","valor_total","valor_pecas","valor_mao_obra","preco_kg","cambio","valor","valor_beneficios","salario","descontos","liquido","preco_unit_moeda","total_moeda","total_brl"): return {"name":name,"fmt":"currency","width":12}
    if name in ("peso_kg","volume_m3","area_ha","produtividade_t_ha","producao_t","mao_obra_horas"): return {"name":name,"fmt":"float","width":12}
    return {"name":name,"fmt":"text","width":max(10,min(26,len(name)+6))}

def build_spec_from_bundle(tema: str, bundle: Dict[str,pd.DataFrame], campos: List[str]) -> Dict[str,Any]:
    base={"workbook":{"title":f"Relatório {tema}","author":"Gerador Interativo","created_at":datetime.now()},
          "dashboard_name":"Dashboard"}

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
        df=bundle["titulos"][campos].copy()
        sac=bundle["sacados"][["empresa","cnpj","cidade","uf","segmento"]]
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
        cli=bundle.get("clientes")
        sheets=[{"name":"Embarques","data":df,"columns":[_col_def(c) for c in df.columns],"freeze":"B2","autofilter":True}]
        if cli is not None:
            sheets.append({"name":"Clientes","data":cli[["empresa","cnpj","cidade","uf"]],"columns":[_col_def(c) for c in ["empresa","cnpj","cidade","uf"]],"freeze":"A2","autofilter":True})
        return {**base,"sheets":sheets,"kpis":[
            {"label":"Frete Total","value":float(bundle["embarques"]["frete"].sum()),"fmt":"currency"},
            {"label":"Peso Total (kg)","value":float(bundle["embarques"]["peso_kg"].sum()),"fmt":"float"},
            {"label":"% Entregue","value":float(bundle["embarques"]["entregue"].mean()*100),"fmt":"float"},
        ],"pivots":[
            {"name":"Frete por Modal","data_sheet":"Embarques","index":["modal"],"columns":[],"values":{"frete":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Frete por Modal","y_title":"R$"}},
        ]}

    if tema=="Agro":
        df=bundle["colheita"][campos].copy()
        ins=bundle["insumos"][["produtor","talhao","cultura","item","sku","qtd","custo_total"]]
        prods=bundle["produtores"][["produtor","cnpj","cpf","cidade","uf"]]
        cat=bundle["catalogo"][["sku","item","cultura","preco_base"]]
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

    if tema=="RH":
        cols=campos
        return {**base,"sheets":[
            {"name":"Colaboradores","data":bundle["colaboradores"][cols if set(cols).issubset(bundle["colaboradores"].columns) else bundle["colaboradores"].columns],"columns":[_col_def(c) for c in (cols if set(cols).issubset(bundle["colaboradores"].columns) else bundle["colaboradores"].columns)],"freeze":"A2","autofilter":True},
            {"name":"Folha","data":bundle["folha"],"columns":[_col_def(c) for c in bundle["folha"].columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Headcount Ativo","value":int(bundle["colaboradores"]["ativo"].sum()),"fmt":"int"},
            {"label":"Massa Salarial","value":float(bundle["colaboradores"]["salario"].sum()),"fmt":"currency"},
            {"label":"Ticket Médio Líquido","value":float(bundle["folha"]["liquido"].mean()),"fmt":"currency"},
        ],"pivots":[
            {"name":"Salário médio por Depto","data_sheet":"Colaboradores","index":["departamento"],"columns":[],"values":{"salario":"mean"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Salário médio por Depto","y_title":"R$"}},
        ]}

    if tema=="Telemarketing":
        df=bundle["chamadas"][campos].copy()
        return {**base,"sheets":[
            {"name":"Chamadas","data":df,"columns":[_col_def(c) for c in df.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Conversões","value":int((bundle["chamadas"]["resultado"]=="Venda").sum()),"fmt":"int"},
            {"label":"Taxa Conversão (%)","value":float((bundle["chamadas"]["resultado"].eq("Venda").mean()*100)),"fmt":"float"},
            {"label":"Duração Média (s)","value":float(bundle["chamadas"]["duracao_s"].mean()),"fmt":"float"},
        ],"pivots":[
            {"name":"Vendas por Campanha","data_sheet":"Chamadas","index":["campanha"],"columns":[],"values":{"valor_venda":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Vendas por Campanha","y_title":"R$"}},
        ]}

    if tema=="Oficina":
        df=bundle["ordens"][campos].copy()
        return {**base,"sheets":[
            {"name":"OS","data":df,"columns":[_col_def(c) for c in df.columns],"freeze":"B2","autofilter":True},
        ],"kpis":[
            {"label":"Faturamento OS","value":float(df["valor_total"].sum()),"fmt":"currency"},
            {"label":"% Finalizadas","value":float(df["status"].eq("Finalizada").mean()*100),"fmt":"float"},
            {"label":"Média Mão de Obra (h)","value":float(df["mao_obra_horas"].mean()),"fmt":"float"},
        ],"pivots":[
            {"name":"Valor por Serviço","data_sheet":"OS","index":["servico"],"columns":[],"values":{"valor_total":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Valor por Serviço","y_title":"R$"}},
        ]}

    if tema=="Hortifruti":
        df=bundle["movimento"][campos].copy()
        return {**base,"sheets":[
            {"name":"Hortifruti","data":df,"columns":[_col_def(c) for c in df.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Receita HF","value":float(df["valor_total"].sum()),"fmt":"currency"},
            {"label":"Kg Vendidos","value":float(df["peso_kg"].sum()),"fmt":"float"},
            {"label":"Preço Médio Kg","value":float((df["valor_total"].sum()/max(1,df["peso_kg"].sum()))),"fmt":"currency"},
        ],"pivots":[
            {"name":"Receita por Item","data_sheet":"Hortifruti","index":["item"],"columns":[],"values":{"valor_total":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"Receita por Item","y_title":"R$"}},
        ]}

    if tema=="Exportação":
        df=bundle["embarques"][campos].copy()
        return {**base,"sheets":[
            {"name":"Exportações","data":df,"columns":[_col_def(c) for c in df.columns],"freeze":"A2","autofilter":True},
        ],"kpis":[
            {"label":"Total BRL","value":float(df["total_brl"].sum()),"fmt":"currency"},
            {"label":"Ticket Médio (moeda)","value":float(df["total_moeda"].mean()),"fmt":"currency"},
            {"label":"Qtde Embarques","value":int(len(df)),"fmt":"int"},
        ],"pivots":[
            {"name":"BRL por Incoterm","data_sheet":"Exportações","index":["incoterm"],"columns":[],"values":{"total_brl":"sum"},"fill_value":0,"round":2,"chart":{"type":"column","title":"BRL por Incoterm","y_title":"R$"}},
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

    raise ValueError("Tema não suportado")

# ================== API pública ==================
def dataset_wrapper(tema: str, n: int, idioma: str):
    if tema=="Market": return dataset_market(n, idioma=idioma)
    if tema=="Supermercado": return dataset_supermercado(n, idioma=idioma)
    if tema=="Estoque": return dataset_estoque(n, idioma=idioma)
    # demais não usam catálogo de produto varejista
    return _TEMAS[tema](n)

_TEMAS = {
    "Market": dataset_market,
    "Financeira": dataset_financeira,
    "Logística": dataset_logistica,
    "Agro": dataset_agro,
    "Supermercado": dataset_supermercado,
    "RH": dataset_rh,
    "Telemarketing": dataset_telemarketing,
    "Oficina": dataset_oficina,
    "Hortifruti": dataset_hortifruti,
    "Exportação": dataset_exportacao,
    "Estoque": dataset_estoque,
}
ALIASES = {
    "market":"Market","mercado":"Market",
    "financeira":"Financeira",
    "logistica":"Logística","logística":"Logística",
    "agro":"Agro",
    "supermercado":"Supermercado","super-mercado":"Supermercado",
    "rh":"RH","recursos-humanos":"RH",
    "telemarketing":"Telemarketing","callcenter":"Telemarketing","call-center":"Telemarketing",
    "oficina":"Oficina","mecanica":"Oficina","mecânica":"Oficina",
    "hortifruti":"Hortifruti",
    "exportacao":"Exportação","exportação":"Exportação",
    "estoque":"Estoque",
}

PERFIS = ["Básico","Completo","Personalizado"]

def listar_temas()->List[str]: return list(_TEMAS.keys())

def gerar_excel_tema(tema: str, n_linhas: int, campos: List[str], output_path: str, estilo="Azul", idioma="pt") -> str:
    if tema not in _TEMAS: raise ValueError(f"Tema inválido. Opções: {listar_temas()}")
    if tema in ("Market","Supermercado","Estoque"):
        bundle=dataset_wrapper(tema, n_linhas, idioma)
    else:
        bundle=_TEMAS[tema](n_linhas)
    spec=build_spec_from_bundle(tema, bundle, campos)
    gerar_planilha(spec, output_path, estilo_key=estilo)
    return output_path

# ================== seleção de campos / CLI ==================
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

def modo_interativo(default_idioma="pt"):
    temas=listar_temas(); estilos=list(ESTILOS.keys())
    tema=temas[prompt_menu("Tema", temas, 0)]
    linhas=prompt_int("Quantidade de linhas", 1000, 1)
    estilo=estilos[prompt_menu("Estilo", estilos, 0)]
    idioma = ["pt","auto"][prompt_menu("Idioma dos produtos (pt/auto)", ["pt","auto"], 0)]
    # para temas sem produtos varejistas, ignorar escolha (mas tudo ok)
    perfil=PERFIS[prompt_menu("Perfil de saída", PERFIS, 0)]
    campos = resolve_campos_por_perfil(tema, perfil)
    saida=input("Arquivo de saída (padrão: saida.xlsx): ").strip() or "saida.xlsx"
    print("\nGerando...")
    bundle_idioma = idioma if tema in ("Market","Supermercado","Estoque") else default_idioma
    caminho=gerar_excel_tema(tema, linhas, campos, saida, estilo=estilo, idioma=bundle_idioma)
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
    p.add_argument("--idioma", default="pt", choices=["pt","auto"], help="Produtos de varejo: pt (força português) ou auto (faker + tradução)")
    p.add_argument("--nao_interativo", action="store_true")
    args=p.parse_args()

    tema=normaliza_tema(args.tema)
    perfil="Básico" if args.perfil.startswith("b") else "Completo" if args.perfil.startswith("c") else "Personalizado"
    campos = resolve_campos_por_perfil(tema, perfil, expr=args.campos if (args.nao_interativo or perfil=="Personalizado") else None)
    caminho=gerar_excel_tema(tema, args.linhas, campos, args.saida, estilo=args.estilo, idioma=args.idioma if tema in ("Market","Supermercado","Estoque") else "pt")
    print(f"✅ Planilha gerada: {caminho}")

if __name__=="__main__":
    if len(sys.argv)==1: modo_interativo()
    else: modo_argparse()
