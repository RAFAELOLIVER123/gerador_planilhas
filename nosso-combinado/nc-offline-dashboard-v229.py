#!/usr/bin/env python3
"""Android 2.1.19: filtro por responsável e consulta offline; dashboard exclusivo da web.
Sem modificar APIs, dados do servidor, agenda, chat, fotos ou alarmes existentes.
"""
from pathlib import Path
import sys
root=Path(sys.argv[1] if len(sys.argv)>1 else "android-v1")
main=root/"app/src/main/java/com/nossocombinado/app/MainActivity.java"
gradle=root/"app/build.gradle"
frag=Path("nosso-combinado/nc-offline-dashboard-v229.javafrag").read_text(encoding="utf-8")
s=main.read_text(encoding="utf-8")
def once(a,b,label):
    global s
    count=s.count(a)
    if count!=1:raise RuntimeError(f'{label}: esperado 1, encontrado {count}')
    s=s.replace(a,b,1)
    print("OK",label)
once("    private void toolCarousel(){",frag+"\n    private void toolCarousel(){",
     "injetar cache financeiro e filtro de pessoa")
# Cabeçalho: filtro e botão de baixar dados para consulta sem internet.
anchor='''        content.addView(selector);gap(content,10);
        LinearLayout views=horizontal();'''
with_buttons='''        content.addView(selector);gap(content,9);
        LinearLayout mobileTools=horizontal();mobileTools.setGravity(Gravity.CENTER_VERTICAL);
        TextView byPerson=button("👤 "+nfOwnerName(),Color.rgb(233,240,250),NAVY,
              ()->nfOwnerPicker(matrix));
        byPerson.setContentDescription("Filtrar por responsável da família");
        LinearLayout.LayoutParams personSize=new LinearLayout.LayoutParams(0,dp(43),1.25f);
        mobileTools.addView(byPerson,personSize);
        TextView sync=button(connected?"⬇ Baixar offline":"◌ Offline",
              Color.rgb(226,241,234),Color.rgb(20,105,81),this::nfSyncOffline);
        LinearLayout.LayoutParams syncSize=new LinearLayout.LayoutParams(0,dp(43),1.05f);
        syncSize.leftMargin=dp(6);
        mobileTools.addView(sync,syncSize);
        content.addView(mobileTools);gap(content,7);
        if(!connected||nfOfflineView)
            text(content,"◌ Dados salvos no aparelho · consulta offline, sem pagamentos ou alterações.",11,MUTED,false);
        LinearLayout views=horizontal();'''
once(anchor,with_buttons,"adicionar filtro e botão baixar offline na navegação")
# Finance: cached local by user+house and filtered correctly.
finance_start=s.index("    private void nfFinanceView(")
finance_end=s.index("    private boolean nfOnline(){",finance_start)
finance=s[finance_start:finance_end]
def finreplace(a,b,label):
    global finance
    count=finance.count(a)
    if count!=1:raise RuntimeError(f"Financeiro - {label}: esperado 1, encontrado {count}")
    finance=finance.replace(a,b,1)
    print("OK",label)
finreplace('''        nfLedger=data==null?new JSONObject():data;
        content.removeAllViews();nfHeader(false);''',
'''        if(!connected){
            JSONObject saved=nfRestoreOffline(NF_LEDGER_KEY);
            nfLedger=saved!=null?saved:(data==null?new JSONObject():data);
            nfOfflineView=true;
        }else{
            nfLedger=data==null?new JSONObject():data;
            nfSaveOffline(NF_LEDGER_KEY,nfLedger);
            nfOfflineView=false;
        }
        content.removeAllViews();nfHeader(false);''',"restituir caixa e receitas offline")
finreplace('''            if(!nfMonth.equals(nfCompetence(tx)))continue;''',
'''            if(!nfMonth.equals(nfCompetence(tx))||!nfMatches(tx))continue;''',
"filtrar lançamentos por responsável")
s=s[:finance_start]+finance+s[finance_end:]
# Matrix: restore last 24 months offline, including when server is down.
load_start=s.index("    private void nfLoadMatrix(){")
load_end=s.index("    private void nfRenderMatrix(",load_start)
updated_load=r'''    private void nfLoadMatrix(){
        if(!page.equals("finance-matrix")||content==null)return;
        int seq=++nfMatrixSeq;
        content.removeAllViews();nfHeader(true);
        text(content,"Preparando Planejamento mensal...",13,MUTED,false);
        JSONObject cached=nfRestoreOffline(NF_MATRIX_KEY);
        if(!connected){
            if(cached!=null&&nfMatrixHasMonth(cached,nfMonth)){
                nfOfflineView=true;nfMatrixSnapshot=cached;nfRenderMatrix(cached);
            }else{
                text(content,"Este mês não foi baixado para consulta offline.",14,MUTED,false);
                text(content,"Conecte-se e toque em Baixar offline antes de viajar ou ficar sem sinal.",12,MUTED,false);
            }
            return;
        }
        final String start=nfMonthAdd(nfMonth,-11);
        JSONObject payload=withHid(data("start",start,"months","24"));
        api.get("finance_matrix",payload,(value,error)->{
            if(seq!=nfMatrixSeq||!page.equals("finance-matrix")||content==null)return;
            if(error!=null||value==null){
                if(cached!=null&&nfMatrixHasMonth(cached,nfMonth)){
                    nfOfflineView=true;nfMatrixSnapshot=cached;nfRenderMatrix(cached);
                    msg("Sem resposta do servidor. Exibindo os últimos dados salvos.");
                }else{
                    content.removeAllViews();nfHeader(true);
                    text(content,"Não foi possível carregar os valores deste mês.",14,PINK,false);
                    addButton(content,"Tentar novamente",Color.rgb(237,241,248),NAVY,this::nfLoadMatrix);
                }
                return;
            }
            try{
                value.put("nc_cache_start",start);
                value.put("nc_cache_months",24);
            }catch(Exception ignored){}
            nfSaveOffline(NF_MATRIX_KEY,value);
            nfMatrixSnapshot=value;
            nfOfflineView=false;
            nfRenderMatrix(value);
        });
    }
'''
s=s[:load_start]+updated_load+s[load_end:]
print("OK", "cache de 24 meses por família e leitura offline do planejamento")
# Matrix: filtering only; dashboard remains exclusively on the website.
matrix_start=s.index("    private void nfRenderMatrix(")
matrix_end=s.index("    private void nfMatrixNewRow(",matrix_start)
matrix=s[matrix_start:matrix_end]
def matreplace(a,b,label):
    global matrix
    count=matrix.count(a)
    if count!=1:raise RuntimeError(f"Planejamento - {label}: esperado 1, encontrado {count}")
    matrix=matrix.replace(a,b,1)
    print("OK",label)
matreplace('''        content.removeAllViews();nfHeader(true);
        addButton(content,"＋ Nova conta ou receita",PINK,WHITE,()->ncProForm(null));''',
'''        nfMatrixSnapshot=data;
        content.removeAllViews();nfHeader(true);
        if(connected&&!nfOfflineView)
            addButton(content,"＋ Nova conta ou receita",PINK,WHITE,()->ncProForm(null));''',
"não oferecer edição no Planejamento offline")
matreplace('''            JSONObject r=row(rows,i),cell=selected.get(r.optLong("id",0));
            if(cell==null)continue;''',
'''            JSONObject r=row(rows,i),cell=selected.get(r.optLong("id",0));
            if(cell==null||!nfMatches(r))continue;''',
"totais do Planejamento por responsável")
matreplace('''            JSONObject r=row(rows,i);
            long rowId=r.optLong("id",0);''',
'''            JSONObject r=row(rows,i);
            if(!nfMatches(r))continue;
            long rowId=r.optLong("id",0);''',
"linhas do Planejamento por responsável")
matreplace('''            TextView options=button("⋮",Color.rgb(245,247,251),NAVY,
                 ()->nfMatrixRowMenu(r,cell));''',
'''            TextView options=button("⋮",Color.rgb(245,247,251),NAVY,
                 ()->{if(!connected||nfOfflineView)msg("Modo offline: somente consulta.");else nfMatrixRowMenu(r,cell);});''',
"menu de ações protege leitura offline")
matreplace('''                nfMatrixRowMenu(r,cell);
                return true;''',
'''                if(!connected||nfOfflineView)msg("Conecte-se para alterar este lançamento.");
                else nfMatrixRowMenu(r,cell);
                return true;''',
"pressionar longo não altera dados offline")
s=s[:matrix_start]+matrix+s[matrix_end:]
# In every mutating financial action, guard disconnected network.
needle='''    private void nfMatrixRowMenu(JSONObject entry,JSONObject cell){'''
once(needle,needle+'''
        if(!connected||nfOfflineView){msg("Planejamento offline: consulta somente. Conecte-se para editar ou pagar.");return;}''',
"bloquear ações por menu offline")
# Bump version, preserve package key separation for test builds.
g=gradle.read_text(encoding="utf-8")
for old,new in [("versionCode 228","versionCode 229"),("versionName '2.1.18'","versionName '2.1.19'")]:
    if g.count(old)!=1:raise RuntimeError("Versão inesperada "+old)
    g=g.replace(old,new,1)
gradle.write_text(g,encoding="utf-8")
main.write_text(s,encoding="utf-8")
for method in ["nfOwnerPicker(","nfSyncOffline(","nfMatrixHasMonth(","nfRestoreOffline("]:
    assert method in s
assert s.count("private void nfLoadMatrix(){")==1
assert s.count("private void nfFinanceView(")==1
assert "ncListEditWithPhotos" in s and "nfMatrixRowMenu" in s
assert "finance_matrix_payment_reverse" in s
assert "nfDashboardSummary(" not in s and "nfMonthChart(" not in s
print("Android 2.1.19 preparado; servidor permanece como fonte de verdade.")
