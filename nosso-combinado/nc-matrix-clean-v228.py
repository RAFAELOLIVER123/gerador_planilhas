#!/usr/bin/env python3
"""Nosso Combinado Android 2.1.18 — Planejamento mensal limpo.
Remove fotos e botões de edição das linhas, concentra ações no pressionar prolongado.
Preserva pagamentos, estornos, demais telas nativas e dados do servidor.
"""
from pathlib import Path
import sys

root=Path(sys.argv[1] if len(sys.argv)>1 else "android-v1")
main=root/"app/src/main/java/com/nossocombinado/app/MainActivity.java"
gradle=root/"app/build.gradle"
source=main.read_text(encoding="utf-8")
fragment=Path("nosso-combinado/nc-matrix-clean-menu-v228.javafrag").read_text(encoding="utf-8")
needle="    private void toolCarousel(){"
if source.count(needle)!=1:raise RuntimeError("toolCarousel: âncora inesperada")
source=source.replace(needle,fragment+"\n"+needle,1)
start=source.index("    private void nfRenderMatrix(")
end=source.index("    private void nfMatrixNewRow(",start)
matrix=source[start:end]
def change(old,new,label):
    global matrix
    count=matrix.count(old)
    if count!=1:raise RuntimeError(f"{label}: esperado 1 ocorrência, encontradas {count}")
    matrix=matrix.replace(old,new,1)
    print("OK",label)

change('''        text(content,"Competências do mês",17,NAVY,true);gap(content,8);''',
'''        text(content,"Lançamentos do mês",16,NAVY,true);
        text(content,"Toque e segure um lançamento para editar, pagar ou estornar.",11,MUTED,false);
        gap(content,7);''',"cabeçalho orientativo compacto")
change('''            LinearLayout tile=card(content);''',
'''            LinearLayout tile=card(content);
            tile.setPadding(dp(14),dp(10),dp(14),dp(11));''',"linhas mais compactas")
v226='''            text(tile,val(r,"title"),16,NAVY,true);
            if(rowId>0){
                TextView details=button("✎ Editar conta, classificação e parcelas",
                    Color.rgb(237,243,253),NAVY,()->ncProForm(r));
                LinearLayout.LayoutParams editP=new LinearLayout.LayoutParams(-1,dp(41));
                editP.topMargin=dp(7);tile.addView(details,editP);
                ncProAttachButton(tile,rowId);
            }'''
menuTitle='''            LinearLayout titleBar=horizontal();
            titleBar.setGravity(Gravity.CENTER_VERTICAL);
            TextView rowTitle=label(val(r,"title"),15,NAVY,true);
            titleBar.addView(rowTitle,new LinearLayout.LayoutParams(0,-2,1));
            TextView options=button("⋮",Color.rgb(245,247,251),NAVY,
                 ()->nfMatrixRowMenu(r,cell));
            options.setContentDescription("Ver opções de "+val(r,"title"));
            options.setTextSize(19);
            titleBar.addView(options,new LinearLayout.LayoutParams(dp(38),dp(38)));
            tile.addView(titleBar);'''
change(v226,menuTitle,"remover botão editar e fotos; trocar por menu discreto")
actionBlock='''                if(virtual||cell.optInt("source_list",0)>0){
                    gap(tile,6);text(tile,"✓ Item pago de lista. Ajuste pagamentos na lista original.",11,MUTED,false);
                    TextView open=button("Abrir listas",Color.rgb(237,242,248),NAVY,()->navigate("shopping"));
                    tile.addView(open,new LinearLayout.LayoutParams(-1,dp(42)));
                }else{
                    long tid=cell.optLong("transaction_id",0);
                    LinearLayout actions=horizontal();
                    TextView edit=button("✎ Valor / data",Color.rgb(239,242,250),NAVY,()->nfEditCell(r,cell));
                    actions.addView(edit,new LinearLayout.LayoutParams(0,dp(42),1));
                    if(tid>0&&remain.signum()>0){
                        JSONObject payTx=new JSONObject();
                        try {
                            payTx.put("transaction_id",tid);
                            payTx.put("title",val(r,"title"));
                            payTx.put("amount",total);
                            payTx.put("paid_amount",done);
                            payTx.put("direction",val(r,"direction"));
                        }catch(Exception ignore){}
                        TextView pay=button(income?"＋ Receber":"✓ Pagar",Color.rgb(227,242,235),Color.rgb(15,108,82),
                             ()->nfPaymentDialog(payTx,true));
                        LinearLayout.LayoutParams p=new LinearLayout.LayoutParams(0,dp(42),1);p.leftMargin=dp(5);
                        actions.addView(pay,p);
                    }
                    tile.addView(actions);
                    if(tid>0){
                        gap(tile,6);
                        TextView history=button("Histórico e estornos",Color.rgb(250,250,252),NAVY,()->nfHistory(tid,true));
                        tile.addView(history,new LinearLayout.LayoutParams(-1,dp(40)));
                    }
                }'''
replacement='''                if(virtual||cell.optInt("source_list",0)>0||
                    r.optLong("source_list_id",0)>0){
                    gap(tile,4);
                    text(tile,"🛒 Vinculado à lista de compras",11,MUTED,false);
                }'''
change(actionBlock,replacement,"remover botões de pagamento, histórico e edição das linhas")
change('''                text(tile,"Sem previsão neste mês.",12,MUTED,false);
                TextView add=button("＋ Planejar valor",Color.rgb(233,243,249),NAVY,()->nfEditCell(r,null));
                tile.addView(add,new LinearLayout.LayoutParams(-1,dp(42)));''',
'''                text(tile,"Sem valor previsto neste mês.",12,MUTED,false);''',
"substituir botão de planejar por ação de pressionar")
change('''            gap(content,7);
        }
        if(displayed==0){''',
'''            tile.setLongClickable(true);
            tile.setOnLongClickListener(v->{
                v.performHapticFeedback(android.view.HapticFeedbackConstants.LONG_PRESS);
                nfMatrixRowMenu(r,cell);
                return true;
            });
            tile.setContentDescription(val(r,"title")+
                ". Pressione e segure para ver as opções.");
            gap(content,6);
        }
        if(displayed==0){''',
"ativar pressionar prolongado em cada lançamento")

assert "ncProAttachButton" not in matrix
assert "Editar conta, classificação e parcelas" not in matrix
assert "nfMatrixRowMenu(r,cell)" in matrix
assert "nfPaymentDialog(payTx" not in matrix
source=source[:start]+matrix+source[end:]

# As fotos continuam disponíveis em novas contas do Financeiro e nos produtos das listas,
# mas não são mostradas nos formulários abertos pelo Planejamento mensal.
formStart=source.index("    private void ncProFormReady(")
photosStart=source.index('        ncProSection(form,"📷","Fotos da conta"',formStart)
photosEnd=source.index("        LinearLayout footer=horizontal();",photosStart)
photoSection=source[photosStart:photosEnd]
source=source[:photosStart]+'''        if(!"finance-matrix".equals(page)){
'''+photoSection+'''        }

'''+source[photosEnd:]
print("OK", "ocultar fotos apenas no Planejamento mensal, mantendo-as no Financeiro")

g=gradle.read_text(encoding="utf-8")
for before,after in [("versionCode 227","versionCode 228"),
                     ("versionName '2.1.17'","versionName '2.1.18'")]:
    if g.count(before)!=1:raise RuntimeError("Versão base inesperada: "+before)
    g=g.replace(before,after,1)
gradle.write_text(g,encoding="utf-8")
main.write_text(source,encoding="utf-8")

assert source.count("private void nfMatrixRowMenu(")==1
assert source.count('if(!"finance-matrix".equals(page)){')==1
assert "ncListEditWithPhotos" in source
assert "ncProPhotoSelect" in source
assert "finance_matrix_payment_reverse" in source
assert "nc_alarm05" in source
print("Android 2.1.18 pronto: Planejamento compacto, menu ao segurar, fotos mantidas fora dele.")
