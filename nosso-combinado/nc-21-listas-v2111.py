#!/usr/bin/env python3
"""Ajuste pontual Android 2.1.11: sem botões textuais de ocultação.
Toque no cabeçalho para expandir/recolher; toque no resumo Comprados para filtrar.
Preserva totais, edição longa, marcação, paginação e os outros módulos.
"""
from pathlib import Path
import sys
root=Path(sys.argv[1])
f=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
s=f.read_text(encoding="utf-8")
def replace_once(old,new,reason):
    global s
    count=s.count(old)
    if count!=1:
        raise RuntimeError(f"{reason}: esperado 1 trecho, encontrado {count}")
    s=s.replace(old,new,1)
    print("OK:",reason)

replace_once(
'''        gap(target,10);
        LinearLayout actions=horizontal();
        TextView toggleBought=button(hideDone?"Mostrar comprados":"Ocultar comprados",Color.rgb(245,246,249),NAVY,()->{});
        actions.addView(toggleBought,new LinearLayout.LayoutParams(0,dp(42),1));
        target.addView(actions);
        toggleBought.setOnClickListener(v->{
            prefs.edit().putBoolean(key+"_hide_done",!hideDone).putInt(key+"_limit",8).apply();
            ncListRenderItems(target,list,members);
        });
        gap(target,9);''',
'''        gap(target,8);''',
"retirar botão textual de comprados")

replace_once(
'''sectionTitle(title,"Resumo compacto • toque em Mostrar itens para expandir.");''',
'''sectionTitle(title,"Itens, quantidades, preços e o que ainda falta.");''',
"remover instrução extra")

replace_once(
'''        final java.util.ArrayList<Integer> counts=new java.util.ArrayList<>();
''',
'',
"remover contador apenas usado no botão")

start='''            final String listKey=category()+":"+val(list,"id");'''
end='''        }
        if(shown==0)placeholder("Crie uma lista para começar as compras.");'''
if s.count(start)!=1 or s.count(end)!=1: raise RuntimeError("estrutura shopping inesperada")
a=s.index(start)
b=s.index(end,a)
replacement=r'''            final String listKey=category()+":"+val(list,"id");
            final boolean initiallyOpen=listKey.equals(ncExpandedListId);
            LinearLayout listCard=card(content);
            LinearLayout titleLine=horizontal();
            TextView listTitle=label("🛒  "+val(list,"title"),17,NAVY,true);
            titleLine.addView(listTitle,new LinearLayout.LayoutParams(0,-2,1));
            TextView chevron=label(initiallyOpen?"⌃":"⌄",22,MUTED,true);
            chevron.setGravity(Gravity.CENTER);
            chevron.setContentDescription(initiallyOpen?"Recolher lista":"Expandir lista");
            titleLine.addView(chevron,new LinearLayout.LayoutParams(dp(38),dp(36)));
            listCard.addView(titleLine);gap(listCard,4);
            LinearLayout head=horizontal();
            String date=val(list,"planned_for");
            if(date.length()>=10&&date.charAt(4)=='-')
                date=date.substring(8,10)+"/"+date.substring(5,7)+"/"+date.substring(0,4);
            head.addView(label("Data: "+date,12,MUTED,false),new LinearLayout.LayoutParams(0,-2,1));
            head.addView(label(doneCount+"/"+members.size()+" comprados",12,MUTED,true));
            listCard.addView(head);gap(listCard,11);
            LinearLayout sums=horizontal();listCard.addView(sums);
            ncListSummary(sums,"COMPRADOS",paid,Color.rgb(28,138,103));
            ncListSummary(sums,"PENDENTES",open,Color.rgb(218,94,78));
            ncListSummary(sums,"TOTAL",paid.add(open),NAVY);
            LinearLayout body=vertical();
            body.setVisibility(initiallyOpen?View.VISIBLE:View.GONE);
            listCard.addView(body);
            details.add(body);buttons.add(chevron);keys.add(listKey);
            if(initiallyOpen)ncListRenderItems(body,list,members);
            // O título e a data expandem/recolhem; nenhum botão extra ocupa espaço.
            titleLine.setContentDescription(initiallyOpen?"Recolher lista":"Expandir lista");
            titleLine.setOnClickListener(v->{
                ncExpandedListId=listKey.equals(ncExpandedListId)?"":listKey;
                for(int n=0;n<details.size();n++){
                    boolean active=keys.get(n).equals(ncExpandedListId);
                    LinearLayout panel=details.get(n);
                    panel.setVisibility(active?View.VISIBLE:View.GONE);
                    TextView arrow=buttons.get(n);
                    arrow.setText(active?"⌃":"⌄");
                    arrow.setContentDescription(active?"Recolher lista":"Expandir lista");
                    if(!active)panel.removeAllViews();
                }
                if(listKey.equals(ncExpandedListId))ncListRenderItems(body,list,members);
            });
            head.setOnClickListener(v->titleLine.performClick());
            // O próprio resumo COMPRADOS alterna os itens já marcados sem botão textual.
            final android.content.SharedPreferences prefs=getSharedPreferences("nc_listas_ui",MODE_PRIVATE);
            final String filterKey=ncListUIKey(list);
            View boughtTile=sums.getChildAt(0);
            boughtTile.setContentDescription("Mostrar ou ocultar produtos comprados");
            boughtTile.setAlpha(prefs.getBoolean(filterKey+"_hide_done",false)?0.55f:1f);
            boughtTile.setOnClickListener(v->{
                boolean hide=prefs.getBoolean(filterKey+"_hide_done",false);
                prefs.edit().putBoolean(filterKey+"_hide_done",!hide).putInt(filterKey+"_limit",8).apply();
                boughtTile.setAlpha(!hide?0.55f:1f);
                if(body.getVisibility()==View.VISIBLE)ncListRenderItems(body,list,members);
            });
'''
s=s[:a]+replacement+s[b:]
for bad in ["Ocultar itens","Ocultar comprados","Mostrar comprados","Mostrar itens ("]:
    # apenas no bloco desta nova implementação
    after=s[s.index("    private String ncExpandedListId="):s.index("    private void newListItem(")]
    if bad in after:raise RuntimeError("texto antigo ainda presente: "+bad)
f.write_text(s,encoding='utf-8')
grad=root/'app/build.gradle'
g=grad.read_text(encoding='utf-8')
if g.count('versionCode 220')!=1 or g.count("versionName '2.1.10'")!=1:
    raise RuntimeError("versão base 2.1.10 não encontrada")
grad.write_text(g.replace('versionCode 220','versionCode 221').replace("versionName '2.1.10'","versionName '2.1.11'"),encoding='utf-8')
print("Android 2.1.11 pronto: apenas toque, sem botões de ocultação.")
