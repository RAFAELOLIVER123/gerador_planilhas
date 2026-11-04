#!/usr/bin/env python3
"""Nosso Combinado Android 2.1.10: listas dobráveis, filtro de comprados e paginação compacta.
Preserva as demais telas e a API existente da versão 2.1.9.
"""
from pathlib import Path
import sys
root=Path(sys.argv[1])
f=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
s=f.read_text(encoding='utf-8')
a=s.index('    private void shopping(JSONObject j){')
b=s.index('    private void newListItem(',a)
new=r'''
    private String ncExpandedListId="";
    private String ncListUIKey(JSONObject list){
        return "list_"+userId+"_"+vault.get("hid")+"_"+category()+"_"+val(list,"id");
    }
    private void ncListRenderItems(LinearLayout target,JSONObject list,java.util.ArrayList<JSONObject> members){
        target.removeAllViews();
        String key=ncListUIKey(list);
        android.content.SharedPreferences prefs=getSharedPreferences("nc_listas_ui",MODE_PRIVATE);
        boolean hideDone=prefs.getBoolean(key+"_hide_done",false);
        int limit=prefs.getInt(key+"_limit",8);
        if(limit<8)limit=8;
        gap(target,10);
        LinearLayout actions=horizontal();
        TextView toggleBought=button(hideDone?"Mostrar comprados":"Ocultar comprados",Color.rgb(245,246,249),NAVY,()->{});
        actions.addView(toggleBought,new LinearLayout.LayoutParams(0,dp(42),1));
        target.addView(actions);
        toggleBought.setOnClickListener(v->{
            prefs.edit().putBoolean(key+"_hide_done",!hideDone).putInt(key+"_limit",8).apply();
            ncListRenderItems(target,list,members);
        });
        gap(target,9);
        LinearLayout headings=horizontal();
        headings.addView(label("",11,MUTED,true),new LinearLayout.LayoutParams(dp(32),-2));
        ncListColumn(headings,"PRODUTO",41,10,true,false);
        ncListColumn(headings,"QTD.",15,10,true,true);
        ncListColumn(headings,"VAL. UN.",21,10,true,true);
        ncListColumn(headings,"TOTAL",23,10,true,true);
        headings.setPadding(0,dp(7),0,dp(7));
        headings.setBackground(shape(Color.rgb(246,247,248),9));
        target.addView(headings);
        int totalVisible=0,rendered=0;
        for(JSONObject item:members){
            if(hideDone&&ncListDone(item))continue;
            totalVisible++;
            if(rendered>=limit)continue;
            rendered++;
            LinearLayout r=horizontal();r.setMinimumHeight(dp(54));
            TextView check=label(ncListDone(item)?"☑":"☐",25,Color.rgb(238,124,110),false);
            check.setGravity(Gravity.CENTER);r.addView(check,new LinearLayout.LayoutParams(dp(32),dp(48)));
            ncListColumn(r,val(item,"title"),41,12,true,false);
            String qty=ncListAmount(item,"quantity",java.math.BigDecimal.ONE)
                .stripTrailingZeros().toPlainString().replace('.',',');
            ncListColumn(r,qty+" "+val(item,"unit"),15,11,false,true);
            ncListColumn(r,ncListNumber(ncListPrice(item)),21,11,false,true);
            ncListColumn(r,ncListNumber(ncListTotal(item)),23,11,true,true);
            target.addView(r);
            View line=new View(this);line.setBackgroundColor(Color.rgb(238,239,241));
            target.addView(line,new LinearLayout.LayoutParams(-1,dp(1)));
            r.setOnClickListener(v->ncListToggle(item,r));
            check.setOnClickListener(v->ncListToggle(item,r));
            r.setOnLongClickListener(v->{ncListEdit(list,item);return true;});
            check.setOnLongClickListener(v->{ncListEdit(list,item);return true;});
        }
        if(totalVisible==0){
            gap(target,7);
            text(target,hideDone?"Nenhum produto pendente nesta lista.":"Nenhum produto nesta lista.",12,MUTED,false);
        }
        if(totalVisible>rendered){
            gap(target,7);
            final int next=limit+8;
            addButton(target,"Ver mais ("+(totalVisible-rendered)+" restantes)",MINT,NAVY,()->{
                prefs.edit().putInt(key+"_limit",next).apply();
                ncListRenderItems(target,list,members);
            });
        }
        if(rendered>8){
            gap(target,6);
            addButton(target,"Mostrar menos",Color.rgb(245,246,249),NAVY,()->{
                prefs.edit().putInt(key+"_limit",8).apply();
                ncListRenderItems(target,list,members);
            });
        }
        gap(target,11);
        addButton(target,"＋ Adicionar item",MINT,NAVY,()->newListItem(list));
    }
    private void shopping(JSONObject j){
        String title=page.equals("pharmacy")?"Farmácia":page.equals("school")?"Materiais escolares":"Lista de supermercado";
        sectionTitle(title,"Resumo compacto • toque em Mostrar itens para expandir.");
        addButton(content,"＋  Criar lista",PINK,WHITE,()->formDialog(
            "Nova lista",new String[]{"Título","Data prevista (AAAA-MM-DD)"},new String[]{"",today()},a->{
                if(a[0].trim().isEmpty())return;
                mutate("list_save",data("title",a[0],"category",category(),"planned_for",a[1]),true);
            }
        ));
        gap(content,13);
        pendingBlock();
        JSONArray lists=array(j,"lists"),items=array(j,"items");int shown=0;
        final java.util.ArrayList<LinearLayout> details=new java.util.ArrayList<>();
        final java.util.ArrayList<TextView> buttons=new java.util.ArrayList<>();
        final java.util.ArrayList<String> keys=new java.util.ArrayList<>();
        final java.util.ArrayList<Integer> counts=new java.util.ArrayList<>();
        for(int i=0;i<lists.length();i++){
            final JSONObject list=row(lists,i);if(!category().equals(val(list,"category")))continue;shown++;
            final java.util.ArrayList<JSONObject> members=new java.util.ArrayList<>();
            java.math.BigDecimal paid=java.math.BigDecimal.ZERO,open=java.math.BigDecimal.ZERO;int doneCount=0;
            for(int k=0;k<items.length();k++){
                JSONObject item=row(items,k);if(!val(item,"list_id").equals(val(list,"id")))continue;
                members.add(item);
                if(ncListDone(item)){doneCount++;paid=paid.add(ncListTotal(item));}
                else open=open.add(ncListTotal(item));
            }
            final String listKey=category()+":"+val(list,"id");
            LinearLayout card=card(content);
            text(card,"🛒  "+val(list,"title"),17,NAVY,true);gap(card,5);
            LinearLayout head=horizontal();
            String date=val(list,"planned_for");
            if(date.length()>=10&&date.charAt(4)=='-')
                date=date.substring(8,10)+"/"+date.substring(5,7)+"/"+date.substring(0,4);
            head.addView(label("Data: "+date,12,MUTED,false),new LinearLayout.LayoutParams(0,-2,1));
            head.addView(label(doneCount+"/"+members.size()+" comprados",12,MUTED,true));
            card.addView(head);gap(card,11);
            LinearLayout sums=horizontal();card.addView(sums);
            ncListSummary(sums,"COMPRADOS",paid,Color.rgb(28,138,103));
            ncListSummary(sums,"PENDENTES",open,Color.rgb(218,94,78));
            ncListSummary(sums,"TOTAL",paid.add(open),NAVY);
            gap(card,10);
            final boolean initiallyOpen=listKey.equals(ncExpandedListId);
            TextView toggle=button(initiallyOpen?"▲ Ocultar itens":"▼ Mostrar itens ("+members.size()+")",
                Color.rgb(235,248,243),NAVY,()->{});
            card.addView(toggle,new LinearLayout.LayoutParams(-1,dp(45)));
            LinearLayout body=vertical();
            body.setVisibility(initiallyOpen?View.VISIBLE:View.GONE);
            card.addView(body);
            details.add(body);buttons.add(toggle);keys.add(listKey);counts.add(members.size());
            if(initiallyOpen)ncListRenderItems(body,list,members);
            toggle.setOnClickListener(v->{
                ncExpandedListId=listKey.equals(ncExpandedListId)?"":listKey;
                for(int n=0;n<details.size();n++){
                    boolean active=keys.get(n).equals(ncExpandedListId);
                    LinearLayout panel=details.get(n);
                    panel.setVisibility(active?View.VISIBLE:View.GONE);
                    buttons.get(n).setText(active?"▲ Ocultar itens":
                        "▼ Mostrar itens ("+counts.get(n)+")");
                    if(!active)panel.removeAllViews();
                }
                if(listKey.equals(ncExpandedListId))ncListRenderItems(body,list,members);
            });
        }
        if(shown==0)placeholder("Crie uma lista para começar as compras.");
        showPendingFor("list_save","Lista pendente");
    }
'''
s=s[:a]+new+s[b:]
f.write_text(s,encoding='utf-8')
grad=root/'app/build.gradle'
g=grad.read_text(encoding='utf-8')
assert 'versionCode 219' in g and "versionName '2.1.9'" in g
grad.write_text(g.replace('versionCode 219','versionCode 220').replace("versionName '2.1.9'","versionName '2.1.10'"),encoding='utf-8')
assert 'ncExpandedListId' in f.read_text(encoding='utf-8')
assert 'ncListRenderItems' in f.read_text(encoding='utf-8')
print('Android 2.1.10: listas recolhíveis, filtros e paginação de 8 produtos.')
