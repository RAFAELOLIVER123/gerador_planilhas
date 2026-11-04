#!/usr/bin/env python3
from pathlib import Path
import sys
root=Path(sys.argv[1]); f=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
s=f.read_text(encoding='utf-8')
a=s.index('    private void shopping(JSONObject j){'); b=s.index('    private void newListItem(',a)
new=r'''    private java.math.BigDecimal ncListDecimal(String raw,java.math.BigDecimal fallback){
        if(raw==null||raw.trim().isEmpty()||"null".equals(raw))return fallback;
        String n=raw.replace("R$","").trim().replace(" ","");
        if(n.contains(",")&&n.contains("."))n=n.replace(".","").replace(",",".");
        else n=n.replace(",",".");
        try{return new java.math.BigDecimal(n);}catch(Exception e){return fallback;}
    }
    private java.math.BigDecimal ncListAmount(JSONObject item,String field,java.math.BigDecimal fallback){
        return ncListDecimal(val(item,field),fallback);
    }
    private boolean ncListDone(JSONObject item){return "1".equals(val(item,"checked"))||item.optBoolean("checked",false);}
    private java.math.BigDecimal ncListPrice(JSONObject item){
        java.math.BigDecimal paid=ncListAmount(item,"paid_price",null);
        return ncListDone(item)&&paid!=null?paid:ncListAmount(item,"estimated_price",java.math.BigDecimal.ZERO);
    }
    private java.math.BigDecimal ncListTotal(JSONObject item){
        return ncListAmount(item,"quantity",java.math.BigDecimal.ONE).multiply(ncListPrice(item)).setScale(2,java.math.RoundingMode.HALF_UP);
    }
    private String ncListMoney(java.math.BigDecimal amount){
        return java.text.NumberFormat.getCurrencyInstance(new java.util.Locale("pt","BR")).format(amount);
    }
    private String ncListNumber(java.math.BigDecimal amount){
        java.text.NumberFormat fmt=java.text.NumberFormat.getNumberInstance(new java.util.Locale("pt","BR"));
        fmt.setMinimumFractionDigits(2);fmt.setMaximumFractionDigits(2);return fmt.format(amount);
    }
    private void ncListColumn(LinearLayout target,String value,int weight,int size,boolean bold,boolean right){
        TextView field=label(value,size,bold?NAVY:MUTED,bold);
        field.setGravity(Gravity.CENTER_VERTICAL|(right?Gravity.RIGHT:Gravity.LEFT));
        field.setMaxLines(2);field.setPadding(dp(2),dp(4),dp(2),dp(4));
        target.addView(field,new LinearLayout.LayoutParams(0,-2,weight));
    }
    private void ncListSummary(LinearLayout target,String title,java.math.BigDecimal amount,int color){
        LinearLayout tile=vertical();tile.setPadding(dp(5),dp(9),dp(4),dp(9));
        tile.setBackground(shape(Color.rgb(249,246,243),12));
        TextView txt=label(title,10,MUTED,true);txt.setGravity(Gravity.CENTER);tile.addView(txt);
        TextView sum=label(ncListMoney(amount),12,color,true);sum.setGravity(Gravity.CENTER);sum.setMaxLines(2);tile.addView(sum);
        LinearLayout.LayoutParams lp=new LinearLayout.LayoutParams(0,-2,1);lp.setMargins(dp(2),0,dp(2),0);
        target.addView(tile,lp);
    }
    private void ncListToggle(JSONObject item,LinearLayout rowView){
        if(!connected){msg("Conecte o celular para alterar esta lista.");return;}
        java.math.BigDecimal price=ncListPrice(item);
        if(!ncListDone(item)&&price.compareTo(java.math.BigDecimal.ZERO)<=0){msg("Defina um valor unitário para comprar este item.");return;}
        JSONObject payload=withHid(data("id",val(item,"id")));
        if(!ncListDone(item)){
            try{payload.put("paid_price",price.toPlainString());payload.put("paid_on",today());}catch(Exception ignored){}
        }
        rowView.setEnabled(false);
        api.post("list_item_toggle",payload,(res,error)->{
            rowView.setEnabled(true);
            if(error!=null){msg("Não foi possível marcar o produto: "+error);return;}
            loadPage(false);
        });
    }
    private void ncListEdit(JSONObject list,JSONObject item){
        String[] fields={"Produto","Quantidade","Unidade (UN, KG, FD)","Valor unitário (R$)"};
        String[] values={val(item,"title"),val(item,"quantity"),val(item,"unit"),ncListNumber(ncListPrice(item))};
        formDialog("Editar produto",fields,values,a->{
            java.math.BigDecimal q=ncListDecimal(a[1],null),p=ncListDecimal(a[3],null);
            if(a[0].trim().isEmpty()||a[2].trim().isEmpty()||q==null||p==null||q.signum()<=0||p.signum()<0){
                msg("Confira nome, quantidade, unidade e valor unitário.");return;
            }
            JSONObject payload=withHid(data("id",val(item,"id"),"list_id",val(list,"id"),
                "title",a[0].trim(),"quantity",q.toPlainString(),"unit",a[2].trim(),
                "estimated_price",p.toPlainString(),"brand",val(item,"brand"),
                "store_name",val(item,"store_name"),"notes",val(item,"notes")));
            if(ncListDone(item))try{payload.put("paid_price",p.toPlainString());}catch(Exception ignored){}
            if(!connected){msg("Conecte o celular para salvar a edição.");return;}
            api.post("list_item_save",payload,(res,error)->{
                if(error!=null){msg("Não foi possível editar: "+error);return;}
                msg("Produto atualizado.");loadPage(false);
            });
        });
    }
    private void shopping(JSONObject j){
        String title=page.equals("pharmacy")?"Farmácia":page.equals("school")?"Materiais escolares":"Lista de supermercado";
        sectionTitle(title,"Itens, quantidades, preços e o que ainda falta.");
        addButton(content,"＋  Criar lista",PINK,WHITE,()->formDialog("Nova lista",new String[]{"Título","Data prevista (AAAA-MM-DD)"},new String[]{"",today()},a->{
            if(a[0].trim().isEmpty())return;
            mutate("list_save",data("title",a[0],"category",category(),"planned_for",a[1]),true);
        }));
        gap(content,13);pendingBlock();JSONArray lists=array(j,"lists"),items=array(j,"items");int shown=0;
        for(int i=0;i<lists.length();i++){
            final JSONObject list=row(lists,i);if(!category().equals(val(list,"category")))continue;shown++;
            java.util.ArrayList<JSONObject> members=new java.util.ArrayList<>();
            java.math.BigDecimal paid=java.math.BigDecimal.ZERO,open=java.math.BigDecimal.ZERO;int doneCount=0;
            for(int k=0;k<items.length();k++){
                JSONObject item=row(items,k);if(!val(item,"list_id").equals(val(list,"id")))continue;
                members.add(item);
                if(ncListDone(item)){doneCount++;paid=paid.add(ncListTotal(item));}
                else open=open.add(ncListTotal(item));
            }
            LinearLayout c=card(content);text(c,"🛒  "+val(list,"title"),17,NAVY,true);gap(c,5);
            LinearLayout head=horizontal();
            String date=val(list,"planned_for");
            if(date.length()>=10&&date.charAt(4)=='-')date=date.substring(8,10)+"/"+date.substring(5,7)+"/"+date.substring(0,4);
            head.addView(label("Data: "+date,12,MUTED,false),new LinearLayout.LayoutParams(0,-2,1));
            head.addView(label(doneCount+"/"+members.size()+" comprados",12,MUTED,true));c.addView(head);gap(c,12);
            LinearLayout sums=horizontal();c.addView(sums);
            ncListSummary(sums,"COMPRADOS",paid,Color.rgb(28,138,103));
            ncListSummary(sums,"PENDENTES",open,Color.rgb(218,94,78));
            ncListSummary(sums,"TOTAL",paid.add(open),NAVY);
            gap(c,12);text(c,"Toque para marcar  •  Segure para editar",11,MUTED,false);gap(c,9);
            LinearLayout headings=horizontal();
            headings.addView(label("",11,MUTED,true),new LinearLayout.LayoutParams(dp(32),-2));
            ncListColumn(headings,"PRODUTO",41,10,true,false);
            ncListColumn(headings,"QTD.",15,10,true,true);
            ncListColumn(headings,"VAL. UN.",21,10,true,true);
            ncListColumn(headings,"TOTAL",23,10,true,true);
            headings.setPadding(0,dp(8),0,dp(8));headings.setBackground(shape(Color.rgb(246,247,248),9));
            c.addView(headings);
            for(JSONObject item:members){
                LinearLayout r=horizontal();r.setMinimumHeight(dp(56));
                TextView check=label(ncListDone(item)?"☑":"☐",25,Color.rgb(238,124,110),false);
                check.setGravity(Gravity.CENTER);r.addView(check,new LinearLayout.LayoutParams(dp(32),dp(48)));
                ncListColumn(r,val(item,"title"),41,12,true,false);
                String qty=ncListAmount(item,"quantity",java.math.BigDecimal.ONE).stripTrailingZeros().toPlainString().replace('.',',');
                ncListColumn(r,qty+" "+val(item,"unit"),15,11,false,true);
                ncListColumn(r,ncListNumber(ncListPrice(item)),21,11,false,true);
                ncListColumn(r,ncListNumber(ncListTotal(item)),23,11,true,true);
                c.addView(r);
                View line=new View(this);line.setBackgroundColor(Color.rgb(238,239,241));
                c.addView(line,new LinearLayout.LayoutParams(-1,dp(1)));
                r.setOnClickListener(v->ncListToggle(item,r));
                check.setOnClickListener(v->ncListToggle(item,r));
                r.setOnLongClickListener(v->{ncListEdit(list,item);return true;});
                check.setOnLongClickListener(v->{ncListEdit(list,item);return true;});
            }
            gap(c,11);addButton(c,"＋ Adicionar item",MINT,NAVY,()->newListItem(list));
        }
        if(shown==0)placeholder("Crie uma lista para começar as compras.");
        showPendingFor("list_save","Lista pendente");
    }
'''
s=s[:a]+new+s[b:];f.write_text(s,encoding='utf-8')
grad=root/'app/build.gradle';g=grad.read_text();assert 'versionCode 218' in g and "versionName '2.1.8'" in g
grad.write_text(g.replace('versionCode 218','versionCode 219').replace("versionName '2.1.8'","versionName '2.1.9'"))
print('Listas profissionais Android 2.1.9 integradas, outras telas preservadas')
