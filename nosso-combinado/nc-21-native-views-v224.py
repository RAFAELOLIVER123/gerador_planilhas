#!/usr/bin/env python3
"""Atualização nativa 2.1.14: compras completas, comparador Android e navegação compacta."""
from pathlib import Path
import re,sys
root=Path(sys.argv[1] if len(sys.argv)>1 else "android-v1")
main=root/"app/src/main/java/com/nossocombinado/app/MainActivity.java"
manifest=root/"app/src/main/AndroidManifest.xml"
gradle=root/"app/build.gradle"
fragment=Path("nosso-combinado/nc-native-views-v224.javafrag").read_text(encoding="utf-8")
s=main.read_text(encoding="utf-8")
def once(before,after,label):
    global s
    count=s.count(before)
    if count!=1:raise RuntimeError(f"{label}: encontrou {count} ocorrências")
    s=s.replace(before,after,1)
    print("OK",label)
# Alterar apenas uma tela nova, preservando o resto dos métodos existentes.
needle="    private void toolCarousel(){"
once(needle,fragment+"\n"+needle,"adicionar UI Android nativa")
start=s.index("    private void toolCarousel(){")
end=s.index("    private void appScreen(){root();",start)
newbar=r'''    private void toolCarousel(){
        HorizontalScrollView reel=new HorizontalScrollView(this);
        reel.setHorizontalScrollBarEnabled(false);reel.setFillViewport(false);
        reel.setBackgroundColor(SURFACE);
        LinearLayout track=horizontal();track.setPadding(dp(8),dp(5),dp(8),dp(5));
        String[][] options={{"home","⌂ Início"},{"new-list","＋ Nova lista"},
          {"shopping","🛒 Listas"},{"price-compare","📊 Comparador"},
          {"finance","R$ Financeiro"},{"matrix","▦ Planejamento"},{"tools","☷ Ferramentas"}};
        for(String[] item:options){
            String target=item[0];boolean active=page.equals(target);
            TextView chip=button(item[1],active?PINK:Color.rgb(245,246,252),active?WHITE:NAVY,()->{
                if("new-list".equals(target))ncNativeListForm(null);
                else if("matrix".equals(target))ncOpenPortal("finance-matrix");
                else if("tools".equals(target))ncToolsMenu();
                else navigate(target);
            });
            chip.setTextSize(12);chip.setPadding(dp(11),dp(8),dp(11),dp(8));
            LinearLayout.LayoutParams p=new LinearLayout.LayoutParams(-2,dp(40));
            p.rightMargin=dp(6);track.addView(chip,p);
        }
        reel.addView(track,new HorizontalScrollView.LayoutParams(-2,-2));
        base.addView(reel,new LinearLayout.LayoutParams(-1,dp(52)));
    }
'''
s=s[:start]+newbar+s[end:]
# Barra inferior compacta, com os principais módulos em telas Android.
old='String[][] nav={{"home","⌂","Início"},{"tasks","✓","Tarefas"},{"calendar","▦","Agenda"},{"network","★","Rede"},{"more","☷","Mais"}};'
new='String[][] nav={{"home","⌂","Início"},{"shopping","🛒","Listas"},{"price-compare","📊","Preços"},{"finance","R$","Financeiro"},{"more","☷","Mais"}};'
once(old,new,"menu inferior com comparação e financeiro")
once('slot.setOnClickListener(v->navigate(n[0]));','slot.setOnClickListener(v->{if("more".equals(n[0]))ncToolsMenu();else navigate(n[0]);});',"menu de todas as ferramentas")
once('if(page.equals("more")){showMore();return;}',
     'if(page.equals("price-compare")){ncComparePage();return;}if(page.equals("more")){showMore();return;}',
     "comparador nativo sem redirecionar")
# Formulário original de duas entradas substituído pelo formulário completo.
begin=s.index('addButton(content,"＋  Criar lista",PINK,WHITE,()->formDialog(')
end=s.index('        gap(content,13);',begin)
s=s[:begin]+'''addButton(content,"＋ Nova lista de compras",PINK,WHITE,()->ncNativeListForm(null));
'''+s[end:]
# Em Compras e listas: todas as categorias podem ser consultadas, incluindo Casa/Outros.
once('if(!category().equals(val(list,"category")))continue;shown++;',
     'if(!page.equals("shopping")&&!category().equals(val(list,"category")))continue;shown++;',
     "mostrar todas categorias em Compras e listas")
once('TextView listTitle=label("🛒  "+val(list,"title"),17,NAVY,true);',
     'TextView listTitle=label("🧺  "+val(list,"title"),16,NAVY,true);',
     "cabeçalho compacto de lista")
needle='titleLine.addView(chevron,new LinearLayout.LayoutParams(dp(38),dp(36)));'
once(needle,needle+'''
            TextView editList=button("Editar",Color.rgb(238,241,250),NAVY,()->ncNativeListForm(list));
            editList.setTextSize(11);
            titleLine.addView(editList,new LinearLayout.LayoutParams(dp(63),dp(37)));''',
            "editar lista completa com datas e loja")
main.write_text(s,encoding="utf-8")
xml=manifest.read_text(encoding="utf-8")
assert "NcListAlarmReceiver" not in xml
assert xml.count("</application>")==1
xml=xml.replace('</application>','''  <receiver android:name=".NcListAlarmReceiver" android:exported="true">
      <intent-filter>
          <action android:name="android.intent.action.BOOT_COMPLETED"/>
          <action android:name="android.intent.action.MY_PACKAGE_REPLACED"/>
      </intent-filter>
  </receiver>
</application>''',1)
if "android.permission.RECEIVE_BOOT_COMPLETED" not in xml:
    xml=xml.replace("<application",'<uses-permission android:name="android.permission.RECEIVE_BOOT_COMPLETED"/>\n<application',1)
manifest.write_text(xml,encoding="utf-8")
g=gradle.read_text(encoding="utf-8")
for a,b in [('versionCode 223','versionCode 224'),("versionName '2.1.13'","versionName '2.1.14'")]:
    assert g.count(a)==1,(a,g.count(a))
    g=g.replace(a,b,1)
gradle.write_text(g,encoding="utf-8")
assert s.count('private void ncNativeListForm(')==1
assert s.count('private void ncComparePage(')==1
print("Android 2.1.14 pronto: formulários, comparador e navegação nativos. Núcleo original mantido.")
