#!/usr/bin/env python3
"""Nosso Combinado Android 2.1.13: adição mínima à versão nativa 2.1.12.
Preserva todo o projeto Android (lembretes, agenda, offline, chat, fotos e mapa).
Adiciona acessos nativos às ferramentas web atualizadas. Sem duplicação de dados.
"""
from pathlib import Path
import sys
root=Path(sys.argv[1] if len(sys.argv)>1 else "android-v1")
main=root/"app/src/main/java/com/nossocombinado/app/MainActivity.java"
manifest=root/"app/src/main/AndroidManifest.xml"
gradle=root/"app/build.gradle"
folder=main.parent
source=main.read_text(encoding="utf-8")
def once(old,new,name):
    global source
    n=source.count(old)
    if n!=1:raise RuntimeError(f"{name}: esperado uma ocorrência, encontrado {n}")
    source=source.replace(old,new,1)
    print("OK",name)

once('{"shopping","🛒  Listas"},{"notes","✎  Notas"}',
     '{"shopping","🛒  Listas"},{"shopping-online","🛒  Listas atualizadas"},'
     '{"price-compare","📊  Comparador de preços"},{"finance-online","R$  Financeiro completo"},'
     '{"finance-matrix","▦  Planejamento mensal"},{"notes","✎  Notas"}',
     "atalhos adicionais no carrossel nativo")

entry='    private void toolCarousel(){'
assert source.count(entry)==1,"Método toolCarousel ausente/duplicado"
helper='''
    /** Abre módulos atualizados da plataforma mantendo todas as telas e serviços Android. */
    private void ncOpenPortal(String key){
        String page=key;
        if("shopping-online".equals(page))page="shopping";
        if("finance-online".equals(page))page="finance";
        android.content.Intent intent=new android.content.Intent(this,WebToolActivity.class);
        intent.putExtra("nc_page",page);
        startActivity(intent);
    }
'''
source=source.replace(entry,helper+entry,1)
begin=source.index(entry)
end=source.index('    private void appScreen(){root();',begin)
portion=source[begin:end]
old='()->navigate(dest)'
assert portion.count(old)==1,("toolCarousel callback mudou",portion.count(old))
new='()->{if(dest.equals("shopping-online")||dest.equals("price-compare")||dest.equals("finance-online")||dest.equals("finance-matrix"))ncOpenPortal(dest);else navigate(dest);}'
portion=portion.replace(old,new,1)
source=source[:begin]+portion+source[end:]
main.write_text(source,encoding="utf-8")
g=gradle.read_text(encoding="utf-8")
for a,b in [('versionCode 222','versionCode 223'),("versionName '2.1.12'","versionName '2.1.13'")]:
    assert g.count(a)==1,(a,g.count(a))
    g=g.replace(a,b,1)
gradle.write_text(g,encoding="utf-8")
xml=manifest.read_text(encoding="utf-8")
assert xml.count('</application>')==1
assert "WebToolActivity" not in xml
xml=xml.replace('</application>','''  <activity android:name=".WebToolActivity"
      android:exported="false"
      android:windowSoftInputMode="adjustResize" />
</application>''',1)
manifest.write_text(xml,encoding="utf-8")
print("Versão nativa 2.1.13 preparada; componentes originais preservados.")
