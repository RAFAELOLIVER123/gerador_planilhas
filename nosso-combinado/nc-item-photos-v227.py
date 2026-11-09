#!/usr/bin/env python3
"""Nosso Combinado Android 2.1.17: fotos nativas em produtos, exceto Veículos.
Preserva banco de dados, eventos de lista, pagamento e a arquitetura original.
"""
from pathlib import Path
import sys

root=Path(sys.argv[1] if len(sys.argv)>1 else "android-v1")
main=root/"app/src/main/java/com/nossocombinado/app/MainActivity.java"
gradle=root/"app/build.gradle"
fragment=Path("nosso-combinado/nc-item-photo-v227.javafrag").read_text(encoding="utf-8")
source=main.read_text(encoding="utf-8")
def once(before,after,label):
    global source
    count=source.count(before)
    if count!=1:raise RuntimeError(f"{label}: esperado 1 trecho, encontrados {count}")
    source=source.replace(before,after,1)
    print("OK",label)

once('    private void toolCarousel(){',
     fragment+'\n    private void toolCarousel(){',
     "Adicionar interface e armazenamento de fotos de itens")

once('    private void ncListEdit(JSONObject list,JSONObject item){',
     '''    private void ncListEdit(JSONObject list,JSONObject item){
        if(ncItemPhotosAllowed(list)){
            ncListEditWithPhotos(list,item);
            return;
        }''',
     "Editar itens com controles de fotos, mantendo editor antigo de Veículos")

once('check.setGravity(Gravity.CENTER);r.addView(check,new LinearLayout.LayoutParams(dp(32),dp(48)));',
     '''check.setGravity(Gravity.CENTER);r.addView(check,new LinearLayout.LayoutParams(dp(32),dp(48)));
            android.widget.ImageView productPhoto=ncItemPhotoThumbnail(list,item);
            if(productPhoto!=null){
                LinearLayout.LayoutParams pp=new LinearLayout.LayoutParams(dp(32),dp(36));
                pp.rightMargin=dp(5);
                r.addView(productPhoto,pp);
            }''',
     "Miniatura de foto clicável em cada linha de produto permitida")

once('fields.addView(label("Quantidade de compra e preço por unidade ou embalagem.",12,MUTED,false));',
     '''fields.addView(label("Quantidade de compra e preço por unidade ou embalagem.",12,MUTED,false));
        TextView photoTip=label("📷 Depois de salvar a lista, toque na câmera do item para adicionar fotos.",11,MUTED,false);
        fields.addView(photoTip);
        cat.setOnItemSelectedListener(new android.widget.AdapterView.OnItemSelectedListener(){
            public void onNothingSelected(android.widget.AdapterView<?> p){}
            public void onItemSelected(android.widget.AdapterView<?> p,View v,int pos,long id){
                photoTip.setVisibility(categories[pos].equals("vehicle")?View.GONE:View.VISIBLE);
            }
        });''',
     "Explicação do uso de fotos apenas nas categorias permitidas")

g=gradle.read_text(encoding="utf-8")
for before,after in [('versionCode 226','versionCode 227'),
                     ("versionName '2.1.16'","versionName '2.1.17'")]:
    if g.count(before)!=1:raise RuntimeError(f"Versão base inesperada: {before}")
    g=g.replace(before,after,1)
gradle.write_text(g,encoding="utf-8")
main.write_text(source,encoding="utf-8")
for expected in ["private boolean ncItemPhotosAllowed(", "private void ncListEditWithPhotos(",
                 "private android.widget.ImageView ncItemPhotoThumbnail(",
                 "private void ncProPhotoSelect(", "ncNativeListForm(", "finance_matrix_schedule_save"]:
    if expected not in source:raise RuntimeError("Função necessária ausente: "+expected)
print("Android 2.1.17: fotos de itens habilitadas em supermercado/farmácia/escola/casa/outros/contas, exceto veículos.")
