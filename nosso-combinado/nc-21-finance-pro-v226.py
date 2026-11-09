#!/usr/bin/env python3
"""NC 2.1.16: novo cadastro de contas/receitas com parcelas e fotos nativas."""
from pathlib import Path
import sys
root=Path(sys.argv[1] if len(sys.argv)>1 else 'android-v1')
main=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
manifest=root/'app/src/main/AndroidManifest.xml'
gradle=root/'app/build.gradle'
s=main.read_text(encoding='utf-8')
helper=Path('nosso-combinado/nc-finance-pro-helpers-v226.javafrag').read_text(encoding='utf-8')
form=Path('nosso-combinado/nc-finance-pro-form-v226.javafrag').read_text(encoding='utf-8')
def once(before,after,label):
    global s
    n=s.count(before)
    if n!=1:raise RuntimeError(f'{label}: esperado 1 trecho, encontrados {n}')
    s=s.replace(before,after,1)
    print('OK:',label)

once('    private void toolCarousel(){',helper+'\n'+form+'\n    private void toolCarousel(){',
    'injetar formulários e operações nativas')
once('addButton(content,"＋ Novo lançamento",PINK,WHITE,this::newFinance);',
     'addButton(content,"＋ Nova conta ou receita",PINK,WHITE,()->ncProForm(null));',
     'novo lançamento do Financeiro com cadastro completo')
once('addButton(content,"＋ Nova linha no Planejamento",PINK,WHITE,()->nfMatrixNewRow());',
     'addButton(content,"＋ Nova conta ou receita",PINK,WHITE,()->ncProForm(null));',
     'planejamento usa o mesmo formulário e as mesmas regras')
original='text(tile,val(r,"title"),16,NAVY,true);'
modified='''text(tile,val(r,"title"),16,NAVY,true);
            if(rowId>0){
                TextView details=button("✎ Editar conta, classificação e parcelas",
                    Color.rgb(237,243,253),NAVY,()->ncProForm(r));
                LinearLayout.LayoutParams editP=new LinearLayout.LayoutParams(-1,dp(41));
                editP.topMargin=dp(7);tile.addView(details,editP);
                ncProAttachButton(tile,rowId);
            }'''
once(original,modified,'editar cadastro no Planejamento, ver fotos')
# Adicionar atalho para tela completa também no menu de ferramentas.
once('case 7:navigate("finance-matrix");break;',
     'case 7:navigate("finance-matrix");break;',
     'manter planejamento nativo no menu')
main.write_text(s,encoding='utf-8')
xml=manifest.read_text(encoding='utf-8')
assert 'NcFinancePhotoActivity' not in xml
assert xml.count('</application>')==1
xml=xml.replace('</application>',
     '  <activity android:name=".NcFinancePhotoActivity" android:exported="false" />\n</application>',1)
manifest.write_text(xml,encoding='utf-8')
g=gradle.read_text(encoding='utf-8')
for old,new in [('versionCode 225','versionCode 226'),
                ("versionName '2.1.15'","versionName '2.1.16'")]:
    if g.count(old)!=1:raise RuntimeError('Versão base inesperada: '+old)
    g=g.replace(old,new,1)
gradle.write_text(g,encoding='utf-8')
assert 'finance_matrix_schedule_save' in s
assert 'NCP_KIND_CODES' in s
assert 'NcFinancePhotoActivity.class' in s
print('Android 2.1.16: estrutura nativa anterior mantida; módulo de cadastro ampliado.')
