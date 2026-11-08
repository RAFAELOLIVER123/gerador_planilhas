#!/usr/bin/env python3
"""2.1.15: mês a mês, pagamento parcial, estorno e planejamento mensal Android nativo."""
from pathlib import Path
import sys
root=Path(sys.argv[1] if len(sys.argv)>1 else 'android-v1')
main=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
grad=root/'app/build.gradle'
source=main.read_text(encoding='utf-8')
fragment=Path('nosso-combinado/nc-native-finance-v225.javafrag').read_text(encoding='utf-8')
def once(before,after,what):
    global source
    found=source.count(before)
    if found!=1:raise RuntimeError(f"{what}: esperado uma ocorrência, encontrado {found}")
    source=source.replace(before,after,1)
    print('OK',what)

once('    private void toolCarousel(){',fragment+'\n    private void toolCarousel(){',
     'acrescentar financeiro, mês a mês e planejamento nativos')
once('case "finance":finance(j);break;',
     'case "finance":nfFinanceView(j);break;',
     'financeiro original agora usa fluxo mensal completo')
once('if(page.equals("price-compare")){ncComparePage();return;}',
     'if(page.equals("finance-matrix")){nfLoadMatrix();return;}if(page.equals("price-compare")){ncComparePage();return;}',
     'carregar planejamento com API oficial')
once('else if("matrix".equals(target))ncOpenPortal("finance-matrix");',
     'else if("matrix".equals(target))navigate("finance-matrix");',
     'planejamento do carrossel agora nativo')
once('case 7:ncOpenPortal("finance-matrix");break;',
     'case 7:navigate("finance-matrix");break;',
     'planejamento do menu de ferramentas agora nativo')
once('String[][] nav={{"home","⌂","Início"},{"shopping","🛒","Listas"},{"price-compare","📊","Preços"},{"finance","R$","Financeiro"},{"more","☷","Mais"}};',
     'String[][] nav={{"home","⌂","Início"},{"shopping","🛒","Listas"},{"finance","R$","Financeiro"},{"finance-matrix","▦","Planejar"},{"more","☷","Mais"}};',
     'menu inferior com financeiro e planejamento')
# No segundo lançamento financeiro: proteger o total da lista de qualquer edição de status.
g=grad.read_text(encoding='utf-8')
assert g.count('versionCode 224')==1
assert g.count("versionName '2.1.14'")==1
g=g.replace('versionCode 224','versionCode 225').replace("versionName '2.1.14'","versionName '2.1.15'")
main.write_text(source,encoding='utf-8')
grad.write_text(g,encoding='utf-8')
for name in ['nfFinanceView(', 'nfLoadMatrix(', 'nfPaymentDialog(', 'nfHistory(', 'nfMatrixNewRow(', 'nfEditCell(']:
    assert name in source
assert 'ncNativeListForm(' in source
assert 'ncComparePage(' in source
print('Versão 2.1.15 pronta para compilar sem alterações no backend ou no sistema de alarmes.')
