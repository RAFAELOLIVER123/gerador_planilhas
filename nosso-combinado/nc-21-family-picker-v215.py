#!/usr/bin/env python3
"""Nosso Combinado nativo 2.1.5: seletor de família na Home e seção Família."""
from pathlib import Path
import sys
root=Path(sys.argv[1]).resolve()
main=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
grad=root/'app/build.gradle'
s=main.read_text(encoding='utf-8')
def change(old,new,label):
    global s
    n=s.count(old)
    if n!=1:raise RuntimeError(f'{label}: expected 1, found {n}')
    s=s.replace(old,new,1)
    print('OK',label)

change('    private void switchFamilyDialog(){', '''    /** Seletor reutilizável mostrando o nome da família ativa. */
    private void familySwitcher(LinearLayout parent){
        LinearLayout picker=horizontal();
        picker.setPadding(dp(13),dp(10),dp(13),dp(10));
        picker.setBackground(outline(SURFACE));
        picker.setElevation(dp(2));
        TextView heart=label("♥",23,PINK,true);
        heart.setGravity(Gravity.CENTER);
        picker.addView(heart,new LinearLayout.LayoutParams(dp(34),dp(48)));
        gapH(picker,7);
        LinearLayout labels=vertical();
        text(labels,"FAMÍLIA ATIVA",11,MUTED,true);
        gap(labels,2);
        TextView current=label(householdName.isEmpty()?"Selecione a família":householdName,16,NAVY,true);
        current.setMaxLines(2);
        current.setEllipsize(android.text.TextUtils.TruncateAt.END);
        labels.addView(current);
        picker.addView(labels,new LinearLayout.LayoutParams(0,-2,1));
        TextView arrow=label("⌄",26,PINK,true);
        arrow.setGravity(Gravity.CENTER);
        arrow.setContentDescription("Abrir lista de famílias");
        picker.addView(arrow,new LinearLayout.LayoutParams(dp(35),dp(48)));
        picker.setContentDescription("Família ativa: "+householdName+". Toque para escolher outra família.");
        picker.setClickable(true);
        picker.setOnClickListener(v->switchFamilyDialog());
        LinearLayout.LayoutParams item=new LinearLayout.LayoutParams(-1,-2);
        item.bottomMargin=dp(9);
        parent.addView(picker,item);
    }
    private void switchFamilyDialog(){''','family switcher component')
change('            String[] options=new String[groups.length()];',
       '            String[] options=new String[groups.length()];int selectedIndex=-1;',
       'selected group index')
change('                boolean active=householdId.equals(val(f,"id"));',
       '                boolean active=householdId.equals(val(f,"id"));if(active)selectedIndex=i;',
       'checked current family')
change('''            dialogBuilder().setTitle("Trocar de família")
              .setItems(options,(dialog,which)->{''',
       '''            dialogBuilder().setTitle("Escolha sua família")
              .setSingleChoiceItems(options,selectedIndex,(dialog,which)->{''',
       'radio list in elegant dialog')
change('''                  if(householdId.equals(val(selected,"id")))return;
                  chooseHousehold(selected);
                  familyTab="group";page="members";
                  appScreen();''',
       '''                  dialog.dismiss();
                  if(householdId.equals(val(selected,"id")))return;
                  final String previousPage=page;
                  // Offline data stays scoped to the original family.
                  chooseHousehold(selected);
                  privateCid=0;
                  privateTitle="";
                  if("members".equals(previousPage))familyTab="group";
                  page="private_thread".equals(previousPage)?"home":previousPage;
                  appScreen();''',
       'keep user on current page and refresh active family')
change('text(artText,"Nossa família · "+householdName,15,NAVY,true);',
       'familySwitcher(artText);',
       'replace home group label with dropdown')
change('sectionTitle("Nossa família ♡","Grupo: "+householdName);',
       'sectionTitle("Nossa família ♡","Selecione qual família deseja visualizar.");familySwitcher(content);',
       'show family selector in members page')
change('addButton(content,"⇄  Trocar de família",PINK,WHITE,this::switchFamilyDialog);gap(content,8);',
       '',
       'remove old duplicate family switching button')
main.write_text(s,encoding='utf-8')
g=grad.read_text(encoding='utf-8')
for old,new in [("versionCode 214","versionCode 215"),("versionName '2.1.4'","versionName '2.1.5'")]:
    if g.count(old)!=1:raise RuntimeError('gradle mismatch: '+old)
    g=g.replace(old,new,1)
grad.write_text(g,encoding='utf-8')
print('APK original 2.1.5: family dropdown ready, other modules preserved.')
