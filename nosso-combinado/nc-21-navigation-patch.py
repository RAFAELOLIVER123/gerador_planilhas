#!/usr/bin/env python3
from pathlib import Path
import sys
root=Path(sys.argv[1])
main=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
grad=root/'app/build.gradle'
s=main.read_text(encoding='utf-8')
def change(a,b,label):
 global s
 assert s.count(a)==1,(label,s.count(a))
 s=s.replace(a,b,1)
 print('OK',label)
change('topBar.addView(bell);','''topBar.addView(bell);
        TextView messagesShortcut=button("💬",MINT,NAVY,()->navigate("messages"));
        messagesShortcut.setTextSize(16);messagesShortcut.setContentDescription("Abrir mensagens privadas");
        LinearLayout.LayoutParams pm=new LinearLayout.LayoutParams(dp(43),dp(43));pm.leftMargin=dp(7);
        topBar.addView(messagesShortcut,pm);''','messages top icon')
change('    private void appScreen(){root();','''    private void toolCarousel(){
       HorizontalScrollView reel=new HorizontalScrollView(this);
       reel.setHorizontalScrollBarEnabled(false);reel.setFillViewport(false);
       reel.setBackgroundColor(SURFACE);
       LinearLayout track=horizontal();track.setPadding(dp(10),dp(6),dp(10),dp(7));
       String[][] tools={{"home","⌂  Início"},{"chat","💬  Família"},{"messages","✉  Privado"},
         {"tasks","✓  Tarefas"},{"calendar","▦  Agenda"},{"map","📍  Mapa"},
         {"shopping","🛒  Listas"},{"notes","✎  Notas"},{"finance","R$  Financeiro"},
         {"members","♥  Família"},{"network","★  Rede"},{"documents","▣  Documentos"},{"more","☷  Mais"}};
       for(String[] tool:tools){
          final String dest=tool[0];boolean selected=dest.equals(page);
          TextView chip=button(tool[1],selected?Color.rgb(255,228,219):Color.rgb(249,246,243),selected?PINK:NAVY,()->navigate(dest));
          chip.setTextSize(12);chip.setPadding(dp(13),dp(9),dp(13),dp(9));chip.setContentDescription("Abrir "+tool[1]);
          LinearLayout.LayoutParams params=new LinearLayout.LayoutParams(-2,dp(39));params.rightMargin=dp(7);track.addView(chip,params);
       }
       reel.addView(track,new HorizontalScrollView.LayoutParams(-2,-2));
       base.addView(reel,new LinearLayout.LayoutParams(-1,dp(56)));
    }
    private void appScreen(){root();''','carousel navigation')
change('smallHeader("Nosso Combinado",householdName.isEmpty()?"Nosso espaço":householdName);',
 'smallHeader("Nosso Combinado",householdName.isEmpty()?"Nosso espaço":householdName);toolCarousel();','insert carousel')
main.write_text(s,encoding='utf-8')
g=grad.read_text(encoding='utf-8')
for a,b in [("versionCode 211","versionCode 212"),("versionName '2.1.1'","versionName '2.1.2'")]:
 assert g.count(a)==1,a
 g=g.replace(a,b,1)
grad.write_text(g,encoding='utf-8')
print('Base 2.1 original preservada; adicionado carrossel de ferramentas.')
