#!/usr/bin/env python3
"""Correção mínima da versão original 2.1.3: Família e botões de teste."""
from pathlib import Path
import sys
root=Path(sys.argv[1])
p=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
s=p.read_text(encoding="utf-8")
def change(old,new,label):
 global s
 count=s.count(old)
 if count!=1:raise RuntimeError(label+": encontrado "+str(count))
 s=s.replace(old,new,1)
 print("OK "+label)
change('addButton(content,"♫  Testar som de mensagem",MINT,NAVY,()->ncPlayAudio(R.raw.nc_mensagem));gap(content,8);addButton(content,"♫  Testar notificação",MINT,NAVY,()->ncPlayAudio(R.raw.nc_notificacao));gap(content,8);addButton(content,"♫  Testar novo post",MINT,NAVY,()->ncPlayAudio(R.raw.nc_novopost));gap(content,8);addButton(content,"♫  Testar interações",MINT,NAVY,()->ncPlayAudio(R.raw.nc_interacoes));gap(content,8);','','retirar botões para testar sons')
change('addButton(content,"🔔  Alertas de tarefas e compromissos",MINT,NAVY,()->{ReminderScheduler.ensureChannel(this);ReminderScheduler.requestNotificationPermission(this);if(android.os.Build.VERSION.SDK_INT>=26){Intent settings=new Intent(android.provider.Settings.ACTION_APP_NOTIFICATION_SETTINGS).putExtra(android.provider.Settings.EXTRA_APP_PACKAGE,getPackageName());startActivity(settings);}else msg("Confira as permissões de notificações do aplicativo.");});gap(content,10);','','retirar atalho de alerta sem desativar alertas')
change('        addButton(content,"⏰  Testar alarme de tarefa (8 segundos)",MINT,NAVY,this::testAlarms);gap(content,10);\n','','retirar botão de testar alarme')
change('    private void chooseHousehold(JSONObject h){householdId=val(h,"id");householdName=val(h,"name");role=val(h,"role");vault.set("hid",householdId);vault.set("household_name",householdName);vault.set("role",role);}', '''    private void chooseHousehold(JSONObject h){householdId=val(h,"id");householdName=val(h,"name");role=val(h,"role");vault.set("hid",householdId);vault.set("household_name",householdName);vault.set("role",role);}
    /** Seleciona somente famílias em que o usuário autenticado participa. */
    private void switchFamilyDialog(){
        if(userId.isEmpty()){authScreen(false);return;}
        api.get("me",null,(result,error)->{
            if(error!=null){msg("Não foi possível consultar suas famílias: "+error);return;}
            final JSONArray groups=array(result,"households");
            if(groups.length()<=1){msg("Você só participa de um grupo. Crie ou entre em outra família para alternar.");return;}
            String[] options=new String[groups.length()];
            for(int i=0;i<groups.length();i++){
                JSONObject f=row(groups,i);
                boolean active=householdId.equals(val(f,"id"));
                options[i]=(active?"✓  ":"○  ")+val(f,"name")+(active?" (atual)":"");
            }
            dialogBuilder().setTitle("Trocar de família")
              .setItems(options,(dialog,which)->{
                  JSONObject selected=row(groups,which);
                  if(householdId.equals(val(selected,"id")))return;
                  chooseHousehold(selected);
                  familyTab="group";page="members";
                  appScreen();
                  msg("Família ativa: "+householdName);
              }).setNegativeButton("Cancelar",null).show();
        });
    }''','dialogo de alternância de famílias')
change('sectionTitle("Nossa família ♡","Grupo: "+householdName);addButton(content,"+ Criar meu próprio grupo",MINT,NAVY,this::createOwnFamily);gap(content,10);',
  'sectionTitle("Nossa família ♡","Grupo: "+householdName);addButton(content,"⇄  Trocar de família",PINK,WHITE,this::switchFamilyDialog);gap(content,8);addButton(content,"+ Criar meu próprio grupo",MINT,NAVY,this::createOwnFamily);gap(content,10);','botão de troca no início da página')
change('gap(content,12);addButton(content,"Trocar de família",MINT,NAVY,()->{householdId="";vault.remove("hid");refreshMe();});return;','gap(content,12);return;','eliminar troca antiga que voltava ao primeiro grupo')
p.write_text(s,encoding="utf-8")
grad=root/'app/build.gradle';g=grad.read_text(encoding="utf-8")
for old,new in [("versionCode 213","versionCode 214"),("versionName '2.1.3'","versionName '2.1.4'")]:
 if g.count(old)!=1:raise RuntimeError(old+" not found")
 g=g.replace(old,new,1)
grad.write_text(g,encoding="utf-8")
print("Versão 2.1.4: recursos e alarmes originais preservados.")
