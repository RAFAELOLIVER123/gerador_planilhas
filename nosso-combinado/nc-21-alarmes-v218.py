#!/usr/bin/env python3
"""2.1.8: lembretes pontuais + notificação ao disparar + diagnóstico de permissões."""
from pathlib import Path
import sys
root=Path(sys.argv[1]);java=root/"app/src/main/java/com/nossocombinado/app"
schedule=java/"ReminderScheduler.java";receiver=java/"ReminderReceiver.java"
main=java/"MainActivity.java";grad=root/"app/build.gradle"
def edit(file,old,new,label):
    s=file.read_text(encoding="utf-8");n=s.count(old)
    if n!=1:raise RuntimeError(label+": anchor matches "+str(n))
    file.write_text(s.replace(old,new,1),encoding="utf-8")
    print("OK",label)

edit(schedule,'    public static void ensureChannel(Context c){',"""
    public static synchronized String nextAlarm(Context c){
        JSONObject all=records(c);java.util.Iterator<String> keys=all.keys();
        long soon=Long.MAX_VALUE,now=System.currentTimeMillis();String title="";
        while(keys.hasNext()){
            JSONObject value=all.optJSONObject(keys.next());if(value==null)continue;
            long at=value.optLong("at",0);
            if(at>now&&at<soon){soon=at;title=value.optString("title","Compromisso");}
        }
        if(soon==Long.MAX_VALUE)return "Nenhum alarme futuro registrado neste aparelho";
        java.text.SimpleDateFormat local=new java.text.SimpleDateFormat("dd/MM/yyyy 'às' HH:mm",new Locale("pt","BR"));
        return local.format(new Date(soon))+" • "+title;
    }
    public static boolean notificationsAllowed(Context c){
        NotificationManager nm=(NotificationManager)c.getSystemService(Context.NOTIFICATION_SERVICE);
        if(nm!=null&&!nm.areNotificationsEnabled())return false;
        if(Build.VERSION.SDK_INT>=33&&c.checkSelfPermission("android.permission.POST_NOTIFICATIONS")!=PackageManager.PERMISSION_GRANTED)return false;
        if(Build.VERSION.SDK_INT>=26&&nm!=null){
            NotificationChannel channel=nm.getNotificationChannel("nc_due_v218");
            if(channel!=null&&channel.getImportance()==NotificationManager.IMPORTANCE_NONE)return false;
        }
        return true;
    }
    public static void ensureChannel(Context c){""","status local de lembretes")

edit(schedule,'    public static void cancelFromOp(Context c,JSONObject item,String uid,String hid){',"""
    /** Toca na hora marcada e, se configurado, antes dela. */
    private static void schedulePair(Context c,String key,String title,long event,int minutes){
        cancel(c,key);
        schedule(c,key+".due",title,event);
        if(minutes>0){
            long early=event-((long)Math.max(0,minutes))*60000L;
            if(early>System.currentTimeMillis()+1200) schedule(c,key+".early",title+" • lembrete antecipado",early);
            else cancel(c,key+".early");
        }else cancel(c,key+".early");
    }
    private static void cancelPair(Context c,String key){
        cancel(c,key);cancel(c,key+".due");cancel(c,key+".early");
    }
    public static void cancelFromOp(Context c,JSONObject item,String uid,String hid){""","alerta na hora e antecipado")

edit(schedule,'        cancel(c,key);\n    }\n    private static void cancel(Context c,String key){','        cancelPair(c,key);\n    }\n    private static void cancel(Context c,String key){',"cancelamento dos dois alarmes")
edit(schedule,'if(when>0)schedule(c,"nc."+uid+"."+hid+"."+section+"."+id,p.optString("title","Compromisso"),alertTime(when,p.optInt("reminder_minutes",30)));',
     'if(when>0)schedulePair(c,"nc."+uid+"."+hid+"."+section+"."+id,p.optString("title","Compromisso"),when,p.optInt("reminder_minutes",30));',"sincronização de mudanças")
edit(schedule,'if(time>0)schedule(c,"nc."+uid+"."+hid+"."+type+".local."+op,p.optString("title","Compromisso"),alertTime(time,before));',
     'if(time>0)schedulePair(c,"nc."+uid+"."+hid+"."+type+".local."+op,p.optString("title","Compromisso"),time,before);',"agendamento offline")
edit(schedule,'cancel(c,"nc."+uid+"."+hid+"."+section+"."+row.optString("id",""));continue;',
     'cancelPair(c,"nc."+uid+"."+hid+"."+section+"."+row.optString("id",""));continue;',"cancelar ao concluir tarefas")
edit(schedule,'schedule(c,key,row.optString("title","Compromisso"),alertTime(time,mins));',
     'schedulePair(c,key,row.optString("title","Compromisso"),time,mins);',"agendar tarefas do servidor")

edit(receiver,'    Intent service=new Intent(context,ReminderAlarmService.class);',"""
    String alertTitle=intent==null?"Compromisso":intent.getStringExtra("title");
    if(alertTitle==null||alertTitle.isEmpty())alertTitle="Compromisso";
    String alertText=intent==null?"Tarefa ou compromisso agendado":intent.getStringExtra("message");
    if(alertText==null||alertText.isEmpty())alertText="Tarefa ou compromisso agendado";
    showDueNotification(context,alertTitle,alertText);
    Intent service=new Intent(context,ReminderAlarmService.class);""","notificação antes de iniciar serviço")

edit(receiver,'\n}',"""
  private static void showDueNotification(Context context,String title,String message){
    try{
      NotificationManager nm=(NotificationManager)context.getSystemService(Context.NOTIFICATION_SERVICE);
      if(nm==null)return;
      final String channelId="nc_due_v218";
      if(Build.VERSION.SDK_INT>=26){
        NotificationChannel ch=new NotificationChannel(channelId,"Tarefas e compromissos",NotificationManager.IMPORTANCE_HIGH);
        ch.setDescription("Mostra notificações de tarefas vencendo, inclusive com aplicativo fechado.");
        ch.setLockscreenVisibility(Notification.VISIBILITY_PRIVATE);ch.enableVibration(true);
        ch.setSound(android.net.Uri.parse("android.resource://"+context.getPackageName()+"/"+R.raw.nc_alarm05),
            new android.media.AudioAttributes.Builder().setUsage(android.media.AudioAttributes.USAGE_ALARM)
            .setContentType(android.media.AudioAttributes.CONTENT_TYPE_SONIFICATION).build());
        nm.createNotificationChannel(ch);
      }
      Intent open=new Intent(context,MainActivity.class).addFlags(Intent.FLAG_ACTIVITY_CLEAR_TOP|Intent.FLAG_ACTIVITY_SINGLE_TOP);
      PendingIntent pi=PendingIntent.getActivity(context,2181,open,PendingIntent.FLAG_UPDATE_CURRENT|PendingIntent.FLAG_IMMUTABLE);
      Notification.Builder note=Build.VERSION.SDK_INT>=26?new Notification.Builder(context,channelId):new Notification.Builder(context);
      note.setSmallIcon(android.R.drawable.ic_lock_idle_alarm).setContentTitle("Nosso Combinado • "+title)
         .setContentText(message).setContentIntent(pi).setAutoCancel(true)
         .setCategory(Notification.CATEGORY_ALARM).setPriority(Notification.PRIORITY_MAX)
         .setVisibility(Notification.VISIBILITY_PRIVATE);
      nm.notify((int)(System.currentTimeMillis()/1000),note.build());
    }catch(Exception ignored){}
  }
}""","aviso persistente na central de notificações")

edit(main,'    private void permissionGuide(){',"""
    private void showAlarmStatus(){
        boolean notify=ReminderScheduler.notificationsAllowed(this);
        boolean exact=ReminderScheduler.exactAllowed(this);
        String current="Notificações: "+(notify?"PERMITIDAS":"BLOQUEADAS")+"\\n"
          +"Alarmes exatos: "+(exact?"PERMITIDOS":"BLOQUEADOS")+"\\n"
          +"Próximo alarme deste aparelho: "+ReminderScheduler.nextAlarm(this)+"\\n\\n"
          +"Novas tarefas criadas no computador precisam chegar ao aplicativo antes do horário. "
          +"Horários já passados não são disparados novamente.";
        dialogBuilder().setTitle("Alarmes e notificações").setMessage(current)
          .setPositiveButton("Sincronizar agora",(d,w)->{
             alarmLastFetch=0;refreshAlarmSchedule();
             msg("Sincronização solicitada. Volte aqui para conferir o próximo horário.");
          }).setNeutralButton("Permissões",(d,w)->{
             if(!notify){
                if(Build.VERSION.SDK_INT>=33&&checkSelfPermission("android.permission.POST_NOTIFICATIONS")!=PackageManager.PERMISSION_GRANTED){
                   ReminderScheduler.requestNotificationPermission(this);
                }else try{
                   Intent i=new Intent(android.provider.Settings.ACTION_APP_NOTIFICATION_SETTINGS)
                       .putExtra(android.provider.Settings.EXTRA_APP_PACKAGE,getPackageName());startActivity(i);
                }catch(Exception ignored){}
             }else if(!exact)ReminderScheduler.openExactSettings(this);
             else try{
                Intent i=new Intent(android.provider.Settings.ACTION_APPLICATION_DETAILS_SETTINGS,
                   android.net.Uri.parse("package:"+getPackageName()));startActivity(i);
             }catch(Exception ignored){}
          }).setNegativeButton("Fechar",null).show();
    }
    private void permissionGuide(){""","diagnóstico real de permissões e agendamento")
edit(main,'        addButton(content,"⚙  Configurar permissões",MINT,NAVY,this::permissionGuide);gap(content,10);',
    '        addButton(content,"⏰  Verificar alarmes e notificações",MINT,NAVY,this::showAlarmStatus);gap(content,10);\n        addButton(content,"⚙  Configurar permissões",MINT,NAVY,this::permissionGuide);gap(content,10);',
    "botão de diagnóstico visível")
edit(main,'handler.removeCallbacks(alarmSyncTicker);handler.postDelayed(alarmSyncTicker,1200);',
     'handler.removeCallbacks(alarmSyncTicker);handler.postDelayed(alarmSyncTicker,1200);\n        ReminderScheduler.restore(this);',"rearmar lembretes ao retomar")
edit(grad,'versionCode 217','versionCode 218',"versionCode")
edit(grad,"versionName '2.1.7'","versionName '2.1.8'","versionName")
print("2.1.8: calendário e horários, sem alterações no mapa/chat/tema.")
