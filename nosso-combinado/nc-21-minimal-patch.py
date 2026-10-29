#!/usr/bin/env python3
"""Only message refresh and alarm playback changes on original 2.1 sources."""
from pathlib import Path
import sys
root=Path(sys.argv[1])
d=root/'app/src/main/java/com/nossocombinado/app'
main=d/'MainActivity.java'
grad=root/'app/build.gradle'
manifest=root/'app/src/main/AndroidManifest.xml'
def sw(p,a,b,label):
 s=p.read_text(encoding='utf-8')
 assert s.count(a)==1,(label,s.count(a))
 p.write_text(s.replace(a,b,1),encoding='utf-8')
 print('OK',label)
sw(main,'if(!syncing && (page.equals("chat")||page.equals("messages")||page.equals("alerts")','if(!syncing && (page.equals("alerts")','avoid full chat rerender by maintenance')
sw(main,'private interface Callback {void done(JSONObject j,String error);}',"""private boolean realtimeChatBusy=false;
    private final Runnable realtimeChat=new Runnable(){@Override public void run(){
        if(!realtimeChatBusy&&!userId.isEmpty()&&!householdId.isEmpty()&&
           (page.equals("chat")||page.equals("messages")||page.equals("private_thread"))){
            String current=page,uid=userId,hid=householdId;long cid=privateCid;
            String action=current.equals("chat")?"chat_list":current.equals("messages")?"private_chat_list":"private_chat_messages";
            JSONObject params=current.equals("private_thread")?
              data("household_id",hid,"conversation_id",String.valueOf(cid)):withHid(new JSONObject());
            String cacheKey=current.equals("private_thread")?"dm_"+cid:current;
            realtimeChatBusy=true;
            api.get(action,params,(j,error)->{
                realtimeChatBusy=false;
                if(error!=null||j==null||!current.equals(page)||!uid.equals(userId)||
                  !hid.equals(householdId)||(current.equals("private_thread")&&cid!=privateCid))return;
                String before=offline.load(uid,hid,cacheKey),after=j.toString();
                if(after.equals(before))return;
                offline.save(uid,hid,cacheKey,after);
                ncCheckSounds(action,j);
                if(scroll==null||content==null)return;
                final int oldY=scroll.getScrollY();
                final boolean nearBottom=scroll.getChildAt(0)!=null&&
                  scroll.getChildAt(0).getHeight()-scroll.getHeight()-oldY<110;
                if(current.equals("private_thread")){
                    String draft=activeDmInput==null?"":activeDmInput.getText().toString();
                    boolean focus=activeDmInput!=null&&activeDmInput.hasFocus();
                    content.removeAllViews();
                    sectionTitle("Chat com "+privateTitle,"Conversa privada entre integrantes.");
                    addButton(content,"← Voltar",MINT,NAVY,()->navigate("messages"));gap(content,12);
                    renderPrivateMessages(j);
                    if(activeDmInput!=null){activeDmInput.setText(draft);
                      activeDmInput.setSelection(activeDmInput.length());
                      if(focus)activeDmInput.requestFocus();}
                }else drawPage(j,false);
                scroll.post(()->{
                    if(scroll==null||scroll.getChildAt(0)==null)return;
                    int max=Math.max(0,scroll.getChildAt(0).getHeight()-scroll.getHeight());
                    scroll.scrollTo(0,nearBottom?max:Math.min(oldY,max));
                });
            });
        }
        handler.postDelayed(this,1200);
    }};
    private interface Callback {void done(JSONObject j,String error);}""",'fast message polling without regenerating same UI')
sw(main,'if("chat_list".equals(action)||"private_chat_list".equals(action)){field="items";kind=R.raw.nc_mensagem;}',
  'if("chat_list".equals(action)||"private_chat_list".equals(action)||"private_chat_messages".equals(action)){field="items";kind=R.raw.nc_mensagem;}','sound when private messages arrive')
sw(main,'@Override protected void onResume(){super.onResume();if(map!=null)map.onResume();',
 '@Override protected void onResume(){super.onResume();handler.removeCallbacks(realtimeChat);handler.postDelayed(realtimeChat,350);if(map!=null)map.onResume();','start polling while active')
sw(main,'@Override protected void onPause(){super.onPause();if(map!=null)map.onPause();}',
 '@Override protected void onPause(){handler.removeCallbacks(realtimeChat);super.onPause();if(map!=null)map.onPause();}','stop polling in background')
sw(main,'handler.removeCallbacks(periodic);if(map!=null)', 'handler.removeCallbacks(periodic);handler.removeCallbacks(realtimeChat);if(map!=null)','cleanup callbacks')
sw(main,'EditText message=edit("Escreva sua mensagem...",false);message.setSingleLine(false);message.setMaxLines(5);',
 'EditText message=edit("Escreva sua mensagem...",false);message.setSingleLine(false);message.setMaxLines(5);activeDmInput=message;', 'retain the typing draft')
(d/'ReminderReceiver.java').write_text("""package com.nossocombinado.app;
import android.app.*;
import android.content.*;
import android.os.Build;
/** Original 2.1 alarm receiver; it now plays the built-in alarm until dismissed. */
public class ReminderReceiver extends android.content.BroadcastReceiver {
  public static final String ACTION_STOP="com.nossocombinado.app.REMINDER_STOP";
  @Override public void onReceive(Context context, Intent intent){
    if(intent!=null&&ACTION_STOP.equals(intent.getAction())){
      context.stopService(new Intent(context,ReminderAlarmService.class));return;
    }
    Intent service=new Intent(context,ReminderAlarmService.class);
    if(intent!=null){service.putExtra("title",intent.getStringExtra("title"));
       service.putExtra("message",intent.getStringExtra("message"));}
    try{
      if(Build.VERSION.SDK_INT>=26)context.startForegroundService(service);
      else context.startService(service);
    }catch(Exception denied){
      ReminderScheduler.ensureChannel(context);
      NotificationManager nm=(NotificationManager)context.getSystemService(Context.NOTIFICATION_SERVICE);
      if(nm!=null){
         Notification.Builder n=Build.VERSION.SDK_INT>=26?
            new Notification.Builder(context,ReminderScheduler.CHANNEL):new Notification.Builder(context);
         n.setSmallIcon(android.R.drawable.ic_popup_reminder).setContentTitle("Nosso Combinado • compromisso")
          .setContentText("Lembrete agendado").setCategory(Notification.CATEGORY_ALARM)
          .setPriority(Notification.PRIORITY_MAX);
         nm.notify(2211,n.build());
      }
    }
  }
}
""",encoding='utf-8')
ms=manifest.read_text(encoding='utf-8')
assert '<application' in ms
if 'android.permission.FOREGROUND_SERVICE' not in ms:
 ms=ms.replace('<application','<uses-permission android:name="android.permission.FOREGROUND_SERVICE"/>\n    <uses-permission android:name="android.permission.FOREGROUND_SERVICE_MEDIA_PLAYBACK"/>\n    <application',1)
assert '<activity ' in ms
ms=ms.replace('<activity ','<service android:name=".ReminderAlarmService" android:exported="false" android:foregroundServiceType="mediaPlayback"/>\n        <activity ',1)
manifest.write_text(ms,encoding='utf-8')
sw(grad,'versionCode 210','versionCode 211','version code for bugfix')
sw(grad,"versionName '2.1.0'","versionName '2.1.1'",'bugfix version name')
print("Same 2.1 original identity, map, layout, assets, API and offline DB.")
