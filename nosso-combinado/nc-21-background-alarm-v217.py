#!/usr/bin/env python3
"""Nosso Combinado 2.1.7 - alarmes registrados/rearmados e um unico controle de som."""
from pathlib import Path
import sys
root=Path(sys.argv[1]).resolve()
pkg=root/'app/src/main/java/com/nossocombinado/app'
main=pkg/'MainActivity.java'
scheduler=pkg/'ReminderScheduler.java'
receiver=pkg/'ReminderReceiver.java'
service=pkg/'ReminderAlarmService.java'
restart=pkg/'RestartReceiver.java'
manifest=root/'app/src/main/AndroidManifest.xml'
grad=root/'app/build.gradle'
def edit(file, old, new, description):
    s=file.read_text(encoding='utf-8')
    n=s.count(old)
    if n!=1:raise RuntimeError(description+' expected once, got '+str(n))
    file.write_text(s.replace(old,new,1),encoding='utf-8')
    print('OK',description)

# Persist scheduled wakeups on device. AlarmManager itself survives app process death.
edit(scheduler,'    private ReminderScheduler(){}',r'''    private ReminderScheduler(){}
    private static final String REGISTRY="nc_pending_alarm_v217";
    private static synchronized JSONObject records(Context ctx){
        try{return new JSONObject(ctx.getSharedPreferences(REGISTRY,Context.MODE_PRIVATE).getString("items","{}"));}
        catch(Exception e){return new JSONObject();}
    }
    private static synchronized void remember(Context ctx,String key,String title,long when){
        try{
            JSONObject all=records(ctx);
            all.put(key,new JSONObject().put("title",title).put("at",when));
            ctx.getSharedPreferences(REGISTRY,Context.MODE_PRIVATE).edit().putString("items",all.toString()).commit();
        }catch(Exception ignored){}
    }
    public static synchronized void forget(Context ctx,String key){
        if(key==null||key.isEmpty())return;
        JSONObject all=records(ctx);all.remove(key);
        ctx.getSharedPreferences(REGISTRY,Context.MODE_PRIVATE).edit().putString("items",all.toString()).commit();
    }
    public static synchronized void restore(Context ctx){
        JSONObject all=records(ctx);
        java.util.Iterator<String> keys=all.keys();
        java.util.ArrayList<String> toRestore=new java.util.ArrayList<>();
        while(keys.hasNext())toRestore.add(keys.next());
        long now=System.currentTimeMillis();
        for(String key:toRestore){
            JSONObject item=all.optJSONObject(key);
            if(item==null){forget(ctx,key);continue;}
            long when=item.optLong("at");
            if(when<=now+1200){forget(ctx,key);continue;}
            schedule(ctx,key,item.optString("title","Compromisso"),when);
        }
    }''','registry for pending alarms')

edit(scheduler,'        if(when<=System.currentTimeMillis()+1200)return;',
'''        if(when<=System.currentTimeMillis()+1200){forget(c,key);return;}
        remember(c,key,title,when);''','remember alarm prior to Android scheduling')
edit(scheduler,'''    private static void cancel(Context c,String key){
       Intent i=new Intent(c,ReminderReceiver.class);''',
'''    private static void cancel(Context c,String key){
       forget(c,key);
       Intent i=new Intent(c,ReminderReceiver.class);''','forget cancelled alarms')

# Receiver continues to work when no activity/process is open.
edit(receiver,'''    Intent service=new Intent(context,ReminderAlarmService.class);''',
'''    if(intent!=null)ReminderScheduler.forget(context,intent.getStringExtra("key"));
    Intent service=new Intent(context,ReminderAlarmService.class);''','do not repeat delivered alarm on boot')
edit(receiver,'''      ReminderScheduler.ensureChannel(context);
      NotificationManager nm=(NotificationManager)context.getSystemService(Context.NOTIFICATION_SERVICE);
      if(nm!=null){
         Notification.Builder n=Build.VERSION.SDK_INT>=26?
            new Notification.Builder(context,ReminderScheduler.CHANNEL):new Notification.Builder(context);''',
'''      // Audible notification fallback if Android prohibits background foreground-service start.
      NotificationManager nm=(NotificationManager)context.getSystemService(Context.NOTIFICATION_SERVICE);
      final String fallbackChannel="nc_alarm_fallback_v217";
      if(nm!=null){
         if(Build.VERSION.SDK_INT>=26){
            android.app.NotificationChannel channel=new android.app.NotificationChannel(
                 fallbackChannel,"Alertas de compromisso",NotificationManager.IMPORTANCE_HIGH);
            channel.setSound(android.net.Uri.parse("android.resource://"+context.getPackageName()+"/"+R.raw.nc_alarm05),
                new android.media.AudioAttributes.Builder().setUsage(android.media.AudioAttributes.USAGE_ALARM)
                  .setContentType(android.media.AudioAttributes.CONTENT_TYPE_SONIFICATION).build());
            channel.enableVibration(true);nm.createNotificationChannel(channel);
         }
         Notification.Builder n=Build.VERSION.SDK_INT>=26?
            new Notification.Builder(context,fallbackChannel):new Notification.Builder(context);''','audible fallback when background service restricted')

# Keep CPU up for playback, with automatic timeout and release.
edit(service,'    private MediaPlayer player;',
'''    private MediaPlayer player;
    private android.os.PowerManager.WakeLock wakeLock;''','wake lock field')
edit(service,'        stopAudio();\n        try{',
'''        stopAudio();
        try{
            android.os.PowerManager pm=(android.os.PowerManager)getSystemService(POWER_SERVICE);
            if(pm!=null){wakeLock=pm.newWakeLock(android.os.PowerManager.PARTIAL_WAKE_LOCK,"NossoCombinado:Alarme");
                wakeLock.acquire(185000L);}
        }catch(Exception ignored){}
        try{''','keep processor awake while playing')
edit(service,'''    @Override public void onDestroy(){handler.removeCallbacksAndMessages(null);stopAudio();super.onDestroy();}''',
'''    @Override public void onDestroy(){
        handler.removeCallbacksAndMessages(null);stopAudio();
        if(wakeLock!=null){try{if(wakeLock.isHeld())wakeLock.release();}catch(Exception ignored){}wakeLock=null;}
        super.onDestroy();
    }''','release wake lock')

# Restore after a reboot, package replacement and granting exact-alarm permission.
edit(restart,'''  if(!Intent.ACTION_BOOT_COMPLETED.equals(i.getAction()))return;''',
'''  if(i==null)return;
  String action=i.getAction();
  if(!Intent.ACTION_BOOT_COMPLETED.equals(action)&&
     !Intent.ACTION_MY_PACKAGE_REPLACED.equals(action)&&
     !"android.app.action.SCHEDULE_EXACT_ALARM_PERMISSION_STATE_CHANGED".equals(action))return;
  ReminderScheduler.restore(context);
  SyncJobService.start(context);''','restore across boot/update/permission')
edit(restart,'''  }catch(Exception ignored){}
 }
}''',
'''  }catch(Exception ignored){}
 }
 /** Rehydrate older-version dashboard reminders while offline. */
 public static void recoverCached(Context context){
  try{
    SecureStore s=new SecureStore(context);String uid=s.get("user_id"),hid=s.get("hid");
    if(uid.isEmpty()||hid.isEmpty())return;
    OfflineDb db=new OfflineDb(context,s);
    for(String page:new String[]{"home","calendar","tasks"}){
      String raw=db.load(uid,hid,page);
      if(!raw.isEmpty())ReminderScheduler.fromDashboard(context,new JSONObject(raw),uid,hid);
    }
  }catch(Exception ignored){}
 }
}''','offline restore helper')
edit(restart,'''  try{SecureStore s=new SecureStore(context);String uid=s.get("user_id"),hid=s.get("hid");if(uid.isEmpty()||hid.isEmpty())return;
    OfflineDb db=new OfflineDb(context,s);
    String snapshot=db.load(uid,hid,"calendar");if(snapshot.isEmpty())snapshot=db.load(uid,hid,"home");
    if(!snapshot.isEmpty())ReminderScheduler.fromDashboard(context,new JSONObject(snapshot),uid,hid);
    snapshot=db.load(uid,hid,"tasks");if(!snapshot.isEmpty())ReminderScheduler.fromDashboard(context,new JSONObject(snapshot),uid,hid);''',
'''  try{recoverCached(context);''','reuse cached alarm recovery')

edit(manifest,'''    <uses-permission android:name="android.permission.FOREGROUND_SERVICE_MEDIA_PLAYBACK"/>''',
'''    <uses-permission android:name="android.permission.FOREGROUND_SERVICE_MEDIA_PLAYBACK"/>
    <uses-permission android:name="android.permission.WAKE_LOCK"/>''','CPU wake lock permission')
edit(manifest,'''<intent-filter><action android:name="android.intent.action.BOOT_COMPLETED" /></intent-filter>''',
'''<intent-filter>
                <action android:name="android.intent.action.BOOT_COMPLETED" />
                <action android:name="android.intent.action.MY_PACKAGE_REPLACED" />
                <action android:name="android.app.action.SCHEDULE_EXACT_ALARM_PERMISSION_STATE_CHANGED" />
            </intent-filter>''','Android recovery broadcasts')

# Restore saved schedules in updated apps without requiring a network connection.
edit(main,'''ReminderScheduler.ensureChannel(this);SyncJobService.start(this);handler.postDelayed(periodic,25000);''',
'''ReminderScheduler.ensureChannel(this);ReminderScheduler.restore(this);
        RestartReceiver.recoverCached(this);SyncJobService.start(this);handler.postDelayed(periodic,25000);''',
'rearm at launch with cached reminders')

# One clear setting for INTERACTION sounds. It must not mute the alarm clock.
edit(main,'''addButton(content,"Som no app: "+("off".equals(vault.get("sound_effects"))?"desativado · ativar":"ativado · desativar"),MINT,NAVY,()->{vault.set("sound_effects","off".equals(vault.get("sound_effects"))?"on":"off");showMore();});gap(content,10);''',
'''LinearLayout soundPanel=card(content);
        android.widget.Switch soundSwitch=new android.widget.Switch(this);
        boolean enabled=!"off".equals(vault.get("sound_effects"));
        soundSwitch.setChecked(enabled);
        soundSwitch.setText(enabled?"Sons de interações: ATIVADOS":"Sons de interações: DESATIVADOS");
        soundSwitch.setTextColor(NAVY);soundSwitch.setTextSize(15);
        soundSwitch.setContentDescription("Ativar ou desativar sons de mensagens e interações");
        soundSwitch.setPadding(dp(8),dp(8),dp(8),dp(8));
        soundSwitch.setOnCheckedChangeListener((button,checked)->{
            vault.set("sound_effects",checked?"on":"off");
            soundSwitch.setText(checked?"Sons de interações: ATIVADOS":"Sons de interações: DESATIVADOS");
        });
        soundPanel.addView(soundSwitch,new LinearLayout.LayoutParams(-1,dp(52)));
        caption(soundPanel,"Alarmes de tarefas e compromissos permanecem ativos.");
        gap(content,10);''','single clear interaction-sound switch')

edit(grad,'versionCode 216','versionCode 217','version code')
edit(grad,"versionName '2.1.6'","versionName '2.1.7'",'version name')
print("Android 2.1.7 - layouts and offline database preserved.")
