#!/usr/bin/env python3
"""Android 2.1.6: altera somente os alarmes sobre a versão original 2.1.5."""
from pathlib import Path
import sys
root=Path(sys.argv[1]).resolve()
folder=root/'app/src/main/java/com/nossocombinado/app'
main=folder/'MainActivity.java'
scheduler=folder/'ReminderScheduler.java'
service=folder/'ReminderAlarmService.java'
grad=root/'app/build.gradle'
def replace_once(file,old,new,name):
 s=file.read_text(encoding='utf-8')
 n=s.count(old)
 if n!=1:raise RuntimeError(name+': esperado um trecho; encontrados '+str(n))
 file.write_text(s.replace(old,new,1),encoding='utf-8')
 print('OK:',name)

replace_once(main,'    private interface Callback {void done(JSONObject j,String error);}', '''    /** Agenda compromissos independentemente da tela ativa. */
    private boolean alarmSyncBusy=false;
    private long alarmLastFetch=0;
    private final Runnable alarmSyncTicker=new Runnable(){@Override public void run(){
        refreshAlarmSchedule();
        handler.postDelayed(this,90000);
    }};
    private void refreshAlarmSchedule(){
        if(api==null||userId.isEmpty()||householdId.isEmpty()||alarmSyncBusy)return;
        long now=android.os.SystemClock.elapsedRealtime();
        if(now-alarmLastFetch<25000)return;
        alarmLastFetch=now;
        alarmSyncBusy=true;
        final String uid=userId,hid=householdId;
        api.get("dashboard",data("household_id",hid),(j,error)->{
            alarmSyncBusy=false;
            if(error!=null||j==null||!uid.equals(userId)||!hid.equals(householdId))return;
            offline.save(uid,hid,"calendar",j.toString());
            ReminderScheduler.fromDashboard(this,j,uid,hid);
        });
    }
    private interface Callback {void done(JSONObject j,String error);}''','agendamento também enquanto chat e outras telas abertas')
replace_once(main,'handler.postDelayed(realtimeChat,350);if(map!=null)map.onResume();',
'''handler.postDelayed(realtimeChat,350);
        handler.removeCallbacks(alarmSyncTicker);handler.postDelayed(alarmSyncTicker,1200);
        if(map!=null)map.onResume();''','buscar alarmes ao abrir')
replace_once(main,'@Override protected void onPause(){handler.removeCallbacks(realtimeChat);',
'@Override protected void onPause(){handler.removeCallbacks(alarmSyncTicker);handler.removeCallbacks(realtimeChat);','não pesquisar em segundo plano')
replace_once(main,'handler.removeCallbacks(periodic);handler.removeCallbacks(realtimeChat);',
'handler.removeCallbacks(periodic);handler.removeCallbacks(realtimeChat);handler.removeCallbacks(alarmSyncTicker);','limpar consulta')
replace_once(main,'    private void chooseHousehold(JSONObject h){householdId=val(h,"id");householdName=val(h,"name");role=val(h,"role");vault.set("hid",householdId);vault.set("household_name",householdName);vault.set("role",role);}',
'''    private void chooseHousehold(JSONObject h){
        String prior=householdId;
        householdId=val(h,"id");householdName=val(h,"name");role=val(h,"role");
        vault.set("hid",householdId);vault.set("household_name",householdName);vault.set("role",role);
        if(!householdId.equals(prior))alarmLastFetch=0;
        refreshAlarmSchedule();
    }''','lembrar compromissos da família selecionada')
replace_once(scheduler,
'''            if(Build.VERSION.SDK_INT>=31&&am.canScheduleExactAlarms())am.setExactAndAllowWhileIdle(AlarmManager.RTC_WAKEUP,when,pi);
            else if(Build.VERSION.SDK_INT>=23&&Build.VERSION.SDK_INT<31)am.setExactAndAllowWhileIdle(AlarmManager.RTC_WAKEUP,when,pi);
            else if(Build.VERSION.SDK_INT>=23)am.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP,when,pi);
            else am.set(AlarmManager.RTC_WAKEUP,when,pi);
        }catch(SecurityException ex){am.set(AlarmManager.RTC_WAKEUP,when,pi);}''',
'''            if(Build.VERSION.SDK_INT>=23){
                Intent screen=new Intent(c,MainActivity.class);
                screen.addFlags(Intent.FLAG_ACTIVITY_CLEAR_TOP|Intent.FLAG_ACTIVITY_SINGLE_TOP);
                PendingIntent show=PendingIntent.getActivity(c,1801,screen,PendingIntent.FLAG_UPDATE_CURRENT|PendingIntent.FLAG_IMMUTABLE);
                am.setAlarmClock(new AlarmManager.AlarmClockInfo(when,show),pi);
            }else am.setExact(AlarmManager.RTC_WAKEUP,when,pi);
        }catch(SecurityException ex){
            if(Build.VERSION.SDK_INT>=23)am.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP,when,pi);
            else am.set(AlarmManager.RTC_WAKEUP,when,pi);
        }''','despertador exato pelo AlarmClock')
replace_once(service,'private static final String CHANNEL="nc_alarm_loop_21_1";',
'private static final String CHANNEL="nc_alarm_loop_21_6";','canal do alarme dedicado')
replace_once(service,'''            player=new MediaPlayer();
            player.setAudioAttributes(new AudioAttributes.Builder()
              .setUsage(AudioAttributes.USAGE_ALARM)
              .setContentType(AudioAttributes.CONTENT_TYPE_SONIFICATION).build());
            Uri sound=Uri.parse("android.resource://"+getPackageName()+"/"+R.raw.nc_alarm05);
            player.setDataSource(this,sound);
            player.setLooping(true);player.prepare();player.start();
        }catch(Exception error){stopAudio();}''',
'''            player=new MediaPlayer();
            player.setAudioAttributes(new AudioAttributes.Builder()
              .setUsage(AudioAttributes.USAGE_ALARM)
              .setContentType(AudioAttributes.CONTENT_TYPE_SONIFICATION).build());
            try(android.content.res.AssetFileDescriptor fd=getResources().openRawResourceFd(R.raw.nc_alarm05)){
                player.setDataSource(fd.getFileDescriptor(),fd.getStartOffset(),fd.getLength());
            }
            player.setLooping(true);player.prepare();player.start();
        }catch(Exception error){
            stopAudio();
            try{
                Uri fallback=android.media.RingtoneManager.getDefaultUri(android.media.RingtoneManager.TYPE_ALARM);
                if(fallback==null)fallback=android.media.RingtoneManager.getDefaultUri(android.media.RingtoneManager.TYPE_NOTIFICATION);
                if(fallback!=null){
                    player=new MediaPlayer();
                    player.setAudioAttributes(new AudioAttributes.Builder().setUsage(AudioAttributes.USAGE_ALARM).setContentType(AudioAttributes.CONTENT_TYPE_SONIFICATION).build());
                    player.setDataSource(this,fallback);player.setLooping(true);player.prepare();player.start();
                }
            }catch(Exception ignored){stopAudio();}
        }''','WAV original repetido e toque de reserva')
replace_once(main,'@Override protected void onResume(){super.onResume();handler.removeCallbacks(realtimeChat);',
'''@Override protected void onResume(){super.onResume();
        if(!userId.isEmpty()&&android.os.Build.VERSION.SDK_INT>=33&&
           checkSelfPermission("android.permission.POST_NOTIFICATIONS")!=android.content.pm.PackageManager.PERMISSION_GRANTED){
            android.content.SharedPreferences prefs=getSharedPreferences("nc_alarm_216",MODE_PRIVATE);
            if(!prefs.getBoolean("notification_permission_asked",false)){
                prefs.edit().putBoolean("notification_permission_asked",true).apply();
                ReminderScheduler.requestNotificationPermission(this);
            }
        }
        handler.removeCallbacks(realtimeChat);''','pedir permissão uma única vez')
replace_once(grad,'versionCode 215','versionCode 216','versão código')
replace_once(grad,"versionName '2.1.5'","versionName '2.1.6'",'versão')
print('OK: 2.1.6 preserva seletor de família, ferramentas e sons originais.')
