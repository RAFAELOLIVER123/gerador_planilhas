package com.nossocombinado.app;

import android.app.AlarmManager;
import android.app.Notification;
import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.app.PendingIntent;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.media.AudioAttributes;
import android.net.Uri;
import android.os.Build;
import android.os.SystemClock;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;
import java.util.TimeZone;

/** Lembrete da lista: agendado no Android; não depende da tela estar aberta. */
public final class NcListAlarmReceiver extends BroadcastReceiver {
    private static final String CHANNEL="nc_list_alarm_214";
    private static final String STORAGE="nc_list_alarm_store";
    private static final String ACTION="com.nossocombinado.app.NATIVE_LIST_ALARM";

    @Override public void onReceive(Context context,Intent intent) {
        String action=intent==null?"":intent.getAction();
        if(Intent.ACTION_BOOT_COMPLETED.equals(action)||Intent.ACTION_MY_PACKAGE_REPLACED.equals(action)){
            restore(context);return;
        }
        if(!ACTION.equals(action))return;
        long id=intent.getLongExtra("id",0);
        String title=intent.getStringExtra("title");
        if(title==null||title.trim().isEmpty())title="Lista de compras";
        NotificationManager manager=(NotificationManager)context.getSystemService(Context.NOTIFICATION_SERVICE);
        Uri sound=Uri.parse("android.resource://"+context.getPackageName()+"/raw/nc_alarm05");
        if(Build.VERSION.SDK_INT>=26) {
            NotificationChannel channel=new NotificationChannel(CHANNEL,"Alarmes de listas",
                 NotificationManager.IMPORTANCE_HIGH);
            AudioAttributes soundSettings=new AudioAttributes.Builder()
               .setUsage(AudioAttributes.USAGE_ALARM)
               .setContentType(AudioAttributes.CONTENT_TYPE_SONIFICATION).build();
            channel.setSound(sound,soundSettings);
            channel.enableVibration(true);
            manager.createNotificationChannel(channel);
        }
        Intent open=new Intent(context,MainActivity.class);
        open.setFlags(Intent.FLAG_ACTIVITY_NEW_TASK|Intent.FLAG_ACTIVITY_CLEAR_TOP);
        PendingIntent go=PendingIntent.getActivity(context, (int)(id%100000),
             open,PendingIntent.FLAG_UPDATE_CURRENT|PendingIntent.FLAG_IMMUTABLE);
        Notification.Builder builder=Build.VERSION.SDK_INT>=26
             ?new Notification.Builder(context,CHANNEL):new Notification.Builder(context);
        builder.setSmallIcon(R.drawable.nosso_icon)
           .setContentTitle("⏰ Hora do combinado")
           .setContentText("Compras: "+title)
           .setContentIntent(go)
           .setAutoCancel(true)
           .setCategory(Notification.CATEGORY_ALARM)
           .setVisibility(Notification.VISIBILITY_PRIVATE)
           .setPriority(Notification.PRIORITY_HIGH);
        if(Build.VERSION.SDK_INT<26)builder.setSound(sound);
        try {manager.notify(6000+(int)(id%100000),builder.build());}catch(SecurityException ignored){}
    }

    private static PendingIntent intent(Context context,long id,String title){
        Intent broadcast=new Intent(context,NcListAlarmReceiver.class);
        broadcast.setAction(ACTION);
        broadcast.putExtra("id",id);
        broadcast.putExtra("title",title);
        return PendingIntent.getBroadcast(context,(int)(id%100000),
             broadcast,PendingIntent.FLAG_IMMUTABLE|PendingIntent.FLAG_UPDATE_CURRENT);
    }

    public static void schedule(Context context,long id,String title,String at){
        try{
            SimpleDateFormat fmt=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss",Locale.US);
            fmt.setTimeZone(TimeZone.getTimeZone("America/Belem"));
            fmt.setLenient(false);
            Date parsed=fmt.parse(at);
            if(parsed==null||parsed.getTime()<=System.currentTimeMillis())return;
            AlarmManager alarms=(AlarmManager)context.getSystemService(Context.ALARM_SERVICE);
            PendingIntent pending=intent(context,id,title);
            if(Build.VERSION.SDK_INT>=31&&alarms.canScheduleExactAlarms()) {
                alarms.setExactAndAllowWhileIdle(AlarmManager.RTC_WAKEUP,parsed.getTime(),pending);
            } else if(Build.VERSION.SDK_INT>=23) {
                alarms.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP,parsed.getTime(),pending);
            } else alarms.set(AlarmManager.RTC_WAKEUP,parsed.getTime(),pending);
            SharedPreferences store=context.getSharedPreferences(STORAGE,Context.MODE_PRIVATE);
            Set<String> ids=new HashSet<>(store.getStringSet("ids",new HashSet<>()));
            ids.add(String.valueOf(id));
            store.edit().putStringSet("ids",ids).putString("title_"+id,title).putString("at_"+id,at).apply();
        }catch(Exception ignored){}
    }

    public static void cancel(Context context,long id){
        try {
            AlarmManager alarms=(AlarmManager)context.getSystemService(Context.ALARM_SERVICE);
            alarms.cancel(intent(context,id,""));
            SharedPreferences store=context.getSharedPreferences(STORAGE,Context.MODE_PRIVATE);
            Set<String> ids=new HashSet<>(store.getStringSet("ids",new HashSet<>()));
            ids.remove(String.valueOf(id));
            store.edit().putStringSet("ids",ids).remove("title_"+id).remove("at_"+id).apply();
        }catch(Exception ignored){}
    }

    public static void restore(Context context){
        SharedPreferences store=context.getSharedPreferences(STORAGE,Context.MODE_PRIVATE);
        Set<String> ids=new HashSet<>(store.getStringSet("ids",new HashSet<>()));
        for(String value:ids){
            try {
                long id=Long.parseLong(value);
                schedule(context,id,store.getString("title_"+id,"Lista"),
                        store.getString("at_"+id,""));
            }catch(Exception ignored){}
        }
    }
}
