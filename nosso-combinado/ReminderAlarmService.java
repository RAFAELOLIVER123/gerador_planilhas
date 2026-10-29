package com.nossocombinado.app;
import android.app.*;
import android.content.*;
import android.media.AudioAttributes;
import android.media.MediaPlayer;
import android.net.Uri;
import android.os.Build;
import android.os.Handler;
import android.os.IBinder;
import android.os.Looper;

/** Plays the unchanged original nc_alarm05 until dismissed or after three minutes. */
public final class ReminderAlarmService extends Service {
    private static final String CHANNEL="nc_alarm_loop_21_1";
    private static final int ID=2211;
    private final Handler handler=new Handler(Looper.getMainLooper());
    private MediaPlayer player;
    @Override public IBinder onBind(Intent i){return null;}
    @Override public int onStartCommand(Intent data,int flags,int startId){
        String title=data==null?null:data.getStringExtra("title");
        if(title==null||title.isEmpty())title="Compromisso da família";
        String message=data==null?null:data.getStringExtra("message");
        if(message==null||message.isEmpty())message="Tarefa ou compromisso agendado";
        NotificationManager nm=(NotificationManager)getSystemService(Context.NOTIFICATION_SERVICE);
        if(Build.VERSION.SDK_INT>=26&&nm!=null){
            NotificationChannel channel=new NotificationChannel(CHANNEL,"Alarme do Nosso Combinado",NotificationManager.IMPORTANCE_HIGH);
            channel.setSound(null,null);
            channel.enableVibration(true);
            nm.createNotificationChannel(channel);
        }
        Intent off=new Intent(this,ReminderReceiver.class).setAction(ReminderReceiver.ACTION_STOP);
        PendingIntent stop=PendingIntent.getBroadcast(this,2211,off,PendingIntent.FLAG_UPDATE_CURRENT|PendingIntent.FLAG_IMMUTABLE);
        Intent open=new Intent(this,MainActivity.class).addFlags(Intent.FLAG_ACTIVITY_CLEAR_TOP|Intent.FLAG_ACTIVITY_SINGLE_TOP);
        PendingIntent app=PendingIntent.getActivity(this,2212,open,PendingIntent.FLAG_UPDATE_CURRENT|PendingIntent.FLAG_IMMUTABLE);
        Notification.Builder builder=Build.VERSION.SDK_INT>=26?new Notification.Builder(this,CHANNEL):new Notification.Builder(this);
        builder.setSmallIcon(android.R.drawable.ic_lock_idle_alarm).setContentTitle("⏰ Nosso Combinado • "+title)
            .setContentText(message).setOngoing(true).setContentIntent(app)
            .setCategory(Notification.CATEGORY_ALARM).setVisibility(Notification.VISIBILITY_PRIVATE)
            .setPriority(Notification.PRIORITY_HIGH)
            .addAction(android.R.drawable.ic_menu_close_clear_cancel,"DESLIGAR",stop);
        try{startForeground(ID,builder.build());}
        catch(Exception denied){stopSelf();return START_NOT_STICKY;}
        stopAudio();
        try{
            player=new MediaPlayer();
            player.setAudioAttributes(new AudioAttributes.Builder()
              .setUsage(AudioAttributes.USAGE_ALARM)
              .setContentType(AudioAttributes.CONTENT_TYPE_SONIFICATION).build());
            Uri sound=Uri.parse("android.resource://"+getPackageName()+"/"+R.raw.nc_alarm05);
            player.setDataSource(this,sound);
            player.setLooping(true);player.prepare();player.start();
        }catch(Exception error){stopAudio();}
        handler.removeCallbacksAndMessages(null);
        handler.postDelayed(this::stopSelf,180000);
        return START_NOT_STICKY;
    }
    private void stopAudio(){
        if(player!=null){try{if(player.isPlaying())player.stop();player.release();}
            catch(Exception ignored){}player=null;}
    }
    @Override public void onDestroy(){handler.removeCallbacksAndMessages(null);stopAudio();super.onDestroy();}
}
