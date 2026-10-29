package com.nossocombinado.app;
import android.app.Service;
import android.app.Notification;
import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.app.PendingIntent;
import android.content.Intent;
import android.media.AudioAttributes;
import android.media.MediaPlayer;
import android.media.RingtoneManager;
import android.net.Uri;
import android.os.Build;
import android.os.Handler;
import android.os.IBinder;
import android.os.Looper;

public final class AlarmService extends Service{
  private static final String CHANNEL="nc_compromissos_26";
  private final Handler handler=new Handler(Looper.getMainLooper());
  private MediaPlayer player;
  @Override public IBinder onBind(Intent i){return null;}
  @Override public int onStartCommand(Intent intent,int flags,int startId){
    String title=intent==null?"Compromisso da família":intent.getStringExtra("title");
    if(title==null||title.trim().isEmpty())title="Compromisso da família";
    title=title.length()>120?title.substring(0,120):title;
    NotificationManager nm=(NotificationManager)getSystemService(NOTIFICATION_SERVICE);
    if(Build.VERSION.SDK_INT>=26){
      NotificationChannel channel=new NotificationChannel(CHANNEL,"Alarmes de tarefas e compromissos",NotificationManager.IMPORTANCE_HIGH);
      channel.setDescription("Alerta de compromisso no horário agendado");
      channel.setSound(null,null);nm.createNotificationChannel(channel);
    }
    Intent dismiss=new Intent(this,AlarmReceiver.class).setAction("com.nossocombinado.app.native28.STOP");
    PendingIntent off=PendingIntent.getBroadcast(this,9111,dismiss,PendingIntent.FLAG_UPDATE_CURRENT|PendingIntent.FLAG_IMMUTABLE);
    Intent open=new Intent(this,AlarmActivity.class).putExtra("title",title);
    PendingIntent screen=PendingIntent.getActivity(this,9112,open,PendingIntent.FLAG_UPDATE_CURRENT|PendingIntent.FLAG_IMMUTABLE);
    Notification.Builder builder=Build.VERSION.SDK_INT>=26?new Notification.Builder(this,CHANNEL):new Notification.Builder(this);
    builder.setSmallIcon(android.R.drawable.ic_lock_idle_alarm)
      .setContentTitle("⏰ Nosso Combinado").setContentText(title)
      .setContentIntent(screen).setAutoCancel(false)
      .setOngoing(true).setCategory(Notification.CATEGORY_ALARM)
      .setPriority(Notification.PRIORITY_HIGH)
      .addAction(android.R.drawable.ic_menu_close_clear_cancel,"DESLIGAR",off);
    if(Build.VERSION.SDK_INT<26)builder.setDefaults(Notification.DEFAULT_VIBRATE);
    try{startForeground(261,builder.build());}catch(Exception e){stopSelf();return START_NOT_STICKY;}
    startAudio();
    handler.removeCallbacksAndMessages(null);
    handler.postDelayed(this::stopSelf,120000);
    return START_NOT_STICKY;
  }
  private void startAudio(){
    if(player!=null){try{player.stop();player.release();}catch(Exception ignored){}}
    try{
      player=new MediaPlayer();
      player.setAudioAttributes(new AudioAttributes.Builder()
         .setUsage(AudioAttributes.USAGE_ALARM)
         .setContentType(AudioAttributes.CONTENT_TYPE_SONIFICATION).build());
      android.content.res.AssetFileDescriptor fd=getResources().openRawResourceFd(R.raw.nc_alarm05);
      try{player.setDataSource(fd.getFileDescriptor(),fd.getStartOffset(),fd.getLength());}finally{fd.close();}
      player.setLooping(true);player.prepare();player.start();
    }catch(Exception failure){
      if(player!=null){try{player.release();}catch(Exception ignored){}player=null;}
      try{
        Uri sound=RingtoneManager.getDefaultUri(RingtoneManager.TYPE_ALARM);
        player=MediaPlayer.create(this,sound);
        if(player!=null){player.setLooping(true);player.start();}
      }catch(Exception ignored){}
    }
  }
  @Override public void onDestroy(){
    handler.removeCallbacksAndMessages(null);
    if(player!=null){try{if(player.isPlaying())player.stop();player.release();}catch(Exception ignored){}player=null;}
    super.onDestroy();
  }
}
