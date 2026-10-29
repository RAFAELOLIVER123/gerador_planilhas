package com.nossocombinado.app;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.os.Build;
public final class AlarmReceiver extends BroadcastReceiver {
  @Override public void onReceive(Context context, Intent intent){
    if("com.nossocombinado.app.native28.STOP".equals(intent.getAction())){
      context.stopService(new Intent(context,AlarmService.class));return;
    }
    Intent service=new Intent(context,AlarmService.class)
      .putExtra("title",intent.getStringExtra("title"))
      .putExtra("id",intent.getStringExtra("id"));
    try{if(Build.VERSION.SDK_INT>=26)context.startForegroundService(service);else context.startService(service);}
    catch(Exception ignored){}
  }
}
