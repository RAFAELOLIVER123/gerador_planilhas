package com.nossocombinado.app;

import android.app.AlarmManager;
import android.app.PendingIntent;
import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.os.Build;
import android.webkit.JavascriptInterface;
import org.json.JSONArray;
import org.json.JSONObject;
import java.util.HashSet;
import java.util.Set;

public final class ReminderBridge {
  private static final String PREF="nc26_alarmas";
  private static final String KEY="agenda";
  private final Context context;
  ReminderBridge(Context c){context=c.getApplicationContext();}
  @JavascriptInterface public boolean canNotify(){return true;}
  @JavascriptInterface public void syncReminders(String json){
    if(json==null||json.length()>80000)return;
    try {
      JSONArray input=new JSONArray(json);
      JSONArray clean=new JSONArray();
      Set<String> ids=new HashSet<>();
      long now=System.currentTimeMillis(),max=now+365L*86400000L;
      for(int i=0;i<input.length()&&i<120;i++){
        JSONObject entry=input.optJSONObject(i);if(entry==null)continue;
        String id=entry.optString("id","").trim();
        String title=entry.optString("title","Compromisso");
        long at=entry.optLong("at",0);
        if(id.isEmpty()||id.length()>80||!ids.add(id)||at<now+3000||at>max)continue;
        clean.put(new JSONObject().put("id",id).put("title",title.substring(0,Math.min(120,title.length()))).put("at",at));
      }
      context.getSharedPreferences(PREF,Context.MODE_PRIVATE).edit().putString(KEY,clean.toString()).apply();
      scheduleList(context,clean,true);
    }catch(Exception ignored){}
  }
  static void restore(Context ctx){
    try{String data=ctx.getSharedPreferences(PREF,Context.MODE_PRIVATE).getString(KEY,"[]");scheduleList(ctx,new JSONArray(data),false);}catch(Exception ignored){}
  }
  private static void scheduleList(Context ctx,JSONArray entries,boolean clearOld){
    AlarmManager am=(AlarmManager)ctx.getSystemService(Context.ALARM_SERVICE);if(am==null)return;
    SharedPreferences prefs=ctx.getSharedPreferences(PREF,Context.MODE_PRIVATE);
    Set<String> next=new HashSet<>();
    for(int i=0;i<entries.length();i++){JSONObject item=entries.optJSONObject(i);if(item!=null)next.add(item.optString("id"));}
    if(clearOld)for(String id:prefs.getStringSet("scheduled",new HashSet<String>()))if(!next.contains(id)){PendingIntent pi=pending(ctx,id,"");if(pi!=null)am.cancel(pi);}
    for(int i=0;i<entries.length();i++){
      JSONObject e=entries.optJSONObject(i);if(e==null)continue;
      String id=e.optString("id"),title=e.optString("title","Compromisso");
      long at=e.optLong("at");if(at<=System.currentTimeMillis()+1000)continue;
      PendingIntent pi=pending(ctx,id,title);
      try{
        if(Build.VERSION.SDK_INT>=31&&am.canScheduleExactAlarms())am.setExactAndAllowWhileIdle(AlarmManager.RTC_WAKEUP,at,pi);
        else if(Build.VERSION.SDK_INT>=23)am.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP,at,pi);
        else am.set(AlarmManager.RTC_WAKEUP,at,pi);
      }catch(SecurityException permission){
        am.set(AlarmManager.RTC_WAKEUP,at,pi);
      }
    }
    prefs.edit().putStringSet("scheduled",next).apply();
  }
  private static PendingIntent pending(Context ctx,String id,String title){
    Intent intent=new Intent(ctx,AlarmReceiver.class);
    intent.setAction("com.nossocombinado.app.native28.ALARM");
    intent.putExtra("id",id);intent.putExtra("title",title);
    return PendingIntent.getBroadcast(ctx,id.hashCode(),intent,PendingIntent.FLAG_UPDATE_CURRENT|PendingIntent.FLAG_IMMUTABLE);
  }
}
