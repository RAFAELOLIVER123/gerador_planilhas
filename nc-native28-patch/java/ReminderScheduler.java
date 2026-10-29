package com.nossocombinado.app;
import android.content.Context;
import org.json.JSONArray;
import org.json.JSONObject;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import java.util.TimeZone;

public final class ReminderScheduler {
  private ReminderScheduler(){}
  private static long parse(String raw) {
    if(raw == null || raw.isEmpty() || "null".equals(raw)) return 0;
    String s=raw.trim().replace('T',' ').replace("Z","");
    if(s.length()>19)s=s.substring(0,19);
    String pattern=s.length()==10?"yyyy-MM-dd":s.length()==16?"yyyy-MM-dd HH:mm":"yyyy-MM-dd HH:mm:ss";
    try{
      SimpleDateFormat f=new SimpleDateFormat(pattern,Locale.US);
      f.setLenient(false);
      f.setTimeZone(TimeZone.getTimeZone("UTC"));
      Date d=f.parse(s);
      return d==null?0:d.getTime();
    }catch(Exception ignored){return 0;}
  }
  public static void fromDashboard(Context c,JSONObject dashboard,String userId,String householdId){
    if(dashboard==null||(!dashboard.has("tasks")&&!dashboard.has("events")))return;
    long now=System.currentTimeMillis();JSONArray out=new JSONArray();
    JSONArray tasks=dashboard.optJSONArray("tasks");
    if(tasks!=null)for(int i=0;i<tasks.length()&&i<80;i++){
      JSONObject t=tasks.optJSONObject(i);if(t==null||"done".equals(t.optString("status")))continue;
      long when=parse(t.optString("due_at"));int prior=Math.max(0,Math.min(1440,t.optInt("reminder_minutes",0)));
      when-=prior*60_000L;
      if(when<now+5000||when>now+365L*86_400_000L)continue;
      try{out.put(new JSONObject().put("id","t-"+householdId+"-"+t.optString("id")).put("title","Tarefa: "+t.optString("title","Compromisso")).put("at",when));}catch(Exception ignored){}
    }
    JSONArray events=dashboard.optJSONArray("events");
    if(events!=null)for(int i=0;i<events.length()&&i<80;i++){
      JSONObject e=events.optJSONObject(i);if(e==null)continue;
      long when=parse(e.optString("begins_at"));
      int prior=Math.max(0,Math.min(1440,e.optInt("reminder_minutes",0)));when-=prior*60_000L;
      if(when<now+5000||when>now+365L*86_400_000L)continue;
      try{out.put(new JSONObject().put("id","e-"+householdId+"-"+e.optString("id")).put("title","Agenda: "+e.optString("title","Compromisso")).put("at",when));}catch(Exception ignored){}
    }
    new ReminderBridge(c).syncReminders(out.toString());
  }
}
