#!/usr/bin/env python3
from pathlib import Path
import shutil, sys
root=Path(sys.argv[1]).resolve()
patch=Path(__file__).resolve().parent
main=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
grad=root/'app/build.gradle'
manifest=root/'app/src/main/AndroidManifest.xml'
s=main.read_text(encoding='utf-8')
def change(a,b,label):
 global s
 if a not in s:raise RuntimeError('Native28 missing '+label+': '+a[:90])
 s=s.replace(a,b,1)
 print('+',label)
change('private LinearLayout base,content,topBar,bottomBar;private ScrollView scroll;',
       'private LinearLayout base,content,topBar,bottomBar;private ScrollView scroll;private TextView unreadMessagesButton;',
       'badge messages')
change('    private interface Callback {void done(JSONObject j,String error);}',
'''    private final Runnable realtimeInbox=new Runnable(){@Override public void run(){
        if(!userId.isEmpty()&&!householdId.isEmpty()){
            refreshInboxCount();
            if(page.equals("chat")||page.equals("messages")||page.equals("private_thread"))loadPage(false);
        }
        handler.postDelayed(this,3600);
    }};
    private void refreshInboxCount(){
        if(api==null||unreadMessagesButton==null)return;
        TextView buttonNow=unreadMessagesButton;
        api.get("notification_list",null,(j,error)->{
            if(error!=null||buttonNow!=unreadMessagesButton)return;
            JSONArray items=array(j,"items");int count=0;
            for(int i=0;i<items.length();i++){
                JSONObject n=row(items,i);String kind=val(n,"kind");
                if(("chat".equals(kind)||"private_chat".equals(kind))&&val(n,"read_at").isEmpty())count++;
            }
            buttonNow.setText(count>0?"✉ "+(count>99?"99+":count):"✉");
            buttonNow.setContentDescription(count>0?count+" mensagens não lidas":"Mensagens");
        });
    }
    private void openPlatform(){startActivity(new Intent(this,PlatformActivity.class));}
    private interface Callback {void done(JSONObject j,String error);}''',
       'notify badge and web platform')
change('handler.postDelayed(periodic,45000);','handler.postDelayed(periodic,45000);handler.postDelayed(realtimeInbox,3600);','run realtime badge')
change('handler.removeCallbacks(periodic);if(map!=null)',
       'handler.removeCallbacks(periodic);handler.removeCallbacks(realtimeInbox);if(map!=null)',
       'cleanup loops')
change('TextView bell=button("🔔",Color.rgb(246,243,241),NAVY,()->navigate("alerts"));topBar.addView(bell);',
'''unreadMessagesButton=button("✉",Color.rgb(255,238,230),NAVY,()->navigate("messages"));
        unreadMessagesButton.setContentDescription("Mensagens");topBar.addView(unreadMessagesButton);
        TextView bell=button("🔔",Color.rgb(246,243,241),NAVY,()->navigate("alerts"));topBar.addView(bell);
        refreshInboxCount();''','top icon chat')
change('if(e==null){connected=true;offline.save(userId,householdId,page,j.toString());drawPage(j,false);}',
'''if(e==null){connected=true;offline.save(userId,householdId,page,j.toString());if("dashboard".equals(action))ReminderScheduler.fromDashboard(this,j,userId,householdId);drawPage(j,false);}''',
       'native alarm scheduler')
change('private void drawPage(JSONObject j,boolean cached){if(content==null)return;content.removeAllViews();',
'''private void drawPage(JSONObject j,boolean cached){if(content==null)return;
        final int previousY=scroll==null?0:scroll.getScrollY();
        final boolean keepPosition=page.equals("chat")||page.equals("messages")||page.equals("private_thread");
        content.removeAllViews();''','preserve scrolling')
change('            default:showMore();\n        }\n    }',
'''            default:showMore();
        }
        if(keepPosition&&scroll!=null&&previousY>0)scroll.post(()->scroll.scrollTo(0,previousY));
    }''','keep chat scroll')
change('lineAction(shortcuts,"✓  Criar tarefa","Abrir",this::newTask);',
       'lineAction(shortcuts,"🌐  Plataforma atualizada","Abrir",this::openPlatform);lineAction(shortcuts,"✓  Criar tarefa","Abrir",this::newTask);',
       'site shortcut')
change('addButton(content,"Mensagens individuais",MINT,NAVY,()->navigate("messages"));gap(content,12);',
       'addButton(content,"💬  Chat em tempo real — versão web",PINK,WHITE,this::openPlatform);gap(content,9);addButton(content,"Mensagens individuais",MINT,NAVY,()->navigate("messages"));gap(content,12);',
       'responsive web chat button')
change('actions.addView(button("Atualizar",MINT,NAVY,()->loadPage(false)),new LinearLayout.LayoutParams(0,dp(47),1));content.addView(actions);gap(content,12);',
       'actions.addView(button("Atualizar",MINT,NAVY,()->loadPage(false)),new LinearLayout.LayoutParams(0,dp(47),1));content.addView(actions);gap(content,8);addButton(content,"🗺  Mapa completo, fotos e endereços",MINT,NAVY,this::openPlatform);gap(content,12);',
       'map interactive web')
change('m.setSubDescription("Atualizado: "+val(p,"captured_at"));map.getOverlays().add(m);',
'''m.setSubDescription("Atualizado: "+val(p,"captured_at")+". Toque para abrir no Google Maps.");
                final String mapsUrl="https://www.google.com/maps/search/?api=1&query="+lat+","+lon;
                m.setOnMarkerClickListener((marker,mapView)->{
                    try{startActivity(new Intent(Intent.ACTION_VIEW,Uri.parse(mapsUrl)));}
                    catch(Exception ex){msg("Instale um aplicativo de mapas para abrir a localização.");}
                    return true;
                });
                map.getOverlays().add(m);''','native map Google Maps')
change('gap(content,12);LinearLayout controls=card(content);text(controls,"Conta e sincronização",17,NAVY,true);',
       'gap(content,12);LinearLayout controls=card(content);text(controls,"Conta e sincronização",17,NAVY,true);gap(controls,8);addButton(controls,"🌐  Abrir site da plataforma",MINT,NAVY,this::openPlatform);',
       'settings link to site')
main.write_text(s,encoding='utf8')
g=grad.read_text()
import re
for pattern,replacement in [(r"applicationId\\s+['\"]com\\.nossocombinado\\.app['\"]","applicationId 'com.nossocombinado.app.native28'"),(r"versionCode\\s+\\d+","versionCode 280"),(r"versionName\\s+['\"][^'\"]+['\"]","versionName '2.8.0'"),(r"targetSdk\\s+\\d+","targetSdk 33")]:
 g,n=re.subn(pattern,replacement,g,count=1)
 if n!=1:raise RuntimeError('Gradle missing '+pattern)
grad.write_text(g)
m=manifest.read_text()
m=m.replace('<application android:', '''<uses-permission android:name="android.permission.POST_NOTIFICATIONS"/>
    <uses-permission android:name="android.permission.FOREGROUND_SERVICE"/>
    <uses-permission android:name="android.permission.FOREGROUND_SERVICE_MEDIA_PLAYBACK"/>
    <uses-permission android:name="android.permission.RECEIVE_BOOT_COMPLETED"/>
    <uses-permission android:name="android.permission.SCHEDULE_EXACT_ALARM"/>
    <uses-permission android:name="android.permission.WAKE_LOCK"/>
    <application android:''')
m=m.replace('        <activity android:name=".MainActivity"', '''        <receiver android:name=".AlarmReceiver" android:exported="false"/>
        <receiver android:name=".BootReceiver" android:exported="false">
          <intent-filter><action android:name="android.intent.action.BOOT_COMPLETED"/></intent-filter>
        </receiver>
        <service android:name=".AlarmService" android:exported="false" android:foregroundServiceType="mediaPlayback"/>
        <activity android:name=".AlarmActivity" android:exported="false"/>
        <activity android:name=".PlatformActivity" android:exported="false" android:windowSoftInputMode="adjustResize"/>
        <activity android:name=".MainActivity"''')
manifest.write_text(m)
folder=root/'app/src/main/java/com/nossocombinado/app'
for f in (patch/'java').glob('*.java'):
 shutil.copy2(f,folder/f.name)
print('Ready: native 2.8 MainActivity + osmdroid + offline + alarm and web chat',len(s))
