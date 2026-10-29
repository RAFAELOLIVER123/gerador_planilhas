#!/usr/bin/env python3
"""Patch incremental sobre o Android 1.8 compilado. Sem troca de banco ou endpoints."""
from pathlib import Path
import sys,zipfile,base64
root=Path(sys.argv[1])
p=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
s=p.read_text()
count=[]
def rep(a,b,k):
 global s
 if a not in s:raise SystemExit('Padrão não localizado ('+k+'): '+a[:100])
 s=s.replace(a,b,1);count.append(k)

rep('private boolean onlyMineTasks=false;private long privateCid=0;', 'private boolean onlyMineTasks=false;private final Calendar agendaMonth=Calendar.getInstance(TimeZone.getTimeZone("America/Belem"));private int agendaSelectedDay=-1;private long privateCid=0;', 'state agenda')
rep('getWindow().setStatusBarColor("dark".equals(vault==null?"":vault.get("theme"))?Color.rgb(10,14,24):Color.rgb(35,48,68));getWindow().setNavigationBarColor("dark".equals(vault==null?"":vault.get("theme"))?Color.rgb(10,14,24):Color.rgb(35,48,68));\n        vault=new SecureStore(this);', 'vault=new SecureStore(this);', 'init janela')
rep('private TextView label(String text,int size,int color,boolean bold){TextView v=new TextView(this);v.setText(text);v.setTextColor(color);v.setTextSize(size);v.setGravity(Gravity.CENTER_VERTICAL);v.setFontFeatureSettings("kern");if(bold)v.setTypeface(Typeface.DEFAULT,Typeface.BOLD);return v;}', '''private TextView label(String text,int size,int color,boolean bold){TextView v=new TextView(this);v.setText(translate(text));v.setTextColor(color);v.setTextSize(size);v.setGravity(Gravity.CENTER_VERTICAL);v.setFontFeatureSettings("kern");v.setTypeface(Typeface.create(bold?"sans-serif-medium":"sans-serif",bold?Typeface.BOLD:Typeface.NORMAL));return v;}''', 'fonte e idioma')
rep('private TextView button(String title,int background,int fg,Runnable click){TextView b=label(title,14,fg,true);', 'private TextView button(String title,int background,int fg,Runnable click){TextView b=label(title,14,adjustContrast(fg,background),true);', 'contraste botao')
rep('private EditText edit(String hint,boolean secret){EditText e=new EditText(this);', 'private EditText edit(String hint,boolean secret){EditText e=new EditText(this);hint=translate(hint);', 'idioma formularios')
rep('private void msg(String s){Toast.makeText(this,s,Toast.LENGTH_LONG).show();}', '''private void msg(String s){Toast.makeText(this,translate(s),Toast.LENGTH_LONG).show();}
    private int adjustContrast(int foreground,int background){double a=luminance(foreground),b=luminance(background);if((Math.max(a,b)+0.05)/(Math.min(a,b)+0.05)>=4.5)return foreground;return b>0.32?Color.rgb(31,46,64):Color.rgb(248,250,255);}
    private double luminance(int color){double r=Color.red(color)/255d,g=Color.green(color)/255d,b=Color.blue(color)/255d;r=r<=.04045?r/12.92:Math.pow((r+.055)/1.055,2.4);g=g<=.04045?g/12.92:Math.pow((g+.055)/1.055,2.4);b=b<=.04045?b/12.92:Math.pow((b+.055)/1.055,2.4);return .2126*r+.7152*g+.0722*b;}
    private String translate(String s){return NcLanguage.tr(vault==null?"pt":vault.get("language"),s);}
    private void chooseLanguage(){final String[] codes={"pt","en","es"};dialogBuilder().setTitle("Idioma / Language / Idioma").setItems(new String[]{"🇧🇷 Português","🇺🇸 English","🇪🇸 Español"},(d,index)->{vault.set("language",codes[index]);if(userId.isEmpty())authScreen(false);else appScreen();}).show();}''','contraste idioma função')
rep('private String itemDate(String s){if(s==null)return "";return s.replace("T"," ").replace("Z","");}', '''private static final TimeZone NC_TZ=TimeZone.getTimeZone("America/Belem");
    private Date parseUtcDate(String raw){if(raw==null||raw.trim().isEmpty())return null;try{
      String v=raw.trim().replace('T',' ').replace("Z","");int dot=v.indexOf('.');if(dot>=0)v=v.substring(0,dot);if(v.length()>19)v=v.substring(0,19);
      SimpleDateFormat f=new SimpleDateFormat(v.length()<=10?"yyyy-MM-dd":v.length()<=16?"yyyy-MM-dd HH:mm":"yyyy-MM-dd HH:mm:ss",Locale.US);
      f.setTimeZone(TimeZone.getTimeZone("UTC"));f.setLenient(false);return f.parse(v);
    }catch(Exception ex){return null;}}
    private String itemDate(String s){Date d=parseUtcDate(s);if(d==null)return s==null?"":s;SimpleDateFormat out=new SimpleDateFormat("dd/MM/yyyy hh:mm:ss a",Locale.US);out.setTimeZone(NC_TZ);return out.format(d);}
    private String dateKey(String s){Date d=parseUtcDate(s);if(d==null)return "";SimpleDateFormat out=new SimpleDateFormat("yyyy-MM-dd",Locale.US);out.setTimeZone(NC_TZ);return out.format(d);}''','UTC --> Belém dd/mm/aaaa 12h')
rep('private String utc(String local)throws Exception{SimpleDateFormat parse=new SimpleDateFormat("yyyy-MM-dd HH:mm",Locale.US);parse.setLenient(false);', 'private String utc(String local)throws Exception{SimpleDateFormat parse=new SimpleDateFormat("yyyy-MM-dd HH:mm",Locale.US);parse.setTimeZone(NC_TZ);parse.setLenient(false);', 'formulario converte Belem UTC')
rep('SimpleDateFormat out=new SimpleDateFormat("yyyy-MM-dd HH:mm",Locale.US);return out.format(date);','SimpleDateFormat out=new SimpleDateFormat("yyyy-MM-dd HH:mm",Locale.US);out.setTimeZone(NC_TZ);return out.format(date);','editor UTC para Belem')
rep('private String today(){return new SimpleDateFormat("yyyy-MM-dd",Locale.US).format(new Date());}', 'private String today(){SimpleDateFormat f=new SimpleDateFormat("yyyy-MM-dd",Locale.US);f.setTimeZone(NC_TZ);return f.format(new Date());}', 'hoje no fuso')
rep('caption(c,localDate(val(e,"begins_at"))+"  ·  "+val(e,"place"));','caption(c,itemDate(val(e,"begins_at"))+"  ·  "+val(e,"place"));','exibir data evento correta')
# Agenda: calendário por dias, pontos coloridos, troca por setas/ano/gesto, seleção filtrável.
start=s.index('    private void calendar(JSONObject j){')
end=s.index('    private void newEvent(){',start)
calendar=r'''    private void calendar(JSONObject j){
        sectionTitle("Agenda da família","Calendário com tarefas e compromissos por dia.");
        JSONArray events=array(j,"events"),tasks=array(j,"tasks");
        agendaMonth.set(Calendar.DAY_OF_MONTH,1);if(agendaSelectedDay<0)agendaSelectedDay=Calendar.getInstance(NC_TZ).get(Calendar.DAY_OF_MONTH);
        int year=agendaMonth.get(Calendar.YEAR),month=agendaMonth.get(Calendar.MONTH);
        int max=agendaMonth.getActualMaximum(Calendar.DAY_OF_MONTH);if(agendaSelectedDay>max)agendaSelectedDay=max;
        LinearLayout panel=card(content);panel.setPadding(dp(9),dp(12),dp(9),dp(14));
        LinearLayout controls=horizontal();
        TextView yearBack=button("«",MINT,NAVY,()->changeAgendaMonth(-12));
        TextView prev=button("‹",MINT,NAVY,()->changeAgendaMonth(-1));
        SimpleDateFormat monthTitle=new SimpleDateFormat("MMMM yyyy",new Locale("pt","BR"));monthTitle.setTimeZone(NC_TZ);
        TextView title=button(monthTitle.format(agendaMonth.getTime()).toUpperCase(new Locale("pt","BR")),SURFACE,NAVY,()->{
            new DatePickerDialog(this,(v,y,m,d)->{agendaMonth.set(y,m,1);agendaSelectedDay=d;loadAgendaLocal();},year,month,agendaSelectedDay).show();
        });title.setTextSize(14);
        TextView next=button("›",MINT,NAVY,()->changeAgendaMonth(1));TextView yearNext=button("»",MINT,NAVY,()->changeAgendaMonth(12));
        controls.addView(yearBack,new LinearLayout.LayoutParams(dp(39),dp(44)));controls.addView(prev,new LinearLayout.LayoutParams(dp(35),dp(44)));controls.addView(title,new LinearLayout.LayoutParams(0,dp(44),1));controls.addView(next,new LinearLayout.LayoutParams(dp(35),dp(44)));controls.addView(yearNext,new LinearLayout.LayoutParams(dp(39),dp(44)));panel.addView(controls);gap(panel,13);
        LinearLayout weekdays=horizontal();for(String day:new String[]{"D","S","T","Q","Q","S","S"}){TextView w=label(day,12,MUTED,true);w.setGravity(Gravity.CENTER);weekdays.addView(w,new LinearLayout.LayoutParams(0,dp(22),1));}panel.addView(weekdays);
        Calendar first=(Calendar)agendaMonth.clone();first.set(Calendar.DAY_OF_MONTH,1);int offset=first.get(Calendar.DAY_OF_WEEK)-1;int n=1,gridRows=(offset+max+6)/7;
        final android.view.GestureDetector gesture=new android.view.GestureDetector(this,new android.view.GestureDetector.SimpleOnGestureListener(){
            @Override public boolean onDown(android.view.MotionEvent e){return true;}
            @Override public boolean onFling(android.view.MotionEvent a,android.view.MotionEvent b,float vx,float vy){if(a!=null&&b!=null&&Math.abs(b.getX()-a.getX())>dp(55)&&Math.abs(vx)>Math.abs(vy)){changeAgendaMonth(b.getX()<a.getX()?1:-1);return true;}return false;}
        });
        for(int row=0;row<gridRows;row++){LinearLayout week=horizontal();for(int c=0;c<7;c++){
            int day=row*7+c-offset+1;LinearLayout.LayoutParams lp=new LinearLayout.LayoutParams(0,dp(62),1);lp.setMargins(dp(1),dp(2),dp(1),dp(2));
            LinearLayout cell=vertical();cell.setGravity(Gravity.CENTER);week.addView(cell,lp);
            if(day<1||day>max)continue;
            String key=String.format(Locale.US,"%04d-%02d-%02d",year,month+1,day);boolean hasTask=false,hasEvent=false;
            for(int k=0;k<tasks.length();k++)if(key.equals(dateKey(val(row(tasks,k),"due_at"))))hasTask=true;
            for(int k=0;k<events.length();k++)if(key.equals(dateKey(val(row(events,k),"begins_at"))))hasEvent=true;
            boolean selected=day==agendaSelectedDay;cell.setBackground(shape(selected?PINK:SURFACE,13));
            TextView value=label(String.valueOf(day),14,selected?Color.WHITE:NAVY,selected);value.setGravity(Gravity.CENTER);cell.addView(value);
            LinearLayout dots=horizontal();dots.setGravity(Gravity.CENTER);if(hasTask)text(dots,"●",10,selected?Color.WHITE:Color.rgb(207,122,54),true);if(hasEvent)text(dots," ●",10,selected?Color.WHITE:Color.rgb(58,149,120),true);cell.addView(dots);
            final int picked=day;cell.setOnClickListener(v->{agendaSelectedDay=picked;loadAgendaLocal();});
            cell.setOnTouchListener((v,e)->{boolean swipe=gesture.onTouchEvent(e);if(e.getAction()==android.view.MotionEvent.ACTION_UP&&!swipe){v.performClick();}return true;});
        }panel.addView(week);}
        gap(panel,8);caption(panel,"● Laranja: tarefas   ● Verde: compromissos   • Deslize para mudar o mês");
        gap(content,11);addButton(content,"+ Novo compromisso",PINK,WHITE,this::newEvent);gap(content,12);
        String chosen=String.format(Locale.US,"%04d-%02d-%02d",year,month+1,agendaSelectedDay);
        text(content,"Atividades de "+String.format(Locale.US,"%02d/%02d/%04d",agendaSelectedDay,month+1,year),17,NAVY,true);gap(content,9);
        int hits=0;
        for(int i=0;i<events.length();i++){JSONObject e=row(events,i);if(!chosen.equals(dateKey(val(e,"begins_at"))))continue;hits++;LinearLayout card=card(content);text(card,"📅 "+val(e,"title"),16,NAVY,true);caption(card,itemDate(val(e,"begins_at"))+" · "+val(e,"place"));gap(card,7);
            LinearLayout actions=horizontal();actions.addView(button("Editar",MINT,NAVY,()->editEvent(e)),new LinearLayout.LayoutParams(0,dp(42),1));gapH(actions,6);
            actions.addView(button("Excluir",Color.rgb(250,239,235),Color.rgb(138,67,60),()->dialogBuilder().setTitle("Excluir compromisso?").setNegativeButton("Cancelar",null).setPositiveButton("Excluir",(d,w)->{if(hasConnection())mutate("event_delete",data("id",val(e,"id")),false);else msg("Conecte-se para excluir.");}).show()),new LinearLayout.LayoutParams(0,dp(42),1));card.addView(actions);
        }
        for(int i=0;i<tasks.length();i++){JSONObject t=row(tasks,i);if(!chosen.equals(dateKey(val(t,"due_at"))))continue;hits++;LinearLayout c=card(content);text(c,"✓ "+val(t,"title"),16,NAVY,true);caption(c,itemDate(val(t,"due_at"))+" · "+("done".equals(val(t,"status"))?"Concluída":"Pendente"));if(!"done".equals(val(t,"status"))&&canCompleteTask(t)){gap(c,7);TextView done=button("✓ Concluir",PINK,WHITE,()->{});done.setOnClickListener(v->completeTask(t,done));c.addView(done,new LinearLayout.LayoutParams(-1,dp(42)));}}
        if(hits==0)caption(content,"Não há tarefas ou compromissos para este dia.");
    }
    private void changeAgendaMonth(int delta){agendaMonth.add(Calendar.MONTH,delta);agendaSelectedDay=1;loadAgendaLocal();}
    private void loadAgendaLocal(){String c=offline.load(userId,householdId,"calendar");try{drawPage(c.isEmpty()?new JSONObject():new JSONObject(c),false);}catch(Exception e){loadPage(true);}}
'''
s=s[:start]+calendar+s[end:];count.append('calendario interativo')
# Remover controles obsoletos da Rede, perfil so nos elementos de autoria
rep('gap(c,4);caption(c,postTypeLabel(val(p,"post_type")));if(mine){gap(c,7);addButton(c,"📷 Adicionar foto ao post",MINT,NAVY,()->{\n                uploadMode="media_upload";uploadEntity="professional_post:"+val(p,"id");pickFile();});}gap(c,9);', 'gap(c,4);caption(c,postTypeLabel(val(p,"post_type")));gap(c,9);','remover adicionar foto após publicar')
rep('            if(p.optLong("professional_id")>0)addButton(c,"Ver perfil profissional",MINT,NAVY,()->openProfessional(data("id",val(p,"professional_id"))));\n', '', 'remover botao perfil')
rep('header.addView(authorAvatar,avatarParams);}', 'header.addView(authorAvatar,avatarParams);if(p.optLong("professional_id")>0)authorAvatar.setOnClickListener(v->openProfessional(data("id",val(p,"professional_id"))));}', 'perfil ao clicar na miniatura')
rep('header.addView(author,new LinearLayout.LayoutParams(0,-2,1));', 'header.addView(author,new LinearLayout.LayoutParams(0,-2,1));if(p.optLong("professional_id")>0)author.setOnClickListener(v->openProfessional(data("id",val(p,"professional_id"))));','perfil ao clicar no nome')
# Em familia: deixar claro grupo ativo + botoes de acao, fallback online e contagem.
rep('private void members(JSONObject j){sectionTitle("Nossa família","Integrantes com contas e permissões individuais.");addButton(content,"+ Gerar convite de uso único",PINK,WHITE,()->{','private void members(JSONObject j){sectionTitle("Nossa família · "+householdName,"Integrantes do grupo e convites com acesso autorizado.");JSONArray currentMembers=array(j,"items");caption(content,currentMembers.length()+" integrante(s) no grupo ativo");gap(content,10);addButton(content,"↻ Atualizar integrantes",MINT,NAVY,()->loadPage(false));gap(content,8);addButton(content,"+ Gerar convite de uso único",PINK,WHITE,()->{','familia grupo e refresh')
rep('JSONArray a=array(j,"items");for(int i=0;i<a.length();i++){JSONObject m=row(a,i);LinearLayout c=card(content);', 'JSONArray a=array(j,"items");if(a.length()==0){placeholder("Nenhum integrante recebido. Confira se esta é a família correta e toque em Atualizar integrantes.");}for(int i=0;i<a.length();i++){JSONObject m=row(a,i);LinearLayout c=card(content);', 'familia estado vazio')
rep('addButton(banner,"Ver integrantes e convites",WHITE,NAVY,()->navigate("members"));', 'addButton(banner,"Ver integrantes e convites",MINT,NAVY,()->navigate("members"));gap(banner,7);addButton(banner,"Acesso rápido à família",MINT,NAVY,()->navigate("members"));', 'atalho familia visivel')
# Marca oficial e banner sem alterar dados.
rep('TextView brand=label("Nosso Combinado",29,NAVY,true);brand.setGravity(Gravity.CENTER);c.addView(brand);', 'ImageView welcome=new ImageView(this);welcome.setImageResource(R.drawable.nc_family_banner);welcome.setScaleType(ImageView.ScaleType.CENTER_CROP);welcome.setAdjustViewBounds(true);LinearLayout.LayoutParams welcomeLp=new LinearLayout.LayoutParams(-1,dp(112));welcomeLp.topMargin=dp(12);welcomeLp.bottomMargin=dp(10);c.addView(welcome,welcomeLp);TextView brand=label("Nosso Combinado",29,NAVY,true);brand.setGravity(Gravity.CENTER);c.addView(brand);', 'capa imagem login')
rep('gap(c,8);TextView desc=label(register?', 'gap(c,8);addButton(c,"🌐 Idioma / Language / Idioma",MINT,NAVY,this::chooseLanguage);gap(c,8);TextView desc=label(register?', 'idioma entrada')
rep('gap(hero,8);text(hero,"Nosso Combinado",23,NAVY,true);', 'ImageView bannerArt=new ImageView(this);bannerArt.setImageResource(R.drawable.nc_family_banner);bannerArt.setScaleType(ImageView.ScaleType.CENTER_CROP);bannerArt.setClipToOutline(true);bannerArt.setBackground(shape(Color.rgb(250,193,175),16));hero.addView(bannerArt,new LinearLayout.LayoutParams(-1,dp(112)));gap(hero,8);text(hero,"Nosso Combinado",23,NAVY,true);','capa na tela inicial')
rep('addButton(controls,"◐  Aparência: quatro temas (inclui preto)",MINT,NAVY,this::chooseTheme);gap(controls,8);', 'addButton(controls,"◐  Aparência: quatro temas (inclui preto)",MINT,NAVY,this::chooseTheme);gap(controls,8);addButton(controls,"🌐 Idioma / Language / Idioma",MINT,NAVY,this::chooseLanguage);gap(controls,8);', 'idioma dentro app')
rep('if("ocean".equals(theme)){BG=Color.rgb(246,249,254);NAVY=Color.rgb(25,52,83);PINK=Color.rgb(31,104,193);MINT=Color.rgb(219,235,253);}', 'if("ocean".equals(theme)){BG=Color.rgb(238,245,253);NAVY=Color.rgb(18,42,72);PINK=Color.rgb(33,95,173);MINT=Color.rgb(215,231,250);}', 'tema oceano')
rep('else if("dark".equals(theme)){BG=Color.rgb(14,20,32);SURFACE=Color.rgb(30,40,56);NAVY=Color.rgb(243,246,253);PINK=Color.rgb(152,164,255);MINT=Color.rgb(48,66,86);MUTED=Color.rgb(190,201,216);}', 'else if("dark".equals(theme)){BG=Color.rgb(13,19,30);SURFACE=Color.rgb(28,38,53);NAVY=Color.rgb(248,250,255);PINK=Color.rgb(246,150,135);MINT=Color.rgb(47,67,88);MUTED=Color.rgb(204,216,230);}', 'tema dark contraste')
rep('private GradientDrawable outline(int color){GradientDrawable g=shape(SURFACE,19);g.setStroke(dp(1),Color.alpha(SURFACE)==255&&SURFACE!=WHITE?Color.rgb(63,72,93):Color.rgb(228,232,238));return g;}', 'private GradientDrawable outline(int color){GradientDrawable g=shape(SURFACE,19);g.setStroke(dp(1),SURFACE!=WHITE?Color.rgb(75,93,114):Color.rgb(228,232,238));return g;}', 'bordas dark')
# Evitar estilo claro para botoes expressamente brancos em dark, antes invisiveis: contraste calculado.
rep('        LinearLayout quick=horizontal();\n        TextView qAgenda=button("📅 Agenda",Color.rgb(232,237,254),NAVY,()->navigate("calendar"));\n        TextView qNotes=button("📄 Arquivos",Color.rgb(229,247,238),NAVY,()->navigate("documents"));\n        quick.addView(qAgenda,new LinearLayout.LayoutParams(0,dp(49),1));gapH(quick,9);quick.addView(qNotes,new LinearLayout.LayoutParams(0,dp(49),1));content.addView(quick);gap(content,14);','        LinearLayout quick=horizontal();\n        quick.addView(microQuick("Agenda",R.drawable.nc_micro_agenda,()->navigate("calendar")),new LinearLayout.LayoutParams(0,dp(77),1));gapH(quick,7);\n        quick.addView(microQuick("Tarefas",R.drawable.nc_micro_tarefas,()->navigate("tasks")),new LinearLayout.LayoutParams(0,dp(77),1));gapH(quick,7);\n        quick.addView(microQuick("Família",R.drawable.nc_micro_familia,()->navigate("members")),new LinearLayout.LayoutParams(0,dp(77),1));gapH(quick,7);\n        quick.addView(microQuick("Rede",R.drawable.nc_micro_rede,()->navigate("network")),new LinearLayout.LayoutParams(0,dp(77),1));content.addView(quick);gap(content,14);','microimagens acesso rapido')
rep('private void gap(LinearLayout col,int d){','private LinearLayout microQuick(String title,int res,Runnable action){LinearLayout box=vertical();box.setGravity(Gravity.CENTER);box.setBackground(shape(SURFACE,15));box.setPadding(dp(4),dp(7),dp(4),dp(7));ImageView ico=new ImageView(this);ico.setImageResource(res);ico.setScaleType(ImageView.ScaleType.FIT_CENTER);box.addView(ico,new LinearLayout.LayoutParams(dp(38),dp(38)));TextView cap=label(title,11,NAVY,true);cap.setGravity(Gravity.CENTER);box.addView(cap);box.setOnClickListener(v->action.run());return box;}\n    private void gap(LinearLayout col,int d){','microimagens helper')
# Salvar versao 1.9
bg=root/'app/build.gradle';g=bg.read_text()
if 'versionCode 180' not in g or "versionName '1.8.0'" not in g:raise SystemExit('Versão 1.8 ausente')
bg.write_text(g.replace('versionCode 180','versionCode 190').replace("versionName '1.8.0'","versionName '1.9.0'"))
p.write_text(s)
print('OK '+str(len(count))+' ajustes: '+', '.join(count))