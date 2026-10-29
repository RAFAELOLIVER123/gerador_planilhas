#!/usr/bin/env python3
"""Nosso Combinado: UI nativa v2, comentários e mensagens inline; sobrepõe a v1.9."""
from pathlib import Path
import sys,re
root=Path(sys.argv[1] if len(sys.argv)>1 else '/mnt/data/nc_v19/localrepo/android-v1')
f=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
s=f.read_text()

def change(a,b,label):
 global s
 if a not in s:raise Exception('Não encontrado '+label+' '+a[:85])
 s=s.replace(a,b,1)
 print('+',label)

# Native palette and contrast: keep white labels dark; dark surface backgrounds selected per theme
change('private int BG=Color.rgb(251,249,246), NAVY=Color.rgb(35,48,68), PINK=Color.rgb(238,118,142), MINT=Color.rgb(225,245,234);',
'private int BG=Color.rgb(255,249,246), NAVY=Color.rgb(34,45,61), PINK=Color.rgb(247,107,98), MINT=Color.rgb(235,248,243);','paleta')
change('private LinearLayout base,content,topBar,bottomBar;private ScrollView scroll;',
'private LinearLayout base,content,topBar,bottomBar;private ScrollView scroll;private EditText activeDmInput;', 'campo mensagens')
# cards more compact, slight shadow and rounded
change('private LinearLayout card(LinearLayout parent){LinearLayout c=vertical();c.setPadding(dp(16),dp(14),dp(16),dp(15));c.setBackground(outline(WHITE));LinearLayout.LayoutParams lp=new LinearLayout.LayoutParams(-1,-2);lp.bottomMargin=dp(10);parent.addView(c,lp);return c;}',
'''private LinearLayout card(LinearLayout parent){LinearLayout c=vertical();c.setPadding(dp(17),dp(15),dp(17),dp(15));c.setBackground(outline(SURFACE));c.setElevation(dp(2));LinearLayout.LayoutParams lp=new LinearLayout.LayoutParams(-1,-2);lp.bottomMargin=dp(11);parent.addView(c,lp);return c;}''','cartões')
# Hero redesign
start=s.index('    private void home(JSONObject j){')
end=s.index('    private void itemCard(',start)
home=r'''    private void home(JSONObject j){
        sectionTitle("Olá, "+(userName.isEmpty()?"família":userName.split(" ")[0])+"! 👋","Vamos organizar o dia em família.");
        LinearLayout hero=card(content);hero.setPadding(dp(0),dp(0),dp(0),dp(0));
        ImageView familyArt=new ImageView(this);familyArt.setImageResource(R.drawable.nc_family_banner);familyArt.setScaleType(ImageView.ScaleType.CENTER_CROP);
        familyArt.setBackground(shape(Color.rgb(247,153,132),21));familyArt.setClipToOutline(true);
        hero.addView(familyArt,new LinearLayout.LayoutParams(-1,dp(148)));
        LinearLayout artText=vertical();artText.setPadding(dp(18),dp(10),dp(18),dp(14));text(artText,"Nossa família · "+householdName,15,NAVY,true);
        gap(artText,3);caption(artText,"Cada combinado aproxima a nossa família.");hero.addView(artText);
        JSONArray tasks=array(j,"tasks"),events=array(j,"events");int open=0;for(int i=0;i<tasks.length();i++)if(!"done".equals(val(row(tasks,i),"status")))open++;
        LinearLayout counter=horizontal();
        String[] values={String.valueOf(open),String.valueOf(events.length()),householdName.isEmpty()?"♡":"♥"};
        String[] titles={"Tarefas","Agenda","Família"};int[] colors={Color.rgb(255,237,232),Color.rgb(235,242,255),Color.rgb(232,248,240)};
        for(int i=0;i<3;i++){final int k=i;LinearLayout tile=vertical();tile.setPadding(dp(10),dp(14),dp(8),dp(12));tile.setBackground(shape("dark".equals(vault.get("theme"))?SURFACE:colors[i],16));
            text(tile,values[i],21,NAVY,true);gap(tile,3);text(tile,titles[i],11,MUTED,false);
            LinearLayout.LayoutParams lp=new LinearLayout.LayoutParams(0,-2,1);lp.rightMargin=dp(i==2?0:7);counter.addView(tile,lp);
            tile.setOnClickListener(v->navigate(k==0?"tasks":k==1?"calendar":"members"));
        }content.addView(counter);gap(content,16);
        text(content,"Acesso rápido",17,NAVY,true);gap(content,8);
        String[][] shortcuts={{"✓","Nova tarefa","tasks"},{"▦","Agenda","calendar"},{"☷","Lista","shopping"},{"♧","Rede","network"}};
        for(int row=0;row<2;row++){
            LinearLayout pair=horizontal();for(int col=0;col<2;col++){int idx=row*2+col;String[] opt=shortcuts[idx];
                LinearLayout tile=horizontal();tile.setPadding(dp(12),dp(13),dp(10),dp(13));tile.setBackground(outline(SURFACE));tile.setElevation(dp(2));
                TextView ico=label(opt[0],23,PINK,true);tile.addView(ico);gapH(tile,10);text(tile,opt[1],13,NAVY,true);
                LinearLayout.LayoutParams lp=new LinearLayout.LayoutParams(0,-2,1);lp.rightMargin=dp(col==0?8:0);pair.addView(tile,lp);
                tile.setOnClickListener(v->{if(idx==0)newTask();else navigate(opt[2]);});
            }content.addView(pair);gap(content,9);
        }
        gap(content,7);LinearLayout header=horizontal();TextView title=label("Próximas tarefas",18,NAVY,true);header.addView(title,new LinearLayout.LayoutParams(0,-2,1));
        header.addView(button("Ver todas",MINT,NAVY,()->navigate("tasks")));content.addView(header);gap(content,9);
        int shown=0;for(int i=0;i<tasks.length()&&shown<4;i++){
            JSONObject t=row(tasks,i);if("done".equals(val(t,"status")))continue;
            LinearLayout item=card(content);LinearLayout line=horizontal();text(line,"☐",23,PINK,true);gapH(line,12);
            LinearLayout info=vertical();text(info,val(t,"title"),15,NAVY,true);if(!val(t,"due_at").isEmpty()){gap(info,3);caption(info,itemDate(val(t,"due_at")));}
            line.addView(info,new LinearLayout.LayoutParams(0,-2,1));item.addView(line);
            if(canCompleteTask(t)){gap(item,10);TextView finish=button("✓ Concluir",Color.rgb(255,232,224),Color.rgb(128,54,47),()->{});finish.setOnClickListener(v->completeTask(t,finish));item.addView(finish);}
            shown++;
        }if(shown==0)placeholder("Nenhuma tarefa pendente.");
    }
'''
s=s[:start]+home+s[end:]
print('+ home/dashboard')
# use real in-line post comments. old method remains unused to avoid regression if other calls.
old='buttons.addView(button("Comentários ("+p.optInt("comment_count")+")",Color.rgb(255,238,230),NAVY,()->postComments(p)),new LinearLayout.LayoutParams(0,dp(44),1));c.addView(buttons);gap(c,7);'
new='''LinearLayout inlineReplies=vertical();inlineReplies.setVisibility(View.GONE);
            buttons.addView(button("♧  Comentar ("+p.optInt("comment_count")+")",Color.rgb(255,238,230),Color.rgb(54,61,82),()->{
                if(inlineReplies.getVisibility()==View.VISIBLE){inlineReplies.setVisibility(View.GONE);return;}
                inlineReplies.setVisibility(View.VISIBLE);renderInlineComments(p,inlineReplies);
            }),new LinearLayout.LayoutParams(0,dp(44),1));c.addView(buttons);gap(c,7);c.addView(inlineReplies);'''
change(old,new,'comentários embutidos')
insert_before='    private void postComments(JSONObject p){'
method=r'''    private void renderInlineComments(JSONObject p,LinearLayout target){
        target.removeAllViews();target.setPadding(dp(10),dp(12),dp(10),dp(12));
        text(target,"Comentários",14,NAVY,true);gap(target,8);
        LinearLayout messages=vertical();target.addView(messages);
        EditText reply=edit("Escreva seu comentário...",false);reply.setSingleLine(false);reply.setMaxLines(4);
        api.get("post_comments",data("post_id",val(p,"id")),(j,error)->{
            if(target.getParent()==null)return;
            messages.removeAllViews();
            if(error!=null){caption(messages,"Não foi possível carregar: "+error);return;}
            JSONArray replies=array(j,"items");if(replies.length()==0)caption(messages,"Seja a primeira pessoa a comentar.");
            for(int i=0;i<replies.length();i++){
                JSONObject comment=row(replies,i);LinearLayout bubble=vertical();
                bubble.setPadding(dp(12),dp(9),dp(12),dp(9));bubble.setBackground(shape("dark".equals(vault.get("theme"))?Color.rgb(44,60,78):Color.rgb(246,245,249),13));
                text(bubble,val(comment,"name"),12,NAVY,true);gap(bubble,4);text(bubble,val(comment,"body"),14,NAVY,false);
                LinearLayout.LayoutParams bp=new LinearLayout.LayoutParams(-1,-2);bp.bottomMargin=dp(7);messages.addView(bubble,bp);
            }
        });
        gap(target,10);LinearLayout compose=horizontal();compose.addView(reply,new LinearLayout.LayoutParams(0,dp(50),1));gapH(compose,8);
        TextView send=button("➤",PINK,WHITE,()->{});send.setOnClickListener(v->{String body=reply.getText().toString().trim();if(body.isEmpty())return;
            send.setEnabled(false);api.post("post_comment_add",data("post_id",val(p,"id"),"body",body),(r,e)->{
                send.setEnabled(true);if(e!=null){msg(e);return;}reply.setText("");renderInlineComments(p,target);
            });});compose.addView(send,new LinearLayout.LayoutParams(dp(52),dp(50)));target.addView(compose);
    }
'''
assert insert_before in s
s=s.replace(insert_before,method+insert_before,1)
# Private chat inline rather than pop-up form: replace last addButton for sending in renderPrivateMessages
old='gap(content,12);addButton(content,"Enviar mensagem privada",PINK,WHITE,()->formDialog("Mensagem para "+privateTitle,new String[]{"Mensagem"},new String[]{""},a1->{mutate("private_chat_send",data("conversation_id",String.valueOf(privateCid),"body",a1[0]),false);}));'
new='''gap(content,12);LinearLayout compose=horizontal();compose.setPadding(dp(5),dp(7),dp(5),dp(7));
        EditText message=edit("Escreva sua mensagem...",false);message.setSingleLine(false);message.setMaxLines(5);
        compose.addView(message,new LinearLayout.LayoutParams(0,dp(54),1));gapH(compose,8);
        TextView send=button("➤",PINK,WHITE,()->{});send.setContentDescription("Enviar mensagem privada");send.setOnClickListener(v->{String body=message.getText().toString().trim();if(body.isEmpty())return;
            send.setEnabled(false);JSONObject payload=data("conversation_id",String.valueOf(privateCid),"body",body);
            api.post("private_chat_send",withHid(payload),(result,error)->{
                send.setEnabled(true);if(error!=null){msg(error);return;}message.setText("");loadPage(false);
            });});compose.addView(send,new LinearLayout.LayoutParams(dp(53),dp(54)));content.addView(compose);'''
change(old,new,'resposta privada na tela')
# family group options: already has member data, simplify header styles by styling in UI and leave data logic
change('private void members(JSONObject j){sectionTitle("Nossa família","Integrantes e convites do grupo ativo: "+householdName);',
'private void members(JSONObject j){sectionTitle("Nossa família ♡","Integrantes e convites · "+householdName);','família')
# style and text contrast in dark previews: buttons colors based on fill and actual background
change('private int readableTextColor(int bg,int preferred){',
'''private int readableTextColor(int bg,int preferred){''','contraste calculado') if False else None
# chat cards with sender bubble lighter/darker relevant to themes
change('private void renderPrivateMessages(JSONObject j){JSONArray a=array(j,"items");for(int i=0;i<a.length();i++){JSONObject m=row(a,i);LinearLayout msgCard=card(content);',
'''private void renderPrivateMessages(JSONObject j){JSONArray a=array(j,"items");for(int i=0;i<a.length();i++){JSONObject m=row(a,i);LinearLayout msgCard=card(content);boolean mine=userId.equals(val(m,"sender_id"));
        msgCard.setBackground(shape(mine?("dark".equals(vault.get("theme"))?Color.rgb(56,48,67):Color.rgb(255,229,220)):SURFACE,18));''','bolhas de mensagem')
# Native version number to 2.0.0
# Keep current s from patch; restructure tasks and family to compact UI.
old='private boolean onlyMineTasks=false;private long privateCid=0;'
new='private boolean onlyMineTasks=false,onlyDoneTasks=false;private String familyTab="members";private long privateCid=0;'
assert old in s
s=s.replace(old,new,1)

taskstart=s.index('    private void tasks(JSONObject j){')
taskend=s.index('    private boolean canCompleteTask(',taskstart)
tasks=r'''    private void tasks(JSONObject j){
        sectionTitle("Tarefas da família","Organize, compartilhe e finalize os combinados.");
        addButton(content,"＋ Nova tarefa",PINK,WHITE,this::newTask);gap(content,11);
        LinearLayout chips=horizontal();String[] names={"Todas","Para mim","Concluídas"};
        int selection=onlyDoneTasks?2:onlyMineTasks?1:0;
        for(int z=0;z<3;z++){final int k=z;TextView chip=button(names[z],z==selection?PINK:MINT,z==selection?WHITE:NAVY,()->{
            onlyMineTasks=k==1;onlyDoneTasks=k==2;loadPage(true);
        });LinearLayout.LayoutParams cp=new LinearLayout.LayoutParams(0,dp(42),1);cp.rightMargin=dp(z==2?0:6);chips.addView(chip,cp);}content.addView(chips);gap(content,13);
        JSONArray list=array(j,"tasks"),subs=array(j,"subtasks");int shown=0;
        for(int i=0;i<list.length();i++){
            JSONObject t=row(list,i);boolean done="done".equals(val(t,"status"));if(onlyMineTasks&&!userId.equals(val(t,"assigned_to")))continue;if(onlyDoneTasks&&!done)continue;
            LinearLayout c=card(content);LinearLayout heading=horizontal();
            TextView check=button(done?"✓":"○",done?MINT:Color.rgb(255,233,227),done?Color.rgb(37,130,100):Color.rgb(197,75,65),()->{});
            heading.addView(check,new LinearLayout.LayoutParams(dp(45),dp(44)));gapH(heading,10);
            LinearLayout info=vertical();TextView title=label(val(t,"title"),15,NAVY,true);title.setMaxLines(3);info.addView(title);
            gap(info,4);String due=val(t,"due_at");caption(info,(due.isEmpty()?"Sem prazo":itemDate(due))+(done?"  ·  Concluída":val(t,"assigned_to").isEmpty()?"":"  ·  Atribuída"));
            heading.addView(info,new LinearLayout.LayoutParams(0,-2,1));
            TextView more=button("⋮",MINT,NAVY,()->{});heading.addView(more,new LinearLayout.LayoutParams(dp(42),dp(44)));c.addView(heading);
            LinearLayout detail=vertical();detail.setVisibility(View.GONE);c.addView(detail);
            if(!val(t,"description").isEmpty()){gap(detail,8);text(detail,val(t,"description"),13,MUTED,false);}
            int subtasks=0;for(int k=0;k<subs.length();k++){JSONObject sub=row(subs,k);if(!val(sub,"task_id").equals(val(t,"id")))continue;
                subtasks++;LinearLayout item=horizontal();TextView box=label(sub.optInt("checked",0)==1?"✓":"○",20,PINK,true);item.addView(box);gapH(item,10);
                text(item,val(sub,"title"),13,NAVY,false);detail.addView(item);gap(detail,6);
                final JSONObject step=sub;item.setOnClickListener(v->mutate("task_subtask_toggle",data("id",val(step,"id"),"checked",step.optInt("checked",0)==1?"0":"1"),true));
            }
            final int subtotal=subtasks;
            more.setOnClickListener(v->{String[] options={"Editar tarefa","Adicionar subtarefa",subtotal>0?"Mostrar / ocultar etapas":"Ver detalhes","Excluir tarefa"};
                dialogBuilder().setTitle(val(t,"title")).setItems(options,(di,which)->{
                    if(which==0)editTask(t);
                    else if(which==1)formDialog("Nova subtarefa",new String[]{"Item ou etapa"},new String[]{""},xs->{if(!xs[0].trim().isEmpty())mutate("task_subtask_add",data("task_id",val(t,"id"),"title",xs[0].trim()),true);});
                    else if(which==2)detail.setVisibility(detail.getVisibility()==View.VISIBLE?View.GONE:View.VISIBLE);
                    else dialogBuilder().setTitle("Excluir esta tarefa?").setMessage("A tarefa será removida da família. Requer internet.")
                        .setNegativeButton("Cancelar",null).setPositiveButton("Excluir",(dd,w)->mutate("task_delete",data("id",val(t,"id")),false)).show();
                }).show();
            });
            if(!done&&canCompleteTask(t)){gap(c,9);TextView finish=button(pendingCompletion(val(t,"id"))?"⏳ Aguardando sincronização":"✓ Concluir",MINT,NAVY,()->{});
                finish.setOnClickListener(v->{if(pendingCompletion(val(t,"id")))syncQueue(true);else completeTask(t,finish);});c.addView(finish,new LinearLayout.LayoutParams(-1,dp(41)));
            }else if(done){gap(c,6);caption(c,"✓ Concluída por integrante autorizado");}
            shown++;
        }
        if(shown==0)placeholder(onlyDoneTasks?"Nenhuma tarefa concluída.":onlyMineTasks?"Nenhuma tarefa atribuída a você.":"Ainda não há tarefas nesta família.");
        showPendingFor("task_add","Tarefa local");
    }
'''
s=s[:taskstart]+tasks+s[taskend:]

famstart=s.index('    private void members(JSONObject j){')
famend=s.index('    private void loadInvites()',famstart)
family=r'''    private void members(JSONObject j){
        sectionTitle("Nossa família ♡","Grupo: "+householdName);
        if(householdId.isEmpty()){placeholder("Escolha uma família.");addButton(content,"Selecionar família",MINT,NAVY,()->householdsScreen());return;}
        LinearLayout tabs=horizontal();String[][] choices={{"group","Grupo"},{"members","Integrantes"},{"invites","Convites"}};
        for(int i=0;i<choices.length;i++){final String key=choices[i][0];boolean active=key.equals(familyTab);
            TextView tab=button(choices[i][1],active?PINK:MINT,active?WHITE:NAVY,()->{familyTab=key;loadPage(false);});
            LinearLayout.LayoutParams lp=new LinearLayout.LayoutParams(0,dp(43),1);lp.rightMargin=dp(i==2?0:5);tabs.addView(tab,lp);
        }content.addView(tabs);gap(content,13);
        JSONArray list=array(j,"items");
        if("group".equals(familyTab)){
            LinearLayout summary=card(content);text(summary,"♥  "+householdName,21,NAVY,true);gap(summary,10);
            caption(summary,list.length()+" integrantes  ·  Acesso protegido");gap(summary,10);
            addButton(summary,"Ver integrantes",MINT,NAVY,()->{familyTab="members";loadPage(false);});gap(summary,7);
            addButton(summary,"Ver convites",MINT,NAVY,()->{familyTab="invites";loadPage(false);});
            gap(content,12);addButton(content,"Trocar de família",MINT,NAVY,()->{householdId="";vault.remove("hid");refreshMe();});return;
        }
        if("invites".equals(familyTab)){
            if(!"owner".equals(role)&&!"admin".equals(role)){placeholder("Somente administradores podem gerar convites.");return;}
            addButton(content,"+ Gerar convite de uso único",PINK,WHITE,()->{
                dialogBuilder().setTitle("Convidar integrante").setItems(new String[]{"Adulto","Administrador","Dependente","Cuidador"},(d,idx)->{
                    String[] roles={"adult","admin","child","caregiver"};
                    api.post("invite_create",withHid(data("role",roles[idx])),(res,err)->{
                        if(err!=null){msg(err);return;}
                        String code=val(res,"code");dialogBuilder().setTitle("Convite de uso único")
                            .setMessage("Código: "+code+"\nExpira em "+val(res,"expires_in_hours")+" horas.\nCompartilhe apenas com a pessoa convidada.")
                            .setPositiveButton("Copiar",(dd,w)->{android.content.ClipboardManager cb=(android.content.ClipboardManager)getSystemService(CLIPBOARD_SERVICE);
                                cb.setPrimaryClip(ClipData.newPlainText("Convite Nosso Combinado",code));msg("Código copiado.");loadPage(false);})
                            .setNegativeButton("Fechar",null).show();
                    });
                }).show();
            });gap(content,12);loadInvites();return;
        }
        if(list.length()==0){placeholder("Nenhum integrante encontrado neste grupo.");addButton(content,"Atualizar",MINT,NAVY,()->loadPage(false));return;}
        for(int i=0;i<list.length();i++){
            JSONObject m=row(list,i);LinearLayout c=card(content);LinearLayout row=horizontal();
            if(m.optLong("avatar_media_id")>0){ImageView avatar=makeMediaImage("/api/index.php?action=media_get&id="+m.optLong("avatar_media_id"),48,true,null);
                LinearLayout.LayoutParams al=new LinearLayout.LayoutParams(dp(48),dp(48));al.rightMargin=dp(12);row.addView(avatar,al);
            }else{TextView avatar=label("♧",23,PINK,true);avatar.setGravity(Gravity.CENTER);avatar.setBackground(shape(MINT,25));
                LinearLayout.LayoutParams ap=new LinearLayout.LayoutParams(dp(48),dp(48));ap.rightMargin=dp(12);row.addView(avatar,ap);}
            LinearLayout info=vertical();text(info,val(m,"name")+(val(m,"id").equals(userId)?" · Eu":""),15,NAVY,true);
            gap(info,4);caption(info,relationshipText(val(m,"family_relationship"))+"  ·  "+val(m,"role"));row.addView(info,new LinearLayout.LayoutParams(0,-2,1));
            TextView menu=button("⋮",MINT,NAVY,()->{});row.addView(menu,new LinearLayout.LayoutParams(dp(43),dp(44)));c.addView(row);
            menu.setOnClickListener(v->{ArrayList<String> options=new ArrayList<>();ArrayList<Runnable> acts=new ArrayList<>();
                if(val(m,"id").equals(userId)||"owner".equals(role)||"admin".equals(role)){options.add("Definir parentesco");acts.add(()->pickRelationship(m));}
                if(!val(m,"id").equals(userId)&&!"caregiver".equals(val(m,"role"))){options.add("Mensagem privada");acts.add(()->api.post("private_chat_open",withHid(data("other_user_id",val(m,"id"))),(res,e)->{if(e!=null){msg(e);return;}openPrivate(res.optLong("conversation_id"),val(m,"name"));}));}
                if(("owner".equals(role)||"admin".equals(role))&&!val(m,"id").equals(userId)&&!"owner".equals(val(m,"role"))){options.add("Remover integrante");acts.add(()->dialogBuilder().setTitle("Remover integrante?").setNegativeButton("Cancelar",null).setPositiveButton("Remover",(d,w)->mutate("member_remove",data("user_id",val(m,"id")),false)).show());}
                if(options.isEmpty()){msg("Nenhuma ação disponível.");return;}
                dialogBuilder().setTitle(val(m,"name")).setItems(options.toArray(new String[0]),(d,w)->acts.get(w).run()).show();
            });
        }
        gap(content,8);addButton(content,"Trocar de família",MINT,NAVY,()->{householdId="";vault.remove("hid");refreshMe();});
    }
'''
s=s[:famstart]+family+s[famend:]
# Show invite list only while user is inside invites tab (avoid asynchronous overpaint)
s=s.replace('if(!"members".equals(page)||!h.equals(householdId)||content==null)return;', 'if(!"members".equals(page)||!"invites".equals(familyTab)||!h.equals(householdId)||content==null)return;',1)
print('+ tasks compact / subtasks menu / family tabs')

f.write_text(s)
grad=root/'app/build.gradle';g=grad.read_text();assert 'versionCode 190' in g and "versionName '1.9.0'" in g
g=g.replace('versionCode 190','versionCode 200').replace("versionName '1.9.0'","versionName '2.0.0'");grad.write_text(g)
print('Updated v2.0 native code',len(s))