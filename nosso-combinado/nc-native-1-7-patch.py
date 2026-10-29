#!/usr/bin/env python3
"""Aplica atualizacoes nativas verificaveis ao codigo v1.6 no runner GitHub."""
from pathlib import Path
import re, sys
root=Path(sys.argv[1] if len(sys.argv)>1 else '/mnt/data/nc_native16src')
src=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
s=src.read_text(encoding='utf8')
changed=[]

def rep(old,new,name):
 global s
 if old not in s: raise SystemExit('Padrão não encontrado: '+name)
 s=s.replace(old,new,1);changed.append(name)

# Guardar a ultima familia ativa e atualizar notificacoes/novos posts.
rep('private long privateCid=0;', 'private boolean onlyMineTasks=false;private long privateCid=0;', 'filtro de tarefas')
rep('page.equals("shopping")))loadPage(false);','page.equals("shopping")||page.equals("network")||page.equals("notes")||page.equals("pharmacy")||page.equals("school")||page.equals("vehicle")||page.equals("bills")))loadPage(false);','atualizacao de secoes')
rep('if(action.equals("dashboard"))ReminderScheduler.fromDashboard(this,j,userId,householdId);drawPage(j,false);','if(action.equals("dashboard"))ReminderScheduler.fromDashboard(this,j,userId,householdId);int prevY=scroll!=null && page.equals("network")?scroll.getScrollY():-1;drawPage(j,false);if(prevY>=0&&scroll!=null)scroll.post(()->scroll.scrollTo(0,prevY));','rolagem rede')

rep('addButton(content,"+  Nova tarefa",PINK,WHITE,this::newTask);gap(content,14);pendingBlock();', 'addButton(content,"+  Nova tarefa",PINK,WHITE,this::newTask);gap(content,7);addButton(content,onlyMineTasks?"✓  Para mim   •   Ver todas":"Todas as tarefas   •   Filtrar para mim",MINT,NAVY,()->{onlyMineTasks=!onlyMineTasks;loadPage(true);});gap(content,14);pendingBlock();', 'filtro visual tarefas')
rep('for(int i=0;i<a.length();i++){JSONObject t=row(a,i);LinearLayout c=card(content);','for(int i=0;i<a.length();i++){JSONObject t=row(a,i);if(onlyMineTasks&&!userId.equals(val(t,"assigned_to")))continue;LinearLayout c=card(content);','filtro por responsavel')
rep('gap(c,4);caption(c,"Prazo: "+(val(t,"due_at").isEmpty()?"sem prazo":itemDate(val(t,"due_at")))+"   •   "+val(t,"status"));', 'gap(c,4);caption(c,"Prazo: "+(val(t,"due_at").isEmpty()?"sem prazo":itemDate(val(t,"due_at")))+"   •   "+val(t,"status"));if("done".equals(val(t,"status")))caption(c,"Concluída em "+itemDate(val(t,"completed_at"))+" • responsável ID "+val(t,"completed_by"));else if(!val(t,"assigned_to").isEmpty())caption(c,"Atribuída ao integrante #"+val(t,"assigned_to"));','historico conclusao')
rep('buttons.addView(button("done".equals(val(t,"status"))?"Reabrir":"Concluir",Color.rgb(255,238,230),NAVY,()->mutate("task_update",data("id",val(t,"id"),"status","done".equals(val(t,"status"))?"open":"done"),true)),new LinearLayout.LayoutParams(0,dp(42),1));c.addView(buttons);', 'if(!"done".equals(val(t,"status")))buttons.addView(button("Concluir",Color.rgb(255,238,230),NAVY,()->mutate("task_update",data("id",val(t,"id"),"status","done"),true)),new LinearLayout.LayoutParams(0,dp(42),1));c.addView(buttons);','bloqueio reabertura')
# Dados dos membros pertencem a membership, nao a role.
rep('caption(c,"Papel: "+val(m,"role"));gap(c,10);', 'caption(c,"Acesso: "+val(m,"role")+"   •   Parentesco: "+relationshipText(val(m,"family_relationship")));gap(c,7);if(val(m,"id").equals(userId)||"owner".equals(role)||"admin".equals(role))addButton(c,"Definir parentesco",MINT,NAVY,()->pickRelationship(m));gap(c,10);','parentesco membro')

net=r'''    private void network(JSONObject j){sectionTitle("✦ Rede de experiências","Publicações, serviços e pessoas procurando trabalho.");
        addButton(content,"+ Publicar experiência",PINK,WHITE,()->editProfessionalPost(null,"experience"));gap(content,7);
        addButton(content,"Estou procurando trabalho",MINT,NAVY,()->editProfessionalPost(null,"availability"));gap(content,7);
        addButton(content,"Anunciar serviço ou oportunidade",MINT,NAVY,()->editProfessionalPost(null,"service"));gap(content,13);
        JSONArray a=array(j,"items");if(a.length()==0)placeholder("Seja a primeira pessoa a publicar.");
        for(int i=0;i<a.length();i++){JSONObject p=row(a,i);LinearLayout c=card(content);
            LinearLayout header=horizontal();LinearLayout author=vertical();author.addView(label("●  "+val(p,"name"),16,NAVY,true));author.addView(label(itemDate(val(p,"created_at")),11,MUTED,false));header.addView(author,new LinearLayout.LayoutParams(0,-2,1));
            boolean mine=userId.equals(val(p,"author_user_id"));if(mine){TextView menu=button("⋯",MINT,NAVY,()->postMenu(p));header.addView(menu,new LinearLayout.LayoutParams(dp(47),dp(42)));}
            c.addView(header);gap(c,8);if(!val(p,"post_title").isEmpty()){text(c,val(p,"post_title"),17,NAVY,true);gap(c,5);}
            text(c,val(p,"body"),14,NAVY,false);gap(c,8);
            JSONArray photos=array(p,"media");for(int k=0;k<photos.length();k++){JSONObject photo=row(photos,k);if("image".equals(val(photo,"kind")))postImage(c,photo);}
            gap(c,4);caption(c,postTypeLabel(val(p,"post_type")));gap(c,9);
            LinearLayout buttons=horizontal();buttons.addView(button((p.optInt("liked_by_me")>0?"♥  Curtido":"♡  Curtir")+" ("+p.optInt("like_count")+")",MINT,NAVY,()->api.post("post_like_toggle",data("post_id",val(p,"id")),(r,e)->{if(e!=null)msg(e);else loadPage(false);})),new LinearLayout.LayoutParams(0,dp(44),1));gapH(buttons,7);
            buttons.addView(button("Comentários ("+p.optInt("comment_count")+")",Color.rgb(255,238,230),NAVY,()->postComments(p)),new LinearLayout.LayoutParams(0,dp(44),1));c.addView(buttons);gap(c,7);
            if(p.optLong("professional_id")>0)addButton(c,"Ver perfil profissional",MINT,NAVY,()->openProfessional(data("id",val(p,"professional_id"))));
        }
        gap(content,10);addButton(content,"Buscar profissionais",MINT,NAVY,()->navigate("jobs"));
    }
    private String postTypeLabel(String type){switch(type){case "availability":return "Procurando trabalho";case "service":return "Serviços e oportunidades";case "course":return "Cursos e formação";default:return "Experiência profissional";}}
    private void postMenu(JSONObject p){dialogBuilder().setTitle("Publicação de "+val(p,"name"))
        .setItems(new String[]{"Editar publicação","Excluir publicação"},(d,which)->{if(which==0)editProfessionalPost(p,val(p,"post_type"));else dialogBuilder().setTitle("Excluir esta publicação?").setMessage("A publicação será excluída da Rede para todos.").setNegativeButton("Cancelar",null).setPositiveButton("Excluir",(di,w)->api.post("professional_post_delete",data("id",val(p,"id")),(r,e)->{if(e!=null)msg(e);else loadPage(false);})).show();}).show();}
    private void editProfessionalPost(JSONObject p,String type){formDialog(p==null?"Publicar na Rede":"Editar publicação",new String[]{"Título","Texto da publicação","Tipo (experience/service/course/availability)"},new String[]{val(p,"post_title"),val(p,"body"),type},values->{if(values[0].trim().isEmpty()||values[1].trim().isEmpty()){msg("Preencha o título e o texto.");return;}String kind=values[2].trim().toLowerCase(Locale.ROOT);if(!Arrays.asList("experience","service","course","availability").contains(kind)){msg("Tipo de publicação inválido.");return;}JSONObject body=data("post_title",values[0],"body",values[1],"post_type",kind);if(p!=null)try{body.put("id",p.optLong("id"));}catch(Exception ignored){}api.post(p==null?"professional_post_add":"professional_post_edit",body,(r,e)->{if(e!=null)msg(e);else{msg("Publicação salva.");loadPage(false);}});});}
    private void postComments(JSONObject p){api.get("post_comments",data("post_id",val(p,"id")),(j,e)->{if(e!=null){msg(e);return;}
        LinearLayout body=vertical();body.setPadding(dp(15),dp(10),dp(15),dp(10));JSONArray comments=array(j,"items");if(comments.length()==0)text(body,"Nenhum comentário. Seja o primeiro!",13,MUTED,false);
        for(int i=0;i<comments.length();i++){JSONObject com=row(comments,i);text(body,val(com,"name")+": "+val(com,"body"),14,NAVY,false);gap(body,9);}
        ScrollView sc=new ScrollView(this);sc.addView(body);dialogBuilder().setTitle("Comentários · "+val(p,"post_title")).setView(sc).setNegativeButton("Fechar",null).setPositiveButton("Comentar",(d,w)->formDialog("Novo comentário",new String[]{"Seu comentário"},new String[]{""},values->{String t=values[0].trim();if(t.isEmpty())return;api.post("post_comment_add",data("post_id",val(p,"id"),"body",t),(r,error)->{if(error!=null)msg(error);else{loadPage(false);postComments(p);}});})).show();});}
    private void postImage(LinearLayout parent,JSONObject media){String id=val(media,"id");if(id.isEmpty())return;
        ImageView iv=new ImageView(this);iv.setScaleType(ImageView.ScaleType.CENTER_CROP);iv.setBackgroundColor(Color.rgb(246,244,242));parent.addView(iv,new LinearLayout.LayoutParams(-1,dp(235)));
        String session=vault.get("session_cookie");new Thread(()->{try{java.net.HttpURLConnection con=(java.net.HttpURLConnection)new java.net.URL("https://violet-grasshopper-887660.hostingersite.com/api/index.php?action=media_get&id="+java.net.URLEncoder.encode(id,"UTF-8")).openConnection();con.setConnectTimeout(12000);con.setReadTimeout(15000);if(!session.isEmpty())con.setRequestProperty("Cookie","nc_session="+session);if(con.getResponseCode()!=200)throw new IOException("Imagem indisponível");byte[] bytes;try(InputStream stream=con.getInputStream();ByteArrayOutputStream buf=new ByteArrayOutputStream()){byte[] chunk=new byte[8192];int n;while((n=stream.read(chunk))!=-1){buf.write(chunk,0,n);if(buf.size()>6000000)throw new IOException("Imagem grande demais");}bytes=buf.toByteArray();}android.graphics.Bitmap bitmap=android.graphics.BitmapFactory.decodeByteArray(bytes,0,bytes.length);if(bitmap==null)throw new IOException("Formato de imagem não suportado");runOnUiThread(()->{iv.setImageBitmap(bitmap);iv.setOnClickListener(v->{ImageView zoom=new ImageView(this);zoom.setImageBitmap(bitmap);zoom.setAdjustViewBounds(true);zoom.setScaleType(ImageView.ScaleType.FIT_CENTER);dialogBuilder().setTitle("Foto da publicação").setView(zoom).setPositiveButton("Fechar",null).show();});});}catch(Exception ex){runOnUiThread(()->{iv.setVisibility(View.GONE);});}}).start();}
    private String relationshipText(String value){switch(value){case "wife":return "Esposa";case "husband":return "Marido";case "mother":return "Mãe";case "father":return "Pai";case "daughter":return "Filha";case "son":return "Filho";case "niece":return "Sobrinha";case "nephew":return "Sobrinho";case "sister":return "Irmã";case "brother":return "Irmão";case "grandmother":return "Avó";case "grandfather":return "Avô";case "partner":return "Companheiro(a)";case "friend":return "Amigo(a)";case "other":return "Outro";default:return "Não informado";}}
    private void pickRelationship(JSONObject member){String[][] options={{"wife","Esposa"},{"husband","Marido"},{"mother","Mãe"},{"father","Pai"},{"daughter","Filha"},{"son","Filho"},{"niece","Sobrinha"},{"nephew","Sobrinho"},{"sister","Irmã"},{"brother","Irmão"},{"grandmother","Avó"},{"grandfather","Avô"},{"partner","Companheiro(a)"},{"friend","Amigo(a)"},{"other","Outro"}};String[] labels=new String[options.length];for(int i=0;i<options.length;i++)labels[i]=options[i][1];dialogBuilder().setTitle("Parentesco de "+val(member,"name")).setItems(labels,(dialog,which)->api.post("family_relationship_save",withHid(data("member_user_id",val(member,"id"),"family_relationship",options[which][0])),(j,e)->{if(e!=null)msg(e);else loadPage(false);})).show();}
'''
pattern=r'    private void network\(JSONObject j\)\{.*?\n    private void professionals\(JSONObject j\)'
if len(re.findall(pattern,s,flags=re.S))!=1:raise SystemExit('Não localizou seção nativa Rede')
s=re.sub(pattern,lambda m:net+'    private void professionals(JSONObject j)',s,count=1,flags=re.S);changed.append('rede nativa e familia')
rep('private void createProfessionalPost(){formDialog("Nova publicação",new String[]{"Título","Texto da experiência","Tipo (experience/course/service/availability)"},new String[]{"","","experience"},a->{api.post("professional_post_add",data("post_title",a[0],"body",a[1],"post_type",a[2]),(j,e)->{if(e!=null)msg(e);else msg("Publicação enviada para revisão.");});});}', 'private void createProfessionalPost(){editProfessionalPost(null,"experience");}', 'post novo')
# Corrigir mensagens obsoletas de moderacao visual no perfil.
rep('else msg("Perfil profissional enviado para moderação.");','else msg("Perfil profissional salvo. Aguarde a confirmação de disponibilidade do currículo.");','feedback perfil')
# Persistir a familia correta em logins: manter papel/descricao sempre atualizados.
rep('if(householdId.equals(val(row(hs,i),"id")))valid=true;', 'if(householdId.equals(val(row(hs,i),"id"))){valid=true;chooseHousehold(row(hs,i));}', 'restauracao familia')
src.write_text(s,encoding='utf8')
grad=root/'app/build.gradle';b=grad.read_text();assert 'versionCode 160' in b and "versionName '1.6.0'" in b
b=b.replace('versionCode 160','versionCode 170').replace("versionName '1.6.0'","versionName '1.7.0'")
grad.write_text(b)
print('OK:', ', '.join(changed))