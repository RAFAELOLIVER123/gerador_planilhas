#!/usr/bin/env python3
"""Atualização incremental nativa 1.8 sobre o pacote 1.7 do Nosso Combinado."""
from pathlib import Path
import re, sys
r=Path(sys.argv[1]); main=r/'app/src/main/java/com/nossocombinado/app/MainActivity.java'; grad=r/'app/build.gradle'
s=main.read_text(); steps=[]
def replace(old,new,label):
 global s
 if old not in s:raise Exception('Não encontrado '+label+' / '+old[:80])
 s=s.replace(old,new,1); steps.append(label)
def regex(pat,repl,label):
 global s
 s,n=re.subn(pat,lambda m:repl,s,count=1,flags=re.S)
 if n!=1:raise Exception('Regex não encontrado '+label)
 steps.append(label)

# Cache de mídia por usuário para evitar recarregar fotos nos ciclos de sincronização.
replace('private String uploadMode="",uploadEntity="",uploadTitle="";', '''private final android.util.LruCache<String,android.graphics.Bitmap> mediaCache=new android.util.LruCache<String,android.graphics.Bitmap>(12*1024*1024){
        @Override protected int sizeOf(String key,android.graphics.Bitmap value){return value.getByteCount();}
    };
    private String uploadMode="",uploadEntity="",uploadTitle="";''','cache protegido de fotos')
# Logomarca usada no site (res/drawable-nodpi/nosso_icon.png).
replace('TextView logo=label("♥",27,PINK,true);logo.setPadding(0,0,dp(7),0);topBar.addView(logo);', '''ImageView logo=new ImageView(this);logo.setImageResource(R.drawable.nosso_icon);logo.setContentDescription("Nosso Combinado");
        LinearLayout.LayoutParams logoParams=new LinearLayout.LayoutParams(dp(37),dp(37));logoParams.rightMargin=dp(9);topBar.addView(logo,logoParams);''','logo do site no cabeçalho')
replace('TextView heart=label("♥",62,PINK,true);heart.setGravity(Gravity.CENTER);c.addView(heart);','''ImageView heart=new ImageView(this);heart.setImageResource(R.drawable.nosso_icon);heart.setContentDescription("Logo do Nosso Combinado");
        LinearLayout.LayoutParams iconParams=new LinearLayout.LayoutParams(dp(86),dp(86));iconParams.gravity=Gravity.CENTER;c.addView(heart,iconParams);''','logo site login')

# Estados de conclusão verificáveis, sem reabertura. Atribuído e adultos podem concluir conforme regra do PHP.
replace('if(!"done".equals(val(t,"status")))buttons.addView(button("Concluir",Color.rgb(255,238,230),NAVY,()->mutate("task_update",data("id",val(t,"id"),"status","done"),true)),new LinearLayout.LayoutParams(0,dp(42),1));c.addView(buttons);', '''if(!"done".equals(val(t,"status"))){
                if(pendingCompletion(val(t,"id"))){TextView queued=button("⏳ Aguardando envio",Color.rgb(251,240,209),NAVY,()->syncQueue(true));buttons.addView(queued,new LinearLayout.LayoutParams(0,dp(42),1));}
                else if(canCompleteTask(t)){TextView finish=button("✓ Concluir",PINK,WHITE,()->{});finish.setOnClickListener(v->completeTask(t,finish));buttons.addView(finish,new LinearLayout.LayoutParams(0,dp(42),1));}
                else {TextView assigned=label("Atribuída a outro integrante",12,MUTED,false);buttons.addView(assigned,new LinearLayout.LayoutParams(0,dp(42),1));}
            }else {TextView done=label("✓ Finalizada",13,Color.rgb(40,128,91),true);done.setGravity(Gravity.CENTER);buttons.addView(done,new LinearLayout.LayoutParams(0,dp(42),1));}
            c.addView(buttons);''','finalizar tarefa com confirmação')
replace('int shown=0;for(int i=0;i<tasks.length()&&shown<4;i++){JSONObject t=row(tasks,i);if("done".equals(val(t,"status")))continue;itemCard("✓ "+val(t,"title"),val(t,"due_at"),()->editTask(t));shown++;}if(shown==0)placeholder("Nenhuma tarefa cadastrada.");', '''int shown=0;for(int i=0;i<tasks.length()&&shown<4;i++){JSONObject t=row(tasks,i);if("done".equals(val(t,"status")))continue;
            LinearLayout taskCard=card(content);text(taskCard,"☐  "+val(t,"title"),15,NAVY,true);if(!val(t,"due_at").isEmpty())caption(taskCard,"Prazo: "+itemDate(val(t,"due_at")));
            if(pendingCompletion(val(t,"id")))caption(taskCard,"⏳ Conclusão aguardando sincronização");
            else if(canCompleteTask(t)){gap(taskCard,8);TextView finish=button("✓ Concluir tarefa",PINK,WHITE,()->{});finish.setOnClickListener(v->completeTask(t,finish));taskCard.addView(finish,new LinearLayout.LayoutParams(-1,dp(43)));}
            shown++;}if(shown==0)placeholder("Nenhuma tarefa pendente.");''','concluir direto no início')
insert_before='    private void gapH(LinearLayout row,int n)'
completion='''    private boolean canCompleteTask(JSONObject t){String r=role;return "owner".equals(r)||"admin".equals(r)||"adult".equals(r)||userId.equals(val(t,"assigned_to"));}
    private boolean pendingCompletion(String taskId){for(JSONObject op:offline.pending(userId,householdId)){
        if("task_update".equals(val(op,"action"))){JSONObject p=op.optJSONObject("payload");if(p!=null&&taskId.equals(val(p,"id"))&&"done".equals(val(p,"status")))return true;}
    }return false;}
    private void completeTask(JSONObject t,TextView finish){String id=val(t,"id");if(id.isEmpty()||"done".equals(val(t,"status")))return;
        if(!canCompleteTask(t)){msg("Apenas um adulto autorizado ou a pessoa responsável pode concluir esta tarefa.");return;}
        if(pendingCompletion(id)){msg("Essa conclusão já aguarda sincronização.");return;}
        JSONObject payload=withHid(data("id",id,"status","done"));
        if(!hasConnection()){
            offline.enqueue(userId,householdId,UUID.randomUUID().toString(),"task_update",payload);
            msg("Conclusão guardada no aparelho. Será enviada quando houver conexão.");loadPage(true);return;
        }
        finish.setEnabled(false);finish.setAlpha(0.65f);finish.setText("Concluindo…");String previousUser=userId,previousHousehold=householdId;
        api.post("task_update",payload,(response,error)->{
            if(!previousUser.equals(userId)||!previousHousehold.equals(householdId))return;
            if(error!=null){finish.setText("✓ Concluir");finish.setEnabled(true);finish.setAlpha(1f);msg("Não foi possível concluir: "+error);return;}
            rememberTaskCompletion(id,response);msg("✓ Tarefa concluída e salva para a família!");loadPage(true);
        });
    }
    private void rememberTaskCompletion(String id,JSONObject result){for(String cachedPage:new String[]{"home","tasks"}){
        try{String raw=offline.load(userId,householdId,cachedPage);if(raw.isEmpty())continue;JSONObject snapshot=new JSONObject(raw);JSONArray list=array(snapshot,"tasks");
            for(int i=0;i<list.length();i++){JSONObject t=row(list,i);if(id.equals(val(t,"id"))){t.put("status","done");t.put("completed_at",val(result,"completed_at"));t.put("completed_by",result.optString("completed_by",userId));}}
            offline.save(userId,householdId,cachedPage,snapshot.toString());
        }catch(Exception ignored){}
    }}
'''
replace(insert_before,completion+insert_before,'método de conclusão e atualização de cache')

# Rede social: avatar do usuário ao lado do post e mídia existente, inclusive imagem legada.
replace('LinearLayout header=horizontal();LinearLayout author=vertical();', '''LinearLayout header=horizontal();if(p.optLong("avatar_media_id")>0){ImageView authorAvatar=makeMediaImage("/api/index.php?action=media_get&id="+p.optLong("avatar_media_id"),46,true,null);LinearLayout.LayoutParams avatarParams=new LinearLayout.LayoutParams(dp(46),dp(46));avatarParams.rightMargin=dp(10);header.addView(authorAvatar,avatarParams);}LinearLayout author=vertical();''','avatares rede')
replace('JSONArray photos=array(p,"media");for(int k=0;k<photos.length();k++){JSONObject photo=row(photos,k);if("image".equals(val(photo,"kind")))postImage(c,photo);}', '''JSONArray photos=array(p,"media");int countPhotos=0;
            for(int k=0;k<photos.length();k++){JSONObject photo=row(photos,k);if("image".equals(val(photo,"kind"))){postImage(c,photo);countPhotos++;}}
            if(countPhotos==0&&!val(p,"legacy_image_url").isEmpty())postImage(c,val(p,"legacy_image_url"));''','imagem legada e múltiplas fotos')
# Chat familiar e privado também devem exibir fotos anexadas no site.
replace('text(c,val(m,"body"),15,NAVY,false);gap(c,5);caption(c,itemDate(val(m,"sent_at")));', '''text(c,val(m,"body"),15,NAVY,false);if("image".equals(val(m,"media_kind"))&&!val(m,"media_id").isEmpty())postImage(c,"/api/index.php?action=media_get&id="+val(m,"media_id"));gap(c,5);caption(c,itemDate(val(m,"sent_at")));''','imagens do chat familiar')
replace('itemCard(val(m,"sender_name")+" · "+itemDate(val(m,"sent_at")),val(m,"body"),null);', '''LinearLayout msgCard=card(content);text(msgCard,val(m,"sender_name")+" · "+itemDate(val(m,"sent_at")),13,PINK,true);gap(msgCard,5);text(msgCard,val(m,"body"),15,NAVY,false);if("image".equals(val(m,"media_kind"))&&!val(m,"media_id").isEmpty())postImage(msgCard,"/api/index.php?action=media_get&id="+val(m,"media_id"));''','imagens do chat privado')
replace('LinearLayout c=card(content);text(c,"♡  "+val(m,"name"),17,NAVY,true);gap(c,5);', '''LinearLayout c=card(content);LinearLayout famHeader=horizontal();if(m.optLong("avatar_media_id")>0){ImageView avatar=makeMediaImage("/api/index.php?action=media_get&id="+m.optLong("avatar_media_id"),48,true,null);LinearLayout.LayoutParams avp=new LinearLayout.LayoutParams(dp(48),dp(48));avp.rightMargin=dp(10);famHeader.addView(avatar,avp);}TextView memberName=label(val(m,"name"),17,NAVY,true);famHeader.addView(memberName,new LinearLayout.LayoutParams(0,-2,1));c.addView(famHeader);gap(c,5);''','fotos da família')

image_methods='''    private void postImage(LinearLayout parent,JSONObject media){String id=val(media,"id");if(id.isEmpty())return;postImage(parent,"/api/index.php?action=media_get&id="+id);}
    private void postImage(LinearLayout parent,String path){if(path==null||path.isEmpty())return;
        LinearLayout imageArea=vertical();imageArea.setPadding(0,dp(7),0,dp(7));parent.addView(imageArea);
        TextView status=label("Carregando foto…",12,MUTED,false);ImageView photo=makeMediaImage(path,238,false,status);
        imageArea.addView(photo,new LinearLayout.LayoutParams(-1,dp(238)));gap(imageArea,4);imageArea.addView(status);
        photo.setOnClickListener(v->{android.graphics.Bitmap bitmap=mediaCache.get(userId+"|"+path);if(bitmap!=null)showImageViewer(bitmap);else loadProtectedImage(photo,path,status);});
        status.setOnClickListener(v->loadProtectedImage(photo,path,status));
    }
    private ImageView makeMediaImage(String path,int size,boolean avatar,TextView status){ImageView iv=new ImageView(this);iv.setImageResource(R.drawable.nosso_icon);
        iv.setScaleType(avatar?ImageView.ScaleType.CENTER_CROP:ImageView.ScaleType.FIT_CENTER);iv.setBackground(shape(Color.rgb(250,245,242),avatar?size/2f:15f));
        iv.setClipToOutline(true);loadProtectedImage(iv,path,status);return iv;
    }
    private void showImageViewer(android.graphics.Bitmap bitmap){ImageView zoom=new ImageView(this);zoom.setImageBitmap(bitmap);zoom.setAdjustViewBounds(true);
        zoom.setScaleType(ImageView.ScaleType.FIT_CENTER);LinearLayout holder=vertical();holder.setPadding(dp(5),dp(8),dp(5),dp(8));holder.addView(zoom,new LinearLayout.LayoutParams(-1,dp(Math.min(590,((int)(getResources().getDisplayMetrics().heightPixels/getResources().getDisplayMetrics().density))*.69f))));
        dialogBuilder().setTitle("Foto · Nosso Combinado").setView(holder).setPositiveButton("Fechar",null).show();
    }
    private void loadProtectedImage(ImageView iv,String path,TextView status){if(path==null||path.trim().isEmpty())return;String key=userId+"|"+path;
        android.graphics.Bitmap cached=mediaCache.get(key);if(cached!=null){iv.setImageBitmap(cached);if(status!=null)status.setVisibility(View.GONE);return;}
        final String url;try{
            java.net.URI requested=new java.net.URI(path);String uriPath=requested.getPath();String query=requested.getQuery();
            if(!"/api/index.php".equals(uriPath)||query==null||!(query.startsWith("action=media_get&")||query.startsWith("action=professional_post_image&")))throw new IOException("Origem de imagem não permitida");
            if(requested.isAbsolute()&&(!"https".equalsIgnoreCase(requested.getScheme())||!"violet-grasshopper-887660.hostingersite.com".equalsIgnoreCase(requested.getHost())))throw new IOException("Endereço não autorizado");
            url=ApiClient.BASE+"?"+query;
        }catch(Exception e){if(status!=null)status.setText("Imagem indisponível.");return;}
        String startedUser=userId;iv.setTag(key);String cookie=vault.get("session_cookie");
        new Thread(()->{java.net.HttpURLConnection conn=null;try{
            conn=(java.net.HttpURLConnection)new java.net.URL(url).openConnection();conn.setUseCaches(false);conn.setConnectTimeout(12000);conn.setReadTimeout(16000);conn.setRequestProperty("Accept","image/jpeg,image/png,image/webp");
            if(!cookie.isEmpty())conn.setRequestProperty("Cookie","nc_session="+cookie);
            if(conn.getResponseCode()!=200)throw new IOException("HTTP "+conn.getResponseCode());
            String type=conn.getContentType();if(type==null||!type.toLowerCase(Locale.ROOT).startsWith("image/"))throw new IOException("Resposta não é imagem");
            byte[] data;try(InputStream stream=conn.getInputStream();ByteArrayOutputStream out=new ByteArrayOutputStream()){
                byte[] chunk=new byte[8192];int n;while((n=stream.read(chunk))!=-1){out.write(chunk,0,n);if(out.size()>6*1024*1024)throw new IOException("Foto acima do limite");}data=out.toByteArray();}
            android.graphics.BitmapFactory.Options bounds=new android.graphics.BitmapFactory.Options();bounds.inJustDecodeBounds=true;android.graphics.BitmapFactory.decodeByteArray(data,0,data.length,bounds);
            int sample=1;while(bounds.outWidth/sample>1600||bounds.outHeight/sample>1600)sample*=2;
            android.graphics.BitmapFactory.Options opts=new android.graphics.BitmapFactory.Options();opts.inSampleSize=sample;android.graphics.Bitmap bitmap=android.graphics.BitmapFactory.decodeByteArray(data,0,data.length,opts);
            if(bitmap==null)throw new IOException("Imagem inválida");mediaCache.put(key,bitmap);
            runOnUiThread(()->{if(startedUser.equals(userId)&&key.equals(iv.getTag())){iv.setImageBitmap(bitmap);if(status!=null)status.setVisibility(View.GONE);}});
        }catch(Exception ex){runOnUiThread(()->{if(startedUser.equals(userId)&&key.equals(iv.getTag())&&status!=null){status.setVisibility(View.VISIBLE);status.setText("Foto não carregada · toque para tentar novamente");}});
        }finally{if(conn!=null)conn.disconnect();}}).start();
    }
'''
regex(r'    private void postImage\(LinearLayout parent,JSONObject media\)\{.*?\n    private String relationshipText\(',image_methods+'    private String relationshipText(', 'visualizador seguro de imagens')
replace('private void logout(){if(userId.isEmpty())return;String old=userId;', 'private void logout(){if(userId.isEmpty())return;mediaCache.evictAll();String old=userId;','limpar fotos no logout')

# Incluir upload de fotos dos posts no mesmo endpoint do site, sem mudar banco.
replace('gap(c,4);caption(c,postTypeLabel(val(p,"post_type")));gap(c,9);', '''gap(c,4);caption(c,postTypeLabel(val(p,"post_type")));if(mine){gap(c,7);addButton(c,"📷 Adicionar foto ao post",MINT,NAVY,()->{
                uploadMode="media_upload";uploadEntity="professional_post:"+val(p,"id");pickFile();});}gap(c,9);''','anexar foto aos posts')
replace('if(uploadMode.equals("avatar_upload")&&mime.equals("application/pdf")){msg("A foto do perfil deve ser uma imagem.");return;}', '''if((uploadMode.equals("avatar_upload")||uploadEntity.startsWith("professional_post:"))&&mime.equals("application/pdf")){msg("Selecione uma foto JPG, PNG ou WEBP.");return;}''','fotos sem pdf')
replace('else if(uploadMode.equals("activity_attachment_upload")){p.put("entity_kind",uploadEntity.split(":")[0]);p.put("entity_id",uploadEntity.split(":")[1]);}', '''else if(uploadMode.equals("activity_attachment_upload")){p.put("entity_kind",uploadEntity.split(":")[0]);p.put("entity_id",uploadEntity.split(":")[1]);}
            else if(uploadMode.equals("media_upload")&&uploadEntity.startsWith("professional_post:")){p.put("scope","professional_post");p.put("post_id",uploadEntity.split(":")[1]);}''','upload media post')
replace('if(e!=null)msg(e);else{msg(uploadMode.equals("avatar_upload")?"Foto de perfil atualizada.":"Documento enviado.");if(content!=null)loadPage(true);}', '''if(e!=null)msg(e);else{msg(uploadMode.equals("avatar_upload")?"Foto de perfil atualizada.":uploadMode.equals("media_upload")?"Foto publicada na Rede.":"Documento enviado.");if(content!=null)loadPage(true);}''','aviso foto enviada')

main.write_text(s)
b=grad.read_text();assert 'versionCode 170' in b and "versionName '1.7.0'" in b
b=b.replace('versionCode 170','versionCode 180').replace("versionName '1.7.0'","versionName '1.8.0'")
grad.write_text(b)
print('Atualização 1.8 aplicada: '+'; '.join(steps))