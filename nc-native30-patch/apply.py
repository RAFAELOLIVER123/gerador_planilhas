#!/usr/bin/env python3
"""Nosso Combinado nativo 3.0: exclusão autorizada e interface preservada da 2.9."""
from pathlib import Path
import sys
root=Path(sys.argv[1]).resolve()
main=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
grad=root/'app/build.gradle'
src=main.read_text(encoding='utf-8')
def swap(before,after,label):
 global src
 count=src.count(before)
 if count!=1:raise RuntimeError(label+': ocorrências='+str(count))
 src=src.replace(before,after,1)
 print('OK',label)
swap('private void chat(JSONObject j){sectionTitle("Chat da família","Mensagens entregues somente quando houver internet.");',
 '''private void chat(JSONObject j){sectionTitle("Chat da família","Mensagens entregues somente quando houver internet.");
        final long chatCid=j.optLong("conversation_id",0);
        if(chatCid>0){addButton(content,"🗑  Limpar histórico só para mim",MINT,NAVY,()->confirmClearConversation(chatCid,()->loadPage(false)));gap(content,8);}
 ''', 'limpar chat familiar')
swap('text(c,val(m,"body"),15,NAVY,false);gap(c,5);caption(c,itemDate(val(m,"sent_at")));}',
 '''text(c,val(m,"body"),15,NAVY,false);gap(c,5);caption(c,itemDate(val(m,"sent_at")));
            if(val(m,"sender_id").equals(userId)&&chatCid>0&&m.optLong("id")>0){
                final long mid=m.optLong("id");
                addButton(c,"Excluir minha mensagem",MINT,NAVY,()->confirmDeleteMessage(mid,chatCid,()->loadPage(false)));
            }
        }''', 'excluir mensagem familiar de autoria própria')
swap('itemCard("✉  "+val(c,"other_name"),val(c,"last_message"),()->openPrivate(c.optLong("id"),val(c,"other_name")));}',
 '''itemCard("✉  "+val(c,"other_name"),val(c,"last_message"),()->openPrivate(c.optLong("id"),val(c,"other_name")));
            final long cid=c.optLong("id");
            if(cid>0)addButton(content,"Excluir conversa com "+val(c,"other_name")+" só para mim",MINT,NAVY,()->confirmClearConversation(cid,()->loadPage(false)));
        }''', 'excluir conversa privada para usuário')
swap('private void renderPrivateMessages(JSONObject j){JSONArray a=array(j,"items");for(int i=0;i<a.length();i++){JSONObject m=row(a,i);itemCard(val(m,"sender_name")+" · "+itemDate(val(m,"sent_at")),val(m,"body"),null);}if(a.length()==0)placeholder("Conversa vazia.");',
 '''private void renderPrivateMessages(JSONObject j){
        if(privateCid>0){addButton(content,"🗑  Limpar conversa só para mim",MINT,NAVY,()->confirmClearConversation(privateCid,()->{
            offline.save(userId,householdId,"dm_"+privateCid,"{\\\"items\\\":[]}");
            openPrivate(privateCid,privateTitle);
        }));gap(content,8);}
        JSONArray a=array(j,"items");for(int i=0;i<a.length();i++){
            JSONObject m=row(a,i);
            itemCard(val(m,"sender_name")+" · "+itemDate(val(m,"sent_at")),val(m,"body"),null);
            if(val(m,"sender_id").equals(userId)&&m.optLong("id")>0){
                final long mid=m.optLong("id");
                addButton(content,"Excluir minha mensagem",MINT,NAVY,()->confirmDeleteMessage(mid,privateCid,()->openPrivate(privateCid,privateTitle)));
            }
        }
        if(a.length()==0)placeholder("Conversa vazia.");''', 'limpar conversa privada e excluir autoria')
swap('    private void openPrivate(long cid,String person){',
 '''    private void confirmDeleteMessage(long messageId,long conversationId,Runnable refresh){
        new AlertDialog.Builder(this).setTitle("Excluir mensagem?")
          .setMessage("Esta mensagem enviada por você deixará de aparecer para todos os participantes.")
          .setNegativeButton("Cancelar",null).setPositiveButton("Excluir",(dialog,which)->{
            JSONObject p=withHid(data("id",String.valueOf(messageId),"conversation_id",String.valueOf(conversationId)));
            api.post("message_delete",p,(result,error)->{
                if(error!=null)msg("Não foi possível excluir: "+error);
                else{msg("Mensagem excluída.");refresh.run();}
            });
          }).show();
    }
    private void confirmClearConversation(long conversationId,Runnable refresh){
        new AlertDialog.Builder(this).setTitle("Limpar conversa?")
          .setMessage("O histórico anterior será ocultado somente para você. As outras pessoas continuam com as mensagens.")
          .setNegativeButton("Cancelar",null).setPositiveButton("Limpar",(dialog,which)->{
            JSONObject p=withHid(data("conversation_id",String.valueOf(conversationId)));
            api.post("conversation_clear",p,(result,error)->{
                if(error!=null)msg("Não foi possível limpar: "+error);
                else{msg("Conversa limpa na sua conta.");refresh.run();}
            });
          }).show();
    }
    private void openPrivate(long cid,String person){''', 'diálogos com confirmação e API segura')
main.write_text(src,encoding='utf-8')
g=grad.read_text(encoding='utf-8')
for a,b in [
 ("applicationId 'com.nossocombinado.app.native29'","applicationId 'com.nossocombinado.app.native30'"),
 ("versionCode 290","versionCode 300"),
 ("versionName '2.9.0'","versionName '3.0.0'")]:
 if g.count(a)!=1:raise RuntimeError('gradle missing '+a)
 g=g.replace(a,b,1)
grad.write_text(g,encoding='utf-8')
print('OK Nosso Combinado 3.0 nativo com opções de exclusão')
