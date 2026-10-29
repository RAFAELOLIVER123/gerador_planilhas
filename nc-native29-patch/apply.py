#!/usr/bin/env python3
"""Nosso Combinado 2.9: corrigir autenticação/sessão compartilhada e sincronização."""
from pathlib import Path
import re,sys
root=Path(sys.argv[1]).resolve()
pkg=root/'app/src/main/java/com/nossocombinado/app'
main=pkg/'MainActivity.java'
api=pkg/'ApiClient.java'
platform=pkg/'PlatformActivity.java'
grad=root/'app/build.gradle'

def replace_once(path, old, new, label):
    text=path.read_text(encoding="utf-8")
    count=text.count(old)
    if count!=1:
        raise RuntimeError(f"{label}: expected once, found {count}")
    path.write_text(text.replace(old,new,1),encoding="utf-8")
    print("OK",label)

replace_once(grad,"applicationId 'com.nossocombinado.app.native28'",
             "applicationId 'com.nossocombinado.app.native29'","pacote separado para não apagar app 2.1")
replace_once(grad,"versionCode 280","versionCode 290","versionCode")
replace_once(grad,"versionName '2.8.0'","versionName '2.9.0'","versionName")

replace_once(main,
  'if(!address.contains("@")||secret.length()<10){msg("Informe um e-mail e senha com pelo menos 10 caracteres.");return;}',
  'if(!address.contains("@")||secret.isEmpty()||(register&&secret.length()<10)){msg(register?"Informe um e-mail válido e senha com pelo menos 10 caracteres.":"Informe seu e-mail e a senha cadastrada.");return;}',
  "login aceita senha legada; cadastro continua exigindo dez")

replace_once(main,
  'if(e.toLowerCase().contains("sessão expirada")||e.toLowerCase().contains("entre na sua conta"))msg("Sessão online expirada. Você pode consultar seu cache offline; entre novamente para sincronizar.");',
  'if(e.toLowerCase(java.util.Locale.ROOT).contains("sessão expirada")||e.toLowerCase(java.util.Locale.ROOT).contains("entre na sua conta")){msg("Sessão expirada. Entre novamente para reativar a sincronização; os dados locais não serão apagados.");authScreen(false);}',
  "reconexão ao expirar sessão sem apagar cache")

replace_once(main,
  'handler.postDelayed(this,3600);',
  'handler.postDelayed(this,(page.equals("chat")||page.equals("messages")||page.equals("private_thread"))?1850:3000);',
  "novas mensagens com atualização mais frequente")
replace_once(main,
  'handler.postDelayed(realtimeInbox,3600);',
  'handler.postDelayed(realtimeInbox,1850);',
  "primeira leitura de mensagens")

replace_once(main,
  'content.removeAllViews();sectionTitle("Chat com "+privateTitle,"Conversa privada entre integrantes.");addButton(content,"← Voltar",MINT,NAVY,()->navigate("messages"));gap(content,12);renderPrivateMessages(j);',
  'final int oldChatY=scroll==null?0:scroll.getScrollY();content.removeAllViews();sectionTitle("Chat com "+privateTitle,"Conversa privada entre integrantes.");addButton(content,"← Voltar",MINT,NAVY,()->navigate("messages"));gap(content,12);renderPrivateMessages(j);if(oldChatY>0&&scroll!=null)scroll.post(()->scroll.scrollTo(0,oldChatY));',
  "conversa privada mantém posição de leitura")

replace_once(api,
  '''String set=c.getHeaderField("Set-Cookie");if(set!=null&&set.startsWith("nc_session=")){
            String v=set.substring("nc_session=".length()).split(";",2)[0];if(!v.isEmpty())store.set("session_cookie",v);
        }''',
  '''// Não depender da posição do cabeçalho: proxies podem incluir outras cookies.
        for(java.util.Map.Entry<String,java.util.List<String>> entry:c.getHeaderFields().entrySet()){
            if(entry.getKey()==null||!"set-cookie".equalsIgnoreCase(entry.getKey())||entry.getValue()==null)continue;
            for(String line:entry.getValue()){
                if(line==null)continue;
                for(String field:line.split(";")){
                    String trimmed=field.trim();
                    if(trimmed.startsWith("nc_session=")){
                        String token=trimmed.substring("nc_session=".length()).trim();
                        if(token.matches("[a-fA-F0-9]{64}"))store.set("session_cookie",token);
                    }
                }
            }
        }''',
  "captura robusta do cookie nc_session")

replace_once(platform,
  '''if (state != null) webView.restoreState(state);
        else webView.loadUrl(SITE);''',
  '''// A interface nativa e o site interno compartilham a mesma sessão HTTPS.
        // Não expor token à camada JavaScript nem ao histórico de navegação.
        String token=new SecureStore(this).get("session_cookie");
        final Runnable start=()->{
            if(isFinishing())return;
            if(state!=null)webView.restoreState(state);
            else webView.loadUrl(SITE);
        };
        if(token.matches("[a-fA-F0-9]{64}")){
            CookieManager.getInstance().setCookie(SITE,
              "nc_session="+token+"; Path=/; Secure; HttpOnly; SameSite=Lax",
              ignored->{CookieManager.getInstance().flush();start.run();});
        }else start.run();''',
  "WebView utiliza sessão segura do aplicativo nativo")

print("Nosso Combinado Android 2.9: autenticação, sessão e mensagens prontas.")
