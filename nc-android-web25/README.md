# Nosso Combinado — Android 2.7 (edição híbrida de teste)

APK Android: **Nosso Combinado 2.7.0** — applicationId `com.nossocombinado.app.web27`, versionCode `270`.

## Novidades
- Exibe o site oficial https://violet-grasshopper-887660.hostingersite.com/.
- **O novo ícone de mensagens** (💬) e o contador de não lidas aparecem **após publicar a correção WEB 2.7 na Hostinger**. A interface web é atualizada online, não incluída dentro do APK.
- Ícone mostra avisos de chat familiar e privado; ao tocar em uma mensagem, abre o chat correspondente. Não reconstrói o editor.
- Consulta os avisos a aproximadamente 2,5 a 3 segundos com o app aberto, internet ativa e sessão válida; não existe push remoto com o app fechado.
- Mantidos: upload de fotos, mapa e rotas do Google Maps pelo site, permissão de localização, agendamento local Android de tarefas e compromissos sincronizados, toque de despertador do sistema e botão DESLIGAR.

## Compatibilidade
O aplicativo **não substitui o APK nativo 2.1** nem a edição de teste 2.6: usa um packageName exclusivo e instala lado a lado. Faça login usando os mesmos dados de acesso ao servidor após instalar.

Este APK de teste usa assinatura debug do GitHub Actions, sem chave de assinatura original de versões anteriores. Os alarmes exigem as permissões do Android; o modo silencioso/Não Perturbe pode impedi-los ou atrasá-los.

## Requisitos / validação
Java 17, Gradle 8.9, SDK 35, AGP 8.7.3. Workflow `.github/workflows/nc-web27.yml`. APK é verificado por `unzip -t`, `apksigner verify` e `aapt dump badging` no CI. Não foi testado em aparelho físico.

## Instalação web antes do APK
Publicar `NossoCombinado_WEB_ICONE_MENSAGENS_2_7.zip` em `public_html` preservando os arquivos privados e o banco. Confirmar pelo arquivo `nc-inbox-versao-27.txt`.
