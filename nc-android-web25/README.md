# Nosso Combinado — Android 2.6 (edição web + alarmes nativos)

APK instalável: `com.nossocombinado.app.web26` / versionCode 260 / versionName 2.6.0.
É um aplicativo de teste independente do APK nativo 2.1 `com.nossocombinado.app.teste21`. **Não substitui, modifica nem preserva o armazenamento privado do app 2.1.**

## Requisitos
- Publicar os arquivos WEB 2.6 na Hostinger, no domínio configurado.
- Abrir o aplicativo com internet e entrar na conta. As tarefas/compromissos existentes são sincronizados enquanto a página estiver ativa.
- Permitir notificações (Android 13+) e, para horário mais exato em Android 12+, autorizar alarmes exatos no sistema quando oferecido.
- Manter volume de alarme audível e verificar configurações de Não Perturbe.

## Funcionalidades
- Site oficial em HTTPS, sem armazenar senhas no código.
- Mapa interativo da web, fotos/endereço e rotas via Google Maps.
- Chat web responsivo 2.6, sem saltos no editor, e consulta periódica enquanto aberto.
- Seleção de arquivos e fotos via Android, localização opcional com permissão.
- Som web nas interações. AlarmManager agenda alarmes locais a partir da agenda sincronizada; BroadcastReceiver inicia serviço Android em primeiro plano que toca o **toque de alarme do sistema** em loop, com aviso e ação DESLIGAR. Autodesliga após 120 segundos.
- Restauração dos alarmes previamente agendados após reinício do dispositivo, quando o Android permitir.

## Limitações
- Não inclui a interface e o armazenamento offline nativos completos do APK 2.1.
- Não recebe mensagens instantaneamente com app fechado (sem FCM ou WebSocket); polling na página aberta.
- Alarmes recém-criados em OUTRO celular enquanto este aplicativo está fechado não podem ser sincronizados sem reabrir o app.
- Permissões do sistema, Doze e fabricantes podem atrasar alarmes inexatos ou impedir execução; não há garantia de tocar em modo silencioso/Não Perturbe.
- APK de teste assinado com chave debug GitHub Actions; não tem a assinatura original 2.1. Não distribuir na Play Store como release.
- O teste realizado foi compilação Gradle, assinatura e análise ZIP/manifesto; **não houve execução em aparelho físico**.

## Build
Branch `nosso-combinado-web26`, diretório `nc-android-web25/` (nome histórico), workflow `.github/workflows/nc-web26.yml`.
JDK 17 / SDK 35 / Gradle 8.9 / plugin 8.7.3.

## Segurança
Use somente sua conta legítima no servidor. Faça backup antes de atualizar arquivos web. Não desinstale a versão nativa 2.1 para instalar esta edição de testes.
