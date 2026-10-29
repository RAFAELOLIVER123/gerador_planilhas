# Nosso Combinado Android — Edição Web 2.5 de teste

Este projeto é uma edição complementar baseada em WebView que acessa o site oficial em HTTPS, não uma recompilação do APK nativo 2.1. Ela mantém a experiência web completa com mapa, rotas via Google Maps e chat da versão publicada na Hostinger.

**Instalação lado a lado:** identificador com.nossocombinado.app.web25, diferente de com.nossocombinado.app.teste21. O APK não substitui o original, nem apaga seus dados.

- Mapa satélite/híbrido, pesquisa e Google Maps: recursos web publicados na Hostinger.
- Chat responsivo com polling da web 2.5: funciona com internet enquanto aberto. Não é websocket nem push com app fechado.
- Fotos e documentos: Android abre seletor de arquivos.
- Geolocalização: solicita autorização apenas após ação do usuário.
- Sons: reprodução web condicionada a volume/permissões/regras do Android. Alarmes nativos em segundo plano da 2.1 **não foram migrados**.
- Não inclui chave privada de assinatura do APK original. Este APK de teste é assinado pelo ambiente GitHub Actions e não deve substituir a versão nativa 2.1.
- Não inclui credenciais, senhas, nem dados dos usuários; login e conteúdo são servidos pelo domínio oficial.

Build: GitHub Actions no branch nosso-combinado-web25, com JDK 17, Android SDK 35, Gradle 8.9 e Android Gradle Plugin 8.7.3.
