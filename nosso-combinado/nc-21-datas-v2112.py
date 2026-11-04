#!/usr/bin/env python3
"""Padrão de datas Android 2.1.12: dd/MM/aaaa na apresentação e nos formulários.
Não altera datas ISO armazenadas, calendário interno, fuso ou APIs.
"""
from pathlib import Path
import sys,re
root=Path(sys.argv[1]);f=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
s=f.read_text(encoding='utf-8')
def once(before,after,label):
    global s
    n=s.count(before)
    if n!=1:raise RuntimeError(label+': esperado 1 trecho, encontrados '+str(n))
    s=s.replace(before,after,1)
    print('OK',label)
helpers=r'''
    /** Datas recebidas continuam ISO; somente os textos visíveis usam dd/MM/aaaa. */
    private String ncDateBR(String source){
        if(source==null)return "";
        return source.replaceAll("(?<!\\d)(\\d{4})-(\\d{2})-(\\d{2})(?!\\d)","$3/$2/$1");
    }
    private String ncDateHint(String hint){
        if(hint==null)return "";
        return hint.replace("AAAA-MM-DD","DD/MM/AAAA")
                   .replace("AAAA/MM/DD","DD/MM/AAAA")
                   .replace("YYYY-MM-DD","DD/MM/AAAA")
                   .replace("yyyy-MM-dd","dd/MM/aaaa");
    }
    private boolean ncDateField(String hint){
        if(hint==null)return false;
        String value=hint.toLowerCase(java.util.Locale.ROOT);
        return value.contains("data")||value.contains("vencimento")
            ||value.contains("nascimento")||value.contains("prazo")||value.contains("date");
    }
    private String ncDateInput(String value,String hint){
        return ncDateField(hint)?ncDateBR(value):value;
    }
    private String ncDateSave(String value,String hint){
        if(value==null||!ncDateField(hint))return value==null?"":value;
        java.util.regex.Matcher m=java.util.regex.Pattern
            .compile("^(\\d{2})/(\\d{2})/(\\d{4})(.*)$")
            .matcher(value.trim());
        if(!m.matches())return value.trim();
        try{
            String dayMonthYear=m.group(1)+"/"+m.group(2)+"/"+m.group(3);
            java.text.SimpleDateFormat parser=new java.text.SimpleDateFormat("dd/MM/yyyy",java.util.Locale.forLanguageTag("pt-BR"));
            parser.setLenient(false);
            java.util.Date date=parser.parse(dayMonthYear);
            String iso=new java.text.SimpleDateFormat("yyyy-MM-dd",java.util.Locale.US).format(date);
            return iso+m.group(4);
        }catch(Exception ignored){return value.trim();}
    }
'''
# Insert helpers immediately before the sole method that constructs text views.
marker='    private TextView label(String text,int size,int color,boolean bold)'
if s.count(marker)!=1:raise RuntimeError('Função label não localizada')
s=s.replace(marker,helpers+'\n'+marker,1)
once('v.setText(tr(text));','v.setText(ncDateBR(tr(text)));','exibir datas pré-existentes em todos os textos da UI')
once('text(form,hints[i],13,NAVY,true);gap(form,6);EditText e=edit(hints[i],false);e.setText(i<initial.length?initial[i]:"");',
     'text(form,ncDateHint(hints[i]),13,NAVY,true);gap(form,6);EditText e=edit(ncDateHint(hints[i]),false);e.setText(ncDateInput(i<initial.length?initial[i]:"",hints[i]));',
     'rótulos e preenchimento dos formulários')
once('vals[i]=fields[i].getText().toString().trim();',
     'vals[i]=ncDateSave(fields[i].getText().toString().trim(),hints[i]);',
     'salvar datas informadas dd/MM/aaaa no padrão da API')
# A seleção de datas nos formulários nativos de tarefa e compromisso utiliza ISO
# apenas em sua variável interna; o texto dos botões deve ser brasileiro.
pat=r'\.setText\("📅 "\+([a-zA-Z_][a-zA-Z_0-9]*\[0\])\)'
s,count=re.subn(pat,lambda m:'.setText("📅 "+ncDateBR('+m.group(1)+'))',s)
print('Botões de data pós-seleção ajustados:',count)
if count<1:raise RuntimeError('Nenhum botão de data nativo encontrado')
# Não alterar localDate(): usado como ISO na filtragem e no calendário.
# utc() continua convertendo datas para o servidor.
f.write_text(s,encoding='utf-8')
grad=root/'app/build.gradle'
g=grad.read_text(encoding='utf-8')
assert g.count('versionCode 221')==1
assert g.count("versionName '2.1.11'")==1
grad.write_text(g.replace('versionCode 221','versionCode 222').replace("versionName '2.1.11'","versionName '2.1.12'"),encoding='utf-8')
print('Android 2.1.12: datas exibidas dd/MM/aaaa, envio ISO preservado.')
