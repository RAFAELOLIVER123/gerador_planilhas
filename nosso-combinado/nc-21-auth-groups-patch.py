#!/usr/bin/env python3
from pathlib import Path
import sys
root=Path(sys.argv[1]);p=root/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
s=p.read_text()
def change(a,b,label):
 global s
 assert s.count(a)==1,(label,s.count(a))
 s=s.replace(a,b,1)
 print("OK",label)
change('if(register){form.addView(name,new LinearLayout.LayoutParams(-1,dp(52)));gap(form,11);}email.setInputType(33);form.addView(email,new LinearLayout.LayoutParams(-1,dp(52)));gap(form,11);form.addView(password,new LinearLayout.LayoutParams(-1,dp(52)));gap(form,19);',
'''if(register){form.addView(name,new LinearLayout.LayoutParams(-1,dp(52)));gap(form,11);}email.setInputType(33);form.addView(email,new LinearLayout.LayoutParams(-1,dp(52)));gap(form,11);form.addView(password,new LinearLayout.LayoutParams(-1,dp(52)));gap(form,6);
        EditText confirmPassword=register?edit("Confirme sua senha",true):null;
        if(register){form.addView(confirmPassword,new LinearLayout.LayoutParams(-1,dp(52)));gap(form,6);}
        CheckBox showPassword=new CheckBox(this);showPassword.setText("Mostrar senhas");showPassword.setTextColor(NAVY);
        showPassword.setOnCheckedChangeListener((button,show)->{
            int cursor=password.getSelectionStart();
            password.setTransformationMethod(show?null:android.text.method.PasswordTransformationMethod.getInstance());
            password.setSelection(Math.min(Math.max(cursor,0),password.length()));
            if(confirmPassword!=null){int confirmCursor=confirmPassword.getSelectionStart();confirmPassword.setTransformationMethod(show?null:android.text.method.PasswordTransformationMethod.getInstance());confirmPassword.setSelection(Math.min(Math.max(confirmCursor,0),confirmPassword.length()));}
        });form.addView(showPassword,new LinearLayout.LayoutParams(-1,dp(42)));gap(form,12);''','visibility and password repeat')
change('String address=email.getText().toString().trim(),secret=password.getText().toString();if(!address.contains("@")||secret.length()<10){msg("Informe um e-mail e senha com pelo menos 10 caracteres.");return;}',
'''String address=email.getText().toString().trim(),secret=password.getText().toString();if(!address.contains("@")||secret.length()<10){msg("Informe um e-mail e senha com pelo menos 10 caracteres.");return;}
            if(register&&(confirmPassword==null||!secret.equals(confirmPassword.getText().toString()))){msg("As senhas não conferem.");return;}''','registration confirmation validation')
change('JSONObject b=data("email",address,"password",secret);if(register)try{b.put("name",name.getText().toString().trim());}catch(Exception ignored){}',
'JSONObject b=data("email",address,"password",secret);if(register)try{b.put("name",name.getText().toString().trim());b.put("password_confirmation",confirmPassword.getText().toString());}catch(Exception ignored){}','pass confirmed field to API')
change('    private void householdsScreen(){root();',
'''    private void createOwnFamily(){formDialog("Novo grupo familiar",new String[]{"Nome da família ou grupo"},new String[]{""},values->{
        if(values[0].trim().isEmpty()){msg("Informe o nome do grupo.");return;}
        api.post("household_create",data("name",values[0].trim()),(j,e)->{
            if(e!=null){msg(e);return;}
            householdId=val(j,"household_id");householdName=values[0].trim();role="owner";
            vault.set("hid",householdId);vault.set("household_name",householdName);vault.set("role",role);
            page="members";familyTab="invites";refreshMe();
            msg("Grupo criado! Você já pode gerar convites para este grupo.");
        });
    });}
    private void householdsScreen(){root();''','creator can create groups')
change('addButton(c,"Criar novo espaço",PINK,WHITE,()->formDialog("Novo espaço familiar",new String[]{"Nome da família"},new String[]{""},values->{api.post("household_create",data("name",values[0]),(j,e)->{if(e!=null){msg(e);return;}householdId=val(j,"household_id");householdName=values[0];vault.set("hid",householdId);vault.set("household_name",householdName);page="home";appScreen();});}));',
'addButton(c,"Criar novo grupo e convidar",PINK,WHITE,this::createOwnFamily);','new account creates group and invites')
change('sectionTitle("Nossa família ♡","Grupo: "+householdName);',
'sectionTitle("Nossa família ♡","Grupo: "+householdName);addButton(content,"+ Criar meu próprio grupo",MINT,NAVY,this::createOwnFamily);gap(content,10);','members can also create own group')
change('if(!"owner".equals(role)&&!"admin".equals(role)){placeholder("Somente administradores podem gerar convites.");return;}',
'if(!"owner".equals(role)){placeholder("Somente quem criou este grupo pode gerar convites. Você pode criar seu próprio grupo para convidar pessoas nele.");return;}','invite only group owner')
change('private void loadInvites(){if(!"owner".equals(role)&&!"admin".equals(role))return;',
'private void loadInvites(){if(!"owner".equals(role))return;','hide invite list for invited admin')
change('    private void inviteHelp(){dialogBuilder().setTitle("Convites da família").setMessage("Os códigos são privados, expiram e só podem ser utilizados uma vez. Use Gerar convite para criar um novo código para compartilhar com a pessoa desejada.")',
'    private void inviteHelp(){dialogBuilder().setTitle("Convites da família").setMessage("Somente quem criou este grupo pode gerar convites. Códigos são privados, expiram e só podem ser utilizados uma vez.")','invite ownership explanation')
p.write_text(s)
g=root/'app/build.gradle';x=g.read_text()
for a,b in [("versionCode 212","versionCode 213"),("versionName '2.1.2'","versionName '2.1.3'")]:
 assert x.count(a)==1,a
 x=x.replace(a,b,1)
g.write_text(x)
print("Native 2.1.3 password confirmation and creator-only invitations applied.")
