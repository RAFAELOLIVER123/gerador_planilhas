package com.nossocombinado.app;
import android.app.Activity;
import android.content.Intent;
import android.graphics.Color;
import android.os.Bundle;
import android.view.Gravity;
import android.view.View;
import android.widget.Button;
import android.widget.LinearLayout;
import android.widget.TextView;
public final class AlarmActivity extends Activity{
  @Override public void onCreate(Bundle state){
    super.onCreate(state);
    LinearLayout root=new LinearLayout(this);root.setOrientation(LinearLayout.VERTICAL);
    root.setGravity(Gravity.CENTER);root.setBackgroundColor(Color.rgb(255,246,244));root.setPadding(34,30,34,30);
    TextView title=new TextView(this);title.setText("⏰ Nosso Combinado");title.setTextSize(26);title.setTextColor(Color.rgb(40,49,65));title.setGravity(Gravity.CENTER);
    TextView detail=new TextView(this);detail.setText(getIntent().getStringExtra("title"));detail.setTextColor(Color.rgb(60,60,75));detail.setTextSize(18);detail.setGravity(Gravity.CENTER);
    Button off=new Button(this);off.setText("DESLIGAR ALARME");off.setAllCaps(false);off.setTextSize(18);
    off.setOnClickListener(v->{stopService(new Intent(this,AlarmService.class));finish();});
    root.addView(title);root.addView(detail);root.addView(off);
    setContentView(root);
  }
}
