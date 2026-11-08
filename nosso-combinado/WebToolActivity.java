package com.nossocombinado.app;

import android.Manifest;
import android.app.Activity;
import android.app.AlertDialog;
import android.app.DownloadManager;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.graphics.Color;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.view.Gravity;
import android.view.View;
import android.view.ViewGroup;
import android.webkit.CookieManager;
import android.webkit.GeolocationPermissions;
import android.webkit.URLUtil;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceError;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.webkit.ValueCallback;
import android.widget.Button;
import android.widget.FrameLayout;
import android.widget.HorizontalScrollView;
import android.widget.LinearLayout;
import android.widget.TextView;
import android.widget.Toast;
import java.net.URI;

/** Telas complementares integradas ao app nativo, com sessão do site oficial. */
public final class WebToolActivity extends Activity {
    private static final String HOME = "https://violet-grasshopper-887660.hostingersite.com/";
    private static final String HOST = "violet-grasshopper-887660.hostingersite.com";
    private static final int FILE_REQUEST = 773;
    private static final int GEO_REQUEST = 774;
    private FrameLayout body;
    private WebView web;
    private ValueCallback<Uri[]> fileCallback;
    private GeolocationPermissions.Callback geoCallback;
    private String geoOrigin;
    private String requestedPage;
    private String lastInjectedPage = "";
    private LinearLayout errorOverlay;

    private int dp(float n) {return Math.round(n*getResources().getDisplayMetrics().density);}
    private Button chip(String text) {
        Button button = new Button(this);
        button.setAllCaps(false);
        button.setText(text);
        button.setTextSize(12);
        button.setTextColor(Color.rgb(32,46,69));
        button.setBackgroundTintList(android.content.res.ColorStateList.valueOf(Color.rgb(238,242,252)));
        button.setPadding(dp(12),0,dp(12),0);
        return button;
    }
    @Override public void onCreate(Bundle state) {
        super.onCreate(state);
        requestedPage = getIntent().getStringExtra("nc_page");
        if (!allowed(requestedPage)) requestedPage = "price-compare";

        LinearLayout root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);
        root.setBackgroundColor(Color.WHITE);

        LinearLayout top = new LinearLayout(this);
        top.setGravity(Gravity.CENTER_VERTICAL);
        top.setBackgroundColor(Color.rgb(248,249,253));
        top.setPadding(dp(8),dp(3),dp(8),dp(3));
        Button back = chip("‹  Voltar ao app");
        back.setOnClickListener(v->finish());
        top.addView(back,new LinearLayout.LayoutParams(-2,dp(47)));
        TextView title = new TextView(this);
        title.setText("Nosso Combinado");
        title.setTextColor(Color.rgb(24,38,60));
        title.setTypeface(null,1);
        title.setTextSize(17);
        title.setGravity(Gravity.CENTER_VERTICAL);
        top.addView(title,new LinearLayout.LayoutParams(0,dp(47),1));
        Button reload = chip("↻");
        reload.setContentDescription("Atualizar tela");
        reload.setOnClickListener(v->web.reload());
        top.addView(reload,new LinearLayout.LayoutParams(dp(54),dp(47)));
        root.addView(top,new LinearLayout.LayoutParams(-1,dp(53)));

        HorizontalScrollView scroll = new HorizontalScrollView(this);
        scroll.setHorizontalScrollBarEnabled(false);
        scroll.setBackgroundColor(Color.rgb(247,248,252));
        LinearLayout nav = new LinearLayout(this);
        nav.setPadding(dp(6),dp(3),dp(6),dp(3));
        String[] routes={"price-compare","finance","finance-matrix","shopping","pharmacy","school","vehicle"};
        String[] labels={"📊 Comparador","R$ Financeiro","▦ Planejamento","🛒 Compras","💊 Farmácia","🏫 Escola","🚗 Veículos"};
        for(int i=0;i<routes.length;i++){
            final String page=routes[i];
            Button button=chip(labels[i]);
            LinearLayout.LayoutParams lp=new LinearLayout.LayoutParams(-2,dp(41));
            lp.rightMargin=dp(5);
            nav.addView(button,lp);
            button.setOnClickListener(v->{ requestedPage=page; lastInjectedPage=""; openRequestedPage(); });
        }
        scroll.addView(nav,new HorizontalScrollView.LayoutParams(-2,dp(47)));
        root.addView(scroll,new LinearLayout.LayoutParams(-1,dp(47)));

        body=new FrameLayout(this);
        web=new WebView(this);
        web.setBackgroundColor(Color.WHITE);
        WebSettings settings=web.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setDatabaseEnabled(true);
        settings.setLoadsImagesAutomatically(true);
        settings.setSupportZoom(true);
        settings.setBuiltInZoomControls(true);
        settings.setDisplayZoomControls(false);
        settings.setMediaPlaybackRequiresUserGesture(false);
        settings.setAllowFileAccess(false);
        settings.setJavaScriptCanOpenWindowsAutomatically(false);
        settings.setMixedContentMode(WebSettings.MIXED_CONTENT_NEVER_ALLOW);
        CookieManager.getInstance().setAcceptCookie(true);
        CookieManager.getInstance().setAcceptThirdPartyCookies(web,false);

        web.setWebViewClient(new WebViewClient() {
            @Override public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
                Uri uri=request.getUrl();
                if(official(uri.toString()))return false;
                String scheme=uri.getScheme();
                if("about".equals(scheme)||"blob".equals(scheme))return false;
                try{startActivity(new Intent(Intent.ACTION_VIEW,uri));}
                catch(Exception e){Toast.makeText(WebToolActivity.this,"Não foi possível abrir o link.",Toast.LENGTH_SHORT).show();}
                return true;
            }
            @Override public void onPageFinished(WebView view, String url) {
                errorOverlay.setVisibility(View.GONE);
                openRequestedPage();
            }
            @Override public void onReceivedError(WebView view, WebResourceRequest request, WebResourceError error) {
                if(request.isForMainFrame())errorOverlay.setVisibility(View.VISIBLE);
            }
        });
        web.setWebChromeClient(new WebChromeClient(){
            @Override public boolean onShowFileChooser(WebView view,ValueCallback<Uri[]> callback,FileChooserParams params){
                if(fileCallback!=null)fileCallback.onReceiveValue(null);
                fileCallback=callback;
                try{startActivityForResult(params.createIntent(),FILE_REQUEST);}
                catch(Exception e){
                    fileCallback.onReceiveValue(null);
                    fileCallback=null;
                    Toast.makeText(WebToolActivity.this,"Não foi possível abrir arquivos.",Toast.LENGTH_SHORT).show();
                }
                return true;
            }
            @Override public void onGeolocationPermissionsShowPrompt(String origin,GeolocationPermissions.Callback callback){
                if(!official(origin)){callback.invoke(origin,false,false);return;}
                if(Build.VERSION.SDK_INT<23 ||
                  checkSelfPermission(Manifest.permission.ACCESS_FINE_LOCATION)==PackageManager.PERMISSION_GRANTED ||
                  checkSelfPermission(Manifest.permission.ACCESS_COARSE_LOCATION)==PackageManager.PERMISSION_GRANTED) {
                    callback.invoke(origin,true,false);return;
                }
                geoOrigin=origin;geoCallback=callback;
                new AlertDialog.Builder(WebToolActivity.this)
                  .setTitle("Localização")
                  .setMessage("Permitir acesso à sua localização no mapa da família?")
                  .setNegativeButton("Não",(dialog,which)->denyGeo())
                  .setPositiveButton("Permitir",(dialog,which)->requestPermissions(
                       new String[]{Manifest.permission.ACCESS_FINE_LOCATION,Manifest.permission.ACCESS_COARSE_LOCATION},GEO_REQUEST))
                  .setCancelable(false).show();
            }
        });
        web.setDownloadListener((url,ua,disposition,mime,size)->{
            if(!official(url)){
                try{startActivity(new Intent(Intent.ACTION_VIEW,Uri.parse(url)));}
                catch(Exception ignored){}
                return;
            }
            try {
                DownloadManager.Request request=new DownloadManager.Request(Uri.parse(url));
                String cookie=CookieManager.getInstance().getCookie(url);
                if(cookie!=null)request.addRequestHeader("Cookie",cookie);
                request.addRequestHeader("User-Agent",ua);
                String name=URLUtil.guessFileName(url,disposition,mime);
                request.setTitle(name);
                request.setMimeType(mime);
                request.setNotificationVisibility(DownloadManager.Request.VISIBILITY_VISIBLE_NOTIFY_COMPLETED);
                request.setDestinationInExternalPublicDir(Environment.DIRECTORY_DOWNLOADS,name);
                ((DownloadManager)getSystemService(DOWNLOAD_SERVICE)).enqueue(request);
                Toast.makeText(WebToolActivity.this,"Download iniciado",Toast.LENGTH_SHORT).show();
            } catch(Exception e){Toast.makeText(WebToolActivity.this,"Não foi possível baixar o arquivo.",Toast.LENGTH_SHORT).show();}
        });
        body.addView(web,new FrameLayout.LayoutParams(-1,-1));
        errorOverlay=new LinearLayout(this);
        errorOverlay.setOrientation(LinearLayout.VERTICAL);
        errorOverlay.setGravity(Gravity.CENTER);
        errorOverlay.setBackgroundColor(Color.WHITE);
        TextView message=new TextView(this);
        message.setText("Não foi possível abrir o site.\nVerifique a internet e tente novamente.");
        message.setGravity(Gravity.CENTER);
        message.setTextSize(15);
        errorOverlay.addView(message);
        Button retry=chip("Tentar novamente");
        retry.setOnClickListener(v->{errorOverlay.setVisibility(View.GONE);web.loadUrl(HOME);});
        errorOverlay.addView(retry);
        errorOverlay.setVisibility(View.GONE);
        body.addView(errorOverlay,new FrameLayout.LayoutParams(-1,-1));
        root.addView(body,new LinearLayout.LayoutParams(-1,0,1));
        setContentView(root);
        if(state==null||web.restoreState(state)==null)web.loadUrl(HOME);
    }
    private static boolean allowed(String page){
        return "price-compare".equals(page)||"finance".equals(page)||"finance-matrix".equals(page)||
          "shopping".equals(page)||"pharmacy".equals(page)||"school".equals(page)||"vehicle".equals(page);
    }
    private static boolean official(String value){
        try {
            URI url=new URI(value);
            return "https".equalsIgnoreCase(url.getScheme()) && HOST.equalsIgnoreCase(url.getHost());
        }catch(Exception e){return false;}
    }
    private void openRequestedPage(){
        if(web==null || !official(web.getUrl()==null?"":web.getUrl())) return;
        final String route=requestedPage;
        // All routes are fixed strings from "allowed"; never interpolate arbitrary input into JavaScript.
        String script="(function(){"+
          "var id='"+route+"';"+
          "if(window.ncNativeRouteObserver){window.ncNativeRouteObserver.disconnect();}"+
          "function nav(){var b=document.querySelector('[data-nav=\"'+id+'\"]');"+
          "if(b){b.click();return true;}return false;}"+
          "if(!nav()){var ob=new MutationObserver(function(){if(nav())ob.disconnect();});"+
          "window.ncNativeRouteObserver=ob;ob.observe(document.documentElement,{childList:true,subtree:true});}"+
          "})()";
        web.evaluateJavascript(script,null);
    }
    private void denyGeo(){
        if(geoCallback!=null){geoCallback.invoke(geoOrigin,false,false);geoCallback=null;geoOrigin=null;}
    }
    @Override public void onRequestPermissionsResult(int request,String[] perm,int[] results){
        super.onRequestPermissionsResult(request,perm,results);
        if(request==GEO_REQUEST && geoCallback!=null){
            boolean granted=false;
            for(int result:results)if(result==PackageManager.PERMISSION_GRANTED)granted=true;
            geoCallback.invoke(geoOrigin,granted,false);geoCallback=null;geoOrigin=null;
        }
    }
    @Override protected void onActivityResult(int request,int result,Intent data){
        super.onActivityResult(request,result,data);
        if(request==FILE_REQUEST&&fileCallback!=null){
            fileCallback.onReceiveValue(WebChromeClient.FileChooserParams.parseResult(result,data));
            fileCallback=null;
        }
    }
    @Override protected void onSaveInstanceState(Bundle state){
        if(web!=null)web.saveState(state);
        super.onSaveInstanceState(state);
    }
    @Override public void onBackPressed(){
        if(web!=null&&web.canGoBack())web.goBack();
        else finish();
    }
    @Override protected void onDestroy(){
        if(web!=null) {
            if(fileCallback!=null){fileCallback.onReceiveValue(null);fileCallback=null;}
            body.removeView(web);web.destroy();web=null;
        }
        denyGeo();
        super.onDestroy();
    }
}
