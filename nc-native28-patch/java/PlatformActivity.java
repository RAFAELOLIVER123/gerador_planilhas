package com.nossocombinado.app;

import android.Manifest;
import android.webkit.JavascriptInterface;
import android.app.Activity;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.graphics.Color;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Message;
import android.view.View;
import android.webkit.CookieManager;
import android.webkit.GeolocationPermissions;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.webkit.WebSettings;
import android.widget.Toast;

public final class PlatformActivity extends Activity {
    private static final String SITE = "https://violet-grasshopper-887660.hostingersite.com/";
    private static final String SITE_HOST = "violet-grasshopper-887660.hostingersite.com";
    private static final int CHOOSE_FILE = 2025;
    private static final int LOCATION_PERMISSION = 2026;
    private WebView webView;
    private ValueCallback<Uri[]> filesCallback;
    private GeolocationPermissions.Callback locationCallback;
    private String pendingLocationOrigin;

    @Override public void onCreate(Bundle state) {
        super.onCreate(state);
        getWindow().setStatusBarColor(Color.rgb(255,249,246));
        getWindow().setNavigationBarColor(Color.rgb(255,249,246));
        getWindow().getDecorView().setSystemUiVisibility(View.SYSTEM_UI_FLAG_LIGHT_STATUS_BAR | View.SYSTEM_UI_FLAG_LIGHT_NAVIGATION_BAR);
        webView = new WebView(this);
        webView.setBackgroundColor(Color.rgb(255,249,246));
        webView.setScrollBarStyle(WebView.SCROLLBARS_INSIDE_OVERLAY);
        setContentView(webView);
        WebSettings settings = webView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setDatabaseEnabled(true);
        settings.setGeolocationEnabled(true);
        settings.setUseWideViewPort(true);
        settings.setLoadWithOverviewMode(true);
        settings.setMediaPlaybackRequiresUserGesture(false);
        settings.setAllowFileAccess(false);
        settings.setAllowContentAccess(true);
        settings.setMixedContentMode(WebSettings.MIXED_CONTENT_NEVER_ALLOW);
        settings.setSupportMultipleWindows(true);
        settings.setJavaScriptCanOpenWindowsAutomatically(true);
        CookieManager.getInstance().setAcceptCookie(true);
        CookieManager.getInstance().setAcceptThirdPartyCookies(webView,false);
        webView.addJavascriptInterface(new ReminderBridge(this), "NCAndroidReminders");
        if (Build.VERSION.SDK_INT >= 33 && checkSelfPermission("android.permission.POST_NOTIFICATIONS") != PackageManager.PERMISSION_GRANTED) requestPermissions(new String[]{"android.permission.POST_NOTIFICATIONS"}, 3026);
        webView.setWebViewClient(new WebViewClient() {
            @Override public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
                return request.isForMainFrame() && redirectIfExternal(request.getUrl());
            }
            @Override public boolean shouldOverrideUrlLoading(WebView view, String url) {
                return redirectIfExternal(Uri.parse(url));
            }
        });
        webView.setWebChromeClient(new WebChromeClient() {
            @Override public boolean onShowFileChooser(WebView view, ValueCallback<Uri[]> callback, FileChooserParams params) {
                if (filesCallback != null) filesCallback.onReceiveValue(null);
                filesCallback = callback;
                try {
                    Intent picker = params.createIntent();
                    picker.addCategory(Intent.CATEGORY_OPENABLE);
                    startActivityForResult(picker, CHOOSE_FILE);
                    return true;
                } catch (Exception error) {
                    filesCallback = null;
                    callback.onReceiveValue(null);
                    Toast.makeText(PlatformActivity.this, "Não foi possível abrir os arquivos.", Toast.LENGTH_SHORT).show();
                    return false;
                }
            }
            @Override public void onGeolocationPermissionsShowPrompt(String origin, GeolocationPermissions.Callback callback) {
                if (Build.VERSION.SDK_INT >= 23 &&
                    checkSelfPermission(Manifest.permission.ACCESS_FINE_LOCATION) != PackageManager.PERMISSION_GRANTED &&
                    checkSelfPermission(Manifest.permission.ACCESS_COARSE_LOCATION) != PackageManager.PERMISSION_GRANTED) {
                    pendingLocationOrigin = origin;
                    locationCallback = callback;
                    requestPermissions(new String[]{Manifest.permission.ACCESS_FINE_LOCATION,Manifest.permission.ACCESS_COARSE_LOCATION},LOCATION_PERMISSION);
                } else callback.invoke(origin,true,false);
            }
            @Override public boolean onCreateWindow(WebView view, boolean isDialog, boolean userGesture, Message resultMsg) {
                if (!userGesture) return false;
                final WebView popup = new WebView(PlatformActivity.this);
                popup.setWebViewClient(new WebViewClient() {
                    @Override public boolean shouldOverrideUrlLoading(WebView v, WebResourceRequest request) {
                        openOutside(request.getUrl());
                        return true;
                    }
                    @Override public boolean shouldOverrideUrlLoading(WebView v, String url) {
                        openOutside(Uri.parse(url));
                        return true;
                    }
                });
                WebView.WebViewTransport transport = (WebView.WebViewTransport) resultMsg.obj;
                transport.setWebView(popup);
                resultMsg.sendToTarget();
                return true;
            }
        });
        if (state != null) webView.restoreState(state);
        else webView.loadUrl(SITE);
    }

    private boolean redirectIfExternal(Uri url) {
        if (url == null) return true;
        String scheme = url.getScheme() == null ? "" : url.getScheme().toLowerCase(java.util.Locale.ROOT);
        String host = url.getHost() == null ? "" : url.getHost().toLowerCase(java.util.Locale.ROOT);
        if (("https".equals(scheme) || "http".equals(scheme)) && SITE_HOST.equals(host)) return false;
        openOutside(url);
        return true;
    }

    private void openOutside(Uri uri) {
        try {
            if (uri == null) return;
            Intent intent = new Intent(Intent.ACTION_VIEW,uri);
            intent.addCategory(Intent.CATEGORY_BROWSABLE);
            startActivity(intent);
        } catch (Exception error) {
            Toast.makeText(this,"Nenhum aplicativo disponível para abrir o link.",Toast.LENGTH_SHORT).show();
        }
    }

    @Override protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode,resultCode,data);
        if (requestCode == CHOOSE_FILE && filesCallback != null) {
            Uri[] result = WebChromeClient.FileChooserParams.parseResult(resultCode,data);
            filesCallback.onReceiveValue(result);
            filesCallback = null;
        }
    }
    @Override public void onRequestPermissionsResult(int requestCode,String[] permissions,int[] results) {
        super.onRequestPermissionsResult(requestCode,permissions,results);
        if(requestCode == LOCATION_PERMISSION && locationCallback != null) {
            boolean granted = false;
            for(int result : results) if(result == PackageManager.PERMISSION_GRANTED) granted = true;
            locationCallback.invoke(pendingLocationOrigin, granted, false);
            locationCallback = null;
            pendingLocationOrigin = null;
        }
    }
    @Override public void onBackPressed() {
        if(webView != null && webView.canGoBack())webView.goBack();
        else super.onBackPressed();
    }
    @Override protected void onPause() {
        if(webView != null) webView.onPause();
        super.onPause();
    }
    @Override protected void onResume() {
        super.onResume();
        if(webView != null)webView.onResume();
    }
    @Override protected void onSaveInstanceState(Bundle out) {
        if(webView != null)webView.saveState(out);
        super.onSaveInstanceState(out);
    }
    @Override protected void onDestroy() {
        if(filesCallback != null){filesCallback.onReceiveValue(null);filesCallback=null;}
        if(webView != null){webView.destroy();webView=null;}
        super.onDestroy();
    }
}
