package com.nossocombinado.app;

import android.app.Activity;
import android.content.ContentValues;
import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.net.Uri;
import android.os.Bundle;
import android.provider.MediaStore;
import android.widget.Toast;
import org.json.JSONArray;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.UUID;

/**
 * Seletor/câmera isolados do MainActivity nativo.
 * O comprovante é armazenado SOMENTE no diretório privado deste aplicativo.
 */
public final class NcFinancePhotoActivity extends Activity {
    private static final int PICK=5501, CAMERA=5502;
    private static final String PREFS="nc_finance_photos";
    private String draft;
    private Uri cameraUri;

    @Override public void onCreate(Bundle state){
        super.onCreate(state);
        draft=getIntent().getStringExtra("draft_id");
        String source=getIntent().getStringExtra("source");
        if(draft==null||!(draft.matches("d_[0-9a-fA-F]{32}")||draft.matches("nci_[0-9a-f]{64}"))){
            Toast.makeText(this,"Seleção de foto inválida.",Toast.LENGTH_SHORT).show();
            finish();return;
        }
        if("camera".equals(source))launchCamera();
        else launchGallery();
    }

    private void launchGallery(){
        try{
            Intent intent=new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            intent.setType("image/*");
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
            startActivityForResult(intent,PICK);
        }catch(Exception error){
            Toast.makeText(this,"Não foi possível abrir a galeria.",Toast.LENGTH_SHORT).show();
            finish();
        }
    }

    private void launchCamera(){
        try{
            ContentValues info=new ContentValues();
            info.put(MediaStore.Images.Media.DISPLAY_NAME,
                "NossoCombinadoFoto_"+System.currentTimeMillis()+".jpg");
            info.put(MediaStore.Images.Media.MIME_TYPE,"image/jpeg");
            cameraUri=getContentResolver().insert(
                MediaStore.Images.Media.EXTERNAL_CONTENT_URI,info);
            if(cameraUri==null)throw new IllegalStateException("cameraUri");
            Intent intent=new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
            intent.putExtra(MediaStore.EXTRA_OUTPUT,cameraUri);
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION|
                            Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
            intent.setClipData(android.content.ClipData.newUri(getContentResolver(),
                 "Foto do comprovante",cameraUri));
            startActivityForResult(intent,CAMERA);
        }catch(Exception error){
            if(cameraUri!=null)try{getContentResolver().delete(cameraUri,null,null);}catch(Exception ignored){}
            Toast.makeText(this,"Câmera indisponível neste aparelho.",Toast.LENGTH_LONG).show();
            finish();
        }
    }

    @Override protected void onActivityResult(int req,int result,Intent resultData){
        super.onActivityResult(req,result,resultData);
        if(req!=PICK&&req!=CAMERA)return;
        if(result!=RESULT_OK){
            if(req==CAMERA&&cameraUri!=null)
                try{getContentResolver().delete(cameraUri,null,null);}catch(Exception ignored){}
            finish();return;
        }
        Uri selected=req==CAMERA?cameraUri:resultData==null?null:resultData.getData();
        if(selected==null){
            Toast.makeText(this,"Imagem não selecionada.",Toast.LENGTH_SHORT).show();
            finish();return;
        }
        try {
            JSONArray existing=new JSONArray(getSharedPreferences(PREFS,MODE_PRIVATE)
                .getString(draft,"[]"));
            if(existing.length()>=6)throw new IllegalArgumentException("Máximo de 6 fotos por cadastro.");
            String mime=getContentResolver().getType(selected);
            if(mime!=null && !mime.startsWith("image/"))
                throw new IllegalArgumentException("Selecione apenas imagens.");
            BitmapFactory.Options size=new BitmapFactory.Options();
            size.inJustDecodeBounds=true;
            try(InputStream in=getContentResolver().openInputStream(selected)){
                if(in==null)throw new IllegalStateException("Sem acesso à foto");
                BitmapFactory.decodeStream(in,null,size);
            }
            if(size.outWidth<1||size.outHeight<1||
               (long)size.outWidth*(long)size.outHeight>65000000L)
                throw new IllegalArgumentException("Imagem inválida ou muito grande.");
            BitmapFactory.Options decode=new BitmapFactory.Options();
            decode.inSampleSize=1;
            while(Math.max(size.outWidth/decode.inSampleSize,size.outHeight/decode.inSampleSize)>1800)
                decode.inSampleSize*=2;
            Bitmap bitmap;
            try(InputStream in=getContentResolver().openInputStream(selected)){
                if(in==null)throw new IllegalStateException("Erro na foto");
                bitmap=BitmapFactory.decodeStream(in,null,decode);
            }
            if(bitmap==null)throw new IllegalArgumentException("Não foi possível ler a imagem.");
            File directory=new File(getFilesDir(),"nc_finance_receipts");
            if(!directory.exists()&&!directory.mkdirs())
                throw new IllegalStateException("Armazenamento indisponível");
            File output=new File(directory,"nc_"+UUID.randomUUID().toString()+".jpg");
            boolean saved;
            try(FileOutputStream stream=new FileOutputStream(output)){
                saved=bitmap.compress(Bitmap.CompressFormat.JPEG,82,stream);
            }
            bitmap.recycle();
            if(!saved||!output.isFile()||output.length()<100)
                throw new IllegalStateException("Falha ao salvar a foto");
            existing.put(output.getCanonicalPath());
            getSharedPreferences(PREFS,MODE_PRIVATE).edit().putString(draft,
                 existing.toString()).apply();
            Toast.makeText(this,"Foto salva neste celular ("+existing.length()+"/6)",
                 Toast.LENGTH_SHORT).show();
        }catch(Exception error){
            Toast.makeText(this,error.getMessage()==null?"Não foi possível salvar a foto.":
              error.getMessage(),Toast.LENGTH_LONG).show();
        }finally{
            finish();
        }
    }
}
