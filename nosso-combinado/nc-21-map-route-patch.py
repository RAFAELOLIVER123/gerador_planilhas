#!/usr/bin/env python3
from pathlib import Path
import sys
main=Path(sys.argv[1])/'app/src/main/java/com/nossocombinado/app/MainActivity.java'
s=main.read_text(encoding='utf-8')
def change(a,b,label):
 global s
 assert s.count(a)==1,(label,s.count(a))
 s=s.replace(a,b,1);print('OK',label)
change('String[] labels={"🗺 Ruas","🛰 Satélite","🌐 Híbrido"};String[] values={"street","satellite","hybrid"};',
 'String[] labels={"🗺 Ruas","🛰 Satélite"};String[] values={"street","satellite"};','remove hybrid selector')
change('ncMapStyle(vault.get("map_layer").isEmpty()?"hybrid":vault.get("map_layer"));',
 'ncMapStyle(vault.get("map_layer").isEmpty()||"hybrid".equals(vault.get("map_layer"))?"street":vault.get("map_layer"));','street labels by default')
change('map.getOverlays().add(m);ncMapLoadAvatar(m,p.optLong("avatar_media_id",p.optLong("avatar_id",0)),name);',
 '''m.setOnMarkerClickListener((marker,mapView)->{
                    dialogBuilder().setTitle(name).setItems(new String[]{"📌 Marcar como origem","🚗 Traçar rota no mapa","🧭 Abrir no Google Maps","ℹ️ Ver detalhes"},(d,which)->{
                        if(which==0){ncRouteStart=new GeoPoint(lat,lon);msg("Origem definida. Escolha agora o destino no mapa.");}
                        else if(which==1)ncDrawRoute(lat,lon);
                        else if(which==2){try{startActivity(new Intent(Intent.ACTION_VIEW,Uri.parse("https://www.google.com/maps/dir/?api=1&destination="+lat+","+lon+"&travelmode=driving&hl=pt-BR")));}catch(Exception x){msg("Não foi possível abrir o Google Maps.");}}
                        else marker.showInfoWindow();
                    }).show();return true;
                });
                map.getOverlays().add(m);ncMapLoadAvatar(m,p.optLong("avatar_media_id",p.optLong("avatar_id",0)),name);''','click markers opens route')
change('    private void mapScreen(JSONObject j){',
 '''    private GeoPoint ncRouteStart;
    private void ncDrawRoute(double lat,double lon){
        if(ncRouteStart==null){msg("Defina a origem: toque primeiro em um marcador e escolha Marcar como origem.");return;}
        JSONObject q=withHid(data("from_lat",String.valueOf(ncRouteStart.getLatitude()),"from_lng",String.valueOf(ncRouteStart.getLongitude()),"to_lat",String.valueOf(lat),"to_lng",String.valueOf(lon)));
        api.get("route",q,(response,error)->{
            if(error!=null){msg("Rota indisponível: "+error+". Use a opção Google Maps.");return;}
            JSONObject geo=response.optJSONObject("geometry");JSONArray coords=geo==null?null:geo.optJSONArray("coordinates");
            if(coords==null||coords.length()<2){msg("Nenhuma rota viária encontrada. Abra o destino no Google Maps.");return;}
            if(map==null)return;
            ArrayList<GeoPoint> points=new ArrayList<>();
            for(int i=0;i<coords.length();i++){JSONArray xy=coords.optJSONArray(i);if(xy==null||xy.length()<2)continue;points.add(new GeoPoint(xy.optDouble(1),xy.optDouble(0)));}
            if(points.size()<2)return;
            Polyline polyline=new Polyline();polyline.setPoints(points);polyline.getOutlinePaint().setColor(PINK);polyline.getOutlinePaint().setStrokeWidth(dp(5));
            map.getOverlays().add(polyline);map.getController().setCenter(points.get(points.size()/2));map.getController().setZoom(13.0);map.invalidate();
            double km=response.optDouble("distance_m",0)/1000d;msg(km>0?String.format(Locale.getDefault(),"Rota traçada: %.1f km",km):"Rota traçada no mapa.");
        });
    }
    private void mapScreen(JSONObject j){''','native routes')
change('Toque na foto para ver nome, data, horário e coordenadas. Imagens © Esri e fornecedores. © OpenStreetMap contributors. Serviços de imagem podem exigir licença.',
 'Toque na foto para traçar rota ou abrir no Google Maps. Ruas mostra nomes das cidades. © OpenStreetMap contributors. Imagens de satélite podem exigir licença.','route instructions')
main.write_text(s,encoding='utf-8')
