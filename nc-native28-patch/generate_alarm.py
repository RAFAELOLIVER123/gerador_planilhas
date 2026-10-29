#!/usr/bin/env python3
"""Som ORIGINAL de despertador do Nosso Combinado, 16 s, WAV PCM mono.
Bipes distintos e repetitivos com envelope suave; sem mídia de terceiros.
Gerado no build, armazenado no APK para alarmes locais mesmo sem internet."""
import wave,math,random,struct,sys
from pathlib import Path
root=Path(sys.argv[1])
res=root/'app/src/main/res/raw';res.mkdir(parents=True,exist_ok=True)
rate=32000;seconds=16;rng=random.Random(288)
p=res/'nc_alarm05.wav'
with wave.open(str(p),'wb') as w:
 w.setnchannels(1);w.setsampwidth(2);w.setframerate(rate)
 for t in range(rate*seconds):
  pos=t/rate;phase=pos%1.65
  vol=0.0
  if 0.0<phase<.54:vol=min(1,phase/.02,(.54-phase)/.07)
  if .72<phase<1.23:vol=min(1,(phase-.72)/.02,(1.23-phase)/.07)
  chirp=math.sin(2*math.pi*(750*pos+80*pos*pos))
  mix=13500*chirp+3200*math.sin(2*math.pi*1130*pos)+rng.uniform(-1200,1200)
  value=int(max(-32767,min(32767,vol*mix)))
  w.writeframesraw(struct.pack('<h',value))
print('Audio do alarme:',p,p.stat().st_size,'bytes')
