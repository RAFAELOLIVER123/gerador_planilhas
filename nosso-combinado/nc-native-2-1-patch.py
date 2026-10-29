#!/usr/bin/env python3
"""Aplicar mapa híbrido e sons à versão nativa 2.0. Scripts/arquivos de áudio são originais do projeto."""
from pathlib import Path
import sys,base64,zipfile,io
root=Path(sys.argv[1] if len(sys.argv)>1 else 'android-v1')
archive=Path('nosso-combinado/nc-native-2-1-files.zip.b64').read_text(encoding='ascii')
binary=base64.b64decode(archive)
with zipfile.ZipFile(io.BytesIO(binary)) as f:
    for entry in f.infolist():
        target=root/entry.filename
        if '..' in Path(entry.filename).parts or target.is_absolute(): raise ValueError('Arquivo inválido')
        target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(f.read(entry))
assert 'versionCode 210' in (root/'app/build.gradle').read_text()
print('Mapa híbrido, áudio integrado e versão 2.1 prontos')