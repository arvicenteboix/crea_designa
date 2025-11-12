# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['a.py'],
    pathex=[],
    binaries=[],
    datas=[('AutorizacionGrabacionYDifusion.pdf', '.'), ('AutorizacionUsoMaterialesAbierto.pdf', '.'), ('CuadroTexto.docx', '.'), ('DATOS PONENTE_NOMBRE.pdf', '.'), ('Evidencias.docx', '.'), ('FITXA ECONÒMICA.xlsx', '.'), ('FSE_Ficha_seguimiento.docx', '.'), ('Informe motivado de necesidad de ponente NO FUNCIONARIO CAST.docx', '.'), ('INSTRUCCIONES FACTURACION FACE_2025_sdgfp.pdf', '.'), ('Manual_detallado_FACe-Manual-Proveedores.pdf', '.'), ('Modelo certificado conformidad contrato menor.docx', '.'), ('Modelo informe necesidad_VAL_V3.docx', '.'), ('README.txt', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='a',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
