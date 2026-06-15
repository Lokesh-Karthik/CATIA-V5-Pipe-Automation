# Building CATIA_Pipe_Suite.exe

## Quick build (one command)

```
pip install pyinstaller pywin32 pyyaml
pyinstaller CATIA_Pipe_Suite.spec
```

The EXE and all bundled dependencies will be in:

```
dist/
└── CATIA_Pipe_Suite/
    ├── CATIA_Pipe_Suite.exe   ← double-click to launch
    ├── main.py
    ├── src/
    │   ├── catia_connection.py
    │   ├── solid_analyzer.py
    │   ├── centerline_builder.py
    │   ├── surface_builder.py
    │   ├── solid_converter.py
    │   ├── parametric_pipeline.py
    │   └── utils.py
    └── ... (pyinstaller runtime files)
```

## Distributing to another machine

Copy the entire `dist/CATIA_Pipe_Suite/` folder to the target machine.
The target machine needs:
- Windows 10 or 11
- CATIA V5 R19–R21 installed (for COM automation)
- NO Python required — it's all bundled

## Single-file EXE (optional)

If you want a single `.exe` instead of a folder, change the spec file:

```python
# In CATIA_Pipe_Suite.spec, replace the EXE(...) block with:
exe = EXE(
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas,
    name='CATIA_Pipe_Suite',
    console=False,
    onefile=True,   # <-- add this
)
```

Or run:
```
pyinstaller --onefile --windowed --name CATIA_Pipe_Suite catia_suite_gui.py
```

Note: single-file EXE is slower to start (extracts on each launch).

## Troubleshooting build

| Error | Fix |
|-------|-----|
| `win32com not found` | Run `pip install pywin32` then `python Scripts/pywin32_postinstall.py -install` |
| `Module not found: yaml` | `pip install pyyaml` |
| App launches but pipeline fails | Make sure `main.py` and `src/` are in the same folder as the EXE |
