# Build (Windows) - PyInstaller

## 1) Ortam
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller
```

## 2) Tek dosya EXE
```bash
pyinstaller -F -w -n MedarYakaKart ^
  --icon assets\medar.ico ^
  -m medar_yakakart
```

- `-w`: konsol penceresini kapatır (GUI)
- `-m medar_yakakart`: modül entrypoint’i (`python -m medar_yakakart`) kullanır

## 3) Çıktılar
- `dist/MedarYakaKart.exe`

## 4) Notlar
- `7z.exe` paketlenmek istenirse exe yanına koyun veya installer ile ekleyin.
- `rarfile` kullanıyorsanız: sistemde unrar/7z gereksinimlerini ayrıca dokümante edin.
