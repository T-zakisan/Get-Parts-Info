# Get-Parts-Info

## 準備
`python`は、インストール済みかつ、パスが通っているものとする。


### モジュールのインストール
```
pip install --force-reinstall -v pyperclip==1.8.2　#1.9：メソッドが使えない
pip install pyodbc
pip install threading
pip install pystray
pip install PIL
pip install PyQt5
pip install pathlib
pip install six
```

### ファイルの準備
下記のファイルをDLする
- `favicon.ico`
- `ClipSearch_qt.py`



## コンパイル

```
pyinstaller.exe   ClipSearch_qt.py   -D   -F   --add-data "./favicon.ico;./"   -w   --clean   --icon="favicon.ico"
```



## 使い方
1. 生成した実行ファイルを適当なディレクトリに移動
2. `部品番号台帳.accdb`を生成した実行ファイルと同じディレクトリに移動あるいはコピー
3. 生成した実行ファイルを起動
