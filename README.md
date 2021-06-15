# 参考
- [Excelマクロ（VBA）をVSCodeで編集したい（Git管理も）](https://kanegolabo.com/vba-edit)
- [VBAサンプル集](https://excel-ubara.com/excelvba5/)
# cmd.exeでワークフォルダに移動
## ディレクトリ構成
```
./
    - bin
        - target.xlsm
    - src
    - readme.md
    - vbac.wsf
```
## 操作
```cmd
-- エクスポート
cscript vbac.wsf decombine
-- インポート
cscript vbac.wsf combine
```
# VSCode plugin
- VSCode VBA
- vba-snippets
- vscode-vba-icons
# 文字化け問題
excelはsjis、vscodeはutf-8の為、.dcmファイルをvscodeで開くと文字化けしている。  
対策は、vscode画面右下のUTF-8欄をクリックし、エンコード付きで再オープン(sjis指定)を行う。
