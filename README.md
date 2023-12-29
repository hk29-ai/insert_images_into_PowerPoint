# 概要  
Pythonにて複数の画像ファイルをMicrosoftのパワーポイント（PowerPoint）に貼り付ける雛形コードです。  
1スライドあたりに貼り付ける列数を指定することで、画像サイズを自動で調整します。  
そして、貼り付ける行方向の枚数を自動で算出して、1スライドで収まらない場合は次のスライドへ順次貼り付けてゆきます。  

# 使用ライブラリ
「python-pptx」と「Pillow」です  
```
pip install python-pptx
pip install Pillow  
```
