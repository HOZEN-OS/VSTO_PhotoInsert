# PhotoInsert


エクセルで写真を取り込む時のアドインです。

![ribbon](https://user-images.githubusercontent.com/78771008/191179031-4eb8a081-abae-4c0b-ba70-c634645f8480.png)


選択されたセル（複数選択可）に合わせて写真を取り込みます。

位置、トリミング、サイズで調整が可能です。（一部無意味な組み合わせがあります）

 

基本的に選択されたセルの大きさに合わせ、はみ出さないように縦横のサイズを調整し、選択されたセルの左上を起点に写真を挿入します。

 

#### 【位置】

横センター：選択されたセル幅が写真の幅より広い場合に左右の空白が同じになるようにします。

縦センター：選択されたセル高が写真の高さより高い場合に上下の空白が同じになるようにします。

 

#### 【トリミング】

選択セル：選択されたセルにぴったり合うように写真をトリミングします。

１６：９：横幅16対高さ9に写真をトリミングします。

４：４  ：横幅４対高さ３に写真をトリミングします。

１：１  ：正方形に写真をトリミングします。

 

#### 【サイズ】

横幅で合わせる：選択されたセルの高さを無視して選択されたセルの横幅に合わせます。

高さで合わせる：選択されたセルの横幅を無視して選択されたセルの高さに合わせます。

圧縮する：取込時に画像をJPEGに圧縮します。　※


#### 【その他】

全て圧縮：シート上の全ての画像を圧縮します。　※


※　圧縮サイズは、オプション→詳細設定→既定の解像度のサイズになります。



この作品は [クリエイティブ・コモンズ 表示 - 非営利 - 改変禁止 4.0 国際 ライセンス](http://creativecommons.org/licenses/by-nc-nd/4.0/) の下に提供されています。
[![CC BY-NC-ND 4.0](https://i.creativecommons.org/l/by-nc-nd/4.0/88x31.png)](http://creativecommons.org/licenses/by-nc-nd/4.0/)

