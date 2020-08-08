[Excel用デフォルトインスタンス設定アドイン](https://github.com/furyutei/DefaultClassInstanceSetting)
====================================================================================================

- License: The MIT license  
- Copyright (c) 2020 風柳(furyu)  
- 対象Excel: Microsoft® Excel® for Microsoft 365 MSO 32ビット
- 対象OS: Windows 10

[Excel VBA クラスモジュールのデフォルトインスタンス有効／無効切替が面倒だった](https://twitter.com/furyutei/status/1291670503506055168)ので試作したアドイン。  

クラスモジュールのデフォルトインスタンスを有効にするには、エクスポートしたソースコードに

```vb
Attribute VB_PredeclaredId = True
```

と設定した後でインポートし直す必要があり、手動で行うと面倒なので、1クリックで変更できる設定フォームを作ってみた。  


■⚠ ご注意 ⚠
---
- 一切の動作は保証しません
- モジュールを直接書き換えるものであるため、事前にバックアップを取るなど十分ご注意の上、自己責任にてお試しください
- ファイル→オプション→トラスト センター→[トラスト センターの設定(T)...]→マクロの設定→開発者向けマクロ設定→「☑ VBA プロジェクト オブジェクト モデルへのアクセスを信頼する(V)」にチェックを入れてからご使用ください


■ インストール
---
1. 右上の [Code ▽] → 「Download ZIP」でダウンロード
2. Excel が起動している場合、終了させる
3. ZIP ファイルを展開して出てくる addin フォルダ中の Install.vbs をダブルクリックし、指示に従う  


■ 使い方
---
インストールすると、リボンの「アドイン」タブに「Default Class Instance Setting」というメニューコマンドが現れるので、これをクリックすると設定フォームが表示される。  

![設定フォーム](https://github.com/furyutei/DefaultClassInstanceSetting/blob/images/DefaultClassInstanceSetting.Menu.png)

- ブック及びクラスモジュール名を選択すると、現在のデフォルトインスタンス作成状態（Enabled:作成される／Disabled:作成されない）を表示  
- ラジオボタンをクリックすると、当該クラスモジュールの設定を切り替え  


■ アンインストール
---
1. Excel が起動している場合、終了させる
2. addin フォルダ中の Uninstall.vbs をダブルクリックし、指示に従う  
