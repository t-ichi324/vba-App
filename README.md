# vba-App

VBA (Microsoft.Access)のアプリケーション制御のクラスモジュールです。 
システムの情報を管理する「APP_DATA」が必要になります。イミディエイトから「App.INIT_APP_TBL()」を実行してください。

#APP_DATAテーブルに登録されたデータにアクセスします。

App.GetData()
App.SetData()
```
MsgBox App.GetData("sample_key", "データが存在しません")

Call App.SetData("sample_key", "初期ユーザ")
MsgBox App.GetData("sample_key")

Call App.SetData("sample_key", "更新しました")
MsgBox App.GetData("sample_key")
```

