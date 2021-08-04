# vba-App

VBA (Microsoft.Access)のアプリケーション制御のクラスモジュールです。 
システムの情報を管理する「APP_DATA」が必要になります。イミディエイトから「App.INIT_APP_TBL()」を実行してください。


## App.GetData() / App.SetData()
APP_DATAテーブルに登録されたデータにアクセスします。
```
MsgBox App.GetData("sample_key", "データが存在しません")

Call App.SetData("sample_key", "初期ユーザ")
MsgBox App.GetData("sample_key")

Call App.SetData("sample_key", "更新しました")
MsgBox App.GetData("sample_key")
```

## App.Env() / App.Mode() / App.SystemTitle() / App.SystemVer()
App.GetData()固定値プロパティです。
```
MsgBox App.GetData("env")
MsgBox App.Env

MsgBox App.GetData("mode")
MsgBox App.Mode

MsgBox App.GetData("sys_title")
MsgBox App.SystemTitle

MsgBox App.GetData("sys_ver")
MsgBox App.SystemVer
```

