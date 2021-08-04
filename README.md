# vba-App

VBA (Microsoft.Access)のアプリケーション汎用ライブラリです。 
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

## App.Env / App.Mode / App.SystemTitle / App.SystemVer
App.GetData()の固定値プロパティです。
```
'MsgBox App.GetData("env")と同じ
MsgBox App.Env

'MsgBox App.GetData("mode")と同じ
MsgBox App.Mode

'MsgBox App.GetData("sys_title")と同じ
MsgBox App.SystemTitle

'MsgBox App.GetData("sys_ver")と同じ
MsgBox App.SystemVer
```

## App.FileName / App.FilePath / App.DirPath
Application.CurrentProjectのエイリアスです。
```
'Application.CurrentProject.nameのエイリアス
MsgBox App.FileName
'Application.CurrentProject.FullNameのエイリアス
MsgBox App.FullName
'Application.CurrentProject.pathのエイリアス
MsgBox App.DirPath
```

## App.MsgLabel()
アクティブフォームのLabelにメッセージを表示する共通メソッドです。MsgBoxのポップアップが煩わしいときに利用してください。
アクティブフォームに指定のコントロー名のLabelが存在しない場合はMsgBoxを表示します。
```
Call App.MsgLabel("こんにちは！")
```
コントロール名を変更するときは下記の定義を変更してください。
```
Private Const MSG_LBL_NAME = "lbl_msg"
```

## App.IsExistsTable() / App.IsExistsQuery()
MSysObjectsを参照してテーブル、クエリが存在しているかチェックします。
```
If App.IsExistsTable("tbl_user") Then
    MsgBox "tbl_userテーブルは存在します"
Else
    MsgBox "tbl_userテーブルは存在しません"
End If

If App.IsExistsQuery("tbl_user") Then
    MsgBox "tbl_userクエリは存在します"
Else
    MsgBox "tbl_userクエリは存在しません"
End If
```
