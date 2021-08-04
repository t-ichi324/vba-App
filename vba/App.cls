VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'-----------------------
' auth : t.ichi
'-----------------------
'■グローバルデフォルトインスタンス■
'
'「Attribute VB_PredeclaredId = true」が設定されています。
'
' 移植時はエディタからコピペではなく、
' エクスポート→テキストエディタでAttributeを確認→インポートの手順を行ってください。

'//////////////////////////////////////
Private Const TBL_APP_DATA = "APP_DATA"
Private Const MSG_LBL_NAME = "lbl_msg"
Private Const TMP_QUERY_NAME = "$_tmp_query"

'--- MemStore
Private m_env As String
Private m_mode As String
Private m_sysTitle As String
Private m_sysVer As String


'//////////////////////////////////////
'--- 初期化用
Public Function INIT_APP_TBL()
On Error GoTo ex
    If Not IsExistsTable(TBL_APP_DATA) Then
        Call CurrentDb.Execute("CREATE TABLE [" & TBL_APP_DATA & "] ([Key] TEXT(50) CONSTRAINT PRIMARYKEY PRIMARY KEY, [Val] CHAR(255));")
        Call CurrentDb.Execute("INSERT INTO [" & TBL_APP_DATA & "] VALUES('env', 'test');")
        Call CurrentDb.Execute("INSERT INTO [" & TBL_APP_DATA & "] VALUES('mode', 'test');")
        Call CurrentDb.Execute("INSERT INTO [" & TBL_APP_DATA & "] VALUES('sys_title', 'System');")
        Call CurrentDb.Execute("INSERT INTO [" & TBL_APP_DATA & "] VALUES('sys_ver', '1.0');")
        Call CurrentDb.Execute("INSERT INTO [" & TBL_APP_DATA & "] VALUES('updated_at', '" & Format(Now, "YYYY-MM-DD") & "');")
    End If
ex:
End Function

'//////////////////////////////////////
'--- デバッグモード判定
Public Function IsDebugMode() As Boolean
    IsDebugMode = App.Env = "test"
End Function


'//////////////////////////////////////
'--- APP_DATA REF
Public Function GetData(ByVal pKey As String, Optional ByVal nullVal As String = "") As String
On Error GoTo ex
    Dim ret As String
    ret = nullVal
    
    pKey = "'" & Replace(pKey, "'", "''") & "'"
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT [Val] FROM [" & TBL_APP_DATA & "] WHERE [Key] = " & pKey & ";")
    
    If Not rs.EOF Then
        ret = Nz(rs.fields("Val"), nullVal)
    End If
ex:
    GetData = ret
    Set rs = Nothing
End Function
Public Function SetData(ByVal pKey As String, ByVal pVal As String)
On Error GoTo ex
    pKey = "'" & Replace(pKey, "'", "''") & "'"
    pVal = "'" & Replace(pVal, "'", "''") & "'"
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT COUNT(0) AS c FROM [" & TBL_APP_DATA & "] WHERE [Key] = " & pKey & ";")
    
    If rs.fields("c") = 0 Then
        Call CurrentDb.Execute("INSERT INTO [" & TBL_APP_DATA & "] VALUES(" & pKey & "," & pVal & ");")
    Else
        Call CurrentDb.Execute("UPDATE [" & TBL_APP_DATA & "] SET [Val] = " & pVal & " WHERE [Key] = " & pKey & ";")
    End If
ex:
    Set rs = Nothing
End Function

'//////////////////////////////////////
'--- PublicProp
Public Property Get Env() As String
    If (m_env = "") Then
        m_env = GetData("env", "real")
    End If
    Env = m_env
End Property
Public Property Get mode() As String
    If (m_mode = "") Then
        m_mode = GetData("mode", "release")
    End If
    mode = m_mode
End Property
Public Property Get SystemTitle() As String
    If (m_sysTitle = "") Then
        m_sysTitle = GetData("sys_title")
    End If
    SystemTitle = m_sysTitle
End Property
Public Property Get SystemVer() As String
    If (SystemVer = "") Then
        m_sysVer = GetData("sys_ver", "1.0.0")
    End If
    SystemVer = m_sysVer
End Property


'//////////////////////////////////////
'--- Accessのファイル名を取得
Public Property Get FileName()
    FileName = Application.CurrentProject.name
End Property

'--- Accessのファイルのパスを取得
Public Property Get FilePath()
    FilePath = Application.CurrentProject.FullName
End Property

'--- Accessのディレクトリのパスを取得
Public Property Get DirPath()
    DirPath = Application.CurrentProject.path
End Property

'--- Accessの終了
Public Sub SystemQuit()
    Call DoCmd.quit(acQuitSaveAll)
End Sub


'//////////////////////////////////////
'--- 共通メッセージラベル
Public Sub MsgLabel(ByVal msg As String)
On Error GoTo ex:
    Dim f As Form
    Set f = Screen.ActiveForm
    If Not f Is Nothing Then
        Dim lbl As Label
        Set lbl = f.Controls.Item("lbl_msg")
        If Not lbl Is Nothing Then
            lbl.Caption = msg
            Exit Sub
        End If
    End If
ex:
    'コントロールが所得できない場合やエラー時
    If msg <> "" Then: Call MsgBox(msg, vbOKOnly)
    
End Sub


'//////////////////////////////////////
'--- テーブル存在チェック
Public Function IsExistsTable(ByVal name As String) As Boolean
On Error GoTo ex
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT COUNT(0) AS c FROM [MSysObjects] WHERE type = 1 AND name = '" & Replace(name, "'", "''") & "'")
    IsExistsTable = (rs.fields("c") <> 0)
ex:
    Set rs = Nothing
End Function
'--- クエリ存在チェック
Public Function IsExistsQuery(ByVal name As String) As Boolean
On Error GoTo ex
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT COUNT(0) AS c FROM [MSysObjects] WHERE type = 5 AND name = '" & Replace(name, "'", "''") & "'")
    IsExistsQuery = (rs.fields("c") <> 0)
ex:
    Set rs = Nothing
End Function

'//////////////////////////////////////
'--- フォームオープン
Public Sub OpenForm(ByVal name As String, Optional ByVal args As String = "")
On Error GoTo ex
    Call DoCmd.Close(acForm, name)
    Call DoCmd.OpenForm(name, acNormal, OpenArgs:=args)
Exit Sub
ex: Call LogException(True, "OpenForm @ " & name)
End Sub
Public Sub CloseForm(ByVal name As String)
On Error GoTo ex
    Call DoCmd.Close(acForm, name)
Exit Sub
ex: Call LogException(True, "CloseForm @ " & name)
End Sub
'--- レポートオープン
Public Sub OpenReport(ByVal name As String, Optional ByVal args As String = "")
On Error GoTo ex
    Call DoCmd.Close(acReport, name)
    Call DoCmd.OpenForm(name, acPreview, OpenArgs:=args)
Exit Sub
ex: Call LogException(True, "OpenReport @ " & name)
End Sub
'--- クエリオープン
Public Sub OpenQuery(ByVal name As String, Optional ByVal mode As AcOpenDataMode = acReadOnly)
On Error GoTo ex
    Call DoCmd.Close(acQuery, name, acSaveNo)
    Call DoCmd.OpenQuery(name, acViewNormal, mode)
Exit Sub
ex: Call LogException(True, "OpenQuery @ " & name)
End Sub
Public Sub OpenQuerySql(ByVal sql As String, Optional ByVal mode As AcOpenDataMode = acReadOnly)
On Error GoTo ex
    Dim def As DAO.QueryDef
    If IsExistsQuery(TMP_QUERY_NAME) Then
        Call DoCmd.Close(acQuery, TMP_QUERY_NAME, acSaveNo)
        Set def = CurrentDb.QueryDefs(TMP_QUERY_NAME)
        def.sql = sql
    Else
        Call CurrentDb.CreateQueryDef(TMP_QUERY_NAME, sql)
    End If

    Call DoCmd.OpenQuery(TMP_QUERY_NAME, acViewNormal, mode)
Exit Sub
Set def = Nothing
ex: Call LogException(True, "OpenQuerySql @ " & """" & sql & """")
End Sub


'//////////////////////////////////////
'--- ログ
Public Function LogInfo(ByVal message As String, Optional ByVal callModuleName As String = "---")
    Call WriteLog("info", LogText(message, callModuleName))
End Function
Public Function LogError(ByVal message As String, Optional ByVal alert As Boolean = True, Optional ByVal callModuleName As String = "---")
    Call WriteLog("error", LogText(message, callModuleName))
    If alert Then
        Call MsgBox(message, vbOKOnly + vbCritical, "ERROR")
    End If
End Function
Public Function LogException(Optional ByVal alert As Boolean = True, Optional ByVal callModuleName As String = "---")
    Dim msg  As String
    msg = err.Description
    Call WriteLog("error", LogText(msg, callModuleName))
    If alert Then
        Call MsgBox(msg, vbOKOnly + vbCritical, "ERROR")
    End If
End Function
Public Function LogSql(ByVal message As String, Optional ByVal callModuleName As String = "---")
    Call WriteLog("sql", LogText(message, callModuleName))
End Function

Private Function LogText(ByVal message As String, ByVal moduleName As String) As String
    Dim txt As String
    txt = "# " & Now & " [ " & moduleName & " ]" & vbCrLf
    txt = txt & "=========================" & vbCrLf
    txt = txt & message & vbCrLf
    txt = txt & "========================="
    LogText = txt
End Function
Private Sub WriteLog(ByVal fName As String, ByVal Log As String)
On Error GoTo ex:
    Dim saveTo As String
    saveTo = Application.CurrentProject.FullName & "." & fName & ".log"
    Dim fso As Object
    Dim f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(saveTo, 8, True)
    Call f.WriteLine(Log & vbCrLf)
ex:
    If Not fso Is Nothing Then: Set fso = Nothing
    If Not f Is Nothing Then: f.Close: Set f = Nothing
End Sub

