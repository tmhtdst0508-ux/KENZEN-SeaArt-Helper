VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==========================================
' ThisWorkbook モジュール
' ==========================================
Private Sub Workbook_Open()
    ' ★ 1. JSONファイルの検索・初期化・読み込みを実行（新規追加）
    Call InitConfigJSON
    
    ' 2. ショートカットキーの割り当て
    Application.OnKey "^+P", "SetPositiveAction"
    Application.OnKey "^+O", "OpenMainWindowAction"

    ' 3. ブックを開いた時に自動でパネルを起動
    Call OpenMainWindowAction
End Sub
' ==========================================
' 【自動実行】ハイパーリンク移動時に左端へ寄せる
' ==========================================
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    On Error Resume Next
    
    ' ハイパーリンクの移動先（名前定義やセル番地）を取得して、
    ' その場所に「左端（左上）固定」でジャンプし直します。
    Dim destination As String
    destination = Target.SubAddress
    
    If destination <> "" Then
        ' Application.Goto を使うことで、スクロール位置を強制的に左上に持っていきます。
        Application.GoTo Reference:=Range(destination), Scroll:=True
    End If
    
    On Error GoTo 0
End Sub

' ブックを閉じる時にショートカットを解除（他のExcelへの影響防止）
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.OnKey "^+P"
    Application.OnKey "^+O"
End Sub

