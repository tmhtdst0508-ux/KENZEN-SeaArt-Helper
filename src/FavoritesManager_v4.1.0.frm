VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FavoritesManager 
   Caption         =   "Favorites Manager"
   ClientHeight    =   6030
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7080
   OleObjectBlob   =   "FavoritesManager_v4.1.0.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FavoritesManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PH_FAV_SEARCH As String = "Search Fav"

' =========================================================================
' ★追加：フォームアクティブ時（Hideからの復帰時も含めて、画面表示時に必ず最新化する）
' =========================================================================
Private Sub UserForm_Activate()
    Call RefreshFavList
End Sub

' =========================================================================
' お気に入り管理ウィンドウ (Favorites Manager) の処理
' =========================================================================

' -------------------------------------------------------------------------
' 1. フォーム初期化時：リストボックスにデータを読み込む
' -------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' リストボックスの列設定（1列目:番号、2列目:Description）
    Me.lstYourAllFav.ColumnCount = 2
    'Me.lstYourAllFav.ColumnWidths = "15;270" ' 幅は適宜調整してください
    
    Call RefreshFavList
End Sub

' -------------------------------------------------------------------------
' ★修正：外部（MainWindow等）から強制再描画できるよう Public に変更
' -------------------------------------------------------------------------
Public Sub RefreshFavList()
    Dim favCol As Object
    Dim i As Long
    Dim descStr As String
    
    Me.lstYourAllFav.Clear
    Me.txtFavFullDescription.text = ""
    
    If ConfigData Is Nothing Then Exit Sub
    If Not ConfigData.Exists("Fav_List") Then Exit Sub
    
    Set favCol = ConfigData("Fav_List")
    
    For i = 1 To favCol.Count
        Me.lstYourAllFav.AddItem "No." & Format(i, "00")
        
        ' ★エッジケース3対策：キーが存在するか厳格チェック
        descStr = ""
        On Error Resume Next
        If favCol(i).Exists("Description") Then
            descStr = favCol(i)("Description")
            ' リストボックス内で改行が黒い四角などで化けるのを防ぐため、スペースに置換
            descStr = Replace(descStr, vbLf, " ")
            descStr = Replace(descStr, vbCr, " ")
        End If
        On Error GoTo 0
        
        Me.lstYourAllFav.List(i - 1, 1) = descStr
    Next i
End Sub

' -------------------------------------------------------------------------
' 2. リスト選択時：詳細テキストボックスにDescriptionの全文を表示
' -------------------------------------------------------------------------
Private Sub lstYourAllFav_Change()
    Dim i As Long
    Dim selIndex As Long
    Dim favCol As Object
    
    selIndex = -1
    ' 複数選択（MultiSelect=2）の場合、最初に選択されている項目の詳細を表示
    For i = 0 To Me.lstYourAllFav.ListCount - 1
        If Me.lstYourAllFav.Selected(i) Then
            selIndex = i
            Exit For
        End If
    Next i
    
    If selIndex <> -1 Then
        Set favCol = ConfigData("Fav_List")
        ' Collectionは1始まり、リストボックスは0始まり
        Me.txtFavFullDescription.text = favCol(selIndex + 1)("Description")
    Else
        Me.txtFavFullDescription.text = ""
    End If
End Sub

' -------------------------------------------------------------------------
' 3. メイン画面へ送信（ボタンクリック ＆ ダブルクリック共通）
' -------------------------------------------------------------------------
Private Sub btnSendToFavWindow_Click()
    Call SendSelectedFavToMain
End Sub

Private Sub lstYourAllFav_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call SendSelectedFavToMain
End Sub

' -------------------------------------------------------------------------
' 3. メイン画面へ送信（ボタンクリック ＆ ダブルクリック共通）※日英併記版
' -------------------------------------------------------------------------
Private Sub SendSelectedFavToMain()
    Dim i As Long
    Dim selIndex As Long
    Dim favCol As Object
    
    selIndex = -1
    For i = 0 To Me.lstYourAllFav.ListCount - 1
        If Me.lstYourAllFav.Selected(i) Then
            selIndex = i
            Exit For
        End If
    Next i
    
    If selIndex = -1 Then
        MsgBox "メイン画面に送るお気に入りを選択してください。" & vbCrLf & _
               "Please select a favorite to send to the main window.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Set favCol = ConfigData("Fav_List")
    
    With MainWindow
        ' ★追加：テキストを入れる直前に変数を同期し、Tweaked!の誤点灯を完全に防ぐ
        .lastTweakedValue = favCol(selIndex + 1)("Prompt")
        .txtFav.text = favCol(selIndex + 1)("Prompt")
        .txtDescription.text = favCol(selIndex + 1)("Description")
        
        ' ★念押しでボタンを明示的に無効化
        .btnTweaked.Enabled = False
        .btnTweaked.ForeColor = RGB(160, 160, 160)
        
        .txtFav.SetFocus
    End With
    
    Me.Hide
End Sub


' -------------------------------------------------------------------------
' 4. 選択されたFavの削除（複数選択対応）※日英併記完全版
' -------------------------------------------------------------------------
Private Sub btnSelectedFavDelete_Click()
    Dim i As Long
    Dim favCol As Object
    Dim delCount As Long
    Dim hasSelection As Boolean
    
    ' 選択されているかチェック
    hasSelection = False
    For i = 0 To Me.lstYourAllFav.ListCount - 1
        If Me.lstYourAllFav.Selected(i) Then hasSelection = True: Exit For
    Next i
    
    ' ★修正：未選択時のメッセージを日英併記化
    If Not hasSelection Then
        MsgBox "削除する項目を選択してください。" & vbCrLf & _
               "Please select the item(s) to delete.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 確認メッセージ（元から日英併記のものを維持）
    If MsgBox("選択したお気に入りを削除しますか？" & vbCrLf & "Are you sure you want to delete the selected item(s)?", _
              vbYesNo + vbQuestion, APP_NAME) <> vbYes Then Exit Sub
              
    Set favCol = ConfigData("Fav_List")
    
    ' 削除時はリストの「下から上へ」ループを回す（インデックスのズレを防ぐため）
    delCount = 0
    For i = Me.lstYourAllFav.ListCount - 1 To 0 Step -1
        If Me.lstYourAllFav.Selected(i) Then
            favCol.Remove i + 1 ' Collectionは1始まり
            delCount = delCount + 1
        End If
    Next i
    
    ' 保存とUI同期
    Call SaveConfigJSON
    Call SyncFavSheetFromJSON
    Call RefreshFavList
    
    ' ★修正：完了通知のメッセージを日英併記化
    MsgBox delCount & " 件のお気に入りを削除しました。" & vbCrLf & _
           delCount & " favorite(s) deleted.", vbInformation, APP_NAME
End Sub

' -------------------------------------------------------------------------
' 5. 全削除 ※日英併記版
' -------------------------------------------------------------------------
Private Sub btnAllFavDelete_Click()
    If Me.lstYourAllFav.ListCount = 0 Then Exit Sub
    
    If MsgBox("★警告 (WARNING)★" & vbCrLf & "すべてのお気に入りを完全に削除しますか？この操作は元に戻せません。" & vbCrLf & vbCrLf & _
              "Are you sure you want to delete ALL favorites? This action cannot be undone.", _
              vbYesNo + vbCritical + vbDefaultButton2, APP_NAME) <> vbYes Then Exit Sub
              
    ' 空のコレクションで上書きして初期化
    Set ConfigData("Fav_List") = New Collection
    
    Call SaveConfigJSON
    Call SyncFavSheetFromJSON
    Call RefreshFavList
    
    MsgBox "すべてのお気に入りを削除しました。" & vbCrLf & _
           "All favorites have been deleted.", vbInformation, APP_NAME
End Sub

' -------------------------------------------------------------------------
' 6. 並び替え（上へ / 下へ）※安全のため単一選択時のみ許可
' -------------------------------------------------------------------------
Private Sub btnFavListMoveUp_Click()
    Call MoveFavItem(-1)
End Sub

Private Sub btnFavListMoveDown_Click()
    Call MoveFavItem(1)
End Sub

' -------------------------------------------------------------------------
' 6. 並び替え（上へ / 下へ）※スクロール自動追従・完全版
' -------------------------------------------------------------------------
Private Sub MoveFavItem(direction As Integer)
    Dim i As Long, selCount As Long, selIndex As Long
    
    ' 選択されている項目数をカウント
    selCount = 0
    For i = 0 To Me.lstYourAllFav.ListCount - 1
        If Me.lstYourAllFav.Selected(i) Then
            selCount = selCount + 1
            selIndex = i
        End If
    Next i
    
    ' 複数選択時は複雑な挙動によるデータ破壊を防ぐためブロック
    If selCount <> 1 Then
        MsgBox "並び替えを行うには、リストから項目を「1つだけ」選択してください。" & vbCrLf & _
               "Please select only one item to reorder.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' ★ここで targetIndex を宣言しています
    Dim targetIndex As Long
    targetIndex = selIndex + direction
    
    ' リストの端（一番上、一番下）を超える移動はブロック
    If targetIndex < 0 Or targetIndex >= Me.lstYourAllFav.ListCount Then Exit Sub
    
    Dim favCol As Object
    Set favCol = ConfigData("Fav_List")
    
    ' 入れ替え処理：一度取り出して、目的の位置に挿入し直す
    Dim moveItem As Object
    Set moveItem = favCol(selIndex + 1)
    favCol.Remove selIndex + 1
    
    If direction = -1 Then
        ' 上へ移動
        favCol.Add moveItem, Before:=targetIndex + 1
    Else
        ' 下へ移動（Removeした後なのでインデックスが1つズレているため、targetIndexそのままがAfterになる）
        favCol.Add moveItem, After:=targetIndex
    End If
    
    ' データの物理保存と再描画
    Call SaveConfigJSON
    Call SyncFavSheetFromJSON
    Call RefreshFavList
    
    ' 移動した項目を再び選択状態にして、連続クリックできるようにする
    Me.lstYourAllFav.Selected(targetIndex) = True
    
    ' =========================================================
    ' ★ここに自動スクロール追従を組み込みました！
    ' =========================================================
    Me.lstYourAllFav.TopIndex = targetIndex
End Sub

' -------------------------------------------------------------------------
' 7. ウィンドウを閉じる
' -------------------------------------------------------------------------
Private Sub btnExitFavManager_Click()
' ★エッジケース1対策：メイン画面に残っている古い表示をリセットして不整合を防ぐ
    With MainWindow
        .txtFav.text = ""
        .txtDescription.text = ""
        .txtFavSearch.text = PH_FAV_SEARCH
    End With
    
    Me.Hide
End Sub


' -------------------------------------------------------------------------
' 右上の「X」ボタンが押されたときの挙動を「Hide」にすり替える安全装置
' -------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' CloseMode = vbFormControlMenu (0) は、ユーザーが右上の「X」を押したことを示す
    If CloseMode = vbFormControlMenu Then
        
        ' 1. デフォルトの強制Unload（メモリ破棄）をキャンセルして割り込む
        Cancel = True

        ' ★エッジケース1対策：メイン画面に残っている古い表示をリセットして不整合を防ぐ
        With MainWindow
            .txtFav.text = ""
            .txtDescription.text = ""
            .txtFavSearch.text = PH_FAV_SEARCH
        End With
        
        ' 2. Exitボタンと全く同じ「Me.Hide」を実行する
        Me.Hide
        
    End If
End Sub

