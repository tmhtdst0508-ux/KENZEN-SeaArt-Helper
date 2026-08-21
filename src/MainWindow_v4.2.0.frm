VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainWindow 
   Caption         =   "KENZEN SeaArt Helper v4.2.0"
   ClientHeight    =   6050
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9740.001
   OleObjectBlob   =   "MainWindow.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- Windows API 宣言（最小化ボタン用） ---
#If VBA7 Then
    ' 64bit / 32bit 両対応
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    ' 旧バージョン用
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

' --- Xボタン無効化のためのWindows API宣言 ---
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr
    Private Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hMenu As LongPtr, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If

' --- ゴースティング（応答なしホワイトアウト）無効化API ---
#If VBA7 Then
    Private Declare PtrSafe Sub DisableProcessWindowsGhosting Lib "user32" ()
#Else
    Private Declare Sub DisableProcessWindowsGhosting Lib "user32" ()
#End If

' ==========================================
' 状態保持用のモジュールレベル変数
' (ユーザーフォームまたはシートの先頭に宣言)
' ==========================================
' Dim lastSearchText As String
' Dim lastFoundCell As Range
Private isSearchingComments As Boolean ' ←これを追加してください（検索モード判定用）

Private Const SC_CLOSE As Long = &HF060&
Private Const MF_BYCOMMAND As Long = &H0&
Private Const MF_GRAYED As Long = &H1&
Private Const MF_ENABLED As Long = &H0&
Private Const GWL_STYLE As Long = -16
Private Const WS_MINIMIZEBOX As Long = &H20000
' ==========================================
' 宣言セクション（モジュールの一番上）
' ==========================================
Const colAction As Long = &HFFE1C8   ' RGB(200, 225, 255) 相当
Const colSuccess As Long = &HC8F0C8  ' RGB(200, 240, 200) 相当
Const colDisabled As Long = &HE0E0E0 ' 明るいグレー
Const PH_FAV_SEARCH As String = "Search Fav"

' --- 検索の継続性を保つための変数（モジュールレベル） ---
Private lastSearchText As String
Private lastFoundCell As Range

' --- 定数の定義 ---
Private Const PH_TEXT As String = "Enter Keyword"
Private Const COLOR_GRAY As Long = &HD0D0D0   ' 灰色
Private Const COLOR_BLACK As Long = &H80000008 ' 黒

' --- お気に入り検索の状態管理用 ---
Private lastSearchKeyFav As String
Private lastSearchIdxFav As Long
Private Const PH_NEGA_INPUT As String = "Enter New Negative"
Private Const PH_POSI_INPUT As String = "Enter New Positive" ' ★この行を追加
Private isTypingSession As Boolean
Private PreEditText As String
' --- フォーム全体のフラグ変数 ---
Private m_IsGachaRunning As Boolean ' ガチャ処理中かどうかを判定するフラグ
Public lastTweakedValue As String ' Tweakedボタンで最後にコピーした内容を保持
'-------------------------------------------------------------------------
' LoRAタブ内のUI要素の状態を動的に更新する（オプションボタン・グレーアウト完全対応版）
'-------------------------------------------------------------------------
Public Sub UpdateLoRAUIState()
    Dim isUse As Boolean
    Dim hasListItems As Boolean
    Dim hasListSelected As Boolean
    Dim hasPreviewText As Boolean
    Dim hasLoRASelected As Boolean
    Dim i As Integer
    Dim bgDefault As Long: bgDefault = vbWindowBackground
    Dim bgDisabled As Long: bgDisabled = vbButtonFace
    
    isUse = Me.chkUseLoRA.Value
    hasListItems = (Me.lstSelectedLoRA.ListCount > 0)
    hasListSelected = (Me.lstSelectedLoRA.ListIndex >= 0)
    hasPreviewText = (Trim(Me.txtLoRAPreview.Value) <> "")
    hasLoRASelected = (Me.cmbLoRAList.ListIndex >= 0) ' LoRAが選ばれているか
    
    ' 1. 各コントロールの有効・無効（Enabled）と背景色（BackColor）の動的制御
    Me.cmbLoRAList.Enabled = isUse
    Me.cmbLoRAList.BackColor = IIf(isUse, bgDefault, bgDisabled)
    
    ' ★仕様変更：LoRAリストで値が選択されて初めて強度(Strength)をアクティブ化
    Dim canEnableStrength As Boolean
    canEnableStrength = (isUse And hasLoRASelected)
    Me.cmbLoRAStrength.Enabled = canEnableStrength
    Me.cmbLoRAStrength.BackColor = IIf(canEnableStrength, bgDefault, bgDisabled)
    
    Me.lstSelectedLoRA.Enabled = isUse
    Me.lstSelectedLoRA.BackColor = IIf(isUse, bgDefault, bgDisabled)
    
    ' ★仕様変更：カート(lstSelectedLoRA)に値が入って初めてトリガー重み(Weight)をアクティブ化
    Dim canEnableTriggerWeight As Boolean
    canEnableTriggerWeight = (isUse And hasListItems)
    Me.cmbTriggerWeight.Enabled = canEnableTriggerWeight
    Me.cmbTriggerWeight.BackColor = IIf(canEnableTriggerWeight, bgDefault, bgDisabled)
    
    Me.cmbPresetList.Enabled = isUse
    Me.cmbPresetList.BackColor = IIf(isUse, bgDefault, bgDisabled)
    
    Me.txtLoRAPreview.Enabled = isUse
    Me.txtLoRAPreview.BackColor = IIf(isUse, bgDefault, bgDisabled)
    
    For i = 1 To 10
        Me.Controls("chkTrigger_" & i).Enabled = isUse
    Next i
    
    ' 2. 状態に関わらず常に有効なもの（生命線）
    Me.btnOpenManageLoRA.Enabled = True
    
    ' 3. 基本操作ボタンのグレーアウト制御
    ' ★仕様変更：コンボボックスでLoRAが選択されて初めて有効化する
    Dim canOperateLoRA As Boolean
    canOperateLoRA = (isUse And hasLoRASelected)
    
    Call SetControlUIState(Me.btnSetLoRA, canOperateLoRA, &HCCFFCC)
    Call SetControlUIState(Me.btnLoRACancel, canOperateLoRA, &HCCE5FF)
    
    Call SetControlUIState(Me.btnCallLoRAPreset, isUse, &HFFE5CC)
    Call SetControlUIState(Me.btnDeleteLoRAPreset, isUse, &HCCCCFF)
    
    ' 4. 状況に合わせて動的にロック・アンロックされるボタン群
    Dim canWrap As Boolean
    canWrap = (isUse And hasListItems)
    
    ' Wrapボタンの有効状態（canWrap）を取得
    Call SetControlUIState(Me.btnWrapLoRA, canWrap, &HFFE5CC)
    
    ' ★追加：opbWrapHash と opbWrapName を btnWrapLoRA の状態（canWrap）に完全同期！
    Me.opbWrapHash.Enabled = canWrap
    Me.opbWrapName.Enabled = canWrap
    
    Call SetControlUIState(Me.btnRemoveLoRA, (isUse And hasListSelected), &HC0E0FF)
    Call SetControlUIState(Me.btnForgetLoRA, (isUse And (hasListItems Or hasPreviewText)), &HCCCCFF)
    Call SetControlUIState(Me.btnSendLoRAtoCockpit, (isUse And hasPreviewText), &HC8F0C8)
    Call SetControlUIState(Me.btnSendLoRAtoFav, (isUse And hasPreviewText), &HC8F0C8)
    Call SetControlUIState(Me.btnClearLoRAPreview, (isUse And hasPreviewText), &HCCE5FF)
    Call SetControlUIState(Me.btnSaveAsPresetLoRA, (isUse And hasPreviewText), &HFFE1C8)
End Sub




' =========================================================================
' 共通処理：新しいポジティブプロンプトの登録（高精度クレンジング・全角排除・JSON対応版）
' =========================================================================
Private Sub AddNewPositiveWord()
    Dim rawInput As String
    Dim cleanedOutput As String
    Dim i As Long, k As Long
    Dim char As String
    Dim inWeight As Boolean
    Dim prevLen As Long
    
    Dim newWords() As String
    Dim word As String
    Dim isDuplicate As Boolean
    Dim addedCount As Long
    Dim skippedFullWidthCount As Long
    Dim finalStock As String
    
    rawInput = Trim(Me.txtInputPositive.text)
    
    If rawInput = "" Or rawInput = PH_POSI_INPUT Then
        MsgBox "インポートするプロンプトをテキストボックスに入力（貼り付け）してください。" & vbCrLf & _
               "Please paste the prompt into the textbox first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' ★修正：StrConvを廃止し、安全な正規化関数を適用（カタカナは全角のまま維持される）
    rawInput = NormalizeToHalfWidth(rawInput)
    
    ' --- 高精度サニタイズ（クレンジング）ロジック ---
    cleanedOutput = ""
    inWeight = False
    
    For i = 1 To Len(rawInput)
        char = Mid(rawInput, i, 1)
        
        If char = "(" Or char = ")" Or char = "[" Or char = "]" Then
            ' カッコ自体は無視してスキップ
        ElseIf char = ":" Then
            inWeight = True
        ElseIf inWeight And (char = "," Or char = vbLf Or char = vbCr) Then
            inWeight = False
            cleanedOutput = cleanedOutput & ","
        Else
            If Not inWeight Then cleanedOutput = cleanedOutput & char
        End If
    Next i
    
    ' --- 整地ロジック：改行のカンマ化とゴミ掃除 ---
    cleanedOutput = Replace(cleanedOutput, vbCrLf, ",")
    cleanedOutput = Replace(cleanedOutput, vbCr, ",")
    cleanedOutput = Replace(cleanedOutput, vbLf, ",")
    cleanedOutput = Replace(cleanedOutput, " ,", ",")
    cleanedOutput = Replace(cleanedOutput, ", ", ",")
    
    Do
        prevLen = Len(cleanedOutput)
        cleanedOutput = Replace(cleanedOutput, ",,", ",")
    Loop While prevLen <> Len(cleanedOutput)
    
    If Left(cleanedOutput, 1) = "," Then cleanedOutput = Mid(cleanedOutput, 2)
    If Right(cleanedOutput, 1) = "," Then cleanedOutput = Left(cleanedOutput, Len(cleanedOutput) - 1)
    
    cleanedOutput = Trim(cleanedOutput)
    If cleanedOutput = "" Then Exit Sub
    
    ' =========================================================
    ' 重複チェック ＆ 全角文字（漢字・カタカナ等）の排除バリデーション
    ' =========================================================
    newWords = Split(cleanedOutput, ",")
    addedCount = 0
    skippedFullWidthCount = 0
    
    For i = 0 To UBound(newWords)
        word = Trim(newWords(i))
        
        If word <> "" Then
            ' Len（文字数）と LenB（バイト数）が一致しなければ全角と判定
            If Len(word) <> LenB(StrConv(word, vbFromUnicode)) Then
                ' ★先ほどの正規化により、全角カタカナも全角のままここへ到達するため、確実に弾かれます
                skippedFullWidthCount = skippedFullWidthCount + 1
            Else
                isDuplicate = False
                
                For k = 0 To Me.lstAllPositive.ListCount - 1
                    If StrComp(Me.lstAllPositive.List(k), word, vbTextCompare) = 0 Then
                        isDuplicate = True
                        Exit For
                    End If
                Next k
                
                If Not isDuplicate Then
                    Me.lstAllPositive.AddItem word
                    addedCount = addedCount + 1
                End If
            End If
        End If
    Next i
    
    ' JSONの「PositiveStock」を最新のリスト状態で再構築して保存
    finalStock = ""
    For k = 0 To Me.lstAllPositive.ListCount - 1
        finalStock = finalStock & Me.lstAllPositive.List(k) & ", "
    Next k
    
    If finalStock <> "" Then finalStock = Left(finalStock, Len(finalStock) - 2)
    
    If ConfigData Is Nothing Then Call InitConfigJSON
    ConfigData("PositiveStock") = finalStock
    Call SaveConfigJSON
    
    Me.txtInputPositive.text = ""
    
    If addedCount > 0 Then
        Dim statusMsg As String
        statusMsg = "● Added " & addedCount & " new prompt(s) to PositiveStock."
        If skippedFullWidthCount > 0 Then
            statusMsg = statusMsg & " (" & skippedFullWidthCount & " item(s) skipped due to full-width characters.)"
        End If
        Application.StatusBar = statusMsg
    Else
        MsgBox "新しいタグは追加されませんでした。" & vbCrLf & _
               "入力されたプロンプトはすべて重複しているか、日本語（漢字・カタカナなど）が含まれているため除外されました。" & vbCrLf & vbCrLf & _
               "No new tags were added." & vbCrLf & _
               "All entered prompts were either duplicates or excluded due to containing full-width characters.", _
               vbInformation, APP_NAME
               
        Application.StatusBar = "● No new prompts added (Duplicates or Invalid characters)."
    End If
End Sub


' =========================================================================
' 共通処理：新しいネガティブプロンプトの登録（高精度クレンジング・全角排除・JSON対応版）
' =========================================================================
Private Sub AddNewNegativeWord()
    Dim rawInput As String
    Dim cleanedOutput As String
    Dim i As Long, k As Long
    Dim char As String
    Dim inWeight As Boolean
    Dim prevLen As Long
    
    Dim newWords() As String
    Dim word As String
    Dim isDuplicate As Boolean
    Dim addedCount As Long
    Dim skippedFullWidthCount As Long
    
    rawInput = Trim(Me.txtInputNegative.text)
    
    If rawInput = "" Or rawInput = PH_NEGA_INPUT Then
        MsgBox "インポートするプロンプトをテキストボックスに入力（貼り付け）してください。" & vbCrLf & _
               "Please paste the prompt into the textbox first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' ★修正：StrConvを廃止し、安全な正規化関数を適用
    rawInput = NormalizeToHalfWidth(rawInput)
    
    ' --- 高精度サニタイズ（クレンジング）ロジック ---
    cleanedOutput = ""
    inWeight = False
    
    For i = 1 To Len(rawInput)
        char = Mid(rawInput, i, 1)
        
        If char = "(" Or char = ")" Or char = "[" Or char = "]" Then
        ElseIf char = ":" Then
            inWeight = True
        ElseIf inWeight And (char = "," Or char = vbLf Or char = vbCr) Then
            inWeight = False
            cleanedOutput = cleanedOutput & ","
        Else
            If Not inWeight Then cleanedOutput = cleanedOutput & char
        End If
    Next i
    
    ' --- 整地ロジック：改行のカンマ化とゴミ掃除 ---
    cleanedOutput = Replace(cleanedOutput, vbCrLf, ",")
    cleanedOutput = Replace(cleanedOutput, vbCr, ",")
    cleanedOutput = Replace(cleanedOutput, vbLf, ",")
    cleanedOutput = Replace(cleanedOutput, " ,", ",")
    cleanedOutput = Replace(cleanedOutput, ", ", ",")
    
    Do
        prevLen = Len(cleanedOutput)
        cleanedOutput = Replace(cleanedOutput, ",,", ",")
    Loop While prevLen <> Len(cleanedOutput)
    
    If Left(cleanedOutput, 1) = "," Then cleanedOutput = Mid(cleanedOutput, 2)
    If Right(cleanedOutput, 1) = "," Then cleanedOutput = Left(cleanedOutput, Len(cleanedOutput) - 1)
    
    cleanedOutput = Trim(cleanedOutput)
    If cleanedOutput = "" Then Exit Sub
    
    ' =========================================================
    ' 重複チェック ＆ 全角文字（漢字・カタカナ等）の排除バリデーション
    ' =========================================================
    newWords = Split(cleanedOutput, ",")
    addedCount = 0
    skippedFullWidthCount = 0
    
    For i = 0 To UBound(newWords)
        word = Trim(newWords(i))
        
        If word <> "" Then
            If Len(word) <> LenB(StrConv(word, vbFromUnicode)) Then
                skippedFullWidthCount = skippedFullWidthCount + 1
            Else
                isDuplicate = False
                
                For k = 0 To Me.lstAllNegative.ListCount - 1
                    If StrComp(Me.lstAllNegative.List(k), word, vbTextCompare) = 0 Then
                        isDuplicate = True
                        Exit For
                    End If
                Next k
                
                If Not isDuplicate Then
                    Me.lstAllNegative.AddItem word
                    addedCount = addedCount + 1
                End If
            End If
        End If
    Next i
    
    ' 最新のリスト状態をJSON（NegativeStock）へ同期保存
    Call SyncNegativeSheet
    
    Me.txtInputNegative.text = ""
    
    If addedCount > 0 Then
        Dim statusMsg As String
        statusMsg = "● Added " & addedCount & " new negative prompt(s) to NegativeStock."
        If skippedFullWidthCount > 0 Then
            statusMsg = statusMsg & " (" & skippedFullWidthCount & " item(s) skipped due to full-width characters.)"
        End If
        Application.StatusBar = statusMsg
    Else
        MsgBox "新しいタグは追加されませんでした。" & vbCrLf & _
               "入力されたプロンプトはすべて重複しているか、日本語（漢字・カタカナなど）が含まれているため除外されました。" & vbCrLf & vbCrLf & _
               "No new tags were added." & vbCrLf & _
               "All entered prompts were either duplicates or excluded due to containing full-width characters.", _
               vbInformation, APP_NAME
               
        Application.StatusBar = "● No new prompts added (Duplicates or Invalid characters)."
    End If
End Sub


'-------------------------------------------------------------------------
' 強度(Strength)のコンボボックスの選択肢を生成するプロシージャ
'-------------------------------------------------------------------------
Private Sub InitLoRAStrengthCombo()
    Dim j As Long
    
    Me.cmbLoRAStrength.Clear
    
    ' -2.5 から 7.0 まで 0.05刻み
    ' (-2.5 * 20 = -50) ～ (7.0 * 20 = 140)
    For j = -50 To 140
        ' 0.00 の形式でフォーマットして追加
        Me.cmbLoRAStrength.AddItem Format(j / 20, "0.00")
    Next j
End Sub
'-------------------------------------------------------------------------
' トリガーワードの重み(Weight)のコンボボックスの選択肢を生成する
'-------------------------------------------------------------------------
Private Sub InitTriggerWeightCombo()
    Dim i As Integer
    
    Me.cmbTriggerWeight.Clear
    
    ' 1.0 から 1.5 まで 0.1 刻みで追加 (10から15で回して10で割る)
    For i = 10 To 15
        Me.cmbTriggerWeight.AddItem Format(i / 10, "0.0")
    Next i
    
    ' デフォルト値として "1.0" を選択状態にしておく（ユーザーへの安全な提案）
    Me.cmbTriggerWeight.ListIndex = 0
End Sub
'-------------------------------------------------------------------------
' [All Reset] ボタン：コンフィグデータを完全に工場出荷状態へリセット
'-------------------------------------------------------------------------
Private Sub btnAllReset_Click()
    Dim userAns As VbMsgBoxResult
    
    ' =========================================================================
    ' 1. 誤クリックを徹底防止する最終警告（DefaultButton2 で「いいえ」にフォーカス）
    ' =========================================================================
    ' ★修正：警告文に「モバイルメモ」を追加
    userAns = MsgBox("【絶対警告】すべての設定、登録したLoRA、お気に入り、プリセット、モバイルメモを完全に消去し、初期状態にリセットしますか？" & vbCrLf & _
                     "※この操作は取り消せません（Undo不可、データは物理的に上書きされます）。" & vbCrLf & vbCrLf & _
                     "[CRITICAL WARNING] Are you absolutely sure you want to completely erase all configuration, registered LoRAs, Favorites, presets, and mobile memos, and reset them to the factory default?" & vbCrLf & _
                     "*This action is destructive and CANNOT be undone.", _
                     vbYesNo + vbCritical + vbDefaultButton2, APP_NAME)
                     
    If userAns <> vbYes Then
        ' 命拾いした時のサイレント離脱
        Application.StatusBar = "● Reset canceled. Your custom Grimoire is safe."
        Exit Sub
    End If
    
    ' =========================================================================
    ' 2. メモリ（Dictionary）の完全パージ ＆ デフォルト構造の再展開
    ' =========================================================================
    Application.StatusBar = "● Purging memory and restoring default config structure..."
    
    ' 既存のデータをメモリから完全に解放
    Set ConfigData = Nothing
    Set ConfigData = CreateObject("Scripting.Dictionary")
    
    ' ご指定のデフォルト構造構築プロシージャを呼び出す
    Call BuildDefaultConfigStructure
    
    ' =========================================================================
    ' ★追加：モバイルメモ用のツリーを初期状態（空）で確実に追加・再構築する
    ' =========================================================================
    If ConfigData.Exists("Mobile_Memo_Stock") Then
        ConfigData("Mobile_Memo_Stock") = ""
    Else
        ConfigData.Add "Mobile_Memo_Stock", ""
    End If
    
    If ConfigData.Exists("Mobile_URL_Stock") Then
        Set ConfigData("Mobile_URL_Stock") = New Collection
    Else
        ConfigData.Add "Mobile_URL_Stock", New Collection
    End If
    
    ' =========================================================================
    ' 3. 鉄壁の保存マクロ（SaveConfigJSON）を呼び出してJSONファイルへ即時同期
    ' =========================================================================
    ' 先ほど実装した「.bak」バックアップ＆ロールバック機能付きの安全保存が走ります
    Call SaveConfigJSON
    ' ★追加：マネージャーが開かれていれば空の状態で即時同期させる
    Call SyncFavoritesManager
    
    ' =========================================================================
    ' 4. UIの安全なリフレッシュ（不整合防止のためのUX案内）
    ' =========================================================================
    'IsDirty = False
    Application.StatusBar = "● System configuration successfully reset to factory default."
    
    ' 日英併記でリセット完了を通知し、フォームの再起動を優しく促す
    ' ★修正：完了メッセージにもモバイルメモが含まれるニュアンスを追加
    MsgBox "設定ファイルとモバイルメモを初期状態に完全リセットしました！" & vbCrLf & _
           "メモリとJSONファイルの同期は完了しています。" & vbCrLf & _
           "画面表示を完全にクリーンアップするため、一度このメインウィンドウを閉じ、開き直すことをお勧めします。" & vbCrLf & vbCrLf & _
           "System configuration and mobile memos have been completely reset!" & vbCrLf & _
           "Memory and JSON files are fully synchronized." & vbCrLf & _
           "To completely clean up and refresh the UI display, it is highly recommended to close and reopen this main window.", _
           vbInformation, APP_NAME
           
    ' ユーザーの利便性を考慮し、自動で一度画面を閉じてあげる親切設計
    Unload Me
End Sub


'-------------------------------------------------------------------------
' [Back to Legend] ボタン（凡例への即時帰還）
'-------------------------------------------------------------------------
Private Sub btnBackToLegend_Click()
    Dim wsName As String: wsName = "KENZEN SeaArt Helper"
    Dim targetRange As String: targetRange = "Legend"
    
    On Error Resume Next
    
    ' 1. 対象のシートを最前面に出す
    Sheets(wsName).Select
    
    ' 2. ネームドセル「Legend」へワープ
    ' Scroll:=True を指定することで、凡例が画面の左上（定位置）に来るようにします
    Application.GoTo Reference:=targetRange, Scroll:=True
    
    ' エラー処理（ネームドセルが消えている等の場合）
    If Err.Number <> 0 Then
        MsgBox "凡例（Legend）が見つかりません。名前定義を確認してください。" & vbCrLf & _
               "Named range 'Legend' not found.", vbCritical, APP_NAME
    End If
    
    On Error GoTo 0
    
    ' 3. ユーザーの操作主体をユーザーフォーム（コックピット）に戻す
    ' これにより、シートが切り替わってもすぐに次の操作が可能になります
    Me.btnCopyNormal.SetFocus
    
    ' 4. （親切設計）カテゴリージャンプの表示を初期値に戻す
    Me.cmbCategoryJump.ListIndex = 0
    
    ' ステータスバーへの通知
    Application.StatusBar = "● Returned to Legend."
End Sub

'-------------------------------------------------------------------------
' [Call Preset] ボタン（JSON対応版：選択したプリセット内容を完全復元・欠落検知＆自動クレンジング版）
'-------------------------------------------------------------------------
Private Sub btnCallLoRAPreset_Click()
    Dim presetName As String
    Dim foundCount As Integer
    Dim missingCount As Integer ' 削除されて見つからなかった件数
    Dim i As Integer ' リスト選択ループ用の変数
    
    ' プリセットデータ内の各項目
    Dim sysHash As String, strengthVal As String, triggers As String, weightVal As String
    Dim aliasName As String, displayStr As String
    
    Dim presetsDict As Object
    Dim presetItems As Collection
    Dim presetItem As Object
    Dim loraList As Object
    Dim loraItem As Object
    Dim foundAlias As String
    
    ' ★追加：生き残った正常なLoRAデータだけを一時退避するコレクション
    Dim validPresetItems As Collection
    
    ' 1. バリデーション：選択チェック
    If Me.cmbPresetList.ListIndex = -1 Then
        MsgBox "呼び出すプリセットを選択してください。" & vbCrLf & _
               "Please select a preset to load.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    presetName = Me.cmbPresetList.Value
    
    ' 2. リストとプレビューを一度「更地」にする
    Me.lstSelectedLoRA.Clear
    Me.txtLoRAPreview.text = ""
    foundCount = 0
    missingCount = 0 ' 初期化
    
    ' 3. JSONから対象のプリセットデータを取得
    If ConfigData Is Nothing Then Exit Sub
    If Not ConfigData.Exists("LoRA_Presets") Then Exit Sub
    
    Set presetsDict = ConfigData("LoRA_Presets")
    If Not presetsDict.Exists(presetName) Then Exit Sub
    
    Set presetItems = presetsDict(presetName)
    Set validPresetItems = New Collection ' ★追加：クレンジング用の空箱を準備
    
    ' メモリ上の大元LoRAデータも取得しておく（存在チェック用）
    If ConfigData.Exists("LoRA_List") Then
        Set loraList = ConfigData("LoRA_List")
    End If
    
    ' 4. プリセット内の各アイテムを展開して存在チェック
    For Each presetItem In presetItems
        sysHash = presetItem("Hash")
        strengthVal = presetItem("Strength")
        triggers = presetItem("Trigger")
        weightVal = presetItem("TriggerWeight")
        
        ' --- SystemHash から Alias を逆引きし、大元に存在するかチェック ---
        foundAlias = ""
        If Not loraList Is Nothing Then
            For Each loraItem In loraList
                If loraItem("Hash") = sysHash Then
                    foundAlias = loraItem("Alias")
                    Exit For
                End If
            Next loraItem
        End If
        
        If foundAlias <> "" Then
            ' 【正常ルート】
            displayStr = foundAlias & " / " & strengthVal
            
            With Me.lstSelectedLoRA
                .AddItem displayStr             ' 1列目: 表示
                .List(.ListCount - 1, 1) = triggers    ' 2列目: トリガー(隠し)
                .List(.ListCount - 1, 2) = foundAlias  ' 3列目: Alias(隠し)
                .List(.ListCount - 1, 3) = strengthVal ' 4列目: Strength(隠し)
            End With
            
            ' 重みコンボボックスの同期（最初の1件目の値で上書き）
            If foundCount = 0 Then
                Me.cmbTriggerWeight.Value = weightVal
            End If
            
            ' ★追加：正常に読み込めたアイテムだけを新しい箱に退避する
            validPresetItems.Add presetItem
            
            foundCount = foundCount + 1
        Else
            ' 【異常ルート】
            missingCount = missingCount + 1
        End If
    Next presetItem
    
' 5. 仕上げと通知
    If foundCount > 0 Then
        
        ' ========================================================
        ' ★修正：btnWrapLoRA_Clickを呼ばず、リストが記憶しているトリガー状態をそのまま結合する
        ' これにより個別のトリガー重み（1.3など）が破壊されるのを防ぐ
        ' ========================================================
        Dim rebuildTriggers As String
        Dim rebuildLoRAs As String
        Dim rAlias As String, rStrength As String, rTrigger As String
        Dim rHash As String, rModelName As String, rTarget As String
        Dim rItem As Object
        
        rebuildTriggers = ""
        rebuildLoRAs = ""
        
        ' ========================================================
        ' ★追加：呼び出したプリセットが「ModelName」で保存されていたか判定し、
        ' UIのラジオボタン（opbWrapName / opbWrapHash）の表示を同期させる
        ' ========================================================
        If validPresetItems.Count > 0 Then
            Dim checkItem As Object
            Set checkItem = validPresetItems(1)
            If checkItem.Exists("TargetName") And checkItem.Exists("ModelName") Then
                ' TargetNameとModelNameが一致していれば、Nameで保存されたと判定してUIを切り替え
                If Trim(checkItem("TargetName")) <> "" And Trim(checkItem("TargetName")) = Trim(checkItem("ModelName")) Then
                    Me.opbWrapName.Value = True
                Else
                    Me.opbWrapHash.Value = True
                End If
            End If
        End If
        
        For i = 0 To Me.lstSelectedLoRA.ListCount - 1
            ' 次の操作のために選択状態は解除しておく（UX向上）
            Me.lstSelectedLoRA.Selected(i) = False
            
            rAlias = Me.lstSelectedLoRA.List(i, 2)
            rStrength = Me.lstSelectedLoRA.List(i, 3)
            rTrigger = Trim(Me.lstSelectedLoRA.List(i, 1))
            rHash = ""
            rModelName = ""
            
            ' メモリ上のLoRA_Listから該当データを検索
            If Not loraList Is Nothing Then
                For Each rItem In loraList
                    If rItem("Alias") = rAlias Then
                        rHash = rItem("Hash")
                        If rItem.Exists("ModelName") Then rModelName = rItem("ModelName")
                        Exit For
                    End If
                Next rItem
            End If
            
            ' ========================================================
            ' ★修正：現在のUIオプションボタンに依存せず、
            ' JSONに保存された「TargetName」を最優先で復元する
            ' ========================================================
            Dim vItem As Object
            Set vItem = validPresetItems(i + 1) ' Collectionは1始まり
            
            rTarget = ""
            If vItem.Exists("TargetName") Then
                rTarget = Trim(vItem("TargetName"))
            End If
            
            ' 旧バージョンのデータ等で TargetName が無い場合のフォールバック
            If rTarget = "" Then
                If Me.opbWrapName.Value = True And Trim(rModelName) <> "" Then
                    rTarget = Trim(rModelName)
                Else
                    If Trim(rHash) <> "" Then rTarget = Trim(rHash) Else rTarget = Trim(rAlias)
                End If
            End If
            
            If rTarget <> "" Then
                ' LoRAタグの結合（重複防止）
                Dim curLoRA As String
                curLoRA = "<lora:" & rTarget & ":" & rStrength & ">"
                If InStr(1, rebuildLoRAs, curLoRA, vbTextCompare) = 0 Then
                    If rebuildLoRAs = "" Then rebuildLoRAs = curLoRA Else rebuildLoRAs = rebuildLoRAs & " " & curLoRA
                End If
                
                ' トリガーワードの結合（カンマ区切りで重複防止）
                If rTrigger <> "" Then
                    Dim tArr() As String
                    Dim t As Integer
                    tArr = Split(rTrigger, ",")
                    For t = 0 To UBound(tArr)
                        Dim singleTrigger As String
                        singleTrigger = Trim(tArr(t))
                        If singleTrigger <> "" Then
                            ' 重複チェック
                            If InStr(1, ", " & rebuildTriggers & ",", ", " & singleTrigger & ",", vbTextCompare) = 0 Then
                                If rebuildTriggers = "" Then
                                    rebuildTriggers = singleTrigger
                                Else
                                    rebuildTriggers = rebuildTriggers & ", " & singleTrigger
                                End If
                            End If
                        End If
                    Next t
                End If
            End If
        Next i
        
        ' プレビュー欄へ最終出力
        
        ' プレビュー欄へ最終出力
        If rebuildTriggers <> "" And rebuildLoRAs <> "" Then
            Me.txtLoRAPreview.text = rebuildTriggers & ", " & rebuildLoRAs
        ElseIf rebuildTriggers <> "" Then
            Me.txtLoRAPreview.text = rebuildTriggers
        ElseIf rebuildLoRAs <> "" Then
            Me.txtLoRAPreview.text = rebuildLoRAs
        End If
        
        ' UI状態の更新
        Call UpdateLoRAUIState
        Application.StatusBar = "● Preset '" & presetName & "' loaded."
        
        ' ========================================================
        ' 一部欠落があった場合の自動クレンジング処理
        ' ========================================================
        If missingCount > 0 Then
            ' 元のプリセットの中身を、生き残った正常なアイテムのみで上書きし、即時保存する
            Set presetsDict(presetName) = validPresetItems
            Call SaveConfigJSON
            
            ' メッセージに自己修復が完了した旨を追記
            MsgBox "プリセットの一部（" & missingCount & "件）は、大元のLoRAデータが削除されているため除外されました。" & vbCrLf & _
                   "※このプリセットは、無効なLoRAを自動的に整理して上書き保存されました。" & vbCrLf & vbCrLf & _
                   "Some LoRAs (" & missingCount & " items) were excluded because their original data was deleted." & vbCrLf & _
                   "*The preset has been automatically cleaned and saved.", _
                   vbExclamation, APP_NAME
        End If
    Else
        ' ========================================================
        ' 1件も読み込めなかった場合（全滅ルート・既存のまま）
        ' ========================================================
        If missingCount > 0 Then
            MsgBox "このプリセットに含まれる全てのLoRAデータが削除されているため、呼び出しをキャンセルしました。" & vbCrLf & _
                   "この無効なプリセットを自動的に削除し、リストから整理します。" & vbCrLf & vbCrLf & _
                   "Load cancelled. All LoRAs in this preset have been deleted. Removing this invalid preset.", _
                   vbCritical, APP_NAME
            
            ' メモリからプリセット名（キー）を指定して一撃で削除し、物理保存
            presetsDict.Remove presetName
            Call SaveConfigJSON
            
            ' コンボボックスの表示を更新
            Call RefreshPresetList
            Me.cmbPresetList.Value = ""
            
        Else
            MsgBox "プリセットデータが見つかりませんでした。" & vbCrLf & _
                   "Preset data not found.", vbCritical, APP_NAME
        End If
    End If
End Sub

' =========================================================================
' [UI制御] プロンプトのクリーンアップ (ゴミカンマの除去とLoRAタグの保護)
' =========================================================================
Private Sub btnCleanUpPrompt_Click()
    Dim currentText As String
    Dim regEx As Object
    
    currentText = Me.txtMain.Value
    
    ' 空の場合は何もしない
    If Trim(currentText) = "" Then Exit Sub
    
    ' 1. 変更前の状態を履歴に保存 (Undo対応：ユーザーがいつでも戻せるようにする)
    Call SaveHistory(currentText)
    
    ' =========================================================
    ' ★正規表現 (VBScript.RegExp) を用いた高度なクリーニング
    ' =========================================================
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.Multiline = True ' 「^ (行頭)」と「$ (行末)」を各行ごとに判定させるため必須
    
    ' (1) 連続するカンマや、スペースを挟んだカンマ (,, や , , ) を1つの「, 」に統合
    regEx.Pattern = "(\s*,\s*){2,}"
    currentText = regEx.Replace(currentText, ", ")
    
    ' (2) 各行の「行頭」にあるカンマを削除 (例: 編集で先頭が ", 1girl" になった場合)
    regEx.Pattern = "^\s*,\s*"
    currentText = regEx.Replace(currentText, "")
    
    ' (3) 各行の「行末」にあるカンマを削除 (例: "1girl," で改行されている場合)
    regEx.Pattern = "\s*,\s*$"
    currentText = regEx.Replace(currentText, "")
    
    ' (4) ★最重要要件：<lora:...> などの「<」の直前にあるカンマを削除し、半角スペースに置換
    ' これにより "1girl, <lora:xxx:1.0>" が美しい "1girl <lora:xxx:1.0>" に変換されます
    ' （複数のLoRAが ", " で繋がっていた場合も " <lora..." に一掃されます）
    regEx.Pattern = "\s*,\s*(<)"
    currentText = regEx.Replace(currentText, " $1")
    
    ' (5) 念のため、「BREAK」の直前にカンマが残っていたらそれも削除
    regEx.Pattern = "\s*,\s*(BREAK)"
    currentText = regEx.Replace(currentText, vbCrLf & "$1")
    
    ' =========================================================
    ' ★テキストボックスとシステム変数の同期
    ' =========================================================
    IsUpdatingBySystem = True
    
    Me.txtMain.Value = currentText
    AccumulatedPrompt = currentText
    Call SetClipboardText(currentText) ' 先ほどリトライ機能を付けた安全なクリップボード関数
    'IsDirty = True
    
    IsUpdatingBySystem = False
    
    ' UI状態の更新 (ウェイト調整ボタンの有効化など)
    Me.chkWeight.Enabled = (Trim(Me.txtMain.Value) <> "")
    
    On Error Resume Next
    'Call Me.UpdateDoneButtonState
    On Error GoTo 0
    
    ' クリーンアップ完了をステータスバーで通知
    Application.StatusBar = "Prompt Cleaned Up!"
End Sub

Private Sub btnClearAllMobiList_Click()
' =========================================================================
' [Mobile Utility] リストとメモを全消去し、母艦コンフィグもリセット(空セーブ)する
' =========================================================================
    
    ' ★追加：すでにリストとメモが空の場合は警告を出して処理を中断する
    If Me.lstMobileMemo.ListCount = 0 And Trim(Me.txtMobileMemo.text) = "" Then
        MsgBox "消去するモバイルデータがありません（既に空です）。" & vbCrLf & _
               "There is no data to clear (already empty).", _
               vbInformation, APP_NAME
        Exit Sub
    End If

    ' 1. 誤操作防止の厳重な確認ダイアログ（vbCriticalで警告音を鳴らす）
    If MsgBox("リストとメモをすべて消去し、マクロ本体のコンフィグ(保存データ)も空にします。" & vbCrLf & _
              "本当によろしいですか？ (Are you sure you want to clear all?)", _
              vbYesNo + vbCritical, APP_NAME) = vbNo Then
        Exit Sub
    End If

    ' 2. UI上のデータをクリア
    Me.lstMobileMemo.Clear
    Me.txtMobileMemo.text = ""

    ' 3. 母艦(ConfigData)のストックを空にリセットする
    If Not ConfigData Is Nothing Then
        ' ① フリーメモのストックを空文字で上書き
        If ConfigData.Exists("Mobile_Memo_Stock") Then
            ConfigData("Mobile_Memo_Stock") = ""
        End If
        
        ' ② URLリストのストックを新しい空のコレクションで上書き
        If ConfigData.Exists("Mobile_URL_Stock") Then
            Set ConfigData("Mobile_URL_Stock") = New Collection
        End If
        
        ' 4. コンフィグJSONへ物理書き込み（空の状態をセーブ）
        Call SaveConfigJSON
    End If

    ' 5. 完了メッセージ
    MsgBox "すべてのモバイルデータを消去し、コンフィグをリセットしました！" & vbCrLf & _
           "All data cleared and config reset successfully!", vbInformation, APP_NAME
End Sub
' =========================================================================
' [Mobile Utility] 編集したメモ・リストをコンフィグ(母艦)へ保存する
' =========================================================================
Private Sub btnSaveMemoToConfig_Click()
    Dim i As Long
    Dim newUrlStock As Collection
    
    If Me.lstMobileMemo.ListCount = 0 And Trim(Me.txtMobileMemo.text) = "" Then
        MsgBox "保存するメモやリストがありません（空です）。" & vbCrLf & _
               "There is no data to save (empty).", _
               vbInformation, APP_NAME
        Exit Sub
    End If
    
    If ConfigData Is Nothing Then
        MsgBox "コンフィグデータが初期化されていません。" & vbCrLf & _
               "Config data is not initialized.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' フリーメモ (txtMobileMemo) をストックに反映
    If ConfigData.Exists("Mobile_Memo_Stock") Then
        ConfigData("Mobile_Memo_Stock") = Me.txtMobileMemo.text
    Else
        ConfigData.Add "Mobile_Memo_Stock", Me.txtMobileMemo.text
    End If
    
    ' リストボックス (lstMobileMemo) の内容を再構築してストックに反映
    Set newUrlStock = New Collection
    
    For i = 0 To Me.lstMobileMemo.ListCount - 1
        ' ★修正：Null対策として明示的に空文字を結合してからデリミタで繋ぐ
        newUrlStock.Add (Me.lstMobileMemo.List(i, 0) & "") & "<|>" & (Me.lstMobileMemo.List(i, 1) & "")
    Next i
    
    If ConfigData.Exists("Mobile_URL_Stock") Then
        Set ConfigData("Mobile_URL_Stock") = newUrlStock
    Else
        ConfigData.Add "Mobile_URL_Stock", newUrlStock
    End If
    
    Call SaveConfigJSON
    
    MsgBox "現在のメモとリストの状態をコンフィグに保存しました！" & vbCrLf & _
           "Memos and lists saved to config successfully!", vbInformation, APP_NAME
End Sub

' =========================================================================
' [Mobile Utility] Deleteキーで選択中のアイテムを削除する
' =========================================================================
Private Sub lstMobileMemo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Deleteキー（KeyCode = 46 または vbKeyDelete）が押されたかを判定
    If KeyCode = vbKeyDelete Then
        ' リストのアイテムが選択されている場合のみ実行
        If Me.lstMobileMemo.ListIndex <> -1 Then
            ' 既に実装済みの個別削除プロシージャを呼び出して処理を委譲する
            Call btnRemoveMobiItem_Click
        End If
    End If
End Sub
'-------------------------------------------------------------------------
' [Clear] ボタン：消去前Undoログ記憶（SaveFavHistory連動・完全版）
'-------------------------------------------------------------------------
Private Sub btnClearBoth_Click()
    ' 1. 空振り防衛：両方のテキストボックスが元から空なら、履歴保存も処理もせず安全に抜ける
    If Trim(Me.txtFav.Value) = "" And Trim(Me.txtDescription.Value) = "" Then
        Exit Sub
    End If
    
    ' 2. 専用の履歴保存関数を呼び出し、現在のプロンプトと説明文をセットで同時に記憶
    Call SaveFavHistory(Me.txtFav.Value, Me.txtDescription.Value)
    
    ' 3. 入力欄を完全に更地（クリア）にする
    Me.txtFav.Value = ""
    Me.txtDescription.Value = ""
    
    ' 4. 変更があったことをシステムに通知（必要に応じてUI更新処理等もここへ追記してください）
    'IsDirty = True
End Sub

'-------------------------------------------------------------------------
' [Clear Preview] ボタン：プレビュー欄のみをスマートに更地にする（UI同期連動版）
'-------------------------------------------------------------------------
Private Sub btnClearLoRAPreview_Click()
    If Trim(Me.txtLoRAPreview.text) = "" Then Exit Sub
    
    Me.txtLoRAPreview.text = ""
    
    ' カートの選択（ハイライト）状態をすべて解除
    Dim i As Integer
    For i = 0 To Me.lstSelectedLoRA.ListCount - 1
        Me.lstSelectedLoRA.Selected(i) = False
    Next i
    
    ' ★追加：選択解除をUIに即座に通知し、btnGetLoRANegativeを即時オフにする
    Call lstSelectedLoRA_Change
    
    Application.StatusBar = "● LoRA Preview and selection cleared."
End Sub


'-------------------------------------------------------------------------
' [Delete Preset] ボタン（JSON対応版：選択したプリセットを完全に削除）
'-------------------------------------------------------------------------
Private Sub btnDeleteLoRAPreset_Click()
    Dim presetName As String
    Dim confirm As VbMsgBoxResult
    Dim presetsDict As Object
    
    ' 1. バリデーション：選択チェック
    If Me.cmbPresetList.ListIndex = -1 Then
        MsgBox "削除するプリセットを選択してください。" & vbCrLf & _
               "Please select a preset to delete.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    presetName = Me.cmbPresetList.Value
    
    ' 2. ユーザーへの最終確認（誤爆防止）
    confirm = MsgBox("プリセット '" & presetName & "' を削除してもよろしいですか？" & vbCrLf & _
                     "Are you sure you want to delete the preset '" & presetName & "'?" & vbCrLf & vbCrLf & _
                     "この操作は元に戻せません。" & vbCrLf & "This action cannot be undone.", _
                     vbQuestion + vbYesNo + vbDefaultButton2, APP_NAME)
    
    If confirm = vbNo Then Exit Sub
    
    ' 3. JSONから該当データを一撃で削除
    If ConfigData Is Nothing Then Exit Sub
    If Not ConfigData.Exists("LoRA_Presets") Then Exit Sub
    
    Set presetsDict = ConfigData("LoRA_Presets")
    
    If presetsDict.Exists(presetName) Then
        ' ★シートを下からループする泥臭い処理が、この1行の魔法に変わります
        presetsDict.Remove presetName
        
        ' 物理ファイルへ即時保存
        Call SaveConfigJSON
    End If
    
    ' 4. 後処理：UIの更新
    Call RefreshPresetList
    
    ' （※ポジティブ/ネガティブのUXに合わせ、コンボボックスが空でなければ先頭に戻す）
    If Me.cmbPresetList.ListCount > 0 Then
        Me.cmbPresetList.ListIndex = 0
    Else
        Me.cmbPresetList.Value = ""
    End If
    
    MsgBox "プリセット '" & presetName & "' を削除しました。" & vbCrLf & _
           "Preset '" & presetName & "' has been deleted.", vbInformation, APP_NAME
    
End Sub

' =========================================================================
' [Export] 選択されたデータをJSONファイルとして書き出す（クローンサニタイズ完全版）
' =========================================================================
Private Sub btnDoExport_Click()
    Dim exportDict As Object
    Dim filePath As Variant
    Dim jsonStr As String
    Dim ado As Object
    Dim hasDataToExport As Boolean
    
    ' 1. チェックボックスが1つ以上選択されているか確認 (★chkExportMobiMemo を追加)
    hasDataToExport = Me.chkExportPosiPresets.Value Or Me.chkExportNegaPresets.Value Or _
                      Me.chkExportLoRA.Value Or Me.chkExportLoRAPresets.Value Or Me.chkExportFav.Value Or _
                      Me.chkExportNegaStock.Value Or Me.chkExportMobiMemo.Value
                      
    If Not hasDataToExport Then
        MsgBox "エクスポートする項目を1つ以上選択してください。" & vbCrLf & _
               "Please select at least one item to export.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If ConfigData Is Nothing Then Exit Sub
    
    ' =========================================================
    ' 2. 出力するデータがあるかチェックして抽出する
    ' =========================================================
    Set exportDict = CreateObject("Scripting.Dictionary")
    
    If Me.chkExportPosiPresets.Value And ConfigData.Exists("Positive_Presets") Then
        If ConfigData("Positive_Presets").Count > 0 Then Set exportDict("Positive_Presets") = ConfigData("Positive_Presets")
    End If
    
    If Me.chkExportNegaPresets.Value And ConfigData.Exists("Negative_Presets") Then
        If ConfigData("Negative_Presets").Count > 0 Then Set exportDict("Negative_Presets") = ConfigData("Negative_Presets")
    End If
    
    If Me.chkExportLoRA.Value And ConfigData.Exists("LoRA_List") Then
        If ConfigData("LoRA_List").Count > 0 Then Set exportDict("LoRA_List") = ConfigData("LoRA_List")
    End If
    
    If Me.chkExportLoRAPresets.Value And ConfigData.Exists("LoRA_Presets") Then
        If ConfigData("LoRA_Presets").Count > 0 Then Set exportDict("LoRA_Presets") = ConfigData("LoRA_Presets")
    End If
    
    If Me.chkExportFav.Value And ConfigData.Exists("Fav_List") Then
        If ConfigData("Fav_List").Count > 0 Then Set exportDict("Fav_List") = ConfigData("Fav_List")
    End If

    If Me.chkExportNegaStock.Value And ConfigData.Exists("NegativeStock") Then
        If Trim(ConfigData("NegativeStock")) <> "" Then
            exportDict("NegativeStock") = ConfigData("NegativeStock")
        End If
    End If
    
    ' ★追加：モバイルメモ ＆ URLリストのエクスポート抽出
    If Me.chkExportMobiMemo.Value Then
        If ConfigData.Exists("Mobile_Memo_Stock") Then
            If Trim(ConfigData("Mobile_Memo_Stock")) <> "" Then
                exportDict("Mobile_Memo_Stock") = ConfigData("Mobile_Memo_Stock")
            End If
        End If
        If ConfigData.Exists("Mobile_URL_Stock") Then
            If ConfigData("Mobile_URL_Stock").Count > 0 Then
                Set exportDict("Mobile_URL_Stock") = ConfigData("Mobile_URL_Stock")
            End If
        End If
    End If
    
    If exportDict.Count = 0 Then
        MsgBox "選択された項目には、現在エクスポートできるデータがありません。" & vbCrLf & _
               "No data found for the selected items.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' =========================================================
    ' 3. エクスポート先を指定するダイアログを表示
    ' =========================================================
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="KENZEN_Backup_" & Format(Now, "yyyymmdd_HHmmss") & ".json", _
        FileFilter:="JSON Files (*.json), *.json", _
        title:="データをエクスポート (Export Data)")
        
    If filePath = False Then Exit Sub
    
    Dim systemConfigPath As String
    systemConfigPath = ThisWorkbook.Path & "\KENZEN_Config.json"
    
    If StrComp(filePath, systemConfigPath, vbTextCompare) = 0 Then
        MsgBox "システムが使用中の本体設定ファイル (KENZEN_Config.json) を直接上書きすることはできません。" & vbCrLf & _
               "データが破損する恐れがあります。別のファイル名やフォルダを指定してください。" & vbCrLf & vbCrLf & _
               "Cannot overwrite the main system configuration file." & vbCrLf & _
               "Please choose a different file name or destination.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' =========================================================
    ' 4. ディープクローン生成 ＆ 鉄壁のサニタイズ処理
    ' =========================================================
    On Error GoTo ExportError
    
    Dim cloneDict As Object
    ' 一度シリアライズしてパースし直すことで、元のConfigDataから完全に独立した複製を作る（VBAシャドウイング回避）
    Set cloneDict = JsonConverter.ParseJson(JsonConverter.ConvertToJson(exportDict))
    
    ' クローン側に対してダブルクォーテーションの無害化を一斉適用（LoRA固有ネガティブを完全救済）
    Call SanitizeQuotesInDictionary(cloneDict)
    
    ' 安全な状態になったクローンを綺麗なインデントでJSON化
    jsonStr = JsonConverter.ConvertToJson(cloneDict, 4)
    
    ' ADODB.StreamによるBOMなしUTF-8書き出し
    Set ado = CreateObject("ADODB.Stream")
    With ado
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText jsonStr
        .SaveToFile filePath, 2 ' adSaveCreateOverWrite
        .Close
    End With
    On Error GoTo 0
    
    ' 日英併記で成功メッセージを出力
    MsgBox "エクスポートが完了しました！" & vbCrLf & _
           "Export complete!", vbInformation, APP_NAME
    Exit Sub
    
ExportError:
    On Error Resume Next
    If Not ado Is Nothing Then ado.Close
    On Error GoTo 0
    
    ' 日英併記でエラーメッセージを出力
    MsgBox "エクスポート中にエラーが発生しました。" & vbCrLf & _
           "Error generating export file.", vbCritical, APP_NAME
End Sub
' =========================================================================
' [Import] 選択されたJSONファイルを読み込み、システムに統合または上書きする（鉄壁互換性版）
' =========================================================================
Private Sub btnDoImport_Click()
    Dim filePath As Variant
    Dim jsonStr As String
    Dim ado As Object
    Dim importDict As Object
    Dim hasDataToImport As Boolean
    Dim modeIsMerge As Boolean
        
    Dim cntPosi As Long, cntNega As Long, cntLoRAPreset As Long
    Dim cntLoRA As Long, cntFav As Long
    Dim cntNegaStock As Long
    Dim cntMobiUrl As Long ' ★追加：モバイルURLのカウント用
    Dim mobiMemoUpdated As Boolean ' ★追加：フリーメモが更新されたかのフラグ
    Dim reportMsg As String
    
    Dim overflowFavCol As Object
    Set overflowFavCol = New Collection
    
    ' 1. チェックボックスとモードの確認 (★chkImportMobiMemo を追加)
    hasDataToImport = Me.chkImportPosiPresets.Value Or Me.chkImportNegaPresets.Value Or _
                      Me.chkImportLoRA.Value Or Me.chkImportLoRAPresets.Value Or Me.chkImportFav.Value Or _
                      Me.chkImportNegaStock.Value Or Me.chkImportMobiMemo.Value
                      
    If Not hasDataToImport Then
        MsgBox "インポートする項目を1つ以上選択してください。" & vbCrLf & _
               "Please select at least one item to import.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Me.opbMerge.Value = False And Me.opbOverwrite.Value = False Then
        MsgBox "インポートの方式（Merge または Replace All）を選択してください。" & vbCrLf & _
               "Please select an import mode.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    modeIsMerge = Me.opbMerge.Value
    
    ' 2. インポートするファイルを選択
    filePath = Application.GetOpenFilename( _
        FileFilter:="JSON Files (*.json), *.json", _
        title:="インポートするファイルを選択 (Select file to import)")
        
    If filePath = False Then Exit Sub
    
    Dim systemConfigPath As String
    systemConfigPath = ThisWorkbook.Path & "\KENZEN_Config.json"
    If StrComp(filePath, systemConfigPath, vbTextCompare) = 0 Then
        MsgBox "現在稼働中の本体ファイル (KENZEN_Config.json) はインポートできません。" & vbCrLf & _
               "バックアップとして出力した別のファイルを選択してください。", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' 3. JSONファイルの読み込み
    On Error GoTo ReadError
    Set ado = CreateObject("ADODB.Stream")
    With ado
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        jsonStr = .ReadText
        .Close
    End With
    
    Set importDict = JsonConverter.ParseJson(jsonStr)
    On Error GoTo 0
    
    If ConfigData Is Nothing Then Call InitConfigJSON
    
    ' =========================================================
    ' 4. データのマージ or 上書き処理（厳格なカウント付き）
    ' =========================================================
    Dim key As Variant
    Dim impItem As Object, existItem As Object
    Dim isDup As Boolean
    Dim impStr As Variant, existStr As Variant ' ★文字列ループ用
    
    ' --- Positive Presets ---
    If Me.chkImportPosiPresets.Value And importDict.Exists("Positive_Presets") Then
        If modeIsMerge Then
            If Not ConfigData.Exists("Positive_Presets") Then ConfigData.Add "Positive_Presets", CreateObject("Scripting.Dictionary")
            cntPosi = 0
            For Each key In importDict("Positive_Presets").Keys
                If Not ConfigData("Positive_Presets").Exists(key) Then
                    cntPosi = cntPosi + 1
                ElseIf ConfigData("Positive_Presets")(key) <> importDict("Positive_Presets")(key) Then
                    cntPosi = cntPosi + 1
                End If
                ConfigData("Positive_Presets")(key) = importDict("Positive_Presets")(key)
            Next key
        Else
            Set ConfigData("Positive_Presets") = importDict("Positive_Presets")
            cntPosi = importDict("Positive_Presets").Count
        End If
    End If
    
    ' --- Negative Presets ---
    If Me.chkImportNegaPresets.Value And importDict.Exists("Negative_Presets") Then
        If modeIsMerge Then
            If Not ConfigData.Exists("Negative_Presets") Then ConfigData.Add "Negative_Presets", CreateObject("Scripting.Dictionary")
            cntNega = 0
            For Each key In importDict("Negative_Presets").Keys
                If Not ConfigData("Negative_Presets").Exists(key) Then
                    cntNega = cntNega + 1
                ElseIf ConfigData("Negative_Presets")(key) <> importDict("Negative_Presets")(key) Then
                    cntNega = cntNega + 1
                End If
                ConfigData("Negative_Presets")(key) = importDict("Negative_Presets")(key)
            Next key
        Else
            Set ConfigData("Negative_Presets") = importDict("Negative_Presets")
            cntNega = importDict("Negative_Presets").Count
        End If
    End If
    
    ' --- LoRA Presets ---
    If Me.chkImportLoRAPresets.Value And importDict.Exists("LoRA_Presets") Then
        If modeIsMerge Then
            If Not ConfigData.Exists("LoRA_Presets") Then ConfigData.Add "LoRA_Presets", CreateObject("Scripting.Dictionary")
            cntLoRAPreset = 0
            For Each key In importDict("LoRA_Presets").Keys
                If Not ConfigData("LoRA_Presets").Exists(key) Then cntLoRAPreset = cntLoRAPreset + 1
                Set ConfigData("LoRA_Presets")(key) = importDict("LoRA_Presets")(key)
            Next key
        Else
            Set ConfigData("LoRA_Presets") = importDict("LoRA_Presets")
            cntLoRAPreset = importDict("LoRA_Presets").Count
        End If
    End If
    
    ' --- LoRA Base Data ---
    If Me.chkImportLoRA.Value And importDict.Exists("LoRA_List") Then
        If modeIsMerge Then
            If Not ConfigData.Exists("LoRA_List") Then ConfigData.Add "LoRA_List", New Collection
            cntLoRA = 0
            For Each impItem In importDict("LoRA_List")
                If Not impItem.Exists("Negative") Then impItem.Add "Negative", ""
                
                isDup = False
                For Each existItem In ConfigData("LoRA_List")
                    If Not existItem.Exists("Negative") Then existItem.Add "Negative", ""
                    
                    If existItem("Hash") = impItem("Hash") Then
                        isDup = True
                        If Trim(existItem("Negative")) <> Trim(impItem("Negative")) Then
                            existItem("Negative") = Trim(impItem("Negative"))
                            cntLoRA = cntLoRA + 1
                        End If
                        Exit For
                    End If
                Next existItem
                
                If Not isDup Then
                    ConfigData("LoRA_List").Add impItem
                    cntLoRA = cntLoRA + 1
                End If
            Next impItem
        Else
            Set ConfigData("LoRA_List") = importDict("LoRA_List")
            cntLoRA = importDict("LoRA_List").Count
            
            For Each impItem In ConfigData("LoRA_List")
                If Not impItem.Exists("Negative") Then impItem.Add "Negative", ""
            Next impItem
        End If
    End If
    
    ' --- Favorite ---
    If Me.chkImportFav.Value And importDict.Exists("Fav_List") Then
        If modeIsMerge Then
            If Not ConfigData.Exists("Fav_List") Then ConfigData.Add "Fav_List", New Collection
            cntFav = 0
            
            Dim cleanImpFav As String
            Dim cleanExistFav As String
            
            For Each impItem In importDict("Fav_List")
                isDup = False
                cleanImpFav = Replace(Replace(impItem("Prompt"), vbCr, ""), vbLf, "")
                
                For Each existItem In ConfigData("Fav_List")
                    cleanExistFav = Replace(Replace(existItem("Prompt"), vbCr, ""), vbLf, "")
                    If cleanExistFav = cleanImpFav Then
                        isDup = True
                        Exit For
                    End If
                Next existItem
                
                If Not isDup Then
                    If ConfigData("Fav_List").Count < 50 Then
                        ConfigData("Fav_List").Add impItem
                        cntFav = cntFav + 1
                    Else
                        overflowFavCol.Add impItem
                    End If
                End If
            Next impItem
        Else
            Set ConfigData("Fav_List") = New Collection
            cntFav = 0
            For Each impItem In importDict("Fav_List")
                If ConfigData("Fav_List").Count < 50 Then
                    ConfigData("Fav_List").Add impItem
                    cntFav = cntFav + 1
                Else
                    overflowFavCol.Add impItem
                End If
            Next impItem
        End If
        Call SyncFavSheetFromJSON
    End If
    
    ' --- Negative Stock ---
    If Me.chkImportNegaStock.Value And importDict.Exists("NegativeStock") Then
        If modeIsMerge Then
            Dim existingNegaStock As String, impNegaStock As String
            Dim tagArr() As String, t As Variant
            Dim tempDict As Object
            Set tempDict = CreateObject("Scripting.Dictionary")
            Dim beforeCount As Long
            
            If ConfigData.Exists("NegativeStock") Then
                existingNegaStock = ConfigData("NegativeStock")
                tagArr = Split(existingNegaStock, ",")
                For Each t In tagArr
                    If Trim(t) <> "" Then tempDict(Trim(t)) = True
                Next t
            End If
            
            beforeCount = tempDict.Count
            
            impNegaStock = importDict("NegativeStock")
            tagArr = Split(impNegaStock, ",")
            For Each t In tagArr
                If Trim(t) <> "" Then tempDict(Trim(t)) = True
            Next t
            
            cntNegaStock = tempDict.Count - beforeCount
            
            If tempDict.Count > 0 Then
                ConfigData("NegativeStock") = Join(tempDict.Keys, ", ")
            Else
                ConfigData("NegativeStock") = ""
            End If
            
        Else
            ConfigData("NegativeStock") = importDict("NegativeStock")
            
            cntNegaStock = 0
            impNegaStock = importDict("NegativeStock")
            tagArr = Split(impNegaStock, ",")
            For Each t In tagArr
                If Trim(t) <> "" Then cntNegaStock = cntNegaStock + 1
            Next t
        End If
    End If

    ' =========================================================
    ' ★追加：Mobile Memo & URL Stock のインポート
    ' =========================================================
    If Me.chkImportMobiMemo.Value Then
        ' 1. URLリストの処理
        If importDict.Exists("Mobile_URL_Stock") Then
            If modeIsMerge Then
                If Not ConfigData.Exists("Mobile_URL_Stock") Then ConfigData.Add "Mobile_URL_Stock", New Collection
                cntMobiUrl = 0
                For Each impStr In importDict("Mobile_URL_Stock")
                    isDup = False
                    For Each existStr In ConfigData("Mobile_URL_Stock")
                        If existStr = impStr Then
                            isDup = True
                            Exit For
                        End If
                    Next existStr
                    
                    If Not isDup Then
                        ConfigData("Mobile_URL_Stock").Add impStr
                        cntMobiUrl = cntMobiUrl + 1
                    End If
                Next impStr
            Else
                Set ConfigData("Mobile_URL_Stock") = importDict("Mobile_URL_Stock")
                cntMobiUrl = importDict("Mobile_URL_Stock").Count
            End If
        End If
        
        ' 2. フリーメモの処理
        If importDict.Exists("Mobile_Memo_Stock") Then
            If modeIsMerge Then
                If Not ConfigData.Exists("Mobile_Memo_Stock") Then ConfigData.Add "Mobile_Memo_Stock", ""
                
                If Trim(importDict("Mobile_Memo_Stock")) <> "" Then
                    ' 全く同じテキストが既に入っている場合は追記しない（無限増殖防止）
                    If InStr(ConfigData("Mobile_Memo_Stock"), importDict("Mobile_Memo_Stock")) = 0 Then
                        If ConfigData("Mobile_Memo_Stock") <> "" Then
                            ' 既存データがあれば、改行を2つ挟んで追記
                            ConfigData("Mobile_Memo_Stock") = ConfigData("Mobile_Memo_Stock") & vbCrLf & vbCrLf & importDict("Mobile_Memo_Stock")
                        Else
                            ConfigData("Mobile_Memo_Stock") = importDict("Mobile_Memo_Stock")
                        End If
                        mobiMemoUpdated = True
                    End If
                End If
            Else
                ConfigData("Mobile_Memo_Stock") = importDict("Mobile_Memo_Stock")
                If Trim(ConfigData("Mobile_Memo_Stock")) <> "" Then mobiMemoUpdated = True
            End If
        End If
    End If
    
    ' =========================================================
    ' 5. 最終保存とUI更新
    ' =========================================================
    Call SaveConfigJSON
    Call InitConfigJSON
    Call RefreshMainWindowLoRA
    Call InitNegativePromptPage
    Call InitPositiveUI
    
    Call RefreshMobileUI ' ★ここを追加！モバイル画面を即座に最新化

    If UserForms.Count > 0 Then
        On Error Resume Next
        Call FavoritesManager.RefreshFavList
        On Error GoTo 0
    End If

    ' =========================================================
    ' 6. サルベージ提案処理
    ' =========================================================
    If overflowFavCol.Count > 0 Then
        Dim ansSalvage As VbMsgBoxResult
        ansSalvage = MsgBox("お気に入りの登録上限（50件）に達したため、" & overflowFavCol.Count & " 件のデータがインポートできませんでした。" & vbCrLf & _
                            "この溢れたデータを別のJSONファイルとして保存（サルベージ）しますか？" & vbCrLf & vbCrLf & _
                            "Limit reached (50 items). Do you want to save the remaining " & overflowFavCol.Count & " items as a new file?", _
                            vbYesNo + vbQuestion, APP_NAME)
        
        If ansSalvage = vbYes Then
            Dim salvagePath As Variant
            Dim salvageDict As Object
            
            salvagePath = Application.GetSaveAsFilename( _
                InitialFileName:="KENZEN_Fav_Overflow_" & Format(Now, "yyyymmdd_HHmmss") & ".json", _
                FileFilter:="JSON Files (*.json), *.json", _
                title:="溢れたデータを保存 (Save Overflow Data)")
                
            If salvagePath <> False Then
                Set salvageDict = CreateObject("Scripting.Dictionary")
                salvageDict.Add "Fav_List", overflowFavCol
                
                On Error GoTo SalvageError
                jsonStr = JsonConverter.ConvertToJson(salvageDict, 4)
                Set ado = CreateObject("ADODB.Stream")
                With ado
                    .Type = 2
                    .Charset = "UTF-8"
                    .Open
                    .WriteText jsonStr
                    .SaveToFile salvagePath, 2
                    .Close
                End With
                On Error GoTo 0
                MsgBox "溢れたデータを無事にサルベージしました！" & vbCrLf & "Overflow data saved successfully!", vbInformation, APP_NAME
            End If
        End If
    End If

    ' =========================================================
    ' ★追加・修正：全てのインポート処理が終わったここで同期をかける
    ' =========================================================
    Call SyncFavoritesManager
    
    ' =========================================================
    ' 7. 結果レポートの動的構築
    ' =========================================================
    reportMsg = "データのインポート処理が完了しました！" & vbCrLf & "(Data import complete!)" & vbCrLf & vbCrLf & "【追加・更新された件数 (Results)】"
    If Me.chkImportPosiPresets.Value Then reportMsg = reportMsg & vbCrLf & "- Positive Presets: " & cntPosi & " 件"
    If Me.chkImportNegaPresets.Value Then reportMsg = reportMsg & vbCrLf & "- Negative Presets: " & cntNega & " 件"
    If Me.chkImportNegaStock.Value Then reportMsg = reportMsg & vbCrLf & "- Negative Stock: " & cntNegaStock & " タグ"
    If Me.chkImportLoRA.Value Then reportMsg = reportMsg & vbCrLf & "- LoRA Base Data: " & cntLoRA & " 件 (※新規追加、および固有ネガティブの統合件数)"
    If Me.chkImportLoRAPresets.Value Then reportMsg = reportMsg & vbCrLf & "- LoRA Presets: " & cntLoRAPreset & " 件"
    If Me.chkImportFav.Value Then reportMsg = reportMsg & vbCrLf & "- Favorites: " & cntFav & " 件"
    
    ' ★追加：モバイルメモのレポート
    If Me.chkImportMobiMemo.Value Then
        reportMsg = reportMsg & vbCrLf & "- Mobile URL Stock: " & cntMobiUrl & " 件"
        If mobiMemoUpdated Then
            reportMsg = reportMsg & vbCrLf & "- Mobile Memo Text: 更新あり (Updated)"
        End If
    End If
    
    ' ★修正：モバイル関連のカウントも空振り判定に含める
    If modeIsMerge And cntPosi = 0 And cntNega = 0 And cntNegaStock = 0 And cntLoRAPreset = 0 And cntLoRA = 0 And cntFav = 0 And cntMobiUrl = 0 And (Not mobiMemoUpdated) Then
        reportMsg = reportMsg & vbCrLf & vbCrLf & "※すべて既存のデータと重複していたため、新しい追加はありませんでした。"
    End If
    
    MsgBox reportMsg, vbInformation, APP_NAME
    Unload Me
    Call OpenMainWindowAction

    Exit Sub
    
    
    
ReadError:
    MsgBox "ファイルの読み込みに失敗しました。JSONの形式が正しくない可能性があります。" & vbCrLf & _
           "Failed to read the file.", vbCritical, APP_NAME
    Exit Sub
    
SalvageError:
    MsgBox "サルベージファイルの作成中にエラーが発生しました。" & vbCrLf & "Failed to create salvage file.", vbCritical, APP_NAME
End Sub


' =========================================================================
' [Mobile Export] お気に入りデータをモバイル閲覧用JSONとして単独書き出し
' =========================================================================
Private Sub btnDoMobiExport_Click()
    Dim exportDict As Object
    Dim filePath As Variant
    Dim jsonStr As String
    Dim ado As Object
    
    ' 門番：シートチェック
    If Not IsFavSheetActive() Then Exit Sub
    
    ' 1. ConfigDataが存在するかチェック
    If ConfigData Is Nothing Then Exit Sub
    
    ' 2. Fav_Listが存在し、かつデータがあるかチェック
    If Not ConfigData.Exists("Fav_List") Then
        MsgBox "エクスポートするお気に入り（Fav）データがありません。" & vbCrLf & _
               "No Favorite data found to export.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If ConfigData("Fav_List").Count = 0 Then
        MsgBox "お気に入り（Fav）リストが空です。" & vbCrLf & _
               "Favorite list is empty.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' =========================================================
    ' 3. モバイル用に抽出 ＆ メモ用の空ツリーを作成
    ' =========================================================
    Set exportDict = CreateObject("Scripting.Dictionary")
    
    ' ① お気に入りデータをセット
    Set exportDict("Fav_List") = ConfigData("Fav_List")
    
    ' ② ★ここを追加：モバイル側で入力するメモ用の空の配列（Collection）を用意
    Dim emptyMemos As Collection
    Set emptyMemos = New Collection
    Set exportDict("Mobile_Memos") = emptyMemos
    
    
    ' 4. エクスポート先を指定するダイアログを表示
    ' ※ファイル名は固定提案に沿い「KENZEN_Mobile_Fav.json」をデフォルトセット
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=ThisWorkbook.Path & "\KENZEN_Mobile_Fav.json", _
        FileFilter:="JSON Files (*.json), *.json", _
        title:="モバイル用Favデータをエクスポート (Export Mobile Fav Data)")
        
    If filePath = False Then Exit Sub
    
    ' ※本体設定ファイル上書き防止チェック（念のためのフェイルセーフ）
    Dim systemConfigPath As String
    systemConfigPath = ThisWorkbook.Path & "\KENZEN_Config.json"
    If StrComp(filePath, systemConfigPath, vbTextCompare) = 0 Then
        MsgBox "システムが使用中の本体設定ファイルは指定できません。" & vbCrLf & _
               "Cannot overwrite the main system configuration file.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' 5. ディープクローン生成 ＆ 鉄壁のサニタイズ処理
    On Error GoTo ExportError
    Dim cloneDict As Object
    ' 一度シリアライズしてパースし直すことで、元のConfigDataから完全に独立した複製を作る
    Set cloneDict = JsonConverter.ParseJson(JsonConverter.ConvertToJson(exportDict))
    
    ' クローン側に対してダブルクォーテーション等の無害化を一斉適用
    Call SanitizeQuotesInDictionary(cloneDict)
    
    ' 安全な状態になったクローンを綺麗なインデントでJSON化
    jsonStr = JsonConverter.ConvertToJson(cloneDict, 4)
    
    ' 6. ADODB.StreamによるBOMなしUTF-8書き出し
    Set ado = CreateObject("ADODB.Stream")
    With ado
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText jsonStr
        .SaveToFile filePath, 2 ' adSaveCreateOverWrite
        .Close
    End With
    On Error GoTo 0
    
    ' 日英併記で成功メッセージを出力
    MsgBox "モバイル用Favデータのエクスポートが完了しました！" & vbCrLf & _
           "Export complete!", vbInformation, APP_NAME
    Exit Sub
    
ExportError:
    On Error Resume Next
    If Not ado Is Nothing Then ado.Close
    On Error GoTo 0
    
    ' 日英併記でエラーメッセージを出力
    MsgBox "エクスポート中にエラーが発生しました。" & vbCrLf & _
           "Error generating mobile export file.", vbCritical, APP_NAME
End Sub
' =========================================================================
' [Mobile Import] モバイル側で編集・追記されたJSONファイルを読み込む
' =========================================================================
Private Sub btnDoMobiImport_Click()
    Dim filePath As Variant
    Dim ado As Object
    Dim jsonStr As String
    Dim parsedJson As Object
    Dim memos As Collection
    Dim item As Variant
    Dim tag As String
    Dim content As String
    Dim i As Long
    Dim charsetToUse As String ' ★追加：文字コード動的指定用
    
    ' 1. インポートするファイルを選択（デフォルトはJSON）
    filePath = Application.GetOpenFilename( _
        FileFilter:="JSON Files (*.json), *.json", _
        title:="モバイル用データを選択 (Select Mobile Data)")
        
    If filePath = False Then Exit Sub
    
    ' =========================================================================
    ' ★最強の解決策：MainCodeに既にある文字コード判定関数を利用
    ' =========================================================================
    If IsFileUTF8(CStr(filePath)) Then
        charsetToUse = "UTF-8"
    Else
        charsetToUse = "Shift_JIS"
    End If
    
    ' 2. ADODB.StreamでJSONファイルを読み込む
    On Error GoTo ImportError
    Set ado = CreateObject("ADODB.Stream")
    With ado
        .Type = 2
        .Charset = charsetToUse ' ★自動判定された文字コードをセット
        .Open
        .LoadFromFile filePath
        jsonStr = .ReadText
        .Close
    End With
    
    ' 3. JSON文字列をパース
    Set parsedJson = JsonConverter.ParseJson(jsonStr)
    
    ' 4. "Mobile_Memos" ツリーが存在するかチェック
    If Not parsedJson.Exists("Mobile_Memos") Then
        MsgBox "指定されたファイルにモバイルメモのデータが存在しません。" & vbCrLf & _
               "No 'Mobile_Memos' found in the selected file.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Set memos = parsedJson("Mobile_Memos")
    
    If memos.Count = 0 Then
        MsgBox "インポートする新しいメモがありません（空です）。" & vbCrLf & _
               "The memo list is empty.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' 5. データを内容に応じてすべてリストボックスに格納する
    For Each item In memos
        tag = ""
        content = ""
        
        ' ① モバイル側からのデータ構造を取得
        If typeName(item) = "Dictionary" Then
            If item.Exists("Type") And item.Exists("Content") Then
                tag = item("Type") ' "Memo", "URL", "Hash" など
                content = item("Content")
            End If
        ElseIf typeName(item) = "String" Then
            content = CStr(item)
            If Left(content, 4) = "http" Then
                tag = "URL"
            Else
                tag = "Memo"
            End If
        End If
        
        ' ② すべてのデータをリストボックスへ統合
        If content <> "" Then
            Dim displayTag As String
            displayTag = tag
            
        ' URLの場合はさらに詳細に判別
            If tag = "URL" Then
                ' ★修正：「/ja/」への依存を無くし、ドメインと「/detail/」の有無でグローバルに判定する
                If InStr(content, "seaart.ai") > 0 And InStr(content, "/detail/") > 0 Then
                    displayTag = "SeaArt"
                    Dim parts() As String
                    parts = Split(content, "/detail/")
                    If UBound(parts) >= 1 Then
                        content = Split(parts(1), "?")(0)
                    End If
                ElseIf InStr(content, "civitai.com") > 0 Then
                    displayTag = "Civitai"
                End If
            End If
            
            ' リストボックスへ追加
            Me.lstMobileMemo.AddItem "[" & displayTag & "]"
            Me.lstMobileMemo.List(Me.lstMobileMemo.ListCount - 1, 1) = content
        End If
    Next item
    
    ' =========================================================================
    ' 6. インポート完了後、モバイル側JSONの "Mobile_Memos" を空にして安全に上書き保存
    ' =========================================================================
    Set parsedJson("Mobile_Memos") = New Collection
    
    Dim jsonOutput As String
    jsonOutput = JsonConverter.ConvertToJson(parsedJson, 4)
    
    Set ado = CreateObject("ADODB.Stream")
    With ado
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText jsonOutput
        .SaveToFile filePath, 2 ' 2 = 上書き保存
        .Close
    End With
    
' =========================================================================
' 7. 母艦（ConfigData）への自動セーブ処理 (btnDoMobiImport_Click の一部)
' =========================================================================
    If Not ConfigData Is Nothing Then
        Dim newUrlStock As Collection
        Set newUrlStock = New Collection
        
        For i = 0 To Me.lstMobileMemo.ListCount - 1
            ' ★修正：Null対策として明示的に空文字を結合してからデリミタで繋ぐ
            newUrlStock.Add (Me.lstMobileMemo.List(i, 0) & "") & "<|>" & (Me.lstMobileMemo.List(i, 1) & "")
        Next i
        
        Set ConfigData("Mobile_URL_Stock") = newUrlStock
        Call SaveConfigJSON
    End If
    
    On Error GoTo 0
    
    MsgBox "モバイルデータのインポートと、母艦への自動セーブが完了しました！" & vbCrLf & _
           "Import and Auto-Save complete!", vbInformation, APP_NAME
    Exit Sub

ImportError:
    On Error Resume Next
    If Not ado Is Nothing Then ado.Close
    On Error GoTo 0
    MsgBox "ファイルの読み込みまたは解析中にエラーが発生しました。" & vbCrLf & _
           "Error loading or parsing the file.", vbCritical, APP_NAME
End Sub

' =========================================================================
' [Mobile Export] メモリストをモバイル用JSONにマージ（任意のパス指定版）
' =========================================================================
Private Sub btnExportMobiMemos_Click()
    Dim fso As Object
    Dim mobiJsonPath As Variant
    Dim selectedFileName As String
    Dim jsonText As String
    Dim dictRoot As Object
    Dim colMemos As Collection
    Dim dictMemo As Object
    Dim i As Long
    Dim stream As Object
    Dim cleanType As String ' ★追加：カッコを取り除いたType文字列用
    
    ' =========================================================
    ' ★追加：リストが空の場合は警告を出して処理を安全に中断する
    ' =========================================================
    If Me.lstMobileMemo.ListCount = 0 Then
        MsgBox "エクスポートするメモがありません（リストが空です）。" & vbCrLf & _
               "There are no memos to export (the list is empty).", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ---------------------------------------------------------
    ' 1. モバイル用JSONファイルのパスをユーザーに指定させる
    ' ---------------------------------------------------------
    mobiJsonPath = Application.GetOpenFilename( _
        FileFilter:="JSONファイル (*.json), *.json", _
        title:="エクスポート先のモバイル用JSONファイルを選択してください (Select the mobile JSON file to export)")
        
    If VarType(mobiJsonPath) = vbBoolean And mobiJsonPath = False Then
        Set fso = Nothing
        Exit Sub
    End If
    
    ' ---------------------------------------------------------
    ' 2. 選択されたファイル名が正しいか（KENZEN_Mobile_Fav.json）チェック
    ' ---------------------------------------------------------
    selectedFileName = fso.GetFileName(mobiJsonPath)
    
    If LCase(selectedFileName) <> LCase("KENZEN_Mobile_Fav.json") Then
        MsgBox "選択されたファイル名が異なります。" & vbCrLf & _
               "The selected file name is incorrect." & vbCrLf & vbCrLf & _
               "選択されたファイル (Selected file): " & selectedFileName & vbCrLf & _
               "指定すべきファイル (Required file): KENZEN_Mobile_Fav.json" & vbCrLf & vbCrLf & _
               "モバイル版との連携を保つため、正しいファイルを選択し直してください。" & vbCrLf & _
               "Please select the correct file to maintain compatibility with the mobile version.", _
               vbCritical, APP_NAME
        Set fso = Nothing
        Exit Sub
    End If
    
    ' ---------------------------------------------------------
    ' 3. 既存のJSONファイルを読み込む (Fav_List を壊さないため)
    ' ---------------------------------------------------------
    If fso.FileExists(mobiJsonPath) Then
        Set stream = CreateObject("ADODB.Stream")
        With stream
            .Type = 2
            .Charset = "UTF-8"
            .Open
            .LoadFromFile mobiJsonPath
            jsonText = .ReadText
            .Close
        End With
        Set dictRoot = JsonConverter.ParseJson(jsonText)
    Else
        Set dictRoot = CreateObject("Scripting.Dictionary")
        dictRoot.Add "Fav_List", New Collection
    End If
    
    ' ---------------------------------------------------------
    ' 4. 現在のリストボックスから Mobile_Memos コレクションを再構築
    ' ---------------------------------------------------------
    Set colMemos = New Collection
    
    For i = 0 To Me.lstMobileMemo.ListCount - 1
        Set dictMemo = CreateObject("Scripting.Dictionary")
        
        ' ★修正：0列目のNullエラーを防ぐ
        cleanType = Me.lstMobileMemo.List(i, 0) & ""
        cleanType = Replace(cleanType, "[", "")
        cleanType = Replace(cleanType, "]", "")
        
        dictMemo.Add "Type", cleanType
        ' ★修正：1列目のNullエラーを防ぐ
        dictMemo.Add "Content", Me.lstMobileMemo.List(i, 1) & ""
        
        colMemos.Add dictMemo
    Next i
    
    If dictRoot.Exists("Mobile_Memos") Then
        Set dictRoot("Mobile_Memos") = colMemos
    Else
        dictRoot.Add "Mobile_Memos", colMemos
    End If
    
    ' ---------------------------------------------------------
    ' 5. ★修正：JSON文字列に変換し、安全なUTF-8形式でシンプルに保存
    ' ---------------------------------------------------------
    jsonText = JsonConverter.ConvertToJson(dictRoot, 4)
    
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText jsonText
        .SaveToFile mobiJsonPath, 2 ' 2 = 上書き保存
        .Close
    End With
    
    Set stream = Nothing
    Set fso = Nothing
    
    MsgBox "モバイルメモをエクスポートしました！" & vbCrLf & _
           "Mobile memos exported successfully!" & vbCrLf & vbCrLf & _
           "保存先 (Saved to): " & vbCrLf & mobiJsonPath, vbInformation, APP_NAME
End Sub
' =========================================================================
' [Mobile Utility] モバイル用UIを最新のConfigDataで再描画する
' =========================================================================
Public Sub RefreshMobileUI()
    ' 既存のリストをクリア
    lstMobileMemo.Clear
    
    Dim i As Integer
    Dim storedItem As String
    Dim splitData() As String
    Dim urlStock As Collection ' ★追加：ConfigDataから受け取るための変数を宣言
    
    ' ★修正：ConfigDataが初期化されており、キーが存在するかを安全にチェック
    If Not ConfigData Is Nothing Then
        If ConfigData.Exists("Mobile_URL_Stock") Then
            ' Dictionaryからコレクションをセット
            Set urlStock = ConfigData("Mobile_URL_Stock")
            
            ' コレクションが正常に取得できた場合のみループを回す
            If Not urlStock Is Nothing Then
                For i = 1 To urlStock.Count
                    storedItem = urlStock(i)
                    
                    ' 空文字チェック
                    If Len(Trim(storedItem)) > 0 Then
                        
                        If InStr(storedItem, "<|>") > 0 Then
                            ' ★修正：ユーザー入力内に万が一 "<|>" が含まれていた場合の破壊を防ぐため Limit=2 を指定
                            splitData = Split(storedItem, "<|>", 2)
                        Else
                            ' 旧仕様データの救済措置
                            splitData = Split(storedItem, " ", 2)
                        End If
                        
                        ' バウンズチェック (正常に分割されていれば要素の最大インデックスは1以上になる)
                        If UBound(splitData) >= 1 Then
                            lstMobileMemo.AddItem splitData(0) ' 1列目 (Tag)
                            lstMobileMemo.List(lstMobileMemo.ListCount - 1, 1) = splitData(1) ' 2列目 (URL/Content)
                        Else
                            ' 区切り文字が見つからないなど、フォーマットが不正な場合のフォールバック
                            lstMobileMemo.AddItem "【破損または未分類】"
                            lstMobileMemo.List(lstMobileMemo.ListCount - 1, 1) = storedItem
                        End If
                    End If
                Next i
            End If
        End If
    End If
End Sub

'-------------------------------------------------------------------------
' [Forget] ボタンが押された時の処理（画面上の全てを初期状態に戻す完全更地化）
'-------------------------------------------------------------------------
Private Sub btnForgetLoRA_Click()
    
    ' 1. ユーザーへの最終確認（誤爆防止の安全装置）
    If MsgBox("リストに追加したLoRAと現在の設定をすべて消去し、最初からやり直しますか？" & vbCrLf & _
              "Are you sure you want to clear all LoRA settings and start over?", _
              vbQuestion + vbYesNo, APP_NAME) = vbNo Then
        Exit Sub
    End If
    
    ' 2. 手元の入力エリアの更地化（Cancelボタンの処理を呼び出して使い回す）
    Call btnLoRACancel_Click
    
    ' 3. トリガーの重み(Weight)をデフォルト(1.0)に戻す（Index 0）
    Me.cmbTriggerWeight.ListIndex = 0
    
    ' 4. 買い物かご（リストボックス）を完全に空にする
    Me.lstSelectedLoRA.Clear
    
    ' 5. プレビュー（出力エリア）を空にする
    Me.txtLoRAPreview.text = ""
    Call lstSelectedLoRA_Change

    If Me.txtYourNegative.Value <> "" Then
       Me.txtYourNegative.Value = ""
    End If

    ' (オプション) 完了を少しだけアピールする場合
    MsgBox "すべての設定をクリアしました。" & vbCrLf & _
           "All settings have been cleared.", vbInformation, APP_NAME
    
End Sub

'-------------------------------------------------------------------------
' [Cancel] ボタンが押された時の処理（入力エリアの更地化）
'-------------------------------------------------------------------------
Private Sub btnLoRACancel_Click()
    Dim i As Integer
    
    ' 1. コンボボックスの選択をクリア（未選択状態にする）
    Me.cmbLoRAList.ListIndex = -1
    Me.cmbLoRAStrength.ListIndex = -1
    
    ' 2. トリガー不要ラベルを非表示にする
    Me.lblNotRequiredTrigger.Visible = False
    
    ' 3. トリガーワードのチェックボックス群を完全に初期化
    For i = 1 To 10
        With Me.Controls("chkTrigger_" & i)
            .Caption = ""          ' 文字を消去
            .Value = False         ' チェックを外す
            .Visible = False       ' 画面から隠す
        End With
    Next i
    
    ' 4. 次の操作がしやすいように、フォーカスをLoRA選択リストに戻す
    Me.cmbLoRAList.SetFocus
End Sub

' -------------------------------------------------------------------------
' [I/Oタブ] 旧バージョンのCSVをJSONに変換（移行ツール）
' -------------------------------------------------------------------------
Private Sub btnMigrateLegacyCSV_Click()
    If MsgBox("旧バージョン(v2系)でエクスポートしたCSV(Fav,LoRA)を、v3.0.0用のJSONデータに変換します。" & vbCrLf & _
              "変換されたJSONファイルは、上部の「Import(Merge) Fav」から追加できます。" & vbCrLf & vbCrLf & _
              "変換を開始しますか？ (Do you want to convert legacy CSV Fav or LoRA to JSON?)", _
              vbYesNo + vbInformation, APP_NAME) = vbYes Then
        
        Call ConvertLegacyCSVtoJSON
        
    End If
End Sub
Private Sub btnOpenFavManage_Click()
    
    If Not IsFavSheetActive() Then Exit Sub ' 門番
    FavoritesManager.Show

End Sub

Private Sub btnOpenManageLoRA_Click()
'-------------------------------------------------------------------------
' [Open LoRA Manage Window] ボタン
'-------------------------------------------------------------------------
    ' 1. 管理画面を表示（モーダル表示なので、閉じるまで下の行には進みません）
    LoRARegister.Show
    
    ' 2. 管理画面を閉じてここに戻ってきた瞬間、メイン画面の選択肢を最新にする
    Call RefreshMainWindowLoRA
    
    ' 3. 修正：不要になった自動Wrapを廃止し、UIのグレーアウト状態だけを正しく最新にする
    Call UpdateLoRAUIState
End Sub

'-------------------------------------------------------------------------
' [Remove] ボタンが押された時の処理（非破壊・プレビュー自動再構築版）
'-------------------------------------------------------------------------
Private Sub btnRemoveLoRA_Click()
    Dim i As Integer
    Dim hasSelection As Boolean
    Dim wasWrapped As Boolean
    
    ' 1. バリデーション：本当に選択(ハイライト)されているかチェック
    hasSelection = False
    For i = 0 To Me.lstSelectedLoRA.ListCount - 1
        If Me.lstSelectedLoRA.Selected(i) = True Then
            hasSelection = True
            Exit For
        End If
    Next i
    
    If Not hasSelection Then
        MsgBox "削除するLoRAをリストから選択してください。" & vbCrLf & _
               "Please select a LoRA from the list to remove.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' --- 削除アクションの「前」の状態を記憶 ---
    wasWrapped = (Trim(Me.txtLoRAPreview.text) <> "")
    
    ' 2. 選択された行を削除する（複数選択時のインデックスずれを防ぐため下からループ）
    For i = Me.lstSelectedLoRA.ListCount - 1 To 0 Step -1
        If Me.lstSelectedLoRA.Selected(i) = True Then
            Me.lstSelectedLoRA.RemoveItem i
        End If
    Next i
    
    ' 3. プレビューエリアを一度「更地」にする
    Me.txtLoRAPreview.text = ""
    
    ' 4. リストにデータが残り、かつ「既にWrapされていた場合」のみ【非破壊】で再構築させる
    If Me.lstSelectedLoRA.ListCount > 0 And wasWrapped Then
        
        ' ★修正の要：btnWrapLoRA_Clickを呼ばず、リストが記憶しているトリガー状態をそのまま結合する
        ' これにより、個別のトリガーに既に設定された重みや、素のトリガーが破壊されるのを防ぐ
        Dim rebuildTriggers As String
        Dim rebuildLoRAs As String
        Dim rAlias As String, rStrength As String, rTrigger As String
        Dim rHash As String, rModelName As String, rTarget As String
        Dim rItem As Object, rList As Object
        
        rebuildTriggers = ""
        rebuildLoRAs = ""
        
        For i = 0 To Me.lstSelectedLoRA.ListCount - 1
            rAlias = Me.lstSelectedLoRA.List(i, 2)
            rStrength = Me.lstSelectedLoRA.List(i, 3)
            rTrigger = Trim(Me.lstSelectedLoRA.List(i, 1)) ' ←リストに保存されている状態（重み含む）をそのまま抽出
            rHash = ""
            rModelName = ""
            
            ' メモリ上のLoRA_Listから該当データを検索
            If Not ConfigData Is Nothing Then
                If ConfigData.Exists("LoRA_List") Then
                    Set rList = ConfigData("LoRA_List")
                    For Each rItem In rList
                        If rItem("Alias") = rAlias Then
                            rHash = rItem("Hash")
                            If rItem.Exists("ModelName") Then rModelName = rItem("ModelName")
                            Exit For
                        End If
                    Next rItem
                End If
            End If
            
            ' オプションボタン（出力ターゲット）の条件分岐判定
            If Me.opbWrapName.Value = True And Trim(rModelName) <> "" Then
                rTarget = Trim(rModelName)
            Else
                If Trim(rHash) <> "" Then rTarget = Trim(rHash) Else rTarget = Trim(rAlias)
            End If
            
            If rTarget <> "" Then
                ' LoRAタグの結合（重複防止）
                Dim curLoRA As String
                curLoRA = "<lora:" & rTarget & ":" & rStrength & ">"
                If InStr(1, rebuildLoRAs, curLoRA, vbTextCompare) = 0 Then
                    If rebuildLoRAs = "" Then rebuildLoRAs = curLoRA Else rebuildLoRAs = rebuildLoRAs & " " & curLoRA
                End If
                
                ' トリガーワードの結合（カンマ区切りで重複防止）
                If rTrigger <> "" Then
                    Dim tArr() As String
                    Dim t As Integer
                    tArr = Split(rTrigger, ",")
                    For t = 0 To UBound(tArr)
                        Dim singleTrigger As String
                        singleTrigger = Trim(tArr(t))
                        If singleTrigger <> "" Then
                            ' 重複チェック
                            If InStr(1, ", " & rebuildTriggers & ",", ", " & singleTrigger & ",", vbTextCompare) = 0 Then
                                If rebuildTriggers = "" Then
                                    rebuildTriggers = singleTrigger
                                Else
                                    rebuildTriggers = rebuildTriggers & ", " & singleTrigger
                                End If
                            End If
                        End If
                    Next t
                End If
            End If
        Next i
        
        ' プレビュー欄へ最終出力
        If rebuildTriggers <> "" And rebuildLoRAs <> "" Then
            Me.txtLoRAPreview.text = rebuildTriggers & ", " & rebuildLoRAs
        ElseIf rebuildTriggers <> "" Then
            Me.txtLoRAPreview.text = rebuildTriggers
        ElseIf rebuildLoRAs <> "" Then
            Me.txtLoRAPreview.text = rebuildLoRAs
        End If
        
    End If
    
    ' 5. UI状態の更新とフォーカス
    Call UpdateLoRAUIState
    Me.lstSelectedLoRA.SetFocus
    Call lstSelectedLoRA_Change
    
    If Me.txtYourNegative.Value <> "" Then
       Me.txtYourNegative.Value = ""
    End If
End Sub
'-------------------------------------------------------------------------
' [Prompt Sort] ボタン：v3.0.6 最新版（Dynamic Prompts完全対応 ＋ 正規表現PACKエンジン）
'-------------------------------------------------------------------------
Private Sub btnPromptSort_Click()
    Dim fullText As String
    Dim paragraphArr() As String
    Dim pIndex As Integer
    Dim finalFullPrompt As String
    
    Dim rawText As String
    Dim basePosText As String, loraBlock As String
    Dim regEx As Object, matches As Object, m As Object
    Dim wsMain As Worksheet
    Dim chunkArr() As String
    Dim i As Integer, j As Integer, wordCount As Integer
    Dim cleanStr As String
    
    ' --- ソート用 配列 ---
    Dim rawArr() As String
    Dim scoreArr() As Double
    Dim indexArr() As Integer
    
    ' --- PACK復元用の変数 ---
    Dim origPhrase As String, origCol As Integer, unbackedChunk As String
    
    fullText = Trim(Me.txtMain.Value)
    
    ' ========================================================
    ' 0. バリデーションと事前準備
    ' ========================================================
    If fullText = "" Then
        MsgBox "ソートするプロンプトが入力されていません。" & vbCrLf & _
               "Please enter prompts to sort.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim emptyCheckStr As String
    emptyCheckStr = Replace(Replace(Replace(Replace(Replace(Replace(fullText, vbCrLf, ""), vbCr, ""), vbLf, ""), ",", ""), " ", ""), " ", "")
    If emptyCheckStr = "" Then
        MsgBox "有効なプロンプトのテキストが見つかりません。" & vbCrLf & _
               "No valid text found to sort.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If UBound(Split(fullText, ",")) > 300 Then
        MsgBox "プロンプトの数が多すぎます（上限300個）。" & vbCrLf & _
               "Systems freeze protection activated.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    If InStr(1, Me.txtYourPositive.Value, "<lora:", vbTextCompare) > 0 Then
        MsgBox "【警告】ベース品質（ポジティブ）入力欄にLoRAタグが混入しています。" & vbCrLf & _
               "LoRA tags found in positive field.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Call SaveHistory(fullText)
    
    Set wsMain = ThisWorkbook.Sheets("KENZEN SeaArt Helper")
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    
    fullText = Replace(Replace(fullText, vbCrLf, vbLf), vbCr, vbLf)
    
    ' ========================================================
    ' ★バグ修正：BREAKの分割時のみ、大文字小文字を厳密に区別する
    ' ========================================================
    regEx.IgnoreCase = False
    regEx.Pattern = "\s*BREAK\s*"
    fullText = regEx.Replace(fullText, vbLf & "BREAK" & vbLf)
    
    ' 以降の処理（PACKエンジン等）のために大文字小文字の区別を無効化に戻す
    regEx.IgnoreCase = True
    ' ========================================================
    
    While InStr(fullText, vbLf & vbLf) > 0
        fullText = Replace(fullText, vbLf & vbLf, vbLf)
    Wend
    paragraphArr = Split(fullText, vbLf)
    finalFullPrompt = ""
    
    ' ========================================================
    ' 1. 激速化エンジン：全辞書データをメモリに1回だけキャッシュする
    ' ========================================================
    Application.StatusBar = "● Building memory cache for ultra-fast sorting..."
    
    Dim loraDict As Object
    Set loraDict = CreateObject("Scripting.Dictionary")
    loraDict.CompareMode = vbTextCompare
    
    If Not ConfigData Is Nothing Then
        If ConfigData.Exists("LoRA_List") Then
            Dim loraItem As Object
            For Each loraItem In ConfigData("LoRA_List")
                
                Dim isActiveLora As Boolean
                isActiveLora = False
                
                If InStr(1, fullText, "<lora:" & loraItem("Hash"), vbTextCompare) > 0 Then isActiveLora = True
                If InStr(1, fullText, "<lora:" & loraItem("Alias"), vbTextCompare) > 0 Then isActiveLora = True
                If loraItem.Exists("ModelName") Then
                    If InStr(1, fullText, "<lora:" & loraItem("ModelName"), vbTextCompare) > 0 Then isActiveLora = True
                End If
                
                If isActiveLora Then
                    loraDict(loraItem("Hash")) = 0.5
                    loraDict(loraItem("Alias")) = 0.5
                    
                    If loraItem.Exists("ModelName") Then
                        If Trim(loraItem("ModelName")) <> "" Then loraDict(Trim(loraItem("ModelName"))) = 0.5
                    End If
                    
                    Dim tArr() As String, t As Variant
                    tArr = Split(loraItem("Trigger"), ",")
                    For Each t In tArr
                        If Trim(t) <> "" Then loraDict(Trim(t)) = 0.5
                    Next
                End If
                
            Next
        End If
    End If
    
    ' ========================================================
    ' 助詞・前置詞（スキップ対象）辞書の作成
    ' ========================================================
    Dim prepDict As Object
    Set prepDict = CreateObject("Scripting.Dictionary")
    prepDict.CompareMode = vbTextCompare
    
    Dim prepData As Variant
    prepData = wsMain.Range("A15:ZZ17").Value
    Dim pc As Long
    For pc = 1 To UBound(prepData, 2)
        If Not IsEmpty(prepData(1, pc)) And Not IsError(prepData(1, pc)) Then
            If Trim(CStr(prepData(1, pc))) <> "" Then prepDict(Trim(CStr(prepData(1, pc)))) = True
        End If
        If Not IsEmpty(prepData(3, pc)) And Not IsError(prepData(3, pc)) Then
            If Trim(CStr(prepData(3, pc))) <> "" Then prepDict(Trim(CStr(prepData(3, pc)))) = True
        End If
    Next pc
    
    prepDict("NOT") = True
    
    ' ========================================================
    ' メイン辞書の作成
    ' ========================================================
    Dim exactDict As Object
    Set exactDict = CreateObject("Scripting.Dictionary")
    exactDict.CompareMode = vbTextCompare
    
    Dim mainData As Variant
    mainData = wsMain.Range("B20:CB500").Value
    
    Dim phraseList() As String, phraseCol() As Integer
    Dim pCount As Long
    pCount = 0
    ReDim phraseList(1 To 2000): ReDim phraseCol(1 To 2000)
    
    Dim r As Long, c As Long
    For c = 1 To UBound(mainData, 2) Step 2
        If c = 7 Then GoTo NextMainCol
        
        For r = 1 To UBound(mainData, 1)
            If Not IsError(mainData(r, c)) And Not IsEmpty(mainData(r, c)) Then
                Dim exactDWord As String
                exactDWord = Trim(CStr(mainData(r, c)))
                
                If exactDWord <> "" Then
                    exactDWord = Replace(exactDWord, Chr(160), " ")
                    exactDWord = Replace(exactDWord, vbTab, "")
                    exactDWord = Replace(exactDWord, vbCr, "")
                    exactDWord = Replace(exactDWord, vbLf, "")
                    exactDWord = Trim(exactDWord)
                    
                    Dim origExactDWord As String
                    origExactDWord = exactDWord
                    While InStr(origExactDWord, "  ") > 0
                        origExactDWord = Replace(origExactDWord, "  ", " ")
                    Wend
                    
                    If Len(exactDWord) >= 4 And Left(exactDWord, 2) = "__" And Right(exactDWord, 2) = "__" Then
                    Else
                        exactDWord = Replace(exactDWord, "_", " ")
                        exactDWord = Replace(exactDWord, "-", " ")
                    End If
                    
                    While InStr(exactDWord, "  ") > 0
                        exactDWord = Replace(exactDWord, "  ", " ")
                    Wend
                    exactDWord = Trim(exactDWord)
                    
                    If exactDWord <> "" Then
                        If Not exactDict.Exists(exactDWord) Then exactDict.Add exactDWord, c + 1
                        
                        If InStr(origExactDWord, ",") > 0 Then
                            pCount = pCount + 1
                            If pCount > UBound(phraseList) Then
                                ReDim Preserve phraseList(1 To pCount + 1000)
                                ReDim Preserve phraseCol(1 To pCount + 1000)
                            End If
                            phraseList(pCount) = origExactDWord
                            phraseCol(pCount) = c + 1
                        End If
                    End If
                End If
            End If
        Next r
NextMainCol:
    Next c
    
    If pCount > 0 Then
        For i = 1 To pCount - 1
            For j = pCount To i + 1 Step -1
                If Len(phraseList(j)) > Len(phraseList(j - 1)) Then
                    Dim tmpS As String, tmpC As Integer
                    tmpS = phraseList(j): phraseList(j) = phraseList(j - 1): phraseList(j - 1) = tmpS
                    tmpC = phraseCol(j): phraseCol(j) = phraseCol(j - 1): phraseCol(j - 1) = tmpC
                End If
            Next j
        Next i
    End If
    
    Application.StatusBar = "● Cache built. Sorting now..."

    ' ========================================================
    ' ★ループ開始：各段落ごとに独立してソートを行う
    ' ========================================================
    For pIndex = 0 To UBound(paragraphArr)
        rawText = Trim(paragraphArr(pIndex))
        
        If UCase(rawText) = "BREAK" Then
            If finalFullPrompt = "" Then finalFullPrompt = "BREAK" Else finalFullPrompt = finalFullPrompt & vbCrLf & "BREAK"
            GoTo NextParagraph
        End If
        
        If rawText = "" Then GoTo NextParagraph
        
        ' 1. Absolute Rearguard (LoRA) Isolation
        loraBlock = ""
        regEx.Pattern = "<lora:[^>]+>"
        Set matches = regEx.Execute(rawText)
        For Each m In matches
            loraBlock = loraBlock & m.Value & " "
            rawText = Replace(rawText, m.Value, "")
        Next
        loraBlock = Trim(loraBlock)
        
        ' ========================================================
        ' ★ 1.5. Dynamic Prompts 専用 PACK エンジン (DP_PACK)
        ' ========================================================
        Dim dpDict As Object
        Set dpDict = CreateObject("Scripting.Dictionary")
        Dim dpIdx As Integer
        dpIdx = 1
        
        ' {} で囲まれたブロック全体を検索
        regEx.Pattern = "\{[^}]+\}"
        regEx.Global = True
        If regEx.Test(rawText) Then
            Set matches = regEx.Execute(rawText)
            For Each m In matches
                Dim dpKey As String
                dpKey = "__DP_PACK_" & dpIdx & "__"
                ' ブロック全体を DP_PACK キーに置換して保護（カンマ分割を回避）
                rawText = Replace(rawText, m.Value, dpKey, 1, 1, vbTextCompare)
                dpDict.Add dpKey, m.Value
                dpIdx = dpIdx + 1
            Next m
        End If

        ' ========================================================
        ' 2. カンマを破壊する「前」にフレーズをパック（保護）する
        ' ========================================================
        Dim packDict As Object
        Set packDict = CreateObject("Scripting.Dictionary")
        Dim packIdx As Integer
        packIdx = 1
        
        If pCount > 0 Then
            For i = 1 To pCount
                Dim searchPattern As String
                searchPattern = phraseList(i)
                
                searchPattern = Replace(searchPattern, "\", "\\")
                searchPattern = Replace(searchPattern, "(", "\(")
                searchPattern = Replace(searchPattern, ")", "\)")
                searchPattern = Replace(searchPattern, "[", "\[")
                searchPattern = Replace(searchPattern, "]", "\]")
                searchPattern = Replace(searchPattern, "+", "\+")
                searchPattern = Replace(searchPattern, ".", "\.")
                searchPattern = Replace(searchPattern, "?", "\?")
                searchPattern = Replace(searchPattern, "*", "\*")
                searchPattern = Replace(searchPattern, "|", "\|")
                searchPattern = Replace(searchPattern, "{", "\{")
                searchPattern = Replace(searchPattern, "}", "\}")
                searchPattern = Replace(searchPattern, "^", "\^")
                searchPattern = Replace(searchPattern, "$", "\$")
                
                searchPattern = Replace(searchPattern, " ", "###SP###")
                searchPattern = Replace(searchPattern, "_", "###SP###")
                searchPattern = Replace(searchPattern, "-", "###SP###")
                searchPattern = Replace(searchPattern, "###SP###", "[-_ ]")
                searchPattern = Replace(searchPattern, ",", "\s*,\s*")
                
                regEx.Pattern = searchPattern
                regEx.Global = True
                regEx.IgnoreCase = True
                
                If regEx.Test(rawText) Then
                    Set matches = regEx.Execute(rawText)
                    For Each m In matches
                        Dim packKey As String
                        packKey = "__PACK_" & packIdx & "__"
                        rawText = Replace(rawText, m.Value, packKey, 1, 1, vbTextCompare)
                        packDict.Add packKey, Array(phraseList(i), phraseCol(i))
                        packIdx = packIdx + 1
                    Next m
                End If
            Next i
        End If
        
        ' 3. カンマの正規化とブラケット除去
        regEx.Pattern = "\s*,\s*"
        rawText = regEx.Replace(rawText, ",")
        regEx.Pattern = ",+"
        rawText = regEx.Replace(rawText, ",")
        
        If Left(rawText, 1) = "," Then rawText = Mid(rawText, 2)
        If Right(rawText, 1) = "," Then rawText = Left(rawText, Len(rawText) - 1)
        
        Dim isBracketParagraph As Boolean
        isBracketParagraph = False
        If InStr(rawText, "[") > 0 Or InStr(rawText, "]") > 0 Then
            isBracketParagraph = True
            rawText = Replace(rawText, "[", "")
            rawText = Replace(rawText, "]", "")
        End If
        
        ' 4. Absolute Vanguard (Base Quality) Dictionary Creation
        basePosText = ""
        If pIndex = 0 Then
            basePosText = Trim(Me.txtYourPositive.Value)
            Do While Right(basePosText, 1) = "," Or Right(basePosText, 1) = " "
                basePosText = Trim(Left(basePosText, Len(basePosText) - 1))
            Loop
            Do While Left(basePosText, 1) = "," Or Left(basePosText, 1) = " "
                basePosText = Trim(Mid(basePosText, 2))
            Loop
        End If
        
        Dim basePosDict As Object
        Set basePosDict = CreateObject("Scripting.Dictionary")
        basePosDict.CompareMode = vbTextCompare
        
        If basePosText <> "" Then
            Dim bArr() As String
            bArr = Split(basePosText, ",")
            For i = 0 To UBound(bArr)
                Dim bWord As String
                bWord = Trim(bArr(i))
                If bWord <> "" Then
                    Dim bClean As String
                    regEx.Pattern = "\s*:\s*\d+(\.\d+)?"
                    bClean = regEx.Replace(bWord, "")
                    regEx.Pattern = "[\(\)\[\]\{\}]"
                    bClean = regEx.Replace(bClean, "")
                    bClean = Trim(bClean)
                    
                    If Len(bClean) >= 4 And Left(bClean, 2) = "__" And Right(bClean, 2) = "__" Then
                    Else
                        bClean = Replace(Replace(bClean, "_", " "), "-", " ")
                    End If
                    bClean = Trim(bClean)
                    If bClean <> "" And Not basePosDict.Exists(bClean) Then basePosDict.Add bClean, True
                End If
            Next i
        End If
        
        ' 5. Chunking, Deduplication, and Scoring
        chunkArr = Split(rawText, ",")
        wordCount = 0
        ReDim rawArr(0 To UBound(chunkArr))
        ReDim scoreArr(0 To UBound(chunkArr))
        ReDim indexArr(0 To UBound(chunkArr))
        
        For i = 0 To UBound(chunkArr)
            Dim rawChunk As String
            rawChunk = Trim(chunkArr(i))
            
            If rawChunk <> "" Then
                Dim isDuplicate As Boolean
                isDuplicate = False
                For j = 0 To wordCount - 1
                    If StrComp(rawChunk, rawArr(j), vbTextCompare) = 0 Then isDuplicate = True: Exit For
                Next j
                
                If Not isDuplicate Then
                    Dim currentScore As Double
                    currentScore = 999
                    
                    regEx.Pattern = "\s*:\s*\d+(\.\d+)?"
                    cleanStr = regEx.Replace(rawChunk, "")
                    regEx.Pattern = "[\(\)\[\]\{\}]"
                    cleanStr = regEx.Replace(cleanStr, "")
                    cleanStr = Trim(cleanStr)
                    
                    ' ========================================================
                    ' ★ DP_PACK (Dynamic Prompts) のスコアリングすり替えロジック
                    ' ========================================================
                    If Left(cleanStr, 10) = "__DP_PACK_" Then
                        If dpDict.Exists(cleanStr) Then
                            Dim origDP As String
                            origDP = dpDict(cleanStr) ' 例: "{ass up, face down | standing}"
                            
                            Dim firstTag As String
                            ' 波括弧を除去
                            firstTag = Mid(origDP, 2, Len(origDP) - 2)
                            ' パイプがある場合は最初のタグのみを抽出
                            If InStr(firstTag, "|") > 0 Then
                                firstTag = Split(firstTag, "|")(0)
                            End If
                            firstTag = Trim(firstTag)
                            
                            ' 先頭タグを通常のタグと同じようにクリーニング
                            regEx.Pattern = "\s*:\s*\d+(\.\d+)?"
                            firstTag = regEx.Replace(firstTag, "")
                            regEx.Pattern = "[\(\)\[\]\{\}]"
                            firstTag = regEx.Replace(firstTag, "")
                            
                            If Len(firstTag) >= 4 And Left(firstTag, 2) = "__" And Right(firstTag, 2) = "__" Then
                            Else
                                firstTag = Replace(firstTag, "_", " ")
                                firstTag = Replace(firstTag, "-", " ")
                            End If
                            firstTag = Replace(firstTag, Chr(160), " ")
                            While InStr(firstTag, "  ") > 0
                                firstTag = Replace(firstTag, "  ", " ")
                            Wend
                            firstTag = Trim(firstTag)
                            
                            ' すり替えた先頭タグでスコアを計算
                            currentScore = 999
                            If loraDict.Exists(firstTag) Then
                                currentScore = 0.5
                            ElseIf exactDict.Exists(firstTag) Then
                                currentScore = exactDict(firstTag)
                            Else
                                Dim dpSubParts() As String
                                dpSubParts = Split(firstTag, " ")
                                If UBound(dpSubParts) > 0 Then
                                    Dim sIdx As Integer, eIdx As Integer
                                    Dim tStr As String
                                    For sIdx = 0 To UBound(dpSubParts)
                                        tStr = ""
                                        For eIdx = sIdx To UBound(dpSubParts)
                                            If tStr = "" Then tStr = dpSubParts(eIdx) Else tStr = tStr & " " & dpSubParts(eIdx)
                                            If StrComp(tStr, firstTag, vbTextCompare) <> 0 Then
                                                If Not prepDict.Exists(tStr) Then
                                                    If exactDict.Exists(tStr) Then
                                                        If exactDict(tStr) < currentScore Then currentScore = exactDict(tStr)
                                                    End If
                                                End If
                                            End If
                                        Next eIdx
                                    Next sIdx
                                End If
                            End If
                            
                            ' スコア計算が終わったら、元の{}ブロック全体を復元して配列に格納
                            Dim unbackedDPChunk As String
                            unbackedDPChunk = Replace(rawChunk, cleanStr, origDP, 1, -1, vbTextCompare)
                            rawArr(wordCount) = unbackedDPChunk
                            scoreArr(wordCount) = currentScore
                            indexArr(wordCount) = wordCount
                            wordCount = wordCount + 1
                        End If
                        GoTo NextChunk
                    End If
                    
                    ' A. PACK Data Assessment
                    If Left(cleanStr, 7) = "__PACK_" Then
                        If packDict.Exists(cleanStr) Then
                            origPhrase = packDict(cleanStr)(0)
                            origCol = packDict(cleanStr)(1)
                            unbackedChunk = Replace(rawChunk, cleanStr, origPhrase, 1, -1, vbTextCompare)
                            rawArr(wordCount) = unbackedChunk
                            scoreArr(wordCount) = origCol
                            indexArr(wordCount) = wordCount
                            wordCount = wordCount + 1
                        End If
                        GoTo NextChunk
                    End If
                    
                    ' B. Normal Word Assessment
                    If Len(cleanStr) >= 4 And Left(cleanStr, 2) = "__" And Right(cleanStr, 2) = "__" Then
                    Else
                        cleanStr = Replace(cleanStr, "_", " ")
                        cleanStr = Replace(cleanStr, "-", " ")
                    End If
                    cleanStr = Replace(cleanStr, Chr(160), " ")
                    While InStr(cleanStr, "  ") > 0
                        cleanStr = Replace(cleanStr, "  ", " ")
                    Wend
                    cleanStr = Trim(cleanStr)
                    
                    If cleanStr <> "" Then
                        If basePosDict.Exists(cleanStr) Then GoTo NextChunk
                        
                        If loraDict.Exists(cleanStr) Then
                            currentScore = 0.5
                        ElseIf exactDict.Exists(cleanStr) Then
                            currentScore = exactDict(cleanStr)
                        Else
                            Dim subParts() As String
                            subParts = Split(cleanStr, " ")
                            
                            If UBound(subParts) > 0 Then
                                Dim startIdx As Integer, endIdx As Integer
                                Dim tryStr As String
                                
                                For startIdx = 0 To UBound(subParts)
                                    tryStr = ""
                                    For endIdx = startIdx To UBound(subParts)
                                        If tryStr = "" Then
                                            tryStr = subParts(endIdx)
                                        Else
                                            tryStr = tryStr & " " & subParts(endIdx)
                                        End If
                                        
                                        If StrComp(tryStr, cleanStr, vbTextCompare) <> 0 Then
                                            If Not prepDict.Exists(tryStr) Then
                                                If exactDict.Exists(tryStr) Then
                                                    If exactDict(tryStr) < currentScore Then
                                                        currentScore = exactDict(tryStr)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next endIdx
                                Next startIdx
                            End If
                        End If
                        
                        rawArr(wordCount) = rawChunk
                        scoreArr(wordCount) = currentScore
                        indexArr(wordCount) = wordCount
                        wordCount = wordCount + 1
                    End If
                End If
            End If
NextChunk:
        Next i
        
        ' 6. Stable Sort (Bubble Sort)
        Dim tempRaw As String, tempScore As Double, tempIdx As Integer
        For i = 0 To wordCount - 2
            For j = wordCount - 1 To i + 1 Step -1
                If scoreArr(j) < scoreArr(j - 1) Or (scoreArr(j) = scoreArr(j - 1) And indexArr(j) < indexArr(j - 1)) Then
                    tempScore = scoreArr(j): scoreArr(j) = scoreArr(j - 1): scoreArr(j - 1) = tempScore
                    tempRaw = rawArr(j): rawArr(j) = rawArr(j - 1): rawArr(j - 1) = tempRaw
                    tempIdx = indexArr(j): indexArr(j) = indexArr(j - 1): indexArr(j - 1) = tempIdx
                End If
            Next j
        Next i
        
        ' 7. Smart Combination within Paragraphs
        Dim paragraphStr As String
        paragraphStr = ""
        For i = 0 To wordCount - 1
            If paragraphStr = "" Then paragraphStr = rawArr(i) Else paragraphStr = paragraphStr & ", " & rawArr(i)
        Next i
        
        If isBracketParagraph Then paragraphStr = "[" & paragraphStr & "]"
        
        If basePosText <> "" And pIndex = 0 Then
            If paragraphStr <> "" Then paragraphStr = basePosText & ", " & paragraphStr Else paragraphStr = basePosText
        End If
        
        If loraBlock <> "" Then
            If paragraphStr <> "" Then paragraphStr = paragraphStr & " " & loraBlock Else paragraphStr = loraBlock
        End If
        
        ' 8. Final Output Integration
        If finalFullPrompt = "" Then finalFullPrompt = paragraphStr Else finalFullPrompt = finalFullPrompt & vbCrLf & paragraphStr

NextParagraph:
    Next pIndex
    
    ' ========================================================
    ' 9. 最終出力とUI/クリップボード同期
    ' ========================================================
    IsUpdatingBySystem = True
    
    Me.txtMain.Value = finalFullPrompt
    AccumulatedPrompt = finalFullPrompt
    
    With Me.txtMain
        .SetFocus
        .SelStart = 0
        .SelLength = Len(finalFullPrompt)
    End With
    
    DoEvents
    
    Call SetClipboardText(finalFullPrompt)
    
    LastCopiedPrompt = finalFullPrompt
    IsUpdatingBySystem = False
    
    Application.StatusBar = "● Prompts successfully sorted and copied to clipboard."
End Sub
' =========================================================================
' [Mobile Utility] 選択したアイテムをリストから削除する
' =========================================================================
Private Sub btnRemoveMobiItem_Click()
    Dim listIdx As Long
    Dim i As Long
    Dim tempCol0 As New Collection
    Dim tempCol1 As New Collection
    
    ' 1. 現在選択されている行のインデックスを取得
    listIdx = Me.lstMobileMemo.ListIndex
    
    ' 2. 未選択の場合は警告を出して終了
    If listIdx = -1 Then
        MsgBox "削除するアイテムを選択してください。" & vbCrLf & _
               "Please select an item to remove.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 3. 誤操作防止の確認ダイアログ
    If MsgBox("選択したアイテムを削除しますか？" & vbCrLf & _
              "Are you sure you want to remove this item?", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then
        Exit Sub
    End If
    
    ' =====================================================================
    ' 4. バグ回避の最終手段：RemoveItemを使わず、リストを再構築する
    ' =====================================================================
    ' ① 削除対象「以外」のデータをコレクションに一時退避
    For i = 0 To Me.lstMobileMemo.ListCount - 1
        If i <> listIdx Then
            ' ★修正：VBA特有の Null エラーを防ぐため、末尾に "" を付与して確実に文字列キャストする
            tempCol0.Add Me.lstMobileMemo.List(i, 0) & ""
            tempCol1.Add Me.lstMobileMemo.List(i, 1) & ""
        End If
    Next i
    
    ' ② リストボックスを完全にリセット（RemoveItemのバグを回避）
    Me.lstMobileMemo.Clear
    
    ' ③ 退避したデータを使ってリストボックスを再構築
    For i = 1 To tempCol0.Count
        Me.lstMobileMemo.AddItem tempCol0(i)
        Me.lstMobileMemo.List(Me.lstMobileMemo.ListCount - 1, 1) = tempCol1(i)
    Next i
    
    ' 5. 連動して詳細テキストボックスの表示もクリア
    Me.txtMobileMemo.text = ""
    
    ' =====================================================================
    ' 6. ★追加：削除結果を母艦（ConfigData）へ自動セーブ（再起動時の復活バグを防止）
    ' =====================================================================
    If Not ConfigData Is Nothing Then
        Dim newUrlStock As Collection
        Set newUrlStock = New Collection
        
        For i = 0 To Me.lstMobileMemo.ListCount - 1
            ' 安全に文字列キャストして結合
            newUrlStock.Add (Me.lstMobileMemo.List(i, 0) & "") & "<|>" & (Me.lstMobileMemo.List(i, 1) & "")
        Next i
        
        If ConfigData.Exists("Mobile_URL_Stock") Then
            Set ConfigData("Mobile_URL_Stock") = newUrlStock
        Else
            ConfigData.Add "Mobile_URL_Stock", newUrlStock
        End If
        
        Call SaveConfigJSON
    End If
End Sub

'-------------------------------------------------------------------------
' [Save As Preset] ボタン（JSON対応：重複・上書き対応版 ＋ ModelName完全対応）
'-------------------------------------------------------------------------
Private Sub btnSaveAsPresetLoRA_Click()
    Dim presetName As String
    Dim i As Long
    Dim listAlias As String
    Dim sysHash As String
    Dim loraModelName As String ' ★追加：ModelName格納用
    Dim targetName As String    ' ★追加：Wrap設定を反映した出力名
    Dim overwriteConfirm As VbMsgBoxResult
    
    Dim presetsDict As Object
    Dim presetItems As Collection
    Dim itemDict As Object
    Dim loraList As Object
    Dim loraItem As Object
    
    ' 1. バリデーション：リストが空なら保存しない
    If Me.lstSelectedLoRA.ListCount = 0 Then
        MsgBox "保存するLoRAがリストにありません。" & vbCrLf & _
               "No LoRA items in the list to save.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' ★追加対策1：プレビューが生成されているか確認（Wrap前対策）
    If Trim(Me.txtLoRAPreview.text) = "" Then
        MsgBox "プレビューが生成されていません。「Wrap LoRA!」を押して内容を確認してから保存してください。" & vbCrLf & _
               "Please click 'Wrap LoRA!' to generate the preview before saving.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 2. ユーザーにプリセット名を入力してもらう
    presetName = InputBox("この組み合わせに名前を付けてください。" & vbCrLf & _
                          "Please enter a name for this preset.", "Save Preset")
    
    ' キャンセルまたは空文字なら終了
    presetName = Trim(presetName)
    If presetName = "" Then Exit Sub
    
    ' 3. JSON用の準備
    If ConfigData Is Nothing Then Exit Sub
    
    ' プリセット用Dictionaryが無ければ作成
    If Not ConfigData.Exists("LoRA_Presets") Then
        Set ConfigData("LoRA_Presets") = CreateObject("Scripting.Dictionary")
    End If
    Set presetsDict = ConfigData("LoRA_Presets")
    
    ' ★追加対策2：重複チェックと上書き確認
    If presetsDict.Exists(presetName) Then
        overwriteConfirm = MsgBox("プリセット '" & presetName & "' は既に存在します。上書きしますか？" & vbCrLf & _
                                  "Preset '" & presetName & "' already exists. Overwrite?", _
                                  vbYesNo + vbExclamation, APP_NAME)
                                  
        If overwriteConfirm = vbNo Then Exit Sub ' いいえなら中止
        
        ' 上書き（はい）の場合は、Dictionaryのキーごと上書きするので削除処理は不要（自動で消え去ります）
    End If
    
    ' 4. リスト内の全項目をループしてJSONのCollectionにまとめる
    Set presetItems = New Collection
    
    For i = 0 To Me.lstSelectedLoRA.ListCount - 1
        listAlias = Me.lstSelectedLoRA.List(i, 2) ' 3列目のAlias
        sysHash = "ERROR_NOT_FOUND"
        loraModelName = "" ' ★初期化
        
        ' ゴーストシートのMatchの代わりに、メモリ上のLoRA_ListからHashを検索
        If ConfigData.Exists("LoRA_List") Then
            Set loraList = ConfigData("LoRA_List")
            For Each loraItem In loraList
                If loraItem("Alias") = listAlias Then
                    sysHash = loraItem("Hash")
                    ' ▼▼▼ 追加：ModelName（正式名称）が存在すれば取得 ▼▼▼
                    If loraItem.Exists("ModelName") Then
                        loraModelName = loraItem("ModelName")
                    End If
                    Exit For
                End If
            Next loraItem
        End If
        
        ' ▼▼▼ 追加：Wrapボタンと同じロジックでTargetName（実際に出力される名前）を決定 ▼▼▼
        If Me.opbWrapName.Value = True And Trim(loraModelName) <> "" Then
            targetName = Trim(loraModelName)
        Else
            If Trim(sysHash) <> "" Then targetName = Trim(sysHash) Else targetName = Trim(listAlias)
        End If
        
        ' 保存用アイテム（Dictionary）を作成
        Set itemDict = CreateObject("Scripting.Dictionary")
        itemDict.Add "Alias", listAlias                           ' A: 表示用としてAliasも保持しておく
        itemDict.Add "Hash", sysHash                              ' B: ハッシュ値
        itemDict.Add "ModelName", loraModelName                   ' ★追加: 正式名称(システム名)
        itemDict.Add "TargetName", targetName                     ' ★追加: Wrap設定を反映した出力名
        itemDict.Add "Strength", Me.lstSelectedLoRA.List(i, 3)    ' C: LoRA強度
        itemDict.Add "Trigger", Me.lstSelectedLoRA.List(i, 1)     ' D: 選択トリガー（生）
        
        ' =================================================================
        ' ★修正: 重みは既に Trigger の文字列内に (tag:1.3) として完成しているため、
        ' コンボボックスの値を拾わず、JSON互換性維持のため常にニュートラルな "1.0" を保存する
        ' =================================================================
        itemDict.Add "TriggerWeight", "1.0"
        
        ' コレクションに追加
        presetItems.Add itemDict
    Next i
    
    ' Dictionaryにプリセット名で紐づけて格納し、物理ファイルに保存
    If presetsDict.Exists(presetName) Then
        Set presetsDict(presetName) = presetItems ' 上書き
    Else
        presetsDict.Add presetName, presetItems   ' 新規登録
    End If
    
    Call SaveConfigJSON
    
    ' 5. コンボボックスのリストを最新状態に更新する
    Call RefreshPresetList
    Me.cmbPresetList.Value = presetName ' 保存した名前を選択状態にする
    
    ' 6. 完了通知
    MsgBox "プリセット '" & presetName & "' を保存しました。" & vbCrLf & _
           "Preset '" & presetName & "' has been saved.", vbInformation, APP_NAME
End Sub

'-------------------------------------------------------------------------
' 保存済みプリセット名をコンボボックスにロードする（JSON対応版）
'-------------------------------------------------------------------------
Private Sub RefreshPresetList()
    Dim k As Variant
    Dim presetsDict As Object
    
    Me.cmbPresetList.Clear
    
    ' JSONデータが存在するかチェック
    If ConfigData Is Nothing Then Exit Sub
    If Not ConfigData.Exists("LoRA_Presets") Then Exit Sub
    
    Set presetsDict = ConfigData("LoRA_Presets")
    
    ' プリセット名（キー）を取り出してコンボボックスに流し込む
    For Each k In presetsDict.Keys
        Me.cmbPresetList.AddItem k
    Next k
End Sub

'-------------------------------------------------------------------------
' 共通処理：ネガティブプリセットのコンボボックス再構築（Default内包版）
'-------------------------------------------------------------------------
Private Sub RefreshNegaPresetList()
    Dim wsPreset As Worksheet
    Dim rngNameBase As Range
    Dim startRow As Long
    Dim lastRow As Long
    Dim nameCol As Long
    Dim i As Long
    
    On Error Resume Next
    Set wsPreset = ThisWorkbook.Sheets("MyNegativePreset")
    Set rngNameBase = wsPreset.Range("NegativePresetName")
    On Error GoTo 0
    
    If wsPreset Is Nothing Or rngNameBase Is Nothing Then Exit Sub
    
    ' コンボボックスを一度空にする
    Me.cmbSelectNegaPreset.Clear
    
    ' ★仕様追加：リストの最上段に「Default」を登録
    Me.cmbSelectNegaPreset.AddItem "Default"
    
    nameCol = rngNameBase.Column
    startRow = rngNameBase.Row + 1
    lastRow = wsPreset.Cells(wsPreset.Rows.Count, nameCol).End(xlUp).Row
    
    ' データが存在すればリストに追加
    If lastRow >= startRow Then
        For i = startRow To lastRow
            Dim currentName As String
            currentName = Trim(CStr(wsPreset.Cells(i, nameCol).Value))
            If currentName <> "" Then
                Me.cmbSelectNegaPreset.AddItem currentName
            End If
        Next i
    End If
End Sub
'=========================================================================
' 【共通ユーティリティ】既存テキストと比較し、未知の要素（単語/タグ）だけを抽出して繋ぐ関数
' （InStrの揺らぎを排除した、完全配列比較による鉄壁の重複ブロック版）
'=========================================================================
Private Function ExtractOnlyNewElements(ByVal sourceText As String, ByVal targetText As String) As String
    Dim sourceElements() As String
    Dim targetElements() As String
    Dim i As Integer, j As Integer
    Dim srcEl As String, tgtEl As String
    Dim resultText As String
    Dim isDup As Boolean
    
    resultText = ""
    
    ' --- 1. Target (Cockpit/Fav) のテキストを要素ごとに分解 ---
    ' LoRAタグ同士の間のスペースをカンマに変換して分割可能にする
    targetText = Replace(targetText, "> <", ">, <")
    ' 改行もカンマ扱いにする
    targetText = Replace(targetText, vbCrLf, ",")
    targetText = Replace(targetText, vbCr, ",")
    targetText = Replace(targetText, vbLf, ",")
    
    targetElements = Split(targetText, ",")
    
    ' --- 2. Source (プレビュー) のテキストを要素ごとに分解 ---
    sourceText = Replace(sourceText, "> <", ">, <")
    sourceElements = Split(sourceText, ",")
    
    ' --- 3. Source の各要素が Target に存在するかチェック ---
    For i = 0 To UBound(sourceElements)
        srcEl = Trim(sourceElements(i))
        
        If srcEl <> "" Then
            isDup = False
            
            For j = 0 To UBound(targetElements)
                tgtEl = Trim(targetElements(j))
                
                ' 大文字小文字を無視して完全一致チェック（スペースやカンマの揺らぎを一切受けない）
                If StrComp(srcEl, tgtEl, vbTextCompare) = 0 Then
                    isDup = True
                    Exit For
                End If
            Next j
            
            ' 重複していなければ結果に追加
            If Not isDup Then
                If resultText = "" Then
                    resultText = srcEl
                Else
                    ' 前がLoRAで後ろもLoRAならスペース、それ以外はカンマにするスマート結合
                    If Right(resultText, 1) = ">" And Left(srcEl, 1) = "<" Then
                        resultText = resultText & " " & srcEl
                    Else
                        resultText = resultText & ", " & srcEl
                    End If
                End If
            End If
        End If
    Next i
    
    ExtractOnlyNewElements = resultText
End Function

'-------------------------------------------------------------------------
' ★新設 [Get LoRA Negative] ボタン：選択中のLoRAから固有ネガティブを抽出しスマート送信
'-------------------------------------------------------------------------
Private Sub btnGetLoRANegative_Click()
    Dim i As Integer
    Dim listAlias As String
    Dim loraList As Object
    Dim loraItem As Object
    Dim rawNegaPrompt As String
    Dim combinedNega As String
    Dim currentNegaMain As String
    Dim appendedNega As String
    
    ' 1. バリデーション：カートに何もなければ終了
    If Me.lstSelectedLoRA.ListCount = 0 Then Exit Sub
    
    combinedNega = ""
    
    ' 2. カート内を走査し、選択（ハイライト）されているLoRAの固有ネガティブを収集
    For i = 0 To Me.lstSelectedLoRA.ListCount - 1
        If Me.lstSelectedLoRA.Selected(i) = True Then
            listAlias = Me.lstSelectedLoRA.List(i, 2) ' 既存のAlias（3列目）を取得
            
            ' メモリ上のLoRA_Listから該当する固有ネガティブ（Negative）を検索
            If Not ConfigData Is Nothing Then
                If ConfigData.Exists("LoRA_List") Then
                    Set loraList = ConfigData("LoRA_List")
                    For Each loraItem In loraList
                        If loraItem("Alias") = listAlias Then
                            ' JSON内に "Negative" キーが存在し、空でなければ取得
                            If loraItem.Exists("Negative") Then
                                rawNegaPrompt = Trim(loraItem("Negative"))
                                If rawNegaPrompt <> "" Then
                                    If combinedNega = "" Then
                                        combinedNega = rawNegaPrompt
                                    Else
                                        combinedNega = combinedNega & ", " & rawNegaPrompt
                                    End If
                                End If
                            End If
                            Exit For ' 見つかったらこのLoRAの検索は終了して次へ
                        End If
                    Next loraItem
                End If
            End If
        End If
    Next i
    
    ' 選択されたLoRAに固有ネガティブが1つも設定されていなかった場合
    If combinedNega = "" Then
        MsgBox "選択されたLoRAには固有のネガティブプロンプトが登録されていません。" & vbCrLf & _
               "No inherent negative prompts are registered for the selected LoRA(s).", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' 3. 鉄壁の重複排除（配列比較関数を流用）
    currentNegaMain = Trim(Me.txtYourNegative.text)
    appendedNega = ExtractOnlyNewElements(combinedNega, currentNegaMain)
    
    If appendedNega = "" Then
        MsgBox "このLoRAの固有ネガティブは既にすべてベースネガティブ欄に含まれています。" & vbCrLf & _
               "The negative prompts are already included in the Negative field.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' 4. ネガティブテキストボックスへスマート挿入
    If currentNegaMain = "" Then
        Me.txtYourNegative.text = appendedNega
    Else
        If Right(currentNegaMain, 1) = "," Then
            Me.txtYourNegative.text = currentNegaMain & " " & appendedNega
        Else
            Me.txtYourNegative.text = currentNegaMain & ", " & appendedNega
        End If
    End If
    
    ' 5. 適用中リストボックス側（lstApplyNegative）も自動更新するために変更イベントを擬似キック
    Call SyncNegativeSheet
    Call UpdateNegativeUIState
    
    ' 処理が終わったらカートの選択状態を解除してQoL向上
    For i = 0 To Me.lstSelectedLoRA.ListCount - 1
        Me.lstSelectedLoRA.Selected(i) = False
    Next i
    
' ★追加：ネガティブタブ（インデックス2）に移動する
    Me.MainWindow.Value = 2
    DoEvents ' 画面描画の完了を確実に待つ安全装置
    
    ' ★修正：YourNegativeにフォーカスを合わせる
    Me.txtYourNegative.SetFocus

    Application.StatusBar = "● Inherent negative prompt(s) injected into Negative field."
End Sub
'-------------------------------------------------------------------------
' [Send to Cockpit] ボタン：LoRAをコクピット(txtMain)の末尾へスマート挿入（完全版）
'-------------------------------------------------------------------------
Private Sub btnSendLoRAtoCockpit_Click()
    Dim newLora As String
    Dim currentMain As String
    Dim appendedPrompt As String
    
    newLora = Trim(Me.txtLoRAPreview.text)
    If newLora = "" Then Exit Sub
    
    currentMain = Trim(Me.txtMain.text)
    
    ' =========================================================================
    ' 1. 改行や「スペース区切りのLoRAタグ」をカンマに正規化
    ' =========================================================================
    newLora = Replace(newLora, vbCrLf, ",")
    newLora = Replace(newLora, vbCr, ",")
    newLora = Replace(newLora, vbLf, ",")
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    ' <lora:...> の直後にスペースを挟んで別の <lora:...> が続く場合、カンマで分断する
    regEx.Pattern = ">\s*<lora:"
    newLora = regEx.Replace(newLora, ">, <lora:")
    
    While InStr(newLora, ",,") > 0
        newLora = Replace(newLora, ",,", ",")
    Wend
    
    Dim rawElements() As String
    Dim cleanElement As String
    Dim i As Integer
    Dim parts() As String
    
    rawElements = Split(newLora, ",")
    appendedPrompt = ""
    
    ' =========================================================================
    ' 2. 個別解体 ＆ 鉄壁の重複チェック
    ' =========================================================================
    For i = 0 To UBound(rawElements)
        cleanElement = Trim(rawElements(i))
        If cleanElement <> "" Then
            Dim isDuplicate As Boolean
            isDuplicate = False
            
            If LCase(Left(cleanElement, 6)) = "<lora:" Then
                ' ★修正：ウェイト変更や大文字小文字、スペース混入対策の正規表現マッチ
                Dim loraName As String
                parts = Split(cleanElement, ":")
                If UBound(parts) >= 1 Then
                    loraName = parts(1) ' コロンで区切った2番目がベース名
                Else
                    loraName = Replace(Replace(cleanElement, "<lora:", ""), ">", "")
                End If
                
                ' ベース名の正規表現エスケープ処理
                Dim escLoraName As String
                escLoraName = loraName
                escLoraName = Replace(escLoraName, "\", "\\")
                escLoraName = Replace(escLoraName, ".", "\.")
                escLoraName = Replace(escLoraName, "^", "\^")
                escLoraName = Replace(escLoraName, "$", "\$")
                escLoraName = Replace(escLoraName, "*", "\*")
                escLoraName = Replace(escLoraName, "+", "\+")
                escLoraName = Replace(escLoraName, "?", "\?")
                escLoraName = Replace(escLoraName, "(", "\(")
                escLoraName = Replace(escLoraName, ")", "\)")
                escLoraName = Replace(escLoraName, "[", "\[")
                escLoraName = Replace(escLoraName, "]", "\]")
                escLoraName = Replace(escLoraName, "{", "\{")
                escLoraName = Replace(escLoraName, "}", "\}")
                escLoraName = Replace(escLoraName, "|", "\|")
                
                ' <lora: ＋ (空白許容) ＋ ベース名 ＋ (空白許容) ＋ : か >
                regEx.Pattern = "<lora:\s*" & escLoraName & "\s*[:>]"
                If regEx.Test(currentMain) Then
                    isDuplicate = True
                End If
            Else
                ' ★修正：トリガーワードの厳密な境界チェック（カッコやウェイト指定のコロンも境界として判定）
                Dim escWord As String
                escWord = cleanElement
                escWord = Replace(escWord, "\", "\\")
                escWord = Replace(escWord, ".", "\.")
                escWord = Replace(escWord, "^", "\^")
                escWord = Replace(escWord, "$", "\$")
                escWord = Replace(escWord, "*", "\*")
                escWord = Replace(escWord, "+", "\+")
                escWord = Replace(escWord, "?", "\?")
                escWord = Replace(escWord, "(", "\(")
                escWord = Replace(escWord, ")", "\)")
                escWord = Replace(escWord, "[", "\[")
                escWord = Replace(escWord, "]", "\]")
                escWord = Replace(escWord, "{", "\{")
                escWord = Replace(escWord, "}", "\}")
                escWord = Replace(escWord, "|", "\|")
                
                ' 前境界: 行頭, カンマ, 空白, 開きカッコ (, [, {
                ' 後境界: 行末, カンマ, 空白, 閉じカッコ ), ], }, または コロン :
                regEx.Pattern = "(^|[,\s\(\[\{])" & escWord & "(?=[,\s\)\]\}:]|$)"
                If regEx.Test(currentMain) Then
                    isDuplicate = True
                End If
            End If
            
            ' =========================================================================
            ' ★ バグ修正：追加する要素がLoRAの場合は、直前のタグに関わらずカンマを打たない
            ' =========================================================================
            If Not isDuplicate Then
                If appendedPrompt = "" Then
                    appendedPrompt = cleanElement
                Else
                    If LCase(Left(cleanElement, 6)) = "<lora:" Then
                        appendedPrompt = appendedPrompt & " " & cleanElement
                    Else
                        appendedPrompt = appendedPrompt & ", " & cleanElement
                    End If
                End If
            End If
        End If
    Next i
    
    ' =========================================================================
    ' 3. コックピットへの最終書き出し
    ' =========================================================================
    If appendedPrompt = "" Then
        MsgBox "このLoRA（またはトリガーワード）は既にコクピットにすべて含まれています。" & vbCrLf & _
               "These items are already included in the Cockpit.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    Call SaveHistory(currentMain)
    
    IsUpdatingBySystem = True
    
    ' ★ バグ修正：追記分全体がLoRAから始まる場合も、カンマを打たずにスペースで結合する
    If currentMain = "" Then
        Me.txtMain.text = appendedPrompt
    Else
        If Right(currentMain, 1) = "," Then
            Me.txtMain.text = currentMain & " " & appendedPrompt
        ElseIf LCase(Left(appendedPrompt, 6)) = "<lora:" Then
            Me.txtMain.text = currentMain & " " & appendedPrompt
        Else
            Me.txtMain.text = currentMain & ", " & appendedPrompt
        End If
    End If
    
    IsUpdatingBySystem = False
    
    On Error Resume Next
    Me.MainWindow.Value = Me.MainWindow.Pages("pgCockpit").Index
    Me.txtMain.SetFocus
    On Error GoTo 0
    
    Application.StatusBar = "● LoRA sent to the end of Cockpit and focus moved."
End Sub
'-------------------------------------------------------------------------
' [Send to Favorites] ボタン（完全版）
'-------------------------------------------------------------------------
Private Sub btnSendLoRAtoFav_Click()
    Dim loraPrompt As String
    Dim currentFav As String
    Dim appendedPrompt As String
    
    loraPrompt = Trim(Me.txtLoRAPreview.text)
    If loraPrompt = "" Then
        MsgBox "送信するLoRAプロンプトがありません。" & vbCrLf & _
               "No LoRA prompt available to send.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    currentFav = Trim(Me.txtFav.text)
    
    loraPrompt = Replace(loraPrompt, vbCrLf, ",")
    loraPrompt = Replace(loraPrompt, vbCr, ",")
    loraPrompt = Replace(loraPrompt, vbLf, ",")
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    regEx.Pattern = ">\s*<lora:"
    loraPrompt = regEx.Replace(loraPrompt, ">, <lora:")
    
    While InStr(loraPrompt, ",,") > 0
        loraPrompt = Replace(loraPrompt, ",,", ",")
    Wend
    
    Dim rawElements() As String
    Dim cleanElement As String
    Dim i As Integer
    Dim parts() As String
    
    rawElements = Split(loraPrompt, ",")
    appendedPrompt = ""
    
    For i = 0 To UBound(rawElements)
        cleanElement = Trim(rawElements(i))
        If cleanElement <> "" Then
            Dim isDuplicate As Boolean
            isDuplicate = False
            
            If LCase(Left(cleanElement, 6)) = "<lora:" Then
                ' ★修正：ウェイト変更や大文字小文字、スペース混入対策の正規表現マッチ
                Dim loraName As String
                parts = Split(cleanElement, ":")
                If UBound(parts) >= 1 Then
                    loraName = parts(1) ' コロンで区切った2番目がベース名
                Else
                    loraName = Replace(Replace(cleanElement, "<lora:", ""), ">", "")
                End If
                
                Dim escLoraName As String
                escLoraName = loraName
                escLoraName = Replace(escLoraName, "\", "\\")
                escLoraName = Replace(escLoraName, ".", "\.")
                escLoraName = Replace(escLoraName, "^", "\^")
                escLoraName = Replace(escLoraName, "$", "\$")
                escLoraName = Replace(escLoraName, "*", "\*")
                escLoraName = Replace(escLoraName, "+", "\+")
                escLoraName = Replace(escLoraName, "?", "\?")
                escLoraName = Replace(escLoraName, "(", "\(")
                escLoraName = Replace(escLoraName, ")", "\)")
                escLoraName = Replace(escLoraName, "[", "\[")
                escLoraName = Replace(escLoraName, "]", "\]")
                escLoraName = Replace(escLoraName, "{", "\{")
                escLoraName = Replace(escLoraName, "}", "\}")
                escLoraName = Replace(escLoraName, "|", "\|")
                
                regEx.Pattern = "<lora:\s*" & escLoraName & "\s*[:>]"
                If regEx.Test(currentFav) Then
                    isDuplicate = True
                End If
            Else
                ' ★修正：トリガーワードの厳密な境界チェック（カッコやウェイト指定のコロンも境界として判定）
                Dim escWord As String
                escWord = cleanElement
                escWord = Replace(escWord, "\", "\\")
                escWord = Replace(escWord, ".", "\.")
                escWord = Replace(escWord, "^", "\^")
                escWord = Replace(escWord, "$", "\$")
                escWord = Replace(escWord, "*", "\*")
                escWord = Replace(escWord, "+", "\+")
                escWord = Replace(escWord, "?", "\?")
                escWord = Replace(escWord, "(", "\(")
                escWord = Replace(escWord, ")", "\)")
                escWord = Replace(escWord, "[", "\[")
                escWord = Replace(escWord, "]", "\]")
                escWord = Replace(escWord, "{", "\{")
                escWord = Replace(escWord, "}", "\}")
                escWord = Replace(escWord, "|", "\|")
                
                ' 前境界: 行頭, カンマ, 空白, 開きカッコ (, [, {
                ' 後境界: 行末, カンマ, 空白, 閉じカッコ ), ], }, または コロン :
                regEx.Pattern = "(^|[,\s\(\[\{])" & escWord & "(?=[,\s\)\]\}:]|$)"
                If regEx.Test(currentFav) Then
                    isDuplicate = True
                End If
            End If
            
            ' =========================================================================
            ' ★ バグ修正：追加する要素がLoRAの場合は、直前のタグに関わらずカンマを打たない
            ' =========================================================================
            If Not isDuplicate Then
                If appendedPrompt = "" Then
                    appendedPrompt = cleanElement
                Else
                    If LCase(Left(cleanElement, 6)) = "<lora:" Then
                        appendedPrompt = appendedPrompt & " " & cleanElement
                    Else
                        appendedPrompt = appendedPrompt & ", " & cleanElement
                    End If
                End If
            End If
        End If
    Next i
    
    If appendedPrompt = "" Then
        MsgBox "このLoRA（またはトリガーワード）は既にお気に入り（Favorites）にすべて含まれています。" & vbCrLf & _
               "These items already exist in the Favorites.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    Call SaveHistory(Me.txtMain.text)
    
    ' ★ バグ修正：追記分全体がLoRAから始まる場合も、カンマを打たずにスペース（改行）で結合する
    If currentFav = "" Then
        Me.txtFav.text = appendedPrompt
    Else
        Dim lastChar As String
        lastChar = Right(currentFav, 1)
        
        If lastChar = vbLf Or lastChar = vbCr Then
            Me.txtFav.text = currentFav & appendedPrompt
        ElseIf lastChar = "," Then
            Me.txtFav.text = currentFav & " " & vbCrLf & appendedPrompt
        ElseIf LCase(Left(appendedPrompt, 6)) = "<lora:" Then
            Me.txtFav.text = currentFav & " " & vbCrLf & appendedPrompt
        Else
            Me.txtFav.text = currentFav & ", " & vbCrLf & appendedPrompt
        End If
    End If
    
    On Error Resume Next
    Me.MainWindow.Value = 4
    On Error GoTo 0
    
    Me.txtFav.SetFocus
    Application.StatusBar = "● Sent to Favorites: " & Left(appendedPrompt, 50) & "..."
End Sub

Private Sub btnUndoFav_Click()

    Call PerformFavUndo
    Call SyncFavoritesManager ' ★追加：非同期ズレ防止

End Sub

'-------------------------------------------------------------------------
' [ btnWrapBracket ] ボタン：選択した文字列をDynamicPrompt用に {} でくくる（トグル式）
' ※正規表現PACKエンジン搭載：データベースに登録された「カンマを含むタグ」を保護します
'-------------------------------------------------------------------------
Private Sub btnWrapBracket_Click()
    Dim targetTextBox As MSForms.TextBox
    Set targetTextBox = Me.txtMain
    
    Dim selectedStr As String
    Dim oldStart As Long
    
    ' PACK用の変数
    Dim wsMain As Worksheet
    Dim mainData As Variant
    Dim r As Long, c As Long
    Dim origPhrase As String
    Dim pCount As Long
    Dim phraseList() As String
    
    Dim regEx As Object, matches As Object, m As Object
    Dim packDict As Object
    Dim packIdx As Integer
    Dim i As Long
    Dim searchPattern As String
    Dim packKey As String
    Dim k As Variant
    
    ' テキストボックスにフォーカスを戻す
    targetTextBox.SetFocus
    
    ' 選択されているテキストがあるか確認
    If targetTextBox.SelLength > 0 Then
        selectedStr = targetTextBox.SelText
        oldStart = targetTextBox.SelStart
        
        ' ========================================================
        ' 1. カンマを含むタグをデータベースから収集（キャッシュ作成）
        ' ========================================================
        Set wsMain = ThisWorkbook.Sheets("KENZEN SeaArt Helper")
        ' 最新のデータベース範囲（20行目スタート、B列～CB列）
        mainData = wsMain.Range("B20:CB500").Value
        
        pCount = 0
        ReDim phraseList(1 To 2000)
        
        For c = 1 To UBound(mainData, 2) Step 2
            ' H列（配列のインデックス7）はBREAK構文前提のため除外
            If c = 7 Then GoTo NextMainCol
            
            For r = 1 To UBound(mainData, 1)
                If Not IsError(mainData(r, c)) And Not IsEmpty(mainData(r, c)) Then
                    origPhrase = Trim(CStr(mainData(r, c)))
                    If origPhrase <> "" Then
                        ' カンマを含むフレーズのみをリストアップ
                        If InStr(origPhrase, ",") > 0 Then
                            pCount = pCount + 1
                            If pCount > UBound(phraseList) Then
                                ReDim Preserve phraseList(1 To pCount + 1000)
                            End If
                            phraseList(pCount) = origPhrase
                        End If
                    End If
                End If
            Next r
NextMainCol:
        Next c
        
        ' 長いフレーズから順にソート（部分一致の誤爆を防ぐため）
        If pCount > 0 Then
            Dim j As Long, tmpS As String
            For i = 1 To pCount - 1
                For j = pCount To i + 1 Step -1
                    If Len(phraseList(j)) > Len(phraseList(j - 1)) Then
                        tmpS = phraseList(j): phraseList(j) = phraseList(j - 1): phraseList(j - 1) = tmpS
                    End If
                Next j
            Next i
        End If

        ' ========================================================
        ' 2. 選択文字列内のセットタグを保護（PACK）
        ' ========================================================
        Set regEx = CreateObject("VBScript.RegExp")
        Set packDict = CreateObject("Scripting.Dictionary")
        packIdx = 1
        
        If pCount > 0 Then
            For i = 1 To pCount
                searchPattern = phraseList(i)
                
                ' 正規表現の特殊文字をエスケープ
                searchPattern = Replace(searchPattern, "\", "\\")
                searchPattern = Replace(searchPattern, "(", "\(")
                searchPattern = Replace(searchPattern, ")", "\)")
                searchPattern = Replace(searchPattern, "[", "\[")
                searchPattern = Replace(searchPattern, "]", "\]")
                searchPattern = Replace(searchPattern, "+", "\+")
                searchPattern = Replace(searchPattern, ".", "\.")
                searchPattern = Replace(searchPattern, "?", "\?")
                searchPattern = Replace(searchPattern, "*", "\*")
                searchPattern = Replace(searchPattern, "|", "\|")
                searchPattern = Replace(searchPattern, "{", "\{")
                searchPattern = Replace(searchPattern, "}", "\}")
                searchPattern = Replace(searchPattern, "^", "\^")
                searchPattern = Replace(searchPattern, "$", "\$")
                
                ' スペース、ハイフン、アンダースコアを同一視
                searchPattern = Replace(searchPattern, " ", "###SP###")
                searchPattern = Replace(searchPattern, "_", "###SP###")
                searchPattern = Replace(searchPattern, "-", "###SP###")
                searchPattern = Replace(searchPattern, "###SP###", "[-_ ]")
                
                ' カンマ前後のスペースブレを許容
                searchPattern = Replace(searchPattern, ",", "\s*,\s*")
                
                regEx.Pattern = searchPattern
                regEx.Global = True
                regEx.IgnoreCase = True
                
                If regEx.Test(selectedStr) Then
                    Set matches = regEx.Execute(selectedStr)
                    For Each m In matches
                        packKey = "__PACK_" & packIdx & "__"
                        ' 一致した部分を保護キーに置換
                        selectedStr = Replace(selectedStr, m.Value, packKey, 1, 1, vbTextCompare)
                        packDict.Add packKey, phraseList(i)
                        packIdx = packIdx + 1
                    Next m
                End If
            Next i
        End If
        
        ' ========================================================
        ' 3. 置換と復元（トグル処理）
        ' ========================================================
        Dim finalStr As String
        
        ' すでに {} でくくられている場合は解除（トグル機能）
        If Left(selectedStr, 1) = "{" And Right(selectedStr, 1) = "}" Then
            ' 中身を取り出す
            finalStr = Mid(selectedStr, 2, Len(selectedStr) - 2)
            
            ' パイプをカンマに戻す
            finalStr = Replace(finalStr, " | ", ", ")
            finalStr = Replace(finalStr, "|", ", ")
            
            ' PACKされたタグを復元（UNPACK）
            For Each k In packDict.Keys
                finalStr = Replace(finalStr, k, packDict(k), 1, -1, vbTextCompare)
            Next k
            
        Else
            ' --- 新規で適用する場合 ---
            
            ' むき出しになっている区切りカンマをパイプに置換
            finalStr = Replace(selectedStr, ", ", " | ")
            finalStr = Replace(finalStr, ",", " | ") ' スペースなしのカンマも念のため対応
            
            ' PACKされたタグを復元（UNPACK）
            For Each k In packDict.Keys
                finalStr = Replace(finalStr, k, packDict(k), 1, -1, vbTextCompare)
            Next k
            
            ' 全体を {} で括る
            finalStr = "{" & finalStr & "}"
        End If
        
        ' ========================================================
        ' 4. テキストボックスへの反映とUI更新
        ' ========================================================
        targetTextBox.SelText = finalStr
        
        ' 処理後も全体を選択状態にしておく
        targetTextBox.SelStart = oldStart
        targetTextBox.SelLength = Len(finalStr)
        
        ' 変更があったのでフラグを立てる（必要に応じて）
        'IsDirty = True
        
    Else
        ' 何も選択されていない場合のアラート
        MsgBox "テキストボックス内で、くくりたい文字列を選択してください。" & vbCrLf & _
               "Please select the text to wrap in the text box.", vbExclamation, APP_NAME
    End If
End Sub
'-------------------------------------------------------------------------
' [Wrap] ボタン：出力切替 ＆ プレビュー重複ブロック ＆ アラートダイアログ完備版
'-------------------------------------------------------------------------
Private Sub btnWrapLoRA_Click()
    Dim i As Integer, j As Integer
    Dim listAlias As String, listStrength As String
    Dim sysHash As String, loraModelName As String, targetName As String
    Dim triggerRaw As String
    Dim triggerArray() As String
    Dim weightVal As String
    
    Dim addedTriggers As String
    Dim addedLoRAs As String
    Dim appendedPrompt As String
    Dim currentPreview As String
    
    Dim loraList As Object
    Dim loraItem As Object
    
    ' ★追加：重複アラート用のメッセージ格納変数
    Dim dupMessages As String
    dupMessages = ""
    
    ' ★既存の重み付けを綺麗に剥がすための正規表現
    Dim reClean As Object
    Set reClean = CreateObject("VBScript.RegExp")
    reClean.Global = True
    reClean.IgnoreCase = True
    reClean.Pattern = "^\(\s*(.+?)\s*:\s*\d+(\.\d+)?\s*\)$"
    
    ' 1. バリデーション
    If Me.lstSelectedLoRA.ListCount = 0 Then
        MsgBox "適用するLoRAがリストにありません。" & vbCrLf & _
               "No LoRA items in the list to wrap.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 何も選択されていない場合は警告を出して終了
    Dim isSelectedAny As Boolean
    isSelectedAny = False
    For i = 0 To Me.lstSelectedLoRA.ListCount - 1
        If Me.lstSelectedLoRA.Selected(i) = True Then
            isSelectedAny = True
            Exit For
        End If
    Next i
    
    If Not isSelectedAny Then
        MsgBox "リストボックスからLoRAを選択してください。" & vbCrLf & _
               "Please select a LoRA from the list.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    weightVal = Me.cmbTriggerWeight.Value
    If weightVal = "" Then weightVal = "1.0"
    
    addedTriggers = ""
    addedLoRAs = ""
    
    ' 既存のプレビュー文字列を取得（重複チェックの防衛用）
    currentPreview = Trim(Me.txtLoRAPreview.Value)
    
    ' 2. リストボックスのループ (【選択されたアイテムのみ】処理)
    For i = 0 To Me.lstSelectedLoRA.ListCount - 1
        If Me.lstSelectedLoRA.Selected(i) = True Then
            
            ' リストボックスからAliasと強度を取得
            listAlias = Me.lstSelectedLoRA.List(i, 2)
            listStrength = Me.lstSelectedLoRA.List(i, 3)
            sysHash = ""
            loraModelName = ""
            
            ' メモリ上のLoRA_Listから該当データを検索し、HashとModelNameを抽出
            If Not ConfigData Is Nothing Then
                If ConfigData.Exists("LoRA_List") Then
                    Set loraList = ConfigData("LoRA_List")
                    For Each loraItem In loraList
                        If loraItem("Alias") = listAlias Then
                            sysHash = loraItem("Hash")
                            If loraItem.Exists("ModelName") Then
                                loraModelName = loraItem("ModelName")
                            End If
                            Exit For
                        End If
                    Next loraItem
                End If
            End If
            
            ' ① オプションボタン（出力ターゲット）の条件分岐判定
            If Me.opbWrapName.Value = True And Trim(loraModelName) <> "" Then
                ' 「正式名称で括る」かつデータが存在する場合
                targetName = Trim(loraModelName)
            Else
                ' 「ハッシュで括る」または正式名称が空の場合はハッシュ値を採用（互換性維持）
                If Trim(sysHash) <> "" Then targetName = Trim(sysHash) Else targetName = Trim(listAlias)
            End If
            
            If targetName <> "" Then
                ' --- A. LoRAタグの収集 ＆ プレビュー重複チェック（ハッシュと名前の両面から厳格に防衛） ---
                Dim currentLoRA As String
                Dim isLoRADup As Boolean
                
                currentLoRA = "<lora:" & targetName & ":" & listStrength & ">"
                isLoRADup = False
                
                ' 1. 「ハッシュ値」がすでに存在するかチェック
                If Trim(sysHash) <> "" Then
                    If InStr(1, currentPreview, "<lora:" & Trim(sysHash) & ":", vbTextCompare) > 0 Or _
                       InStr(1, addedLoRAs, "<lora:" & Trim(sysHash) & ":", vbTextCompare) > 0 Then
                        isLoRADup = True
                    End If
                End If
                
                ' 2. 「正式名称(ModelName)」がすでに存在するかチェック
                If Trim(loraModelName) <> "" Then
                    If InStr(1, currentPreview, "<lora:" & Trim(loraModelName) & ":", vbTextCompare) > 0 Or _
                       InStr(1, addedLoRAs, "<lora:" & Trim(loraModelName) & ":", vbTextCompare) > 0 Then
                        isLoRADup = True
                    End If
                End If
                
                ' 3. 「エイリアス(表示名)」がすでに存在するかチェック（旧仕様向けフォールバック）
                If Trim(listAlias) <> "" Then
                    If InStr(1, currentPreview, "<lora:" & Trim(listAlias) & ":", vbTextCompare) > 0 Or _
                       InStr(1, addedLoRAs, "<lora:" & Trim(listAlias) & ":", vbTextCompare) > 0 Then
                        isLoRADup = True
                    End If
                End If
                
                ' ハッシュ・正式名称・エイリアスの【どれ一つとして重複していない場合】のみ結合を許可する
                If Not isLoRADup Then
                    If addedLoRAs = "" Then
                        addedLoRAs = currentLoRA
                    Else
                        addedLoRAs = addedLoRAs & " " & currentLoRA
                    End If
                Else
                    ' ★追加：重複したLoRAをアラートメッセージのリストに蓄積
                    If InStr(1, dupMessages, "LoRA: " & listAlias, vbTextCompare) = 0 Then
                        dupMessages = dupMessages & "  - LoRA: " & listAlias & vbCrLf
                    End If
                End If
                
                ' --- B. トリガーワードの処理と重み付け ＆ 重複チェック ---
                triggerRaw = Me.lstSelectedLoRA.List(i, 1)
                Dim newTriggerRaw As String
                newTriggerRaw = ""
                
                If triggerRaw <> "" Then
                    triggerArray = Split(triggerRaw, ",")
                    For j = 0 To UBound(triggerArray)
                        Dim tWord As String
                        tWord = Trim(triggerArray(j))
                        
                        If tWord <> "" Then
                            ' 既存の重み付けを剥がす
                            tWord = reClean.Replace(tWord, "$1")
                            
                            ' コンボボックスの重みを適用
                            If weightVal <> "1.0" Then
                                tWord = "(" & tWord & ":" & weightVal & ")"
                            End If
                            
                            ' 重複チェック：すでにプレビュー欄、または今回追加分に同じトリガーがないか判定
                            Dim isTriggerDup As Boolean
                            isTriggerDup = False
                            
                            If currentPreview <> "" Then
                                ' カンマ区切りの独立要素として安全に部分一致をチェック
                                If InStr(1, ", " & currentPreview & ",", ", " & tWord & ",", vbTextCompare) > 0 Or _
                                   InStr(1, currentPreview, tWord, vbTextCompare) > 0 Then
                                    isTriggerDup = True
                                End If
                            End If
                            
                            If InStr(1, addedTriggers, tWord, vbTextCompare) > 0 Then
                                isTriggerDup = True
                            End If
                            
                            ' 重複していない場合のみ結合
                            If Not isTriggerDup Then
                                If newTriggerRaw = "" Then
                                    newTriggerRaw = tWord
                                Else
                                    newTriggerRaw = newTriggerRaw & ", " & tWord
                                End If
                            Else
                                ' ★追加：重複したトリガーワードをアラートメッセージのリストに蓄積
                                If InStr(1, dupMessages, "Tag: " & tWord, vbTextCompare) = 0 Then
                                    dupMessages = dupMessages & "  - Tag: " & tWord & vbCrLf
                                End If
                            End If
                        End If
                    Next j
                End If
                
                ' リストボックスの表示（トリガー部分）を更新して状態を記憶
                Me.lstSelectedLoRA.List(i, 1) = newTriggerRaw
                
                ' 全体（今回追加分）のトリガー文字列へ結合
                If newTriggerRaw <> "" Then
                    If addedTriggers = "" Then
                        addedTriggers = newTriggerRaw
                    Else
                        addedTriggers = addedTriggers & ", " & newTriggerRaw
                    End If
                End If
                
            End If
            
            ' 処理が終わったら選択状態を解除する（連続操作しやすくするため）
            Me.lstSelectedLoRA.Selected(i) = False
            
        End If
    Next i
    
' ========================================================
    ' 3. 今回追加するプロンプトの結合とプレビュー欄への【追記】
    ' ========================================================
    If addedTriggers <> "" And addedLoRAs <> "" Then
        appendedPrompt = addedTriggers & ", " & addedLoRAs
    ElseIf addedTriggers <> "" Then
        appendedPrompt = addedTriggers
    ElseIf addedLoRAs <> "" Then
        appendedPrompt = addedLoRAs
    Else
        appendedPrompt = ""
    End If
    
    If appendedPrompt <> "" Then
        ' プレビュー欄が空でなければ、状態に合わせて賢く繋ぐ
        If currentPreview <> "" Then
            ' ★修正：既存の末尾が ">" (LoRA) で、かつ追加分の先頭が "<" (LoRA) の場合はカンマを打たない
            If Right(Trim(currentPreview), 1) = ">" And Left(Trim(appendedPrompt), 1) = "<" Then
                Me.txtLoRAPreview.Value = Trim(currentPreview) & " " & Trim(appendedPrompt)
            
            ' すでに末尾にカンマがある場合はスペースだけ追加
            ElseIf Right(Trim(currentPreview), 1) = "," Then
                Me.txtLoRAPreview.Value = Trim(currentPreview) & " " & Trim(appendedPrompt)
            
            ' それ以外（通常のタグからLoRAに繋ぐ場合など）はカンマとスペースで繋ぐ
            Else
                Me.txtLoRAPreview.Value = Trim(currentPreview) & ", " & Trim(appendedPrompt)
            End If
        Else
            ' プレビュー欄が空ならそのまま入れる
            Me.txtLoRAPreview.Value = Trim(appendedPrompt)
        End If
    End If
    
    ' ========================================================
    ' ★ 4. 最後に重複があった場合のみ、1回だけダイアログを表示
    ' ========================================================
    If dupMessages <> "" Then
        MsgBox "以下の項目は既にプレビュー欄に存在するため、重複をスキップしました：" & vbCrLf & _
               "The following items were skipped to prevent duplication:" & vbCrLf & vbCrLf & _
               dupMessages, vbInformation, APP_NAME
    End If
End Sub

Private Sub btnSendToCockpit_Click()
    Call SendFavoriteToMain
End Sub

'-------------------------------------------------------------------------
' [Set LoRA] ボタンが押された時の処理（リストへ追加・重複ブロック版）
'-------------------------------------------------------------------------
Private Sub btnSetLoRA_Click()
    Dim aliasName As String
    Dim strengthVal As String
    Dim displayStr As String
    Dim triggerStr As String
    Dim i As Integer
    Dim newIndex As Integer
    
    ' （追記）安全装置：リストボックスを強制的に4列・幅指定にする
    Me.lstSelectedLoRA.ColumnCount = 4
    Me.lstSelectedLoRA.ColumnWidths = "180;0;0;0"
    
    ' 1. バリデーション（未入力チェック）
    If Me.cmbLoRAList.ListIndex = -1 Then
        MsgBox "リストに追加するLoRAを選択してください。" & vbCrLf & _
               "Please select a LoRA to add to the list.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Me.cmbLoRAStrength.Value = "" Then
        MsgBox "LoRAの強度(Strength)が設定されていません。" & vbCrLf & _
               "LoRA Strength is not set.", vbExclamation, "Setting Error"
        Exit Sub
    End If
    
    aliasName = Me.cmbLoRAList.Value
    
    ' ★追加対策：リスト内重複チェック（既に同じLoRAが入っていないか？）
    For i = 0 To Me.lstSelectedLoRA.ListCount - 1
        ' 3列目 (Index 2) に純粋なAliasが隠されているので、それと照合する
        If Me.lstSelectedLoRA.List(i, 2) = aliasName Then
            MsgBox "「" & aliasName & "」は既にリストに追加されています。" & vbCrLf & _
                   "'" & aliasName & "' is already in the list.", vbExclamation, APP_NAME
            Exit Sub ' 重複していたらここで処理を強制終了！
        End If
    Next i
    
    ' 2. 1カラム目（表示用）の文字列を作成
    strengthVal = Me.cmbLoRAStrength.Value
    displayStr = aliasName & " / " & strengthVal
    
    ' 3. 2カラム目（裏データ）用のトリガーワードを収集
    triggerStr = ""
    For i = 1 To 10
        ' チェックボックスが表示されていて、かつチェックが入っているものだけを拾う
        If Me.Controls("chkTrigger_" & i).Visible = True Then
            If Me.Controls("chkTrigger_" & i).Value = True Then
                ' すでに単語が入っていれば、カンマとスペースで区切る
                If triggerStr <> "" Then
                    triggerStr = triggerStr & ", "
                End If
                triggerStr = triggerStr & Me.Controls("chkTrigger_" & i).Caption
            End If
        End If
    Next i
    
    ' 4. リストボックスへ追加
    Me.lstSelectedLoRA.AddItem "" ' まず空の行を1つ追加する
    newIndex = Me.lstSelectedLoRA.ListCount - 1 ' 今追加した行のインデックス番号を取得
    
    ' 1列目 (Index 0) : 表示文字列（Alias / Strength） ※画面に見える
    Me.lstSelectedLoRA.List(newIndex, 0) = displayStr
    ' 2列目 (Index 1) : トリガーワード（カンマ区切り） ※Width=0で隠す
    Me.lstSelectedLoRA.List(newIndex, 1) = triggerStr
    ' 3列目 (Index 2) : 純粋なAlias（検索用） ※Width=0で隠す
    Me.lstSelectedLoRA.List(newIndex, 2) = aliasName
    ' 4列目 (Index 3) : 純粋なStrength（出力用） ※Width=0で隠す
    Me.lstSelectedLoRA.List(newIndex, 3) = strengthVal
    ' ★追加：カートにデータが入ったため、UI状態を再判定させてWeightコンボを点灯させる
    Call UpdateLoRAUIState

End Sub



Private Sub chkImportLoRAPresets_Click()

End Sub

' chkUseLoRA がクリックされた時のイベント
Private Sub chkUseLoRA_Click()
    Dim state As Boolean
    state = Me.chkUseLoRA.Value
    
    ' オフになった瞬間の初期化処理（古いToggleLoRAUIから移管）
    If state = False Then
        Dim i As Integer
        For i = 1 To 10
            Me.Controls("chkTrigger_" & i).Value = False
            Me.Controls("chkTrigger_" & i).Visible = False
        Next i
        Me.lblNotRequiredTrigger.Visible = False
        Me.cmbLoRAList.ListIndex = -1
        Me.cmbLoRAStrength.ListIndex = -1
        Me.cmbPresetList.ListIndex = -1
        Me.txtLoRAPreview.text = ""
        Me.lstSelectedLoRA.Clear
    End If
    
    ' 最新のUI一括管理を呼び出す
    Call UpdateLoRAUIState
' LoRA機能全体のON/OFFが切り替わった最後に、カートの状態を強制的に再評価させて上書きする
    Call lstSelectedLoRA_Change
End Sub

Private Sub btnAddToFav_Click()
    Call AddToFavorite
    Call SyncFavoritesManager ' ★追加：非同期ズレ防止
End Sub

Private Sub btnFavClear_Click()
    Call FavClear
    Call SyncFavoritesManager ' ★追加：非同期ズレ防止
End Sub

Private Sub btnFavCopy_Click()
    Call FavoriteCopy
End Sub

Private Sub btnPullFromCockpit_Click()
    Call TransferPrompt
End Sub

Private Sub btnReplaceFav_Click()
    Call ReplaceFavoriteCell
    Call SyncFavoritesManager ' ★追加：非同期ズレ防止
End Sub

Private Sub btnSendToFav_Click()
    ' 1. プロンプトの転送処理を実行
    Call TransferPrompt
    
    If Trim(Me.txtMain) = "" Then
        Exit Sub
    End If
    
    ' 2. タブを「Favorite」（インデックス4）に切り替える
    Me.MainWindow.Value = 4
    
    ' 3. Favoriteタブ内のテキストボックスにフォーカスを当てる
    With Me.txtFav
        .SetFocus
        ' 追記しやすいよう、カーソルをテキストの最後に移動させる（おまけ）
        .SelStart = Len(.text)
    End With
End Sub

' ==========================================
' 【Tweaked!】編集したお気に入りプロンプトをコピーする
' ==========================================
Private Sub btnTweaked_Click()
    Dim clip As Object
    
    ' クリップボードへ送る
    'On Error Resume Next
    'Set clip = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    'clip.SetText Me.txtFav.Value
    'clip.PutInClipboard
    'On Error GoTo 0
    Call SetClipboardText(Me.txtFav.Value)
    
    ' --- 【追加】全選択の演出 ---
    With Me.txtFav
        .SetFocus            ' フォーカスを当てないと選択状態が見えません
        .SelStart = 0        ' 選択開始位置を先頭に
        .SelLength = Len(.Value) ' 全文字分を選択
    End With
    DoEvents                 ' 描画を強制して選択状態を画面に反映させる
    ' ----------------------------
    
    ' 状態を更新
    lastTweakedValue = Me.txtFav.Value
    Me.btnTweaked.Enabled = False
    Me.btnTweaked.ForeColor = RGB(160, 160, 160)
    
    ' フィードバック
    Application.StatusBar = "Tweaked prompt copied to clipboard!"
    
    ' 3秒後に消去（待機中に全選択状態が維持されます）
    Application.Wait [Now() + "00:00:03"]
    Application.StatusBar = False
    
    ' 必要であれば、待機後に選択を解除するなら以下を追加
    ' Me.txtFav.SelLength = 0
End Sub

'-------------------------------------------------------------------------
' カテゴリ選択時のワープ処理（エラー解消・再選択対応版）
'-------------------------------------------------------------------------
Private Sub cmbCategoryJump_Change()
    Dim targetName As String
    Dim wsName As String: wsName = "KENZEN SeaArt Helper"
    
    ' システムによる表示リセット中は処理をスキップ
    If IsUpdatingBySystem Then Exit Sub
    
    ' 選択されたアイテムの2列目（名前定義）を取得
    ' 「Jump to Category」（Index 0）の場合は何もしない
    If Me.cmbCategoryJump.ListIndex <= 0 Then Exit Sub
    targetName = Me.cmbCategoryJump.List(Me.cmbCategoryJump.ListIndex, 1)
    
    If targetName = "" Then Exit Sub
    
    On Error Resume Next
    ' ★重要：以前の行で発生した可能性のあるエラー情報をクリア
    Err.Clear
    
    ' 1. シートを切り替える（既にアクティブならスキップして安定性を高める）
    If ActiveSheet.Name <> wsName Then
        Sheets(wsName).Select
    End If
    
    ' 2. 名前定義へジャンプ
    ' ジャンプ直前に再度クリアし、ジャンプ自体の成功を確実に判定する
    Err.Clear
    Application.GoTo Reference:=targetName, Scroll:=True
    
    ' --- 判定ロジックの修正 ---
    If Err.Number <> 0 Then
        ' 本当にジャンプに失敗した場合（名前定義が存在しない場合）のみ通知
        MsgBox "指定されたカテゴリ「" & targetName & "」が見つかりませんでした。" & vbCrLf & _
               "Excel側の「名前の管理」で定義されているか確認してください。", vbCritical, APP_NAME
    Else
        ' ジャンプ成功時：表示をリセットして「同じ項目の再選択」を可能にする
        IsUpdatingBySystem = True
        Me.cmbCategoryJump.ListIndex = 0
        IsUpdatingBySystem = False
    End If
    
    On Error GoTo 0
    
    ' 3. ユーザーフォームにフォーカスを戻す
    Me.btnCopyNormal.SetFocus
End Sub

'-------------------------------------------------------------------------
' cmbLoRAList の選択が変更された瞬間に、強度とトリガーを自動展開する (JSON対応版)
'-------------------------------------------------------------------------
Private Sub cmbLoRAList_Change()
    Dim selectedAlias As String
    Dim loraCol As Object
    Dim item As Object
    Dim foundItem As Object
    Dim recStrength As Variant
    Dim triggerStr As String
    Dim triggerArray() As String
    Dim i As Integer
    
    ' 1. 一旦UIを「更地化」して前回の表示を消す
    Me.cmbLoRAStrength.ListIndex = -1
    Me.lblNotRequiredTrigger.Visible = False
    For i = 1 To 10
        Me.Controls("chkTrigger_" & i).Visible = False
        Me.Controls("chkTrigger_" & i).Value = False
    Next i
    
    ' 2. 未選択状態（Clearされた時など）ならここで終了
    If Me.cmbLoRAList.ListIndex = -1 Then
        Call UpdateLoRAUIState
        Exit Sub
    End If
    
    selectedAlias = Me.cmbLoRAList.Value
    
    ' 3. JSONデータ（ConfigData）から該当するAliasのアイテムを検索
    If ConfigData Is Nothing Then Exit Sub
    If Not ConfigData.Exists("LoRA_List") Then Exit Sub
    
    Set loraCol = ConfigData("LoRA_List")
    Set foundItem = Nothing
    
    ' コレクション内の全LoRAを走査し、選択されたAliasと一致するものを探す
    For Each item In loraCol
        If item("Alias") = selectedAlias Then
            Set foundItem = item
            Exit For ' 見つかったらループを抜ける（高速化）
        End If
    Next item
    
    ' 万が一見つからなかった場合はUIを更新して終了
    If foundItem Is Nothing Then
        Call UpdateLoRAUIState
        Exit Sub
    End If
    
    ' ==========================================
    ' 4. RecommendedStrength の取得と自動セット
    ' ==========================================
    recStrength = foundItem("Strength")
    
    If IsNumeric(recStrength) Then
        Me.cmbLoRAStrength.Value = Format(recStrength, "0.00")
    End If
    
    ' ==========================================
    ' 5. TriggerWords の取得とチェックボックスへの展開
    ' ==========================================
    triggerStr = Trim(foundItem("Trigger"))
    
    If triggerStr = "" Then
        Me.lblNotRequiredTrigger.Visible = True
    Else
        triggerArray = Split(triggerStr, ",")
        For i = 0 To UBound(triggerArray)
            If i < 10 Then
                With Me.Controls("chkTrigger_" & (i + 1))
                    .Caption = Trim(triggerArray(i))
                    .Visible = True
                    .Enabled = True
                    .Value = True
                End With
            End If
        Next i
    End If
    
    ' ★動的な有効化条件を反映させる
    Call UpdateLoRAUIState
End Sub


Private Sub lblGachaCount_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim newLimit As String
    Dim currentLimit As String
    
    ' 現在の設定値を取得
    currentLimit = GetSetting(APP_NAME, "Gacha", "MaxLimit", "20")
    
    ' 入力ボックスを日英併記に
    newLimit = InputBox("1日の最大ガチャ回数を設定してください。" & vbCrLf & _
                        "(課金済みの場合は 1500 や 9999 を入力)" & vbCrLf & vbCrLf & _
                        "Please set the maximum daily Gacha limit." & vbCrLf & _
                        "(Enter 1500 or 9999 if you are a paid user)", _
                        "Limit Setting", currentLimit)
    
    ' 入力内容が数値（かつ空でない）場合のみ保存
    If IsNumeric(newLimit) And newLimit <> "" Then
        SaveSetting APP_NAME, "Gacha", "MaxLimit", newLimit
        Me.lblGachaCount.Caption = GetGachaStatus()
    End If
End Sub


Private Sub lstAllPositive_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
' Shift=1, Ctrl=2 なので、両方押しは 3
    If Shift = 3 And KeyCode = vbKeyP Then
        Call SetPositiveAction
        KeyCode = 0 ' 他の処理にキーを渡さない（ビープ音防止）
    End If
End Sub

Private Sub lstApplyPositive_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
' Shift=1, Ctrl=2 なので、両方押しは 3
    If Shift = 3 And KeyCode = vbKeyP Then
        Call SetPositiveAction
        KeyCode = 0 ' 他の処理にキーを渡さない（ビープ音防止）
    End If
End Sub

' =========================================================================
' [Mobile UI] リスト選択時に詳細テキストボックスへ内容を展開する
' =========================================================================
Private Sub lstMobileMemo_Click()
    Dim listIdx As Long
    
    ' 現在選択されている行のインデックスを取得
    listIdx = Me.lstMobileMemo.ListIndex
    
    ' 選択が解除されている場合（ListIndex = -1）はテキストボックスを空にして終了
    If listIdx = -1 Then
        Me.txtMobileMemo.text = ""
        Exit Sub
    End If
    
    ' リストの2列目（インデックス1）のテキストを、そのままテキストボックスに代入
    Me.txtMobileMemo.text = Me.lstMobileMemo.List(listIdx, 1)
End Sub

' =========================================================================
' [Mobile Utility] 詳細テキストボックスの内容をクリップボードに全コピー
' =========================================================================
Private Sub btnCopyMobiDetail_Click()
    Dim cb As Object
    
    ' テキストボックスが空の場合は何もしない
    If Me.txtMobileMemo.text = "" Then
        MsgBox "コピーするテキストがありません" & vbCrLf & _
               "No text.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' 確実なLate Binding方式でクリップボードへ転送
    On Error Resume Next
    Set cb = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    cb.SetText Me.txtMobileMemo.text
    cb.PutInClipboard
    
    If Err.Number = 0 Then
        MsgBox "テキストをクリップボードにコピーしました！" & vbCrLf & _
               "Text copied to clipboard!", vbInformation, APP_NAME
    Else
        MsgBox "クリップボードへのコピーに失敗しました。" & vbCrLf & _
               "Failed to copy to clipboard.", vbCritical, APP_NAME
    End If
    On Error GoTo 0
    Set cb = Nothing
    
    ' --- 視覚演出: MsgBoxが閉じた後にフォーカスを当てて全選択状態にする ---
    Me.txtMobileMemo.SetFocus
    Me.txtMobileMemo.SelStart = 0
    Me.txtMobileMemo.SelLength = Len(Me.txtMobileMemo.text)
End Sub

' =========================================================================
' [Mobile Utility] 詳細テキストボックスの表示を消去
' =========================================================================
Private Sub btnClearMobiDetail_Click()
    ' テキストボックスの文字列を空にするだけ（リストのデータ自体は消えません）
    Me.txtMobileMemo.text = ""
End Sub


Private Sub txtAPIKey_Change()
    Call UpdateGachaUI
End Sub

' フォーカスが当たった時：プレースホルダーを消す
Private Sub txtInputNegative_Enter()
    With Me.txtInputNegative
        If .text = PH_NEGA_INPUT Then
            .text = ""
            .ForeColor = vbBlack
        End If
    End With
End Sub

' フォーカスが外れた時：空ならプレースホルダーを戻す
Private Sub txtInputNegative_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.txtInputNegative
        If Trim(.text) = "" Then
            .text = PH_NEGA_INPUT
            .ForeColor = COLOR_GRAY
        End If
    End With
End Sub

' テキストボックス内で Ctrl+Shift+P を押したときも発火させる
Private Sub txtMain_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Shift=1, Ctrl=2 なので、両方押しは 3
    If Shift = 3 And KeyCode = vbKeyP Then
        Call SetPositiveAction
        KeyCode = 0 ' 他の処理にキーを渡さない（ビープ音防止）
    End If
End Sub

' =========================================================================
' [Mobile UI] テキストボックスの編集をリストにリアルタイム反映
' =========================================================================
Private Sub txtMobileMemo_Change()
    Dim listIdx As Long
    
    ' 現在選択されている行のインデックスを取得
    listIdx = Me.lstMobileMemo.ListIndex
    
    ' アイテムが選択されている場合のみ、リストの2列目（インデックス1）を更新
    If listIdx <> -1 Then
        Me.lstMobileMemo.List(listIdx, 1) = Me.txtMobileMemo.text
    End If
End Sub

Private Sub txtRequest_Change()
    ' システムによる自動入力中でなければUI更新
    If Not isAutoFilling Then Call UpdateGachaUI
End Sub

' --- txtFav が編集されたらボタンを光らせる ---
Private Sub txtFav_Change()
    ' 現在の内容が、前回コピーした内容と異なる場合のみ有効化
    If Me.txtFav.Value <> lastTweakedValue And Trim(Me.txtFav.Value) <> "" Then
        Me.btnTweaked.Enabled = True
        Me.btnTweaked.ForeColor = RGB(0, 0, 0) ' 黒（有効）
    Else
        Me.btnTweaked.Enabled = False
        Me.btnTweaked.ForeColor = RGB(160, 160, 160) ' グレー（無効）
    End If
End Sub

' フォーカスが当たった時：プレースホルダーを消す
Private Sub txtFavSearch_Enter()
    With Me.txtFavSearch
        If .text = PH_FAV_SEARCH Then
            .text = ""
            .ForeColor = COLOR_BLACK
        End If
    End With
End Sub

' フォーカスが外れた時：空ならプレースホルダーを戻す
Private Sub txtFavSearch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.txtFavSearch
        If Trim(.text) = "" Then
            .text = PH_FAV_SEARCH
            .ForeColor = COLOR_GRAY
        End If
    End With
End Sub

'-------------------------------------------------------------------------
' テキストボックスの中身が変わった瞬間に実行される（リアルタイムUI制御）
'-------------------------------------------------------------------------
Private Sub txtMain_Change()
    ' システムによる自動更新中は処理を抜ける
    If IsUpdatingBySystem Or isAutoFilling Then Exit Sub
    
    ' =======================================================
    ' ★ 新規追加：手動入力の「最初の1文字目」だけ履歴に保存
    ' =======================================================
    If Not isTypingSession Then
        isTypingSession = True
        
        Dim needSave As Boolean
        needSave = True
        
        ' 最新の履歴と「入力前の状態」が同じなら二重保存しない
        If Not UndoMemory Is Nothing Then
            If UndoMemory.Count > 0 Then
                If UndoMemory(UndoMemory.Count) = PreEditText Then
                    needSave = False
                End If
            End If
        End If
        
        If needSave Then
            Call SaveHistory(PreEditText) ' ここでUndoボタンが自動的にEnabledになります
        End If
    End If
    ' =======================================================
    
    Dim hasText As Boolean
    hasText = (Trim(Me.txtMain.text) <> "")
    
    ' --- 重み付けガードレール ＆ クリアボタン制御 ---
    Me.chkWeight.Enabled = hasText
    Me.btnClear.Enabled = hasText
    Me.btnAllClear.Enabled = hasText
    
    If Not hasText Then
        Me.chkWeight.Value = False
        Call ApplyWeightUIAppearance(Me)
    End If
    
    'Call UpdateDoneButtonState
End Sub

' --- Done! ボタンクリック ---
Private Sub btnDone_Click()
    ' 1. 最新の状態を変数へ同期
    AccumulatedPrompt = Me.txtMain.Value
    
    ' 2. 決定稿として履歴に刻む
    Call SaveHistory(AccumulatedPrompt)
    
    ' 3. クリップボードへ転送（ここで全選択状態にする演出も可能）
    'Dim clip As Object
    'Set clip = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    'clip.SetText AccumulatedPrompt
    'clip.PutInClipboard
    Call SetClipboardText(AccumulatedPrompt)
    
    ' 4. UIフィードバック
    Application.StatusBar = "Done! Prompt copied to clipboard."
    
    ' 5. ボタンを消灯（連続クリック防止）
    'Me.btnDone.Enabled = False
    'Me.btnDone.BackColor = &HE0E0E0
    
    ' 最後に、コピーしたことを知らせるためにテキストボックスを一瞬全選択する
    Me.txtMain.SetFocus
    Me.txtMain.SelStart = 0
    Me.txtMain.SelLength = Len(Me.txtMain.text)
    LastCopiedPrompt = Me.txtMain.Value
' コピー処理が完了した後のコードに追記
    LastCopiedPrompt = Trim(Me.txtMain.text)
    'IsDirty = False ' ここでリセット！
    'UpdateDoneButtonState
End Sub

'-------------------------------------------------------------------------
' 出力欄 (txtYourPositive) で Ctrl+Shift+P を押したときも発火させる
'-------------------------------------------------------------------------
Private Sub txtYourPositive_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Shift=1, Ctrl=2 なので、両方押しは 3
    If Shift = 3 And KeyCode = vbKeyP Then
        Call SetPositiveAction
        KeyCode = 0 ' 他の処理にキーを渡さない（ビープ音防止）
    End If
End Sub
' =========================================================================
' 1. フォーム起動時の処理
' =========================================================================
Private Sub UserForm_Initialize()

    ' ---------------------------------------------------------------------
    ' システム・ウィンドウ周りの初期設定
    ' ---------------------------------------------------------------------
    ' Windowsの「応答なし（ホワイトアウト）」お節介機能を停止
    Call DisableProcessWindowsGhosting
    
    Dim hwnd As LongPtr
    Dim style As Long
    
    ' ウィンドウハンドルを取得して最小化ボタンを付与
    hwnd = FindWindow("ThunderDFrame", Me.Caption)
    #If VBA7 Then
        style = GetWindowLongPtr(hwnd, GWL_STYLE)
        SetWindowLongPtr hwnd, GWL_STYLE, style Or WS_MINIMIZEBOX
    #Else
        style = GetWindowLong(hwnd, GWL_STYLE)
        SetWindowLong hwnd, GWL_STYLE, style Or WS_MINIMIZEBOX
    #End If
    
    ' ---------------------------------------------------------------------
    ' データの初期ロード・同期
    ' ---------------------------------------------------------------------
    ' すべてのUI構築が始まる前に、JSONデータをメモリに完全ロードする
    If ConfigData Is Nothing Then Call InitConfigJSON
    
    ' JSONのデータをMy Favoriteシートに展開して同期する
    Call SyncFavSheetFromJSON
    
    If dictJPtoEN Is Nothing Then InitDictionaries
    
    ' ---------------------------------------------------------------------
    ' コクピット（Cockpit）タブの初期化
    ' ---------------------------------------------------------------------
    ' タブを強制的に一番左（Cockpit: Index 0）にセット
    Me.MainWindow.Value = 0
    
    Call InitCategoryJump
    
    ' 検索ボックスの初期化
    With Me.txtSearch
        .text = PH_TEXT
        .ForeColor = COLOR_GRAY
    End With
    Me.btnSearch.ForeColor = COLOR_BLACK
    Me.btnSearch.BackColor = RGB(200, 225, 255) ' colAction
    
    lastSearchText = ""
    Set lastFoundCell = Nothing
    
    ' --- ★追加：Cockpitのクリアボタンを初期封印 ---
    'Me.btnClear.Enabled = False
    'Me.btnAllClear.Enabled = False
    
    ' --- Done! ボタンの初期封印（統一外観を適用） ---
    'Call SetControlUIState(Me.btnDone, False)
    
    ' --- Undoボタンの初期設定 ---
    'Call UpdateUndoButtonState
    
    ' ---------------------------------------------------------------------
    ' コクピット（Cockpit）の重み付け・Wrap周りの初期封印
    ' ---------------------------------------------------------------------
    With Me.cmbWeightValue
        .AddItem "0.5"
        .AddItem "0.6"
        .AddItem "0.7"
        .AddItem "0.8"
        .AddItem "0.9"
        .AddItem "1.1"
        .AddItem "1.2"
        .AddItem "1.3"
        .AddItem "1.4"
        .AddItem "1.5"
        .Value = "1.1"
    End With
    
    Me.chkWeight.Value = False
    Me.chkWeight.Enabled = False
    Call ApplyWeightUIAppearance(Me)
    
    Call SetControlUIState(Me.cmbLoRAList, False)
    Call SetControlUIState(Me.btnWrapBlock, False)
    Call SetControlUIState(Me.cmbWeightValue, False)
    Call SetControlUIState(Me.opbWrapHash, False)
    Call SetControlUIState(Me.opbWrapName, False)
    
    Me.btnCopyNormal.BackColor = RGB(200, 225, 255) ' colAction
    
    ' ---------------------------------------------------------------------
    ' 各種タブ（LoRA, Positive, Negative, Favorite, Gacha）の初期化
    ' ---------------------------------------------------------------------
    ' --- LoRAタブ ---
    Call RefreshMainWindowLoRA
    Call InitLoRAStrengthCombo
    Call InitTriggerWeightCombo
    Call RefreshPresetList
    
    ' トリガーワードUIの初期化（非表示・文字クリア）
    Dim tIdx As Integer
    For tIdx = 1 To 10
        With Me.Controls("chkTrigger_" & tIdx)
            .Caption = ""
            .Value = False
            .Visible = False
        End With
    Next tIdx
    Me.lblNotRequiredTrigger.Visible = False
    Call LockLoRAUIIfEmpty
    Call lstSelectedLoRA_Change
    
    ' --- Positive / Negative / Favorite タブ ---
    Call InitPositiveUI
    Call InitNegativePromptPage
    Call UpdateFavUndoStatus(False) ' FavoriteのUndo履歴リセット
    
    ' --- Gacha! (Gemini API) タブ ---
    Me.txtAPIKey.text = GetSetting(APP_NAME, "Settings", "GeminiAPIKey", "")
    Me.lblGachaCount.Caption = GetGachaStatus()
    Call AlignGachaLabel
    Call UpdateGachaUI(False)
    Call UpdateSurpriseMeUI

    ' =========================================================
    ' モバイル用タブ (pgMobile) のUI初期化
    ' =========================================================
    With Me.lstMobileMemo
        .ColumnCount = 2
        .ColumnWidths = "44;400" ' 1列目:タグ(44)、2列目:内容(400)
        .MultiSelect = fmMultiSelectSingle ' 単一選択に明示的に設定 (値:0)
    End With
    
    ' =========================================================
    ' ★追加：モバイル用タブ (pgMobile) のストックデータ復元処理
    ' =========================================================
    Call RefreshMobileUI
    
End Sub

' =========================================================
' v.3.0.0 MainWindow：ポジティブプロンプト関連イベント (JSON版)
' =========================================================

' --- A. フォーム初期化時（UserForm_Initialize）の末尾などに仕込んでください ---
Private Sub InitPositiveUI()
    ' JSONの初期化エンジンを確実に走らせる
    Call InitPositiveJSON
    
    ' プリセットコンボボックスにリストを充填（先頭にDefaultが自動挿入されます）
    Call LoadPositivePresetsToCombo(Me.cmbSelectPosiPreset)
    
    ' ★ 起動時に「Default」を最初からカチッと選択状態にする（UI同期）
    If Me.cmbSelectPosiPreset.ListCount > 0 Then
        Me.cmbSelectPosiPreset.ListIndex = 0
    End If
    
    ' =========================================================
    ' 追加：起動時にJSON内の「PositiveStock」からストックを復元し、
    '       管理用リストボックス（lstAllPositive）へバラして即時展開
    ' =========================================================
    Me.lstAllPositive.Clear
    
    Dim stockPrompt As String
    stockPrompt = ""
    
    ' メモリ（JSON）から前回のストック文字列を安全に取得
    On Error Resume Next
    stockPrompt = Trim(ConfigData("PositiveStock"))
    On Error GoTo 0
    
    ' ストックが存在する場合のみ分解処理を実行
    If stockPrompt <> "" Then
        Dim words() As String
        Dim word As String
        Dim k As Long
        
        ' カンマ前後のスペースのブレを綺麗に整地
        stockPrompt = Replace(stockPrompt, " ,", ",")
        stockPrompt = Replace(stockPrompt, ", ", ",")
        words = Split(stockPrompt, ",")
        
        ' 画面更新の一時停止などを行わなくても、メモリ上からのAddItemなので一瞬で終わります
        For k = 0 To UBound(words)
            word = Trim(words(k))
            ' 空文字でなければ、リストボックスへ1単語ずつ切り分けて直投入
            If word <> "" Then Me.lstAllPositive.AddItem word
        Next k
    End If
        
    With Me.txtInputPositive
        .text = PH_POSI_INPUT
        .ForeColor = COLOR_GRAY
    End With

End Sub


Private Sub btnCopySpace_Click()
    Dim targetText As String
    
    ' 0. ★追加：My Favoriteシートからの直接コピーを制限するバリデーション
    ' お気に入りのプロンプトは、メインウィンドウの「Favorite」タブから流し込むのが正規ルートです
    If ActiveSheet.Name = "My Favorite" Then
        MsgBox "「My Favorite」シートから直接コピーすることはできません。" & vbCrLf & _
               "お気に入りのプロンプトを追加する場合は、メインウィンドウの「Favorite」タブから操作してください。" & vbCrLf & vbCrLf & _
               "Direct copying from the 'My Favorite' sheet is disabled." & vbCrLf & _
               "To use favorite prompts, please operate from the 'Favorite' tab in the main window.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 1. ガード：空白なら警告して終了
    If IsError(ActiveCell.Value) Or Trim(CStr(ActiveCell.Value)) = "" Then
        MsgBox "選択されたセルは空白、またはエラー値です。" & vbCrLf & _
               "The selected cell is empty or contains an error.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 2. 値の取得と重み付け
    targetText = Trim(CStr(ActiveCell.Value))
    targetText = GetWeightedValue(targetText)
    
    ' 3. ロジック呼び出し
    Call ExecutePromptLogic(targetText, False)
    ' 変更があったことを明示する
    'IsDirty = True
    'UpdateDoneButtonState
    ' サーチウィンドウの場合はフォーカスを戻す（既存の挙動を維持）
    On Error Resume Next
    Me.btnSearch.SetFocus
    On Error GoTo 0
End Sub

' --- 重み付けチェックボックスの連動 ---
Private Sub chkWeight_Click()
    Dim isEnabled As Boolean
    isEnabled = Me.chkWeight.Value
    
    ' Wrap Blockボタンの有効化と色の変更
    Me.btnWrapBlock.Enabled = isEnabled
    If isEnabled Then
        Me.btnWrapBlock.BackColor = colAction   ' 有効時は「実行ブルー」
        Me.btnWrapBlock.ForeColor = vbBlack     ' 文字をはっきりさせる
    Else
        Me.btnWrapBlock.BackColor = colDisabled ' 無効時は「中立グレー」
    End If
    
    ' コンボボックスも同様
    Me.cmbWeightValue.Enabled = isEnabled
    Me.cmbWeightValue.BackColor = IIf(isEnabled, vbWhite, colDisabled)
End Sub

' ==========================================
' 検索ボックスのクリア処理
' ==========================================
Private Sub btnSearchClear_Click()
    ' 1. テキストボックスをプレースホルダー状態に戻す
    With Me.txtSearch
        .text = PH_TEXT
        .ForeColor = COLOR_GRAY
    End With
    
    ' 2. 検索の継続性（FindNext用）をリセット
    ' これをしないと、新しい言葉で検索した時に変な場所から始まってしまいます
    lastSearchText = ""
    Set lastFoundCell = Nothing
    
    ' 3. すぐに次の入力を始められるようにフォーカスを当てる
    Me.txtSearch.SetFocus
    
    ' ステータスバーも軽くリセット
    Application.StatusBar = "Search box cleared."
End Sub


' ==========================================
' 2. btnSearch（検索窓）の挙動
' ==========================================

' --- 1. フォーカスが当たった時（入力開始） ---
Private Sub txtSearch_Enter()
    With Me.txtSearch
        ' もしプレースホルダーが表示されていたら、中身を消して文字色を黒にする
        If .text = PH_TEXT Then
            .text = ""
            .ForeColor = COLOR_BLACK
        End If
    End With
End Sub

' --- 2. フォーカスが外れた時（入力終了） ---
Private Sub txtSearch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.txtSearch
        ' もし中身が空っぽだったら、プレースホルダーを復活させて文字色をグレーにする
        If Trim(.text) = "" Then
            .text = PH_TEXT
            .ForeColor = COLOR_GRAY
        End If
    End With
End Sub


' ==========================================
' 統合検索ロジック（Find & FindNext + メモ検索対応）
' ==========================================
Private Sub btnSearch_Click()
    Dim searchText As String
    Dim foundCell As Range
    Dim isHalfWidth As Boolean
    Dim i As Integer
    Dim firstFoundAddr As String
    
    ' テキストボックス（検索窓）から入力を取得
    searchText = Trim(Me.txtSearch.text)
    
    ' 未入力またはプレースホルダーならエラー
    If searchText = "" Or searchText = PH_TEXT Then
        MsgBox "検索ワードが入力されていません。" & vbCrLf & _
               "Please enter a search keyword.", vbExclamation, APP_NAME
        Me.txtSearch.SetFocus
        Exit Sub
    End If
    
    ' 1. 半角判定（日本語検索なら英語セルへ飛ばすため）
    isHalfWidth = True
    For i = 1 To Len(searchText)
        If AscW(Mid(searchText, i, 1)) > 255 Then
            isHalfWidth = False
            Exit For
        End If
    Next i
    
    ' 2. 検索実行の分岐設定
    ' 前回と異なるワードなら、検索履歴をリセットして「値検索」から開始
    If searchText <> lastSearchText Or lastFoundCell Is Nothing Then
        lastSearchText = searchText
        Set lastFoundCell = Nothing
        isSearchingComments = False ' セルの値検索モードからスタート
    End If

    ' --- [検索実行ループ] ---
    ' 「セルの値検索」が一周したら「コメント検索」へ移行する
    Do
        If Not isSearchingComments Then
            ' ▼ セルの値を検索（LookIn:=xlValues）
            If lastFoundCell Is Nothing Then
                Set foundCell = Cells.Find(What:=searchText, LookIn:=xlValues, LookAt:=xlPart)
            Else
                Set foundCell = Cells.FindNext(After:=lastFoundCell)
            End If
            
            ' 値検索でヒットしなかった場合は、コメント検索モードへ切り替えてループを回す
            If foundCell Is Nothing Then
                isSearchingComments = True
                Set lastFoundCell = Nothing
            Else
                Exit Do ' 見つかったら抜ける
            End If
        End If
        
        If isSearchingComments Then
            ' ▼ メモ（コメント）を検索（LookIn:=xlComments）
            If lastFoundCell Is Nothing Then
                Set foundCell = Cells.Find(What:=searchText, LookIn:=xlComments, LookAt:=xlPart)
            Else
                Set foundCell = Cells.FindNext(After:=lastFoundCell)
            End If
            Exit Do ' コメント検索が終了したらループを抜ける
        End If
    Loop

    ' 3. ヒットした場合の処理
    If Not foundCell Is Nothing Then
        Set lastFoundCell = foundCell
        
        ' シートとセルをアクティブにする
        foundCell.Parent.Activate
        foundCell.Select
        
        ' 場所を強調する（点滅）
        Call FlashCell(foundCell)
        
        ' --- 【修正箇所】右隣へのジャンプ制御 ---
        ' メモ（コメント）内でヒットした場合は、そのままのセルをアクティブにする
        If isSearchingComments Then
            ' 何もしない（foundCellのまま）
            ' コメント表示を明示的に出したい場合は、以下を有効化しても良いです
            ' foundCell.Comment.Visible = True
        Else
            ' 値検索でヒットし、日本語検索、かつアクティブシートが「My Favorite」ではない場合のみ右へ飛ぶ
            If Not isHalfWidth And ActiveSheet.Name <> "My Favorite" Then
                On Error Resume Next
                foundCell.Offset(0, 1).Select
                On Error GoTo 0
            End If
        End If
        ' --------------------------------------
        
        ' フォーカスを検索窓に戻す
        Me.txtSearch.SetFocus
        Me.txtSearch.SelStart = 0
        Me.txtSearch.SelLength = Len(Me.txtSearch.text)
        
    Else
        ' セルの値にもメモにもヒットしなかった場合
        MsgBox "「" & searchText & "」は見つかりませんでした。" & vbCrLf & vbCrLf & _
               "'" & searchText & "' was not found.", vbExclamation, APP_NAME
               
        Set lastFoundCell = Nothing
        lastSearchText = ""
        isSearchingComments = False ' リセット
        Me.txtSearch.SetFocus
    End If
End Sub

' ==========================================================
' 4. その他のボタン（My Favoriteシート誤操作防衛・日英併記版）
' ==========================================================
Private Sub btnCopyNormal_Click()
    Dim targetText As String
    
    ' 0. ★追加：My Favoriteシートからの直接コピーを制限するバリデーション
    ' お気に入りのプロンプトは、メインウィンドウの「Favorite」タブから流し込むのが正規ルートです
    If ActiveSheet.Name = "My Favorite" Then
        MsgBox "「My Favorite」シートから直接コピーすることはできません。" & vbCrLf & _
               "お気に入りのプロンプトを追加する場合は、メインウィンドウの「Favorite」タブから操作してください。" & vbCrLf & vbCrLf & _
               "Direct copying from the 'My Favorite' sheet is disabled." & vbCrLf & _
               "To use favorite prompts, please operate from the 'Favorite' tab in the main window.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 1. 変換前に空白・エラーをチェックしてアラートを出す
    If IsError(ActiveCell.Value) Or Trim(CStr(ActiveCell.Value)) = "" Then
        MsgBox "選択されたセルは空白、またはエラー値です。" & vbCrLf & _
               "The selected cell is empty or contains an error.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 2. 値を取得して重み付けを適用
    targetText = Trim(CStr(ActiveCell.Value))
    targetText = GetWeightedValue(targetText)
    
    ' 3. 統合ロジックを呼び出す
    Call ExecutePromptLogic(targetText, True)
    
    ' 変更があったことを明示する
    'IsDirty = True
    'UpdateDoneButtonState
End Sub

' ==========================================
' アンドゥボタン：直前のプロンプト操作を取り消す
' ==========================================
Private Sub btnUndo_Click()
    ' ★超重要バグフィックス：Undoによるテキスト変更が「手入力」として誤検知されるのを防ぐ
    IsUpdatingBySystem = True
    
    ' 1. アンドゥ処理を実行
    Call PerformUndo
    
    ' 処理が終わったらシステムフラグを解除
    IsUpdatingBySystem = False
    
    ' 2. Undo履歴の残弾をチェックし、即座に同期
    'Call UpdateUndoButtonState
    
    ' 3. Doneボタンや各Clearボタンの有効/無効も一括同期
    'Call UpdateDoneButtonState
    
    ' ★UX向上：Undo後は再びテキストボックスにフォーカスを戻す（連続タイピングしやすくするため）
    On Error Resume Next
    If Me.txtMain.Visible And Me.txtMain.Enabled Then
        Me.txtMain.SetFocus
    End If
    On Error GoTo 0
End Sub


Private Sub btnClear_Click()
    ' =========================================================
    ' ★追加対策：テキストを消去する前に、現在の状態をアンドゥログに保存
    ' （※すでに空っぽの場合は無駄な履歴を積まないように条件分岐します）
    ' =========================================================
    If Trim(Me.txtMain.text) <> "" Then
        Call SaveHistory(Me.txtMain.text)
    End If

    Call ClearPrompt
    'Call UpdateUndoButtonState
    
End Sub

Private Sub btnAllClear_Click()
    Call AllClearPrompt
End Sub

Private Sub btnWrapBlock_Click()
    Call WrapLastBlock  '一括重み付けを呼び出す
End Sub
' ==========================================
' Favoritesタブ内検索（シンプル版）
' ==========================================
Private Sub btnFavSearch_Click()
    Dim wsFav As Worksheet
    Dim searchKey As String
    Dim i As Long, startIdx As Long, targetIdx As Long
    Dim currentCell As Range
    
    ' 門番：シートチェック
    If Not IsFavSheetActive() Then Exit Sub
    Set wsFav = ActiveSheet
    
    ' キーワード取得
    searchKey = Trim(Me.txtFavSearch.text)
    If searchKey = "" Or searchKey = PH_FAV_SEARCH Then
        MsgBox "検索キーワードを入力してください。" & vbCrLf & _
               "Please enter a search keyword.", vbExclamation, APP_NAME
        Exit Sub
    End If

    ' 検索開始位置の決定
    If searchKey <> lastSearchKeyFav Then
        startIdx = 1
    Else
        startIdx = (lastSearchIdxFav Mod 50) + 1
    End If
    
    ' 50件のループ検索
    For i = 0 To 49
        targetIdx = ((startIdx + i - 1) Mod 50) + 1
        On Error Resume Next
        ' No_n の1行下（Description行）をターゲットにする
        Set currentCell = wsFav.Range("No_" & targetIdx).Offset(1, 0)
        On Error GoTo 0
        
        If Not currentCell Is Nothing Then
            If InStr(1, CStr(currentCell.Value), searchKey, vbTextCompare) > 0 Then
                
                ' --- 検索ヒット時の挙動 ---
                ' 画面更新を止めずに自然に選択（ウィンドウ枠固定にお任せ）
                currentCell.Select
                
                ' 点滅演出の呼び出し
                Call FlashCell(currentCell.MergeArea)
                
                ' 状態保存
                lastSearchKeyFav = searchKey
                lastSearchIdxFav = targetIdx
                
                ' ステータスバー通知
                Application.StatusBar = "Found: " & targetIdx
                
                ' 検索窓にフォーカスを戻す
                Me.txtFavSearch.SetFocus
                Me.txtFavSearch.SelStart = 0
                Me.txtFavSearch.SelLength = Len(Me.txtFavSearch.text)
                
                Exit Sub ' ヒットしたので終了
            End If
        End If
    Next i ' ここで次のループへ
    
    ' --- 見つからなかった場合 ---
    MsgBox "キーワード「" & searchKey & "」は見つかりませんでした。" & vbCrLf & _
           "No match found for: " & searchKey, vbInformation, APP_NAME
    
    lastSearchKeyFav = ""
    lastSearchIdxFav = 0
    Application.StatusBar = False
End Sub




' --- Favoritesタブの検索窓でEnter ---
Private Sub txtFavSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Enterキー (13) が押されたかチェック
    If KeyCode = 13 Then
        ' Windows特有の「ポーン」という警告音を防ぐおまじない
        KeyCode = 0
        ' 検索ボタンのクリック処理を呼び出す
        Call btnFavSearch_Click
    End If
End Sub

' --- Cockpitタブの検索窓でもEnterを使えるようにする ---
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        Call btnSearch_Click ' Cockpit側の検索ボタン名に合わせてください
    End If
End Sub

' ==========================================
' Favoritesタブの検索窓をリセットする
' ==========================================
Private Sub btnClearSearchFav_Click()
    ' 1. テキストボックスの表示をプレースホルダーに戻す
    With Me.txtFavSearch
        .Value = PH_FAV_SEARCH
        .ForeColor = COLOR_GRAY
    End With
    
    ' 2. 検索の内部状態もリセット
    ' これにより、次回検索時に1件目から確実に探し直します
    lastSearchKeyFav = ""
    lastSearchIdxFav = 0
    
    ' 3. UX：フォーカスを窓に戻して、すぐに次の入力ができるようにする
    Me.txtFavSearch.SetFocus
    
    ' 4. ステータスバー通知
    Application.StatusBar = "Favorite search cleared and reset."
    
    ' 1秒後に表示を消す
    Application.Wait [Now() + "00:00:01"]
    Application.StatusBar = False
End Sub
' --- タブが切り替わった時の処理（全タブ対応版） ---
Private Sub MainWindow_Change()
    Select Case Me.MainWindow.Value
        
        Case 0 ' --- Cockpitタブ ---
            If Trim(Me.txtSearch.text) = "" Then
                Me.txtSearch.text = PH_TEXT
                Me.txtSearch.ForeColor = COLOR_GRAY
            End If
        
        Case 1 ' --- Positiveタブ ---
            Call UpdatePositiveUIState
        
        Case 2 ' --- Negativeタブ ---
            Call UpdateNegativeUIState
        
        Case 3 ' --- Loraタブ ---
            Call UpdateLoRAUIState
            
        Case 4 ' --- Favoritesタブ ---
            If Trim(Me.txtFavSearch.text) = "" Then
                Me.txtFavSearch.text = PH_FAV_SEARCH
                Me.txtFavSearch.ForeColor = COLOR_GRAY
            End If

        Case 5 ' --- Gacha! タブ ---
            Call UpdateGachaUI
            
    End Select
End Sub

' --- MainWindow (UserForm) 内、または共通モジュール ---

' ボタンやコントロールの状態と色味を「一括」で同期するプロシージャ
Private Sub SetControlUIState(ByRef ctrl As Object, ByVal isEnabled As Boolean, Optional ByVal activeColor As Long = -1)
    ' 定数定義（パステル調に合わせたグレーを選択）
    'Const COL_DISABLED_BACK As Long = &HE0E0E0 ' 統一された無効時背景（明るめのグレー）
    Const COL_DISABLED_FORE As Long = &H808080 ' 統一された無効時文字色（中間グレー）
    Const COL_DEFAULT_BACK  As Long = &H8000000F ' 標準のボタン表面色

    ctrl.Enabled = isEnabled

    If isEnabled Then
        ' 有効時
        If activeColor <> -1 Then
            ctrl.BackColor = activeColor ' 指定の色があれば適用
        Else
            ctrl.BackColor = COL_DEFAULT_BACK
        End If
        ctrl.ForeColor = &H0      ' 文字は黒
    Else
        ' 無効時（ここを完全に統一）
        'ctrl.BackColor = COL_DISABLED_BACK
        ctrl.ForeColor = COL_DISABLED_FORE
    End If
End Sub
' 新規作成したボタンの処理
Private Sub btnSetPositive_Click()
    Call SetPositiveAction
End Sub

'-------------------------------------------------------------------------
' LoRAデータが空の場合にUIをロックアウトする処理（JSON対応版）
'-------------------------------------------------------------------------
Public Sub LockLoRAUIIfEmpty()
    Dim isLoRAEmpty As Boolean
    Dim isPresetEmpty As Boolean
    Dim i As Integer
    Dim colGray As Long
    
    ' グレーアウト用のシステムカラー（chkUseLoRA未チェック時と同じ）
    colGray = &H8000000F
    
    ' 1. JSON（メモリ）から状態判定
    ' デフォルトは「空」として扱う
    isLoRAEmpty = True
    isPresetEmpty = True
    
    If Not ConfigData Is Nothing Then
        ' LoRA_List の判定
        If ConfigData.Exists("LoRA_List") Then
            If ConfigData("LoRA_List").Count > 0 Then isLoRAEmpty = False
        End If
        
        ' LoRA_Presets の判定
        If ConfigData.Exists("LoRA_Presets") Then
            If ConfigData("LoRA_Presets").Count > 0 Then isPresetEmpty = False
        End If
    End If
    
    ' 2. 両方のデータが空の場合のみ、ロックアウト処理を実行
    If isLoRAEmpty And isPresetEmpty Then
        
        'btnOpenManageLoRA だけは生命線として有効化
        Me.btnOpenManageLoRA.Enabled = True
        
        'それ以外のボタン・チェックボックスを無効化
        Me.chkUseLoRA.Enabled = False
        Me.lblNotRequiredTrigger.Enabled = False
        Me.btnSetLoRA.Enabled = False
        Me.btnLoRACancel.Enabled = False
        Me.btnWrapLoRA.Enabled = False
        Me.btnRemoveLoRA.Enabled = False
        Me.btnForgetLoRA.Enabled = False
        Me.btnSendLoRAtoCockpit.Enabled = False
        Me.btnSendLoRAtoFav.Enabled = False
        Me.btnClearLoRAPreview.Enabled = False
        Me.btnSaveAsPresetLoRA.Enabled = False
        Me.btnCallLoRAPreset.Enabled = False
        Me.btnDeleteLoRAPreset.Enabled = False
        Me.btnGetLoRANegative.Enabled = False
        
        'トリガーワードのチェックボックス (1～10) をループで無効化
        For i = 1 To 10
            Me.Controls("chkTrigger_" & i).Enabled = False
        Next i
        
        'リスト・コンボ・テキストを無効化 ＆ グレーアウト（見た目の変更）
        With Me.cmbLoRAList
            .Enabled = False
            .BackColor = colGray
        End With
        
        With Me.cmbLoRAStrength
            .Enabled = False
            .BackColor = colGray
        End With
        
        With Me.lstSelectedLoRA
            .Enabled = False
            .BackColor = colGray
        End With
        
        With Me.cmbTriggerWeight
            .Enabled = False
            .BackColor = colGray
        End With
        
        With Me.cmbPresetList
            .Enabled = False
            .BackColor = colGray
        End With
        
        With Me.txtLoRAPreview
            .Enabled = False
            .BackColor = colGray
        End With
        
    End If
    
End Sub

'-------------------------------------------------------------------------
' メインウィンドウのLoRAコンボボックスを更新する処理（堅牢版）
'-------------------------------------------------------------------------
Private Sub RefreshMainWindowLoRA()
    Dim loraList As Object
    Dim loraItem As Object
    Dim aliasName As String
    Dim validCount As Long ' 正常に追加できた件数をカウント
    
    ' A. 一度コンボボックスを空にする
    Me.cmbLoRAList.Clear
    validCount = 0
    
    ' B. JSONデータがメモリ上に存在するかチェック
    If Not ConfigData Is Nothing Then
        If ConfigData.Exists("LoRA_List") Then
            Set loraList = ConfigData("LoRA_List")
            
            ' 登録データがある場合
            If loraList.Count > 0 Then
                
                ' コンボボックスにAliasを流し込む（安全装置付き）
                For Each loraItem In loraList
                    aliasName = ""
                    ' 悪意のあるJSON（キー欠落）でクラッシュしないように保護
                    On Error Resume Next
                    If loraItem.Exists("Alias") Then
                        aliasName = Trim(CStr(loraItem("Alias")))
                    End If
                    On Error GoTo 0
                    
                    ' 空白じゃなければ追加、空白ならダミー名で追加して存在だけはさせる
                    If aliasName <> "" Then
                        Me.cmbLoRAList.AddItem aliasName
                        validCount = validCount + 1
                    ElseIf loraItem.Exists("Hash") Then
                        ' Aliasが消されていてもHashが生きている場合への救済処置
                        Me.cmbLoRAList.AddItem "(Unnamed_" & Left(loraItem("Hash"), 5) & ")"
                        validCount = validCount + 1
                    End If
                Next loraItem
                
                ' 1件でも正常にコンボボックスに入った場合のみ有効化する
                If validCount > 0 Then
                    Me.chkUseLoRA.Enabled = True
                    Exit Sub
                End If
                
            End If
        End If
    End If
    
    ' C. 登録が1件もない、またはデータ自体が存在しない場合のフォールバック処理
    Me.chkUseLoRA.Value = False
    Me.chkUseLoRA.Enabled = False
End Sub

Private Sub btnFeelLucky_Click()
    ' --- 冷却用変数の追加 ---
    Dim cooldownTime As Integer: cooldownTime = 15
    Dim startTime As Single
    Dim originalCaption As String: originalCaption = "I'm Feeling Lucky"
    
    If dictJPtoEN Is Nothing Or dictENtoEN Is Nothing Then InitDictionaries
    
    Dim apiKey As String: apiKey = Me.txtAPIKey.text
    Dim apiResponse As String
    
    ' ========================================================
    ' ★変更箇所：userInput の動的生成（Gacha! 対応・最新DBインデックス版）
    ' ========================================================
    Dim userInput As String
    
    ' おまかせ機能（Surprise Me!）がONの場合
    If Me.chkSupriseMe.Value = True Then
        Randomize
        Dim wordsList As Object
        Set wordsList = CreateObject("System.Collections.ArrayList")
        
        ' 【全モード共通確定枠】
        Call DrawWord(3, 1#, wordsList)    ' (3)キャラ数
        Call DrawWord(34, 1#, wordsList)   ' (34)場所
        Call DrawWord(2, 1#, wordsList)   ' (2)アングル
        
        ' 【松竹梅モード別抽出】
        If Me.opbSFW.Value = True Then
            ' 梅 (SFW)
            Call DrawWord(16, 1#, wordsList)    ' (16)一般的行為
            If Rnd < 0.5 Then Call DrawWord(20, 1#, wordsList) Else Call DrawWord(22, 1#, wordsList)   ' (20)職業 or (22)服装
            Call DrawWord(9, 0.5, wordsList)    ' (9)髪の長さ(50%)
            Call DrawWord(12, 0.5, wordsList)   ' (12)髪の色(50%)
            Call DrawWord(26, 0.3, wordsList)   ' (26)ヘッドウェア(30%)
            Call DrawWord(28, 0.3, wordsList)   ' (28)足元周り(30%)
            Call DrawWord(29, 0.3, wordsList)   ' (29)アクセサリー(30%)
            Call DrawWord(13, 0.5, wordsList)   ' (13)表情(50%)
            Call DrawWord(36, 0.2, wordsList)   ' (36)その他アイテム(20%)
            
        ElseIf Me.opbNSFW.Value = True Then
            ' 竹 (NSFW)
            Call DrawWord(18, 1#, wordsList)    ' (18)性的行為
            If Rnd < 0.5 Then Call DrawWord(21, 1#, wordsList) Else Call DrawWord(23, 1#, wordsList)   ' (21)下着 or (23)服の状態
            Call DrawWord(15, 0.8, wordsList)   ' (15)体位(80%)
            Call DrawWord(31, 0.5, wordsList)   ' (31)身体の部位(50%)
            Call DrawWord(32, 0.5, wordsList)   ' (32)行為の状態(50%)
            Call DrawWord(14, 0.4, wordsList)   ' (14)体毛(40%)
            Call DrawWord(33, 0.5, wordsList)   ' (33)体液(50%)
            Call DrawWord(13, 0.8, wordsList)   ' (13)表情(80%)
            
        ElseIf Me.opbHardcore.Value = True Then
            ' 松 (Hardcore)
            Call DrawWord(19, 1#, wordsList)    ' (19)ボンデージ行為
            Call DrawWord(17, 0.8, wordsList)   ' (17)性的行為大分類(80%)
            Call DrawWord(30, 0.9, wordsList)   ' (30)手段・道具(90%)
            Call DrawWord(23, 0.8, wordsList)   ' (23)服の状態(80%)
            Call DrawWord(15, 0.8, wordsList)   ' (15)体位(80%)
            Call DrawWord(33, 0.8, wordsList)   ' (33)体液(80%)
            Call DrawWord(13, 0.9, wordsList)   ' (13)表情(90%)
        End If
        
        ' 抽出した単語をカンマで結合
        Dim i As Integer
        userInput = ""
        For i = 0 To wordsList.Count - 1
            If userInput = "" Then
                userInput = wordsList(i)
            Else
                userInput = userInput & ", " & wordsList(i)
            End If
        Next i
        
        ' ガチャで生成されたプロンプトをUIのテキストボックスにも反映してあげる（ユーザーが何が出たか見えるように）
        Me.txtRequest.text = userInput
        
    Else
        ' おまかせ機能がOFFの場合（手入力のプロンプトをそのまま使用）
        userInput = Trim(Me.txtRequest.text)
    End If
    
    ' 入力が空の場合は処理終了
    If userInput = "" Then Exit Sub
    
    ' バリデーション
    If Trim(userInput) = "" Or Trim(apiKey) = "" Then
        MsgBox "入力内容（欲望またはAPIキー）を確認してください。" & vbCrLf & _
               "Please check your input (desire or API key)!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' ==========================================================
    ' 1. 【物理封鎖】処理開始直後にハッチをロック
    ' ==========================================================
    m_IsGachaRunning = True
    Call SetCloseButtonState(False)
    
    ' UIの初期描画を済ませる
    Call SetControlUIState(Me.btnFeelLucky, False)
    Call SetControlUIState(Me.btnClearlRequest, False)
    Me.txtResult.text = "思考中... / Thinking... (Calling Gemini API)"
    Application.StatusBar = "API通信中... / Calling Gemini API..."
    
    ' ここで一度OSに描画を任せつつ、ロックが確実に適用されるよう待ちます
    Me.Repaint
    DoEvents
    
    ' 念押し：通信直前にもう一度ロックを確認
    Call SetCloseButtonState(False)
    
    ' --- 2. Geminiにリクエストを投げる ---
    apiResponse = GetKeywordsFromGemini(userInput, apiKey)
    
    ' --- 通信失敗（またはセーフガード）時の処理 ---
    If apiResponse = "" Then
        Me.txtResult.text = ""
        Application.StatusBar = "Canceled by Safety Filter."
        GoTo StartCooldown ' 冷却へ
    End If

    ' --- 3. 通信成功時の処理 ---
    Call IncrementGachaCount
    Me.lblGachaCount.Caption = GetGachaStatus()
    
    SaveSetting APP_NAME, "Settings", "GeminiAPIKey", apiKey
    
    ' ここで画面が大きく書き換わり、Windowsが勝手にXボタンを復活させてしまう！
    Me.txtResult.text = GeneratePrompt(apiResponse)
    
    Application.StatusBar = "Prompt Generated Successfully!"
    Call UpdateGachaUI(True)
    
    Me.txtRequest.Locked = False
    Me.txtRequest.BackColor = &H80000005

' ==========================================================
' 4. 冷却カウントダウン（鉄壁のシーリング）
' ==========================================================
StartCooldown:
    Application.Cursor = xlDefault
    
    ' ページの封鎖（マルチページ全体をロック）
    Me.MainWindow.Pages("pgCockpit").Enabled = False
    Me.MainWindow.Pages("pgFavorites").Enabled = False
    Me.MainWindow.Pages("pgLoRA").Enabled = False
    Me.MainWindow.Pages("pgNegative").Enabled = False
    Me.MainWindow.Pages("pgPositive").Enabled = False
    Me.MainWindow.Pages("pgIO").Enabled = False
    Me.MainWindow.Pages("pgMobile").Enabled = False
    
' その他、誤操作を防ぎたい入力欄も念のため封鎖
    Me.txtRequest.Enabled = False
    Call SetControlUIState(Me.btnFeelLucky, False)
    Call SetControlUIState(Me.btnClearlRequest, False)
    
    startTime = Timer
    Do While Timer < startTime + cooldownTime
        If Timer < startTime Then startTime = startTime - 86400
        
        ' カウントダウン更新
        Me.btnFeelLucky.Caption = "Cooling... (" & Int(startTime + cooldownTime - Timer + 1) & "s)"
        
        ' ==========================================================
        ' ★ 最終防衛線：DoEvents（操作の受付）の「直前」に必ずロックを再適用する！
        ' Windowsが勝手に復活させたXボタンを、クリック判定される前に無効化
        ' ==========================================================
        If m_IsGachaRunning Then Call SetCloseButtonState(False)
        
        ' ここでようやくOSに操作（クリックの残響など）を消化させる
        DoEvents
        
        Sleep 100
        If Application.Cursor <> xlDefault Then Application.Cursor = xlDefault
    Loop

    ' --- 5. UIの最終復帰処理 ---
    Me.MainWindow.Pages("pgCockpit").Enabled = True
    Me.MainWindow.Pages("pgFavorites").Enabled = True
    Me.MainWindow.Pages("pgLoRA").Enabled = True
    Me.MainWindow.Pages("pgNegative").Enabled = True
    Me.MainWindow.Pages("pgPositive").Enabled = True
    Me.MainWindow.Pages("pgIO").Enabled = True
    Me.MainWindow.Pages("pgMobile").Enabled = True
    
    Me.txtRequest.Enabled = True
    Me.btnFeelLucky.Caption = originalCaption
    Call SetControlUIState(Me.btnFeelLucky, True, &H80FF80)
    Call SetControlUIState(Me.btnClearlRequest, True, &H8080FF)
    
    ' --- 6. 【解錠】全ての静寂が終わった後の「最後の一手」 ---
    DoEvents
    
    On Error Resume Next
    Application.Cursor = xlAutomatic
    Call SetCloseButtonState(True) ' ここで初めてXボタンが復活
    m_IsGachaRunning = False
    On Error GoTo 0
    
    Application.StatusBar = False
End Sub

Private Sub btnClearlRequest_Click()
    Me.txtRequest.text = ""
End Sub

Private Sub btnCopyResult_Click()
    If Trim(Me.txtResult.text) <> "" Then
        With Me.txtResult
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.text)
        End With
        ' 自作の安全な関数を使用
        Call SetClipboardText(Me.txtResult.text)
        Application.StatusBar = "Result copied to clipboard."
    End If
    Call SetControlUIState(Me.btnCopyResult, False)

End Sub

' Cockpitへ送る
Private Sub btnSendFeeltoCockpit_Click()
    Dim resultPrompt As String
    Dim ans As VbMsgBoxResult
    Dim shouldAppend As Boolean: shouldAppend = True
    
    resultPrompt = Trim(Me.txtResult.text)
    If resultPrompt = "" Then Exit Sub
    
    ' 末尾をカンマで整える
    If Right(resultPrompt, 1) <> "," Then
        resultPrompt = resultPrompt & ","
    End If
    
    ' 既存内容がある場合の確認
    If Trim(Me.txtMain.text) <> "" Then
        ans = MsgBox("Cockpitに既に内容が存在します。新しいプロンプトを追記しますか？" & vbCrLf & _
                     "Content already exists in Cockpit. Do you want to append the new prompt?" & vbCrLf & vbCrLf & _
                     "【Yes】末尾に追記 (Append)" & vbCrLf & _
                     "【No】上書きする (Overwrite)" & vbCrLf & _
                     "【Cancel】中止 (Cancel)", _
                     vbYesNoCancel + vbQuestion, APP_NAME)
        
        Select Case ans
            Case vbYes:    shouldAppend = True
            Case vbNo:     shouldAppend = False
            Case vbCancel: Exit Sub
        End Select
    End If
    
    ' 上書き時は、まずテキストボックスをクリアする
    If Not shouldAppend Then
        Me.txtMain.text = ""
    End If
    
    ' 万能エンジンを通す
    Call ExecutePromptLogic(resultPrompt, shouldAppend)
    
    ' --- 安全なタブ切り替えとフォーカス処理 ---
    ' 1. タブをCockpitへ（MainWindowはMultiPageのオブジェクト名）
    Me.MainWindow.Value = 0
    
    ' 2. 画面の描画を強制的に更新して「見える状態」にする
    DoEvents
    
    ' 3. 安全にフォーカスを当てる（冷却中などで無効な場合は無視する）
    On Error Resume Next
    If Me.txtMain.Enabled And Me.txtMain.Visible Then
        Me.txtMain.SetFocus
    End If
    
    On Error GoTo 0
' --- ★追加：送信完了後、このボタンを「使用済み」として無効化 ---
    Call SetControlUIState(Me.btnSendFeeltoCockpit, False)

End Sub

' Favoritesへ送る
Private Sub btnSendFeeltoFav_Click()
    Dim resultPrompt As String, currentFav As String
    Dim ans As VbMsgBoxResult
    Dim shouldAppend As Boolean: shouldAppend = True
    
    resultPrompt = Trim(Me.txtResult.text)
    If resultPrompt = "" Then Exit Sub
    
    ' 末尾をカンマで整える
    If Right(resultPrompt, 1) <> "," Then
        resultPrompt = resultPrompt & ","
    End If
    
    currentFav = Trim(Me.txtFav.text)
    
    ' 既存内容がある場合の確認
    If currentFav <> "" Then
        ans = MsgBox("Favoritesに既に内容が存在します。新しいプロンプトを追記しますか？" & vbCrLf & _
                     "Content already exists in Favorites. Do you want to append the new prompt?" & vbCrLf & vbCrLf & _
                     "【Yes】末尾に追記 (Append)" & vbCrLf & _
                     "【No】上書きする (Overwrite)" & vbCrLf & _
                     "【Cancel】中止 (Cancel)", _
                     vbYesNoCancel + vbQuestion, APP_NAME)
        
        Select Case ans
            Case vbYes:    shouldAppend = True
            Case vbNo:     shouldAppend = False
            Case vbCancel: Exit Sub
        End Select
    End If
    
    ' 履歴保存
    Call SaveHistory(Me.txtMain.text)
    
    ' --- 条件分岐を整理 ---
    If shouldAppend And currentFav <> "" Then
        ' 追記かつ既存ありの場合のみ連結
        Me.txtFav.text = currentFav & vbCrLf & resultPrompt
    Else
        ' 上書き、または既存が空の場合はそのまま代入
        Me.txtFav.text = resultPrompt
    End If
    
    ' --- 【修正点】安全なタブ切り替えとフォーカス処理 ---
    ' 1. タブをFavoritesへ（Value = 1 は左から2番目のタブを指します）
    Me.MainWindow.Value = 4
    
    ' 2. 画面の描画を強制的に更新して「見える状態」にする
    DoEvents
    
    ' 3. 安全にフォーカスを当てる（冷却中などで無効な場合は無視する）
    On Error Resume Next ' 万が一の2110エラーを無視して処理を続行させる
    If Me.txtFav.Enabled And Me.txtFav.Visible Then
        Me.txtFav.SetFocus
    End If
    On Error GoTo 0 ' エラーハンドリングを通常に戻す

    ' --- ★追加：送信完了後、このボタンを「使用済み」として無効化 ---
    Call SetControlUIState(Me.btnSendFeeltoFav, False)
End Sub

Private Sub btnClearResult_Click()
    Me.txtResult.text = ""
    Call SetControlUIState(Me.btnClearResult, False)
    Call SetControlUIState(Me.btnCopyResult, False)
    Call SetControlUIState(Me.btnSendFeeltoCockpit, False)
    Call SetControlUIState(Me.btnSendFeeltoFav, False)
    

End Sub

' ラベルをフレーム中央に整列させる専用サブ
Private Sub AlignGachaLabel()
    With Me.lblGachaCount
        .AutoSize = True
        .Top = (Me.frmFrameGacha.InsideHeight / 2) - (.Height / 2)
        .Left = (Me.frmFrameGacha.InsideWidth / 2) - (.Width / 2)
    End With
End Sub
' ==========================================
' Gacha! タブのUI状態を一括管理する（Gacha! 対応版）
' ==========================================
Public Sub UpdateGachaUI(Optional isSuccess As Boolean = False)
    Dim hasKey As Boolean: hasKey = (Trim(Me.txtAPIKey.text) <> "")
    Dim hasRequest As Boolean: hasRequest = (Trim(Me.txtRequest.text) <> "")
    
    ' 1. APIキーの有無でリクエスト欄(txtRequest)を制御
    Me.txtRequest.Enabled = hasKey
    Me.txtRequest.Locked = Not hasKey
    ' &H80000005 はシステムカラーのためそのままでOK
    Me.txtRequest.BackColor = IIf(hasKey, &H80000005, colDisabled)

    ' --- ★新機能：Gacha! の準備ができているか判定 ---
    Dim isGachaReady As Boolean
    isGachaReady = False
    ' おまかせ機能がONで、かつ松竹梅のどれかが選ばれていればTrue
    If Me.chkSupriseMe.Value = True Then
        If Me.opbSFW.Value = True Or Me.opbNSFW.Value = True Or Me.opbHardcore.Value = True Then
            isGachaReady = True
        End If
    End If

    ' 2. リクエストの有無 ＆ Gacha!の状態で「錬成」系ボタンを制御
    ' APIキーがあり、かつ「手入力がある」または「ガチャの準備OK」なら錬成ボタンを有効化
    Dim canLucky As Boolean
    canLucky = hasKey And (hasRequest Or isGachaReady)
    
    ' ★末尾に & を付けて Long型として明示的に渡す
    Call SetControlUIState(Me.btnFeelLucky, canLucky, &H80FF80)
    
    ' ※微修正：クリアボタンは「テキスト欄に文字がある時」だけ押せるように分離しました
    Call SetControlUIState(Me.btnClearlRequest, hasRequest, &H8080FF)

    ' 3. 成功時のみ解放される出力系パーツの制御
    Dim hasResult As Boolean: hasResult = (Trim(Me.txtResult.text) <> "" Or isSuccess)
    
    ' &HD7FF& とすることで、正の数として認識させ、エラー380を回避します
    Call SetControlUIState(Me.btnSendFeeltoCockpit, hasResult, &HFFE0C0)
    Call SetControlUIState(Me.btnSendFeeltoFav, hasResult, &HFFE0C0)
    Call SetControlUIState(Me.btnCopyResult, hasResult, &HFFE0C0)
    Call SetControlUIState(Me.btnClearResult, hasResult, &H8080FF)
    
End Sub
'-------------------------------------------------------------------------
' [修正版] カテゴリジャンプの初期化（行継続文字を一切使用しない）
'-------------------------------------------------------------------------
Private Sub InitCategoryJump()
    Dim dataStr As String
    Dim items() As String
    Dim parts() As String
    Dim i As Integer
    
    ' 1. 1行ずつ変数に継ぎ足していく（これなら制限に引っかかりません）
    dataStr = "Jump to Category,"
    dataStr = dataStr & ";1.画風(Art Style),ArtStyle"
    dataStr = dataStr & ";2.アングル(Camera Angle),CameraAngle"
    dataStr = dataStr & ";3.キャラ数(Character Count),Count"
    dataStr = dataStr & ";4.キャラの役割(Character Roles),Roles"
    dataStr = dataStr & ";5.関係性(Relationship),Relationship"
    dataStr = dataStr & ";6.肌の色・属性(Skin & Attributes),Skin"
    dataStr = dataStr & ";7.体型(Body Type),BodyType"
    dataStr = dataStr & ";8.髪の毛全てへのワイルドカード(Wildcard for Hair),WildcardforHair"
    dataStr = dataStr & ";9.髪の長さ(Hair length),HairLength"
    dataStr = dataStr & ";10.前髪(Bangs),Bangs"
    dataStr = dataStr & ";11.髪の結び目(Tying),Tying"
    dataStr = dataStr & ";12.髪の色(Hair Color),HairColor"
    dataStr = dataStr & ";13.瞳の色(Eye Color),EyeColor"
    dataStr = dataStr & ";14.表情(Expressions),Expressions"
    dataStr = dataStr & ";15.体毛(Body Hair),BodyHair"
    dataStr = dataStr & ";16.体位(Position),Position"
    dataStr = dataStr & ";17.一般的行為(General Action),GeneralAction"
    dataStr = dataStr & ";18.性的行為大分類(Major Categories of Sexual Act),MajorSexualAct"
    dataStr = dataStr & ";19.性的行為(Sexual Act),SexualAct"
    dataStr = dataStr & ";20.ボンデージ行為(Bondage Action & Movement),BondageAction"
    dataStr = dataStr & ";21.職業(Occupation),Occupation"
    dataStr = dataStr & ";22.下着(Underwear),Underwear"
    dataStr = dataStr & ";23.服装(Outfit),Outfit"
    dataStr = dataStr & ";24.服の状態(Outfit State),OutfitState"
    dataStr = dataStr & ";25.服の色 (Clothing Color),ClothingColor"
    dataStr = dataStr & ";26.服の素材 (Clothing Material),ClothingMaterial"
    dataStr = dataStr & ";27.ヘッドウェア(Headwear),Headwear"
    dataStr = dataStr & ";28.手元周り(Hands & Wrists),Wrists"
    dataStr = dataStr & ";29.足元周り(Footwear & Legwear),Footwear"
    dataStr = dataStr & ";30.アクセサリー類(Accessories),Accessories"
    dataStr = dataStr & ";31.手段・道具(Means & Props),Means"
    dataStr = dataStr & ";32.身体の部位(Body Parts),BodyParts"
    dataStr = dataStr & ";33.行為の状態(Interaction State),InteractionState"
    dataStr = dataStr & ";34.体液(Body fluids),BodyFluids"
    dataStr = dataStr & ";35.場所(Location),Location"
    dataStr = dataStr & ";36.時間帯・周囲の状況(Time & Surroundings),Time"
    dataStr = dataStr & ";37.その他アイテム(Misc Items),Items"
    dataStr = dataStr & ";38.光源(Lighting),Lighting"
    dataStr = dataStr & ";39.効果(Effects),Effects"
    dataStr = dataStr & ";40.修正 / その他(Censorship Fixes & Others),Censorship"

    ' 2. セミコロンで分割して配列にする
    items = Split(dataStr, ";")
    
    ' 3. コンボボックスの基本設定
    With Me.cmbCategoryJump
        .Clear
        .ColumnCount = 2
        ' 1列目（表示名）を表示し、2列目（ネームドセル名）を隠す
        '.ColumnWidths = .Width & ";0"
        .ColumnWidths = (.Width - 15) & ";0"
        
        ' 4. 配列をループして流し込む
        For i = 0 To UBound(items)
            parts = Split(items(i), ",")
            
            .AddItem parts(0) ' 1列目: 表示名
            If UBound(parts) > 0 Then
                .List(i, 1) = parts(1) ' 2列目: ネームドセル名
            Else
                .List(i, 1) = ""
            End If
        Next i
        
        .Value = .List(0, 0) ' 初期値をセット
    End With
End Sub

' ==========================================
' 1. 項目移動：ストック (All) → 適用 (Apply)
' ==========================================
'-------------------------------------------------------------------------
' [ → ] ボタン：選択したストックを適用中リスト(Apply)へ送る (完全対称版)
'-------------------------------------------------------------------------
Private Sub btnApplyNegative_Click()
    Dim i As Long
    Dim word As String
    Dim isDuplicate As Boolean
    Dim k As Long
    Dim addedCount As Long
    
    addedCount = 0
    
    ' ストックリストの選択されたアイテムを上から順に走査
    For i = 0 To Me.lstAllNegative.ListCount - 1
        If Me.lstAllNegative.Selected(i) Then
            word = Me.lstAllNegative.List(i)
            
            ' 適用中リスト(lstApplyNegative)に既に存在しないか重複チェック
            isDuplicate = False
            For k = 0 To Me.lstApplyNegative.ListCount - 1
                If StrComp(Me.lstApplyNegative.List(k), word, vbTextCompare) = 0 Then
                    isDuplicate = True
                    Exit For
                End If
            Next k
            
            ' 重複していなければ適用中リストに追加（★ストック側からのRemoveItemは行わない）
            If Not isDuplicate Then
                Me.lstApplyNegative.AddItem word
                addedCount = addedCount + 1
            End If
            
            ' 選択状態を解除（ポジティブ側のUXに合わせる場合）
            Me.lstAllNegative.Selected(i) = False
        End If
    Next i
    
    ' 1つでも追加されたら全体の出力テキスト（txtYourNegative）を再ビルドしてUIを更新
    If addedCount > 0 Then
        Call UpdateNegativeUIState
    End If
End Sub

' =========================================================================
' [適用解除(<-)] ボタン：適用中リストから選んだ項目を外す（ネガティブ側）
' =========================================================================
Private Sub btnDismissNegative_Click()
    Dim i As Integer, k As Integer
    Dim hasDeleted As Boolean
    Dim hasAddedToStock As Boolean
    Dim word As String
    Dim isDuplicate As Boolean
    
    If Me.lstApplyNegative.ListIndex = -1 Then Exit Sub
    
    For i = Me.lstApplyNegative.ListCount - 1 To 0 Step -1
        If Me.lstApplyNegative.Selected(i) Then
            ' ★ストックに回収・比較する前に、ウェイトやカッコを剥がして純粋なタグに戻す
            word = GetPureTagName(Me.lstApplyNegative.List(i))
            
            ' ストック側に既に存在するか重複チェック
            isDuplicate = False
            For k = 0 To Me.lstAllNegative.ListCount - 1
                If StrComp(Me.lstAllNegative.List(k), word, vbTextCompare) = 0 Then
                    isDuplicate = True
                    Exit For
                End If
            Next k
            
            ' ストックに存在しない「未知のタグ」なら自動的にストックへ回収
            If Not isDuplicate Then
                Me.lstAllNegative.AddItem word
                hasAddedToStock = True
            End If
            
            ' 適用中リストからは削除
            Me.lstApplyNegative.RemoveItem i
            hasDeleted = True
        End If
    Next i
    
    ' ストックへの自動追加が発生した場合のみ、JSON（Stock）へ同期保存
    If hasAddedToStock Then Call SyncNegativeSheet
    
    ' 削除が発生した場合、出力テキストの再構築とUIの更新を行う
    If hasDeleted Then
        Call RefreshNegativeOutputText ' ★新設：適用リストの変更を出力テキストに即時同期
        Call UpdateNegativeUIState
        
        ' ステータスバーのメッセージを状況に応じて切り替え
        If hasAddedToStock Then
            Application.StatusBar = "● Removed from active list (New tags auto-added to stock)."
        Else
            Application.StatusBar = "● Removed selected item(s) from the active list."
        End If
    End If
End Sub


' ==========================================
' 4. ソート：適用側 (Apply) 上下移動
' ==========================================
Private Sub btnApplyNegativeSortUp_Click()
    Call MoveListItem(Me.lstApplyNegative, -1)
End Sub

Private Sub btnApplyNegativeSortDown_Click()
    Call MoveListItem(Me.lstApplyNegative, 1)
End Sub

' -------------------------------------------------------------------------
' 共通ヘルパー：リスト項目の移動ロジック（MultiSelect完全対応・誤爆防止版）
' -------------------------------------------------------------------------
Private Sub MoveListItem(ByRef lst As MSForms.ListBox, direction As Integer)
    Dim i As Integer
    Dim selCount As Integer
    Dim selIndex As Integer
    Dim temp As String

    ' 1. 「本当に選択（ハイライト）されているアイテム」を数え、そのインデックスを取得
    selCount = 0
    selIndex = -1
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            selCount = selCount + 1
            selIndex = i
        End If
    Next i

    ' 2. バリデーション：未選択なら完全に無視（何も起きない）
    If selCount = 0 Then Exit Sub
    
    ' 3. バリデーション：複数選択されている場合はバグの元なのでブロック
    If selCount > 1 Then
        MsgBox "並び替えは1件ずつ行ってください。" & vbCrLf & _
               "Please select only one item to sort.", vbExclamation, APP_NAME
        Exit Sub
    End If

    ' 4. 移動不能な位置（一番上をさらに上へ、一番下をさらに下へ等）なら終了
    If direction = -1 And selIndex = 0 Then Exit Sub
    If direction = 1 And selIndex = lst.ListCount - 1 Then Exit Sub

    ' 5. 値の入れ替え
    temp = lst.List(selIndex)
    lst.List(selIndex) = lst.List(selIndex + direction)
    lst.List(selIndex + direction) = temp

    ' 6. 選択状態とフォーカス（点線枠）を移動先へ追従させる
    lst.Selected(selIndex) = False
    lst.Selected(selIndex + direction) = True
    lst.ListIndex = selIndex + direction
End Sub




' --- 各種イベントでの呼び出し ---
Private Sub lstAllNegative_Change()
    Call UpdateNegativeUIState
End Sub

Private Sub lstApplyNegative_Change()
    Call UpdateNegativeUIState
End Sub

Private Sub chkWeightNegative_Change()
    Call UpdateNegativeUIState
End Sub

Private Sub txtYourNegative_Change()
    Call UpdateNegativeUIState
End Sub

' ==========================================
' 2. ボタン挙動：文字列生成と重み付け
' ==========================================

Private Sub btnAddNegative_Click()
    Dim i As Integer
    Dim currentText As String
    Dim wordsToAdd As String
    
    currentText = Trim(Me.txtYourNegative.text)
    wordsToAdd = ""
    
    For i = 0 To Me.lstApplyNegative.ListCount - 1
        If Me.lstApplyNegative.Selected(i) Then
            ' 重複チェック
            If InStr(1, currentText, Me.lstApplyNegative.List(i), vbTextCompare) = 0 And _
               InStr(1, wordsToAdd, Me.lstApplyNegative.List(i), vbTextCompare) = 0 Then
                
                If wordsToAdd <> "" Then wordsToAdd = wordsToAdd & ", "
                wordsToAdd = wordsToAdd & Me.lstApplyNegative.List(i)
            End If
            
            ' ★追加：処理した項目の選択を解除
            Me.lstApplyNegative.Selected(i) = False
        End If
    Next i
    
    If wordsToAdd <> "" Then
        ' 追記ロジック
        If currentText = "" Then
            Me.txtYourNegative.text = wordsToAdd
        Else
            Me.txtYourNegative.text = currentText & IIf(Right(currentText, 1) = ",", " ", ", ") & wordsToAdd
        End If
    End If
    
    Call UpdateNegativeUIState ' 状態を再判定
End Sub

'-------------------------------------------------------------------------
' [Weighten & Add] ボタン：選択項目を重み付けして出力欄へ送り、UIをリセット
'-------------------------------------------------------------------------
Private Sub btnWeightenNegative_Click()
    ' ★追加：念のためのガードレール（チェックがオフなら強制終了）
    If Not Me.chkWeightNegative.Value Then Exit Sub

    Dim i As Integer
    Dim weight As String
    Dim combinedWords As String
    Dim currentText As String
    
    Dim rawWord As String
    Dim cleanWord As String
    Dim regEx As Object
    
    ' ★追加：既存の重み付けを剥がすための正規表現エンジンを準備
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    weight = Me.cmbNegativeWeight.Value
    If weight = "" Then weight = "1.1"
    
    combinedWords = ""
    For i = 0 To Me.lstApplyNegative.ListCount - 1
        If Me.lstApplyNegative.Selected(i) Then
            rawWord = Me.lstApplyNegative.List(i)
            
            ' ★スマート・オーバーライド処理：既存のカッコと数値を一旦破壊する
            cleanWord = rawWord
            regEx.Pattern = "\s*:\s*\d+(\.\d+)?"  ' ":1.4" などの数値を検出
            cleanWord = regEx.Replace(cleanWord, "")
            regEx.Pattern = "[\(\)\[\]\{\}]"      ' 各種カッコを検出
            cleanWord = regEx.Replace(cleanWord, "")
            cleanWord = Trim(cleanWord)           ' 前後の空白を掃除
            
            If combinedWords <> "" Then
                combinedWords = combinedWords & ", " & cleanWord
            Else
                combinedWords = cleanWord
            End If
            
            ' 処理した項目の選択を解除
            Me.lstApplyNegative.Selected(i) = False
        End If
    Next i
    
    If combinedWords = "" Then Exit Sub
    
    ' 新しい重み付け文字列の作成（ここで剥がしたタグを新しいweightで包み直す）
    Dim weightedResult As String
    weightedResult = "(" & combinedWords & ":" & weight & ")"
    currentText = Trim(Me.txtYourNegative.text)
    
    ' 重複チェックと出力
    If InStr(1, currentText, weightedResult, vbTextCompare) = 0 Then
        If currentText = "" Then
            Me.txtYourNegative.text = weightedResult
        Else
            Me.txtYourNegative.text = currentText & IIf(Right(currentText, 1) = ",", " ", ", ") & weightedResult
        End If
    End If
    
    ' 重み付けの適用が終わったので、チェックボックスを強制的にオフに戻す
    Me.chkWeightNegative.Value = False
    
    ' 全体のUI状態を再判定（これでコンボボックスもDisabledになります）
    Call UpdateNegativeUIState
    
    Application.StatusBar = "Weighted negative prompt added and UI reset."
End Sub

' テキストボックスで Enterキー が押された時
Private Sub txtInputNegative_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then ' Enter
        KeyCode = 0      ' ビープ音防止
        Call AddNewNegativeWord
    End If
End Sub


' [Clear Input] ボタン：入力欄をプレースホルダーに戻す
Private Sub btnClearNegaInput_Click()
    With Me.txtInputNegative
        .text = PH_NEGA_INPUT
        .ForeColor = COLOR_GRAY
    End With
End Sub

'-------------------------------------------------------------------------
' [All Dismiss] ボタン：適用中リストをすべて空にする (ネガティブ側)
'-------------------------------------------------------------------------
Private Sub btnAllNegativeDismiss_Click()
    Dim ans As VbMsgBoxResult
    Dim i As Integer, k As Integer
    Dim hasAddedToStock As Boolean
    Dim word As String
    Dim isDuplicate As Boolean
    
    If Me.lstApplyNegative.ListCount = 0 Then Exit Sub
    
    ' 確認ダイアログ（誤爆防止のため「いいえ」をデフォルトに設定）
    ans = MsgBox("適用リストのすべての項目をクリア（解除）しますか？" & vbCrLf & _
                 "Are you sure you want to dismiss all active items?", _
                 vbYesNo + vbQuestion + vbDefaultButton2, APP_NAME)
                 
    If ans = vbNo Then Exit Sub
    
    ' ★全消しする前に、未知のタグがあればストックへ一括回収する
    For i = 0 To Me.lstApplyNegative.ListCount - 1
        ' ウェイトやカッコを剥がして純粋なタグに戻す
        word = GetPureTagName(Me.lstApplyNegative.List(i))
        
        isDuplicate = False
        For k = 0 To Me.lstAllNegative.ListCount - 1
            If StrComp(Me.lstAllNegative.List(k), word, vbTextCompare) = 0 Then
                isDuplicate = True
                Exit For
            End If
        Next k
        
        If Not isDuplicate Then
            Me.lstAllNegative.AddItem word
            hasAddedToStock = True
        End If
    Next i
    
    ' ストックへの自動追加が発生した場合のみ、JSON（Stock）へ同期保存
    If hasAddedToStock Then Call SyncNegativeSheet
    
    ' 適用リストを一括クリア
    Me.lstApplyNegative.Clear
    
    ' ★出力テキストの再構築（空になる）とUI状態の更新
    Call RefreshNegativeOutputText
    Call UpdateNegativeUIState
    
    If hasAddedToStock Then
        Application.StatusBar = "● All items dismissed (New tags auto-added to stock)."
    Else
        Application.StatusBar = "● All negative active items dismissed."
    End If
End Sub

' [Clear Output] ボタン：出力欄の消去
Private Sub btnClearNegative_Click()
    Me.txtYourNegative.text = ""
    Call UpdateNegativeUIState ' UI状態（コピーボタン等の無効化）を同期
End Sub

' [Copy] ボタン：全選択演出とクリップボード転送
Private Sub btnCopyNegative_Click()
    Dim targetText As String
    targetText = Trim(Me.txtYourNegative.text)
    
    If targetText = "" Then Exit Sub
    
    ' 1. 全選択の視覚的演出
    With Me.txtYourNegative
        .SetFocus
        .SelStart = 0
        .SelLength = Len(targetText)
    End With
    DoEvents ' 描画を強制して選択状態を画面に反映
    
    ' 2. クリップボードへの転送（既存の共通関数を使用）
    Call SetClipboardText(targetText)
    
    ' 3. 通知
    Application.StatusBar = "Negative prompt copied to clipboard!"
    Application.Wait [Now() + "00:00:01"]
    Application.StatusBar = False
End Sub

' --- 通信中の強制終了をブロックするガードレール（サイレント・キル版） ---
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' CloseMode = 0 は、ユーザーが右上の「X」ボタンを押したことを指します
    If CloseMode = 0 Then
        ' 確実な内部フラグで判定する
        If m_IsGachaRunning = True Then
            ' ★修正：MsgBoxを出すと、連打されたキューが消化される時に
            ' マクロが一時停止して「固まった」ように感じるため、警告ダイアログを廃止。
            ' 代わりにシステム音だけで「操作無効」を伝えます。
            Beep
            Cancel = True ' 閉じる動作を強制キャンセル（画面はそのまま進む）
        End If
    End If
End Sub


' ==========================================================
' 【UI制御】フォームの「X」ボタン（閉じる）を物理的に有効/無効化する
' ==========================================================
Private Sub SetCloseButtonState(ByVal bEnabled As Boolean)
    Dim hwnd As LongPtr
    Dim hMenu As LongPtr
    
    ' すでに宣言されているFindWindowを使ってフォームのハンドル（認識番号）を取得
    #If VBA7 Then
        hwnd = FindWindow("ThunderDFrame", Me.Caption)
    #Else
        hwnd = FindWindow("ThunderDFrame", Me.Caption)
    #End If
    
    If hwnd <> 0 Then
        hMenu = GetSystemMenu(hwnd, 0) ' ウィンドウのシステムメニューを取得
        If hMenu <> 0 Then
            If bEnabled Then
                ' 有効化（元の状態）
                EnableMenuItem hMenu, SC_CLOSE, MF_BYCOMMAND Or MF_ENABLED
            Else
                ' 無効化（グレーアウトさせて押せなくする）
                EnableMenuItem hMenu, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED
            End If
            ' 描画を強制更新して即座に見た目に反映させる
            DrawMenuBar hwnd
        End If
    End If
End Sub
'-------------------------------------------------------------------------
' [Save Preset] ボタン：ポジティブプロンプトをプリセットとして保存（JSON版）
'-------------------------------------------------------------------------
Private Sub btnSaveAsPrtesetPositive_Click()
    Dim presetName As String
    Dim promptValue As String
    Dim defaultPrompt As String
    Dim presets As Object
    Dim ans As Integer
    Dim defaultInput As String
    Dim k As Variant
    
    ' 1. 保存対象のプロンプトとデフォルト値を取得
    promptValue = Trim(Me.txtYourPositive.text)
    defaultPrompt = Trim(ConfigData("Positive_Default"))
    
    ' 先行バリデーション①：保存する中身自体が空欄なら弾く
    If promptValue = "" Then
        MsgBox "出力欄（txtYourPositive）が空欄です。保存する内容がありません。" & vbCrLf & _
               "The prompt field is empty.", vbExclamation, APP_NAME
        Me.txtYourPositive.SetFocus
        Exit Sub
    End If
    
    ' 先行バリデーション②：デフォルトのBase呪文との「順不同一致」をチェック
    If IsSamePromptIgnoreOrder(promptValue, defaultPrompt) Then
        MsgBox "入力されたプロンプトは、デフォルトのBase呪文と全く同じ構成です。" & vbCrLf & _
               "別のお気に入り構成を登録してください。" & vbCrLf & vbCrLf & _
               "The entered prompt consists of the exact same tags as the default Base spell.", vbExclamation, APP_NAME
        Me.txtYourPositive.SetFocus
        Exit Sub
    End If
    
    ' 先行バリデーション③：既存プリセットとの「順不同重複」を鉄壁ブロック
    Set presets = ConfigData("Positive_Presets")
    
    For Each k In presets.Keys
        If IsSamePromptIgnoreOrder(promptValue, presets(k)) Then
            MsgBox "入力されたプロンプトは、既にプリセット「" & k & "」と全く同じタグ群で構成されています。" & vbCrLf & _
                   "重複する構成を複数登録することはできません。内容を変更してください。" & vbCrLf & vbCrLf & _
                   "The entered prompt has the exact same tag composition as the existing preset '" & k & "'." & vbCrLf & _
                   "You cannot register multiple presets with identical content.", vbExclamation, APP_NAME
            Me.txtYourPositive.SetFocus
            Exit Sub
        End If
    Next k

    ' ダイアログの初期入力値の決定
    defaultInput = Trim(Me.cmbSelectPosiPreset.text)
    If UCase(defaultInput) = "DEFAULT" Then defaultInput = ""
    
    ' =========================================================
    ' 日英併記の入力ダイアログ（InputBox）の呼び出し
    ' =========================================================
    presetName = InputBox("ポジティブプリセットの登録名を入力してください。" & vbCrLf & _
                          "Please enter the registration name for the positive preset.", _
                          "Save Positive Preset", defaultInput)
    
    presetName = Trim(presetName)
    
    ' バリデーション④：空欄・キャンセルチェック
    If presetName = "" Then Exit Sub
    
    ' バリデーション⑤：「Default」という文字列の含有を完全に禁止
    If InStr(UCase(presetName), "DEFAULT") > 0 Then
        MsgBox "「Default」という文字列を含む名前はシステム専用、または混同防止のため、プリセット名として使用できません。" & vbCrLf & _
               "The name containing 'Default' is reserved for the system and cannot be used as a preset name.", _
               vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' =========================================================
    ' JSONへの書き込み・上書き処理
    ' =========================================================
    ' 同名プリセットの上書きチェック
    If presets.Exists(presetName) Then
        ans = MsgBox("「" & presetName & "」は既に存在します。上書きしますか？" & vbCrLf & _
                     "'" & presetName & "' already exists. Overwrite?", _
                     vbYesNo + vbQuestion + vbDefaultButton2, APP_NAME)
        If ans = vbNo Then Exit Sub
    End If
    
    ' JSONデータ（メモリ上）へ格納・上書き
    If presets.Exists(presetName) Then
        presets(presetName) = promptValue
    Else
        presets.Add presetName, promptValue
    End If
    
    ' 物理JSONファイルへ即時書き出し
    Call SaveConfigJSON
    
    ' コンボボックスのリストを再更新
    Call LoadPositivePresetsToCombo(Me.cmbSelectPosiPreset)
    Me.cmbSelectPosiPreset.text = presetName
    
    ' 成功メッセージ
    MsgBox "プリセット「" & presetName & "」を保存しました！" & vbCrLf & _
           "Preset '" & presetName & "' saved successfully!", vbInformation, APP_NAME
End Sub

'-------------------------------------------------------------------------
' [Save Preset] ボタン：ネガティブプロンプトをプリセットとして保存（JSON版）
'-------------------------------------------------------------------------
Private Sub btnSaveAsPresetNegative_Click()
    Dim presetName As String
    Dim promptValue As String
    Dim defaultPrompt As String
    Dim presets As Object
    Dim ans As Integer
    Dim defaultInput As String
    Dim k As Variant
    
    ' 1. 保存対象のプロンプトとデフォルト値を取得
    promptValue = Trim(Me.txtYourNegative.text)
    defaultPrompt = Trim(ConfigData("Negative_Default"))
    
    ' 先行バリデーション①：保存する中身自体が空欄なら弾く
    If promptValue = "" Then
        MsgBox "出力欄（txtYourNegative）が空欄です。保存する内容がありません。" & vbCrLf & _
               "The prompt field is empty.", vbExclamation, APP_NAME
        Me.txtYourNegative.SetFocus
        Exit Sub
    End If
    
    ' 先行バリデーション②：デフォルトのBase呪文との「順不同一致」をチェック
    If IsSamePromptIgnoreOrder(promptValue, defaultPrompt) Then
        MsgBox "入力されたプロンプトは、デフォルトのBase呪文と全く同じ構成です。" & vbCrLf & _
               "別のお気に入り構成を登録してください。" & vbCrLf & vbCrLf & _
               "The entered prompt consists of the exact same tags as the default Base spell.", vbExclamation, APP_NAME
        Me.txtYourNegative.SetFocus
        Exit Sub
    End If
    
    ' 先行バリデーション③：既存プリセットとの「順不同重複」を鉄壁ブロック
    Set presets = ConfigData("Negative_Presets")
    
    For Each k In presets.Keys
        If IsSamePromptIgnoreOrder(promptValue, presets(k)) Then
            MsgBox "入力されたプロンプトは、既にプリセット「" & k & "」と全く同じタグ群で構成されています。" & vbCrLf & _
                   "重複する構成を複数登録することはできません。内容を変更してください。" & vbCrLf & vbCrLf & _
                   "The entered prompt has the exact same tag composition as the existing preset '" & k & "'." & vbCrLf & _
                   "You cannot register multiple presets with identical content.", vbExclamation, APP_NAME
            Me.txtYourNegative.SetFocus
            Exit Sub
        End If
    Next k

    ' ダイアログの初期入力値の決定
    defaultInput = Trim(Me.cmbSelectNegaPreset.text)
    If UCase(defaultInput) = "DEFAULT" Then defaultInput = ""
    
    ' =========================================================
    ' 日英併記の入力ダイアログ（InputBox）の呼び出し
    ' =========================================================
    presetName = InputBox("ネガティブプリセットの登録名を入力してください。" & vbCrLf & _
                          "Please enter the registration name for the negative preset.", _
                          "Save Negative Preset", defaultInput)
    
    presetName = Trim(presetName)
    
    ' バリデーション④：空欄・キャンセルチェック
    If presetName = "" Then Exit Sub
    
    ' バリデーション⑤：「Default」という文字列の含有を完全に禁止
    If InStr(UCase(presetName), "DEFAULT") > 0 Then
        MsgBox "「Default」という文字列を含む名前はシステム専用、または混同防止のため、プリセット名として使用できません。" & vbCrLf & _
               "The name containing 'Default' is reserved for the system and cannot be used as a preset name.", _
               vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' =========================================================
    ' JSONへの書き込み・上書き処理
    ' =========================================================
    ' 同名プリセットの上書きチェック
    If presets.Exists(presetName) Then
        ans = MsgBox("「" & presetName & "」は既に存在します。上書きしますか？" & vbCrLf & _
                     "'" & presetName & "' already exists. Overwrite?", _
                     vbYesNo + vbQuestion + vbDefaultButton2, APP_NAME)
        If ans = vbNo Then Exit Sub
    End If
    
    ' JSONデータ（メモリ上）へ格納・上書き
    If presets.Exists(presetName) Then
        presets(presetName) = promptValue
    Else
        presets.Add presetName, promptValue
    End If
    
    ' 物理JSONファイルへ即時書き出し
    Call SaveConfigJSON
    
    ' コンボボックスのリストを再更新
    Call LoadNegativePresetsToCombo(Me.cmbSelectNegaPreset)
    Me.cmbSelectNegaPreset.text = presetName
    
    ' 成功メッセージ
    MsgBox "プリセット「" & presetName & "」を保存しました！" & vbCrLf & _
           "Preset '" & presetName & "' saved successfully!", vbInformation, APP_NAME
End Sub



' --- E. Civitai等からの「一括コピペ登録＆自動クレンジング」ボタン ---
' ブラウザからカッコ付きのまま持ってきた呪文のウェイト表記「(tags:1.2)」を「tags」にサニタイズして取り込む
Private Sub btnBulkImportPositive_Click()
    Dim rawInput As String
    Dim cleanedOutput As String
    Dim i As Long
    Dim char As String
    Dim inWeight As Boolean
    
    ' ユーザーが貼り付けた生データをテキストボックスなどから取得（またはInputBox）
    rawInput = Trim(Me.txtYourPositive.text)
    
    If rawInput = "" Then
        MsgBox "インポートするプロンプトをテキストボックスに入力（貼り付け）してください。" & vbCrLf & _
               "Please paste the prompt into the textbox first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' --- 高精度サニタイズ（クレンジング）ロジック ---
    ' カッコ () や []、およびコロン以降の数値「:1.3」を物理的にスルーして純粋なDanbooruタグだけを抽出する
    cleanedOutput = ""
    inWeight = False
    
    For i = 1 To Len(rawInput)
        char = Mid(rawInput, i, 1)
        
        If char = "(" Or char = ")" Or char = "[" Or char = "]" Then
            ' カッコ自体は無視してスキップ
        ElseIf char = ":" Then
            ' コロンを検知したら、次のカンマかカッコが来るまで「ウェイト数値ゾーン」とみなしてブロック
            inWeight = True
        ElseIf inWeight And (char = "," Or char = vbLf Or char = vbCr) Then
            ' ウェイトゾーン中に区切り（カンマや改行）が来たら、ゾーン解除して文字取得を再開
            inWeight = False
            cleanedOutput = cleanedOutput & char
        Else
            ' 通常文字であれば、ウェイトゾーン外のときのみ結合
            If Not inWeight Then
                cleanedOutput = cleanedOutput & char
            End If
        End If
    Next i
    
    ' 連続したスペースや不要なカンマの並びを綺麗に整形（QoL職人芸）
    cleanedOutput = Replace(cleanedOutput, ", ,", ",")
    Me.txtYourPositive.text = Trim(cleanedOutput)
    
    MsgBox "プロンプトのクレンジングが完了しました！カッコやウェイト数値が自動除去されました。" & vbCrLf & _
           "Prompt cleaning complete! Brackets and weight values have been auto-removed.", vbInformation, APP_NAME
End Sub

' -------------------------------------------------------------------------
' 1. 初期化とUI状態一括管理
' -------------------------------------------------------------------------
' =========================================================================
' ポジティブ側のUI状態一括更新（マルチセレクト判定・ソート完全対応版）
' =========================================================================
Private Sub UpdatePositiveUIState()
    Dim hasSelectAll As Boolean, hasSelectApply As Boolean
    Dim hasItemAll As Boolean, hasItemApply As Boolean
    Dim hasTextResult As Boolean
    Dim i As Integer
    
    Dim selCountAll As Integer: selCountAll = 0
    Dim selIdxAll As Integer: selIdxAll = -1
    Dim selCountApply As Integer: selCountApply = 0
    Dim selIdxApply As Integer: selIdxApply = -1
    
    hasItemAll = (Me.lstAllPositive.ListCount > 0)
    hasItemApply = (Me.lstApplyPositive.ListCount > 0)
    hasTextResult = (Trim(Me.txtYourPositive.text) <> "")
    
    ' ★修正：選択状態と正確なインデックスをループで確実にカウントする
    For i = 0 To Me.lstAllPositive.ListCount - 1
        If Me.lstAllPositive.Selected(i) Then
            hasSelectAll = True
            selCountAll = selCountAll + 1
            selIdxAll = i
        End If
    Next i
    
    For i = 0 To Me.lstApplyPositive.ListCount - 1
        If Me.lstApplyPositive.Selected(i) Then
            hasSelectApply = True
            selCountApply = selCountApply + 1
            selIdxApply = i
        End If
    Next i

    ' --- 各種操作ボタンのEnabledとカラーの制御 ---
    Call SetControlUIState(Me.btnDeletePositive, hasSelectAll, &HC0E0FF)
    Call SetControlUIState(Me.btnAllDeletePositive, hasItemAll, &HC0E0FF)
    Call SetControlUIState(Me.btnApplyPositive, hasSelectAll, &HC8F0C8)
    Call SetControlUIState(Me.btnAddPositive, hasSelectApply, &HC8F0C8)
    Call SetControlUIState(Me.btnDismissPositive, hasSelectApply, &HC0E0FF)
    Call SetControlUIState(Me.btnAllPositiveSelect, hasItemApply, &HFFE1C8)
    Call SetControlUIState(Me.btnAllPositiveDismiss, hasItemApply, &HC0E0FF)
    
    ' --- ★修正：ソートボタンの制御（1件のみ選択されている場合だけアクティブ化） ---
    Call SetControlUIState(Me.btnAllPositiveSortUp, (selCountAll = 1 And selIdxAll > 0), &HFFE1C8)
    Call SetControlUIState(Me.btnAllPositiveSortDown, (selCountAll = 1 And selIdxAll >= 0 And selIdxAll < Me.lstAllPositive.ListCount - 1), &HFFE1C8)
    
    Call SetControlUIState(Me.btnApplyPositiveSortUp, (selCountApply = 1 And selIdxApply > 0), &HC8F0C8)
    Call SetControlUIState(Me.btnApplyPositiveSortDown, (selCountApply = 1 And selIdxApply >= 0 And selIdxApply < Me.lstApplyPositive.ListCount - 1), &HC8F0C8)
    
    ' --- 出力テキストボックスと関連UIの制御 ---
    Me.txtYourPositive.Enabled = (hasItemApply Or hasTextResult)
    Me.txtYourPositive.BackColor = IIf(Me.txtYourPositive.Enabled, vbWindowBackground, &HE0E0E0)
    
    Call SetControlUIState(Me.btnSetPositiveDefault, hasTextResult, &HC8F0C8)
    Call SetControlUIState(Me.btnSaveAsPrtesetPositive, hasTextResult, &HFFE1C8)
    Call SetControlUIState(Me.btnSetPositiveToCockpit, hasTextResult, &HC8F0C8)
    Call SetControlUIState(Me.btnClearPositive, hasTextResult, &HC0E0FF)
End Sub
' -------------------------------------------------------------------------
' [適用(->)] ボタン：ストックから適用中リストへコピー（仮決め状態を維持）
' -------------------------------------------------------------------------
Private Sub btnApplyPositive_Click()
    Dim i As Long
    Dim word As String
    Dim isDuplicate As Boolean
    Dim k As Long
    Dim addedCount As Long
    
    addedCount = 0
    
    ' ストックリストの選択されたアイテムを上から順に走査
    For i = 0 To Me.lstAllPositive.ListCount - 1
        If Me.lstAllPositive.Selected(i) Then
            word = Me.lstAllPositive.List(i)
            
            ' 適用中リスト(lstApplyPositive)に既に存在しないか重複チェック
            isDuplicate = False
            For k = 0 To Me.lstApplyPositive.ListCount - 1
                If StrComp(Me.lstApplyPositive.List(k), word, vbTextCompare) = 0 Then
                    isDuplicate = True
                    Exit For
                End If
            Next k
            
            ' 重複していなければ適用中リストに追加（★ストック側からのRemoveItemは行わない）
            If Not isDuplicate Then
                Me.lstApplyPositive.AddItem word
                addedCount = addedCount + 1
            End If
            
            ' 選択状態を解除（ポジティブ側のUXに合わせる場合）
            Me.lstAllPositive.Selected(i) = False
        End If
    Next i
    
    ' 1つでも追加されたら全体の出力テキスト（txtYourNegative）を再ビルドしてUIを更新
    If addedCount > 0 Then
        Call UpdatePositiveUIState
    End If
End Sub

' -------------------------------------------------------------------------
' [適用解除(<-)] ボタン：適用中リストから選んだ項目を外す（ポジティブ側）
' -------------------------------------------------------------------------
' -------------------------------------------------------------------------
' [適用解除(<-)] ボタン：適用中リストから選んだ項目を外す（ポジティブ側）
' -------------------------------------------------------------------------
Private Sub btnDismissPositive_Click()
    Dim i As Integer, k As Integer
    Dim hasDeleted As Boolean
    Dim hasAddedToStock As Boolean
    Dim word As String
    Dim isDuplicate As Boolean
    
    For i = Me.lstApplyPositive.ListCount - 1 To 0 Step -1
        If Me.lstApplyPositive.Selected(i) Then
            ' ★ストックに回収・比較する前に、ウェイトやカッコを剥がして純粋なタグに戻す
            word = GetPureTagName(Me.lstApplyPositive.List(i))
            
            ' ストック側に既に存在するか重複チェック
            isDuplicate = False
            For k = 0 To Me.lstAllPositive.ListCount - 1
                If StrComp(Me.lstAllPositive.List(k), word, vbTextCompare) = 0 Then
                    isDuplicate = True
                    Exit For
                End If
            Next k
            
            ' ストックに存在しない「未知のタグ」なら自動的にストックへ回収
            If Not isDuplicate Then
                Me.lstAllPositive.AddItem word
                hasAddedToStock = True
            End If
            
            ' 適用中リストからは削除
            Me.lstApplyPositive.RemoveItem i
            hasDeleted = True
        End If
    Next i
    
    ' ストックへの自動追加が発生した場合のみ、JSON（Stock）へ同期保存
    If hasAddedToStock Then Call SyncPositiveSheet
    
    ' 削除が発生した場合、出力テキストの再構築とUIの更新を行う
    If hasDeleted Then
        Call RefreshPositiveOutputText ' ★追加：適用リストの変更を出力テキストに即時同期
        Call UpdatePositiveUIState
        
        ' ステータスバーのメッセージを状況に応じて切り替え
        If hasAddedToStock Then
            Application.StatusBar = "● Removed from active list (New tags auto-added to stock)."
        Else
            Application.StatusBar = "● Removed selected item(s) from the active list."
        End If
    End If
End Sub

'-------------------------------------------------------------------------
' [All Dismiss] ボタン：適用中リストをすべて空にする (ポジティブ側)
'-------------------------------------------------------------------------
Private Sub btnAllPositiveDismiss_Click()
    Dim ans As VbMsgBoxResult
    Dim i As Integer, k As Integer
    Dim hasAddedToStock As Boolean
    Dim word As String
    Dim isDuplicate As Boolean
    
    If Me.lstApplyPositive.ListCount = 0 Then Exit Sub
    
    ' 確認ダイアログ（誤爆防止のため「いいえ」をデフォルトに設定）
    ans = MsgBox("適用リストのすべての項目をクリア（解除）しますか？" & vbCrLf & _
                 "Are you sure you want to dismiss all active items?", _
                 vbYesNo + vbQuestion + vbDefaultButton2, APP_NAME)
                 
    If ans = vbNo Then Exit Sub
    
    ' ★全消しする前に、未知のタグがあればストックへ一括回収する
    For i = 0 To Me.lstApplyPositive.ListCount - 1
        ' ウェイトやカッコを剥がして純粋なタグに戻す
        word = GetPureTagName(Me.lstApplyPositive.List(i))
        
        isDuplicate = False
        For k = 0 To Me.lstAllPositive.ListCount - 1
            If StrComp(Me.lstAllPositive.List(k), word, vbTextCompare) = 0 Then
                isDuplicate = True
                Exit For
            End If
        Next k
        
        If Not isDuplicate Then
            Me.lstAllPositive.AddItem word
            hasAddedToStock = True
        End If
    Next i
    
    ' ストックへの自動追加が発生した場合のみ、JSON（Stock）へ同期保存
    If hasAddedToStock Then Call SyncPositiveSheet
    
    ' 適用リストを一括クリア
    Me.lstApplyPositive.Clear
    
    ' ★出力テキストの再構築（空になる）とUI状態の更新
    Call RefreshPositiveOutputText
    Call UpdatePositiveUIState
    
    If hasAddedToStock Then
        Application.StatusBar = "● All items dismissed (New tags auto-added to stock)."
    Else
        Application.StatusBar = "● All positive active items dismissed."
    End If
End Sub

' 共通ヘルパー：適用中リストの状態を出力用テキストボックスへ反映・自動連結
Private Sub RefreshPositiveOutputText()
    Dim i As Integer
    Dim combined As String
    
    combined = ""
    For i = 0 To Me.lstApplyPositive.ListCount - 1
        If combined <> "" Then
            combined = combined & ", " & Me.lstApplyPositive.List(i)
        Else
            combined = Me.lstApplyPositive.List(i)
        End If
    Next i
    
    ' コクピットへの安全な出力
    Me.txtYourPositive.text = combined
    Call UpdatePositiveUIState
End Sub

' -------------------------------------------------------------------------
' 4. ソート操作（順序変更）
' -------------------------------------------------------------------------

' リストボックスのアイテムを上下にシフトさせる共通ロジック
Private Sub MovePositiveListItem(lst As MSForms.ListBox, direction As Integer)
    Dim idx As Integer
    Dim temp As String
    
    idx = lst.ListIndex
    If idx < 0 Then Exit Sub
    
    If direction = -1 Then ' Up
        If idx = 0 Then Exit Sub
        temp = lst.List(idx)
        lst.List(idx) = lst.List(idx - 1)
        lst.List(idx - 1) = temp
        lst.ListIndex = idx - 1
    Else ' Down
        If idx = lst.ListCount - 1 Then Exit Sub
        temp = lst.List(idx)
        lst.List(idx) = lst.List(idx + 1)
        lst.List(idx + 1) = temp
        lst.ListIndex = idx + 1
    End If
End Sub

'-------------------------------------------------------------------------
' [デフォルト登録] ボタン：現在の出力テキストをデフォルト（Baseポジティブ呪文）として登録
'-------------------------------------------------------------------------
Private Sub btnSetPositiveDefault_Click()
    Dim currentDefault As String
    Dim newDefault As String
    
    ' 1. JSONデータが未ロードの場合は初期化（安全対策）
    If ConfigData Is Nothing Then Call InitConfigJSON
    
    ' 2. JSONから現在のデフォルト値を取得（安全のためのExists判定）
    If ConfigData.Exists("Positive_Default") Then
        currentDefault = Trim(ConfigData("Positive_Default"))
    Else
        currentDefault = ""
    End If
    
    ' 3. 入力値（テキストボックス）を取得
    newDefault = Trim(Me.txtYourPositive.text)
    
    ' 4. 先行バリデーション
    If newDefault = "" Then
        MsgBox "出力欄が空欄です。登録する内容がありません。" & vbCrLf & _
               "The positive prompt output field is empty.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 5. 変更チェック
    If StrComp(currentDefault, newDefault, vbBinaryCompare) = 0 Then
        MsgBox "内容に変更がありませんでした。" & vbCrLf & _
               "No changes were made.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' 6. JSONデータ（メモリ上）へ上書き更新 ＆ 物理ファイルへ即時書き出し
    ConfigData("Positive_Default") = newDefault
    Call SaveConfigJSON
    
    ' 7. 結果通知
    MsgBox "デフォルトのポジティブプロンプト（Base呪文）を更新しました！" & vbCrLf & _
           "Default positive prompt updated successfully!", vbInformation, APP_NAME
           
    Application.StatusBar = "● Default positive prompt updated (JSON updated)."
End Sub


'-------------------------------------------------------------------------
' [デフォルト登録] ボタン：現在の出力テキストをデフォルト（Baseネガティブ呪文）として登録
'-------------------------------------------------------------------------
Private Sub btnSetNegativeDefault_Click()
    Dim currentDefault As String
    Dim newDefault As String
    
    ' 1. JSONデータが未ロードの場合は初期化（安全対策）
    If ConfigData Is Nothing Then Call InitConfigJSON
    
    ' 2. JSONから現在のデフォルト値を取得（安全のためのExists判定）
    If ConfigData.Exists("Negative_Default") Then
        currentDefault = Trim(ConfigData("Negative_Default"))
    Else
        currentDefault = ""
    End If
    
    ' 3. 入力値（テキストボックス）を取得
    newDefault = Trim(Me.txtYourNegative.text)
    
    ' 4. 先行バリデーション
    If newDefault = "" Then
        MsgBox "出力欄が空欄です。登録する内容がありません。" & vbCrLf & _
               "The negative prompt output field is empty.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 5. 変更チェック
    If StrComp(currentDefault, newDefault, vbBinaryCompare) = 0 Then
        MsgBox "内容に変更がありませんでした。" & vbCrLf & _
               "No changes were made.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' 6. JSONデータ（メモリ上）へ上書き更新 ＆ 物理ファイルへ即時書き出し
    ConfigData("Negative_Default") = newDefault
    Call SaveConfigJSON
    
    ' 7. 結果通知
    MsgBox "デフォルトのネガティブプロンプト（Base呪文）を更新しました！" & vbCrLf & _
           "Default negative prompt updated successfully!", vbInformation, APP_NAME
           
    Application.StatusBar = "● Default negative prompt updated (JSON updated)."
End Sub

' -------------------------------------------------------------------------
' 6. リスト操作時のイベントハンドリングトリガー
' -------------------------------------------------------------------------
Private Sub lstAllPositive_Click()
    Call UpdatePositiveUIState
End Sub

Private Sub lstApplyPositive_Click()
    Call UpdatePositiveUIState
End Sub



'-------------------------------------------------------------------------
' [Clear] ボタン：出力欄（txtYourPositive）の内容を消去
'-------------------------------------------------------------------------
Private Sub btnClearPositive_Click()
    ' 出力欄を空にする
    Me.txtYourPositive.text = ""
    
    ' UI状態（ボタンのグレーアウトや背景色）を即座に再判定
    Call UpdatePositiveUIState
    
    Application.StatusBar = "● Positive output field cleared."
End Sub


' -------------------------------------------------------------------------
' 6. 各種イベントハンドリングトリガー（修正・追加版）
' -------------------------------------------------------------------------

' リストボックスの選択変更を確実に検知する（ClickからChangeへ修正）
Private Sub lstAllPositive_Change()
    Call UpdatePositiveUIState
End Sub

Private Sub lstApplyPositive_Change()
    Call UpdatePositiveUIState
End Sub

' 出力欄（txtYourPositive）の内容変更を検知する（クリアボタン等の連動用に追加）
Private Sub txtYourPositive_Change()
    Call UpdatePositiveUIState
End Sub

' フォーカスが当たった時：プレースホルダーを消す
Private Sub txtInputPositive_Enter()
    With Me.txtInputPositive
        If .text = PH_POSI_INPUT Then
            .text = ""
            .ForeColor = vbBlack
        End If
    End With
End Sub

' フォーカスが外れた時：空ならプレースホルダーを戻す
Private Sub txtInputPositive_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.txtInputPositive
        If Trim(.text) = "" Then
            .text = PH_POSI_INPUT
            .ForeColor = COLOR_GRAY
        End If
    End With
End Sub

' 一括入力欄（txtInputPositive）の文字監視（プレースホルダー考慮版に更新）
Private Sub txtInputPositive_Change()
    Dim currentText As String
    Dim hasValidInput As Boolean
    
    currentText = Trim(Me.txtInputPositive.text)
    
    ' ★修正：入力が空欄ではなく、かつプレースホルダーでもないかを判定
    hasValidInput = (currentText <> "") And (currentText <> PH_POSI_INPUT)
    
    ' 状態に合わせて追加ボタンの有効・無効をパステルカラーで制御
    Call SetControlUIState(Me.btnAddPositive, hasValidInput, &HFFE1C8)
End Sub


' -------------------------------------------------------------------------
' [Add New] ボタンが押された時
' （※ボタン名が btnAddPositive の場合は適宜名前を合わせてください）
' -------------------------------------------------------------------------
Private Sub btnAddNewPositive_Click()
    Call AddNewPositiveWord
End Sub

'-------------------------------------------------------------------------
' 入力欄 (txtInputPositive) でのキーボード操作（Enter登録 ＆ ショートカット）
'-------------------------------------------------------------------------
Private Sub txtInputPositive_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' 1. Enterキー (13) でのストック登録処理
    If KeyCode = 13 Then
        KeyCode = 0
        Call AddNewPositiveWord
        Exit Sub
    End If
    
    ' 2. Ctrl+Shift+P でのコクピット送信処理
    If Shift = 3 And KeyCode = vbKeyP Then
        Call SetPositiveAction
        KeyCode = 0 ' ビープ音防止
    End If
End Sub

'-------------------------------------------------------------------------
' [追加(btnAddPositive)] ボタン：適用中リストの選択項目を出力欄へ送る（重複スキップ＆全ダブり通知版）
'-------------------------------------------------------------------------
Private Sub btnAddPositive_Click()
    Dim i As Integer
    Dim combinedWords As String
    Dim currentText As String
    Dim addedCount As Integer
    Dim selectedCount As Integer
    Dim word As String
    
    currentText = Trim(Me.txtYourPositive.text)
    combinedWords = ""
    addedCount = 0
    selectedCount = 0
    
    ' 適用中リスト (lstApplyPositive) の選択された項目をループ処理
    For i = 0 To Me.lstApplyPositive.ListCount - 1
        If Me.lstApplyPositive.Selected(i) Then
            selectedCount = selectedCount + 1
            word = Trim(Me.lstApplyPositive.List(i))
            
            ' 1. 出力欄（txtYourPositive）に既に存在するかをチェック
            ' ★修正：Exit Subで処理を止めず、重複していない場合のみ追加処理へ進む
            If InStr(1, currentText, word, vbTextCompare) = 0 Then
                
                ' 2. 今回の同時追加分の中での重複もチェック
                If InStr(1, combinedWords, word, vbTextCompare) = 0 Then
                    If combinedWords <> "" Then
                        combinedWords = combinedWords & ", " & word
                    Else
                        combinedWords = word
                    End If
                    addedCount = addedCount + 1
                End If
            End If
            
            ' 処理した項目の選択を解除（追加されたものも、スキップされたものも解除）
            Me.lstApplyPositive.Selected(i) = False
        End If
    Next i
    
    ' 何も選択されていなかった場合は静かに終了
    If selectedCount = 0 Then Exit Sub
    
    ' =========================================================
    ' ?? 結果の判定（すべて重複だった場合のMsgBox通知）
    ' =========================================================
    If addedCount = 0 Then
        ' 選択した項目がすべて出力欄に存在していた場合
        MsgBox "選択されたプロンプトはすべて既に出力欄に含まれています。" & vbCrLf & _
               "（新しいプロンプトは追加されませんでした）" & vbCrLf & vbCrLf & _
               "All selected prompts are already included in the output field." & vbCrLf & _
               "(No new prompts were added.)", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' =========================================================
    ' 出力欄へ反映とUI更新
    ' =========================================================
    If currentText = "" Then
        Me.txtYourPositive.text = combinedWords
    Else
        Me.txtYourPositive.text = currentText & IIf(Right(currentText, 1) = ",", " ", ", ") & combinedWords
    End If
    
    ' UI状態の更新
    Call UpdatePositiveUIState
    
    Application.StatusBar = "● Added " & addedCount & " prompt(s) to the output field."
End Sub


'-------------------------------------------------------------------------
' [Send to Cockpit] ボタン：出力をコクピット(txtMain)へ送る（究極スマート結合版）
'-------------------------------------------------------------------------
Public Sub btnSetPositiveToCockpit_Click()
    Dim newPrompt As String
    Dim currentMain As String
    
    newPrompt = Trim(Me.txtYourPositive.text)
    
    ' 出力欄が空なら処理しない
    If newPrompt = "" Then Exit Sub
    
    currentMain = Trim(Me.txtMain.text)
    
    ' 1. 重複チェック
    If currentMain <> "" Then
        If InStr(1, currentMain, newPrompt, vbTextCompare) > 0 Then
            MsgBox "このポジティブプロンプトは既にコクピットに含まれています。" & vbCrLf & _
                   "This positive prompt is already included in the Cockpit.", vbInformation, APP_NAME
            Exit Sub
        End If
    End If
    
    ' 2. 変更前の状態をUndo用履歴シートへ保存
    Call SaveHistory(Me.txtMain.text)
    
    ' 3. 挿入位置の決定（常に先頭へ挿入＆スマート結合）
    If currentMain = "" Then
        Me.txtMain.text = newPrompt
    Else
        ' ★修正：コクピットの先頭が "<lora:" で始まっている場合（トリガーがなくLoRA単体の場合）
        ' カンマで繋ぐと「masterpiece, <lora:」となりLoRA消滅後にカンマが残るため、スペースで繋ぐ
        If Left(currentMain, 6) = "<lora:" Then
            Me.txtMain.text = newPrompt & " " & currentMain
        Else
            ' それ以外（一般プロンプトやトリガーがある場合）はカンマで繋ぐ
            Me.txtMain.text = newPrompt & ", " & currentMain
        End If
    End If
    
    ' 4. タブのフォーカスを強制的にコクピット(txtMain)へ移動
    On Error Resume Next
    Me.MainWindow.Value = Me.MainWindow.Pages("pgCockpit").Index
    On Error GoTo 0
    
    Me.txtMain.SetFocus
    
    ' 5. ステータスバーへの通知
    Application.StatusBar = "● Positive prompt sent to Cockpit."
End Sub

' =========================================================
' v.3.0.0 ポジティブストック：操作・ソート・JSON同期エンジン
' =========================================================

' [選択削除] ボタン：ストックから選んだ項目を消去
Private Sub btnDeletePositive_Click()
    Dim i As Integer
    Dim hasDeleted As Boolean
    
    ' 複数選択に対応するため、下から逆順に走査して削除
    For i = Me.lstAllPositive.ListCount - 1 To 0 Step -1
        If Me.lstAllPositive.Selected(i) Then
            Me.lstAllPositive.RemoveItem i
            hasDeleted = True
        End If
    Next i
    
    ' 削除が発生した場合のみ、JSONへ同期してUI状態を更新
    If hasDeleted Then
        Call SyncPositiveSheet ' 新しいJSON同期ヘルパーを呼び出し
        Call UpdatePositiveUIState
        Application.StatusBar = "● Selected stock item(s) deleted (JSON updated)."
    End If
End Sub

' [全削除] ボタン：ストックを完全クリア（日英併記維持）
Private Sub btnAllDeletePositive_Click()
    Dim ans As VbMsgBoxResult
    If Me.lstAllPositive.ListCount = 0 Then Exit Sub
    
    ans = MsgBox("ストックされている全てのポジティブプロンプトを削除しますか？" & vbCrLf & _
                 "Are you sure you want to delete all stocked positive prompts?", _
                 vbYesNo + vbQuestion + vbDefaultButton2, APP_NAME)
    If ans = vbNo Then Exit Sub
    
    ' リストボックスをクリアしてJSONに同期（空の状態で保存されます）
    Me.lstAllPositive.Clear
    Call SyncPositiveSheet
    Call UpdatePositiveUIState
    Application.StatusBar = "● All stock items cleared (JSON updated)."
End Sub

' ??【完全改装】共通ヘルパー：lstAllPositiveの状態をそのままJSON（PositiveStock）へ同期保存
' シート操作を100%完全に廃止し、現在の並び順をそのままファイルへ記憶させます。
Private Sub SyncPositiveSheet()
    Dim finalStock As String
    Dim k As Long
    
    finalStock = ""
    
    ' 1. 現在のリストボックスのアイテムを、現在の並び順のままカンマ区切りで結合
    If Me.lstAllPositive.ListCount > 0 Then
        For k = 0 To Me.lstAllPositive.ListCount - 1
            finalStock = finalStock & Me.lstAllPositive.List(k) & ", "
        Next k
        
        ' 末尾の余分なカンマとスペースを綺麗にカット
        If finalStock <> "" Then finalStock = Left(finalStock, Len(finalStock) - 2)
    End If
    
    ' 2. メモリ上のConfigDataへ格納し、物理ファイルに一撃で保存
    If ConfigData Is Nothing Then Call InitConfigJSON
    ConfigData("PositiveStock") = finalStock
    Call SaveConfigJSON
End Sub

' -------------------------------------------------------------------------
' 4. ソート操作（順序変更） - 共通ヘルパー(MoveListItem)完全対応版
' -------------------------------------------------------------------------

' ストック側のソート：移動後に SyncPositiveSheet が走るため、並び替えた順序が永続化されます。
Private Sub btnAllPositiveSortUp_Click()
    Call MoveListItem(Me.lstAllPositive, -1)
    Call SyncPositiveSheet
    Call UpdatePositiveUIState
End Sub

Private Sub btnAllPositiveSortDown_Click()
    Call MoveListItem(Me.lstAllPositive, 1)
    Call SyncPositiveSheet
    Call UpdatePositiveUIState
End Sub

' 適用中（コックピット側）のソート：画面上の一時的な編集のため、テキストの再マッハ整形のみ実行
Private Sub btnApplyPositiveSortUp_Click()
    Call MoveListItem(Me.lstApplyPositive, -1)
    Call RefreshPositiveOutputText
End Sub

Private Sub btnApplyPositiveSortDown_Click()
    Call MoveListItem(Me.lstApplyPositive, 1)
    Call RefreshPositiveOutputText
End Sub

' =========================================================================
' ネガティブ側のUI状態一括更新（マルチセレクト判定・ソート完全対応版）
' =========================================================================
Private Sub UpdateNegativeUIState()
    Dim hasSelectAll As Boolean, hasSelectApply As Boolean
    Dim hasItemAll As Boolean, hasItemApply As Boolean
    Dim hasTextResult As Boolean
    Dim i As Integer
    
    Dim selCountAll As Integer: selCountAll = 0
    Dim selIdxAll As Integer: selIdxAll = -1
    Dim selCountApply As Integer: selCountApply = 0
    Dim selIdxApply As Integer: selIdxApply = -1
    
    hasItemAll = (Me.lstAllNegative.ListCount > 0)
    hasItemApply = (Me.lstApplyNegative.ListCount > 0)
    hasTextResult = (Trim(Me.txtYourNegative.text) <> "")
    
    ' ★修正：選択状態と正確なインデックスをループで確実にカウントする
    For i = 0 To Me.lstAllNegative.ListCount - 1
        If Me.lstAllNegative.Selected(i) Then
            hasSelectAll = True
            selCountAll = selCountAll + 1
            selIdxAll = i
        End If
    Next i
    
    For i = 0 To Me.lstApplyNegative.ListCount - 1
        If Me.lstApplyNegative.Selected(i) Then
            hasSelectApply = True
            selCountApply = selCountApply + 1
            selIdxApply = i
        End If
    Next i

    ' --- 各種操作ボタンのEnabledとカラーの制御 ---
    Call SetControlUIState(Me.btnDeleteNegative, hasSelectAll, &HC0E0FF)
    Call SetControlUIState(Me.btnAllDeleteNegative, hasItemAll, &HC0E0FF)
    Call SetControlUIState(Me.btnApplyNegative, hasSelectAll, &HC8F0C8)
    Call SetControlUIState(Me.btnAddNegative, hasSelectApply, &HC8F0C8)
    Call SetControlUIState(Me.btnDismissNegative, hasSelectApply, &HC0E0FF)
    Call SetControlUIState(Me.btnAllNegativeSelect, hasItemApply, &HFFE1C8)
    Call SetControlUIState(Me.btnAllNegativeDismiss, hasItemApply, &HC0E0FF)


    ' --- ★修正：ソートボタンの制御（1件のみ選択されている場合だけアクティブ化） ---
    Call SetControlUIState(Me.btnAllNegativeSortUp, (selCountAll = 1 And selIdxAll > 0), &HFFE1C8)
    Call SetControlUIState(Me.btnAllNegativeSortDown, (selCountAll = 1 And selIdxAll >= 0 And selIdxAll < Me.lstAllNegative.ListCount - 1), &HFFE1C8)
    
    Call SetControlUIState(Me.btnApplyNegativeSortUp, (selCountApply = 1 And selIdxApply > 0), &HC8F0C8)
    Call SetControlUIState(Me.btnApplyNegativeSortDown, (selCountApply = 1 And selIdxApply >= 0 And selIdxApply < Me.lstApplyNegative.ListCount - 1), &HC8F0C8)

    ' --- 出力テキストボックスと関連UIの制御 ---
    Dim shouldEnableResult As Boolean
    shouldEnableResult = (hasSelectApply Or hasTextResult)
    
    Me.txtYourNegative.Enabled = shouldEnableResult
    Me.txtYourNegative.BackColor = IIf(shouldEnableResult, vbWindowBackground, &HE0E0E0)
    
    Me.chkWeightNegative.Enabled = hasSelectApply
    
    Dim isWeightOn As Boolean
    ' ※VBAのチェックボックスは稀にNullを返すことがあるため、安全のために True かどうかを明示的に判定します
    isWeightOn = (Me.chkWeightNegative.Value = True)
    
    Call SetControlUIState(Me.btnWeightenNegative, (hasSelectApply And isWeightOn), &HC8F0C8)
    Me.cmbNegativeWeight.Enabled = isWeightOn
    Me.cmbNegativeWeight.BackColor = IIf(isWeightOn, vbWhite, &HE0E0E0)
    
    Call SetControlUIState(Me.btnSetNegativeDefault, hasTextResult, &HC8F0C8)
    Call SetControlUIState(Me.btnSaveAsPresetNegative, hasTextResult, &HFFE1C8)
    Call SetControlUIState(Me.btnCopyNegative, hasTextResult, &HC8F0C8)
    Call SetControlUIState(Me.btnClearNegative, hasTextResult, &HC0E0FF)
End Sub

' --- LoRAタブ：リストやテキストの変化を検知してUIをリアルタイム更新する ---
'-------------------------------------------------------------------------
' ★カート（選択中リスト）の選択状態が変わった瞬間にUIをリアルタイム同期
'-------------------------------------------------------------------------
Private Sub lstSelectedLoRA_Change()
    Dim i As Integer
    Dim listAlias As String
    Dim loraList As Object
    Dim loraItem As Object
    Dim hasNega As Boolean
    
    hasNega = False
    
    ' 1. カート内を走査し、現在「選択（ハイライト）」されているLoRAをスキャン
    If Me.lstSelectedLoRA.ListCount > 0 Then
        For i = 0 To Me.lstSelectedLoRA.ListCount - 1
            If Me.lstSelectedLoRA.Selected(i) = True Then
                listAlias = Me.lstSelectedLoRA.List(i, 2) ' 3列目のAliasを取得
                
                ' メモリ上のJSONから固有ネガティブの有無をルックアップ
                If Not ConfigData Is Nothing Then
                    If ConfigData.Exists("LoRA_List") Then
                        Set loraList = ConfigData("LoRA_List")
                        For Each loraItem In loraList
                            If loraItem("Alias") = listAlias Then
                                If loraItem.Exists("Negative") Then
                                    If Trim(loraItem("Negative")) <> "" Then
                                        hasNega = True ' 1つでもネガティブ持ちが選択されていればフラグON
                                        Exit For
                                    End If
                                End If
                                Exit For
                            End If
                        Next loraItem
                    End If
                End If
            End If
            If hasNega Then Exit For ' 既に発見していればこれ以上のループは不要
        Next i
    End If
    
    ' 2. ★判定結果をリネーム後のボタン（btnGetLoRANegative）のEnabledへダイレクト反映
    Me.btnGetLoRANegative.Enabled = hasNega
End Sub


Private Sub lstSelectedLoRA_Click()
    Call UpdateLoRAUIState
End Sub

Private Sub txtLoRAPreview_Change()
    Call UpdateLoRAUIState
End Sub
'-------------------------------------------------------------------------
' UI制御：APIキーの有無で「Surprise Me!」のオンオフを切り替える
' ※txtAPIKeyの値が変更された時や、フォーム起動時(Initialize)に呼び出してください
'-------------------------------------------------------------------------
Private Sub UpdateSurpriseMeUI()
    If Trim(Me.txtAPIKey.text) = "" Then
        ' APIキーがない場合は強制的にオフ＆グレーアウト
        Me.chkSupriseMe.Value = False
        Me.chkSupriseMe.Enabled = False
    Else
        ' APIキーがあれば操作可能
        Me.chkSupriseMe.Enabled = True
    End If
    
    ' フレームの状態を同期
    Call chkSupriseMe_Click
End Sub

'-------------------------------------------------------------------------
' UI制御：chkSupriseMe のチェック状態でフレームの有効/無効を切り替え
'-------------------------------------------------------------------------
Private Sub chkSupriseMe_Click()
    Me.frmMenu.Visible = Me.chkSupriseMe.Value
    Me.frmMenu.Enabled = Me.chkSupriseMe.Value
    Call UpdateGachaUI
End Sub

Private Sub opbSFW_Click()
    Call UpdateGachaUI
End Sub

Private Sub opbNSFW_Click()
    Call UpdateGachaUI
End Sub

Private Sub opbHardcore_Click()
    Call UpdateGachaUI
End Sub

'-------------------------------------------------------------------------
' [Helper] DrawWord：指定された列（論理インデックス）から、確率に応じて単語を引く
' colLogicalIndex : 1～36 の論理的な列番号（1ならB列、2ならD列...）
' probability     : 0.0(0%) ～ 1.0(100%) の抽出確率
' outList         : 抽出に成功した場合、単語を格納するリスト
'-------------------------------------------------------------------------
Private Sub DrawWord(colLogicalIndex As Integer, probability As Double, ByRef outList As Object)
    ' 確率判定（外れたら何もせず終了）
    If Rnd > probability Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("KENZEN SeaArt Helper")
    
    ' 論理インデックス(1,2,3...)を、実際のExcelの列番号(2,4,6...)に変換
    Dim actualCol As Integer
    actualCol = colLogicalIndex * 2
    
    ' その列の最終行を取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, actualCol).End(xlUp).Row
    
    ' ★修正1：1行目～19行目はヘッダや凡例エリアのため、20行目以降にデータが無ければ終了
    If lastRow < 20 Then Exit Sub
    
    ' 空白やエラーセル、システム文字列を除外して、有効な単語だけを配列にストックする
    Dim validWords() As String
    Dim wCount As Integer
    wCount = 0
    ReDim validWords(1 To lastRow)
    
    Dim r As Long
    ' ★修正1：データの読み込みを「20行目」から開始する
    For r = 20 To lastRow
        ' 辞書破壊テロ対策（エラー値は無視）
        If Not IsError(ws.Cells(r, actualCol).Value) Then
            Dim cellVal As String
            cellVal = Trim(CStr(ws.Cells(r, actualCol).Value))
            
            ' 空白チェック
            If cellVal <> "" Then
                ' ★修正2：文字列に「凡例」「プロンプト」や「ワイルドカード」関連が含まれる場合は除外
                If InStr(1, cellVal, "凡例", vbTextCompare) = 0 And _
                   InStr(1, cellVal, "プロンプト", vbTextCompare) = 0 And _
                   InStr(1, cellVal, "ワイルドカード", vbTextCompare) = 0 And _
                   InStr(1, cellVal, "Wildcard", vbTextCompare) = 0 And _
                   InStr(1, cellVal, "__", vbTextCompare) = 0 Then
                   
                    wCount = wCount + 1
                    validWords(wCount) = cellVal
                    
                End If
            End If
        End If
    Next r
    
    ' 有効な単語が1つ以上あれば、ランダムに1つ選んでリストに追加
    If wCount > 0 Then
        Dim randIdx As Integer
        randIdx = Int((wCount * Rnd) + 1)
        outList.Add validWords(randIdx)
    End If
End Sub

' ??【新規追加】2つのプロンプトが、並び順を無視して「同じタグ群」で構成されているかを判定する
Private Function IsSamePromptIgnoreOrder(ByVal p1 As String, ByVal p2 As String) As Boolean
    Dim dict1 As Object, dict2 As Object
    Dim words1() As String, words2() As String
    Dim w As Variant, k As Variant
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    dict1.CompareMode = 1 ' 大文字小文字を区別しない (vbTextCompare)
    dict2.CompareMode = 1
    
    ' p1（比較元）を改行・カンマでバラしてDictionaryにカウント登録
    p1 = Replace(p1, vbCrLf, ",")
    p1 = Replace(p1, vbCr, ",")
    p1 = Replace(p1, vbLf, ",")
    words1 = Split(p1, ",")
    For Each w In words1
        w = Trim(CStr(w))
        If w <> "" Then
            dict1(w) = dict1(w) + 1 ' 出現回数をカウント
        End If
    Next
    
    ' p2（既存データ）を同様にバラしてカウント登録
    p2 = Replace(p2, vbCrLf, ",")
    p2 = Replace(p2, vbCr, ",")
    p2 = Replace(p2, vbLf, ",")
    words2 = Split(p2, ",")
    For Each w In words2
        w = Trim(CStr(w))
        If w <> "" Then
            dict2(w) = dict2(w) + 1
        End If
    Next
    
    ' 双方とも完全に空なら「一致」とみなす
    If dict1.Count = 0 And dict2.Count = 0 Then
        IsSamePromptIgnoreOrder = True
        Exit Function
    End If
    
    ' 登録されている単語の種類数が違えば、その時点で「不一致」
    If dict1.Count <> dict2.Count Then
        IsSamePromptIgnoreOrder = False
        Exit Function
    End If
    
    ' 各単語の存在と、出現回数が過不足なく一致するかクロスチェック
    For Each k In dict1.Keys
        If Not dict2.Exists(k) Then
            IsSamePromptIgnoreOrder = False ' 相手方に存在しない単語があればアウト
            Exit Function
        End If
        If dict1(k) <> dict2(k) Then
            IsSamePromptIgnoreOrder = False ' 単語の出現回数が違えばアウト（強調の差を検知）
            Exit Function
        End If
    Next k
    
    ' すべての関門を突破すれば「順不同の完全一致」と判定
    IsSamePromptIgnoreOrder = True

End Function
' ==========================================
' pgNegative（ネガティブプロンプト管理）の初期化 (JSON完全対応版)
' ==========================================
Private Sub InitNegativePromptPage()
    ' JSONの初期化エンジンを確実に走らせる
    Call InitNegativeJSON
    
    ' 水平スクロールバーを表示させるための設定
    Me.lstAllNegative.ColumnWidths = "300 pt"
    Me.lstApplyNegative.ColumnWidths = "300 pt"
    
    ' プリセットコンボボックスにリストを充填（先頭にDefaultが自動挿入されます）
    Call LoadNegativePresetsToCombo(Me.cmbSelectNegaPreset)
    If Me.cmbSelectNegaPreset.ListCount > 0 Then
        Me.cmbSelectNegaPreset.ListIndex = 0
    End If
    
    ' =========================================================
    ' JSON内の「NegativeStock」からストックを復元しリストボックスへ即時展開
    ' =========================================================
    Me.lstAllNegative.Clear
    Dim stockPrompt As String
    stockPrompt = ""
    
    On Error Resume Next
    stockPrompt = Trim(ConfigData("NegativeStock"))
    On Error GoTo 0
    
    If stockPrompt <> "" Then
        Dim words() As String
        Dim word As String
        Dim k As Long
        
        ' カンマ前後のスペースのブレを綺麗に整地
        stockPrompt = Replace(stockPrompt, " ,", ",")
        stockPrompt = Replace(stockPrompt, ", ", ",")
        words = Split(stockPrompt, ",")
        
        For k = 0 To UBound(words)
            word = Trim(words(k))
            If word <> "" Then Me.lstAllNegative.AddItem word
        Next k
    End If
    
    ' cmbNegativeWeight のセット（0.5 ～ 1.3 まで 0.1 刻み）
    Me.cmbNegativeWeight.Clear
    Dim i As Long
    For i = 5 To 15
        If i <> 10 Then ' 1.0は飛ばす
            Me.cmbNegativeWeight.AddItem Format(i / 10, "0.0")
        End If
    Next i
    Me.cmbNegativeWeight.Value = "1.1" ' 初期値
    
    ' 入力欄のプレースホルダー初期化
    With Me.txtInputNegative
        .text = PH_NEGA_INPUT
        .ForeColor = COLOR_GRAY
    End With
    
    ' 手動の個別ロックを廃止し、一括更新プロシージャにすべて委ねる
    Call UpdateNegativeUIState
End Sub




' =========================================================================
' ネガティブストック：操作・ソート・JSON同期エンジン (v3.0.0 完全対称版)
' =========================================================================

' --- 1. ソート：ストック側 (All) 上下移動 ＆ JSON同期 ---
Private Sub btnAllNegativeSortUp_Click()
    Call MoveListItem(Me.lstAllNegative, -1)
    Call SyncNegativeSheet ' JSONへ即時同期
    Call UpdateNegativeUIState
End Sub

Private Sub btnAllNegativeSortDown_Click()
    Call MoveListItem(Me.lstAllNegative, 1)
    Call SyncNegativeSheet ' JSONへ即時同期
    Call UpdateNegativeUIState
End Sub

' --- 2. 新規入力欄(txtInputNegative) の状態監視（プレースホルダー考慮版） ---
Private Sub txtInputNegative_Change()
    Dim currentText As String
    Dim hasValidInput As Boolean
    
    currentText = Trim(Me.txtInputNegative.text)
    
    ' 入力が空欄ではなく、かつプレースホルダーでもないかを判定
    hasValidInput = (currentText <> "") And (currentText <> PH_NEGA_INPUT)
    
    ' 状態に合わせて追加ボタンの有効・無効をパステルカラーで制御
    Call SetControlUIState(Me.btnAddNewNegative, hasValidInput, &HFFE1C8)
    Call SetControlUIState(Me.btnClearNegaInput, hasValidInput, &HC0E0FF)
End Sub

' --- 3. [Add New] ボタンおよび Enter キーからの受け皿 ---
Private Sub btnAddNewNegative_Click()
    Call AddNewNegativeWord
End Sub



' --- 4. 【完全改装】物理ファイル保存同期エンジン（ワークシート操作を100%完全廃止） ---
Private Sub SyncNegativeSheet()
    Dim finalStock As String
    Dim k As Long
    
    finalStock = ""
    
    ' 現在のリストボックスのアイテムを、現在の並び順のままカンマ区切りで結合
    If Me.lstAllNegative.ListCount > 0 Then
        For k = 0 To Me.lstAllNegative.ListCount - 1
            finalStock = finalStock & Me.lstAllNegative.List(k) & ", "
        Next k
        
        ' 末尾の余分なカンマとスペースをカット
        If finalStock <> "" Then finalStock = Left(finalStock, Len(finalStock) - 2)
    End If
    
    ' メモリ上のConfigDataへ格納し、物理ファイルに保存
    If ConfigData Is Nothing Then Call InitConfigJSON
    ConfigData("NegativeStock") = finalStock
    Call SaveConfigJSON
End Sub

' --- 5. [Delete] ボタン：ストック(All)から選択された項目を削除する（複数選択対応版） ---
Private Sub btnDeleteNegative_Click()
    Dim i As Integer
    Dim hasDeleted As Boolean
    
    ' 複数選択に対応するため、下から逆順に走査して削除
    For i = Me.lstAllNegative.ListCount - 1 To 0 Step -1
        If Me.lstAllNegative.Selected(i) Then
            Me.lstAllNegative.RemoveItem i
            hasDeleted = True
        End If
    Next i
    
    ' 削除が発生した場合のみ、JSONへ同期してUI状態を更新
    If hasDeleted Then
        Call SyncNegativeSheet
        Call UpdateNegativeUIState
        Application.StatusBar = "● Selected stock item(s) deleted (JSON updated)."
    End If
End Sub

' --- 6. [All Delete] ボタン：ストック(All)の全データを物理削除する ---
Private Sub btnAllDeleteNegative_Click()
    Dim ans As VbMsgBoxResult
    
    ' リストが空なら何もしない
    If Me.lstAllNegative.ListCount = 0 Then Exit Sub
    
    ' 最終確認
    ans = MsgBox("ストックされている全てのネガティブプロンプトを削除しますか？" & vbCrLf & _
                 "Are you sure you want to delete all stocked negative prompts?", _
                 vbYesNo + vbQuestion + vbDefaultButton2, APP_NAME)
                 
    If ans = vbNo Then Exit Sub
    
    ' リストボックスをクリアしてJSONに同期（空の状態で上書き保存されます）
    Me.lstAllNegative.Clear
    Call SyncNegativeSheet
    Call UpdateNegativeUIState
    
    Application.StatusBar = "● All stock data cleared (JSON updated)."
End Sub

' =========================================================================
' 適用中リスト(Apply)の内容を出力テキスト(txtYourNegative)へ再構築・同期する
' =========================================================================
Private Sub RefreshNegativeOutputText()
    Dim i As Long
    Dim newText As String
    
    newText = ""
    
    ' 適用リスト内のアイテムを順にカンマ区切りで結合
    If Me.lstApplyNegative.ListCount > 0 Then
        For i = 0 To Me.lstApplyNegative.ListCount - 1
            newText = newText & Me.lstApplyNegative.List(i) & ", "
        Next i
        
        ' 末尾の余分なカンマとスペースをカット
        newText = Left(newText, Len(newText) - 2)
    End If
    
    ' 出力テキストボックスへ反映
    Me.txtYourNegative.text = newText
End Sub

'-------------------------------------------------------------------------
' Undoボタンの有効/無効を動的に制御する（独立メモリ監視版）
'-------------------------------------------------------------------------
Public Sub UpdateUndoButtonState()
    Dim hasHistory As Boolean
    hasHistory = False
    
    ' JSONではなく、専用のUndoメモリをチェックする
    If Not UndoMemory Is Nothing Then
        If UndoMemory.Count > 0 Then
            hasHistory = True
        End If
    End If
    
    ' ボタンの有効/無効を切り替え
    Me.btnUndo.Enabled = hasHistory
    
End Sub

' =========================================================================
' [Export側] Select All ボタン：すべて選択 / すべて解除のスマートトグル
' =========================================================================
Private Sub btnExportAll_Click()
    Dim targetState As Boolean
    
    ' 全てTrueならFalseに、それ以外ならTrueにする
    targetState = Not (Me.chkExportPosiPresets.Value And _
                       Me.chkExportNegaPresets.Value And _
                       Me.chkExportLoRA.Value And _
                       Me.chkExportLoRAPresets.Value And _
                       Me.chkExportFav.Value And _
                       Me.chkExportMobiMemo.Value)
                       
    Me.chkExportPosiPresets.Value = targetState
    Me.chkExportNegaPresets.Value = targetState
    Me.chkExportLoRA.Value = targetState
    Me.chkExportLoRAPresets.Value = targetState
    Me.chkExportFav.Value = targetState
    Me.chkExportMobiMemo.Value = targetState
    
End Sub

' =========================================================================
' [Import側] Select All ボタン：すべて選択 / すべて解除のスマートトグル
' =========================================================================
Private Sub btnImportAll_Click()
    Dim targetState As Boolean
    
    targetState = Not (Me.chkImportPosiPresets.Value And _
                       Me.chkImportNegaPresets.Value And _
                       Me.chkImportLoRA.Value And _
                       Me.chkImportLoRAPresets.Value And _
                       Me.chkImportFav.Value And _
                       Me.chkImportMobiMemo.Value)
                       
    Me.chkImportPosiPresets.Value = targetState
    Me.chkImportNegaPresets.Value = targetState
    Me.chkImportLoRA.Value = targetState
    Me.chkImportLoRAPresets.Value = targetState
    Me.chkImportFav.Value = targetState
    Me.chkImportMobiMemo.Value = targetState
    
End Sub


'-------------------------------------------------------------------------
' [UI制御] Undoボタンの有効・無効（グレーアウト）状態を切り替える処理
'-------------------------------------------------------------------------
Public Sub UpdateFavUndoStatus(ByVal hasHistory As Boolean)
    With Me.btnUndoFav
        If hasHistory Then
            .Enabled = True
            .BackColor = &HC0E0FF    ' 本来の色（薄いオレンジ）
        Else
            .Enabled = False
            .BackColor = &H8000000F  ' グレーアウト用のシステムカラー（通常ボタン色）
        End If
    End With
End Sub

'-------------------------------------------------------------------------
' [Call Preset] ボタン：選択されたポジティブプリセットを展開（JSON＆カッコ内カンマ完全対応版）
'-------------------------------------------------------------------------
Private Sub btnCallPosiPreset_Click()
    Dim selectedTarget As String
    Dim targetPrompt As String
    Dim presets As Object
    
    selectedTarget = Trim(Me.cmbSelectPosiPreset.text)
    
    ' バリデーション：何も選択されていない場合は警告して離脱
    If selectedTarget = "" Then
        MsgBox "呼び出すプリセットを選択してください。" & vbCrLf & _
               "Please select a preset to load.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 1. デフォルト判定（大文字小文字を区別せず、"Default" だった場合）
    If UCase(selectedTarget) = "DEFAULT" Then
        targetPrompt = Trim(ConfigData("Positive_Default"))
        
    ' 2. 通常プリセット検索（JSONの Dictionary からダイレクトに取得）
    Else
        Set presets = ConfigData("Positive_Presets")
        If presets.Exists(selectedTarget) Then
            targetPrompt = Trim(presets(selectedTarget))
        Else
            ' 万が一、JSON側に存在しない幽霊プリセット名だった場合のセーフティ
            MsgBox "指定されたプリセット「" & selectedTarget & "」は設定ファイル内に見つかりません。" & vbCrLf & _
                   "Preset '" & selectedTarget & "' not found in configuration.", vbCritical, APP_NAME
            Exit Sub
        End If
    End If
    
    ' 3. 出力欄に反映するとともに、適用中リスト（ListBox）側にも分解して即時展開する
    Me.txtYourPositive.text = targetPrompt
    Me.lstApplyPositive.Clear
    
    If targetPrompt <> "" Then
        Dim i As Long
        Dim charStr As String
        Dim currentWord As String
        Dim nestLevel As Integer
        
        currentWord = ""
        nestLevel = 0
        
        ' ========================================================
        ' ★ シンメトリカル修正：カッコ内のカンマを保護するスマート・パーサー
        ' ========================================================
        For i = 1 To Len(targetPrompt)
            charStr = Mid(targetPrompt, i, 1)
            
            ' カッコの開始なら階層を深くする
            If charStr = "(" Or charStr = "[" Or charStr = "{" Then
                nestLevel = nestLevel + 1
                currentWord = currentWord & charStr
                
            ' カッコの終了なら階層を浅くする
            ElseIf charStr = ")" Or charStr = "]" Or charStr = "}" Then
                nestLevel = nestLevel - 1
                If nestLevel < 0 Then nestLevel = 0 ' エラー回避のフェイルセーフ
                currentWord = currentWord & charStr
                
            ' カッコの外（nestLevel = 0）にあるカンマなら、そこで単語を分割！
            ElseIf charStr = "," And nestLevel = 0 Then
                If Trim(currentWord) <> "" Then Me.lstApplyPositive.AddItem Trim(currentWord)
                currentWord = ""
                
            ' それ以外の文字はそのまま結合
            Else
                currentWord = currentWord & charStr
            End If
        Next i
        
        ' ループを抜けた後に残っている最後の単語を追加
        If Trim(currentWord) <> "" Then Me.lstApplyPositive.AddItem Trim(currentWord)
    End If
    
    ' 4. 画面全体の表示状態をアップデート
    Call UpdatePositiveUIState
    
    ' 5. 鉄壁のデバッグ：実行時エラー2110を完全に封じるフォーカスガード
    On Error Resume Next
    If Me.txtYourPositive.Visible And Me.txtYourPositive.Enabled Then
        Me.txtYourPositive.SetFocus
    End If
    On Error GoTo 0
    
    Application.StatusBar = "● Loaded positive preset: " & selectedTarget
End Sub


'-------------------------------------------------------------------------
' [Call Preset] ボタン：選択されたネガティブプリセットを展開（JSON＆カッコ内カンマ完全対応版）
'-------------------------------------------------------------------------
Private Sub btnCallNegaPreset_Click()
    Dim selectedTarget As String
    Dim targetPrompt As String
    Dim presets As Object
    
    selectedTarget = Trim(Me.cmbSelectNegaPreset.text)
    
    ' バリデーション：何も選択されていない場合は警告して離脱
    If selectedTarget = "" Then
        MsgBox "呼び出すプリセットを選択してください。" & vbCrLf & _
               "Please select a preset to load.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 1. デフォルト判定（大文字小文字を区別せず、"Default" だった場合）
    If UCase(selectedTarget) = "DEFAULT" Then
        targetPrompt = Trim(ConfigData("Negative_Default"))
        
    ' 2. 通常プリセット検索（JSONの Dictionary からダイレクトに取得）
    Else
        Set presets = ConfigData("Negative_Presets")
        If presets.Exists(selectedTarget) Then
            targetPrompt = Trim(presets(selectedTarget))
        Else
            ' 万が一、JSON側に存在しない幽霊プリセット名だった場合のセーフティ
            MsgBox "指定されたプリセット「" & selectedTarget & "」は設定ファイル内に見つかりません。" & vbCrLf & _
                   "Preset '" & selectedTarget & "' not found in configuration.", vbCritical, APP_NAME
            Exit Sub
        End If
    End If
    
    ' 3. 出力欄に反映するとともに、適用中リスト（ListBox）側にも分解して即時展開する
    Me.txtYourNegative.text = targetPrompt
    Me.lstApplyNegative.Clear
    
    If targetPrompt <> "" Then
        Dim i As Long
        Dim charStr As String
        Dim currentWord As String
        Dim nestLevel As Integer
        
        currentWord = ""
        nestLevel = 0
        
        ' ========================================================
        ' ★ 修正：カッコ内のカンマを保護するスマート・パーサー
        ' ========================================================
        For i = 1 To Len(targetPrompt)
            charStr = Mid(targetPrompt, i, 1)
            
            ' カッコの開始なら階層を深くする
            If charStr = "(" Or charStr = "[" Or charStr = "{" Then
                nestLevel = nestLevel + 1
                currentWord = currentWord & charStr
                
            ' カッコの終了なら階層を浅くする
            ElseIf charStr = ")" Or charStr = "]" Or charStr = "}" Then
                nestLevel = nestLevel - 1
                If nestLevel < 0 Then nestLevel = 0 ' エラー回避のフェイルセーフ
                currentWord = currentWord & charStr
                
            ' カッコの外（nestLevel = 0）にあるカンマなら、そこで単語を分割！
            ElseIf charStr = "," And nestLevel = 0 Then
                If Trim(currentWord) <> "" Then Me.lstApplyNegative.AddItem Trim(currentWord)
                currentWord = ""
                
            ' それ以外の文字はそのまま結合
            Else
                currentWord = currentWord & charStr
            End If
        Next i
        
        ' ループを抜けた後に残っている最後の単語を追加
        If Trim(currentWord) <> "" Then Me.lstApplyNegative.AddItem Trim(currentWord)
    End If
    
    ' 4. 画面全体の表示状態をアップデート
    Call UpdateNegativeUIState
    
    ' 5. 鉄壁のデバッグ：実行時エラー2110を完全に封じるフォーカスガード
    On Error Resume Next
    If Me.txtYourNegative.Visible And Me.txtYourNegative.Enabled Then
        Me.txtYourNegative.SetFocus
    End If
    On Error GoTo 0
    
    Application.StatusBar = "● Loaded negative preset: " & selectedTarget
End Sub


' =========================================================================
' プリセット管理：削除（Delete）処理
' =========================================================================

'-------------------------------------------------------------------------
' [Delete Preset] ボタン：選択中のポジティブプリセットを削除（JSON完全対応版）
'-------------------------------------------------------------------------
Private Sub btnDeletePosiPreset_Click()
    Dim presetName As String
    Dim presets As Object
    Dim ans As Integer
    
    presetName = Trim(Me.cmbSelectPosiPreset.text)
    If presetName = "" Then Exit Sub
    
    ' 1. 鉄壁の防衛ガード：「Default」の削除要請は絶対に弾く
    If UCase(presetName) = "DEFAULT" Then
        MsgBox "デフォルトのプリセットDefaultを削除することはできません。" & vbCrLf & _
               "You cannot delete the default Preset spell.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    Set presets = ConfigData("Positive_Presets")
    If Not presets.Exists(presetName) Then Exit Sub
    
    ' 2. 削除確認ダイアログ
    ans = MsgBox("プリセット「" & presetName & "」を削除します。よろしいですか？" & vbCrLf & _
                 "Are you sure you want to delete preset '" & presetName & "'?", _
                 vbYesNo + vbExclamation + vbDefaultButton2, APP_NAME)
    If ans = vbNo Then Exit Sub
    
    ' 3. メモリから削除して物理ファイルへ即時保存
    presets.Remove presetName
    Call SaveConfigJSON
    
    ' 4. UI（コンボボックス）の再構築
    Call LoadPositivePresetsToCombo(Me.cmbSelectPosiPreset)
    
    ' 5. UX向上：削除後はコンボボックスの先頭（Default）を自動選択する
    If Me.cmbSelectPosiPreset.ListCount > 0 Then
        Me.cmbSelectPosiPreset.ListIndex = 0
    End If
    
    MsgBox "削除しました。" & vbCrLf & "Preset deleted.", vbInformation, APP_NAME
    
    ' 6. UX向上：選択されたDefaultの中身を即座にUI（出力欄・リスト）へ展開・同期する
    Call btnCallPosiPreset_Click
End Sub

'-------------------------------------------------------------------------
' [Delete Preset] ボタン：選択中のネガティブプリセットを削除（JSON完全対応版）
'-------------------------------------------------------------------------
Private Sub btnDeleteNegaPreset_Click()
    Dim presetName As String
    Dim presets As Object
    Dim ans As Integer
    
    presetName = Trim(Me.cmbSelectNegaPreset.text)
    If presetName = "" Then Exit Sub
    
    ' 1. 鉄壁の防衛ガード：「Default」の削除要請は絶対に弾く
    If UCase(presetName) = "DEFAULT" Then
        MsgBox "デフォルトのプリセットDefaultを削除することはできません。" & vbCrLf & _
               "You cannot delete the default Preset spell.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    Set presets = ConfigData("Negative_Presets")
    If Not presets.Exists(presetName) Then Exit Sub
    
    ' 2. 削除確認ダイアログ
    ans = MsgBox("プリセット「" & presetName & "」を削除します。よろしいですか？" & vbCrLf & _
                 "Are you sure you want to delete preset '" & presetName & "'?", _
                 vbYesNo + vbExclamation + vbDefaultButton2, APP_NAME)
    If ans = vbNo Then Exit Sub
    
    ' 3. メモリから削除して物理ファイルへ即時保存
    presets.Remove presetName
    Call SaveConfigJSON
    
    ' 4. UI（コンボボックス）の再構築
    Call LoadNegativePresetsToCombo(Me.cmbSelectNegaPreset)
    
    ' 5. UX向上：削除後はコンボボックスの先頭（Default）を自動選択する
    If Me.cmbSelectNegaPreset.ListCount > 0 Then
        Me.cmbSelectNegaPreset.ListIndex = 0
    End If
    
    MsgBox "削除しました。" & vbCrLf & "Preset deleted.", vbInformation, APP_NAME
    
    ' 6. UX向上：選択されたDefaultの中身を即座にUI（出力欄・リスト）へ展開・同期する
    Call btnCallNegaPreset_Click
End Sub


' =========================================================================
' 適用中リスト（Apply List）の全選択処理
' =========================================================================

'-------------------------------------------------------------------------
' [Select All] ボタン：ポジティブ側の適用中リストを全選択する
'-------------------------------------------------------------------------
Private Sub btnAllPositiveSelect_Click()
    Dim i As Integer
    
    ' 1. バリデーション：リストが空なら何もしない
    If Me.lstApplyPositive.ListCount = 0 Then Exit Sub
    
    ' 2. リスト内の全アイテムを選択状態 (Selected = True) にする
    For i = 0 To Me.lstApplyPositive.ListCount - 1
        Me.lstApplyPositive.Selected(i) = True
    Next i
    
    ' 3. 画面全体の表示状態をアップデート（削除ボタン等を有効化させる）
    Call UpdatePositiveUIState
    
    ' 4. ステータスバー通知
    Application.StatusBar = "● All items selected in Positive list."
End Sub

'-------------------------------------------------------------------------
' [Select All] ボタン：ネガティブ側の適用中リストを全選択する
'-------------------------------------------------------------------------
Private Sub btnAllNegativeSelect_Click()
    Dim i As Integer
    
    ' 1. バリデーション：リストが空なら何もしない
    If Me.lstApplyNegative.ListCount = 0 Then Exit Sub
    
    ' 2. リスト内の全アイテムを選択状態 (Selected = True) にする
    For i = 0 To Me.lstApplyNegative.ListCount - 1
        Me.lstApplyNegative.Selected(i) = True
    Next i
    
    ' 3. 画面全体の表示状態をアップデート（削除ボタン等を有効化させる）
    Call UpdateNegativeUIState
    
    ' 4. ステータスバー通知
    Application.StatusBar = "● All items selected in Negative list."
End Sub

' =========================================================================
' タイピングセッション管理（手動入力のUndoをスマートに記録する）
' =========================================================================

' テキストボックスにフォーカスが入った時（入力準備）
Private Sub txtMain_Enter()
    isTypingSession = False
    PreEditText = Me.txtMain.text
End Sub

' テキストボックスからフォーカスが外れた時（入力終了）
Private Sub txtMain_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isTypingSession = False
End Sub

' =========================================================================
' [Mobile Action] リストのダブルクリック時のイベント（共通処理を呼び出す）
' =========================================================================
Private Sub lstMobileMemo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ExecuteMobiAction
End Sub

' =========================================================================
' [Mobile Action] 実行ボタンクリック時のイベント（共通処理を呼び出す）
' =========================================================================
Private Sub btnDoMobiAction_Click()
    Call ExecuteMobiAction
End Sub

' =========================================================================
' [Mobile Action] 選択されたアイテムの種別（タグ）に応じてアクションを実行する共通処理
' =========================================================================
Private Sub ExecuteMobiAction()
    Dim listIdx As Long
    Dim itemTag As String
    Dim itemContent As String
    Dim cb As Object
    
    ' 1. アイテムが選択されているかチェック
    listIdx = Me.lstMobileMemo.ListIndex
    If listIdx = -1 Then
        MsgBox "アクションを実行するアイテムを選択してください。" & vbCrLf & _
               "Please select an item to execute.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' 2. 選択行のタグ（0列目）と内容（1列目）を取得
    itemTag = Me.lstMobileMemo.List(listIdx, 0)
    itemContent = Me.lstMobileMemo.List(listIdx, 1)
    
    ' 3. タグに応じた処理のルーティング
    Select Case itemTag
        Case "[URL]"
            ' 既定のブラウザでURLを開く
            On Error Resume Next
            CreateObject("WScript.Shell").Run itemContent
            If Err.Number <> 0 Then
                MsgBox "URLを開けませんでした。URLの形式を確認してください。" & vbCrLf & _
                       "Could not open URL.", vbCritical, APP_NAME
            End If
            On Error GoTo 0
            
        Case "[Hash]"
            ' ハッシュ値をクリップボードにコピー
            On Error Resume Next
            ' DataObjectのLate Binding（環境依存エラーを防ぐ確実なアプローチ）
            Set cb = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
            cb.SetText itemContent
            cb.PutInClipboard
            
            If Err.Number = 0 Then
                MsgBox "ハッシュ値をクリップボードにコピーしました！" & vbCrLf & _
                       "Hash value copied to clipboard!", vbInformation, APP_NAME
            Else
                MsgBox "クリップボードへのコピーに失敗しました。" & vbCrLf & _
                       "Failed to copy to clipboard.", vbCritical, APP_NAME
            End If
            On Error GoTo 0
            Set cb = Nothing
            
        Case Else
            ' [Memo] など、アクション対象外の場合
            MsgBox "この項目には、開けるURLやコピー可能なハッシュ値が含まれていません。" & vbCrLf & _
                   "This item does not contain a valid URL or Hash to execute.", vbInformation, APP_NAME
    End Select
End Sub

'=========================================================================
' ★追加：FavoritesManager の非同期ズレを防止する安全同期関数
'=========================================================================
Public Sub SyncFavoritesManager()
    Dim i As Integer
    Dim isManagerLoaded As Boolean
    isManagerLoaded = False
    
    ' FavoritesManager がメモリ上にロードされているか（起動中か）をチェック
    For i = 0 To VBA.UserForms.Count - 1
        If VBA.UserForms(i).Name = "FavoritesManager" Then
            isManagerLoaded = True
            Exit For
        End If
    Next i
    
    ' 起動している場合のみ、リストを強制リフレッシュして整合性を保つ
    If isManagerLoaded Then
        On Error Resume Next
        FavoritesManager.RefreshFavList
        On Error GoTo 0
    End If
End Sub

