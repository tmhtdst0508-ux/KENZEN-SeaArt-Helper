VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoRARegister 
   Caption         =   "LoRA Register & Manage"
   ClientHeight    =   7170
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10160
   OleObjectBlob   =   "LoRARegister_v4.0.0.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "LoRARegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==========================================================
' 【宣言部】 SHA256ハッシュ計算用 Windows API（一番上に記述）
' ==========================================================
#If VBA7 Then
    Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, ByVal pbHashObject As LongPtr, ByVal cbHashObject As Long, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As LongPtr) As Long
    Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByRef pbInput As Byte, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByRef pbOutput As Byte, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByRef phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As Long) As Long
    Private Declare Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As Long, ByRef pbInput As Byte, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As Long, ByRef pbOutput As Byte, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
#End If

Private Const BCRYPT_SHA256_ALGORITHM As String = "SHA256"
Private Const STATUS_SUCCESS As Long = &H0

Private m_EditIndex As Integer
Private m_HashPlaceholder As String
Private m_IsUpdating As Boolean
Private IsHandlingRegister As Boolean ' 登録ボタン処理中のガードフラグ

Private Sub btnRegisterLoRA_Click()

    Dim loraAlias As String, loraHash As String, loraStrength As String
    Dim loraTrigger As String, loraModelName As String, loraNegative As String
    Dim r As Long, i As Long, isEdit As Boolean
    
    IsHandlingRegister = True ' ガードオン

    ' 1. UIから値を取得
    loraAlias = Trim(Me.txtLoRAAlias.Value)
    loraStrength = Trim(Me.cmbReccomendStrengthRG.Value)
    loraModelName = Trim(Me.txtLoRAModelName.Value)
    
    If Me.txtLoRAHash.Value = m_HashPlaceholder Then loraHash = "" Else loraHash = Trim(Me.txtLoRAHash.Value)
    
    ' 2. トリガーワードの正規化
    If Me.chkTrigger_None_Rg.Value = True Then
        loraTrigger = ""
    Else
        loraTrigger = Replace(Replace(Replace(Trim(Me.txtTriggerWordRG.Value), " ,", ","), ", ", ","), ",", ", ")
    End If
    If Replace(Replace(loraTrigger, ",", ""), " ", "") = "" Then loraTrigger = ""
    
    ' 3. ネガティブプロンプトの正規化
    If Me.chkLoRANegativeNone.Value = True Then
        loraNegative = ""
    Else
        loraNegative = Replace(Replace(Replace(Trim(Me.txtLoRANegativePrompts.Value), " ,", ","), ", ", ","), ",", ", ")
    End If
    If Replace(Replace(loraNegative, ",", ""), " ", "") = "" Then loraNegative = ""
    
    ' 4. 更新モードの判定とJSONからの旧データ取得
    Dim oldAlias As String, oldHash As String, oldTrigger As String
    Dim oldStrength As String, oldModelName As String, oldNegative As String
    
    If m_EditIndex >= 0 Then
        ' 現在のリスト値をベースに取得
        oldAlias = Me.lstYourAllLoRA.List(m_EditIndex, 0)
        oldHash = Me.lstYourAllLoRA.List(m_EditIndex, 1)
        oldTrigger = Me.lstYourAllLoRA.List(m_EditIndex, 2)
        oldStrength = Me.lstYourAllLoRA.List(m_EditIndex, 3)
        oldModelName = Me.lstYourAllLoRA.List(m_EditIndex, 4)
        
        ' ネガティブだけはJSONから取得（安全策）
        oldNegative = ""
        If Not ConfigData Is Nothing Then
            If ConfigData.Exists("LoRA_List") Then
                Dim itemObj As Object
                For Each itemObj In ConfigData("LoRA_List")
                    If itemObj("Alias") = oldAlias And itemObj("Hash") = oldHash Then
                        If itemObj.Exists("Negative") Then oldNegative = Trim(itemObj("Negative"))
                        Exit For
                    End If
                Next itemObj
            End If
        End If
        
        ' 比較（すべて一致したら変更なし）
        If oldAlias = loraAlias And _
           oldHash = loraHash And _
           Replace(oldTrigger, " ", "") = Replace(loraTrigger, " ", "") And _
           oldStrength = loraStrength And _
           oldModelName = loraModelName And _
           Replace(oldNegative, " ", "") = Replace(loraNegative, " ", "") Then
            
            ' ★ メッセージを日英併記に変更
            MsgBox "内容に変更がありませんでした。" & vbCrLf & _
                   "No changes were detected.", vbInformation, APP_NAME
            Call ClearAllInputs
            IsHandlingRegister = False ' ★ デッドロック防止：ここでもガードを確実にオフ
            Exit Sub
        End If
        
        ' 更新処理
        Me.lstYourAllLoRA.List(m_EditIndex, 0) = loraAlias
        Me.lstYourAllLoRA.List(m_EditIndex, 1) = loraHash
        Me.lstYourAllLoRA.List(m_EditIndex, 2) = loraTrigger
        Me.lstYourAllLoRA.List(m_EditIndex, 3) = loraStrength
        Me.lstYourAllLoRA.List(m_EditIndex, 4) = loraModelName
        Me.lstYourAllLoRA.List(m_EditIndex, 5) = loraNegative
        isEdit = True
        
    Else
        ' 新規登録チェック
        For i = 0 To Me.lstYourAllLoRA.ListCount - 1
            If StrComp(Me.lstYourAllLoRA.List(i, 0), loraAlias, vbTextCompare) = 0 Then
                
                ' ★ メッセージを日英併記に変更
                MsgBox "このエイリアスは既に登録されています。" & vbCrLf & _
                       "This Alias is already registered.", vbExclamation, APP_NAME
                Me.txtLoRAAlias.SetFocus
                IsHandlingRegister = False ' ★ デッドロック防止：ここでもガードを確実にオフ
                Exit Sub
            End If
        Next i
        
        Me.lstYourAllLoRA.AddItem loraAlias
        r = Me.lstYourAllLoRA.ListCount - 1
        Me.lstYourAllLoRA.List(r, 1) = loraHash
        Me.lstYourAllLoRA.List(r, 2) = loraTrigger
        Me.lstYourAllLoRA.List(r, 3) = loraStrength
        Me.lstYourAllLoRA.List(r, 4) = loraModelName
        Me.lstYourAllLoRA.List(r, 5) = loraNegative
        isEdit = False
    End If
    
    Call SyncLoRAJSON
    
    ' ★ メッセージを日英併記に変更（可読性のためIf文を展開）
    If isEdit Then
        MsgBox "LoRAを更新しました！" & vbCrLf & _
               "LoRA updated successfully!", vbInformation, APP_NAME
    Else
        MsgBox "新しいLoRAを登録しました！" & vbCrLf & _
               "New LoRA registered successfully!", vbInformation, APP_NAME
    End If
    
    Call ClearAllInputs
    IsHandlingRegister = False ' ★ 正常終了時のガードオフ

End Sub

'-------------------------------------------------------------------------
' [無し] チェックボックス（固有ネガティブ用）の連動UI制御
'-------------------------------------------------------------------------
Private Sub chkLoRANegativeNone_Click()
    If Me.chkLoRANegativeNone.Value = True Then
        ' チェック時は入力欄をクリアして、標準システムカラーでグレーアウト
        Me.txtLoRANegativePrompts.Value = ""
        Me.txtLoRANegativePrompts.Enabled = False
        Me.txtLoRANegativePrompts.BackColor = &H8000000F  ' ★システムグレーに変更
    Else
        ' 解除時は入力欄を有効化
        Me.txtLoRANegativePrompts.Enabled = True
        Me.txtLoRANegativePrompts.BackColor = vbWhite
    End If
    If IsHandlingRegister Then Exit Sub ' ★ガード中は判定しない
    ' ボタンの状態を再評価
    Call CheckRegisterReady
End Sub
'-------------------------------------------------------------------------
' ★新設：ネガティブプロンプト入力欄の自動整地（一括コピペ対応版）
'-------------------------------------------------------------------------
Private Sub txtLoRANegativePrompts_AfterUpdate()
    Dim rawText As String
    Dim cleanedText As String
    
    If IsHandlingRegister Then Exit Sub ' ★ガード中は判定しない
    
    rawText = Trim(Me.txtLoRANegativePrompts.text)
    If rawText = "" Then Exit Sub
    
    ' 1. 全角英数字・スペース・句読点を半角・カンマに正規化
    cleanedText = NormalizeToHalfWidth(rawText)
    
    ' 2. 改行や連続するカンマ、不要なスペースのブレを鉄壁トラップ
    cleanedText = Replace(cleanedText, vbCrLf, ",")
    cleanedText = Replace(cleanedText, vbCr, ",")
    cleanedText = Replace(cleanedText, vbLf, ",")
    cleanedText = Replace(cleanedText, " ,", ",")
    cleanedText = Replace(cleanedText, ", ", ",")
    
    ' 連続カンマの集約（,, を , に）
    Do While InStr(cleanedText, ",,") > 0
        cleanedText = Replace(cleanedText, ",,", ",")
    Loop
    
    ' 先頭と末尾のカンマをトリミング
    If Left(cleanedText, 1) = "," Then cleanedText = Mid(cleanedText, 2)
    If Right(cleanedText, 1) = "," Then cleanedText = Left(cleanedText, Len(cleanedText) - 1)
    
    ' 3. 整地された美呪文をテキストボックスへ書き戻し
    Me.txtLoRANegativePrompts.text = Trim(cleanedText)
    
' 4. 編集ボタンの有効化状態を再評価
    Call CheckRegisterReady
    
End Sub

'-------------------------------------------------------------------------
' ★新設：ネガティブプロンプト入力欄の変更イベント（リアルタイムUI連動）
'-------------------------------------------------------------------------
Private Sub txtLoRANegativePrompts_Change()
    If IsHandlingRegister Then Exit Sub ' ★ガード中は判定しない
    ' 文字が入力されるたびに登録ボタンとキャンセルボタンの状態を再評価する
    Call CheckRegisterReady
End Sub
'-------------------------------------------------------------------------
' フォーム起動時の初期化処理
'-------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim i As Long
    m_EditIndex = -1
    m_HashPlaceholder = "Dbl-click" & ChrW(160) & "to" & ChrW(160) & "get" & ChrW(160) & "Auto" & ChrW(160) & "V2" & ChrW(160) & "Hash" & ChrW(160) & "from" & ChrW(160) & ".safetensors" & ChrW(160) & "file"
    
    m_IsUpdating = True
    
    Me.cmbReccomendStrengthRG.Clear
    For i = -50 To 140
        Me.cmbReccomendStrengthRG.AddItem Format(i / 20, "0.00")
    Next i
    
    ' ★リストボックスを6列に変更（Alias, Hash, Trigger, Strength, ModelName,LoRANegative）
    With Me.lstYourAllLoRA
        .ColumnCount = 6
        .ColumnWidths = "250 pt;0 pt;0 pt;54 pt;0 pt;0 pt" ' 5,6列目は隠す
        .Clear
    End With
    
    Me.btnRegisterLoRA.Enabled = False
    Me.btnClearLoRARGInput.Enabled = False
    Me.btnManageLoRA.Enabled = False
    Me.btnDeleteLoRARG.Enabled = False
    
    m_IsUpdating = False
End Sub

Private Sub UserForm_Activate()
    Dim loraCol As Object
    Dim item As Object
    
    Call ClearAllInputs
    Me.lstYourAllLoRA.Clear
    
    If ConfigData Is Nothing Then Exit Sub
    If Not ConfigData.Exists("LoRA_List") Then ConfigData.Add "LoRA_List", New Collection
    
    Set loraCol = ConfigData("LoRA_List")
    If Not loraCol Is Nothing Then
        If loraCol.Count > 0 Then
            For Each item In loraCol
                With Me.lstYourAllLoRA
                    .AddItem item("Alias")
                    .List(.ListCount - 1, 1) = item("Hash")
                    .List(.ListCount - 1, 2) = item("Trigger")
                    .List(.ListCount - 1, 3) = item("Strength")
                    
                    ' ★過去データの互換性維持（ModelNameが無ければ空にする）
                    If item.Exists("ModelName") Then
                        .List(.ListCount - 1, 4) = item("ModelName")
                    Else
                        .List(.ListCount - 1, 4) = ""
                    End If
                End With
            Next item
        End If
    End If
    Call lstYourAllLoRA_Click
End Sub

'-------------------------------------------------------------------------
' 全ての入力欄を更地にする処理
'-------------------------------------------------------------------------
Private Sub ClearAllInputs()
    m_IsUpdating = True
    
    Me.txtLoRAAlias.Value = ""
    Me.txtLoRAModelName.Value = "" ' ★追加
    
    Me.txtLoRAHash.Value = m_HashPlaceholder
    Me.txtLoRAHash.ForeColor = RGB(150, 150, 150)
    
    Me.cmbReccomendStrengthRG.ListIndex = -1
    Me.cmbReccomendStrengthRG.Value = ""
    
    Me.txtTriggerWordRG.Value = ""
    Me.chkTrigger_None_Rg.Value = False
    Me.cmbReccomendStrengthRG.Value = "1.0"
    Me.txtLoRANegativePrompts.Value = ""
    
    m_EditIndex = -1
    Me.btnRegisterLoRA.Caption = "Register LoRA"
    Me.btnClearLoRARGInput.Caption = "Clear Inputs"
    
    Me.txtLoRAHash.Locked = False
    Me.txtLoRAHash.BackColor = &HFFFFFF
    
    m_IsUpdating = False
    
' 登録・更新完了後のネガティブプロンプトUIの初期化（システムグレー化）
    Me.chkLoRANegativeNone.Value = True
    Me.txtLoRANegativePrompts.Value = ""
    Me.txtLoRANegativePrompts.Enabled = False
    Me.txtLoRANegativePrompts.BackColor = &H8000000F  ' ★システムグレーで封印
    
    ' 最後にボタンの状態を新規登録モード用にリセット
    If IsHandlingRegister Then Exit Sub ' ★ガード中は判定しない
    Call CheckRegisterReady
    Call lstYourAllLoRA_Click
End Sub

'-------------------------------------------------------------------------
' リアルタイム監視イベント
'-------------------------------------------------------------------------
Private Sub txtLoRAAlias_Change()
    If m_IsUpdating Then Exit Sub
    Dim currentText As String, cleanText As String, cursorPos As Integer
    currentText = Me.txtLoRAAlias.Value
    If currentText = "" Then
        Call CheckRegisterReady: Exit Sub
    End If
    cursorPos = Me.txtLoRAAlias.SelStart
    cleanText = CleanAliasString(currentText)
    If currentText <> cleanText Then
        m_IsUpdating = True
        Me.txtLoRAAlias.Value = cleanText
        m_IsUpdating = False
        If cursorPos > Len(cleanText) Then cursorPos = Len(cleanText)
        Me.txtLoRAAlias.SelStart = cursorPos
    End If
    If IsHandlingRegister Then Exit Sub ' ★ガード中は判定しない
    Call CheckRegisterReady
End Sub

Private Sub txtLoRAHash_Change()
    If m_IsUpdating Then Exit Sub
    Dim currentText As String, cleanText As String
    currentText = Me.txtLoRAHash.Value
    If currentText = m_HashPlaceholder Then Exit Sub
    cleanText = CleanHashString(currentText)
    If currentText <> cleanText Then
        m_IsUpdating = True: Me.txtLoRAHash.Value = cleanText: m_IsUpdating = False
    End If
    If IsHandlingRegister Then Exit Sub ' ★ガード中は判定しない
    Call CheckRegisterReady
End Sub

Private Sub txtTriggerWordRG_Change()
    If m_IsUpdating Then Exit Sub
    Dim currentText As String, preProcessedText As String, cleanText As String, cursorPos As Integer
    currentText = Me.txtTriggerWordRG.Value
    If currentText = "" Then
        Call CheckRegisterReady: Exit Sub
    End If
    cursorPos = Me.txtTriggerWordRG.SelStart
    preProcessedText = Replace(Replace(Replace(Replace(currentText, vbCrLf, ","), vbCr, ","), vbLf, ","), vbTab, ",")
    cleanText = CleanTriggerString(preProcessedText)
    If currentText <> cleanText Then
        m_IsUpdating = True: Me.txtTriggerWordRG.Value = cleanText: m_IsUpdating = False
        If cursorPos > Len(cleanText) Then cursorPos = Len(cleanText)
        Me.txtTriggerWordRG.SelStart = cursorPos
    End If
    If IsHandlingRegister Then Exit Sub ' ★ガード中は判定しない
    Call CheckRegisterReady
End Sub

Private Sub cmbReccomendStrengthRG_Change()
    If m_IsUpdating Then Exit Sub
    Call CheckRegisterReady
End Sub

'-------------------------------------------------------------------------
' 登録・更新ボタン（btnRegisterLoRA）の有効化条件を判定する統合ガードレール
'-------------------------------------------------------------------------
Private Sub CheckRegisterReady()
    Dim isReady As Boolean
    Dim hasAlias As Boolean
    Dim hasModelName As Boolean ' ★旧関数から吸収
    Dim hasTrigger As Boolean
    Dim hasNega As Boolean
    
    ' 1. Alias（LoRA名）と ModelName が入力されているか
    hasAlias = (Trim(Me.txtLoRAAlias.Value) <> "")
    hasModelName = (Trim(Me.txtLoRAModelName.Value) <> "")
    
    ' 2. トリガーワードの条件（文字があるか、または無しチェックがあるか）
    hasTrigger = (Trim(Me.txtTriggerWordRG.Value) <> "") Or (Me.chkTrigger_None_Rg.Value = True)
    
    ' 3. ネガティブプロンプトの条件（文字があるか、または無しチェックがあるか）
    hasNega = (Trim(Me.txtLoRANegativePrompts.Value) <> "") Or (Me.chkLoRANegativeNone.Value = True)
    
    ' 4. 総合判定（すべての必須条件を満たして初めて True）
    isReady = hasAlias And hasModelName And hasTrigger And hasNega
    
    ' 5. UIへの即時反映
    Me.btnRegisterLoRA.Enabled = isReady
    Me.btnClearLoRARGInput.Enabled = True
End Sub

'-------------------------------------------------------------------------
' [Manage LoRA] ボタン：選択したLoRAを編集モードで画面に呼び出す
'-------------------------------------------------------------------------
Private Sub btnManageLoRA_Click()
    Dim idx As Integer
    Dim selectedAlias As String, selectedHash As String
    Dim loraCol As Object, item As Object, foundItem As Object
    Dim targetNega As String ' ★追加：ネガティブプロンプト展開用
    
    idx = Me.lstYourAllLoRA.ListIndex
    If idx = -1 Then Exit Sub
    
    m_EditIndex = idx
    Me.btnRegisterLoRA.Caption = "Update LoRA"
    Me.btnClearLoRARGInput.Caption = "Cancel Update"
    
    selectedAlias = Me.lstYourAllLoRA.List(idx, 0)
    selectedHash = Me.lstYourAllLoRA.List(idx, 1)
    
    Set foundItem = Nothing
    If Not ConfigData Is Nothing Then
        If ConfigData.Exists("LoRA_List") Then
            Set loraCol = ConfigData("LoRA_List")
            For Each item In loraCol
                If item("Alias") = selectedAlias And item("Hash") = selectedHash Then
                    Set foundItem = item
                    Exit For
                End If
            Next item
        End If
    End If
    
    targetNega = "" ' ★初期化
    
    If Not foundItem Is Nothing Then
        ' -----------------------------------------------------
        ' JSONからデータを展開するルート
        ' -----------------------------------------------------
        Me.txtLoRAAlias.Value = foundItem("Alias")
        Me.txtLoRAHash.Value = foundItem("Hash")
        Me.cmbReccomendStrengthRG.Value = foundItem("Strength")
        
        ' 互換性処理: ModelName
        If foundItem.Exists("ModelName") Then
            Me.txtLoRAModelName.Value = foundItem("ModelName")
        Else
            Me.txtLoRAModelName.Value = ""
        End If
        
        ' ★互換性処理: 固有ネガティブプロンプト
        If foundItem.Exists("Negative") Then
            targetNega = Trim(foundItem("Negative"))
        End If
        
        ' トリガーワードのUI展開
        If Trim(foundItem("Trigger")) = "" Then
            Me.chkTrigger_None_Rg.Value = True
            Me.txtTriggerWordRG.Value = ""
        Else
            Me.chkTrigger_None_Rg.Value = False
            Me.txtTriggerWordRG.Value = foundItem("Trigger")
        End If
    Else
        ' -----------------------------------------------------
        ' リストボックスからデータを展開するルート（フォールバック）
        ' -----------------------------------------------------
        Me.txtLoRAAlias.Value = Me.lstYourAllLoRA.List(idx, 0)
        Me.txtLoRAHash.Value = Me.lstYourAllLoRA.List(idx, 1)
        Me.cmbReccomendStrengthRG.Value = Me.lstYourAllLoRA.List(idx, 3)
        Me.txtLoRAModelName.Value = Me.lstYourAllLoRA.List(idx, 4)
        
        ' ★追加：リストの6列目(Index 5)からネガティブを取得（エラー回避付き）
        On Error Resume Next
        targetNega = Trim(Me.lstYourAllLoRA.List(idx, 5))
        On Error GoTo 0
        
        ' トリガーワードのUI展開
        If Me.lstYourAllLoRA.List(idx, 2) = "" Then
            Me.chkTrigger_None_Rg.Value = True
            Me.txtTriggerWordRG.Value = ""
        Else
            Me.chkTrigger_None_Rg.Value = False
            Me.txtTriggerWordRG.Value = Me.lstYourAllLoRA.List(idx, 2)
        End If
    End If
    
    ' =========================================================
    ' ★追加：ネガティブプロンプトのUI連動（テキスト＆チェックボックス）
    ' =========================================================
    If targetNega = "" Then
        ' ネガティブ設定がない場合は「無し」にチェックを入れ、入力欄をロック
        Me.chkLoRANegativeNone.Value = True
        Me.txtLoRANegativePrompts.Value = ""
        Me.txtLoRANegativePrompts.Enabled = False
        Me.txtLoRANegativePrompts.BackColor = &HE0E0E0 ' 無効状態のグレー
    Else
        ' ネガティブ設定がある場合はチェックを外し、入力欄をアクティブにして展開
        Me.chkLoRANegativeNone.Value = False
        Me.txtLoRANegativePrompts.Value = targetNega
        Me.txtLoRANegativePrompts.Enabled = True
        Me.txtLoRANegativePrompts.BackColor = vbWhite
    End If
    ' =========================================================
    
    Me.txtLoRAHash.Locked = True
    Me.txtLoRAHash.BackColor = &H8000000F
    
    ' 最後に登録ボタンの有効・無効を最新化
    If IsHandlingRegister Then Exit Sub ' ★ガード中は判定しない
    Call CheckRegisterReady
End Sub

Private Sub lstYourAllLoRA_Click()
    Dim colGray As Long, colBlue As Long, colPink As Long
    colGray = &H8000000F: colBlue = &HFFE6CC: colPink = &HC8C8FF
    
    Call UpdateNoLoRALabel
    If Me.lstYourAllLoRA.ListIndex = -1 Then
        Me.btnManageLoRA.Enabled = False: Me.btnManageLoRA.BackColor = colGray
        Me.btnDeleteLoRARG.Enabled = False: Me.btnDeleteLoRARG.BackColor = colGray
        Me.btnLoRaListMoveUp.Enabled = False: Me.btnLoRaListMoveDown.Enabled = False
    Else
        Me.btnManageLoRA.Enabled = True: Me.btnManageLoRA.BackColor = colBlue
        Me.btnDeleteLoRARG.Enabled = True: Me.btnDeleteLoRARG.BackColor = colPink
        Me.btnLoRaListMoveUp.Enabled = (Me.lstYourAllLoRA.ListIndex > 0)
        Me.btnLoRaListMoveDown.Enabled = (Me.lstYourAllLoRA.ListIndex < Me.lstYourAllLoRA.ListCount - 1)
    End If
End Sub

'-------------------------------------------------------------------------
' その他UIイベント等
'-------------------------------------------------------------------------
Private Sub txtLoRAHash_Enter()
    If Me.txtLoRAHash.Value = m_HashPlaceholder Then
        m_IsUpdating = True: Me.txtLoRAHash.Value = "": Me.txtLoRAHash.ForeColor = RGB(0, 0, 0): m_IsUpdating = False
    End If
End Sub

Private Sub txtLoRAHash_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim currentText As String
    currentText = Trim(Me.txtLoRAHash.Value)
    If currentText = "" Or Replace(currentText, " ", "") = Replace(m_HashPlaceholder, " ", "") Then
        m_IsUpdating = True: Me.txtLoRAHash.Value = m_HashPlaceholder: Me.txtLoRAHash.ForeColor = RGB(150, 150, 150): m_IsUpdating = False
    End If
End Sub

Private Sub chkTrigger_None_Rg_Click()
    Dim isNone As Boolean, colGray As Long, colWhite As Long
    colGray = &H8000000F: colWhite = &HFFFFFF
    isNone = Me.chkTrigger_None_Rg.Value
    If isNone Then
        Me.txtTriggerWordRG.Enabled = False: Me.txtTriggerWordRG.BackColor = colGray: Me.txtTriggerWordRG.Value = ""
    Else
        Me.txtTriggerWordRG.Enabled = True: Me.txtTriggerWordRG.BackColor = colWhite
    End If
    Call CheckRegisterReady
End Sub

Private Sub btnClearTrigger_Click()
    Me.txtTriggerWordRG.Value = ""
    Me.txtTriggerWordRG.SetFocus
End Sub

Private Sub btnClearLoRARGInput_Click()
    If MsgBox("入力内容をクリアしますか？" & vbCrLf & "Are you sure you want to clear inputs?", vbYesNo + vbQuestion, APP_NAME) = vbYes Then Call ClearAllInputs
End Sub

Private Sub btnDeleteLoRARG_Click()
    Dim idx As Integer
    idx = Me.lstYourAllLoRA.ListIndex
    If idx = -1 Then Exit Sub
    If MsgBox("選択したLoRAを削除しますか？" & vbCrLf & "Are you sure you want to delete this LoRA?" & vbCrLf & vbCrLf & "Alias: " & Me.lstYourAllLoRA.List(idx, 0), vbYesNo + vbExclamation, APP_NAME) = vbYes Then
        Me.lstYourAllLoRA.RemoveItem idx
        Call SyncLoRAJSON
        Call ClearAllInputs
    End If
End Sub

Private Sub btnLoRaListMoveUp_Click()
    Dim idx As Integer, c As Integer, temp As String
    idx = Me.lstYourAllLoRA.ListIndex
    If idx <= 0 Then Exit Sub
    For c = 0 To 4 ' ★5列対応に変更
        temp = Me.lstYourAllLoRA.List(idx, c)
        Me.lstYourAllLoRA.List(idx, c) = Me.lstYourAllLoRA.List(idx - 1, c)
        Me.lstYourAllLoRA.List(idx - 1, c) = temp
    Next c
    Me.lstYourAllLoRA.ListIndex = idx - 1
    Call SyncLoRAJSON
End Sub

Private Sub btnLoRaListMoveDown_Click()
    Dim idx As Integer, c As Integer, temp As String
    idx = Me.lstYourAllLoRA.ListIndex
    If idx = -1 Or idx >= Me.lstYourAllLoRA.ListCount - 1 Then Exit Sub
    For c = 0 To 4 ' ★5列対応に変更
        temp = Me.lstYourAllLoRA.List(idx, c)
        Me.lstYourAllLoRA.List(idx, c) = Me.lstYourAllLoRA.List(idx + 1, c)
        Me.lstYourAllLoRA.List(idx + 1, c) = temp
    Next c
    Me.lstYourAllLoRA.ListIndex = idx + 1
    Call SyncLoRAJSON
End Sub

Private Sub btnExitWindow_Click()
    Me.Hide
End Sub

Private Sub lblHelpIcon_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not Me.lblHelpBalloon.Visible Then Me.lblHelpBalloon.Visible = True
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me.lblHelpBalloon.Visible Then Me.lblHelpBalloon.Visible = False
End Sub

Private Sub UpdateNoLoRALabel()
    If Me.lstYourAllLoRA.ListCount = 0 Then Me.lblNoLoRA.Visible = True Else Me.lblNoLoRA.Visible = False
End Sub

'-------------------------------------------------------------------------
' データ浄化・JSON同期関数群
'-------------------------------------------------------------------------
Private Function CleanAliasString(ByVal inputStr As String) As String
    Dim result As String
    result = inputStr
    result = Replace(result, ",", "_"): result = Replace(result, """", "'")
    result = Replace(result, "|", "_"): result = Replace(result, "\", "_")
    result = Replace(result, "/", "_"): result = Replace(result, ":", "_")
    result = Replace(result, "*", "_"): result = Replace(result, "?", "_")
    result = Replace(result, "<", "_"): result = Replace(result, ">", "_")
    CleanAliasString = result
End Function

Private Function CleanHashString(ByVal inputStr As String) As String
    Dim i As Integer, ch As String, result As String
    inputStr = StrConv(inputStr, vbNarrow)
    result = ""
    For i = 1 To Len(inputStr)
        ch = Mid(inputStr, i, 1)
        If ch Like "[a-zA-Z0-9]" Or ch = "?" Or ch = "=" Or ch = "-" Or ch = "_" Then result = result & ch
    Next i
    CleanHashString = result
End Function

Private Function CleanTriggerString(ByVal inputStr As String) As String
    Dim i As Integer, ch As String, result As String
    inputStr = StrConv(inputStr, vbNarrow)
    result = ""
    For i = 1 To Len(inputStr)
        ch = Mid(inputStr, i, 1)
        If ch Like "[a-zA-Z0-9]" Or InStr(" ,._-()[]{}:+/", ch) > 0 Then result = result & ch
    Next i
    CleanTriggerString = result
End Function

'-------------------------------------------------------------------------
' メモリ上のリストボックスの内容をJSON（LoRA_List）へ同期して保存する
'-------------------------------------------------------------------------
Private Sub SyncLoRAJSON()
    Dim newCol As Collection, dict As Object, i As Long
    Dim negaVal As String ' ★追加：ネガティブプロンプト取得用
    
    Set newCol = New Collection
    
    If Me.lstYourAllLoRA.ListCount > 0 Then
        For i = 0 To Me.lstYourAllLoRA.ListCount - 1
            Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "Alias", Me.lstYourAllLoRA.List(i, 0)
            dict.Add "Hash", Me.lstYourAllLoRA.List(i, 1)
            dict.Add "Trigger", Me.lstYourAllLoRA.List(i, 2)
            dict.Add "Strength", Me.lstYourAllLoRA.List(i, 3)
            dict.Add "ModelName", Me.lstYourAllLoRA.List(i, 4)
            
            ' =========================================================
            ' ★追加：6列目（Index 5）からネガティブプロンプトを取得
            ' ※旧バージョンで登録したデータなど、6列目が存在しない場合のエラーを完全に防ぐ
            ' =========================================================
            negaVal = ""
            On Error Resume Next
            negaVal = Trim(Me.lstYourAllLoRA.List(i, 5))
            On Error GoTo 0
            
            dict.Add "Negative", negaVal
            
            newCol.Add dict
        Next i
    End If
    
    If ConfigData Is Nothing Then Exit Sub
    If ConfigData.Exists("LoRA_List") Then
        Set ConfigData("LoRA_List") = newCol
    Else
        ConfigData.Add "LoRA_List", newCol
    End If
    Call SaveConfigJSON
End Sub


'-------------------------------------------------------------------------
' [神機能] Hash自動取得 ＆ メタデータ抽出 (連続読込時の残骸防衛・手動入力特化版)
'-------------------------------------------------------------------------
Private Sub txtLoRAHash_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim filePath As Variant, resultHash As String, extractedModelName As String
    Dim dummyTriggers As String ' ★関数呼び出しの引数エラーを防ぐためのダミー変数
    
    filePath = Application.GetOpenFilename("Safetensors Files (*.safetensors), *.safetensors", , "ハッシュを計算するLoRAを選択 (Select a LoRA file to calculate Hash)")
    ' キャンセルされた場合は、現在入力されている内容を守るため何もせず終了
    If filePath = False Then Exit Sub
    
    Application.Cursor = xlWait
    
    ' 1. ハッシュの計算
    resultHash = GetAutoV2Hash(CStr(filePath))
    
    ' 2. メタデータ（正式名称のみ）の抽出
    extractedModelName = GetSafetensorsMetadata(CStr(filePath), dummyTriggers)
    
    ' 読み込み成功時の処理
    If Left(resultHash, 5) <> "ERROR" Then
        Cancel = True
        m_IsUpdating = True ' イベント暴走阻止フラグをON
        
        ' ========================================================
        ' ★古いファイル情報の残骸を完全に一掃（更地化）
        ' ========================================================
        Me.txtLoRAHash.Value = ""
        Me.txtLoRAModelName.Value = ""
        Me.txtLoRAAlias.Value = ""
        Me.txtTriggerWordRG.Value = ""
        
        ' 3. 新しいファイルから取得したデータを安全に代入
        Me.txtLoRAHash.Value = resultHash
        Me.txtLoRAHash.ForeColor = RGB(0, 0, 0) ' 文字色を黒に
        
        ' 正式名称(ModelName)をセット
        Me.txtLoRAModelName.Value = extractedModelName
        
        ' エイリアスには、新しく読み込んだファイルの正式名称を確実に無条件セット
        Me.txtLoRAAlias.Value = extractedModelName
        
        ' ========================================================
        ' ★トリガーワード手動コピペ入力のためのUIリセット
        ' （Noneチェックを外し、テキストボックスを有効化して白紙で待機させる）
        ' ========================================================
        Me.chkTrigger_None_Rg.Value = False
        Me.txtTriggerWordRG.Enabled = True
        Me.txtTriggerWordRG.BackColor = &HFFFFFF ' 背景色を白に
        
        m_IsUpdating = False ' フラグを解除
        
        ' 完了メッセージの作成（トリガーワード関係のメッセージも削除しスッキリと）
        Dim msgText As String
        msgText = "Auto V2 ハッシュ値とモデル名を自動取得しました！" & vbCrLf & _
                  "Metadata auto-retrieved successfully!" & vbCrLf & vbCrLf & _
                  "Model: " & extractedModelName & vbCrLf & _
                  "Hash: " & resultHash
                  
        MsgBox msgText, vbInformation, APP_NAME
    Else
        ' エラー時は既存の入力を壊さないようそのままにしてアラートを出す
        MsgBox "ハッシュ値の取得に失敗しました。" & vbCrLf & "Failed to retrieve the hash." & vbCrLf & vbCrLf & "Detail: " & resultHash, vbExclamation, APP_NAME
    End If
    Application.Cursor = xlDefault
End Sub

'-------------------------------------------------------------------------
' [内部機能] Safetensorsのファイルハッシュ(AutoV2)を取得する (2GB超過ファイル対応版)
'-------------------------------------------------------------------------
Private Function GetAutoV2Hash(ByVal filePath As String) As String
    Dim fileNum As Integer, buffer() As Byte
    Dim chunkSize As Long, bytesToRead As Long
    Dim fileLen As Double, bytesRemaining As Double ' ★修正1：Long(2GBの壁)からDoubleへ変更
    
    #If VBA7 Then
        Dim hAlg As LongPtr, hHash As LongPtr
    #Else
        Dim hAlg As Long, hHash As Long
    #End If
    Dim algName() As Byte, hashVal(0 To 31) As Byte
    Dim status As Long, i As Long, resultStr As String
    Dim fso As Object, fileObj As Object
    
    If Dir(filePath) = "" Then GetAutoV2Hash = "ERROR: File Not Found": Exit Function
    
    ' ★修正2：LOF関数の2GB超過エラーを防ぐため、FSOを利用して正確なファイルサイズを取得
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObj = fso.GetFile(filePath)
    fileLen = fileObj.Size
    
    algName = BCRYPT_SHA256_ALGORITHM & Chr(0)
    status = BCryptOpenAlgorithmProvider(hAlg, VarPtr(algName(0)), 0&, 0&)
    If status <> STATUS_SUCCESS Then GetAutoV2Hash = "ERROR: OpenAlgorithmProvider Failed": Exit Function
    
    status = BCryptCreateHash(hAlg, hHash, 0&, 0&, 0&, 0&, 0&)
    If status <> STATUS_SUCCESS Then BCryptCloseAlgorithmProvider hAlg, 0&: GetAutoV2Hash = "ERROR: CreateHash Failed": Exit Function
    
    On Error GoTo FileError
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    chunkSize = 1048576 ' 1MB
    bytesRemaining = fileLen
    
    Do While bytesRemaining > 0
        If bytesRemaining < chunkSize Then bytesToRead = CLng(bytesRemaining) Else bytesToRead = chunkSize
        ReDim buffer(0 To bytesToRead - 1)
        Get #fileNum, , buffer
        status = BCryptHashData(hHash, buffer(0), bytesToRead, 0&)
        If status <> STATUS_SUCCESS Then Close #fileNum: GoTo HashError
        bytesRemaining = bytesRemaining - bytesToRead
        DoEvents
    Loop
    Close #fileNum
    On Error GoTo 0
    
    status = BCryptFinishHash(hHash, hashVal(0), 32, 0&)
    If status = STATUS_SUCCESS Then
        For i = 0 To 31: resultStr = resultStr & Right$("0" & Hex(hashVal(i)), 2): Next i
        GetAutoV2Hash = LCase(Left(resultStr, 10))
    Else
        GetAutoV2Hash = "ERROR: FinishHash Failed"
    End If
    
    BCryptDestroyHash hHash: BCryptCloseAlgorithmProvider hAlg, 0&
    Exit Function
FileError:
    If fileNum > 0 Then Close #fileNum
HashError:
    If hHash <> 0 Then BCryptDestroyHash hHash
    If hAlg <> 0 Then BCryptCloseAlgorithmProvider hAlg, 0&
    GetAutoV2Hash = "ERROR: Processing Failed"
End Function

'-------------------------------------------------------------------------
' [内部機能] SafetensorsのJSONヘッダから正式名称(ModelName)のみを抽出する（超軽量版）
' ※トリガーワードの自動抽出機能は廃止（ユーザーの手動コピペ入力前提）
'-------------------------------------------------------------------------
Private Function GetSafetensorsMetadata(ByVal filePath As String, ByRef outTriggers As String) As String
    Dim fileNum As Integer, headerSizeBytes(0 To 7) As Byte, headerSize As Long
    Dim headerJsonBytes() As Byte, headerJson As String, ado As Object
    Dim regEx As Object, matches As Object, fso As Object, defaultName As String
    Dim searchChunk As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    defaultName = fso.GetBaseName(filePath)
    
    ' ★自動抽出を廃止したため、常に空文字を返す（呼び出し元のUI生成を安全にスキップさせます）
    outTriggers = ""
    
    On Error GoTo Fallback
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    Get #fileNum, 1, headerSizeBytes
    
    ' 4バイト目まで計算に含め、真のヘッダサイズを正確に取得
    headerSize = CLng(headerSizeBytes(0)) + _
                 CLng(headerSizeBytes(1)) * 256& + _
                 CLng(headerSizeBytes(2)) * 65536 + _
                 CLng(headerSizeBytes(3)) * 16777216
    
    If headerSize <= 0 Or headerSize > 52428800 Then Close #fileNum: GoTo Fallback
    
    ReDim headerJsonBytes(0 To headerSize - 1)
    Get #fileNum, 9, headerJsonBytes
    Close #fileNum
    On Error GoTo 0
    
    ' ADODBでUTF-8デコード
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 1: ado.Open: ado.Write headerJsonBytes: ado.Position = 0
    ado.Type = 2: ado.Charset = "UTF-8": headerJson = ado.ReadText: ado.Close
    
    ' 巨大JSON対策：最初と最後だけ切り出して検索空間を極小化する
    If Len(headerJson) > 500000 Then
        searchChunk = Left(headerJson, 250000) & Right(headerJson, 250000)
    Else
        searchChunk = headerJson
    End If
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = False
    regEx.IgnoreCase = True
    
    ' =========================================================================
    ' 正式名称(ModelName)の抽出ロジックのみ実行
    ' =========================================================================
    regEx.Pattern = """modelspec\.title""\s*:\s*""([^""]+)"""
    Set matches = regEx.Execute(searchChunk)
    If matches.Count > 0 Then GetSafetensorsMetadata = matches(0).subMatches(0): Exit Function
    
    regEx.Pattern = """ss_output_name""\s*:\s*""([^""]+)"""
    Set matches = regEx.Execute(searchChunk)
    If matches.Count > 0 Then GetSafetensorsMetadata = matches(0).subMatches(0): Exit Function

Fallback:
    If fileNum > 0 Then Close #fileNum
    GetSafetensorsMetadata = defaultName
End Function


