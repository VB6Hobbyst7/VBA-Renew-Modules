Attribute VB_Name = "ModRenewModule"
Option Explicit

Sub 必要モジュール更新()
'実行サンプルなどモジュールを常に更新したいワークブックにて、起動時イベントで実行するようにする。
'20210824

    '指定ユーザーでないと動作しないようにしておく
    If GetUserName <> "YF215008" Then
        Exit Sub
    End If
    
    Stop
    
    '↓ワークブックに合わせて内容を変更すること
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim ModuleList$(1 To 5) '←←←←←←←←←←←←←←←←←←←←←←←
    ModuleList(1) = "frmKaiso.frm"
    ModuleList(2) = "ModExtProcedure.bas"
    ModuleList(3) = "classModule.cls"
    ModuleList(4) = "classProcedure.cls"
    ModuleList(5) = "classVBProject.cls"
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    
    Dim TmpModulePath$
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    For I = 1 To UBound(ModuleList)
        Call DeleteModule(ModuleList(I), ThisWorkbook)
        TmpModulePath = ThisWorkbook.Path & "\" & ModuleList(I)
        Call ImportModule(TmpModulePath, ThisWorkbook)
    Next I
    
    '確認メッセージ表示
    Dim Message$
    Message = "下記モジュールを更新しました"
    For I = 1 To UBound(ModuleList)
        Message = Message & vbLf & "○" & ModuleList(I)
    Next I
    
    MsgBox (Message)

End Sub

Sub ImportModule(ModulePath$, Optional TargetBook As Workbook)
'指定パスのモジュールをインポートする
'20210823

'引数
'ModulePath :インポートするモジュールのフルパス
'TargetBook :インポート先のワークブック。未入力なら自ブック(ThisWorkBook)を対象

    '引数チェック
    If TargetBook Is Nothing Then '対象ブックが未入力ならこのブックを対象とする。
        Set TargetBook = ThisWorkbook
    End If
    
    '指定パスのモジュールの存在確認
    If Dir(ModulePath) = "" Then
        MsgBox ("「" & ModulePath & "」" & vbLf & _
               "は存在しません")
        Stop
        End
    End If
    
    'モジュールの名前取得
    Dim ModuleName$
    ModuleName = GetFileName(ModulePath)
    ModuleName = Split(ModuleName, ".")(0)
    
    'インポートするモジュールが既にあるか確認
    Dim TmpModuleName$
    Dim TargetModule As VBComponent
    Dim TmpModule As VBComponent
    Dim Hantei As Boolean
    Hantei = False
    For Each TmpModule In TargetBook.VBProject.VBComponents
        If TmpModule.Name = ModuleName Then
            Hantei = True
            Set TargetModule = TmpModule '消去対象のモジュール設定
            Exit For
        End If
    Next
    
    'インポートするモジュールが既に存在する場合は確認のメッセージ
    If Hantei Then
        If MsgBox("モジュール" & "「" & ModulePath & "」" & vbLf & _
               "はすでにプロジェクトに存在します。" & _
               "上書きインポートしますか？", vbYesNo) = vbYes Then
               
            Call DeleteModule(ModuleName, TargetBook)
        Else
            Exit Sub
        End If
    End If
    
    'モジュールのインポート
    Call TargetBook.VBProject.VBComponents.Import(ModulePath)

End Sub

Sub DeleteModule(ModuleNameWithExtention$, Optional TargetBook As Workbook)
'指定モジュールを消去する
'20210823

'引数
'ModuleNameWithExtention    :消去するモジュールの名前。拡張子をつけること（例：Module1.bas）
'TargetBook                 :インポート先のワークブック。未入力なら自ブック(ThisWorkBook)を対象

    '引数チェック
    If TargetBook Is Nothing Then '対象ブックが未入力ならこのブックを対象とする。
        Set TargetBook = ThisWorkbook
    End If
    
    'モジュール名が拡張子つきの場合
    Dim ModuleName$, ModuleType$
    If InStr(1, ModuleNameWithExtention, ".") = 0 Then
        MsgBox ("「ModuleNameWithExtention」は拡張子も付けて入力してください。" & vbLf & _
               "「**.frm」→ユーザーフォーム" & vbLf & _
               "「**.bas」→標準モジュール" & vbLf & _
               "「**.cls」→クラスモジュール")
        Stop
        End
    Else
        ModuleName = Split(ModuleNameWithExtention, ".")(0)
        ModuleType = StrConv(Split(ModuleNameWithExtention, ".")(1), vbNarrow)
        
        If ModuleType <> "frm" And ModuleType <> "bas" And ModuleType <> "cls" Then
            MsgBox ("「" & ModuleType & "」は拡張子として認識できません。" & vbLf & _
                   "「**.frm」→ユーザーフォーム" & vbLf & _
                   "「**.bas」→標準モジュール" & vbLf & _
                   "「**.cls」→クラスモジュール")
            Stop
            End
        End If
    End If

    '指定名のモジュールがあるか確認
    Dim TmpModuleName$
    Dim TargetModule As VBComponent
    Dim TmpModule As VBComponent
    Dim TmpModuleType$
    Dim Hantei As Boolean
    Hantei = False
    For Each TmpModule In TargetBook.VBProject.VBComponents
        TmpModuleType = モジュール種類判定(TmpModule)
        If TmpModule.Name = ModuleName And TmpModuleType = ModuleType Then
            Hantei = True
            Set TargetModule = TmpModule '消去対象のモジュール設定
            Exit For
        End If
    Next
    
    '指定名のモジュールが見つからなかった場合は終了
    If Hantei = False Then
        MsgBox ("モジュール" & "「" & ModuleName & "」" & vbLf & _
               "は見つかりませんでした")
        Exit Sub
    End If
    
    'モジュールの消去
    Call TargetBook.VBProject.VBComponents.Remove(TargetModule)
    
End Sub

Private Function GetFileName$(FilePath$)
'ファイルのフルパスからファイル名取得
'関数思い出し用
'20210824
    
    Dim Output$
    Dim TmpList
    TmpList = Split(FilePath, "\")
    Output = TmpList(UBound(TmpList))
    GetFileName = Output
    
End Function

Private Function モジュール種類判定(InputModule As VBComponent)
'http://officetanaka.net/excel/vba/vbe/04.htm

    Dim Output$
    Select Case InputModule.Type
    Case 1
        Output = "bas"
    Case 2
        Output = "cls"
    Case 3
        Output = "frm"
    Case 11
        Output = "ActiveX デザイナ"
    Case 100
        Output = "Document モジュール"
    Case Else
        MsgBox ("モジュール種類が判定できません")
        Stop
    End Select
    
    モジュール種類判定 = Output
    
End Function

Private Function GetUserName$()
'現在稼働しているWindowsにログインしているユーザー名を取得する
'20210726
    GetUserName = Environ("USERNAME")

End Function
