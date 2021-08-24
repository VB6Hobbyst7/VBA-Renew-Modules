Attribute VB_Name = "ModRenewModule"
Option Explicit

Sub �K�v���W���[���X�V()
'���s�T���v���Ȃǃ��W���[������ɍX�V���������[�N�u�b�N�ɂāA�N�����C�x���g�Ŏ��s����悤�ɂ���B
'20210824

    '�w�胆�[�U�[�łȂ��Ɠ��삵�Ȃ��悤�ɂ��Ă���
    If GetUserName <> "YF215008" Then
        Exit Sub
    End If
    
    Stop
    
    '�����[�N�u�b�N�ɍ��킹�ē��e��ύX���邱��
    '������������������������������������������������������
    Dim ModuleList$(1 To 5) '����������������������������������������������
    ModuleList(1) = "frmKaiso.frm"
    ModuleList(2) = "ModExtProcedure.bas"
    ModuleList(3) = "classModule.cls"
    ModuleList(4) = "classProcedure.cls"
    ModuleList(5) = "classVBProject.cls"
    '������������������������������������������������������
    
    Dim TmpModulePath$
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    For I = 1 To UBound(ModuleList)
        Call DeleteModule(ModuleList(I), ThisWorkbook)
        TmpModulePath = ThisWorkbook.Path & "\" & ModuleList(I)
        Call ImportModule(TmpModulePath, ThisWorkbook)
    Next I
    
    '�m�F���b�Z�[�W�\��
    Dim Message$
    Message = "���L���W���[�����X�V���܂���"
    For I = 1 To UBound(ModuleList)
        Message = Message & vbLf & "��" & ModuleList(I)
    Next I
    
    MsgBox (Message)

End Sub

Sub ImportModule(ModulePath$, Optional TargetBook As Workbook)
'�w��p�X�̃��W���[�����C���|�[�g����
'20210823

'����
'ModulePath :�C���|�[�g���郂�W���[���̃t���p�X
'TargetBook :�C���|�[�g��̃��[�N�u�b�N�B�����͂Ȃ玩�u�b�N(ThisWorkBook)��Ώ�

    '�����`�F�b�N
    If TargetBook Is Nothing Then '�Ώۃu�b�N�������͂Ȃ炱�̃u�b�N��ΏۂƂ���B
        Set TargetBook = ThisWorkbook
    End If
    
    '�w��p�X�̃��W���[���̑��݊m�F
    If Dir(ModulePath) = "" Then
        MsgBox ("�u" & ModulePath & "�v" & vbLf & _
               "�͑��݂��܂���")
        Stop
        End
    End If
    
    '���W���[���̖��O�擾
    Dim ModuleName$
    ModuleName = GetFileName(ModulePath)
    ModuleName = Split(ModuleName, ".")(0)
    
    '�C���|�[�g���郂�W���[�������ɂ��邩�m�F
    Dim TmpModuleName$
    Dim TargetModule As VBComponent
    Dim TmpModule As VBComponent
    Dim Hantei As Boolean
    Hantei = False
    For Each TmpModule In TargetBook.VBProject.VBComponents
        If TmpModule.Name = ModuleName Then
            Hantei = True
            Set TargetModule = TmpModule '�����Ώۂ̃��W���[���ݒ�
            Exit For
        End If
    Next
    
    '�C���|�[�g���郂�W���[�������ɑ��݂���ꍇ�͊m�F�̃��b�Z�[�W
    If Hantei Then
        If MsgBox("���W���[��" & "�u" & ModulePath & "�v" & vbLf & _
               "�͂��łɃv���W�F�N�g�ɑ��݂��܂��B" & _
               "�㏑���C���|�[�g���܂����H", vbYesNo) = vbYes Then
               
            Call DeleteModule(ModuleName, TargetBook)
        Else
            Exit Sub
        End If
    End If
    
    '���W���[���̃C���|�[�g
    Call TargetBook.VBProject.VBComponents.Import(ModulePath)

End Sub

Sub DeleteModule(ModuleNameWithExtention$, Optional TargetBook As Workbook)
'�w�胂�W���[������������
'20210823

'����
'ModuleNameWithExtention    :�������郂�W���[���̖��O�B�g���q�����邱�Ɓi��FModule1.bas�j
'TargetBook                 :�C���|�[�g��̃��[�N�u�b�N�B�����͂Ȃ玩�u�b�N(ThisWorkBook)��Ώ�

    '�����`�F�b�N
    If TargetBook Is Nothing Then '�Ώۃu�b�N�������͂Ȃ炱�̃u�b�N��ΏۂƂ���B
        Set TargetBook = ThisWorkbook
    End If
    
    '���W���[�������g���q���̏ꍇ
    Dim ModuleName$, ModuleType$
    If InStr(1, ModuleNameWithExtention, ".") = 0 Then
        MsgBox ("�uModuleNameWithExtention�v�͊g���q���t���ē��͂��Ă��������B" & vbLf & _
               "�u**.frm�v�����[�U�[�t�H�[��" & vbLf & _
               "�u**.bas�v���W�����W���[��" & vbLf & _
               "�u**.cls�v���N���X���W���[��")
        Stop
        End
    Else
        ModuleName = Split(ModuleNameWithExtention, ".")(0)
        ModuleType = StrConv(Split(ModuleNameWithExtention, ".")(1), vbNarrow)
        
        If ModuleType <> "frm" And ModuleType <> "bas" And ModuleType <> "cls" Then
            MsgBox ("�u" & ModuleType & "�v�͊g���q�Ƃ��ĔF���ł��܂���B" & vbLf & _
                   "�u**.frm�v�����[�U�[�t�H�[��" & vbLf & _
                   "�u**.bas�v���W�����W���[��" & vbLf & _
                   "�u**.cls�v���N���X���W���[��")
            Stop
            End
        End If
    End If

    '�w�薼�̃��W���[�������邩�m�F
    Dim TmpModuleName$
    Dim TargetModule As VBComponent
    Dim TmpModule As VBComponent
    Dim TmpModuleType$
    Dim Hantei As Boolean
    Hantei = False
    For Each TmpModule In TargetBook.VBProject.VBComponents
        TmpModuleType = ���W���[����ޔ���(TmpModule)
        If TmpModule.Name = ModuleName And TmpModuleType = ModuleType Then
            Hantei = True
            Set TargetModule = TmpModule '�����Ώۂ̃��W���[���ݒ�
            Exit For
        End If
    Next
    
    '�w�薼�̃��W���[����������Ȃ������ꍇ�͏I��
    If Hantei = False Then
        MsgBox ("���W���[��" & "�u" & ModuleName & "�v" & vbLf & _
               "�͌�����܂���ł���")
        Exit Sub
    End If
    
    '���W���[���̏���
    Call TargetBook.VBProject.VBComponents.Remove(TargetModule)
    
End Sub

Private Function GetFileName$(FilePath$)
'�t�@�C���̃t���p�X����t�@�C�����擾
'�֐��v���o���p
'20210824
    
    Dim Output$
    Dim TmpList
    TmpList = Split(FilePath, "\")
    Output = TmpList(UBound(TmpList))
    GetFileName = Output
    
End Function

Private Function ���W���[����ޔ���(InputModule As VBComponent)
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
        Output = "ActiveX �f�U�C�i"
    Case 100
        Output = "Document ���W���[��"
    Case Else
        MsgBox ("���W���[����ނ�����ł��܂���")
        Stop
    End Select
    
    ���W���[����ޔ��� = Output
    
End Function

Private Function GetUserName$()
'���݉ғ����Ă���Windows�Ƀ��O�C�����Ă��郆�[�U�[�����擾����
'20210726
    GetUserName = Environ("USERNAME")

End Function
