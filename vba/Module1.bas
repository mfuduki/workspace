Attribute VB_Name = "Module1"
Option Explicit

'
' �J���Ă��郏�[�N�u�b�N�̃V�[�g���ړ����邽�߂�
' ���[�U�t�H�[����\�����A�I�������V�[�g�ֈړ�����B
'
' [Ctrl]+m
'
Sub �V�[�g�ړ�()

    UserForm1.Show
    
End Sub


'
' �l�݂̂̓\��t�� Macro
'
' Keyboard Shortcut: Ctrl+Shift+v
'
Sub �l�݂̂̓\��t��()
Attribute �l�݂̂̓\��t��.VB_ProcData.VB_Invoke_Func = "V\n14"
'    On Error Resume Next
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
    
    ActiveSheet.PasteSpecial Format:="Unicode �e�L�X�g", Link:=False, DisplayAsIcon:=False

End Sub

'
' �r���\���ؑ� Macro
'
' Keyboard Shortcut: Ctrl+Shift+e
'
Sub �g���\���ؑ�()
Attribute �g���\���ؑ�.VB_ProcData.VB_Invoke_Func = "E\n14"

    If ActiveWindow.DisplayGridlines = True Then
        ActiveWindow.DisplayGridlines = False
    Else
        ActiveWindow.DisplayGridlines = True
    End If
End Sub

'
' �n�C�p�[�����N�W�����v Macro
'
' Keyboard Shortcut: Ctrl+Shift+J
'
Sub �n�C�p�[�����N�W�����v()
Attribute �n�C�p�[�����N�W�����v.VB_ProcData.VB_Invoke_Func = "J\n14"
    
'    ActiveCell.Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    
'    ActiveSheet.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    
End Sub


' �A�N�e�B�u�V�[�g�̘g���\����؂�ւ���B
'
' [Ctrl]+[Shift]+j
Sub �g���\���ؑ�()

    If ActiveWindow.DisplayGridlines = True Then
        ActiveWindow.DisplayGridlines = False
    Else
        ActiveWindow.DisplayGridlines = True
    End If
End Sub


' �I�����Ă���V�[�g���̌��o���F���N���A����B
' �i�����I���j
'
' [Ctrl]+[Shift]+q
'
Sub �V�[�g���̌��o���F�N���A()

    Dim sheet As Worksheet
    
    For Each sheet In ActiveWindow.SelectedSheets
        sheet.Tab.ColorIndex = xlColorIndexNone
    Next sheet
    
End Sub

' �I�������I�[�g�V�F�C�v��I���ς݃Z���̍��E���S�ɍ��킹��B
'
' [Ctrl]+[Shift]+L
'
Sub �I�[�g�V�F�C�v���`���Z�����E���S()

    If TypeName(Selection) = "Range" Then Exit Sub
    
    Dim shape As Object
    
    For Each shape In Selection.ShapeRange
        With shape
            .Left = ActiveCell.Left + (ActiveCell.Width - .Width) / 2
        End With
    Next shape

End Sub

' �A�N�e�B�u�Z���̃R�����g�ʒu�����Z�b�g
'
'
Sub �R�����g�ʒu���Z�b�g()
    
End Sub

