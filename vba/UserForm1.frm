VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "シート移動"
   ClientHeight    =   3228
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4644
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const onDebug As Boolean = False
Const hiddenSheetStr As String = "*"


Private Sub UserForm_Initialize()
    If onDebug Then
        MsgBox ("#UserForm_Initialize")
    End If
    
    With lvSheetList
        .ColumnHeaders.Add , , "シート名", .Width
    End With
    
    Application.OnKey "{ESC}", "cbCancel_Click"

End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If onDebug Then
        MsgBox ("#UserForm_KeyUp :: onKey ->" & KeyCode)
    End If
    
    If KeyCode = vbKeyEscape Then
        cbCancel_Click
    End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If onDebug Then
        MsgBox ("#UserForm_QueryClose")
    End If
    
    Application.OnKey "{ESC}"    '設定取り消し

End Sub

Private Sub UserForm_Activate()
    If onDebug Then
        MsgBox ("#UserForm_Activate")
    End If
    
    Dim sheet As Object
    Dim actSheetName As String
    Dim str As String
    
    actSheetName = ActiveSheet.Name
    
    With lvSheetList
        .ListItems.Clear
        
        For Each sheet In ActiveWorkbook.Sheets
        
            ' 非表示のシートには"*"を付与
            If sheet.Visible = xlSheetVisible Then
                str = ""
            Else
                str = hiddenSheetStr
            End If
            
            .ListItems.Add.Text = str & sheet.Name
            
        Next
    End With
    
End Sub


Private Sub cbOk_Click()
    If onDebug Then
        MsgBox ("#cbOk_Click")
    End If
    
    Dim selectedSheet As Object
    Dim sheetName As String
    
    sheetName = lvSheetList.SelectedItem.Text
    
    If Left(sheetName, 1) = hiddenSheetStr Then
        sheetName = Replace(sheetName, hiddenSheetStr, "")
    End If
    
    Set selectedSheet = ActiveWorkbook.Sheets(sheetName)
    
    With selectedSheet
        .Visible = xlSheetVisible
        .Activate
    End With
    
    Unload Me
    
End Sub


Private Sub cbCancel_Click()
    If onDebug Then
        MsgBox ("#cbCancel_Click")
    End If
    
    Unload Me

End Sub


Private Sub lvSheetList_DblClick()
    If onDebug Then
        MsgBox ("#lvSheetList_DblClick")
    End If
    
    cbOk_Click    'リストビューダブルクリック時はOKボタン押下時の処理と同じ

End Sub


Private Sub lvSheetList_KeyUp(KeyCode As Integer, ByVal Shift As Integer)

    If onDebug Then
        MsgBox ("#lvSheetList_KeyUp :: onKey ->" & KeyCode)
    End If
    
    If KeyCode = vbKeyReturn Then      ' Enterキー押下時
        cbOk_Click
    ElseIf KeyCode = vbKeyEscape Then  ' ESCキー押下時
        cbCancel_Click
    End If
    
End Sub


