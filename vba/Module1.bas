Attribute VB_Name = "Module1"
Option Explicit

'
' 開いているワークブックのシートを移動するための
' ユーザフォームを表示し、選択したシートへ移動する。
'
' [Ctrl]+m
'
Sub シート移動()

    UserForm1.Show
    
End Sub


'
' 値のみの貼り付け Macro
'
' Keyboard Shortcut: Ctrl+Shift+v
'
Sub 値のみの貼り付け()
Attribute 値のみの貼り付け.VB_ProcData.VB_Invoke_Func = "V\n14"
'    On Error Resume Next
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
    
    ActiveSheet.PasteSpecial Format:="Unicode テキスト", Link:=False, DisplayAsIcon:=False

End Sub

'
' 罫線表示切替 Macro
'
' Keyboard Shortcut: Ctrl+Shift+e
'
Sub 枠線表示切替()
Attribute 枠線表示切替.VB_ProcData.VB_Invoke_Func = "E\n14"

    If ActiveWindow.DisplayGridlines = True Then
        ActiveWindow.DisplayGridlines = False
    Else
        ActiveWindow.DisplayGridlines = True
    End If
End Sub

'
' ハイパーリンクジャンプ Macro
'
' Keyboard Shortcut: Ctrl+Shift+J
'
Sub ハイパーリンクジャンプ()
Attribute ハイパーリンクジャンプ.VB_ProcData.VB_Invoke_Func = "J\n14"
    
'    ActiveCell.Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    
'    ActiveSheet.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    
End Sub


' アクティブシートの枠線表示を切り替える。
'
' [Ctrl]+[Shift]+j
Sub 枠線表示切替()

    If ActiveWindow.DisplayGridlines = True Then
        ActiveWindow.DisplayGridlines = False
    Else
        ActiveWindow.DisplayGridlines = True
    End If
End Sub


' 選択しているシート名の見出し色をクリアする。
' （複数選択可）
'
' [Ctrl]+[Shift]+q
'
Sub シート名の見出し色クリア()

    Dim sheet As Worksheet
    
    For Each sheet In ActiveWindow.SelectedSheets
        sheet.Tab.ColorIndex = xlColorIndexNone
    Next sheet
    
End Sub

' 選択したオートシェイプを選択済みセルの左右中心に合わせる。
'
' [Ctrl]+[Shift]+L
'
Sub オートシェイプ整形→セル左右中心()

    If TypeName(Selection) = "Range" Then Exit Sub
    
    Dim shape As Object
    
    For Each shape In Selection.ShapeRange
        With shape
            .Left = ActiveCell.Left + (ActiveCell.Width - .Width) / 2
        End With
    Next shape

End Sub

' アクティブセルのコメント位置をリセット
'
'
Sub コメント位置リセット()
    
End Sub

