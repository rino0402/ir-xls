Attribute VB_Name = "MenuMacro"
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++
'メニュー用マクロ
'++++++++++++++++++++++++++++++++++++++++++++++++++
Dim Frag(10) As Boolean

'++++++++++++++++++++++++++++++++++++++++++++++++++
'"MENU1"の初期化
Sub Menu1_Init()
Attribute Menu1_Init.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook.Worksheets("LINK")
    .Range("B4,E4,F4").Value = 0
    .Range("G4").Value = 1
    .Range("D3:D31,I4:I8,I11:I13,K4:K6,M4:M8,O4:O5,Q4:Q5").Value = False
    .Range("F9").ClearContents
    SHEET_Check
    OPT_SYORI
    OPT_HYODAI
    EDT_HYODAI
    CHK_PRINT1
    CHK_BUSYO_INIT
    SPN_BUSU
    SetMessage
  End With
End Sub

Private Sub SetMessage()
    Dim wbk As Workbook
    Dim wst As Worksheet
    For Each wbk In Application.Workbooks
        For Each wst In wbk.Worksheets
            If wst.Name = "FILE" Then
                'Debug.Print wbk.Name & " " & wst.Name
                'Debug.Print wst.Range("I10")
'                Call wst.Unprotect(password:="sdc2035")
'                wst.Range("I10") = "2008.04.28 新項目対応"
'                wst.Range("I10").Font.Color = vbWhite
'                Call wst.Protect(password:="SDC2035")
                Exit For
            End If
        Next
    Next
End Sub


'++++++++++++++++++++++++++++++++++++++++++++++++++
'処理月
Sub OPT_TUKI()
Attribute OPT_TUKI.VB_ProcData.VB_Invoke_Func = " \n14"
  Frag_Check
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'部署（センター）名
Sub CHK_BUSYO()
Attribute CHK_BUSYO.VB_ProcData.VB_Invoke_Func = " \n14"
  Frag_Check
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'一般販売
Sub CHK_BUSYOIPPAN()
Attribute CHK_BUSYOIPPAN.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim Ch As Boolean
  
  With ThisWorkbook
    Ch = Application.Or(.Worksheets("LINK").Range("D28:D31").Value)
    .DialogSheets("Menu1").CheckBoxes("P3_1_CHK").Enabled = Ch
    .DialogSheets("Menu1").CheckBoxes("P3_2_CHK").Enabled = Ch
    .DialogSheets("Menu1").CheckBoxes("P3_3_CHK").Enabled = Ch
    .DialogSheets("Menu1").CheckBoxes("P3_4_CHK").Enabled = Ch
    .DialogSheets("Menu1").CheckBoxes("P3_5_CHK").Enabled = Ch
    Frag_Check
  End With
End Sub

' b.OnAction = "CHK_BUSYOIPPAN"
' b.OnAction = "CHK_BUSYO"

Private Function B_CHK_Value(strOnAction As String) As Boolean
    Dim b As CheckBox
    Dim dlgMenu As DialogSheet
    Set dlgMenu = ThisWorkbook.DialogSheets("Menu1")
    Dim strOnAct As String
    B_CHK_Value = False
    
    For Each b In dlgMenu.CheckBoxes
        If b.Name Like "B_CHK*" Then
            strOnAct = b.OnAction
            If InStr(strOnAct, "!") Then
                strOnAct = Split(strOnAct, "!")(1)
            End If
            If strOnAct = strOnAction Then
                If b.Value = 1 Then
                    B_CHK_Value = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++
'部署（センター）名
Sub CHK_BUSYO_INIT()
Attribute CHK_BUSYO_INIT.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim i As Integer
    Dim b As CheckBox
    Dim lngRow As Long
    Dim wst As Worksheet
    Dim wstFile As Worksheet
    Dim dlgMenu As DialogSheet
    Dim strSheetNo As String
    Dim strBName As String
    Dim lngW As Double
    Dim lngH As Double
    Set dlgMenu = ThisWorkbook.DialogSheets("Menu1")
    Set wstFile = Workbooks(BookName(1)).Worksheets("FILE")
    lngW = 0
    lngH = 0
    For Each b In dlgMenu.CheckBoxes
        If b.Name Like "B_CHK*" Then
            If lngW = 0 Then
                lngW = b.Width
            End If
            If lngH = 0 Then
                lngH = b.Height
            End If
            b.Width = lngW
            b.Height = lngH
            Debug.Print b.Name & " " & b.Text & " " & b.Caption & " " & b.Value & " " & b.LinkedCell
'            b.Caption = b.Name
            b.Caption = ""
            b.LinkedCell = ""
            b.Value = False
            Select Case b.Name
            Case "B_CHK201", "B_CHK202", "B_CHK203", "B_CHK204"
                b.OnAction = "CHK_BUSYOIPPAN"
                Select Case b.Name
                Case "B_CHK201"
                    b.Caption = RTrim(wstFile.Range("J19"))
                Case "B_CHK202"
                    b.Caption = RTrim(wstFile.Range("K19"))
                Case "B_CHK203"
                    b.Caption = RTrim(wstFile.Range("L19"))
                Case "B_CHK204"
                    b.Caption = RTrim(wstFile.Range("M19"))
                End Select
            Case Else
                b.OnAction = "CHK_BUSYO"
            End Select
            '    .DrawingObjects(Array("B_CHK1", "B_CHK2", "B_CHK3", "B_CHK4", "B_CHK5", "B_CHK6", "B_CHK7", "B_CHK8", "B_CHK9", "B_CHK10", "B_CHK11", "B_CHK12", "B_CHK13", "B_CHK14", "B_CHK15", "B_CHK16", "B_CHK17", "B_CHK18", "B_CHK19", "B_CHK20", "B_CHK21", "B_CHK22", "B_CHK23", "B_CHK24", "B_CHK25")).OnAction = "CHK_BUSYO"
            '    .DrawingObjects(Array("B_CHK26", "B_CHK27", "B_CHK28", "B_CHK29")).OnAction = "CHK_BUSYOIPPAN"
        End If
    Next
    For lngRow = 10 To 55
        strSheetNo = RTrim(wstFile.Range("C" & lngRow))
        Debug.Print wstFile.Range("C" & lngRow)
        If strSheetNo <> "" Then
            If strSheetNo = "201" Then
                Debug.Print ""
            End If
            strBName = "B_CHK" & strSheetNo
            For Each b In dlgMenu.CheckBoxes
                If b.Name = strBName Then
                    b.Caption = RTrim(wstFile.Range("E" & lngRow))
                    If b.Caption <> "" Then
                        b.Value = True
                    End If
                    Exit For
                End If
            Next
        End If
    Next
End Sub
    
'Sub kkkk()
'  With ThisWorkbook
''    Dim dlgMenu As DialogSheet
''    Set dlgMenu = .DialogSheets("Menu1")
''    Debug.Print dlgMenu.Show
'    For i = 1 To 25
'' 2008.05.09
''      .Worksheets("LINK").Cells(i + 2, "C").Value = Workbooks(BookName(1)).Worksheets("FILE").Buttons("Sh_BT" & i).Text
'      .DialogSheets("Menu1").CheckBoxes("B_CHK" & i).Text = Workbooks(BookName(1)).Worksheets("FILE").Cells(i + 9, "E").Value
'      If Trim(.DialogSheets("Menu1").CheckBoxes("B_CHK" & i).Text) <> "" Then
'          .DialogSheets("Menu1").CheckBoxes("B_CHK" & i).Value = True
'      End If
'    Next
'    For i = 26 To 29
'' 2008.05.09
''      .Worksheets("LINK").Cells(i + 2, "C").Value = Workbooks(BookName(1)).Worksheets("FILE").Buttons("Sh_BT" & i).Text
'      .DialogSheets("Menu1").CheckBoxes("B_CHK" & (i)).Text = Workbooks(BookName(1)).Worksheets("FILE").Cells(19, i - 16).Value
''      If Trim(.DialogSheets("Menu1").CheckBoxes("B_CHK" & i).Text) <> "" Then
''          .DialogSheets("Menu1").CheckBoxes("B_CHK" & i).Value = True
''      End If
'    Next
'    Dim w As Worksheet
'    .DialogSheets("Menu1").CheckBoxes("B_CHK241").Text = ""
'    For Each w In Workbooks(BookName(1)).Worksheets
'        If w.Name = "641" Then
'            .DialogSheets("Menu1").CheckBoxes("B_CHK241").Text = "自販機家賃"
'        End If
'    Next
'  End With
'End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'処理方法
Sub OPT_SYORI()
Attribute OPT_SYORI.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook
    If .Worksheets("LINK").Range("E4").Value = 2 Then
      .DialogSheets("Menu1").GroupBoxes("H_GRP1").Enabled = True
      .DialogSheets("Menu1").OptionButtons("H_OPT1").Enabled = True
      .DialogSheets("Menu1").OptionButtons("H_OPT2").Enabled = True
      .DialogSheets("Menu1").OptionButtons("H_OPT3").Enabled = True
      .DialogSheets("Menu1").OptionButtons("H_OPT4").Enabled = True
      .DialogSheets("Menu1").OptionButtons("H_OPT5").Enabled = True
    Else
      .DialogSheets("Menu1").GroupBoxes("H_GRP1").Enabled = False
      .DialogSheets("Menu1").OptionButtons("H_OPT1").Enabled = False
      .DialogSheets("Menu1").OptionButtons("H_OPT2").Enabled = False
      .DialogSheets("Menu1").OptionButtons("H_OPT3").Enabled = False
      .DialogSheets("Menu1").OptionButtons("H_OPT4").Enabled = False
      .DialogSheets("Menu1").OptionButtons("H_OPT5").Enabled = False
      .DialogSheets("Menu1").EditBoxes("H_EDT1").Enabled = False
    End If
    Frag_Check
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'表題選択
Sub OPT_HYODAI()
Attribute OPT_HYODAI.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook
    If .Worksheets("LINK").Range("F4").Value = 5 Then
      .DialogSheets("Menu1").EditBoxes("H_EDT1").Enabled = True
      .DialogSheets("Menu1").EditBoxes("H_EDT1").Text = "（表題入力）"
      .DialogSheets("Menu1").Focus = "H_EDT1"
    Else
      .DialogSheets("Menu1").EditBoxes("H_EDT1").Enabled = False
      .DialogSheets("Menu1").EditBoxes("H_EDT1").Text = ""
    End If
    Frag_Check
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'表題その他
Sub EDT_HYODAI()
Attribute EDT_HYODAI.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook
    .Worksheets("LINK").Range("F9").Value = .DialogSheets("MENU1").EditBoxes("H_EDT1").Text
    Frag_Check
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'発行部数
Sub SPN_BUSU()
Attribute SPN_BUSU.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook
    .DialogSheets("Menu1").Labels("B_LAB").Text = .Worksheets("LINK").Range("G4").Value
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'発行書類1（通期一覧表）
Sub CHK_PRINT1()
Attribute CHK_PRINT1.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim Ch As Boolean
  
  With ThisWorkbook
    Ch = Application.Or(.Worksheets("LINK").Range("I4:I8"))
    .DialogSheets("Menu1").CheckBoxes("P1_P1_CHK").Enabled = Ch
    .DialogSheets("Menu1").CheckBoxes("P1_P2_CHK").Enabled = Ch
    .DialogSheets("Menu1").CheckBoxes("P1_P3_CHK").Enabled = Ch
    Frag_Check
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'発行書類2（収支対比表）
Sub CHK_PRINT2()
Attribute CHK_PRINT2.VB_ProcData.VB_Invoke_Func = " \n14"
  Frag_Check
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'発行書類3（一般販売）
Sub CHK_PRINT3()
Attribute CHK_PRINT3.VB_ProcData.VB_Invoke_Func = " \n14"
  Frag_Check
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'発行書類4（分析表）
Sub CHK_PRINT4()
Attribute CHK_PRINT4.VB_ProcData.VB_Invoke_Func = " \n14"
  Frag_Check
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'発行書類5（計画対比表）
Sub CHK_PRINT5()
Attribute CHK_PRINT5.VB_ProcData.VB_Invoke_Func = " \n14"
  Frag_Check
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'発行のチェック
Sub Frag_Check()
Attribute Frag_Check.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim Ch As Boolean
  Dim i As Integer
  
  With ThisWorkbook
    With .Worksheets("LINK")
'      If Application.Or(.Range("D3:D27").Value, .Range("D32").Value) Then '部署（センター）名
       If B_CHK_Value("CHK_BUSYO") = True Then
        Ch = .Range("J2").Value               '発行書類
        Ch = Ch And (.Range("B4").Value > 0)  '処理月
        Ch = Ch And (.Range("E4").Value > 0)  '処理方法
        If .Range("E4").Value = 2 Then Ch = Ch And (.Range("F4").Value > 0)
      End If
'      If Application.Or(.Range("D28:D31")) Then Ch = .Range("M3").Value
      If B_CHK_Value("CHK_BUSYOIPPAN") Then Ch = .Range("M3").Value
    End With
    
    .DialogSheets("Menu1").Buttons("発行").Enabled = Ch
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'シートチェック
Sub SHEET_Check()
Attribute SHEET_Check.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim i As Integer
  
  With ThisWorkbook
    For i = 1 To 29
'      .DialogSheets("Menu1").CheckBoxes("B_CHK" & i).Enabled = .Worksheets("管理").Cells(i + 5, "C").Value
    Next
'    .DialogSheets("Menu1").CheckBoxes("B_CHK241").Enabled = True
'    .DialogSheets("Menu1").CheckBoxes("B_CHK241").Value = False
  End With
End Sub

Sub x()
    Dim dlgMenu1 As DialogSheet
    Set dlgMenu1 = ThisWorkbook.DialogSheets("Menu1")
'    dlgMenu1.Activate
    dlgMenu1.Visible = True
    ThisWorkbook.IsAddin = False
'    dlgMenu1.Show
End Sub
