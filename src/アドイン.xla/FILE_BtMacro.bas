Attribute VB_Name = "FILE_BtMacro"
Option Explicit

'    dlgS.DialogFrame.Caption = "書類の発行 ･･･ 2008.04.28 新項目対応＆式補正対応＆シート保護解除"
'    dlgS.DialogFrame.Caption = "書類の発行 ･･･ 2008.04.30 ディスク作成の対応"
'    dlgS.DialogFrame.Caption = "書類の発行 ･･･ 2008.05.01 200販売管理の式対応"
'    dlgS.DialogFrame.Caption = "書類の発行 ･･･ 2008.06.05 収支対比表 累計 売上合計/原価その他 の不具合修正"
'    dlgS.DialogFrame.Caption = "書類の発行 ･･･ 2008.07.07 SHIKI SYUSEI"
' Const cstVersion = "2009.02.07 直接／間接人件費表示対応(大阪)"
' Const cstVersion = "ver.007 2009.02.07 直接／間接人件費 年間合計式追加"
' Const cstVersion = "ver.007o 2009.02.07 直接／間接人件費 年間合計式追加(小野版)"
' Const cstVersion = "ver.008 2009.08.05 粗利追加／共通化"
' Const cstVersion = "ver.009 2009.08.06 収支対比表 粗利 追加"
'Const cstVersion = "ver.010 2012.03.27 集計表 来期計画の場合 期をプラス１"
Const cstVersion = "ver.011 2020.03.07 2020年度版 売上原価 B 小計 の式を戻るボタンでセットしない"

'++++++++++++++++++++++++++++++++++++++++++++++++++
'ｼｰﾄ名ﾎﾞﾀﾝのアクション
Sub SH_BT()
Attribute SH_BT.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim Sh As String
  
'  If SystemSecure(0) = False Then
'    sysAllClose (0)
'    Application.Quit
'    Exit Sub
'  Else
    With Workbooks(BookName(1))
      Sh = .Worksheets("FILE").Buttons(Application.Caller).Text
      .Worksheets(Sh).Activate
      .Worksheets(Sh).Unprotect Password:="sdc2035"
' 2008.04.26
'      If InStr(.Worksheets(Sh).Range("AL2"), "商品化") > 0 Then
'        .Worksheets(Sh).Range("C10") = "販売(商品化)"
'        .Worksheets(Sh).Range("C11") = "工料(商品化)"
'        .Worksheets(Sh).Range("C17") = "資材仕入(商品化)"
'        .Worksheets(Sh).Range("C18") = "工料仕入(商品化)"
'      ElseIf InStr(.Worksheets(Sh).Range("AL2"), "販売") > 0 Then
'        .Worksheets(Sh).Range("C10") = "販売(資材)"
'        .Worksheets(Sh).Range("C17") = "資材仕入(販売)"
'        .Worksheets(Sh).Range("C17") = "工料仕入(商品化)"
'      ElseIf InStr(.Worksheets(Sh).Range("AL2"), "出荷") > 0 Then
'        .Worksheets(Sh).Range("C11") = "工料(出荷)"
'        .Worksheets(Sh).Range("C18") = "工料仕入(出荷)"
'      End If
      .Worksheets(Sh).Range("C17").HorizontalAlignment = xlCenter
      .Worksheets(Sh).Range("C17").VerticalAlignment = xlCenter
      .Worksheets(Sh).Range("C17").ShrinkToFit = True
      .Worksheets(Sh).Range("C18").HorizontalAlignment = xlCenter
      .Worksheets(Sh).Range("C18").VerticalAlignment = xlCenter
      .Worksheets(Sh).Range("C18").ShrinkToFit = True
'      .Worksheets(Sh).Protect password:="sdc2035"
    End With
'  End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'戻るﾎﾞﾀﾝのアクション
Sub SH_FILE()
Attribute SH_FILE.VB_ProcData.VB_Invoke_Func = " \n14"
  If SystemSecure(0) = False Then
    sysAllClose (0)
    Application.Quit
    Exit Sub
  Else
    SIKI_SetUp ActiveSheet.Name
    Workbooks(BookName(1)).Worksheets("FILE").Activate
  End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'終了
Sub END_BT()
Attribute END_BT.VB_ProcData.VB_Invoke_Func = " \n14"
  If SystemSecure(0) = False Then
    sysAllClose (0)
    Application.Quit
    Exit Sub
  Else
    Workbooks(BookName(1)).RunAutoMacros xlAutoClose
  End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'発行
Sub PRINT_BT()
Attribute PRINT_BT.VB_ProcData.VB_Invoke_Func = " \n14"
  If SystemSecure(0) = False Then
    sysAllClose (0)
    Application.Quit
    Exit Sub
  Else
    PRINT_BTsub
    Workbooks.Open DirName & BookName(3)
    Menu1_Init
    Dim dlgS As DialogSheet
    Set dlgS = ThisWorkbook.DialogSheets("Menu1")
'    dlgS.DialogFrame.Caption = "書類の発行 ･･･ 2008.04.28 新項目対応＆式補正対応＆シート保護解除"
'    dlgS.DialogFrame.Caption = "書類の発行 ･･･ 2008.04.30 ディスク作成の対応"
'    dlgS.DialogFrame.Caption = "書類の発行 ･･･ 2008.05.01 200販売管理の式対応"
    dlgS.DialogFrame.Caption = "書類の発行 ･･･ " & cstVersion
    If ThisWorkbook.DialogSheets("Menu1").Show Then
      PrintOut_Action
    End If
    Workbooks(BookName(3)).Close savechanges:=False
    'Workbooks(BookName(3)).Close savechanges:=True  '袋井用
    Workbooks(BookName(1)).Worksheets("FILE").Activate
  End If
End Sub
Sub PRINT_BTsub()
  With ThisWorkbook.DialogSheets("Menu1")
    .DrawingObjects(Array("OPTm4", "OPTm5", "OPTm6", "OPTm7", "OPTm8", "OPTm9", "OPTm10", "OPTm11", "OPTm12", "OPTm1", "OPTm2", "OPTm3")).OnAction = "OPT_TUKI"
    .DrawingObjects(Array("OPT_As", "OPT_At")).OnAction = "OPT_SYORI"
    .DrawingObjects(Array("H_OPT1", "H_OPT2", "H_OPT3", "H_OPT4", "H_OPT5")).OnAction = "OPT_HYODAI"
    .DrawingObjects(Array("H_EDT1")).OnAction = "EDT_HYODAI"
    .DrawingObjects(Array("P1_1_CHK", "P1_2_CHK", "P1_3_CHK", "P1_4_CHK", "P1_5_CHK", "P1_P1_CHK", "P1_P2_CHK", "P1_P3_CHK")).OnAction = "CHK_PRINT1"
    .DrawingObjects(Array("P2_1_CHK", "P2_2_CHK", "P2_3_CHK")).OnAction = "CHK_PRINT2"
    .DrawingObjects(Array("P3_1_CHK", "P3_2_CHK", "P3_3_CHK", "P3_4_CHK", "P3_5_CHK")).OnAction = "CHK_PRINT3"
    .DrawingObjects(Array("P4_1_CHK", "P4_2_CHK")).OnAction = "CHK_PRINT4"
    .DrawingObjects(Array("P5_1_CHK", "P5_2_CHK")).OnAction = "CHK_PRINT5"
'    .DrawingObjects(Array("B_CHK1", "B_CHK2", "B_CHK3", "B_CHK4", "B_CHK5", "B_CHK6", "B_CHK7", "B_CHK8", "B_CHK9", "B_CHK10", "B_CHK11", "B_CHK12", "B_CHK13", "B_CHK14", "B_CHK15", "B_CHK16", "B_CHK17", "B_CHK18", "B_CHK19", "B_CHK20", "B_CHK21", "B_CHK22", "B_CHK23", "B_CHK24", "B_CHK25")).OnAction = "CHK_BUSYO"
'    .DrawingObjects(Array("B_CHK26", "B_CHK27", "B_CHK28", "B_CHK29")).OnAction = "CHK_BUSYOIPPAN"
    .DrawingObjects(Array("SPN_ct")).OnAction = "SPN_BUSU"
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'ﾃﾞｨｽｸ作成
Sub DOWN_BT()
Attribute DOWN_BT.VB_ProcData.VB_Invoke_Func = " \n14"
  If SystemSecure(0) = False Then
    sysAllClose (0)
    Application.Quit
    Exit Sub
  Else
    Make_Disk
    MsgBox "できあがり。"
  End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'メンテ
Sub MENTE_BT()
Attribute MENTE_BT.VB_ProcData.VB_Invoke_Func = " \n14"
  If SystemSecure(0) = False Then
    sysAllClose (0)
    Application.Quit
    Exit Sub
  Else
    MENTE_Year
  End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'式の書き換え
Sub SIKI_SetUp(ByVal WriShName As String)
Attribute SIKI_SetUp.VB_ProcData.VB_Invoke_Func = " \n14"
  With Workbooks(BookName(1))
    .Sheets(WriShName).Unprotect Password:="sdc2035"
    Select Case WriShName
      Case "200"
        WriteSIKI_200
        WriteSIKI .Sheets(WriShName)
      Case "000"
        WriteSIKI_000
        WriteSIKI .Sheets(WriShName)
      Case "201", "202", "203", "204"
        WriteSIKI_20X .Sheets(WriShName)
      Case "P1", "P2", "P3", "P4"
      
      Case Else
        WriteSIKI .Sheets(WriShName)
    End Select
    .Sheets(WriShName).Protect Password:="sdc2035"
  End With
End Sub

'+++++++++++++++++++++++++
'式の入力（シート000の式）
Sub WriteSIKI_000()
Attribute WriteSIKI_000.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim lngRowPay1 As Long
    Dim lngRowPay2 As Long
  
  With Workbooks(BookName(1)).Worksheets("000")
    .Range("AA10:AL37,BA10:BL37,CA10:CL37,DA10:DL37,EA10:EL37").FormulaR1C1 = "=SUM('P2:P3'!RC)"
    .Range("AA50:AL88,BA50:BL88,CA50:CL88,DA50:DL88,EA50:EL88").FormulaR1C1 = "=SUM('P2:P3'!RC)"
        
    .Range("AA50:AL88,BA50:BL88,CA50:CL88,DA50:DL88,EA50:EL88").FormulaR1C1 = "=SUM('P2:P3'!RC)"
    ' 2009.02.07 直接／間接人件費
    Select Case Workbooks(BookName(1)).Worksheets("FILE").Range("K2")
    Case "草津商品化センター"
        lngRowPay1 = 177        ' ②直接人件費
        lngRowPay2 = 178        ' ③間接人件費
    Case "小野PC"
        lngRowPay1 = 187        ' ②直接人件費
        lngRowPay2 = 188        ' ③間接人件費
    Case "大阪ＰＣ"
        lngRowPay1 = 148        ' ②直接人件費
        lngRowPay2 = 149        ' ③間接人件費
    Case "滋賀ＰＣ"
        lngRowPay1 = 178        ' ②直接人件費
        lngRowPay2 = 179        ' ③間接人件費
    Case "袋井PC"
        lngRowPay1 = 176        ' ②直接人件費
        lngRowPay2 = 177        ' ③間接人件費
    Case "奈良営業所"
        lngRowPay1 = 197        ' ②直接人件費
        lngRowPay2 = 198        ' ③間接人件費
    Case Else
        lngRowPay1 = 65530      ' ②直接人件費
        lngRowPay2 = 65531      ' ③間接人件費
    End Select
    If lngRowPay1 < 65530 Then
        .Range("C" & lngRowPay1) = "②直接人件費"
        .Range("C" & lngRowPay2) = "③間接人件費"
        .Range("AA" & lngRowPay1 & ":AL" & lngRowPay2).FormulaR1C1 = "=SUM('P2:P3'!RC)"
        .Range("BA" & lngRowPay1 & ":BL" & lngRowPay2).FormulaR1C1 = "=SUM('P2:P3'!RC)"
        .Range("CA" & lngRowPay1 & ":CL" & lngRowPay2).FormulaR1C1 = "=SUM('P2:P3'!RC)"
        .Range("DA" & lngRowPay1 & ":DL" & lngRowPay2).FormulaR1C1 = "=SUM('P2:P3'!RC)"
        .Range("EA" & lngRowPay1 & ":EL" & lngRowPay2).FormulaR1C1 = "=SUM('P2:P3'!RC)"
    End If
  End With
  
End Sub

'+++++++++++++++++++++++++
'式の入力（シート200の式）
Sub WriteSIKI_200()
Attribute WriteSIKI_200.VB_ProcData.VB_Invoke_Func = " \n14"
  ' 2008.04.26
  With Workbooks(BookName(1)).Worksheets("200")
    .Range("AA10:AL10,BA10:BL10,CA10:CL10,DA10:DL10,EA10:EL10").FormulaR1C1 = "=SUM('P3:P4'!R[91]C)"
    .Range("AA17:AL17,BA17:BL17,CA17:CL17,DA17:DL17,EA17:EL17").FormulaR1C1 = "=SUM('P3:P4'!R[85]C)"
    .Range("AA18:AL18,BA18:BL18,CA18:CL18,DA18:DL18,EA18:EL18").FormulaR1C1 = "=SUM('P3:P4'!R[85]C)"
    .Range("AA21:AL21,BA21:BL21,CA21:CL21,DA21:DL21,EA21:EL21").FormulaR1C1 = "=SUM('P3:P4'!R[83]C)"
  End With
End Sub

'+++++++++++++++++++++++++
'式の入力（小計・合計など）
Sub WriteSIKI(ByVal A_Sh As Object)
Attribute WriteSIKI.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim lngRowPay1 As Long
    Dim lngRowPay2 As Long
  ' 2008.04.26
  With A_Sh
    .Range("AA16:AL16,BA16:BL16,CA16:CL16,DA16:DL16,EA16:EL16").FormulaR1C1 = "=SUM(R[-6]C:R[-1]C)"
    .Range("AA22:AL22,BA22:BL22,CA22:CL22,DA22:DL22,EA22:EL22").FormulaR1C1 = "=SUM(R[-5]C:R[-2]C)-SUM(R[-1]C)"
    .Range("AA26:AL26,BA26:BL26,CA26:CL26,DA26:DL26,EA26:EL26").FormulaR1C1 = "=R[-3]C + R[-2]C - R[-1]C"
    .Range("AA27:AL27,BA27:BL27,CA27:CL27,DA27:DL27,EA27:EL27").FormulaR1C1 = "=SUM(R[-5]C,R[-1]C)"
    .Range("AA28:AL28,BA28:BL28,CA28:CL28,DA28:DL28,EA28:EL28").FormulaR1C1 = "=SUM(R[22]C,R[24]C,R[25]C,R[26]C,R[27]C,R[55]C,R[57]C)"
'    .Range("AA29:AL29,BA29:BL29,CA29:CL29,DA29:DL29,EA29:EL29").FormulaR1C1 = "=SUM(R[-17]C)-SUM(R[-7]C)"
'    .Range("AA30:AL30,BA30:BL30,CA30:CL30,DA30:DL30,EA30:EL30").FormulaR1C1 = "=SUM(R[-17]C)-SUM(R[-4]C)"
    .Range("AA32:AL32,BA32:BL32,CA32:CL32,DA32:DL32,EA32:EL32").FormulaR1C1 = "=R[59]C - R[1]C" ' 91-32
    .Range("AA34:AL34,BA34:BL34,CA34:CL34,DA34:DL34,EA34:EL34").FormulaR1C1 = "=SUM(R[-6]C:R[-1]C)"
'    .Range("AA32:AL32,BA32:BL32,CA32:CL32,DA32:DL32,EA32:EL32").FormulaR1C1 = "=SUM(R[58]C)"
'    .Range("AA33:AL33,BA33:BL33,CA33:CL33,DA33:DL33,EA33:EL33").FormulaR1C1 = "=SUM(R[56]C)-SUM(R[57]C)"
'    .Range("AA34:AL34,BA34:BL34,CA34:CL34,DA34:DL34,EA34:EL34").FormulaR1C1 = "=SUM(R[55]C)"
    .Range("AA35:AL35,BA35:BL35,CA35:CL35,DA35:DL35,EA35:EL35").FormulaR1C1 = "=R[-19]C-R[-8]C-R[-1]C"
    .Range("AA51:AL51,BA51:BL51,CA51:CL51,DA51:DL51,EA51:EL51").FormulaR1C1 = "=sum(R[-22]C:R[-20]C)" ' 29:31-51
    .Range("AA37:AL37,BA37:BL37,CA37:CL37,DA37:DL37,EA37:EL37").FormulaR1C1 = "=R[-2]C-R[-1]C"
    .Range("AA89:AL89,BA89:BL89,CA89:CL89,DA89:DL89,EA89:EL89").FormulaR1C1 = "=SUM(R[-39]C:R[-1]C)"
    .Range("AA90:AL90,BA90:BL90,CA90:CL90,DA90:DL90,EA90:EL90").FormulaR1C1 = "=SUM(R[-40]C:R[-35]C,R[-7]C,R[-5]C)"
    .Range("AA91:AL91,BA91:BL91,CA91:CL91,DA91:DL91,EA91:EL91").FormulaR1C1 = "=R[-2]C-R[-1]C"
    
    .Range("AM10:AM16,AM18:AM20,AM22:AM37,AM50:AM91").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("BM10:BM16,BM18:BM20,BM22:BM37,BM50:BM91").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("CM10:CM16,CM18:CM20,CM22:CM37,CM50:CM91").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("DM10:DM16,DM18:DM20,DM22:DM37,DM50:DM91").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("EM10:EM16,EM18:EM20,EM22:EM37,EM50:EM91").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("AM21,BM21,CM21,DM21,EM21").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])-SUM(R[-4]C[-12]:R[-4]C[-1])"
    
    .Range("E10:E37,E50:E88").FormulaR1C1 = "=SUM(RC[22]:RC[30],RC[57]:RC[59])"
    .Range("E21").FormulaR1C1 = "=SUM(RC[22]:RC[30],RC[57]:RC[59])-SUM(R[-4]C[22]:R[-4]C[30],R[-4]C[57]:R[-4]C[59])"
    
    ' 2009.02.07 直接／間接人件費
    Select Case Workbooks(BookName(1)).Worksheets("FILE").Range("K2")
    Case "草津商品化センター"
        lngRowPay1 = 177        ' ②直接人件費
        lngRowPay2 = 178        ' ③間接人件費
    Case "小野PC"
        lngRowPay1 = 187        ' ②直接人件費
        lngRowPay2 = 188        ' ③間接人件費
    Case "大阪ＰＣ"
        lngRowPay1 = 148        ' ②直接人件費
        lngRowPay2 = 149        ' ③間接人件費
    Case "滋賀ＰＣ"
        lngRowPay1 = 178        ' ②直接人件費
        lngRowPay2 = 179        ' ③間接人件費
    Case "袋井PC"
        lngRowPay1 = 176        ' ②直接人件費
        lngRowPay2 = 177        ' ③間接人件費
    Case "奈良営業所"
        lngRowPay1 = 197        ' ②直接人件費
        lngRowPay2 = 198        ' ③間接人件費
    Case Else
        lngRowPay1 = 65530      ' ②直接人件費
        lngRowPay2 = 65531      ' ③間接人件費
    End Select
    If lngRowPay1 < 65530 Then
        .Range("C" & lngRowPay1) = "②直接人件費"
        .Range("C" & lngRowPay2) = "③間接人件費"
        .Range("AM" & lngRowPay1 & ":AM" & lngRowPay2).FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
        .Range("BM" & lngRowPay1 & ":BM" & lngRowPay2).FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
        .Range("CM" & lngRowPay1 & ":CM" & lngRowPay2).FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
        .Range("DM" & lngRowPay1 & ":DM" & lngRowPay2).FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
        .Range("EM" & lngRowPay1 & ":EM" & lngRowPay2).FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
        
        .Range("AM" & lngRowPay1 & ":AM" & lngRowPay2).Interior.Pattern = xlSolid
        .Range("BM" & lngRowPay1 & ":BM" & lngRowPay2).Interior.Pattern = xlSolid
        .Range("CM" & lngRowPay1 & ":CM" & lngRowPay2).Interior.Pattern = xlSolid
        .Range("DM" & lngRowPay1 & ":DM" & lngRowPay2).Interior.Pattern = xlSolid
        .Range("EM" & lngRowPay1 & ":EM" & lngRowPay2).Interior.Pattern = xlSolid
        
        .Range("C" & lngRowPay1 & ":EM" & lngRowPay2).Interior.ColorIndex = 35
    
    End If
  End With
End Sub

'+++++++++++++++++++++++++
'式の入力（シート201,202,203,204の式）
Sub WriteSIKI_20X(ByVal A_Sh As Object)
Attribute WriteSIKI_20X.VB_ProcData.VB_Invoke_Func = " \n14"
  With A_Sh
    .Range("AA15:AL15,BA15:BL15,CA15:CL15,DA15:DL15,EA15:EL15").FormulaR1C1 = "=SUM(R[-3]C:R[-2]C)-SUM(R[-1]C)"
    .Range("AA25:AL25,BA25:BL25,CA25:CL25,DA25:DL25,EA25:EL25").FormulaR1C1 = "=SUM(R[-3]C:R[-2]C)-SUM(R[-1]C)"
    .Range("AA35:AL35,BA35:BL35,CA35:CL35,DA35:DL35,EA35:EL35").FormulaR1C1 = "=SUM(R[-3]C:R[-2]C)-SUM(R[-1]C)"
    .Range("AA45:AL45,BA45:BL45,CA45:CL45,DA45:DL45,EA45:EL45").FormulaR1C1 = "=SUM(R[-3]C:R[-2]C)-SUM(R[-1]C)"
    .Range("AA55:AL55,BA55:BL55,CA55:CL55,DA55:DL55,EA55:EL55").FormulaR1C1 = "=SUM(R[-3]C:R[-2]C)-SUM(R[-1]C)"
    .Range("AA65:AL65,BA65:BL65,CA65:CL65,DA65:DL65,EA65:EL65").FormulaR1C1 = "=SUM(R[-3]C:R[-2]C)-SUM(R[-1]C)"
    .Range("AA75:AL75,BA75:BL75,CA75:CL75,DA75:DL75,EA75:EL75").FormulaR1C1 = "=SUM(R[-3]C:R[-2]C)-SUM(R[-1]C)"
    .Range("AA85:AL85,BA85:BL85,CA85:CL85,DA85:DL85,EA85:EL85").FormulaR1C1 = "=SUM(R[-3]C:R[-2]C)-SUM(R[-1]C)"
    .Range("AA95:AL95,BA95:BL95,CA95:CL95,DA95:DL95,EA95:EL95").FormulaR1C1 = "=SUM(R[-3]C:R[-2]C)-SUM(R[-1]C)"
    .Range("AA105:AL105,BA105:BL105,CA105:CL105,DA105:DL105,EA105:EL105").FormulaR1C1 = "=SUM(R[-3]C:R[-2]C)-SUM(R[-1]C)"
    
    .Range("AA16:AL16,BA16:BL16,CA16:CL16,DA16:DL16,EA16:EL16").FormulaR1C1 = "=SUM(R[-5]C)-SUM(R[-1]C)"
    .Range("AA26:AL26,BA26:BL26,CA26:CL26,DA26:DL26,EA26:EL26").FormulaR1C1 = "=SUM(R[-5]C)-SUM(R[-1]C)"
    .Range("AA36:AL36,BA36:BL36,CA36:CL36,DA36:DL36,EA36:EL36").FormulaR1C1 = "=SUM(R[-5]C)-SUM(R[-1]C)"
    .Range("AA46:AL46,BA46:BL46,CA46:CL46,DA46:DL46,EA46:EL46").FormulaR1C1 = "=SUM(R[-5]C)-SUM(R[-1]C)"
    .Range("AA56:AL56,BA56:BL56,CA56:CL56,DA56:DL56,EA56:EL56").FormulaR1C1 = "=SUM(R[-5]C)-SUM(R[-1]C)"
    .Range("AA66:AL66,BA66:BL66,CA66:CL66,DA66:DL66,EA66:EL66").FormulaR1C1 = "=SUM(R[-5]C)-SUM(R[-1]C)"
    .Range("AA76:AL76,BA76:BL76,CA76:CL76,DA76:DL76,EA76:EL76").FormulaR1C1 = "=SUM(R[-5]C)-SUM(R[-1]C)"
    .Range("AA86:AL86,BA86:BL86,CA86:CL86,DA86:DL86,EA86:EL86").FormulaR1C1 = "=SUM(R[-5]C)-SUM(R[-1]C)"
    .Range("AA96:AL96,BA96:BL96,CA96:CL96,DA96:DL96,EA96:EL96").FormulaR1C1 = "=SUM(R[-5]C)-SUM(R[-1]C)"
    .Range("AA106:AL106,BA106:BL106,CA106:CL106,DA106:DL106,EA106:EL106").FormulaR1C1 = "=SUM(R[-5]C)-SUM(R[-1]C)"
    
    .Range("AA101:AL104,BA101:BL104,CA101:CL104,DA101:DL104,EA101:EL104").FormulaR1C1 = "=SUM(R[-90]C,R[-80]C,R[-70]C,R[-60]C,R[-50]C,R[-40]C,R[-30]C,R[-20]C,R[-10]C)"
    
    .Range("AM11,AM13,AM15:AM16,BM11,BM13,BM15:BM16,CM11,CM13,CM15:CM16,DM11,DM13,DM15:DM16,EM11,EM13,EM15:EM16").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("AM21,AM23,AM25:AM26,BM21,BM23,BM25:BM26,CM21,CM23,CM25:CM26,DM21,DM23,DM25:DM26,EM21,EM23,EM25:EM26").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("AM31,AM33,AM35:AM36,BM31,BM33,BM35:BM36,CM31,CM33,CM35:CM36,DM31,DM33,DM35:DM36,EM31,EM33,EM35:EM36").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("AM41,AM43,AM45:AM46,BM41,BM43,BM45:BM46,CM41,CM43,CM45:CM46,DM41,DM43,DM45:DM46,EM41,EM43,EM45:EM46").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("AM51,AM53,AM55:AM56,BM51,BM53,BM55:BM56,CM51,CM53,CM55:CM56,DM51,DM53,DM55:DM56,EM51,EM53,EM55:EM56").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("AM61,AM63,AM65:AM66,BM61,BM63,BM65:BM66,CM61,CM63,CM65:CM66,DM61,DM63,DM65:DM66,EM61,EM63,EM65:EM66").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("AM71,AM73,AM75:AM76,BM71,BM73,BM75:BM76,CM71,CM73,CM75:CM76,DM71,DM73,DM75:DM76,EM71,EM73,EM75:EM76").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("AM81,AM83,AM85:AM86,BM81,BM83,BM85:BM86,CM81,CM83,CM85:CM86,DM81,DM83,DM85:DM86,EM81,EM83,EM85:EM86").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("AM91,AM93,AM95:AM96,BM91,BM93,BM95:BM96,CM91,CM93,CM95:CM96,DM91,DM93,DM95:DM96,EM91,EM93,EM95:EM96").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Range("AM101,AM103,AM105:AM106,BM101,BM103,BM105:BM106,CM101,CM103,CM105:CM106,DM101,DM103,DM105:DM106,EM101,EM103,EM105:EM106").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    
    .Range("AM14,BM14,CM14,DM14,EM14,AM24,BM24,CM24,DM24,EM24,AM34,BM34,CM34,DM34,EM34,AM44,BM44,CM44,DM44,EM44,AM54,BM54,CM54,DM54,EM54").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])-SUM(R[-2]C[-12]:R[-2]C[-1])"
    .Range("AM64,BM64,CM64,DM64,EM64,AM74,BM74,CM74,DM74,EM74,AM84,BM84,CM84,DM84,EM84,AM94,BM94,CM94,DM94,EM94,AM104,BM104,CM104,DM104,EM104").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])-SUM(R[-2]C[-12]:R[-2]C[-1])"
    
    .Range("E11,E13,E15:E16,E21,E23,E25:E26,E31,E33,E35:E36,E41,E43,E45:E46,E51,E53,E55:E56").FormulaR1C1 = "=SUM(RC[22]:RC[30],RC[57],RC[59])"
    .Range("E61,E63,E65:E66,E71,E73,E75:E76,E81,E83,E85:E86,E91,E93,E95:E96,E101,E103,E105:E106").FormulaR1C1 = "=SUM(RC[22]:RC[30],RC[57],RC[59])"
    .Range("E14,E24,E34,E44,E54,E64,E74,E84,E94,E104").FormulaR1C1 = "=SUM(RC[22]:RC[30],RC[57]:RC[59])-SUM(R[-2]C[22]:R[-2]C[30],R[-2]C[57]:R[-2]C[59])"
  End With
End Sub

