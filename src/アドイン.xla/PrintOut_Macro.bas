Attribute VB_Name = "PrintOut_Macro"
Option Explicit
Option Base 1
' 001 2009.02.07 ���ځ^�Ԑڐl����Ή�(����SC)
' 002 2009.02.07 ���ځ^�Ԑڐl����Ή�(���PC)
' 003 2009.02.07 ���ځ^�Ԑڐl����Ή�(����PC)
' 004 2009.02.07 ���ځ^�Ԑڐl����Ή�(�܈�PC) ' �l����̍s���s����
' 005 2009.02.07 ���ځ^�Ԑڐl����Ή�(�ޗ�)
' 006 2009.02.07 ���ځ^�Ԑڐl����Ή� ��������ݒ�
' 006o 2009.02.07 ���ځ^�Ԑڐl����Ή� ��������ݒ�(�����)

'++++++++++++++++++++++++++++++++++++++++++++++++++
'���s����
'++++++++++++++++++++++++++++++++++++++++++++++++++
Dim ActSHnum, TUKI As Integer
Dim ActSHname(), SBname() As String

'++++++++++++++++++++++++++++++++++++++++++++++++++
'���s����
Sub PrintOut_Action()
Attribute PrintOut_Action.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim i As Integer, Ct As Integer
  Dim b As CheckBox
  Dim dlgMenu As DialogSheet
  Dim strOnAct As String

  Set dlgMenu = ThisWorkbook.DialogSheets("Menu1")
  Ct = 0
  With ThisWorkbook.Worksheets("LINK")
    For Each b In dlgMenu.CheckBoxes
        If b.Name Like "B_CHK*" Then
            strOnAct = b.OnAction
            If InStr(strOnAct, "!") Then
                strOnAct = Split(strOnAct, "!")(1)
            End If
            If strOnAct = "CHK_BUSYO" Then
                If b.Value = 1 Then
                    Ct = Ct + 1                               '�I�����ꂽ�V�[�g�����擾
                    ReDim Preserve ActSHname(Ct)
                    ReDim Preserve SBname(Ct)
                    ActSHname(Ct) = Replace(b.Name, "B_CHK", "") '�uLINK�v�V�[�g�ɏ����Ă���V�[�g�����擾
                    SBname(Ct) = "[" & BookName(1) & "]" & ActSHname(Ct) & "!"
                End If
            End If
        End If
    Next
    
'    For i = 1 To 25
'      If .Cells(i + 2, "D").Value = True Then
'        Ct = Ct + 1                               '�I�����ꂽ�V�[�g�����擾
'        ReDim Preserve ActSHname(Ct)
'        ReDim Preserve SBname(Ct)
'        ActSHname(Ct) = .Cells(i + 2, "C").Value  '�uLINK�v�V�[�g�ɏ����Ă���V�[�g�����擾
'        SBname(Ct) = "[" & BookName(1) & "]" & ActSHname(Ct) & "!"
'      End If
'    Next
'    If .Cells(32, "D").Value = True Then
'      Ct = Ct + 1                               '�I�����ꂽ�V�[�g�����擾
'      ReDim Preserve ActSHname(Ct)
'      ReDim Preserve SBname(Ct)
'      ActSHname(Ct) = .Cells(32, "C").Value  '�uLINK�v�V�[�g�ɏ����Ă���V�[�g�����擾
'      SBname(Ct) = "[" & BookName(1) & "]" & ActSHname(Ct) & "!"
'    End If
    TUKI = .Range("B4").Value   '������No.
    ActSHnum = Ct
    
    '�ʊ��ꗗ�\�E���x�Δ�\
    If .Range("I3").Value Or .Range("K3").Value Then
      Make_SHT  '�W�v�\�E���x�Δ�\�̍쐻
    End If
    
    '���͕\�i��c�p�j
    If .Range("O3").Value Then
      Make_Sh4  '���͕\�i��c�p�j�̍쐻
    End If
    
    '��ʂ̏W�v
    If .Range("M3").Value Then
      Make_Sh3  '��ʂ̏W�v�̍쐻
    End If
    
    '�v��Δ�
    If .Range("Q3").Value Then
      Make_Sh5  '���ƌv����x�Δ�\�̍쐻
    End If
    
  End With
  
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'���ލ쐬
'++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++
'�W�v�\�E���x�Δ�\
Sub Make_SHT()
Attribute Make_SHT.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim i As Integer
    Dim TS1() As String, TS2() As String, TS3() As String, TS4() As String, TS5() As String 'S:���x���� 1:���� 2:��/�v�� 3:���ƌv�� 4:�O������ 5:�v��Δ�
    Dim TT1() As String, TT2() As String, TT3() As String, TT4() As String, TT5() As String 'S:���x���� 1:���� 2:��/�v�� 3:���ƌv�� 4:�O������ 5:�v��Δ�
    Dim TK1() As String, TK2() As String, TK3() As String, TK4() As String, TK5() As String 'K:�o��� 1:���� 2:��/�v�� 3:���ƌv�� 4:�O������ 5:�v��Δ�
    Dim TZ1() As String, TZ2() As String, TZ3() As String, TZ4() As String, TZ5() As String 'K:���ځ^�Ԑڐl���� 1:���� 2:��/�v�� 3:���ƌv�� 4:�O������ 5:�v��Δ�
    Dim INPbk As Object, PRTbk As Object
    Dim lngRowPay1 As Long
    Dim lngRowPay2 As Long
  
  Set INPbk = Workbooks(BookName(1))
  Set PRTbk = Workbooks(BookName(3))
  
  PRTbk.Worksheets("�W�v�\").Range("B2").Value = "�� " & INPbk.Worksheets("FILE").Range("C2").Value & " ��"
'  PRTbk.Worksheets("�W�v�\(�����)").Range("B2").Value = "�� " & INPbk.Worksheets("FILE").Range("C2").Value & " ��"
  PRTbk.Worksheets("�Δ�\").Range("B2").Value = "�� " & INPbk.Worksheets("FILE").Range("C2").Value & " ��"
  PRTbk.Worksheets("�Δ�\").Range("AA2").Value = ThisWorkbook.Worksheets("LINK").Range("B3").Value
'  PRTbk.Worksheets("�Δ�\(�����)").Range("B2").Value = "�� " & INPbk.Worksheets("FILE").Range("C2").Value & " ��"
'  PRTbk.Worksheets("�Δ�\(�����)").Range("AA2").Value = ThisWorkbook.Worksheets("LINK").Range("B3").Value
  
  With PRTbk
    Select Case INPbk.Worksheets("FILE").Range("K2")
    Case "���Ï��i���Z���^�["
        lngRowPay1 = 177        ' �A���ڐl����
        lngRowPay2 = 178        ' �B�Ԑڐl����
    Case "����PC"
        lngRowPay1 = 187        ' �A���ڐl����
        lngRowPay2 = 188        ' �B�Ԑڐl����
    Case "���o�b"
        lngRowPay1 = 148        ' �A���ڐl����
        lngRowPay2 = 149        ' �B�Ԑڐl����
    Case "����o�b"
        lngRowPay1 = 178        ' �A���ڐl����
        lngRowPay2 = 179        ' �B�Ԑڐl����
    Case "�܈�PC"
        lngRowPay1 = 176        ' �A���ڐl����
        lngRowPay2 = 177        ' �B�Ԑڐl����
    Case "�ޗǉc�Ə�"
        lngRowPay1 = 197        ' �A���ڐl����
        lngRowPay2 = 198        ' �B�Ԑڐl����
    Case Else
        lngRowPay1 = 65530      ' �A���ڐl����
        lngRowPay2 = 65531      ' �B�Ԑڐl����
    End Select
    .Worksheets("�W�v�\").Range("C38") = "�A���ڐl����"
    .Worksheets("�W�v�\").Range("C39") = "�B�Ԑڐl����"
    .Worksheets("�W�v�\").Range("38:39").EntireRow.RowHeight = .Worksheets("�W�v�\").Range("37:37").EntireRow.Height
    .Worksheets("�W�v�\").Range("38:39").VerticalAlignment = xlCenter
    
    .Worksheets("�W�v�\").Range("AA38:EM39").Font.Name = .Worksheets("�W�v�\").Range("AA37").Font.Name
    .Worksheets("�W�v�\").Range("AA38:EM39").NumberFormatLocal = .Worksheets("�W�v�\").Range("AA37").NumberFormatLocal
    .Worksheets("�W�v�\").Range("AA38:EM39").Font.Size = .Worksheets("�W�v�\").Range("AA37").Font.Size
    
    Select Case ThisWorkbook.Worksheets("LINK").Range("E4").Value  '1:�����P�� ,2:�W�v
      Case 1  '������������������������������������������������������ �����P��
        ReDim TS1(1), TS2(1), TS3(1), TS4(1), TS5(1)
        ReDim TK1(1), TK2(1), TK3(1), TK4(1), TK5(1)
        ReDim TZ1(1), TZ2(1), TZ3(1), TZ4(1), TZ5(1)
        
        For i = 1 To ActSHnum
          '�W�v�\
          TS1(1) = SBname(i) & "R10C27:R37C39"   '[AA10:AM37] ���x����
          TS2(1) = SBname(i) & "R10C53:R37C65"   '[BA10:BM37] ���x����
          TS3(1) = SBname(i) & "R10C79:R37C91"   '[CA10:CM37] ���x����
          TS4(1) = SBname(i) & "R10C105:R37C117" '[DA10:DM37] ���x����
          TS5(1) = SBname(i) & "R10C131:R37C143" '[EA10:EM37] ���x����
          
          TK1(1) = SBname(i) & "R50C27:R91C39"   '[AA50:AM90] �o���
          TK2(1) = SBname(i) & "R50C53:R91C65"   '[BA50:BM90] �o���
          TK3(1) = SBname(i) & "R50C79:R91C91"   '[CA50:CM90] �o���
          TK4(1) = SBname(i) & "R50C105:R91C117" '[DA50:DM90] �o���
          TK5(1) = SBname(i) & "R50C131:R91C143" '[EA50:EM90] �o���
          
          TZ1(1) = SBname(i) & "R" & lngRowPay1 & "C27:R" & lngRowPay2 & "C39"   '[AA177:AM178]
          TZ2(1) = SBname(i) & "R" & lngRowPay1 & "C53:R" & lngRowPay2 & "C65"   '[BA177:BM178]
          TZ3(1) = SBname(i) & "R" & lngRowPay1 & "C79:R" & lngRowPay2 & "C91"   '[CA177:CM178]
          TZ4(1) = SBname(i) & "R" & lngRowPay1 & "C105:R" & lngRowPay2 & "C117" '[DA177:DM178]
          TZ5(1) = SBname(i) & "R" & lngRowPay1 & "C131:R" & lngRowPay2 & "C143" '[EA177:EM178]
          
          .Worksheets("�W�v�\").Range("AA10").Consolidate sources:=TS1, Function:=xlSum
          .Worksheets("�W�v�\").Range("BA10").Consolidate sources:=TS2, Function:=xlSum
          .Worksheets("�W�v�\").Range("CA10").Consolidate sources:=TS3, Function:=xlSum
          .Worksheets("�W�v�\").Range("DA10").Consolidate sources:=TS4, Function:=xlSum
          .Worksheets("�W�v�\").Range("EA10").Consolidate sources:=TS5, Function:=xlSum
          
          .Worksheets("�W�v�\").Range("AA50").Consolidate sources:=TK1, Function:=xlSum
          .Worksheets("�W�v�\").Range("BA50").Consolidate sources:=TK2, Function:=xlSum
          .Worksheets("�W�v�\").Range("CA50").Consolidate sources:=TK3, Function:=xlSum
          .Worksheets("�W�v�\").Range("DA50").Consolidate sources:=TK4, Function:=xlSum
          .Worksheets("�W�v�\").Range("EA50").Consolidate sources:=TK5, Function:=xlSum
          
          .Worksheets("�W�v�\").Range("AA38").Consolidate sources:=TZ1, Function:=xlSum
          .Worksheets("�W�v�\").Range("BA38").Consolidate sources:=TZ2, Function:=xlSum
          .Worksheets("�W�v�\").Range("CA38").Consolidate sources:=TZ3, Function:=xlSum
          .Worksheets("�W�v�\").Range("DA38").Consolidate sources:=TZ4, Function:=xlSum
          .Worksheets("�W�v�\").Range("EA38").Consolidate sources:=TZ5, Function:=xlSum
          

'          .Worksheets("�W�v�\(�����)").Range("AA10").Consolidate sources:=TS1, Function:=xlSum
'          .Worksheets("�W�v�\(�����)").Range("BA10").Consolidate sources:=TS2, Function:=xlSum
'          .Worksheets("�W�v�\(�����)").Range("CA10").Consolidate sources:=TS3, Function:=xlSum
'          .Worksheets("�W�v�\(�����)").Range("DA10").Consolidate sources:=TS4, Function:=xlSum
'          .Worksheets("�W�v�\(�����)").Range("EA10").Consolidate sources:=TS5, Function:=xlSum
          
          '�\��
          .Worksheets("�W�v�\").Range("AK2").Value = INPbk.Worksheets(ActSHname(i)).Range("AK2").Value
          .Worksheets("�W�v�\").Range("AL2").Value = INPbk.Worksheets(ActSHname(i)).Range("AL2").Value
'          .Worksheets("�W�v�\(�����)").Range("AK2").Value = INPbk.Worksheets(ActSHname(i)).Range("AK2").Value
'          .Worksheets("�W�v�\(�����)").Range("AL2").Value = INPbk.Worksheets(ActSHname(i)).Range("AL2").Value
                                        
          '�Δ�\
'          .Worksheets("�Δ�\").Range("AO2").Value = INPbk.Worksheets(ActSHname(i)).Range("AK2").Value
          .Worksheets("�Δ�\").Range("AR2").Value = INPbk.Worksheets(ActSHname(i)).Range("AK2").Value
'          .Worksheets("�Δ�\").Range("AP2").ClearContents
          .Worksheets("�Δ�\").Range("AS2").ClearContents
'          .Worksheets("�Δ�\").Range("AQ2").Value = INPbk.Worksheets(ActSHname(i)).Range("AL2").Value
          .Worksheets("�Δ�\").Range("AV2").Value = INPbk.Worksheets(ActSHname(i)).Range("AL2").Value
          Make_Sh2
          ' �����
          Call PrintOut_ShT(False)
        Next
      Case 2  '������������������������������������������������������ �W�v
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        ReDim TT1(ActSHnum), TT2(ActSHnum), TT3(ActSHnum), TT4(ActSHnum), TT5(ActSHnum)
        ReDim TK1(ActSHnum), TK2(ActSHnum), TK3(ActSHnum), TK4(ActSHnum), TK5(ActSHnum)
        ReDim TZ1(ActSHnum), TZ2(ActSHnum), TZ3(ActSHnum), TZ4(ActSHnum), TZ5(ActSHnum)

        For i = 1 To ActSHnum
          TS1(i) = SBname(i) & "R10C27:R27C39"   '[AA10:AM27]
          TS2(i) = SBname(i) & "R10C53:R27C65"   '[BA10:BM27]
          TS3(i) = SBname(i) & "R10C79:R27C91"   '[CA10:CM27]
          TS4(i) = SBname(i) & "R10C105:R27C117" '[DA10:DM27]
          TS5(i) = SBname(i) & "R10C131:R27C143" '[EA10:EM27]
          
          TT1(i) = SBname(i) & "R27C27:R37C39"   '[AA27:AM37]
          TT2(i) = SBname(i) & "R27C53:R37C65"   '[BA27:BM37]
          TT3(i) = SBname(i) & "R27C79:R37C91"   '[CA27:CM37]
          TT4(i) = SBname(i) & "R27C105:R37C117" '[DA27:DM37]
          TT5(i) = SBname(i) & "R27C131:R37C143" '[EA27:EM37]

          TK1(i) = SBname(i) & "R50C27:R91C39"   '[AA50:AM90]
          TK2(i) = SBname(i) & "R50C53:R91C65"   '[BA50:BM90]
          TK3(i) = SBname(i) & "R50C79:R91C91"   '[CA50:CM90]
          TK4(i) = SBname(i) & "R50C105:R91C117" '[DA50:DM90]
          TK5(i) = SBname(i) & "R50C131:R91C143" '[EA50:EM90]

          TZ1(i) = SBname(i) & "R" & lngRowPay1 & "C27:R" & lngRowPay2 & "C39"   '[AA177:AM178]
          TZ2(i) = SBname(i) & "R" & lngRowPay1 & "C53:R" & lngRowPay2 & "C65"   '[BA177:BM178]
          TZ3(i) = SBname(i) & "R" & lngRowPay1 & "C79:R" & lngRowPay2 & "C91"   '[CA177:CM178]
          TZ4(i) = SBname(i) & "R" & lngRowPay1 & "C105:R" & lngRowPay2 & "C117" '[DA177:DM178]
          TZ5(i) = SBname(i) & "R" & lngRowPay1 & "C131:R" & lngRowPay2 & "C143" '[EA177:EM178]

        Next
        
        .Worksheets("�W�v�\").Range("AA10").Consolidate sources:=TS1, Function:=xlSum
        .Worksheets("�W�v�\").Range("BA10").Consolidate sources:=TS2, Function:=xlSum
        .Worksheets("�W�v�\").Range("CA10").Consolidate sources:=TS3, Function:=xlSum
        .Worksheets("�W�v�\").Range("DA10").Consolidate sources:=TS4, Function:=xlSum
        .Worksheets("�W�v�\").Range("EA10").Consolidate sources:=TS5, Function:=xlSum
        .Worksheets("�W�v�\").Range("AA27").Consolidate sources:=TT1, Function:=xlSum
        .Worksheets("�W�v�\").Range("BA27").Consolidate sources:=TT2, Function:=xlSum
        .Worksheets("�W�v�\").Range("CA27").Consolidate sources:=TT3, Function:=xlSum
        .Worksheets("�W�v�\").Range("DA27").Consolidate sources:=TT4, Function:=xlSum
        .Worksheets("�W�v�\").Range("EA27").Consolidate sources:=TT5, Function:=xlSum
        .Worksheets("�W�v�\").Range("AA50").Consolidate sources:=TK1, Function:=xlSum
        .Worksheets("�W�v�\").Range("BA50").Consolidate sources:=TK2, Function:=xlSum
        .Worksheets("�W�v�\").Range("CA50").Consolidate sources:=TK3, Function:=xlSum
        .Worksheets("�W�v�\").Range("DA50").Consolidate sources:=TK4, Function:=xlSum
        .Worksheets("�W�v�\").Range("EA50").Consolidate sources:=TK5, Function:=xlSum
        
        
        .Worksheets("�W�v�\").Range("AA38").Consolidate sources:=TZ1, Function:=xlSum
        .Worksheets("�W�v�\").Range("BA38").Consolidate sources:=TZ2, Function:=xlSum
        .Worksheets("�W�v�\").Range("CA38").Consolidate sources:=TZ3, Function:=xlSum
        .Worksheets("�W�v�\").Range("DA38").Consolidate sources:=TZ4, Function:=xlSum
        .Worksheets("�W�v�\").Range("EA38").Consolidate sources:=TZ5, Function:=xlSum
        
        ' ������v�N���A
'        .Worksheets("�W�v�\(�����)").Range("AA10:EM17").ClearContents
        Dim iNum As Integer
        Dim aryTS1() As String, aryTS2() As String, aryTS3() As String, aryTS4() As String, aryTS5() As String 'S:���x���� 1:���� 2:��/�v�� 3:���ƌv�� 4:�O������ 5:�v��Δ�
        ' �̔�(����)
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        iNum = 0
        For i = 1 To ActSHnum
'            If Not Workbooks("System�o�c����.xls").Worksheets(ActSHname(i)).Range("AL2") Like "*���i��*" Then
                iNum = iNum + 1
                TS1(iNum) = SBname(i) & "R10C27:R10C39"   '[AA10:AM10]
                TS2(iNum) = SBname(i) & "R10C53:R10C65"   '[BA10:BM10]
                TS3(iNum) = SBname(i) & "R10C79:R10C91"   '[CA10:CM10]
                TS4(iNum) = SBname(i) & "R10C105:R10C117" '[DA10:DM10]
                TS5(iNum) = SBname(i) & "R10C131:R10C143" '[EA10:EM10]
'            End If
        Next
        If iNum > 0 Then
            ReDim aryTS1(iNum) As String, aryTS2(iNum) As String, aryTS3(iNum) As String, aryTS4(iNum) As String, aryTS5(iNum) As String
            For i = 1 To iNum
                aryTS1(i) = TS1(i)
                aryTS2(i) = TS2(i)
                aryTS3(i) = TS3(i)
                aryTS4(i) = TS4(i)
                aryTS5(i) = TS5(i)
            Next
'            .Worksheets("�W�v�\(�����)").Range("AA10").Consolidate sources:=aryTS1, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("BA10").Consolidate sources:=aryTS2, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("CA10").Consolidate sources:=aryTS3, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("DA10").Consolidate sources:=aryTS4, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("EA10").Consolidate sources:=aryTS5, Function:=xlSum
        End If
        ' �̔�(���i��)
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        iNum = 0
        For i = 1 To ActSHnum
'            If Workbooks("System�o�c����.xls").Worksheets(ActSHname(i)).Range("AL2") Like "*���i��*" Then
                iNum = iNum + 1
                TS1(iNum) = SBname(i) & "R10C27:R10C39"   '[AA10:AM10]
                TS2(iNum) = SBname(i) & "R10C53:R10C65"   '[BA10:BM10]
                TS3(iNum) = SBname(i) & "R10C79:R10C91"   '[CA10:CM10]
                TS4(iNum) = SBname(i) & "R10C105:R10C117" '[DA10:DM10]
                TS5(iNum) = SBname(i) & "R10C131:R10C143" '[EA10:EM10]
'            End If
        Next
        If iNum > 0 Then
            ReDim aryTS1(iNum) As String, aryTS2(iNum) As String, aryTS3(iNum) As String, aryTS4(iNum) As String, aryTS5(iNum) As String
            For i = 1 To iNum
                aryTS1(i) = TS1(i)
                aryTS2(i) = TS2(i)
                aryTS3(i) = TS3(i)
                aryTS4(i) = TS4(i)
                aryTS5(i) = TS5(i)
            Next
'            .Worksheets("�W�v�\(�����)").Range("AA11").Consolidate sources:=aryTS1, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("BA11").Consolidate sources:=aryTS2, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("CA11").Consolidate sources:=aryTS3, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("DA11").Consolidate sources:=aryTS4, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("EA11").Consolidate sources:=aryTS5, Function:=xlSum
        End If
        ' �H��(���i��)
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        iNum = 0
        For i = 1 To ActSHnum
'            If Not Workbooks("System�o�c����.xls").Worksheets(ActSHname(i)).Range("AL2") Like "*�o��*" Then
                iNum = iNum + 1
                TS1(iNum) = SBname(i) & "R11C27:R11C39"   '[AA11:AM11]
                TS2(iNum) = SBname(i) & "R11C53:R11C65"   '[BA11:BM11]
                TS3(iNum) = SBname(i) & "R11C79:R11C91"   '[CA11:CM11]
                TS4(iNum) = SBname(i) & "R11C105:R11C117" '[DA11:DM11]
                TS5(iNum) = SBname(i) & "R11C131:R11C143" '[EA11:EM11]
'            End If
        Next
        If iNum > 0 Then
            ReDim aryTS1(iNum) As String, aryTS2(iNum) As String, aryTS3(iNum) As String, aryTS4(iNum) As String, aryTS5(iNum) As String
            For i = 1 To iNum
                aryTS1(i) = TS1(i)
                aryTS2(i) = TS2(i)
                aryTS3(i) = TS3(i)
                aryTS4(i) = TS4(i)
                aryTS5(i) = TS5(i)
            Next
'            .Worksheets("�W�v�\(�����)").Range("AA12").Consolidate sources:=aryTS1, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("BA12").Consolidate sources:=aryTS2, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("CA12").Consolidate sources:=aryTS3, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("DA12").Consolidate sources:=aryTS4, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("EA12").Consolidate sources:=aryTS5, Function:=xlSum
        End If
        ' �H��(�o��)
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        iNum = 0
        For i = 1 To ActSHnum
'            If Workbooks("System�o�c����.xls").Worksheets(ActSHname(i)).Range("AL2") Like "*�o��*" Then
                iNum = iNum + 1
                TS1(iNum) = SBname(i) & "R11C27:R11C39"   '[AA11:AM11]
                TS2(iNum) = SBname(i) & "R11C53:R11C65"   '[BA11:BM11]
                TS3(iNum) = SBname(i) & "R11C79:R11C91"   '[CA11:CM11]
                TS4(iNum) = SBname(i) & "R11C105:R11C117" '[DA11:DM11]
                TS5(iNum) = SBname(i) & "R11C131:R11C143" '[EA11:EM11]
'            End If
        Next
        If iNum > 0 Then
            ReDim aryTS1(iNum) As String, aryTS2(iNum) As String, aryTS3(iNum) As String, aryTS4(iNum) As String, aryTS5(iNum) As String
            For i = 1 To iNum
                aryTS1(i) = TS1(i)
                aryTS2(i) = TS2(i)
                aryTS3(i) = TS3(i)
                aryTS4(i) = TS4(i)
                aryTS5(i) = TS5(i)
            Next
'            .Worksheets("�W�v�\(�����)").Range("AA13").Consolidate sources:=aryTS1, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("BA13").Consolidate sources:=aryTS2, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("CA13").Consolidate sources:=aryTS3, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("DA13").Consolidate sources:=aryTS4, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("EA13").Consolidate sources:=aryTS5, Function:=xlSum
        End If
        ' �h���̕�
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        For i = 1 To ActSHnum
          TS1(i) = SBname(i) & "R12C27:R12C39"   '[AA12:AM12]
          TS2(i) = SBname(i) & "R12C53:R12C65"   '[BA12:BM12]
          TS3(i) = SBname(i) & "R12C79:R12C91"   '[CA12:CM12]
          TS4(i) = SBname(i) & "R12C105:R12C117" '[DA12:DM12]
          TS5(i) = SBname(i) & "R12C131:R12C143" '[EA12:EM12]
        Next
'        .Worksheets("�W�v�\(�����)").Range("AA14").Consolidate sources:=TS1, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("BA14").Consolidate sources:=TS2, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("CA14").Consolidate sources:=TS3, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("DA14").Consolidate sources:=TS4, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("EA14").Consolidate sources:=TS5, Function:=xlSum
        ' �ƒ��̕�
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        For i = 1 To ActSHnum
          TS1(i) = SBname(i) & "R13C27:R13C39"   '[AA13:AM13]
          TS2(i) = SBname(i) & "R13C53:R13C65"   '[BA13:BM13]
          TS3(i) = SBname(i) & "R13C79:R13C91"   '[CA13:CM13]
          TS4(i) = SBname(i) & "R13C105:R13C117" '[DA13:DM13]
          TS5(i) = SBname(i) & "R13C131:R13C143" '[EA13:EM13]
        Next
'        .Worksheets("�W�v�\(�����)").Range("AA15").Consolidate sources:=TS1, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("BA15").Consolidate sources:=TS2, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("CA15").Consolidate sources:=TS3, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("DA15").Consolidate sources:=TS4, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("EA15").Consolidate sources:=TS5, Function:=xlSum
        ' ���̑�
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        For i = 1 To ActSHnum
          TS1(i) = SBname(i) & "R14C27:R14C39"   '[AA14:AM14]
          TS2(i) = SBname(i) & "R14C53:R14C65"   '[BA14:BM14]
          TS3(i) = SBname(i) & "R14C79:R14C91"   '[CA14:CM14]
          TS4(i) = SBname(i) & "R14C105:R14C117" '[DA14:DM14]
          TS5(i) = SBname(i) & "R14C131:R14C143" '[EA14:EM14]
        Next
'        .Worksheets("�W�v�\(�����)").Range("AA16").Consolidate sources:=TS1, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("BA16").Consolidate sources:=TS2, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("CA16").Consolidate sources:=TS3, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("DA16").Consolidate sources:=TS4, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("EA16").Consolidate sources:=TS5, Function:=xlSum
        ' ���v
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        For i = 1 To ActSHnum
          TS1(i) = SBname(i) & "R15C27:R15C39"   '[AA15:AM15]
          TS2(i) = SBname(i) & "R15C53:R15C65"   '[BA15:BM15]
          TS3(i) = SBname(i) & "R15C79:R15C91"   '[CA15:CM15]
          TS4(i) = SBname(i) & "R15C105:R15C117" '[DA15:DM15]
          TS5(i) = SBname(i) & "R15C131:R15C143" '[EA15:EM15]
        Next
'        .Worksheets("�W�v�\(�����)").Range("AA17").Consolidate sources:=TS1, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("BA17").Consolidate sources:=TS2, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("CA17").Consolidate sources:=TS3, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("DA17").Consolidate sources:=TS4, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("EA17").Consolidate sources:=TS5, Function:=xlSum
        
        ' ���㌴���N���A
'        .Worksheets("�W�v�\(�����)").Range("AA18:EM31").ClearContents
        ' �O���݌�
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        For i = 1 To ActSHnum
          TS1(i) = SBname(i) & "R16C27:R16C39"   '[AA16:AM16]
          TS2(i) = SBname(i) & "R16C53:R16C65"   '[BA16:BM16]
          TS3(i) = SBname(i) & "R16C79:R16C91"   '[CA16:CM16]
          TS4(i) = SBname(i) & "R16C105:R16C117" '[DA16:DM16]
          TS5(i) = SBname(i) & "R16C131:R16C143" '[EA16:EM16]
        Next
'        .Worksheets("�W�v�\(�����)").Range("AA18").Consolidate sources:=TS1, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("BA18").Consolidate sources:=TS2, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("CA18").Consolidate sources:=TS3, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("DA18").Consolidate sources:=TS4, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("EA18").Consolidate sources:=TS5, Function:=xlSum
        
        ' ���ގd��(�̔�)
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        iNum = 0
        For i = 1 To ActSHnum
'            If Not Workbooks("System�o�c����.xls").Worksheets(ActSHname(i)).Range("AL2") Like "*���i��*" Then
                iNum = iNum + 1
                TS1(iNum) = SBname(i) & "R17C27:R17C39"   '[AA10:AM10]
                TS2(iNum) = SBname(i) & "R17C53:R17C65"   '[BA10:BM10]
                TS3(iNum) = SBname(i) & "R17C79:R17C91"   '[CA10:CM10]
                TS4(iNum) = SBname(i) & "R17C105:R17C117" '[DA10:DM10]
                TS5(iNum) = SBname(i) & "R17C131:R17C143" '[EA10:EM10]
'            End If
        Next
        If iNum > 0 Then
            ReDim aryTS1(iNum) As String, aryTS2(iNum) As String, aryTS3(iNum) As String, aryTS4(iNum) As String, aryTS5(iNum) As String
            For i = 1 To iNum
                aryTS1(i) = TS1(i)
                aryTS2(i) = TS2(i)
                aryTS3(i) = TS3(i)
                aryTS4(i) = TS4(i)
                aryTS5(i) = TS5(i)
            Next
'            .Worksheets("�W�v�\(�����)").Range("AA19").Consolidate sources:=aryTS1, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("BA19").Consolidate sources:=aryTS2, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("CA19").Consolidate sources:=aryTS3, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("DA19").Consolidate sources:=aryTS4, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("EA19").Consolidate sources:=aryTS5, Function:=xlSum
        End If
        ' ���ގd��(���i��)
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        iNum = 0
        For i = 1 To ActSHnum
'            If Workbooks("System�o�c����.xls").Worksheets(ActSHname(i)).Range("AL2") Like "*���i��*" Then
                iNum = iNum + 1
                TS1(iNum) = SBname(i) & "R17C27:R17C39"   '[AA10:AM10]
                TS2(iNum) = SBname(i) & "R17C53:R17C65"   '[BA10:BM10]
                TS3(iNum) = SBname(i) & "R17C79:R17C91"   '[CA10:CM10]
                TS4(iNum) = SBname(i) & "R17C105:R17C117" '[DA10:DM10]
                TS5(iNum) = SBname(i) & "R17C131:R17C143" '[EA10:EM10]
'            End If
        Next
        If iNum > 0 Then
            ReDim aryTS1(iNum) As String, aryTS2(iNum) As String, aryTS3(iNum) As String, aryTS4(iNum) As String, aryTS5(iNum) As String
            For i = 1 To iNum
                aryTS1(i) = TS1(i)
                aryTS2(i) = TS2(i)
                aryTS3(i) = TS3(i)
                aryTS4(i) = TS4(i)
                aryTS5(i) = TS5(i)
            Next
'            .Worksheets("�W�v�\(�����)").Range("AA20").Consolidate sources:=aryTS1, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("BA20").Consolidate sources:=aryTS2, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("CA20").Consolidate sources:=aryTS3, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("DA20").Consolidate sources:=aryTS4, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("EA20").Consolidate sources:=aryTS5, Function:=xlSum
        End If
        ' �H���d��(���i��)
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        iNum = 0
        For i = 1 To ActSHnum
'            If Workbooks("System�o�c����.xls").Worksheets(ActSHname(i)).Range("AL2") Like "*���i��*" Then
                iNum = iNum + 1
                TS1(iNum) = SBname(i) & "R18C27:R18C39"   '[AA10:AM10]
                TS2(iNum) = SBname(i) & "R18C53:R18C65"   '[BA10:BM10]
                TS3(iNum) = SBname(i) & "R18C79:R18C91"   '[CA10:CM10]
                TS4(iNum) = SBname(i) & "R18C105:R18C117" '[DA10:DM10]
                TS5(iNum) = SBname(i) & "R18C131:R18C143" '[EA10:EM10]
'            End If
        Next
        If iNum > 0 Then
            ReDim aryTS1(iNum) As String, aryTS2(iNum) As String, aryTS3(iNum) As String, aryTS4(iNum) As String, aryTS5(iNum) As String
            For i = 1 To iNum
                aryTS1(i) = TS1(i)
                aryTS2(i) = TS2(i)
                aryTS3(i) = TS3(i)
                aryTS4(i) = TS4(i)
                aryTS5(i) = TS5(i)
            Next
'            .Worksheets("�W�v�\(�����)").Range("AA21").Consolidate sources:=aryTS1, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("BA21").Consolidate sources:=aryTS2, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("CA21").Consolidate sources:=aryTS3, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("DA21").Consolidate sources:=aryTS4, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("EA21").Consolidate sources:=aryTS5, Function:=xlSum
        End If
        
        ' �H���d��(�o��)
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        iNum = 0
        For i = 1 To ActSHnum
'            If Not Workbooks("System�o�c����.xls").Worksheets(ActSHname(i)).Range("AL2") Like "*���i��*" Then
                iNum = iNum + 1
                TS1(iNum) = SBname(i) & "R18C27:R18C39"   '[AA10:AM10]
                TS2(iNum) = SBname(i) & "R18C53:R18C65"   '[BA10:BM10]
                TS3(iNum) = SBname(i) & "R18C79:R18C91"   '[CA10:CM10]
                TS4(iNum) = SBname(i) & "R18C105:R18C117" '[DA10:DM10]
                TS5(iNum) = SBname(i) & "R18C131:R18C143" '[EA10:EM10]
'            End If
        Next
        If iNum > 0 Then
            ReDim aryTS1(iNum) As String, aryTS2(iNum) As String, aryTS3(iNum) As String, aryTS4(iNum) As String, aryTS5(iNum) As String
            For i = 1 To iNum
                aryTS1(i) = TS1(i)
                aryTS2(i) = TS2(i)
                aryTS3(i) = TS3(i)
                aryTS4(i) = TS4(i)
                aryTS5(i) = TS5(i)
            Next
'            .Worksheets("�W�v�\(�����)").Range("AA22").Consolidate sources:=aryTS1, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("BA22").Consolidate sources:=aryTS2, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("CA22").Consolidate sources:=aryTS3, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("DA22").Consolidate sources:=aryTS4, Function:=xlSum
'            .Worksheets("�W�v�\(�����)").Range("EA22").Consolidate sources:=aryTS5, Function:=xlSum
        End If
        ' ���̑��`�c�Ɨ��v
        ReDim TS1(ActSHnum), TS2(ActSHnum), TS3(ActSHnum), TS4(ActSHnum), TS5(ActSHnum)
        For i = 1 To ActSHnum
          TS1(i) = SBname(i) & "R19C27:R37C39"   '[AA15:AM15]
          TS2(i) = SBname(i) & "R19C53:R37C65"   '[BA15:BM15]
          TS3(i) = SBname(i) & "R19C79:R37C91"   '[CA15:CM15]
          TS4(i) = SBname(i) & "R19C105:R37C117" '[DA15:DM15]
          TS5(i) = SBname(i) & "R19C131:R37C143" '[EA15:EM15]
        Next
'        .Worksheets("�W�v�\(�����)").Range("AA23").Consolidate sources:=TS1, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("BA23").Consolidate sources:=TS2, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("CA23").Consolidate sources:=TS3, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("DA23").Consolidate sources:=TS4, Function:=xlSum
'        .Worksheets("�W�v�\(�����)").Range("EA23").Consolidate sources:=TS5, Function:=xlSum
        
        '�\��
        .Worksheets("�W�v�\").Range("AK2").Value = INPbk.Worksheets("FILE").Range("K2").Value
        .Worksheets("�W�v�\").Range("AL2").Value = ThisWorkbook.Worksheets("LINK").Range("F3").Value
'        .Worksheets("�W�v�\(�����)").Range("AK2").Value = INPbk.Worksheets("FILE").Range("K2").Value
'        .Worksheets("�W�v�\(�����)").Range("AL2").Value = ThisWorkbook.Worksheets("LINK").Range("F3").Value
        
        '�Δ�\
'        .Worksheets("�Δ�\").Range("AO2").Value = INPbk.Worksheets("FILE").Range("K2").Value
        .Worksheets("�Δ�\").Range("AR2").Value = INPbk.Worksheets("FILE").Range("K2").Value
'        .Worksheets("�Δ�\(�����)").Range("AR2").Value = INPbk.Worksheets("FILE").Range("K2").Value
'        .Worksheets("�Δ�\").Range("AP2").ClearContents
        .Worksheets("�Δ�\").Range("AS2").ClearContents
'        .Worksheets("�Δ�\(�����)").Range("AS2").ClearContents
'        .Worksheets("�Δ�\").Range("AQ2").Value = ThisWorkbook.Worksheets("LINK").Range("F3").Value
        .Worksheets("�Δ�\").Range("AV2").Value = ThisWorkbook.Worksheets("LINK").Range("F3").Value
'        .Worksheets("�Δ�\(�����)").Range("AV2").Value = ThisWorkbook.Worksheets("LINK").Range("F3").Value
        Make_Sh2
        ' �����v
        Call PrintOut_ShT(True)
    End Select
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'���x�Δ�\
Sub Make_Sh2()
Attribute Make_Sh2.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim i As Integer
  Dim SRg1() As String, SRg2() As String, SRg3() As String
  Dim SRg11() As String, SRg22() As String, SRg33() As String
  Dim TRg1() As String, TRg2() As String, TRg3() As String
  Dim TRg11() As String, TRg22() As String, TRg33() As String
  Dim KRg1() As String, KRg2() As String, KRg3() As String
  Dim PRTbk As Object
  
  Set PRTbk = Workbooks(BookName(3))
  
  With PRTbk
    ReDim SRg1(TUKI), SRg2(TUKI), SRg3(TUKI)  'S:���x����  1:���ƌv�� 2:���� 3:�O������
    ReDim SRg11(TUKI), SRg22(TUKI), SRg33(TUKI)  'S:���x����  1:���ƌv�� 2:���� 3:�O������
    ReDim TRg1(TUKI), TRg2(TUKI), TRg3(TUKI)  'S:���x����  1:���ƌv�� 2:���� 3:�O������
    ReDim TRg11(TUKI), TRg22(TUKI), TRg33(TUKI)  'S:���x����  1:���ƌv�� 2:���� 3:�O������
    ReDim KRg1(TUKI), KRg2(TUKI), KRg3(TUKI)  'K:�o���  1:���ƌv�� 2:���� 3:�O������
    
    For i = 1 To TUKI '���̗݌v
      '���x����
      SRg1(i) = "[" & PRTbk.Name & "]�W�v�\!R10C" & CStr(78 + i) & ":R27C" & CStr(78 + i)    '���ƌv��
      SRg2(i) = "[" & PRTbk.Name & "]�W�v�\!R10C" & CStr(26 + i) & ":R27C" & CStr(26 + i)    '����
      SRg3(i) = "[" & PRTbk.Name & "]�W�v�\!R10C" & CStr(104 + i) & ":R27C" & CStr(104 + i)  '�O������
      TRg1(i) = "[" & PRTbk.Name & "]�W�v�\!R28C" & CStr(78 + i) & ":R37C" & CStr(78 + i)    '���ƌv��
      TRg2(i) = "[" & PRTbk.Name & "]�W�v�\!R28C" & CStr(26 + i) & ":R37C" & CStr(26 + i)    '����
      TRg3(i) = "[" & PRTbk.Name & "]�W�v�\!R28C" & CStr(104 + i) & ":R37C" & CStr(104 + i)  '�O������
      '���x����(�����)
'      SRg11(i) = "[" & PRTbk.Name & "]�W�v�\(�����)!R10C" & CStr(78 + i) & ":R41C" & CStr(78 + i)    '���ƌv��
'      SRg22(i) = "[" & PRTbk.Name & "]�W�v�\(�����)!R10C" & CStr(26 + i) & ":R41C" & CStr(26 + i)    '����
'      SRg33(i) = "[" & PRTbk.Name & "]�W�v�\(�����)!R10C" & CStr(104 + i) & ":R41C" & CStr(104 + i)  '�O������
      '�o���
      KRg1(i) = "[" & PRTbk.Name & "]�W�v�\!R50C" & CStr(78 + i) & ":R91C" & CStr(78 + i)    '���ƌv��
      KRg2(i) = "[" & PRTbk.Name & "]�W�v�\!R50C" & CStr(26 + i) & ":R91C" & CStr(26 + i)    '����
      KRg3(i) = "[" & PRTbk.Name & "]�W�v�\!R50C" & CStr(104 + i) & ":R91C" & CStr(104 + i)  '�O������
    Next
    
    '���x����+(�����)
    .Worksheets("�Δ�\").Range("AA10").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R10C" & CStr(78 + TUKI) & ":R27C" & CStr(78 + TUKI), Function:=xlSum
    .Worksheets("�Δ�\").Range("AC10").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R10C" & CStr(52 + TUKI) & ":R27C" & CStr(52 + TUKI), Function:=xlSum
    .Worksheets("�Δ�\").Range("AE10").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R10C" & CStr(26 + TUKI) & ":R27C" & CStr(26 + TUKI), Function:=xlSum
    .Worksheets("�Δ�\").Range("AO10").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R10C" & CStr(104 + TUKI) & ":R27C" & CStr(104 + TUKI), Function:=xlSum
    .Worksheets("�Δ�\").Range("AA29").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R28C" & CStr(78 + TUKI) & ":R37C" & CStr(78 + TUKI), Function:=xlSum
    .Worksheets("�Δ�\").Range("AC29").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R28C" & CStr(52 + TUKI) & ":R37C" & CStr(52 + TUKI), Function:=xlSum
    .Worksheets("�Δ�\").Range("AE29").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R28C" & CStr(26 + TUKI) & ":R37C" & CStr(26 + TUKI), Function:=xlSum
    .Worksheets("�Δ�\").Range("AO29").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R28C" & CStr(104 + TUKI) & ":R37C" & CStr(104 + TUKI), Function:=xlSum
    
    .Worksheets("�Δ�\").Range("AI10").Consolidate sources:=SRg1, Function:=xlSum  '���ƌv��
    .Worksheets("�Δ�\").Range("AI29").Consolidate sources:=TRg1, Function:=xlSum  '���ƌv��
    
    .Worksheets("�Δ�\").Range("AI17").Value = .Worksheets("�W�v�\").Cells(17, 79).Value            '�O���݌� 2008.06.05
    .Worksheets("�Δ�\").Range("AI21").Value = .Worksheets("�W�v�\").Cells(21, 78 + TUKI).Value     '�����݌� 2008.06.05
    
    .Worksheets("�Δ�\").Range("AK10").Consolidate sources:=SRg2, Function:=xlSum  '����
    .Worksheets("�Δ�\").Range("AK29").Consolidate sources:=TRg2, Function:=xlSum  '����
    .Worksheets("�Δ�\").Range("AK17").Value = .Worksheets("�W�v�\").Cells(17, 27).Value            '�O���݌� 2008.06.05
    .Worksheets("�Δ�\").Range("AK21").Value = .Worksheets("�W�v�\").Cells(21, 26 + TUKI).Value     '�����݌� 2008.06.05
    
    .Worksheets("�Δ�\").Range("AS10").Consolidate sources:=SRg3, Function:=xlSum  '�O������
    .Worksheets("�Δ�\").Range("AS29").Consolidate sources:=TRg3, Function:=xlSum  '�O������
    .Worksheets("�Δ�\").Range("AS17").Value = .Worksheets("�W�v�\").Cells(17, 105).Value            '�O���݌� 2008.06.05
    .Worksheets("�Δ�\").Range("AS21").Value = .Worksheets("�W�v�\").Cells(21, 104 + TUKI).Value     '�����݌� 2008.06.05
    
    '�o���
    .Worksheets("�Δ�\").Range("AA51").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R50C" & CStr(78 + TUKI) & ":R91C" & CStr(78 + TUKI), Function:=xlSum
    .Worksheets("�Δ�\").Range("AC51").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R50C" & CStr(52 + TUKI) & ":R91C" & CStr(52 + TUKI), Function:=xlSum
    .Worksheets("�Δ�\").Range("AE51").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R50C" & CStr(26 + TUKI) & ":R91C" & CStr(26 + TUKI), Function:=xlSum
    .Worksheets("�Δ�\").Range("AO51").Consolidate sources:="[" & PRTbk.Name & "]�W�v�\!R50C" & CStr(104 + TUKI) & ":R91C" & CStr(104 + TUKI), Function:=xlSum
    
    .Worksheets("�Δ�\").Range("AI51").Consolidate sources:=KRg1, Function:=xlSum
    .Worksheets("�Δ�\").Range("AK51").Consolidate sources:=KRg2, Function:=xlSum
    .Worksheets("�Δ�\").Range("AS51").Consolidate sources:=KRg3, Function:=xlSum
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'��ʔ̔�
Sub Make_Sh3()
Attribute Make_Sh3.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook.Worksheets("LINK")
    If .Range("D28").Value Then  '201
      PrintOut_Sh3_1
    End If
    If .Range("D29").Value Then  '202
      PrintOut_Sh3_2
    End If
    If .Range("D30").Value Then  '203
      PrintOut_Sh3_3
    End If
    If .Range("D31").Value Then  '204
      PrintOut_Sh3_4
    End If
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'���͕\�i��c�p�j
Sub Make_Sh4()
Attribute Make_Sh4.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim i As Integer
  Dim Rg_R1022C104() As String, Rg_R26C104() As String, Rg_R3437C104() As String
  Dim Rg_R1022C78() As String, Rg_R26C78() As String, Rg_R3437C78() As String
  Dim Rg_R1022C52() As String, Rg_R26C52() As String, Rg_R3437C52() As String
  Dim Rg_R1022C26() As String, Rg_R26C26() As String, Rg_R3437C26() As String
  
  Workbooks(BookName(3)).Worksheets("���͕\").Range("AD3").Value = ThisWorkbook.Worksheets("LINK").Range("B3").Value
  
  With Workbooks(BookName(3)).Worksheets("���͕\")
    Select Case ThisWorkbook.Worksheets("LINK").Range("E4").Value  '1:�����P�� ,2:�W�v
      Case 1  '������������������������������������������������������ �����P��
        ReDim Rg_R1022C104(1), Rg_R26C104(1), Rg_R3437C104(1)
        ReDim Rg_R1022C78(1), Rg_R26C78(1), Rg_R3437C78(1)
        ReDim Rg_R1022C52(1), Rg_R26C52(1), Rg_R3437C52(1)
        ReDim Rg_R1022C26(1), Rg_R26C26(1), Rg_R3437C26(1)
        
        For i = 1 To ActSHnum
          Rg_R1022C104(1) = SBname(i) & "R10C" & CStr(104 + TUKI) & ":R22C" & CStr(104 + TUKI)
          Rg_R26C104(1) = SBname(i) & "R26C" & CStr(104 + TUKI)
          Rg_R3437C104(1) = SBname(i) & "R34C" & CStr(104 + TUKI) & ":R37C" & CStr(104 + TUKI)
          Rg_R1022C78(1) = SBname(i) & "R10C" & CStr(78 + TUKI) & ":R22C" & CStr(78 + TUKI)
          Rg_R26C78(1) = SBname(i) & "R26C" & CStr(78 + TUKI)
          Rg_R3437C78(1) = SBname(i) & "R34C" & CStr(78 + TUKI) & ":R37C" & CStr(78 + TUKI)
          Rg_R1022C52(1) = SBname(i) & "R10C" & CStr(52 + TUKI) & ":R22C" & CStr(52 + TUKI)
          Rg_R26C52(1) = SBname(i) & "R26C" & CStr(52 + TUKI)
          Rg_R3437C52(1) = SBname(i) & "R34C" & CStr(52 + TUKI) & ":R37C" & CStr(52 + TUKI)
          Rg_R1022C26(1) = SBname(i) & "R10C" & CStr(26 + TUKI) & ":R22C" & CStr(26 + TUKI)
          Rg_R26C26(1) = SBname(i) & "R26C" & CStr(26 + TUKI)
          Rg_R3437C26(1) = SBname(i) & "R34C" & CStr(26 + TUKI) & ":R37C" & CStr(26 + TUKI)
          
          '�\��
          .Range("AB3").Value = Workbooks(BookName(1)).Worksheets(ActSHname(i)).Range("AK2").Value
          .Range("AC3").Value = Workbooks(BookName(1)).Worksheets(ActSHname(i)).Range("AL2").Value
          '�O������
          .Range("AA6").Consolidate sources:=Rg_R1022C104, Function:=xlSum
          .Range("AA19").Consolidate sources:=Rg_R26C104, Function:=xlSum
          .Range("AA22").Consolidate sources:=Rg_R3437C104, Function:=xlSum
          '���ƌv��
          .Range("AB6").Consolidate sources:=Rg_R1022C78, Function:=xlSum
          .Range("AB19").Consolidate sources:=Rg_R26C78, Function:=xlSum
          .Range("AB22").Consolidate sources:=Rg_R3437C78, Function:=xlSum
          '���ƌv��
          .Range("AC6").Consolidate sources:=Rg_R1022C52, Function:=xlSum
          .Range("AC19").Consolidate sources:=Rg_R26C52, Function:=xlSum
          .Range("AC22").Consolidate sources:=Rg_R3437C52, Function:=xlSum
          '���ƌv��
          .Range("AD6").Consolidate sources:=Rg_R1022C26, Function:=xlSum
          .Range("AD19").Consolidate sources:=Rg_R26C26, Function:=xlSum
          .Range("AD22").Consolidate sources:=Rg_R3437C26, Function:=xlSum
          
          PrintOut_Sh4
        Next
      Case 2  '������������������������������������������������������ �W�v
        ReDim Rg_R1022C104(ActSHnum), Rg_R26C104(ActSHnum), Rg_R3437C104(ActSHnum)
        ReDim Rg_R1022C78(ActSHnum), Rg_R26C78(ActSHnum), Rg_R3437C78(ActSHnum)
        ReDim Rg_R1022C52(ActSHnum), Rg_R26C52(ActSHnum), Rg_R3437C52(ActSHnum)
        ReDim Rg_R1022C26(ActSHnum), Rg_R26C26(ActSHnum), Rg_R3437C26(ActSHnum)
        
        For i = 1 To ActSHnum
          Rg_R1022C104(i) = SBname(i) & "R10C" & CStr(104 + TUKI) & ":R22C" & CStr(104 + TUKI)
          Rg_R26C104(i) = SBname(i) & "R26C" & CStr(104 + TUKI)
          Rg_R3437C104(i) = SBname(i) & "R34C" & CStr(104 + TUKI) & ":R37C" & CStr(104 + TUKI)
          Rg_R1022C78(i) = SBname(i) & "R10C" & CStr(78 + TUKI) & ":R22C" & CStr(78 + TUKI)
          Rg_R26C78(i) = SBname(i) & "R26C" & CStr(78 + TUKI)
          Rg_R3437C78(i) = SBname(i) & "R34C" & CStr(78 + TUKI) & ":R37C" & CStr(78 + TUKI)
          Rg_R1022C52(i) = SBname(i) & "R10C" & CStr(52 + TUKI) & ":R22C" & CStr(52 + TUKI)
          Rg_R26C52(i) = SBname(i) & "R26C" & CStr(52 + TUKI)
          Rg_R3437C52(i) = SBname(i) & "R34C" & CStr(52 + TUKI) & ":R37C" & CStr(52 + TUKI)
          Rg_R1022C26(i) = SBname(i) & "R10C" & CStr(26 + TUKI) & ":R22C" & CStr(26 + TUKI)
          Rg_R26C26(i) = SBname(i) & "R26C" & CStr(26 + TUKI)
          Rg_R3437C26(i) = SBname(i) & "R34C" & CStr(26 + TUKI) & ":R37C" & CStr(26 + TUKI)
        Next
        
        '�\��
        .Range("AB3").Value = Workbooks(BookName(1)).Worksheets("FILE").Range("K2").Value
        .Range("AC3").Value = ThisWorkbook.Worksheets("LINK").Range("F3").Value
        '�O������
        .Range("AA6").Consolidate sources:=Rg_R1022C104, Function:=xlSum
        .Range("AA19").Consolidate sources:=Rg_R26C104, Function:=xlSum
        .Range("AA22").Consolidate sources:=Rg_R3437C104, Function:=xlSum
        '���ƌv��
        .Range("AB6").Consolidate sources:=Rg_R1022C78, Function:=xlSum
        .Range("AB19").Consolidate sources:=Rg_R26C78, Function:=xlSum
        .Range("AB22").Consolidate sources:=Rg_R3437C78, Function:=xlSum
        '���ƌv��
        .Range("AC6").Consolidate sources:=Rg_R1022C52, Function:=xlSum
        .Range("AC19").Consolidate sources:=Rg_R26C52, Function:=xlSum
        .Range("AC22").Consolidate sources:=Rg_R3437C52, Function:=xlSum
        '���ƌv��
        .Range("AD6").Consolidate sources:=Rg_R1022C26, Function:=xlSum
        .Range("AD19").Consolidate sources:=Rg_R26C26, Function:=xlSum
        .Range("AD22").Consolidate sources:=Rg_R3437C26, Function:=xlSum
        
        PrintOut_Sh4
    End Select
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'�v��Δ�\
Sub Make_Sh5()
Attribute Make_Sh5.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim i As Integer
  Dim Rg_R10C91() As String, Rg_R10C5() As String, Rg_R10C143() As String
  Dim Rg_R50C91() As String, Rg_R50C5() As String, Rg_R50C143() As String
  
  Workbooks(BookName(3)).Worksheets("�v��Δ�").Range("B2").Value = "�� " & Workbooks(BookName(1)).Worksheets("FILE").Range("C2").Value & " ��"
  
  With Workbooks(BookName(3)).Worksheets("�v��Δ�")
    Select Case ThisWorkbook.Worksheets("LINK").Range("E4").Value  '1:�����P�� ,2:�W�v
      Case 1  '������������������������������������������������������ �����P��
        ReDim Rg_R10C91(1), Rg_R10C5(1), Rg_R10C143(1)
        ReDim Rg_R50C91(1), Rg_R50C5(1), Rg_R50C143(1)
        
        For i = 1 To ActSHnum
          Rg_R10C91(1) = SBname(i) & "R10C91:R37C91"
          Rg_R10C5(1) = SBname(i) & "R10C5:R37C5"
          Rg_R10C143(1) = SBname(i) & "R10C143:R37C143"
          Rg_R50C91(1) = SBname(i) & "R50C91:R91C91"
          Rg_R50C5(1) = SBname(i) & "R50C5:R91C5"
          Rg_R50C143(1) = SBname(i) & "R50C143:R91C143"
          
          '�\��
          .Range("AI2").Value = Workbooks(BookName(1)).Worksheets(ActSHname(i)).Range("AK2").Value & Workbooks(BookName(1)).Worksheets(ActSHname(i)).Range("AL2").Value
          '���ƌv��
          .Range("AA10").Consolidate sources:=Rg_R10C91, Function:=xlSum
          .Range("AA50").Consolidate sources:=Rg_R50C91, Function:=xlSum
          '���ѐ���
          .Range("AB10").Consolidate sources:=Rg_R10C5, Function:=xlSum
          .Range("AB50").Consolidate sources:=Rg_R50C5, Function:=xlSum
          '�����̎��ƌv��
          .Range("AE10").Consolidate sources:=Rg_R10C143, Function:=xlSum
          .Range("AE50").Consolidate sources:=Rg_R50C143, Function:=xlSum
          
          PrintOut_Sh5
        Next
      Case 2  '������������������������������������������������������ �W�v
        ReDim Rg_R10C91(ActSHnum), Rg_R10C5(ActSHnum), Rg_R10C143(ActSHnum)
        ReDim Rg_R50C91(ActSHnum), Rg_R50C5(ActSHnum), Rg_R50C143(ActSHnum)
        
        For i = 1 To ActSHnum
          Rg_R10C91(i) = SBname(i) & "R10C91:R37C91"
          Rg_R10C5(i) = SBname(i) & "R10C5:R37C5"
          Rg_R10C143(i) = SBname(i) & "R10C143:R37C143"
          Rg_R50C91(i) = SBname(i) & "R50C91:R91C91"
          Rg_R50C5(i) = SBname(i) & "R50C5:R91C5"
          Rg_R50C143(i) = SBname(i) & "R50C143:R91C143"
        Next
        
        '�\��
        .Range("AI2").Value = Workbooks(BookName(1)).Worksheets("FILE").Range("K2").Value & " " & ThisWorkbook.Worksheets("LINK").Range("F3").Value
        '���ƌv��
        .Range("AA10").Consolidate sources:=Rg_R10C91, Function:=xlSum
        .Range("AA50").Consolidate sources:=Rg_R50C91, Function:=xlSum
        '���ѐ���
        .Range("AB10").Consolidate sources:=Rg_R10C5, Function:=xlSum
        .Range("AB50").Consolidate sources:=Rg_R50C5, Function:=xlSum
        '�����̎��ƌv��
        .Range("AE10").Consolidate sources:=Rg_R10C143, Function:=xlSum
        .Range("AE50").Consolidate sources:=Rg_R50C143, Function:=xlSum
        
        PrintOut_Sh5
    End Select
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'���s
'++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++
'�������сA��/�v��A���ƌv��A�O�����сA���x�Δ�\
Sub PrintOut_ShT(ByVal blnBumon As Boolean)
Attribute PrintOut_ShT.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook.Worksheets("LINK")
    Workbooks(BookName(3)).Worksheets("�W�v�\").Range("AA38:EM39").NumberFormatLocal = Workbooks(BookName(3)).Worksheets("�W�v�\").Range("AA37").NumberFormatLocal
    If .Range("I4").Value Then  '��������
      If .Range("I11").Value Then  '���x����
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$AA$9:$AM$39"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
      If .Range("I12").Value Then  '�o���
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$AA$49:$AM$91"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
      If .Range("I13").Value Then '���v����_
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$AA$99:$AM$119"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
    End If
    If .Range("I5").Value Then  '��/�v��
      If .Range("I11").Value Then  '���x����
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$BA$9:$BM$39"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
      If .Range("I12").Value Then  '�o���
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$BA$49:$BM$91"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
      If .Range("I13").Value Then '���v����_
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$BA$99:$BM$119"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
    End If
    If .Range("I6").Value Then  '���ƌv��
      If .Range("I11").Value Then  '���x����
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$CA$9:$CM$39"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
      If .Range("I12").Value Then  '�o���
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$CA$49:$CM$91"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
      If .Range("I13").Value Then '���v����_
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$CA$99:$CM$119"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
    End If
    If .Range("I7").Value Then  '�O������
      If .Range("I11").Value Then  '���x����
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$DA$9:$DM$39"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
      If .Range("I12").Value Then  '�o���
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$DA$49:$DM$91"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
      If .Range("I13").Value Then '���v����_
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$DA$99:$DM$119"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
    End If
    If .Range("I8").Value Then  '�����v��
      If .Range("I11").Value Then  '���x����
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$EA$9:$EM$39"
        '2012.03.27 �������ƌv��̊����v���X�P
        Workbooks(BookName(3)).Worksheets("�W�v�\").Range("B2").Value = "�� " & Workbooks(BookName(1)).Worksheets("FILE").Range("C2").Value + 1 & " ��"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
      If .Range("I12").Value Then  '�o���
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$EA$49:$EM$91"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
      If .Range("I13").Value Then '���v����_
        Workbooks(BookName(3)).Worksheets("�W�v�\").PageSetup.PrintArea = "$EA$99:$EM$119"
        Workbooks(BookName(3)).Worksheets("�W�v�\").PrintOut copies:=Val(.Range("G4").Value)
      End If
    End If
    '���x�Δ�\
    If .Range("K4").Value Then  '���x����
'      Workbooks(BookName(3)).Worksheets("�Δ�\").PageSetup.PrintArea = "$AA$8:$AR$37"
      Workbooks(BookName(3)).Worksheets("�Δ�\").PageSetup.PrintArea = "$AA$8:$AV$38"
      Workbooks(BookName(3)).Worksheets("�Δ�\").PrintOut copies:=Val(.Range("G4").Value)
      If blnBumon = True Then
'        Workbooks(BookName(3)).Worksheets("�Δ�\(�����)").PageSetup.PrintArea = "$AA$8:$AV$41"
'        Workbooks(BookName(3)).Worksheets("�Δ�\(�����)").PrintOut copies:=Val(.Range("G4").Value)
      End If
    End If
    If .Range("K5").Value Then  '�o���
'      Workbooks(BookName(3)).Worksheets("�Δ�\").PageSetup.PrintArea = "$AA$48:$AR$90"
      Workbooks(BookName(3)).Worksheets("�Δ�\").PageSetup.PrintArea = "$AA$49:$AV$92"
      Workbooks(BookName(3)).Worksheets("�Δ�\").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("K6").Value Then '���v����_
'      Workbooks(BookName(3)).Worksheets("�Δ�\").PageSetup.PrintArea = "$AA$98:$AR$119"
      Workbooks(BookName(3)).Worksheets("�Δ�\").PageSetup.PrintArea = "$AA$99:$AV$120"
      Workbooks(BookName(3)).Worksheets("�Δ�\").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("K4").Value Or .Range("K5").Value Or .Range("K6").Value Then  '�Δ�\�̃O���t
'      If MsgBox("�Δ�\�̃O���t�𔭍s���܂����H", vbQuestion + vbYesNo) = vbYes Then
'        Workbooks(BookName(3)).Worksheets("�Δ�f").PrintOut copies:=Val(.Range("G4").Value)
'      End If
    End If
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'�V�[�g�u201�v
Sub PrintOut_Sh3_1()
Attribute PrintOut_Sh3_1.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook.Worksheets("LINK")
    If .Range("M4").Value Then  '��������
      Workbooks(BookName(1)).Worksheets("201").PageSetup.PrintArea = "$AA$10:$AM$106"
      Workbooks(BookName(1)).Worksheets("201").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M5").Value Then  '��/�v��
      Workbooks(BookName(1)).Worksheets("201").PageSetup.PrintArea = "$BA$10:$BM$106"
      Workbooks(BookName(1)).Worksheets("201").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M6").Value Then  '���ƌv��
      Workbooks(BookName(1)).Worksheets("201").PageSetup.PrintArea = "$CA$10:$CM$106"
      Workbooks(BookName(1)).Worksheets("201").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M7").Value Then  '�O������
      Workbooks(BookName(1)).Worksheets("201").PageSetup.PrintArea = "$DA$10:$DM$106"
      Workbooks(BookName(1)).Worksheets("201").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M8").Value Then  '�����v��
      Workbooks(BookName(1)).Worksheets("201").PageSetup.PrintArea = "$EA$10:$EM$106"
      Workbooks(BookName(1)).Worksheets("201").PrintOut copies:=Val(.Range("G4").Value)
    End If
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'�V�[�g�u202�v
Sub PrintOut_Sh3_2()
Attribute PrintOut_Sh3_2.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook.Worksheets("LINK")
    If .Range("M4").Value Then  '��������
      Workbooks(BookName(1)).Worksheets("202").PageSetup.PrintArea = "$AA$10:$AM$106"
      Workbooks(BookName(1)).Worksheets("202").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M5").Value Then  '��/�v��
      Workbooks(BookName(1)).Worksheets("202").PageSetup.PrintArea = "$BA$10:$BM$106"
      Workbooks(BookName(1)).Worksheets("202").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M6").Value Then  '���ƌv��
      Workbooks(BookName(1)).Worksheets("202").PageSetup.PrintArea = "$CA$10:$CM$106"
      Workbooks(BookName(1)).Worksheets("202").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M7").Value Then  '�O������
      Workbooks(BookName(1)).Worksheets("202").PageSetup.PrintArea = "$DA$10:$DM$106"
      Workbooks(BookName(1)).Worksheets("202").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M8").Value Then  '�����v��
      Workbooks(BookName(1)).Worksheets("202").PageSetup.PrintArea = "$EA$10:$EM$106"
      Workbooks(BookName(1)).Worksheets("202").PrintOut copies:=Val(.Range("G4").Value)
    End If
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'�V�[�g�u203�v
Sub PrintOut_Sh3_3()
Attribute PrintOut_Sh3_3.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook.Worksheets("LINK")
    If .Range("M4").Value Then  '��������
      Workbooks(BookName(1)).Worksheets("203").PageSetup.PrintArea = "$AA$10:$AM$106"
      Workbooks(BookName(1)).Worksheets("203").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M5").Value Then  '��/�v��
      Workbooks(BookName(1)).Worksheets("203").PageSetup.PrintArea = "$BA$10:$BM$106"
      Workbooks(BookName(1)).Worksheets("203").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M6").Value Then  '���ƌv��
      Workbooks(BookName(1)).Worksheets("203").PageSetup.PrintArea = "$CA$10:$CM$106"
      Workbooks(BookName(1)).Worksheets("203").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M7").Value Then  '�O������
      Workbooks(BookName(1)).Worksheets("203").PageSetup.PrintArea = "$DA$10:$DM$106"
      Workbooks(BookName(1)).Worksheets("203").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M8").Value Then  '�����v��
      Workbooks(BookName(1)).Worksheets("203").PageSetup.PrintArea = "$EA$10:$EM$106"
      Workbooks(BookName(1)).Worksheets("203").PrintOut copies:=Val(.Range("G4").Value)
    End If
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'�V�[�g�u204�v
Sub PrintOut_Sh3_4()
Attribute PrintOut_Sh3_4.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook.Worksheets("LINK")
    If .Range("M4").Value Then  '��������
      Workbooks(BookName(1)).Worksheets("204").PageSetup.PrintArea = "$AA$10:$AM$106"
      Workbooks(BookName(1)).Worksheets("204").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M5").Value Then  '��/�v��
      Workbooks(BookName(1)).Worksheets("204").PageSetup.PrintArea = "$BA$10:$BM$106"
      Workbooks(BookName(1)).Worksheets("204").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M6").Value Then  '���ƌv��
      Workbooks(BookName(1)).Worksheets("204").PageSetup.PrintArea = "$CA$10:$CM$106"
      Workbooks(BookName(1)).Worksheets("204").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M7").Value Then  '�O������
      Workbooks(BookName(1)).Worksheets("204").PageSetup.PrintArea = "$DA$10:$DM$106"
      Workbooks(BookName(1)).Worksheets("204").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("M8").Value Then  '�����v��
      Workbooks(BookName(1)).Worksheets("204").PageSetup.PrintArea = "$EA$10:$EM$106"
      Workbooks(BookName(1)).Worksheets("204").PrintOut copies:=Val(.Range("G4").Value)
    End If
  End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'���͕\�i��c�p�j
Sub PrintOut_Sh4()
Attribute PrintOut_Sh4.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook.Worksheets("LINK")
    If .Range("O4").Value Then  '��/�v��
      Workbooks(BookName(3)).Worksheets("���͕\").PageSetup.PrintArea = "$A$1:$M$25"
      Workbooks(BookName(3)).Worksheets("���͕\").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("O5").Value Then  '����
      Workbooks(BookName(3)).Worksheets("���͕\").PageSetup.PrintArea = "$A$26:$M$50"
      Workbooks(BookName(3)).Worksheets("���͕\").PrintOut copies:=Val(.Range("G4").Value)
    End If
  End With
End Sub
     
'++++++++++++++++++++++++++++++++++++++++++++++++++
'�v��Δ�
Sub PrintOut_Sh5()
Attribute PrintOut_Sh5.VB_ProcData.VB_Invoke_Func = " \n14"
  With ThisWorkbook.Worksheets("LINK")
    If .Range("Q4").Value Then  '���x����
      Workbooks(BookName(3)).Worksheets("�v��Δ�").PageSetup.PrintArea = "$AA$8:$AI$37"
      Workbooks(BookName(3)).Worksheets("�v��Δ�").PrintOut copies:=Val(.Range("G4").Value)
    End If
    If .Range("Q5").Value Then  '�o���
      Workbooks(BookName(3)).Worksheets("�v��Δ�").PageSetup.PrintArea = "$AA$48:$AI$90"
      Workbooks(BookName(3)).Worksheets("�v��Δ�").PrintOut copies:=Val(.Range("G4").Value)
    End If
  End With
End Sub

