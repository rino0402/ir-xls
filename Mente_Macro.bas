Attribute VB_Name = "Mente_Macro"
Option Explicit
Option Base 1

'++++++++++++++++++++++++++++++++++++++++++++++++++
'�����e
'++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++
'�N�x�������e
Sub MENTE_Year()
Attribute MENTE_Year.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim Ch As Integer, Ct As Integer, i As Integer, j As Integer, ThisYear As Integer
  Dim ORGName As String, NewName As String, SheetName As String, ReadBk As String
  Dim WriteSheet As Object
  Dim Frg As Boolean
  Dim wst As Worksheet
  
  With Workbooks(BookName(1))
    Beep
    Ch = MsgBox("�N�x���X�V���܂��B�����ł����H", vbYesNo + vbQuestion, Title:="�����e")
'   2009.05.09 kubota
'    Application.ScreenUpdating = False  '�i��ʕ\�����Ȃ��j
    If Ch = vbYes Then
      ThisYear = Val(.Worksheets("FILE").Range("C2").Value) '����
      ORGName = "��" & CStr(ThisYear - 1) & "��"            '�O��
      Workbooks.Open DirName & BookName(4)  '���ѕۑ��p�u�b�N�̃I�[�v��
      
      '�f�[�^�̍X�V
'      For i = 6 To 34 '�����̐��Ԃ�A�J��Ԃ����
'        If ThisWorkbook.Worksheets("�Ǘ�").Cells(i, "C").Value = True Then
'          SheetName = ThisWorkbook.Worksheets("�Ǘ�").Cells(i, "B").Value
      For Each wst In Workbooks(BookName(1)).Worksheets
        SheetName = wst.Name
        If Len(SheetName) <> 3 Then
            SheetName = ""
        End If
        If SheetName = "000" Then
            SheetName = ""
        End If
        If SheetName = "ALL" Then
            SheetName = ""
        End If
        Debug.Print SheetName
        If SheetName <> "" Then
          Set WriteSheet = Workbooks(BookName(4)).Worksheets(SheetName)   '���ѕۑ��p�̃u�b�N�i�����݂悤�j
          With .Worksheets(WriteSheet.Name)
            .Unprotect password:="sdc2035"            '���b�N�����i�o�c�����̓��͗p�u�b�N�j
            '�\��
            WriteSheet.Range("B2").Value = "�� " & CStr(ThisYear - 1) & " ��"
            WriteSheet.Range("AK2").Value = .Range("AK2").Value
            WriteSheet.Range("AL2").Value = .Range("AL2").Value
            Select Case SheetName
              Case "201", "202", "203", "204"
                ReadBk = "[" & BookName(1) & "]" & SheetName & "!"
                For j = 11 To 101 Step 10
                  '������
                  WriteSheet.Range("AM" & j - 1).Value = .Range("AM" & j - 1).Value
                  '�O�����т̃f�[�^�����т̃u�b�N��
                  WriteSheet.Range("AA" & j).Consolidate sources:=ReadBk & "R" & j & "C105:R" & j + 5 & "C117", Function:=xlSum
                  '�������т̃f�[�^��O�����т�
                  .Range("DA" & j).Consolidate sources:=ReadBk & "R" & j & "C27:R" & j + 5 & "C39", Function:=xlSum
                  '�����v��̃f�[�^�����ƌv���
                  .Range("CA" & j).Consolidate sources:=ReadBk & "R" & j & "C131:R" & j + 5 & "C143", Function:=xlSum
                  '�������сA��/�v��A�����v��̃f�[�^������
                  .Range("AA" & j & ":AM" & j + 5).ClearContents
                  .Range("BA" & j & ":BM" & j + 5).ClearContents
                  .Range("EA" & j & ":EM" & j + 5).ClearContents
                Next
              Case Else
                ReadBk = "[" & BookName(1) & "]" & SheetName & "!"
                '�O�����т̃f�[�^�����т̃u�b�N��
                WriteSheet.Range("AA10").Consolidate sources:=ReadBk & "R10C105:R37C117", Function:=xlSum
                WriteSheet.Range("AA50").Consolidate sources:=ReadBk & "R50C105:R91C117", Function:=xlSum
                '�������т̃f�[�^��O�����т�
                .Range("DA10").Consolidate sources:=ReadBk & "R10C27:R37C39", Function:=xlSum
                .Range("DA50").Consolidate sources:=ReadBk & "R50C27:R91C39", Function:=xlSum
                '�����v��̃f�[�^�����ƌv���
                .Range("CA10").Consolidate sources:=ReadBk & "R10C131:R37C143", Function:=xlSum
                .Range("CA50").Consolidate sources:=ReadBk & "R50C131:R91C143", Function:=xlSum
                '�������сA��/�v��A�����v��̃f�[�^������
                .Range("AA10:AM37,AA50:AM91").ClearContents
                .Range("BA10:BM37,BA50:BM91").ClearContents
                .Range("EA10:EM37,EA50:EM91").ClearContents
            End Select
            SIKI_SetUp SheetName  '���̏�������
          End With
          WriteSheet.Protect password:="sdc2035"  '���b�N�i���ѕۑ��p�u�b�N�j
        End If
      Next
      
      NewName = ORGName
      Ct = 1
      Do
        If Dir(DirName & NewName & ".xls") = "" Then Exit Do
        Ct = Ct + 1
        Frg = True
        NewName = ORGName & "_" & CStr(Ct)
      Loop
      Workbooks(BookName(4)).SaveAs DirName & NewName
      Workbooks(NewName & ".xls").Close
      .Worksheets("FILE").Range("C4").Value = CStr(ThisYear + 1)
      Application.ScreenUpdating = True  '�i��ʕ\������j
      Beep
      MsgBox "�O�����т̃f�[�^���t�@�C�����F" & NewName & " �ŕۑ����܂����B"
    End If
  End With
End Sub
