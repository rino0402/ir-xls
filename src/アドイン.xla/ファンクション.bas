Attribute VB_Name = "�t�@���N�V����"
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'�ޯ����̎擾
Function BookName(ByVal Number As Integer) As String
Attribute BookName.VB_ProcData.VB_Invoke_Func = " \n14"
  Select Case Number
    Case 1  '���͗p�u�b�N��
      BookName = ThisWorkbook.Worksheets("�Ǘ�").Range("B1").Value
    Case 2  '�z���グ�p�u�b�N��
      BookName = ThisWorkbook.Worksheets("�Ǘ�").Range("B2").Value
    Case 3  '���s�p�u�b�N��
      BookName = ThisWorkbook.Worksheets("�Ǘ�").Range("B3").Value
    Case 4  '���ѕۑ��p�u�b�N
      BookName = ThisWorkbook.Worksheets("�Ǘ�").Range("B4").Value
  End Select
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'�f�B���N�g�����̎擾
Function DirName() As String
Attribute DirName.VB_ProcData.VB_Invoke_Func = " \n14"
  DirName = ThisWorkbook.Worksheets("�Ǘ�").Range("C1").Value
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'�t���b�s�[�f�B�X�N�h���C�u���̎擾
Function FD() As String
Attribute FD.VB_ProcData.VB_Invoke_Func = " \n14"
  FD = ThisWorkbook.Worksheets("�Ǘ�").Range("C2").Value
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'�V�X�e���Z�L�����e�B�[�̃`�F�b�N
Function SystemSecure(ByVal n As Integer) As Boolean
  Dim FF As Boolean, SSC_FF As Boolean

  If Dir("C:\Y\SSC.xla") = "" Then
    SSC_FF = False
  Else
    Workbooks.Open "C:\Y\SSC.xla"
    SSC_FF = Application.Run("SystemStartCheck")
    Workbooks("SSC.xla").Close
    FF = True: While FF = False: Wend: FF = False
  End If
  SystemSecure = SSC_FF
  SystemSecure = True
End Function
Sub sysAllClose(ByVal n As Integer)
  Dim i As Object
  For Each i In Workbooks
    i.Close savechanges:=False
  Next
End Sub

Sub AddinTrue()
    Workbooks("�A�h�C��.xla").IsAddin = False
    Workbooks("�A�h�C��.xla").IsAddin = True
End Sub
