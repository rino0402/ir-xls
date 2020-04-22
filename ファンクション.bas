Attribute VB_Name = "ファンクション"
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'ﾌﾞｯｸ名の取得
Function BookName(ByVal Number As Integer) As String
Attribute BookName.VB_ProcData.VB_Invoke_Func = " \n14"
  Select Case Number
    Case 1  '入力用ブック名
      BookName = ThisWorkbook.Worksheets("管理").Range("B1").Value
    Case 2  '吸い上げ用ブック名
      BookName = ThisWorkbook.Worksheets("管理").Range("B2").Value
    Case 3  '発行用ブック名
      BookName = ThisWorkbook.Worksheets("管理").Range("B3").Value
    Case 4  '実績保存用ブック
      BookName = ThisWorkbook.Worksheets("管理").Range("B4").Value
  End Select
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'ディレクトリ名の取得
Function DirName() As String
Attribute DirName.VB_ProcData.VB_Invoke_Func = " \n14"
  DirName = ThisWorkbook.Worksheets("管理").Range("C1").Value
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'フロッピーディスクドライブ名の取得
Function FD() As String
Attribute FD.VB_ProcData.VB_Invoke_Func = " \n14"
  FD = ThisWorkbook.Worksheets("管理").Range("C2").Value
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'システムセキュリティーのチェック
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
    Workbooks("アドイン.xla").IsAddin = False
    Workbooks("アドイン.xla").IsAddin = True
End Sub
