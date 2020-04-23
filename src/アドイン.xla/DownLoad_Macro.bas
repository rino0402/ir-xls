Attribute VB_Name = "DownLoad_Macro"
Option Explicit
Option Base 1

'++++++++++++++++++++++++++++++++++++++++++++++++++
'ﾃﾞｨｽｸ作成
'++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++
'ﾌｧｲﾙのｵｰﾌﾟﾝ
Sub Make_Disk()
  Dim Ch As Variant
  Dim BKn As String, mlDir As String
  
  mlDir = "C:\My Documents\メール（本社行）\"
  BKn = "経営資料_" & ThisWorkbook.Worksheets("管理").Range("C3").Value & ".xls"
  Ch = Dir(mlDir & BKn)
'  If Ch = "" Then
    Workbooks.Open DirName & BookName(2)
    Make_TotalSh BookName(2)
    Workbooks(BookName(2)).SaveAs mlDir & BKn
    Workbooks(BKn).Close
'  Else
'    Workbooks.Open mlDir & BKn
'    Make_TotalSh BKn
'    Workbooks(BKn).Close savechanges:=True
'  End If
  MsgBox "データを作成しました。" & Chr(10) & " フォルダ：" & mlDir & Chr(10) & " ファイル：" & BKn
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'集計表の作成
Sub Make_TotalSh(ByVal ACTname As String)
Attribute Make_TotalSh.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim ReadBk As String, HyoDai As String
  
  With Workbooks(ACTname).Worksheets(1)
    ReadBk = "[" & BookName(1) & "]000!"
    .Range("AA10").Consolidate sources:=ReadBk & "R10C27:R37C39", Function:=xlSum   '[AA10:AM37]
    .Range("BA10").Consolidate sources:=ReadBk & "R10C53:R37C65", Function:=xlSum   '[BA10:BM37]
    .Range("CA10").Consolidate sources:=ReadBk & "R10C79:R37C91", Function:=xlSum   '[CA10:CM37]
    .Range("DA10").Consolidate sources:=ReadBk & "R10C105:R37C117", Function:=xlSum '[DA10:DM37]
    .Range("EA10").Consolidate sources:=ReadBk & "R10C131:R37C143", Function:=xlSum '[EA10:EM37]
    .Range("AA50").Consolidate sources:=ReadBk & "R50C27:R91C39", Function:=xlSum   '[AA50:AM90]
    .Range("BA50").Consolidate sources:=ReadBk & "R50C53:R91C65", Function:=xlSum   '[BA50:BM90]
    .Range("CA50").Consolidate sources:=ReadBk & "R50C79:R91C91", Function:=xlSum   '[CA50:CM90]
    .Range("DA50").Consolidate sources:=ReadBk & "R50C105:R91C117", Function:=xlSum '[DA50:DM90]
    .Range("EA50").Consolidate sources:=ReadBk & "R50C131:R91C143", Function:=xlSum '[EA50:EM90]
    .Name = ThisWorkbook.Worksheets("管理").Range("C3").Value
  End With
End Sub

