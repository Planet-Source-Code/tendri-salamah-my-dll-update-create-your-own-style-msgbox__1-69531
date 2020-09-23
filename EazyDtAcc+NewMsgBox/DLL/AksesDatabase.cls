VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AksesData"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'%#########################################%'
'%Author   : Tendri S (20)                 %'
'%Date     : October 03, 2007 (Update)     %'
'%Location : Bekasi, Indonesia             %'
'%Email    : mizz_daeng@plasa.com          %'
'%Please Do Not Removes Any Copyrights     %'
'%#########################################%'

Private Cn As New ADODB.Connection
Private rs As New ADODB.Recordset
Private Rs2 As New ADODB.Recordset
Private TransOpen As Boolean
Private RsOpen As Boolean
Private RsOpen2 As Boolean
Private ConnectString As String

Private mvarConnectStr As String
Private mvarConnectionType As ConnectType

Public Enum ConnectType
 OnDemand = 0
 Persist = 1
End Enum
Public Property Let TipeKoneksi(ByVal vData As ConnectType)
'mengubah nilai
mvarConnectionType = vData
End Property
Public Property Get TipeKoneksi() As ConnectType
'membaca nilai
TipeKoneksi = mvarConnectionType
End Property
Public Property Let KonekStr(ByVal vData As String)
mvarConnectStr = vData
ConnectString = mvarConnectStr
End Property
Public Property Get KonekStr() As String
KonekStr = mvarConnectStr
End Property
Public Function GetAllData(Tabel As String, OrderBy As String) As Recordset
On Error GoTo ErrorCheck

If Not Cn.State = adStateOpen Then
 ConnectDb
End If

yy = "select * from " & Tabel & " order by " & OrderBy
Set rs = New ADODB.Recordset
rs.Open yy, Cn, adOpenDynamic, adLockOptimistic
RsOpen = True
Set GetAllData = rs

Exit Function

If TipeKoneksi = OnDemand Then
 Set Cn = Nothing
End If

rs.Close
Set rs = Nothing

ErrorCheck:
 MsgBox Err.Description, , "Data Access"
 If TipeKoneksi = OnDemand Then
  DisconnectDB
 End If
End Function
Public Function DeleteData(Tabel As String, Field As String, Data As String) As Boolean
'UNTUK NGEHAPUS RECORD-RECORD DARI DATABASE
On Error GoTo ErrorCheck

If Not Cn.State = adStateOpen Then
 ConnectDb
End If

yy = "select * from " & Tabel & " where " & Field & "='" & Data & "'"
Set rs = New ADODB.Recordset
rs.Open yy, Cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
 xx = "delete from " & Tabel & " where " & Field & "='" & Data & "'"
 Cn.Execute xx
 DeleteData = True
 'MsgBox "Data is completely removed..", vbInformation, "Data Access"
 rs.Close
 Set rs = Nothing
Else
 'MsgBox "Record not found, Delete Failed..", vbCritical, "Pesan"
 mb.TendriMsg "Record not found, Delete Failed..", vbCritical, "Tendri Data Access"
 DeleteData = False
 rs.Close
 Set rs = Nothing
 Exit Function
End If
  
If TipeKoneksi = OnDemand Then
 DisconnectDB
End If

Exit Function

ErrorCheck:
 MsgBox Err.Description, , "Data Access"
 If TipeKoneksi = OnDemand Then
  DisconnectDB
 End If
 DeleteData = False
End Function

Public Function GetData(Tabel As String, NmField As String, PKey As String, Data As String)
On Error GoTo ErrorCheck

If Not Cn.State = adStateOpen Then
 ConnectDb
End If

yy = "select " & NmField & " from " & Tabel & " where " & PKey & "='" & Data & "'"
Set rs = New ADODB.Recordset
rs.Open yy, Cn, adOpenDynamic, adLockOptimistic

If Not rs.EOF Then
 GetData = rs(0)
Else
 GetData = ""
End If

rs.Close
Set rs = Nothing

Exit Function

If TipeKoneksi = OnDemand Then
 Set Cn = Nothing
End If

ErrorCheck:
 MsgBox Err.Description, , "Data Access"
 If TipeKoneksi = OnDemand Then
  DisconnectDB
 End If
End Function
Public Function CekData(Tabel As String, NmField As String, Data As Variant) As Boolean
On Error GoTo ErrorCheck
 
Dim xx As String
    
If Not Cn.State = adStateOpen Then
 ConnectDb
End If

xx = "select * from " & Tabel & " where " & NmField & "='" & Data & "'"
rs.Open xx, Cn, adOpenDynamic, adLockOptimistic

If Not rs.EOF Then
 CekData = True
Else
 CekData = False
End If

rs.Close
Set rs = Nothing
    
If TipeKoneksi = OnDemand Then
 DisconnectDB
End If
    
Exit Function

ErrorCheck:
 MsgBox Err.Description, , "Data Access"
 CekData = False
 If TipeKoneksi = OnDemand Then
  DisconnectDB
 End If
End Function

Public Function FillListCombo(Field As String, Tabel As String, Combo As Object)
On Error Resume Next

Dim x As String
    
Combo.Clear

If Not Cn.State = adStateOpen Then
 ConnectDb
End If

Set rs = New ADODB.Recordset
x = "select " & Field & " from " & Tabel & " order by " & Field & ""
rs.Open x, Cn, adOpenDynamic, adLockOptimistic

While Not rs.EOF
 Combo.AddItem rs(0)
 rs.MoveNext
Wend

rs.Close
Set rs = Nothing

If TipeKoneksi = OnDemand Then
 DisconnectDB
End If
End Function
Public Function InsertData(ByVal SqlStr As String, RecordArray() As Variant) As Boolean
'UNTUK TAMBAH DATA BARU
Dim y As Integer
Dim Size As Integer
Dim RsCount As Integer

On Error GoTo ErrorCheck
    
'jml array
Size = UBound(RecordArray)
    
'konek ke database
If Not Cn.State = adStateOpen Then
 ConnectDb
End If

Cn.BeginTrans

TransOpen = True
Rs2.Open SqlStr, Cn, adOpenKeyset, adLockOptimistic
RsOpen = True
Rs2.AddNew

For y = 0 To Size
 If Not RecordArray(y) = Empty Then
  Rs2.Fields(y).Value = RecordArray(y)
 End If
Next

Rs2.Update
InsertData = True
'MsgBox "Data is completely saved..", vbInformation, "Data Access"

Rs2.Close
RsOpen = False

Cn.CommitTrans

TransOpen = False
Set Rs2 = Nothing

If TipeKoneksi = OnDemand Then
 DisconnectDB
End If

Exit Function

ErrorCheck:
 MsgBox Err.Description, , "Data Access"
 If TransOpen = True Then
  Cn.RollbackTrans
 End If
 InsertData = False
 If TipeKoneksi = OnDemand Then
  DisconnectDB
 End If
End Function
Public Function UpdateData(ByVal SqlStr As String, RecordArray() As Variant, Optional dtField As String) As Boolean
'UNTUK NGEDIT RECORD
Dim y As Integer
Dim Size As Integer
Dim RsCount As Integer
    
On Error GoTo ErrorCheck
    
'jml array
Size = UBound(RecordArray)
    
If Not Cn.State = adStateOpen Then
 ConnectDb
End If

Cn.BeginTrans

TransOpen = True
Rs2.Open SqlStr, Cn, adOpenKeyset, adLockOptimistic
RsOpen = True

If Not Rs2.EOF Then
 'Loop semua record
 y = 0
 RsCount = Rs2.RecordCount
 
 If RsCount > 1 Then
  'MsgBox "More than one record, Update Failed..", vbCritical, "Data Access"
  mb.TendriMsg "More than one record, Update Failed..", vbCritical, "Tendri Data Access"
  UpdateData = False
  Exit Function
 Else
  For y = 0 To Size
   If Not RecordArray(y) = Empty Then
    Rs2.Fields(y).Value = RecordArray(y)
   End If
  Next
            
  Rs2.Update
  UpdateData = True
  If dtField = Empty Then mb.TendriMsg "Data has been updated..", vbInformation, "Tendri Data Access"
  'MsgBox "Data is completely updated..", vbInformation, "Data Access"
  If dtField <> Empty Then
   For I = 0 To RsCount
    mb.TendriMsg "Data " & dtField & " " & Rs2.Fields(I) & " has been updated..", vbInformation, "Tendri Data Access"
    Exit For
   Next
  End If
 End If
 
 Rs2.Close
 RsOpen = False
 
 Cn.CommitTrans
 
 TransOpen = False
 Set Rs2 = Nothing
Else
 Cn.RollbackTrans
 
 Rs2.Close
 RsOpen = False
 Set Rs2 = Nothing
 'MsgBox "No record found, Update Failed..", vbCritical, "Pesan"
 mb.TendriMsg "No record found, Update Failed..", vbCritical, "Tendri Data Access"
 UpdateData = False
 Exit Function
End If

If TipeKoneksi = OnDemand Then
 DisconnectDB
End If

Exit Function

ErrorCheck:
 MsgBox Err.Description, , "Data Access"
 If TransOpen = True Then
  Cn.RollbackTrans
 End If
 UpdateData = False
 If TipeKoneksi = OnDemand Then
  DisconnectDB
 End If
End Function

Private Sub ConnectDb()
'UNTUK MEMBUAT KONEKSI ADO KE DB
On Error GoTo ErrorCheck

Cn.ConnectionTimeout = 0
Cn.CommandTimeout = 0
Cn.CursorLocation = adUseClient
Cn.Open ConnectString

Exit Sub

ErrorCheck:
 MsgBox Err.Description, , "Data Access"
End Sub
Public Function SearchData(ByVal SqlStr As String) As Recordset
On Error Resume Next

If Not Cn.State = adStateOpen Then
 ConnectDb
End If

Rs2.Open SqlStr, Cn, adOpenDynamic, adLockOptimistic
If Not Rs2.EOF Then
 RsOpen = True
 Set SearchData = Rs2
Else
 Rs2.Close
 Set Rs2 = Nothing
 RsOpen = False
End If

Exit Function

If ConnectionType = OnDemand Then
 Set Cn = Nothing
End If
End Function

Private Sub DisconnectDB()
On Error Resume Next
'UNTUK MENUTUP RECORDSET DAN KONEKSI KE DATABASE

If RsOpen = True Then
 Set rs = Nothing
 RsOpen = False
End If

If RsOpen2 = True Then
 Set Rs2 = Nothing
 RsOpen2 = False
End If

Cn.Close
Set Cn = Nothing

Exit Sub
End Sub


'Private Sub DisplayError()
'ConcatString = ""
'Dim ec As Integer
'
'If Cn.Errors.Count > 0 Then
' For ec = 0 To Cn.Errors.Count - 1
'  ConcatString = ConcatString & "DB Error Number: " & Cn.Errors(ec).Number & vbNewLine
'  ConcatString = ConcatString & "     Source: " & Cn.Errors(ec).Source & vbNewLine
'  ConcatString = ConcatString & "             " & Cn.Errors(ec).Description & vbNewLine
'  ConcatString = ConcatString & vbNewLine
' Next
'
' If TransOpen = True Then
'  ConcatString = ConcatString & "ADO.RollBack" & vbNewLine
'  ConcatString = ConcatString & "          Proses penyimpanan data diroll back telah terjadi error. Data tidak tersimpan."
'  ConcatString = ConcatString & vbNewLine
'  RollBackOccured = False
' End If
'Else
' ConcatString = ConcatString & "In Code Error Number: " & Err.Number & vbNewLine
' ConcatString = ConcatString & "          " & Err.Description & vbNewLine
' ConcatString = ConcatString & vbNewLine
'End If
'End Sub
Private Sub Class_Terminate()
 DisconnectDB
End Sub







