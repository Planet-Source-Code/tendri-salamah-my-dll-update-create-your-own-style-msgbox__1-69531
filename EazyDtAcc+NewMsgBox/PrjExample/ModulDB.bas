Attribute VB_Name = "ModDB"
Public AksesDB As New Tendri_S.AksesData
Public MB As New Tendri_S.MessageBox
'Public cn As New ADODB.Connection
'Public rs As New ADODB.Recordset

Sub bukaDB()
Dim konek As String

On Error GoTo ERR_koneksi

konek = "DSN=APOTIK"
'konek = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db_apotik.mdb"

AksesDB.KonekStr = konek

AksesDB.TipeKoneksi = OnDemand

Exit Sub

ERR_koneksi:
MsgBox "Error Koneksi..!", vbCritical, "Pesan Error"
End
End Sub


