Attribute VB_Name = "koneksi"
Option Explicit
Public KonekDb As New ADODB.Connection
Public rsLogin As New ADODB.Recordset
Public SQLSimpan As String
Public SQLUpdate As String
Public SQLDelete As String
Public SQL As String


Sub BukaDataBase()
    Set KonekDb = New ADODB.Connection
    KonekDb.CursorLocation = adUseClient
    KonekDb.ConnectionString = "DRIVER={MYSQL ODBC 5.1 Driver};SERVER=localhost;DATABASE=upj;UID=root;PWD=;OPTION="
    On Error Resume Next
    If KonekDb.State = adStateOpen Then
        KonekDb.Close
        Set KonekDb = New ADODB.Connection
        KonekDb.Open
    Else
        KonekDb.Open
    End If
    If Err.Number <> 0 Then
        MsgBox "Gagal menyambungkan!", vbOKOnly + vbCritical, "Kesalahan"
    End If
End Sub

