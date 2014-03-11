Attribute VB_Name = "modMain"
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit

Public conn     As cConnection
Public strSql   As String

Public Function dbGetValue(ByVal query As String, ByVal defValue As Variant) As Variant
    Dim rsDbGetValue  As cRecordset
    
    On Error GoTo errHandle
        
    Set rsDbGetValue = conn.OpenRecordset(query, True)
    If Not rsDbGetValue.EOF Then
        If Not IsEmpty(rsDbGetValue(0).Value) Then
            dbGetValue = rsDbGetValue(0).Value
        Else
            dbGetValue = defValue
        End If
    Else
        dbGetValue = defValue
    End If
        
    Exit Function
errHandle:
    dbGetValue = defValue
End Function

Private Function openDb() As Boolean
    On Error GoTo errHandle

    Set conn = New cConnection
    conn.openDb App.Path & "\db\dbsms.db3"
    openDb = True

    Exit Function
errHandle:
    openDb = False
End Function

Public Sub Main()
    If openDb Then
        Form1.Show
        
    Else
        MsgBox "Koneksi ke database gagal !!!", vbExclamation, "Peringatan"
    End If
End Sub
