VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo membuat SMS Gateway"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtModem 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1020
      TabIndex        =   7
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3660
      TabIndex        =   5
      Top             =   1395
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   2570
      TabIndex        =   4
      Top             =   1395
      Width           =   975
   End
   Begin VB.ComboBox cmbStorage 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1020
      List            =   "Form1.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   555
      Width           =   3615
   End
   Begin VB.ComboBox cmbPORT 
      Height          =   315
      ItemData        =   "Form1.frx":0023
      Left            =   1020
      List            =   "Form1.frx":0039
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Timer tmrReceiveSms 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1320
   End
   Begin VB.Label Label3 
      Caption         =   "MODEM"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1000
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "STORAGE"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "PORT"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   170
      Width           =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub closeDevice()
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    
    cmbPORT.Enabled = True
    cmbStorage.Enabled = True
    
    tmrReceiveSms.Enabled = False
    txtModem.Text = ""
End Sub

Private Function connectToDevice(ByVal device As String) As Boolean
    Dim objGsm      As ASmsCtrl.GsmOut
    Dim manufaktur  As String
    
    On Error GoTo errHandle
    
    Set objGsm = New ASmsCtrl.GsmOut
    With objGsm
        .Activate SERIAL_NUMBER
        .device = device
    
        manufaktur = .SendCommand("AT+CGMI", 500)  'menampilkan informasi manufactur
        manufaktur = Replace$(manufaktur, vbCrLf, "")
        manufaktur = Replace$(manufaktur, "OK", "")
        manufaktur = Replace$(manufaktur, "ERROR", "")
        manufaktur = Replace$(manufaktur, "AT+CGMI", "")
    End With
    Set objGsm = Nothing
    
    If Len(manufaktur) > 0 Then
        txtModem.Text = manufaktur
        connectToDevice = True
    End If
    
    Exit Function
errHandle:
    connectToDevice = False
End Function

Private Function cekSMSIn() As Boolean
    Dim ret As Integer
    
    strSql = "SELECT COUNT(*) FROM sms_in WHERE status = 0" 'jika status = 0 berarti sms masuk belum di proses
    ret = CInt(dbGetValue(strSql, 0))
    If ret > 0 Then 'ada sms yg belum diproses
        cekSMSIn = True
    End If
End Function

Private Sub sendSMS()
    Dim rsSend          As cRecordset
    Dim cmd             As cCommand
    
    Dim objGsmOut       As ASmsCtrl.GsmOut
    Dim objConstants    As ASmsCtrl.Constants

    Dim phoneNumber     As String
    Dim keyword         As String
    Dim smsBalasan      As String
    
    On Error GoTo errHandle
    
    'cek sms yang belum di proses, ditandai dg status = 0
    strSql = "SELECT id, phone_number, sms_keyword " & _
             "FROM sms_in " & _
             "WHERE status = 0 " & _
             "ORDER BY id"
    Set rsSend = conn.OpenRecordset(strSql)
    If Not rsSend.EOF Then
        Set objGsmOut = New ASmsCtrl.GsmOut
        Set objConstants = New ASmsCtrl.Constants
        
        objGsmOut.Activate SERIAL_NUMBER
        objGsmOut.device = cmbPORT.Text
        objGsmOut.DeviceSpeed = 0
        objGsmOut.RequestStatusReport = False
        objGsmOut.MessageType = objConstants.asMESSAGETYPE_TEXT_MULTIPART
    
        Do While Not rsSend.EOF
            'ganti prefix nomor hp 0 -> +62
            phoneNumber = rep0to62("" & rsSend("phone_number").Value)
            keyword = rsSend("sms_keyword").Value
            
            smsBalasan = getBalasanSms(keyword, phoneNumber)
            
            objGsmOut.MessageRecipient = phoneNumber
            objGsmOut.MessageData = smsBalasan
            objGsmOut.Send
                                                    
            If objGsmOut.LastError = 0 Or objGsmOut.LastError = 23140 Then 'sms sukses dikirim
                'update status sms -> 1
                strSql = "UPDATE sms_in SET status = ?, no_ref = ? " & _
                         "WHERE id = ?"
                Set cmd = conn.CreateCommand(strSql)
                With cmd
                    .SetInt32 1, 1
                    .SetInt32 2, objGsmOut.MessageReference
                    .SetInt32 3, rsSend("id").Value
                    
                    .Execute
                End With
                Set cmd = Nothing
                                         
                'insert ke tabel sms_out, untuk histori sms keluar
                strSql = "INSERT INTO sms_out (phone_number, replay_msg, date_out, time_out) VALUES (?, ?, ?, ?)"
                Set cmd = conn.CreateCommand(strSql)
                With cmd
                    .SetText 1, phoneNumber
                    .SetText 2, smsBalasan
                    .SetDate 3, Format(Now, "yyyy/MM/dd")
                    .SetTime 4, Format(Now, "hh:mm:ss")
                    
                    .Execute
                End With
                Set cmd = Nothing
                
            Else 'sms gagal dikirim
                'update status sms -> 1
                
                'ini masih bisa dikembangkan lagi dengan menambah kolom max_jumlah_kirim di tabel sms_in
                'jadi bisa diberi aturan sms yg gagal dikirim > 3 x baru status smsnya diupdate menjadi 1
                strSql = "UPDATE sms_in SET status = ?, no_ref = ? " & _
                         "WHERE id = ?"
                Set cmd = conn.CreateCommand(strSql)
                With cmd
                    .SetInt32 1, 1
                    .SetInt32 2, objGsmOut.MessageReference
                    .SetInt32 3, rsSend("id").Value
                    
                    .Execute
                End With
                Set cmd = Nothing
            End If
            
            Call Wait(5000)
            
            rsSend.MoveNext
        Loop
        Set objConstants = Nothing
        Set objGsmOut = Nothing
    End If
    
    Exit Sub
errHandle:
    Resume Next
End Sub

Private Sub readSMS()
    Dim objGsmIn        As ASmsCtrl.GsmIn
    Dim objConstants    As ASmsCtrl.Constants
    Dim cmd             As cCommand

    Dim keyword         As String
    Dim phoneNumber     As String
    Dim i               As Integer
    
    On Error GoTo errHandle
    
    Screen.MousePointer = vbHourglass
    cmdStop.Enabled = False
    tmrReceiveSms.Enabled = False
    
    DoEvents
    
    Set objGsmIn = New ASmsCtrl.GsmIn
    Set objConstants = New ASmsCtrl.Constants
    
    With objGsmIn
        .Activate SERIAL_NUMBER
        .device = cmbPORT.Text
        .DeviceSpeed = 0

        .Storage = cmbStorage.ItemData(cmbStorage.ListIndex)
        .DeleteAfterReceive = True 'hapus sms jika sudah dibaca
        .Receive
        
        If .LastError = 0 Or .LastError = 23140 Then 'baca sms sukses
            .GetFirstMessage
        
            strSql = "INSERT INTO sms_in (phone_number, sms_keyword, date_in, time_in) VALUES (?, ?, ?, ?)"
            Set cmd = conn.CreateCommand(strSql)
            conn.BeginTrans
            
            i = 1
            While .LastError = 0
                phoneNumber = rep0to62(.MessageSender)
                keyword = .MessageData
                
                cmd.SetText 1, phoneNumber
                cmd.SetText 2, keyword
                cmd.SetDate 3, Format(Now, "yyyy/MM/dd")
                cmd.SetTime 4, Format(Now, "hh:mm:ss")
                
                cmd.Execute
                
                If i Mod 10 = 0 Then
                    conn.CommitTrans
                    DoEvents
        
                    conn.BeginTrans
                End If
                
                i = i + 1
                
                .GetNextMessage
            Wend
            
            conn.CommitTrans
            Set cmd = Nothing

        End If
    End With
    Set objGsmIn = Nothing
        
    If cekSMSIn Then
        Call sendSMS
    Else
        Call Wait(5000)
    End If
        
    cmdStop.Enabled = True
    Screen.MousePointer = vbDefault
    
    tmrReceiveSms.Enabled = True
    
    Exit Sub
errHandle:
    tmrReceiveSms.Enabled = True
End Sub

Private Sub cmdStart_Click()
    If connectToDevice(cmbPORT.Text) Then
        cmdStart.Enabled = False
        cmbPORT.Enabled = False
        cmbStorage.Enabled = False
        
        cmdStop.Enabled = True
        tmrReceiveSms.Enabled = True
        
    Else
        MsgBox "Koneksi ke modem di port " & cmbPORT.Text & " gagal !!!", vbExclamation, "Peringatan"
    End If
End Sub

Private Sub cmdStop_Click()
    Call closeDevice
End Sub

Private Sub Form_Load()
    cmbPORT.ListIndex = 0
    cmbStorage.ListIndex = 2
End Sub

Private Sub tmrReceiveSms_Timer()
    Call readSMS
End Sub
