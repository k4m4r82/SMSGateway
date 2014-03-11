Attribute VB_Name = "modSMS"
Option Explicit

Dim ret As Integer

Public Function rep0to62(ByVal phoneNumber As String) As String
    'fungsi untuk mengganti prefix 0 -> +62
    
    rep0to62 = phoneNumber
    If Left(phoneNumber, 1) = "0" Then rep0to62 = "+62" & Right(phoneNumber, Len(phoneNumber) - 1)
End Function

Private Function isValidNIS(ByVal nis As String) As Boolean
'    On Error GoTo errHandle
    
    strSql = "SELECT COUNT(*) " & _
             "FROM siswa " & _
             "WHERE nis = '" & nis & "'"
    ret = CLng(dbGetValue(strSql, 0))
    If ret > 0 Then isValidNIS = True
    
    Exit Function
errHandle:
    isValidNIS = False
End Function

Private Function isValidHPSiswa(ByVal nis As String, ByVal phoneNumber As String) As Boolean
'    On Error GoTo errHandle
        
    strSql = "SELECT COUNT(*) " & _
             "FROM siswa " & _
             "WHERE nis = '" & nis & "' AND no_hp = '" & phoneNumber & "'"
    ret = CLng(dbGetValue(strSql, 0))
    If ret > 0 Then isValidHPSiswa = True
    
    Exit Function
errHandle:
    isValidHPSiswa = False
End Function

Public Function getBalasanSms(ByVal keywordSms As String, ByVal phoneNumber As String) As String
    Dim rs              As cRecordset
    Dim param1          As String
    Dim arrKeyword()    As String
    
    Dim prefix          As String
    Dim nilai           As String
    Dim nama            As String
    
    Dim tha             As String
    Dim semester        As String
    
    If Len(keywordSms) > 0 Then
        If InStr(1, keywordSms, "#") > 0 Then 'karakter # -> separator keyword
            arrKeyword = Split(keywordSms, "#")
            If Not (Len(arrKeyword(0)) > 0) Then
                getBalasanSms = "Keyword sms salah"
                Exit Function
                
            Else
                'do nothing
            End If
            
        Else
            ReDim arrKeyword(0)
            arrKeyword(0) = keywordSms
        End If
        
    Else
        getBalasanSms = "Keyword sms salah"
        Exit Function
    End If
    
    prefix = arrKeyword(0)
    prefix = UCase$(prefix)
    
    If UBound(arrKeyword) > 0 Then param1 = arrKeyword(1) 'untuk contoh disini param1 bernilai nomor induk siswa
    
    'untuk pengembangan lebih lanjut tahun ajaran dan semester dibuat settingan tersendiri
    tha = "2009/2010"
    semester = 2
    
    Select Case prefix
        Case "TGS"
            'validasi nis siswa
            If Not isValidNIS(param1) Then getBalasanSms = Replace(NIS_SALAH, "<nis>", param1): Exit Function
            
            'validasi no hp siswa
            'nama sekolah sebaiknya disimpan didalam variabel
            If Not isValidHPSiswa(param1, phoneNumber) Then
                getBalasanSms = Replace(HP_UNREG, "<nama_sekolah>", "SMA Negeri Yogyakarta")
                getBalasanSms = Replace(getBalasanSms, "<no_hp>", phoneNumber): Exit Function
            End If
            
            strSql = "SELECT UPPER(nama) FROM siswa WHERE nis = '" & param1 & "'"
            nama = CStr(dbGetValue(strSql, ""))
            
            'mulai proses pencarian nilai
            strSql = "SELECT matapelajaran_kode, nilai " & _
                     "FROM nilai_tugas " & _
                     "WHERE siswa_nis = '" & param1 & "' AND tahun_ajaran = '" & tha & "' AND semester = " & semester & " " & _
                     "ORDER BY matapelajaran_kode"
            Set rs = conn.OpenRecordset(strSql)
            If Not rs.EOF Then
                Do While Not rs.EOF
                    nilai = nilai & rs("matapelajaran_kode").Value & "=" & rs("nilai").Value & ", "
                    rs.MoveNext
                Loop
            End If
            
            If Len(nilai) > 0 Then
                nilai = Left(nilai, Len(nilai) - 2)
                getBalasanSms = "Nilai tugas (" & nama & ") : " & nilai
                
            Else
                getBalasanSms = "Nilai tugas (" & nama & ") sedang dalam proses pendataan"
            End If
            
        Case "UH"
            'validasi nis siswa
            If Not isValidNIS(param1) Then getBalasanSms = Replace(NIS_SALAH, "<nis>", param1): Exit Function
            
            'validasi no hp siswa
            'nama sekolah sebaiknya disimpan didalam variabel
            If Not isValidHPSiswa(param1, phoneNumber) Then
                getBalasanSms = Replace(HP_UNREG, "<nama_sekolah>", "SMA Negeri Yogyakarta")
                getBalasanSms = Replace(getBalasanSms, "<no_hp>", phoneNumber): Exit Function
            End If
            
            strSql = "SELECT UPPER(nama) FROM siswa WHERE nis = '" & param1 & "'"
            nama = CStr(dbGetValue(strSql, ""))
            
            'mulai proses pencarian nilai
            strSql = "SELECT matapelajaran_kode, nilai " & _
                     "FROM nilai_ulangan " & _
                     "WHERE siswa_nis = '" & param1 & "' AND tahun_ajaran = '" & tha & "' AND semester = " & semester & " " & _
                     "ORDER BY matapelajaran_kode"
            Set rs = conn.OpenRecordset(strSql)
            If Not rs.EOF Then
                Do While Not rs.EOF
                    nilai = nilai & rs("matapelajaran_kode").Value & "=" & rs("nilai").Value & ", "
                    rs.MoveNext
                Loop
            End If
            
            If Len(nilai) > 0 Then
                nilai = Left(nilai, Len(nilai) - 2)
                getBalasanSms = "Nilai ulangan (" & nama & ") : " & nilai
                
            Else
                getBalasanSms = "Nilai ulangan (" & nama & ") sedang dalam proses pendataan"
            End If
            
        Case Else
            getBalasanSms = "Keyword sms salah"
    End Select
End Function
