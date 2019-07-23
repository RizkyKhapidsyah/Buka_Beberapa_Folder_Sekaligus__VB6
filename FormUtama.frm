Option Explicit

'Author by Rizky Khapidsyah
'Bantu saya menyederhanakan source code ini :)

Dim myPath1 As String
Dim myPath2 As String
Dim myPath3 As String
Dim myPath4 As String
Dim myPath5 As String
Dim myPath6 As String
Dim myPath7 As String
Dim myPath8 As String
Dim myPath9 As String
Dim myPath10 As String
Dim myPath11 As String
Dim myPath12 As String

Function pathOfFile(fileName As String) As String
    Dim posn As Integer
    posn = InStrRev(fileName, "\")
    If posn > 0 Then
        pathOfFile = Left$(fileName, posn)
    Else
        pathOfFile = ""
    End If
End Function

'Atur bagian ini sesuai direktori anda, jika ingin menambahkan direktori, tambahkan variabel baru untuk mendefinisikan direktori baru
Sub AturDirektori()
    myPath1 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\"
    myPath2 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Keperluan Lamaran - Kerja\"
    myPath3 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Keperluan Lamaran - Kerja\#FOTO FORMAL\"
    myPath4 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Keperluan Lamaran - Kerja\baru lagI\"
    myPath5 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Keperluan Lamaran - Kerja\baru lagI\ijazah S1\"
    myPath6 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Keperluan Lamaran - Kerja\Cover Letter (Application Letter)\"
    myPath7 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Keperluan Lamaran - Kerja\Curriculum VItae & Resume\JADI\"
    myPath8 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Keperluan Lamaran - Kerja\DATA PENTING BARU MASUK\"
    myPath9 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Keperluan Lamaran - Kerja\ijazah S1\"
    myPath10 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Keperluan Lamaran - Kerja\kartu keluarga\"
    myPath11 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Ijazah Rina\"
    myPath12 = "D:\Rizky Hafitsyah's Documents\#DATA_PENTING_LAINNYA\PRIBADI\Haldina Suaida\"
End Sub
    
Private Sub Form_Load()
    Call AturDirektori
    Shell "explorer " & pathOfFile(myPath1)
    Shell "explorer " & pathOfFile(myPath2)
    Shell "explorer " & pathOfFile(myPath3)
    Shell "explorer " & pathOfFile(myPath4)
    Shell "explorer " & pathOfFile(myPath5)
    Shell "explorer " & pathOfFile(myPath6)
    Shell "explorer " & pathOfFile(myPath7)
    Shell "explorer " & pathOfFile(myPath8)
    Shell "explorer " & pathOfFile(myPath9)
    Shell "explorer " & pathOfFile(myPath10)
    Shell "explorer " & pathOfFile(myPath11)
    Shell "explorer " & pathOfFile(myPath12)
    End
End Sub
