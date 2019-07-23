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
    myPath1 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath2 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath3 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath4 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath5 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath6 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath7 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath8 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath9 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath10 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath11 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
    myPath12 = "(Isi Direktori Anda Disini, 'contoh D:\New Folder\New Folder\XXX\')"
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
