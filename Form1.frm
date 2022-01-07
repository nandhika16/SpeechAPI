VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Speech dan Region API"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   479
   ScaleMode       =   0  'User
   ScaleWidth      =   1283.962
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deklarasi untuk drag and drop region
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'Speech API
Public WithEvents RC As SpeechLib.SpSharedRecoContext
Attribute RC.VB_VarHelpID = -1
Public myGrammar As SpeechLib.ISpeechRecoGrammar
Public speak As SpeechLib.SpVoice

'Variabel Global
Dim lebar As Long, tinggi As Long
Private Sub Form_Load()
    'Maximize form
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    
    'Variabel untuk posisi tengah region
    lebar = ScaleX(Screen.Width, vbTwips, vbPixels) / 2
    tinggi = ScaleY(Screen.Height, vbTwips, vbPixels) / 2
    
    'Variabel untuk lebar region utama
    lebarRgnParent = 0
    
    'Membuat parent region dan set ke window
    init
    
    On Error GoTo EH
    Set RC = New SpeechLib.SpSharedRecoContext
    With RC
        .EventInterests = SRERecognition Or SREFalseRecognition Or SREStreamStart
        Set myGrammar = .CreateGrammar()
    End With
    myGrammar.DictationSetState SGDSActive
    Set speak = New SpVoice
EH:
    If Err.Number Then ShowErrMsg
End Sub
Private Sub Form_Unload(Cancel As Integer)
    delete
End Sub

Private Sub RC_FalseRecognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal Result As SpeechLib.ISpeechRecoResult)
    speak.speak ("No Recognition")
End Sub

Private Sub RC_Recognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal RecognitionType As SpeechLib.SpeechRecognitionType, ByVal Result As SpeechLib.ISpeechRecoResult)
    'Hasil dari recognition
    res = Result.PhraseInfo.GetText
    
    'Hasil Memanggil Huruf'
    If (res = "p" Or res = "P" Or res = "nine") Then
        speak.speak ("Set region P")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf p
        p lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
        
   ElseIf (res = "u" Or res = "U" Or res = "two") Then
        speak.speak ("Set region U")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf u
        u lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
    
   ElseIf (res = "t" Or res = "T" Or res = "seven") Then
        speak.speak ("Set region T")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf t
        t lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
        
    ElseIf (res = "one") Then
        speak.speak ("Move to Left")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 50px
        geserKiri

        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
        
    ElseIf (res = "n" Or res = "N" Or res = "you") Then
        speak.speak ("Set region N")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf n
        n lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
        
    ElseIf (res = "a" Or res = "A") Then
        speak.speak ("Set region A")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf a
        a lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
        
    ElseIf (res = "d" Or res = "D") Then
        speak.speak ("Set region D")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf d
        d lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
        
    ElseIf (res = "h" Or res = "H" Or res = "three") Then
        speak.speak ("Set region H")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf h
        h lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
        
    ElseIf (res = "i" Or res = "I") Then
        speak.speak ("Set region I")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf i
        i lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
        
    ElseIf (res = "K" Or res = "key") Then
        speak.speak ("Set region K")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf k
        k lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
        
    ElseIf (res = "r" Or res = "R") Then
        speak.speak ("Set region R")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf r
        r lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
        
    ElseIf (res = "m" Or res = "M") Then
        speak.speak ("Set region M")
        
        'Geser region ke kiri sebanyak lebar region utama
        geserKiriSet
        
        'Geser region ke kiri sebanyak 70px
        geserKiri
        
        'Tampilkan region huruf m
        m lebar, tinggi
        
        'Geser region ke kanan sebanyak lebar region utama
        geserKananSet
    Else
        speak.speak ("Wrong Command")
    End If
End Sub
Private Sub ShowErrMsg()

    ' Declare identifiers:
    Const NL = vbNewLine
    Dim t As String

    t = "Desc: " & Err.Description & NL
    t = t & "Err #: " & Err.Number
    MsgBox t, vbExclamation, "Run-Time Error"
    End

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
