VERSION 5.00
Begin VB.Form PianoTest 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PianoTest"
   ClientHeight    =   3060
   ClientLeft      =   915
   ClientTop       =   3600
   ClientWidth     =   17190
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PianoTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   17190
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   17175
   End
   Begin ¸ÖÇÙ.Piano Piano1 
      Height          =   2670
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   4710
      AutoPlay        =   0   'False
      mYinSe          =   0
   End
   Begin VB.Menu Close 
      Caption         =   "¹Ø±Õ"
   End
End
Attribute VB_Name = "PianoTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ys As MidiMu

Private Sub Close_Click()
  Unload Me
End Sub

Private Sub Combo1_Click()
  Ys = Combo1.ListIndex
  Piano1.ChangeYinSe Ys
End Sub

Private Sub Form_Load()
  Piano1.OpenDevices
  Piano1.AutoPlay = True

  Dim i As Long

  For i = 0 To 127
  Combo1.AddItem Piano1.GetMidiMu(i)
  Next

  Ys = Piano1.mYinSe
  Combo1.ListIndex = Ys
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Piano1.CloseDevices
End Sub

Private Function GetNote(ByVal KeyCode As Integer, Shift As Integer) As Byte

  Dim Note As Byte

  Note = 0

  Select Case Chr(KeyCode)

  Case "Z"
    Note = 36

  Case "X"
    Note = 38

  Case "C"
    Note = 40

  Case "V"
    Note = 41

  Case "B"
    Note = 43

  Case "N"
    Note = 45

  Case "M"
    Note = 47

  Case "A"
    Note = 48

  Case "S"
    Note = 50

  Case "D"
    Note = 52

  Case "F"
    Note = 53

  Case "G"
    Note = 55

  Case "H"
    Note = 57

  Case "J"
    Note = 59

  Case "K"
    Note = 60

  Case "L"
    Note = 62

  Case "Q"
    Note = 64

    Case "W"
      Note = 65

    Case "E"
      Note = 67

    Case "R"
      Note = 69

    Case "T"
      Note = 71

    Case "Y"
      Note = 72

    Case "U"
      Note = 74

    Case "I"
      Note = 76

    Case "O"
      Note = 77

    Case "P"
      Note = 79

    Case "1"
      Note = 81

    Case "2"
      Note = 83

    Case "3"
      Note = 84

    Case "4"
      Note = 86

    Case "5"
      Note = 88

    Case "6"
      Note = 89

    Case "7"
      Note = 91

    Case "8"
      Note = 93

    Case "9"
      Note = 95

    Case "0"
      Note = 96
  End Select

  'If Shift Then Note = Note + 1
  GetNote = Note
End Function

Private Sub Piano1_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim Note As Byte

  Note = GetNote(KeyCode, Shift)

  If Note = 0 Then Exit Sub
  If Piano1.isSelc(Note) = True Then Exit Sub
  Piano1.selc Note
  'PlayNote Note
End Sub

Private Sub Piano1_KeyUp(KeyCode As Integer, Shift As Integer)

  Dim Note As Byte

  Note = GetNote(KeyCode, Shift)

  If Note = 0 Then Exit Sub
  Piano1.unSelc Note
  'PlayNote Note, 0
End Sub
