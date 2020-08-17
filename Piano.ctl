VERSION 5.00
Begin VB.UserControl Piano 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17205
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1147
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   0
      Left            =   0
      Picture         =   "Piano.ctx":0000
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1144
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   17160
   End
   Begin VB.PictureBox KW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   8
      Left            =   0
      Picture         =   "Piano.ctx":97A2
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox KW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Index           =   7
      Left            =   2880
      Picture         =   "Piano.ctx":9D0F
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox KW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   6
      Left            =   1080
      Picture         =   "Piano.ctx":A189
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox KW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   5
      Left            =   1800
      Picture         =   "Piano.ctx":A616
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox KW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   4
      Left            =   2160
      Picture         =   "Piano.ctx":AB55
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox KW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   2
      Left            =   2520
      Picture         =   "Piano.ctx":B0B4
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox KW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   1
      Left            =   1440
      Picture         =   "Piano.ctx":B5F3
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox KW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   0
      Left            =   720
      Picture         =   "Piano.ctx":BB52
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox Buffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6480
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox KW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   3
      Left            =   360
      Picture         =   "Piano.ctx":BFDF
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Index           =   1
      Left            =   4800
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1162
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   17430
   End
End
Attribute VB_Name = "Piano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'作者: zxrobin   2009.7.5
'Download by http://www.codefans.net
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Declare Function BitBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal dwRop As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function ScreenToClient _
                Lib "user32" (ByVal hwnd As Long, _
                              lpPoint As POINTAPI) As Long
                              
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long

Private Declare Function midiOutOpen _
                Lib "winmm.dll" (lphMidiOut As Long, _
                                 ByVal uDeviceID As Long, _
                                 ByVal dwCallback As Long, _
                                 ByVal dwInstance As Long, _
                                 ByVal dwFlags As Long) As Long

Private Declare Function midiOutShortMsg _
                Lib "winmm.dll" (ByVal hMidiOut As Long, _
                                 ByVal dwMsg As Long) As Long

Public Enum MidiMu

    AcousticGrandPiano大钢琴声学钢琴
    BrightAcousticPiano明亮的钢琴
    ElectricGrandPiano电钢琴
    Honky_tonkPiano酒吧钢琴
    RhodesPiano柔和的电钢琴
    ChorusedPiano加合唱效果的电钢琴
    Harpsichord羽管键琴拨弦古钢琴
    Clavichord科拉维科特琴击弦古钢琴
    Celesta钢片琴
    Glockenspiel钟琴
    Musicbox八音盒
    Vibraphone颤音琴
    Marimba马林巴
    Xylophone木琴
    TubularBells管钟
    Dulcimer大扬琴
    HammondOrgan击杆风琴
    PercussiveOrgan打击式风琴
    RockOrgan摇滚风琴
    ChurchOrgan教堂风琴
    ReedOrgan簧管风琴
    Accordian手风琴
    Harmonica口琴
    TangoAccordian探戈手风琴
    AcousticGuitar_nylon_尼龙弦吉他
    AcousticGuitar_steel_钢弦吉他
    ElectricGuitar_jazz_爵士电吉他
    ElectricGuitar_clean_清音电吉他
    ElectricGuitar_muted_闷音电吉他
    OverdrivenGuitar加驱动效果的电吉他
    DistortionGuitar加失真效果的电吉他
    GuitarHarmonics吉他和音
    AcousticBass大贝司声学贝司
    ElectricBass_finger_电贝司指弹
    ElectricBass_pick_电贝司拨片
    FretlessBass无品贝司
    SlapBass1掌击Bass1
    SlapBass2掌击Bass2
    SynthBass1电子合成Bass1
    SynthBass2电子合成Bass2
    Violin小提琴
    Viola中提琴
    Cello大提琴
    Contrabass低音大提琴
    TremoloStrings弦乐群颤音音色
    PizzicatoStrings弦乐群拨弦音色
    OrchestralHarp竖琴
    Timpani定音鼓
    StringEnsemble1弦乐合奏音色1
    StringEnsemble2弦乐合奏音色2
    SynthStrings1合成弦乐合奏音色1
    SynthStrings2合成弦乐合奏音色2
    ChoirAahs人声合唱_啊
    VoiceOohs人声_嘟
    SynthVoice合成人声
    OrchestraHit管弦乐敲击齐奏
    Trumpet小号
    Trombone长号
    Tuba大号
    MutedTrumpet加弱音器小号
    FrenchHorn法国号圆号
    BrassSection铜管组铜管乐器合奏音色
    SynthBrass1合成铜管音色1
    SynthBrass2合成铜管音色2
    SopranoSax高音萨克斯风
    AltoSax次中音萨克斯风
    TenorSax中音萨克斯风
    BaritoneSax低音萨克斯风
    Oboe双簧管
    EnglishHorn英国管
    Bassoon巴松大管
    Clarinet单簧管黑管
    Piccolo短笛
    Flute长笛
    Recorder竖笛
    PanFlute排箫
    BottleBlow
    Shakuhachi日本尺八
    Whistle口哨声
    Ocarina奥卡雷那
    Lead_square_合成主音1方波
    Lead_sawtooth_合成主音2锯齿波
    Lead_caliopelead_合成主音
    Lead_chifflead_合成主音
    Lead_charang_合成主音
    Lead_voice_合成主音6人声
    Lead_fifths_合成主音7平行五度
    Lead_bass_lead_合成主音8贝司加主音
    Pad_newage_合成音色1新世纪
    Pad_warm_合成音色温暖
    Pad_polysynth_合成音色
    Pad_choir_合成音色合唱
    Pad_bowed_合成音色
    Pad_metallic_合成音色金属声
    Pad_halo_合成音色光环
    Pad_sweep_合成音色
    FX_rain_合成效果雨声
    FX_soundtrack_合成效果音轨
    FX_crystal_合成效果水晶
    FX_atmosphere_合成效果大气
    FX_brightness_合成效果明亮
    FX_goblins_合成效果鬼怪
    FX_echoes_合成效果回声
    FX_sci_fi_合成效果科幻
    Sitar西塔尔印度
    Banjo班卓琴美洲
    Shamisen三昧线日本
    Koto十三弦筝日本
    Kalimba卡林巴
    Bagpipe风笛
    Fiddle民族提琴
    Shanai山奈
    TinkleBell叮当铃
    Agogo
    SteelDrums钢鼓
    Woodblock木鱼
    TaikoDrum太鼓
    MelodicTom通通鼓
    SynthDrum合成鼓
    ReverseCymbal铜钹
    GuitarFretNoise吉他换把杂音
    BreathNoise呼吸声
    Seashore海浪声
    BirdTweet鸟鸣
    TelephoneRing电话铃
    Helicopter直升机
    Applause鼓掌声
    Gunshot枪声

End Enum

Private Type POINTAPI

    x As Long
    y As Long

End Type

Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private hMidiIn As Long, hMidiOut As Long

Dim Sel(128)    As Boolean, tempNote As Byte

Dim AP          As Boolean, YinSe As MidiMu

Public Event Click(ByVal Note As Byte)

Public Event Move(ByVal Note As Byte, _
                  Button As Integer, _
                  Shift As Integer, _
                  x As Single, _
                  y As Single)

Public Event Down(ByVal Note As Byte, _
                  Button As Integer, _
                  Shift As Integer, _
                  x As Single, _
                  y As Single)

Public Event Up(ByVal Note As Byte, _
                Button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)

Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Public Function OpenDevices(Optional ByVal ID As Long = 0) As Long

    If hMidiOut <> 0 Then Exit Function
    If 0 = midiOutOpen(hMidiOut, ID, 0, 0, 0) Then
        OpenDevices = hMidiOut
        ChangeYinSe YinSe
    Else
        OpenDevices = 0
    End If

End Function

Public Function CloseDevices() As Long

    If hMidiOut <> 0 Then midiOutClose hMidiOut: hMidiOut = 0
End Function

Public Sub PlayNote(ByVal Note As Byte, _
                    Optional ByVal s As Byte = 112, _
                    Optional ByVal n As Byte = 0)

    If hMidiOut = 0 Then Exit Sub

    Dim MidiMsg As Long

    MidiMsg = &H90& + Note * &H100& + s * &H10000 + n
    midiOutShortMsg hMidiOut, MidiMsg
End Sub

Public Sub PlayMsg(ByVal MidiMsg As Long)

    If hMidiOut = 0 Then Exit Sub
    midiOutShortMsg hMidiOut, MidiMsg
End Sub

Public Sub ChangeYinSe(Optional ByVal Ys As MidiMu = 0, Optional ByVal n As Byte = 0)

    PlayMsg &HC0& + n + Ys * &H100&
      
End Sub

Public Sub selc(ByVal Index As Byte)
    Sel(Index) = True

    If AP Then PlayNote Index
    DrawKeybd
End Sub

Public Sub unSelc(ByVal Index As Byte)
    Sel(Index) = False

    If AP Then PlayNote Index, 0
    DrawKeybd
End Sub

Public Function isSelc(ByVal Index As Byte)
    isSelc = Sel(Index)
End Function

Private Sub Picture1_Click(Index As Integer)

    Dim Po As POINTAPI

    GetCursorPos Po
    ScreenToClient Picture1(0).hwnd, Po
    RaiseEvent Click(GetNote(Po.x, Po.y))
End Sub

Private Sub Picture1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Picture1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Picture1_MouseDown(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

    Dim Po As POINTAPI, hei As Boolean

    GetCursorPos Po
    ScreenToClient Picture1(0).hwnd, Po
    tempNote = GetNote(Po.x, Po.y)
    selc tempNote
    RaiseEvent Down(tempNote, Button, Shift, x, y)
End Sub

Private Sub Picture1_MouseMove(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

    Dim Po As POINTAPI, hei As Boolean

    GetCursorPos Po
    ScreenToClient Picture1(0).hwnd, Po

    If Button <> 0 Then
        If tempNote <> GetNote(Po.x, Po.y) Then
            unSelc tempNote
            tempNote = GetNote(Po.x, Po.y)
            selc tempNote
        End If

    Else
        tempNote = GetNote(Po.x, Po.y)
    End If

    RaiseEvent Move(tempNote, Button, Shift, x, y)
End Sub

Private Sub Picture1_MouseUp(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    Dim Po As POINTAPI, hei As Boolean

    GetCursorPos Po
    ScreenToClient Picture1(0).hwnd, Po
    unSelc tempNote
    RaiseEvent Up(tempNote, Button, Shift, x, y)
End Sub

Private Sub UserControl_Initialize()
   
    Dim i As Long

    Picture1(0).Top = 0
    Picture1(0).Left = 0
    UserControl.Width = Picture1(0).Width * 15 + 15
    UserControl.Height = Picture1(0).Height * 15 + 15
    Buffer.Width = Picture1(0).Width
    Buffer.Height = Picture1(0).Height
    Picture1(1).Height = Picture1(0).Height
    Picture1(1).Width = Picture1(0).Width
    Picture1(1).Picture = Picture1(0).Picture

    For i = 0 To 128
        Sel(i) = False
    Next

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    AP = PropBag.ReadProperty("AutoPlay", False)
    YinSe = PropBag.ReadProperty("mYinSe", 0)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = Picture1(0).Width * 15 + 15
    UserControl.Height = Picture1(0).Height * 15 + 15
    Buffer.Width = Picture1(0).Width
    Buffer.Height = Picture1(0).Height
End Sub

Private Function DrawKeybd()

    Dim i As Long, x As Long, y As Long

    Buffer.Cls
    BitBlt Buffer.hDC, 0, 0, Picture1(0).Width, Picture1(0).Height, Picture1(1).hDC, 0, 0, SRCCOPY

    For i = 21 To 110

        If Sel(i) Then
            GetPos i, x, y
            BitBlt Buffer.hDC, x, 0, KW(y).Width, KW(y).Height, KW(y).hDC, 0, 0, SRCINVERT
        End If

    Next

    BitBlt Picture1(0).hDC, 0, 0, Picture1(0).Width, Picture1(0).Height, Buffer.hDC, 0, 0, SRCCOPY
    Picture1(0).Refresh
End Function
   
Public Function GetNote(ByVal x As Long, _
                        ByVal y As Long, _
                        Optional ByRef isHei As Boolean = False) As Byte

    Dim TH As Long, TW As Long, TX As Long, TS As Long, S1 As Long, S2 As Long, No As Long

    Dim Uw As Long

    Uw = Int(Round(Picture1(0).Width / 51))
    isHei = False
    TH = Picture1(0).Height
    TW = Picture1(0).Width
    TX = x \ Uw
    TS = x Mod Uw
    S1 = TX \ 7
    S2 = TX Mod 7
    No = S1 * 12 + 21

    If y < TH / 6 * 4 Then
        If TS > Uw * 2 / 3 And S2 + 3 <> 7 And S2 + 6 <> 7 Then No = No + 1
        If TS < Uw / 3 And S2 + 2 <> 7 And S2 + 5 <> 7 Then No = No - 1
    End If

    No = No + S2 * 2

    If S2 > 1 Then No = No - 1
    If S2 > 4 Then No = No - 1
    GetNote = No
End Function

Public Function GetPos(ByVal Note As Byte, ByRef x As Long, ByRef y As Long) As Boolean

    Dim x1 As Long, x2 As Long, X3 As Long

    Dim Uw As Long

    Uw = Int(Round(Picture1(0).Width / 51))

    x1 = (Note - 21) \ 12
    x2 = (Note - 21) Mod 12
    X3 = x2

    If x2 > 2 Then X3 = X3 + 1
    If x2 > 7 Then X3 = X3 + 1
   
    x = (x1 * 7 + X3 \ 2) * Uw
    y = X3 \ 2

    If (X3 + 1) Mod 2 = 0 Then y = 7
    If Note = 21 Then y = 2
    If Note = 108 Then y = 8
    If y = 7 Then x = x + Uw / 3 * 2

End Function

Public Property Get AutoPlay() As Boolean
    AutoPlay = AP
End Property

Public Property Let AutoPlay(ByVal vNewValue As Boolean)
    AP = vNewValue
End Property

Public Property Get mYinSe() As MidiMu
    mYinSe = YinSe
End Property

Public Property Let mYinSe(ByVal vNewValue As MidiMu)
    YinSe = vNewValue
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoPlay", AP)
    Call PropBag.WriteProperty("mYinSe", YinSe)
End Sub

Public Function GetMidiMu(ByVal Ys As MidiMu) As String

    Select Case Ys

        Case 0
            GetMidiMu = "0 Acoustic Grand Piano   大钢琴（声学钢琴） "

        Case 1
            GetMidiMu = "1 Bright Acoustic Piano         明亮的钢琴 "

        Case 2
            GetMidiMu = "2 Electric Grand Piano              电钢琴 "

        Case 3
            GetMidiMu = "3 Honky-tonk Piano                酒吧钢琴 "

        Case 4
            GetMidiMu = "4 Rhodes Piano                柔和的电钢琴 "

        Case 5
            GetMidiMu = "5 Chorused Piano        加合唱效果的电钢琴 "

        Case 6
            GetMidiMu = "6 Harpsichord        羽管键琴（拨弦古钢琴） "

        Case 7
            GetMidiMu = "7 Clavichord     科拉维科特琴（击弦古钢琴） "

        Case 8
            GetMidiMu = "8 Celesta                           钢片琴 "

        Case 9
            GetMidiMu = "9 Glockenspiel                        钟琴 "

        Case 10
            GetMidiMu = "10 Music box                        八音盒 "

        Case 11
            GetMidiMu = "11 Vibraphone                       颤音琴 "

        Case 12
            GetMidiMu = "12 Marimba                          马林巴 "

        Case 13
            GetMidiMu = "13 Xylophone                          木琴 "

        Case 14
            GetMidiMu = "14 Tubular Bells                      管钟 "

        Case 15
            GetMidiMu = "15 Dulcimer                         大扬琴 "

        Case 16
            GetMidiMu = "16 Hammond Organ                  击杆风琴 "

        Case 17
            GetMidiMu = "17 Percussive Organ             打击式风琴 "

        Case 18
            GetMidiMu = "18 Rock Organ                     摇滚风琴 "

        Case 19
            GetMidiMu = "19 Church Organ                   教堂风琴 "

        Case 20
            GetMidiMu = "20 Reed Organ                     簧管风琴 "

        Case 21
            GetMidiMu = "21 Accordian                        手风琴 "

        Case 22
            GetMidiMu = "22 Harmonica                          口琴 "

        Case 23
            GetMidiMu = "23 Tango Accordian              探戈手风琴 "

        Case 24
            GetMidiMu = "24 Acoustic Guitar (nylon)      尼龙弦吉他 "

        Case 25
            GetMidiMu = "25 Acoustic Guitar (steel)        钢弦吉他 "

        Case 26
            GetMidiMu = "26 Electric Guitar (jazz)       爵士电吉他 "

        Case 27
            GetMidiMu = "27 Electric Guitar (clean)      清音电吉他 "

        Case 28
            GetMidiMu = "28 Electric Guitar (muted)      闷音电吉他 "

        Case 29
            GetMidiMu = "29 Overdriven Guitar    加驱动效果的电吉他 "

        Case 30
            GetMidiMu = "30 Distortion Guitar    加失真效果的电吉他 "

        Case 31
            GetMidiMu = "31 Guitar Harmonics               吉他和音 "

        Case 32
            GetMidiMu = "32 Acoustic Bass         大贝司（声学贝司） "

        Case 33
            GetMidiMu = "33 Electric Bass(finger)     电贝司（指弹） "

        Case 34
            GetMidiMu = "34 Electric Bass (pick)      电贝司（拨片） "

        Case 35
            GetMidiMu = "35 Fretless Bass                  无品贝司 "

        Case 36
            GetMidiMu = "36 Slap Bass 1                   掌击Bass 1 "

        Case 37
            GetMidiMu = "37 Slap Bass 2                  掌击Bass 2 "

        Case 38
            GetMidiMu = "38 Synth Bass 1             电子合成Bass 1 "

        Case 39
            GetMidiMu = "39 Synth Bass 2             电子合成Bass 2 "

        Case 40
            GetMidiMu = "40 Violin                           小提琴 "

        Case 41
            GetMidiMu = "41 Viola                            中提琴 "

        Case 42
            GetMidiMu = "42 Cello                            大提琴 "

        Case 43
            GetMidiMu = "43 Contrabass                   低音大提琴 "

        Case 44
            GetMidiMu = "44 Tremolo Strings          弦乐群颤音音色 "

        Case 45
            GetMidiMu = "45 Pizzicato Strings        弦乐群拨弦音色 "

        Case 46
            GetMidiMu = "46 Orchestral Harp                    竖琴 "

        Case 47
            GetMidiMu = "47 Timpani                          定音鼓 "

        Case 48
            GetMidiMu = "48 String Ensemble 1         弦乐合奏音色1 "

        Case 49
            GetMidiMu = "49 String Ensemble 2         弦乐合奏音色2 "

        Case 50
            GetMidiMu = "50 Synth Strings 1       合成弦乐合奏音色1 "

        Case 51
            GetMidiMu = "51 Synth Strings 2       合成弦乐合奏音色2 "

        Case 52
            GetMidiMu = "52 Choir Aahs               人声合唱“啊” "

        Case 53
            GetMidiMu = "53 Voice Oohs                   人声“嘟” "

        Case 54
            GetMidiMu = "54 Synth Voice                    合成人声 "

        Case 55
            GetMidiMu = "55 Orchestra Hit            管弦乐敲击齐奏 "

        Case 56
            GetMidiMu = "56 Trumpet                            小号 "

        Case 57
            GetMidiMu = "57 Trombone                           长号 "

        Case 58
            GetMidiMu = "58 Tuba                               大号 "

        Case 59
            GetMidiMu = "59 Muted Trumpet              加弱音器小号 "

        Case 60
            GetMidiMu = "60 French Horn               法国号（圆号） "

        Case 61
            GetMidiMu = "61 Brass Section 铜管组（铜管乐器合奏音色） "

        Case 62
            GetMidiMu = "62 Synth Brass 1             合成铜管音色1 "

        Case 63
            GetMidiMu = "63 Synth Brass 2             合成铜管音色2 "

        Case 64
            GetMidiMu = "64 Soprano Sax                高音萨克斯风 "

        Case 65
            GetMidiMu = "65 Alto Sax                 次中音萨克斯风 "

        Case 66
            GetMidiMu = "66 Tenor Sax                  中音萨克斯风 "

        Case 67
            GetMidiMu = "67 Baritone Sax               低音萨克斯风 "

        Case 68
            GetMidiMu = "68 Oboe                             双簧管 "

        Case 69
            GetMidiMu = "69 English Horn                     英国管 "

        Case 70
            GetMidiMu = "70 Bassoon                     巴松（大管） "

        Case 71
            GetMidiMu = "71 Clarinet                  单簧管（黑管） "

        Case 72
            GetMidiMu = "72 Piccolo                            短笛 "

        Case 73
            GetMidiMu = "73 Flute                               长笛 "

        Case 74
            GetMidiMu = "74 Recorder                           竖笛 "

        Case 75
            GetMidiMu = "75 Pan Flute                          排箫 "

        Case 76
            GetMidiMu = "76 Bottle Blow               [中文名称暂缺]"

        Case 77
            GetMidiMu = "77 Shakuhachi                     日本尺八 "

        Case 78
            GetMidiMu = "78 Whistle                          口哨声 "

        Case 79
            GetMidiMu = "79 Ocarina                         奥卡雷那 "

        Case 80
            GetMidiMu = "80 Lead 1 (square)        合成主音1（方波） "

        Case 81
            GetMidiMu = "81 Lead 2 (sawtooth)    合成主音2（锯齿波） "

        Case 82
            GetMidiMu = "82 Lead 3 (caliope lead)         合成主音3 "

        Case 83
            GetMidiMu = "83 Lead 4 (chiff lead)           合成主音4 "

        Case 84
            GetMidiMu = "84 Lead 5 (charang)              合成主音5 "

        Case 85
            GetMidiMu = "85 Lead 6 (voice)         合成主音6（人声） "

        Case 86
            GetMidiMu = "86 Lead 7 (fifths)    合成主音7（平行五度） "

        Case 87
            GetMidiMu = "87 Lead 8 (bass+lead)合成主音8（贝司加主音） "

        Case 88
            GetMidiMu = "88 Pad 1 (new age)      合成音色1（新世纪） "

        Case 89
            GetMidiMu = "89 Pad 2 (warm)          合成音色2 （温暖） "

        Case 90
            GetMidiMu = "90 Pad 3 (polysynth)              合成音色3 "

        Case 91
            GetMidiMu = "91 Pad 4 (choir)         合成音色4 （合唱） "

        Case 92
            GetMidiMu = "92 Pad 5 (bowed)                 合成音色5 "

        Case 93
            GetMidiMu = "93 Pad 6 (metallic)    合成音色6 （金属声） "

        Case 94
            GetMidiMu = "94 Pad 7 (halo)          合成音色7 （光环） "

        Case 95
            GetMidiMu = "95 Pad 8 (sweep)                 合成音色8 "

        Case 96
            GetMidiMu = "96 FX 1 (rain)             合成效果 1 雨声 "

        Case 97
            GetMidiMu = "97 FX 2 (soundtrack)       合成效果 2 音轨 "

        Case 98
            GetMidiMu = "98 FX 3 (crystal)          合成效果 3 水晶 "

        Case 99
            GetMidiMu = "99 FX 4 (atmosphere)       合成效果 4 大气 "

        Case 100
            GetMidiMu = "100 FX 5 (brightness)      合成效果 5 明亮 "

        Case 101
            GetMidiMu = "101 FX 6 (goblins)         合成效果 6 鬼怪 "

        Case 102
            GetMidiMu = "102 FX 7 (echoes)          合成效果 7 回声 "

        Case 103
            GetMidiMu = "103 FX 8 (sci-fi)          合成效果 8 科幻 "

        Case 104
            GetMidiMu = "104 Sitar                    西塔尔（印度） "

        Case 105
            GetMidiMu = "105 Banjo                    班卓琴（美洲） "

        Case 106
            GetMidiMu = "106 Shamisen                 三昧线（日本） "

        Case 107
            GetMidiMu = "107 Koto                   十三弦筝（日本） "

        Case 108
            GetMidiMu = "108 Kalimba                         卡林巴 "

        Case 109
            GetMidiMu = "109 Bagpipe                           风笛 "

        Case 110
            GetMidiMu = "110 Fiddle                        民族提琴 "

        Case 111
            GetMidiMu = "111 Shanai                            山奈 "

        Case 112
            GetMidiMu = "112 Tinkle Bell                     叮当铃 "

        Case 113
            GetMidiMu = "113 Agogo                    [中文名称暂缺]"

        Case 114
            GetMidiMu = "114 Steel Drums                       钢鼓 "

        Case 115
            GetMidiMu = "115 Woodblock                         木鱼 "

        Case 116
            GetMidiMu = "116 Taiko Drum                        太鼓 "

        Case 117
            GetMidiMu = "117 Melodic Tom                     通通鼓 "

        Case 118
            GetMidiMu = "118 Synth Drum                      合成鼓 "

        Case 119
            GetMidiMu = "119 Reverse Cymbal                    铜钹 "

        Case 120
            GetMidiMu = "120 Guitar Fret Noise         吉他换把杂音 "

        Case 121
            GetMidiMu = "121 Breath Noise                    呼吸声 "

        Case 122
            GetMidiMu = "122 Seashore                        海浪声 "

        Case 123
            GetMidiMu = "123 Bird Tweet                        鸟鸣 "

        Case 124
            GetMidiMu = "124 Telephone Ring                  电话铃 "

        Case 125
            GetMidiMu = "125 Helicopter                      直升机 "

        Case 126
            GetMidiMu = "126 Applause                        鼓掌声 "

        Case 127
            GetMidiMu = "127 Gunshot                           枪声"
    End Select

End Function
'Download by http://www.codefans.net
