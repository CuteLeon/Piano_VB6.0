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
'����: zxrobin   2009.7.5
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

    AcousticGrandPiano�������ѧ����
    BrightAcousticPiano�����ĸ���
    ElectricGrandPiano�����
    Honky_tonkPiano�ưɸ���
    RhodesPiano��͵ĵ����
    ChorusedPiano�Ӻϳ�Ч���ĵ����
    Harpsichord��ܼ��ٲ��ҹŸ���
    Clavichord����ά�����ٻ��ҹŸ���
    Celesta��Ƭ��
    Glockenspiel����
    Musicbox������
    Vibraphone������
    Marimba���ְ�
    Xylophoneľ��
    TubularBells����
    Dulcimer������
    HammondOrgan���˷���
    PercussiveOrgan���ʽ����
    RockOrganҡ������
    ChurchOrgan���÷���
    ReedOrgan�ɹܷ���
    Accordian�ַ���
    Harmonica����
    TangoAccordian̽���ַ���
    AcousticGuitar_nylon_�����Ҽ���
    AcousticGuitar_steel_���Ҽ���
    ElectricGuitar_jazz_��ʿ�缪��
    ElectricGuitar_clean_�����缪��
    ElectricGuitar_muted_�����缪��
    OverdrivenGuitar������Ч���ĵ缪��
    DistortionGuitar��ʧ��Ч���ĵ缪��
    GuitarHarmonics��������
    AcousticBass��˾��ѧ��˾
    ElectricBass_finger_�籴˾ָ��
    ElectricBass_pick_�籴˾��Ƭ
    FretlessBass��Ʒ��˾
    SlapBass1�ƻ�Bass1
    SlapBass2�ƻ�Bass2
    SynthBass1���Ӻϳ�Bass1
    SynthBass2���Ӻϳ�Bass2
    ViolinС����
    Viola������
    Cello������
    Contrabass����������
    TremoloStrings����Ⱥ������ɫ
    PizzicatoStrings����Ⱥ������ɫ
    OrchestralHarp����
    Timpani������
    StringEnsemble1���ֺ�����ɫ1
    StringEnsemble2���ֺ�����ɫ2
    SynthStrings1�ϳ����ֺ�����ɫ1
    SynthStrings2�ϳ����ֺ�����ɫ2
    ChoirAahs�����ϳ�_��
    VoiceOohs����_�
    SynthVoice�ϳ�����
    OrchestraHit�������û�����
    TrumpetС��
    Trombone����
    Tuba���
    MutedTrumpet��������С��
    FrenchHorn������Բ��
    BrassSectionͭ����ͭ������������ɫ
    SynthBrass1�ϳ�ͭ����ɫ1
    SynthBrass2�ϳ�ͭ����ɫ2
    SopranoSax��������˹��
    AltoSax����������˹��
    TenorSax��������˹��
    BaritoneSax��������˹��
    Oboe˫�ɹ�
    EnglishHornӢ����
    Bassoon���ɴ��
    Clarinet���ɹܺڹ�
    Piccolo�̵�
    Flute����
    Recorder����
    PanFlute����
    BottleBlow
    Shakuhachi�ձ��߰�
    Whistle������
    Ocarina�¿�����
    Lead_square_�ϳ�����1����
    Lead_sawtooth_�ϳ�����2��ݲ�
    Lead_caliopelead_�ϳ�����
    Lead_chifflead_�ϳ�����
    Lead_charang_�ϳ�����
    Lead_voice_�ϳ�����6����
    Lead_fifths_�ϳ�����7ƽ�����
    Lead_bass_lead_�ϳ�����8��˾������
    Pad_newage_�ϳ���ɫ1������
    Pad_warm_�ϳ���ɫ��ů
    Pad_polysynth_�ϳ���ɫ
    Pad_choir_�ϳ���ɫ�ϳ�
    Pad_bowed_�ϳ���ɫ
    Pad_metallic_�ϳ���ɫ������
    Pad_halo_�ϳ���ɫ�⻷
    Pad_sweep_�ϳ���ɫ
    FX_rain_�ϳ�Ч������
    FX_soundtrack_�ϳ�Ч������
    FX_crystal_�ϳ�Ч��ˮ��
    FX_atmosphere_�ϳ�Ч������
    FX_brightness_�ϳ�Ч������
    FX_goblins_�ϳ�Ч�����
    FX_echoes_�ϳ�Ч������
    FX_sci_fi_�ϳ�Ч���ƻ�
    Sitar������ӡ��
    Banjo��׿������
    Shamisen�������ձ�
    Kotoʮ�������ձ�
    Kalimba���ְ�
    Bagpipe���
    Fiddle��������
    Shanaiɽ��
    TinkleBell������
    Agogo
    SteelDrums�ֹ�
    Woodblockľ��
    TaikoDrum̫��
    MelodicTomͨͨ��
    SynthDrum�ϳɹ�
    ReverseCymbalͭ��
    GuitarFretNoise������������
    BreathNoise������
    Seashore������
    BirdTweet����
    TelephoneRing�绰��
    Helicopterֱ����
    Applause������
    Gunshotǹ��

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
            GetMidiMu = "0 Acoustic Grand Piano   ����٣���ѧ���٣� "

        Case 1
            GetMidiMu = "1 Bright Acoustic Piano         �����ĸ��� "

        Case 2
            GetMidiMu = "2 Electric Grand Piano              ����� "

        Case 3
            GetMidiMu = "3 Honky-tonk Piano                �ưɸ��� "

        Case 4
            GetMidiMu = "4 Rhodes Piano                ��͵ĵ���� "

        Case 5
            GetMidiMu = "5 Chorused Piano        �Ӻϳ�Ч���ĵ���� "

        Case 6
            GetMidiMu = "6 Harpsichord        ��ܼ��٣����ҹŸ��٣� "

        Case 7
            GetMidiMu = "7 Clavichord     ����ά�����٣����ҹŸ��٣� "

        Case 8
            GetMidiMu = "8 Celesta                           ��Ƭ�� "

        Case 9
            GetMidiMu = "9 Glockenspiel                        ���� "

        Case 10
            GetMidiMu = "10 Music box                        ������ "

        Case 11
            GetMidiMu = "11 Vibraphone                       ������ "

        Case 12
            GetMidiMu = "12 Marimba                          ���ְ� "

        Case 13
            GetMidiMu = "13 Xylophone                          ľ�� "

        Case 14
            GetMidiMu = "14 Tubular Bells                      ���� "

        Case 15
            GetMidiMu = "15 Dulcimer                         ������ "

        Case 16
            GetMidiMu = "16 Hammond Organ                  ���˷��� "

        Case 17
            GetMidiMu = "17 Percussive Organ             ���ʽ���� "

        Case 18
            GetMidiMu = "18 Rock Organ                     ҡ������ "

        Case 19
            GetMidiMu = "19 Church Organ                   ���÷��� "

        Case 20
            GetMidiMu = "20 Reed Organ                     �ɹܷ��� "

        Case 21
            GetMidiMu = "21 Accordian                        �ַ��� "

        Case 22
            GetMidiMu = "22 Harmonica                          ���� "

        Case 23
            GetMidiMu = "23 Tango Accordian              ̽���ַ��� "

        Case 24
            GetMidiMu = "24 Acoustic Guitar (nylon)      �����Ҽ��� "

        Case 25
            GetMidiMu = "25 Acoustic Guitar (steel)        ���Ҽ��� "

        Case 26
            GetMidiMu = "26 Electric Guitar (jazz)       ��ʿ�缪�� "

        Case 27
            GetMidiMu = "27 Electric Guitar (clean)      �����缪�� "

        Case 28
            GetMidiMu = "28 Electric Guitar (muted)      �����缪�� "

        Case 29
            GetMidiMu = "29 Overdriven Guitar    ������Ч���ĵ缪�� "

        Case 30
            GetMidiMu = "30 Distortion Guitar    ��ʧ��Ч���ĵ缪�� "

        Case 31
            GetMidiMu = "31 Guitar Harmonics               �������� "

        Case 32
            GetMidiMu = "32 Acoustic Bass         ��˾����ѧ��˾�� "

        Case 33
            GetMidiMu = "33 Electric Bass(finger)     �籴˾��ָ���� "

        Case 34
            GetMidiMu = "34 Electric Bass (pick)      �籴˾����Ƭ�� "

        Case 35
            GetMidiMu = "35 Fretless Bass                  ��Ʒ��˾ "

        Case 36
            GetMidiMu = "36 Slap Bass 1                   �ƻ�Bass 1 "

        Case 37
            GetMidiMu = "37 Slap Bass 2                  �ƻ�Bass 2 "

        Case 38
            GetMidiMu = "38 Synth Bass 1             ���Ӻϳ�Bass 1 "

        Case 39
            GetMidiMu = "39 Synth Bass 2             ���Ӻϳ�Bass 2 "

        Case 40
            GetMidiMu = "40 Violin                           С���� "

        Case 41
            GetMidiMu = "41 Viola                            ������ "

        Case 42
            GetMidiMu = "42 Cello                            ������ "

        Case 43
            GetMidiMu = "43 Contrabass                   ���������� "

        Case 44
            GetMidiMu = "44 Tremolo Strings          ����Ⱥ������ɫ "

        Case 45
            GetMidiMu = "45 Pizzicato Strings        ����Ⱥ������ɫ "

        Case 46
            GetMidiMu = "46 Orchestral Harp                    ���� "

        Case 47
            GetMidiMu = "47 Timpani                          ������ "

        Case 48
            GetMidiMu = "48 String Ensemble 1         ���ֺ�����ɫ1 "

        Case 49
            GetMidiMu = "49 String Ensemble 2         ���ֺ�����ɫ2 "

        Case 50
            GetMidiMu = "50 Synth Strings 1       �ϳ����ֺ�����ɫ1 "

        Case 51
            GetMidiMu = "51 Synth Strings 2       �ϳ����ֺ�����ɫ2 "

        Case 52
            GetMidiMu = "52 Choir Aahs               �����ϳ������� "

        Case 53
            GetMidiMu = "53 Voice Oohs                   ������ཡ� "

        Case 54
            GetMidiMu = "54 Synth Voice                    �ϳ����� "

        Case 55
            GetMidiMu = "55 Orchestra Hit            �������û����� "

        Case 56
            GetMidiMu = "56 Trumpet                            С�� "

        Case 57
            GetMidiMu = "57 Trombone                           ���� "

        Case 58
            GetMidiMu = "58 Tuba                               ��� "

        Case 59
            GetMidiMu = "59 Muted Trumpet              ��������С�� "

        Case 60
            GetMidiMu = "60 French Horn               �����ţ�Բ�ţ� "

        Case 61
            GetMidiMu = "61 Brass Section ͭ���飨ͭ������������ɫ�� "

        Case 62
            GetMidiMu = "62 Synth Brass 1             �ϳ�ͭ����ɫ1 "

        Case 63
            GetMidiMu = "63 Synth Brass 2             �ϳ�ͭ����ɫ2 "

        Case 64
            GetMidiMu = "64 Soprano Sax                ��������˹�� "

        Case 65
            GetMidiMu = "65 Alto Sax                 ����������˹�� "

        Case 66
            GetMidiMu = "66 Tenor Sax                  ��������˹�� "

        Case 67
            GetMidiMu = "67 Baritone Sax               ��������˹�� "

        Case 68
            GetMidiMu = "68 Oboe                             ˫�ɹ� "

        Case 69
            GetMidiMu = "69 English Horn                     Ӣ���� "

        Case 70
            GetMidiMu = "70 Bassoon                     ���ɣ���ܣ� "

        Case 71
            GetMidiMu = "71 Clarinet                  ���ɹܣ��ڹܣ� "

        Case 72
            GetMidiMu = "72 Piccolo                            �̵� "

        Case 73
            GetMidiMu = "73 Flute                               ���� "

        Case 74
            GetMidiMu = "74 Recorder                           ���� "

        Case 75
            GetMidiMu = "75 Pan Flute                          ���� "

        Case 76
            GetMidiMu = "76 Bottle Blow               [����������ȱ]"

        Case 77
            GetMidiMu = "77 Shakuhachi                     �ձ��߰� "

        Case 78
            GetMidiMu = "78 Whistle                          ������ "

        Case 79
            GetMidiMu = "79 Ocarina                         �¿����� "

        Case 80
            GetMidiMu = "80 Lead 1 (square)        �ϳ�����1�������� "

        Case 81
            GetMidiMu = "81 Lead 2 (sawtooth)    �ϳ�����2����ݲ��� "

        Case 82
            GetMidiMu = "82 Lead 3 (caliope lead)         �ϳ�����3 "

        Case 83
            GetMidiMu = "83 Lead 4 (chiff lead)           �ϳ�����4 "

        Case 84
            GetMidiMu = "84 Lead 5 (charang)              �ϳ�����5 "

        Case 85
            GetMidiMu = "85 Lead 6 (voice)         �ϳ�����6�������� "

        Case 86
            GetMidiMu = "86 Lead 7 (fifths)    �ϳ�����7��ƽ����ȣ� "

        Case 87
            GetMidiMu = "87 Lead 8 (bass+lead)�ϳ�����8����˾�������� "

        Case 88
            GetMidiMu = "88 Pad 1 (new age)      �ϳ���ɫ1�������ͣ� "

        Case 89
            GetMidiMu = "89 Pad 2 (warm)          �ϳ���ɫ2 ����ů�� "

        Case 90
            GetMidiMu = "90 Pad 3 (polysynth)              �ϳ���ɫ3 "

        Case 91
            GetMidiMu = "91 Pad 4 (choir)         �ϳ���ɫ4 ���ϳ��� "

        Case 92
            GetMidiMu = "92 Pad 5 (bowed)                 �ϳ���ɫ5 "

        Case 93
            GetMidiMu = "93 Pad 6 (metallic)    �ϳ���ɫ6 ���������� "

        Case 94
            GetMidiMu = "94 Pad 7 (halo)          �ϳ���ɫ7 ���⻷�� "

        Case 95
            GetMidiMu = "95 Pad 8 (sweep)                 �ϳ���ɫ8 "

        Case 96
            GetMidiMu = "96 FX 1 (rain)             �ϳ�Ч�� 1 ���� "

        Case 97
            GetMidiMu = "97 FX 2 (soundtrack)       �ϳ�Ч�� 2 ���� "

        Case 98
            GetMidiMu = "98 FX 3 (crystal)          �ϳ�Ч�� 3 ˮ�� "

        Case 99
            GetMidiMu = "99 FX 4 (atmosphere)       �ϳ�Ч�� 4 ���� "

        Case 100
            GetMidiMu = "100 FX 5 (brightness)      �ϳ�Ч�� 5 ���� "

        Case 101
            GetMidiMu = "101 FX 6 (goblins)         �ϳ�Ч�� 6 ��� "

        Case 102
            GetMidiMu = "102 FX 7 (echoes)          �ϳ�Ч�� 7 ���� "

        Case 103
            GetMidiMu = "103 FX 8 (sci-fi)          �ϳ�Ч�� 8 �ƻ� "

        Case 104
            GetMidiMu = "104 Sitar                    ��������ӡ�ȣ� "

        Case 105
            GetMidiMu = "105 Banjo                    ��׿�٣����ޣ� "

        Case 106
            GetMidiMu = "106 Shamisen                 �����ߣ��ձ��� "

        Case 107
            GetMidiMu = "107 Koto                   ʮ�����ݣ��ձ��� "

        Case 108
            GetMidiMu = "108 Kalimba                         ���ְ� "

        Case 109
            GetMidiMu = "109 Bagpipe                           ��� "

        Case 110
            GetMidiMu = "110 Fiddle                        �������� "

        Case 111
            GetMidiMu = "111 Shanai                            ɽ�� "

        Case 112
            GetMidiMu = "112 Tinkle Bell                     ������ "

        Case 113
            GetMidiMu = "113 Agogo                    [����������ȱ]"

        Case 114
            GetMidiMu = "114 Steel Drums                       �ֹ� "

        Case 115
            GetMidiMu = "115 Woodblock                         ľ�� "

        Case 116
            GetMidiMu = "116 Taiko Drum                        ̫�� "

        Case 117
            GetMidiMu = "117 Melodic Tom                     ͨͨ�� "

        Case 118
            GetMidiMu = "118 Synth Drum                      �ϳɹ� "

        Case 119
            GetMidiMu = "119 Reverse Cymbal                    ͭ�� "

        Case 120
            GetMidiMu = "120 Guitar Fret Noise         ������������ "

        Case 121
            GetMidiMu = "121 Breath Noise                    ������ "

        Case 122
            GetMidiMu = "122 Seashore                        ������ "

        Case 123
            GetMidiMu = "123 Bird Tweet                        ���� "

        Case 124
            GetMidiMu = "124 Telephone Ring                  �绰�� "

        Case 125
            GetMidiMu = "125 Helicopter                      ֱ���� "

        Case 126
            GetMidiMu = "126 Applause                        ������ "

        Case 127
            GetMidiMu = "127 Gunshot                           ǹ��"
    End Select

End Function
'Download by http://www.codefans.net
