VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "NMEA-Parser v.1.0"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame2 
      Caption         =   "GPGGA-Data"
      Height          =   3255
      Left            =   4800
      TabIndex        =   23
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox txtALEU 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtALSU 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtALE 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtALS 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtHD 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtSA 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtQU 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtLOD2 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtLAD2 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtUT2 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtLA2 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtLO2 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "AltEllipsoid:"
         Height          =   285
         Left            =   120
         TabIndex        =   43
         ToolTipText     =   "Altitude over Ellipsoid"
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "AltSea:"
         Height          =   285
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Altitude over Sea"
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "HDOP:"
         Height          =   285
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "horizontal dilution of precision"
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "SatellitesIV:"
         Height          =   285
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Sat. in view"
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Quality:"
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Longitude:"
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Latitude:"
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "UtcTime:"
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "GPRMC-Data"
      Height          =   3255
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox txtMD 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtDS 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtCO 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtSK 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtLAD 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtLOD 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtLO 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtLA 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtRW 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtUT 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "MagDec:"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Magnetic Declination"
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Date:"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Course:"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Speed:"
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Latitude:"
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Longitude:"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "RecWarn:"
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "UtcTime:"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&start"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtComPort 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "5"
      Top             =   1680
      Width           =   375
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "ComPort:"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Buffer As String, x As String
Dim rmc As String, ut As String, rw As Boolean, la As Double, lad As String, _
    lo As Double, lod As String, sk As Double, co As Double, ds As String, _
    md As Double, cs As Boolean, _
    gga As String, qu As String, sa As Integer, hd As Double, als As Double, _
    alsu As String, ale As Double, aleu As String
    
Private Sub Command1_Click()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
        Command1.Caption = "&start"
    Else
        MSComm1.CommPort = txtComPort.Text
        MSComm1.Settings = "4800,n,8,1"
        MSComm1.RThreshold = 1
        MSComm1.PortOpen = True
        Command1.Caption = "&stop"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) > 6000 Then Text1.Text = Right(Text1.Text, 3000)
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub MSComm1_OnComm()
If MSComm1.CommEvent = comEvReceive Then
    x = MSComm1.Input
    Buffer = Buffer & x
    'Me.Caption = Len(Buffer)
    
    'Is there a completed GPRMC Datasentence in the buffer?
    rmcS = InStr(1, Buffer, "$GPRMC")
    If rmcS > 0 Then
        rmcE = InStr(rmcS, Buffer, vbCrLf)
        If rmcE > 0 Then 'GPRMC Datasentence found
            rmc = Mid(Buffer, rmcS, rmcE - rmcS)
            decodeRMC rmc, ut, rw, la, lad, lo, lod, sk, co, ds, md, cs
            If cs Then 'checksum is correct
                Text1.Text = Text1.Text & rmc & vbCrLf
                txtUT.Text = ut
                txtRW.Text = rw
                txtLA.Text = la
                txtLAD.Text = lad
                txtLO.Text = lo
                txtLOD.Text = lod
                txtSK.Text = sk
                txtCO.Text = co
                txtDS.Text = ds
                txtMD.Text = md
            End If
            Buffer = Right(Buffer, Len(Buffer) - rmcE) 'remove parsed data from the buffer
        End If
    End If
    
    'Is there a completed GPGGA Datasentence in the buffer?
    ggaS = InStr(1, Buffer, "$GPGGA")
    If ggaS > 0 Then
        ggaE = InStr(ggaS, Buffer, vbCrLf)
        If ggaE > 0 Then 'GPGGA Datasentence found
            gga = Mid(Buffer, ggaS, ggaE - ggaS)
            decodeGGA gga, ut, la, lad, lo, lod, qu, sa, hd, als, alsu, ale, aleu, cs
            If cs Then 'checksum is correct
                Text1.Text = Text1.Text & gga & vbCrLf
                txtUT2.Text = ut
                txtLA2.Text = la
                txtLAD2.Text = lad
                txtLO2.Text = lo
                txtLOD2.Text = lod
                txtQU.Text = qu
                txtSA.Text = sa
                txtHD.Text = hd
                txtALS.Text = als
                txtALSU.Text = alsu
                txtALE.Text = ale
                txtALEU.Text = aleu
            End If
            Buffer = Right(Buffer, Len(Buffer) - ggaE) 'remove parsed data from the buffer
        End If
    End If
End If
End Sub

