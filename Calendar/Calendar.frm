VERSION 5.00
Begin VB.Form Calendar 
   BorderStyle     =   0  'None
   Caption         =   "year"
   ClientHeight    =   4515
   ClientLeft      =   4095
   ClientTop       =   2430
   ClientWidth     =   3810
   BeginProperty Font 
      Name            =   "Lucida Calligraphy"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   3585
      TabIndex        =   1
      Top             =   600
      Width           =   3615
      Begin VB.CommandButton cmdToday 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Today"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox ComboYear 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   345
         Left            =   2040
         TabIndex        =   5
         Text            =   "year"
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox ComboMonth 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Text            =   "month"
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdIndietro 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdAvanti 
         BackColor       =   &H00E0E0E0&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   0
         X2              =   3960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   330
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1560
         Width           =   435
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   0
         X2              =   3960
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   0
         X2              =   3960
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblDayWeek 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dom"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblDayWeek 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   46
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblDayWeek 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mar"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   45
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblDayWeek 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mer"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   44
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblDayWeek 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gio"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   43
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblDayWeek 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ven"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   42
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblDayWeek 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sab"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   41
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   33
         Left            =   2520
         TabIndex        =   40
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   32
         Left            =   2040
         TabIndex        =   39
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   31
         Left            =   1560
         TabIndex        =   38
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   30
         Left            =   1080
         TabIndex        =   37
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   29
         Left            =   600
         TabIndex        =   36
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   28
         Left            =   120
         TabIndex        =   35
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   27
         Left            =   3000
         TabIndex        =   34
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   26
         Left            =   2520
         TabIndex        =   33
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   25
         Left            =   2040
         TabIndex        =   32
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   24
         Left            =   1560
         TabIndex        =   31
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   23
         Left            =   1080
         TabIndex        =   30
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   22
         Left            =   600
         TabIndex        =   29
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   21
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   20
         Left            =   3000
         TabIndex        =   27
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   19
         Left            =   2520
         TabIndex        =   26
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   18
         Left            =   2040
         TabIndex        =   25
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   1560
         TabIndex        =   24
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   1080
         TabIndex        =   23
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   600
         TabIndex        =   22
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   600
         TabIndex        =   20
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   3000
         TabIndex        =   18
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   3000
         TabIndex        =   16
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   2520
         TabIndex        =   15
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   2040
         TabIndex        =   14
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   10
         Left            =   1560
         TabIndex        =   13
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   1080
         TabIndex        =   12
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   600
         TabIndex        =   11
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   2520
         TabIndex        =   10
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2040
         TabIndex        =   9
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1560
         TabIndex        =   8
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1080
         TabIndex        =   7
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   34
         Left            =   3000
         TabIndex        =   6
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   48
         Top             =   600
         Width           =   3675
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   450
      End
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   3600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      Index           =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      Index           =   2
      X1              =   3600
      X2              =   3600
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   3600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblDataSelect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "00/00/00"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   50
      Top             =   120
      Width           =   2835
   End
   Begin VB.Label lblBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Private Declare Function ReleaseCapture Lib "user32" () As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTBOTTOMRIGHT = 17
    Const HTCAPTION = 2
    
Dim DaySelect, MonthSelect, IndiceAnno As Integer
Dim TodayDay, NowMonth As Integer
Dim Flag As Boolean

Sub LoadAll()
ComboYear.Clear
ComboMonth.Clear
ComboMonth.AddItem "January"
ComboMonth.AddItem "February"
ComboMonth.AddItem "March"
ComboMonth.AddItem "April"
ComboMonth.AddItem "May"
ComboMonth.AddItem "june"
ComboMonth.AddItem "July"
ComboMonth.AddItem "August"
ComboMonth.AddItem "September"
ComboMonth.AddItem "October"
ComboMonth.AddItem "November"
ComboMonth.AddItem "December"

lblDayWeek(0).Caption = "Sun"
lblDayWeek(1).Caption = "Mon"
lblDayWeek(2).Caption = "Tue"
lblDayWeek(3).Caption = "Wed"
lblDayWeek(4).Caption = "Thu"
lblDayWeek(5).Caption = "Fri"
lblDayWeek(6).Caption = "Sat"


For i = 1990 To 2100
ComboYear.AddItem i
Next i
End Sub
Sub Pulisci()
For i = 0 To 34
lblDay(i).Caption = " "
Next i
End Sub

Private Sub SettaCal()
On Error Resume Next
MonthSelect = ComboMonth.ListIndex + 1
anno = Val(ComboYear)

Date1 = DateSerial(anno, MonthSelect, 1)
Date2 = DateSerial(anno, MonthSelect + 1, 1)
NumGiorni = Date2 - Date1
GiornoSett = Weekday(Date1) - 1
Pulisci
For i = 1 To NumGiorni
    lblDay(GiornoSett).Caption = i
    GiornoSett = GiornoSett + 1
Next i

End Sub

Private Sub cmdToday_Click()
ComboMonth.ListIndex = NowMonth
ComboYear.ListIndex = IndiceAnno
Shape1.Visible = True
Shape1.Top = lblDay(TodayDay).Top
Shape1.Left = lblDay(TodayDay).Left
End Sub

Sub IndiceAnnoCorrente()
IndiceAnno = 0
For i = 1990 To 2100
If i = DatePart("Yyyy", Date) Then

Exit Sub
End If
IndiceAnno = IndiceAnno + 1
Next i
End Sub
Private Sub ComboMonth_Click()
SettaCal
SettaDatalbl
Shape1.Visible = False
End Sub

Private Sub ComboMonth_LostFocus()
SettaCal
SettaDatalbl
Shape1.Visible = False
End Sub

Private Sub ComboYear_Click()
SettaCal
SettaDatalbl
Shape1.Visible = False
End Sub

Sub SettaDatalbl()
anno = Val(ComboYear)
lblDataSelect.Caption = DaySelect & "  " & ComboMonth.List(MonthSelect - 1) & "  " & anno
End Sub


Private Sub ComboYear_LostFocus()
SettaCal
SettaDatalbl
Shape1.Visible = False
End Sub

Private Sub cmdAvanti_Click()
On Error Resume Next
If ComboMonth.List(ComboMonth.ListIndex) = "December" Then
ComboYear.ListIndex = ComboYear.ListIndex + 1
ComboMonth.ListIndex = 0
lblDataSelect.Caption = "00 " & ComboMonth.List(MonthSelect - 1) & "  " & ComboYear.List(ComboYear.ListIndex)
Exit Sub
Else
ComboMonth.ListIndex = ComboMonth.ListIndex + 1
lblDataSelect.Caption = "00 " & ComboMonth.List(MonthSelect - 1) & "  " & ComboYear.List(ComboYear.ListIndex)
End If
End Sub

Private Sub cmdIndietro_Click()
On Error Resume Next
If ComboMonth.List(ComboMonth.ListIndex) = "January" Then
ComboYear.ListIndex = ComboYear.ListIndex - 1
ComboMonth.ListIndex = 11
lblDataSelect.Caption = "00 " & ComboMonth.List(MonthSelect - 1) & "  " & ComboYear.List(ComboYear.ListIndex)
Exit Sub
Else
ComboMonth.ListIndex = ComboMonth.ListIndex - 1
lblDataSelect.Caption = "00 " & ComboMonth.List(MonthSelect - 1) & "  " & ComboYear.List(ComboYear.ListIndex)
End If
End Sub

Private Sub Command3_Click()
ComboMonth.ListIndex = 7
End Sub


Private Sub Command1_Click()
If Flag = False Then
Picture1.Visible = False
Command1.Caption = ">"
Calendar.Height = lblBarra.Height
Calendar.Width = lblBarra.Width
Flag = True
Else
Picture1.Visible = True
Command1.Caption = "-"
Calendar.Height = Picture1.Top + Picture1.Height
Calendar.Width = lblBarra.Width

Flag = flase
End If

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
LoadAll
Pulisci
IndiceAnnoCorrente

ComboMonth.ListIndex = DatePart("m", Date) - 1
ComboYear.ListIndex = IndiceAnno

'lblDataSelect.Caption = Date

TodayDay = DatePart("d", Date) + 1
NowMonth = DatePart("m", Date) - 1
DaySelect = 0
lblDataSelect.Caption = TodayDay & "  " & ComboMonth.List(MonthSelect - 1) & "  " & ComboYear.List(ComboYear.ListIndex)

Calendar.Height = Picture1.Top + Picture1.Height
Calendar.Width = lblBarra.Width
lblDay(TodayDay).ForeColor = &HC000C0
Shape1.Visible = True
Shape1.Top = lblDay(TodayDay).Top
Shape1.Left = lblDay(TodayDay).Left

End Sub





Private Sub lblBarra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub lblDataSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub lblDay_Click(Index As Integer)
DaySelect = lblDay(Index).Caption
If lblDay(Index).Caption = " " Then
Exit Sub
Else
Shape1.Visible = True
Shape1.Left = lblDay(Index).Left
Shape1.Top = lblDay(Index).Top
SettaDatalbl
End If
End Sub
