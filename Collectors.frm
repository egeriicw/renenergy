VERSION 5.00
Begin VB.Form Collectors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collector Parameters"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CollectorsCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Collectors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   20
      Top             =   1800
      Width           =   3855
      Begin VB.TextBox NumCollectorsText 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Number of Collectors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame CollectorData 
      Caption         =   "Collector Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3960
      TabIndex        =   14
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox Check1 
         Caption         =   "Selective Surface"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox AreaText 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox GlazingsText 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox FrULText 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox FrTaText 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Area (per collector panel)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Num. of Glazings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "FrUL (W/m2-C)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "FrTa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.CommandButton CollectorsOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame InputFrame 
      Caption         =   "Collector Alignment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox RhoGText 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox BetaText 
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox GammaText 
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Ground Reflectivity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Collector Slope (degrees)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Azimuth (degrees)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Collectors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CollectorsCancel_Click()
    Collectors.Hide
End Sub

Private Sub CollectorsOk_Click()
    Dim j
    Dim AreaTemp
    
    Gamma = Val(GammaText.Text)
    Beta = Val(BetaText.Text)
    RhoG = Val(RhoGText.Text)
    
    NumCollectors = Val(NumCollectorsText.Text)
    FrTa = Val(FrTaText.Text)
    FrUL = Val(FrULText.Text) * 0.83 * 10.76 * 1.055 * 1.8
    numGlazings = Val(GlazingsText.Text)
    AreaTemp = Val(AreaText.Text)
    If AreaTemp = "" Then
        Area = 0
    Else
        Area = AreaTemp * NumCollectors * 1 / 10.76
    End If
    
    If Gamma = "" Or Beta = "" Or RhoG = "" Or NumCollectors = "" Or FrTa = "" Or FrUL = "" Or numGlazings = "" Or Area = "" Then
        MsgBox ("Must define all variables!")
    Else
        For j = 1 To 8760
            Call SolDat(j)                      '  Calls SolDat() function, to calculate beam and diffuse radiation
        Next j
        Collectors.Hide
        CDefine = 1
    End If
    
    
End Sub

Private Sub Form_Load()
    GammaText.Text = ""
    BetaText.Text = ""
    RhoGText.Text = ""
End Sub

