VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   6264
   ClientLeft      =   1572
   ClientTop       =   1596
   ClientWidth     =   8208
   LinkTopic       =   "Form1"
   ScaleHeight     =   6264
   ScaleWidth      =   8208
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6480
      TabIndex        =   74
      Top             =   4320
      Width           =   1332
   End
   Begin VB.CommandButton BlackRook1 
      BackColor       =   &H8000000A&
      Caption         =   "R"
      Height          =   492
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   1800
      Width           =   492
   End
   Begin VB.CommandButton WhiteQueen 
      Caption         =   "Q"
      Height          =   492
      Left            =   3120
      TabIndex        =   70
      Top             =   5160
      Width           =   492
   End
   Begin VB.CommandButton WhiteKing 
      Caption         =   "KG"
      Height          =   492
      Left            =   3600
      TabIndex        =   69
      Top             =   5160
      Width           =   492
   End
   Begin VB.CommandButton WhiteBishop2 
      Caption         =   "B"
      Height          =   492
      Left            =   4080
      TabIndex        =   68
      Top             =   5160
      Width           =   492
   End
   Begin VB.CommandButton WhiteRook2 
      Caption         =   "R"
      Height          =   492
      Left            =   5040
      TabIndex        =   67
      Top             =   5160
      Width           =   492
   End
   Begin VB.CommandButton WhiteRook1 
      Caption         =   "R"
      Height          =   492
      Left            =   1680
      TabIndex        =   66
      Top             =   5160
      Width           =   492
   End
   Begin VB.CommandButton WhitePawn4 
      Caption         =   "P"
      Height          =   492
      Left            =   3120
      TabIndex        =   65
      Top             =   4680
      Width           =   492
   End
   Begin VB.CommandButton WhiteHorse1 
      Caption         =   "K"
      Height          =   492
      Left            =   2160
      TabIndex        =   0
      Top             =   5160
      Width           =   492
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   63
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   1800
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   62
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   1800
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   61
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   1800
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   60
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   1800
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   59
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   1800
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   58
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   1800
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   57
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   1800
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   56
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   1800
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   55
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2280
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   54
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2280
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   53
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2280
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   52
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2280
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   51
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2280
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   50
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2280
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   49
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2280
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   48
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2280
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   47
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2760
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   46
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2760
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   45
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2760
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   44
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2760
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   43
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2760
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   42
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2760
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   41
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2760
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   40
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2760
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   39
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3240
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   38
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3240
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   37
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3240
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   36
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3240
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   35
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3240
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   34
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3240
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   33
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3240
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   32
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3240
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   31
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3720
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   30
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3720
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   29
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3720
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   28
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3720
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   27
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3720
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   26
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3720
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   25
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3720
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   24
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3720
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   23
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4200
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   22
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4200
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   21
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4200
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   20
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4200
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   19
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4200
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   18
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4200
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   17
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   16
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   15
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4680
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   14
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   13
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   12
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   11
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4680
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   10
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   9
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   8
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   7
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   6
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   5
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   4
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   3
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   2
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   1
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   0
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   732
      Left            =   2160
      TabIndex        =   72
      Top             =   360
      Width           =   3012
   End
   Begin VB.Label lblMessage 
      Caption         =   "R = Rook K=Knight   P = Pawn Q= Queen KG=King   B=Bishop  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      Left            =   6480
      TabIndex        =   71
      Top             =   1920
      Width           =   1452
   End
   Begin VB.Shape Square 
      BackColor       =   &H80000004&
      FillColor       =   &H80000001&
      Height          =   492
      Index           =   63
      Left            =   5040
      Shape           =   1  'Square
      Top             =   1800
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   62
      Left            =   4560
      Shape           =   1  'Square
      Top             =   1800
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   61
      Left            =   4080
      Shape           =   1  'Square
      Top             =   1800
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   60
      Left            =   3600
      Shape           =   1  'Square
      Top             =   1800
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   59
      Left            =   3120
      Shape           =   1  'Square
      Top             =   1800
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   58
      Left            =   2640
      Shape           =   1  'Square
      Top             =   1800
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   57
      Left            =   2160
      Shape           =   1  'Square
      Top             =   1800
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   56
      Left            =   1680
      Shape           =   1  'Square
      Top             =   1800
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   55
      Left            =   5040
      Shape           =   1  'Square
      Top             =   2280
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   54
      Left            =   4560
      Shape           =   1  'Square
      Top             =   2280
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   53
      Left            =   4080
      Shape           =   1  'Square
      Top             =   2280
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   52
      Left            =   3600
      Shape           =   1  'Square
      Top             =   2280
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   51
      Left            =   3120
      Shape           =   1  'Square
      Top             =   2280
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   50
      Left            =   2640
      Shape           =   1  'Square
      Top             =   2280
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   49
      Left            =   2160
      Shape           =   1  'Square
      Top             =   2280
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   48
      Left            =   1680
      Shape           =   1  'Square
      Top             =   2280
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   47
      Left            =   5040
      Shape           =   1  'Square
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   46
      Left            =   4560
      Shape           =   1  'Square
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   45
      Left            =   4080
      Shape           =   1  'Square
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   44
      Left            =   3600
      Shape           =   1  'Square
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   43
      Left            =   3120
      Shape           =   1  'Square
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   42
      Left            =   2640
      Shape           =   1  'Square
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   41
      Left            =   2160
      Shape           =   1  'Square
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   40
      Left            =   1680
      Shape           =   1  'Square
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   39
      Left            =   5040
      Shape           =   1  'Square
      Top             =   3240
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   38
      Left            =   4560
      Shape           =   1  'Square
      Top             =   3240
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   37
      Left            =   4080
      Shape           =   1  'Square
      Top             =   3240
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   36
      Left            =   3600
      Shape           =   1  'Square
      Top             =   3240
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   35
      Left            =   3120
      Shape           =   1  'Square
      Top             =   3240
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   34
      Left            =   2640
      Shape           =   1  'Square
      Top             =   3240
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   33
      Left            =   2160
      Shape           =   1  'Square
      Top             =   3240
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   32
      Left            =   1680
      Shape           =   1  'Square
      Top             =   3240
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   31
      Left            =   5040
      Shape           =   1  'Square
      Top             =   3720
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   30
      Left            =   4560
      Shape           =   1  'Square
      Top             =   3720
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   29
      Left            =   4080
      Shape           =   1  'Square
      Top             =   3720
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   28
      Left            =   3600
      Shape           =   1  'Square
      Top             =   3720
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   27
      Left            =   3120
      Shape           =   1  'Square
      Top             =   3720
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   26
      Left            =   2640
      Shape           =   1  'Square
      Top             =   3720
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   25
      Left            =   2160
      Shape           =   1  'Square
      Top             =   3720
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   24
      Left            =   1680
      Shape           =   1  'Square
      Top             =   3720
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   23
      Left            =   5040
      Shape           =   1  'Square
      Top             =   4200
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   22
      Left            =   4560
      Shape           =   1  'Square
      Top             =   4200
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   21
      Left            =   4080
      Shape           =   1  'Square
      Top             =   4200
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   20
      Left            =   3600
      Shape           =   1  'Square
      Top             =   4200
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   19
      Left            =   3120
      Shape           =   1  'Square
      Top             =   4200
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   18
      Left            =   2640
      Shape           =   1  'Square
      Top             =   4200
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   17
      Left            =   2160
      Shape           =   1  'Square
      Top             =   4200
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   16
      Left            =   1680
      Shape           =   1  'Square
      Top             =   4200
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   15
      Left            =   5040
      Shape           =   1  'Square
      Top             =   4680
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   14
      Left            =   4560
      Shape           =   1  'Square
      Top             =   4680
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   13
      Left            =   4080
      Shape           =   1  'Square
      Top             =   4680
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   12
      Left            =   3600
      Shape           =   1  'Square
      Top             =   4680
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   11
      Left            =   3120
      Shape           =   1  'Square
      Top             =   4680
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   10
      Left            =   2640
      Shape           =   1  'Square
      Top             =   4680
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   9
      Left            =   2160
      Shape           =   1  'Square
      Top             =   4680
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   8
      Left            =   1680
      Shape           =   1  'Square
      Top             =   4680
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   7
      Left            =   5040
      Shape           =   1  'Square
      Top             =   5160
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   6
      Left            =   4560
      Shape           =   1  'Square
      Top             =   5160
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   5
      Left            =   4080
      Shape           =   1  'Square
      Top             =   5160
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   4
      Left            =   3600
      Shape           =   1  'Square
      Top             =   5160
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   3
      Left            =   3120
      Shape           =   1  'Square
      Top             =   5160
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   2
      Left            =   2640
      Shape           =   1  'Square
      Top             =   5160
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   1
      Left            =   2160
      Shape           =   1  'Square
      Top             =   5160
      Width           =   492
   End
   Begin VB.Shape Square 
      BackColor       =   &H00C0C0FF&
      Height          =   492
      Index           =   0
      Left            =   1680
      Shape           =   1  'Square
      Top             =   5160
      Width           =   492
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private astrSquare(0 To 63) As String
Private aPiece(63) As CommandButton
Dim isWhiteTurn As Boolean
Dim intWhiteHorse1 As Integer
Dim intWhitePawn4 As Integer
Dim intWhiteRook1 As Integer
Dim intWhiteRook2 As Integer
Dim intWhiteBishop2 As Integer
Dim intWhiteKing As Integer
Dim intWhiteQueen As Integer
Dim IntBlackRook1 As Integer
Dim intSquareNumber As Integer
Dim intPreviousSquareNumber As Integer
Dim x As Integer
Dim y As Integer
Dim ActivePiece As CommandButton
Dim pieceIsBlocking1 As Boolean
Dim pieceIsBlocking2 As Boolean
Dim pieceIsBlocking3 As Boolean
Dim pieceIsBlocking4 As Boolean
Dim pieceIsBlocking5 As Boolean
Dim pieceIsBlocking6 As Boolean
Dim pieceIsBlocking7 As Boolean
Dim pieceIsBlocking8 As Boolean


'Linear Search Function Used to find position of Pieces in the array
Function FindItemIndex(ByRef astrSquare() As String, ByVal strSearchItem) As String
Dim intIndex As Integer

intIndex = LBound(astrSquare)
    Do While (astrSquare(intIndex) <> strSearchItem) And (intIndex < UBound(astrSquare))
    intIndex = intIndex + 1
Loop
If astrSquare(intIndex) = strSearchItem Then
    FindItemIndex = intIndex
Else
    FindItemIndex = -1
    End If
End Function

Private Sub BlackRook1_Click()



IntBlackRook1 = FindItemIndex(astrSquare(), "BlackRook1")
If isWhiteTurn = False Then
Call MovePiece(IntBlackRook1, "BlackRook1")
Call resetCommandButtons
intPreviousSquareNumber = IntBlackRook1

Dim u As Integer
For u = 1 To 7
If y + u <= 8 Then
    Call findSquare(x, y + u)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking1 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking1 = True
    End If
End If

If y - u >= 1 Then
    Call findSquare(x, y - u)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking2 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking2 = True
    End If
End If

If x + u <= 8 Then
    Call findSquare(x + u, y)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking3 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking3 = True
    End If
End If

If x - u >= 1 Then
    Call findSquare(x - u, y)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking4 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking4 = True
    End If
End If
Next u
ElseIf isWhiteTurn = True And cmdSquare(IntBlackRook1).Visible = True Then

    BlackRook1.Visible = False
    astrSquare(IntBlackRook1) = ActivePiece.Name
    astrSquare(intPreviousSquareNumber) = ""
    
    ActivePiece.Top = BlackRook1.Top
    ActivePiece.Left = BlackRook1.Left
    
    'Set whiteturn to false
    isWhiteTurn = False
    Call resetCommandButtons


End If
pieceIsBlocking1 = False
pieceIsBlocking2 = False
pieceIsBlocking3 = False
pieceIsBlocking4 = False

Set ActivePiece = Me.Controls("BlackRook1")
End Sub


Private Sub cmdReset_Click()
Call resetCommandButtons
Call resetPosition

End Sub

Private Sub WhiteBishop2_Click()
intWhiteBishop2 = FindItemIndex(astrSquare(), "WhiteBishop2")

If isWhiteTurn = True Then
Call MovePiece(intWhiteBishop2, "WhiteBishop2")
Call resetCommandButtons
intPreviousSquareNumber = intWhiteBishop2

Dim q As Integer
For q = 1 To 7

    If x + q <= 8 And y + q <= 8 And x + q > 0 And y + q > 0 Then
        Call findSquare(x + q, y + q)
        If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking1 = False Then
        cmdSquare(intSquareNumber).Visible = True
        Else
        pieceIsBlocking1 = True
        End If
    
    End If
    If x - q <= 8 And y + q <= 8 And x - q > 0 And y + q > 0 Then
        Call findSquare(x - q, y + q)
        If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking2 = False Then
            cmdSquare(intSquareNumber).Visible = True
        Else
        pieceIsBlocking2 = True
        End If
    End If
    If x + q <= 8 And y - q <= 8 And x + q > 0 And y - q > 0 Then
        Call findSquare(x + q, y - q)
        If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking3 = False Then
            cmdSquare(intSquareNumber).Visible = True
        Else
        pieceIsBlocking3 = True
        End If
    End If
    If x - q <= 8 And y - q <= 8 And x - q > 0 And y - q > 0 Then
        Call findSquare(x - q, y - q)
        If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking4 = False Then
            cmdSquare(intSquareNumber).Visible = True
        Else
        pieceIsBlocking4 = True
        End If
    End If
Next q

ElseIf isWhiteTurn = False And cmdSquare(intWhiteBishop2).Visible = True Then
 WhiteBishop2.Visible = False
    astrSquare(intWhiteBishop2) = ActivePiece.Name
    astrSquare(intPreviousSquareNumber) = ""
    
    ActivePiece.Top = WhiteBishop2.Top
    ActivePiece.Left = WhiteBishop2.Left
    
    'Set whiteturn to True
    isWhiteTurn = True
Call resetCommandButtons
End If
Set ActivePiece = Me.Controls("WhiteBishop2")
pieceIsBlocking1 = False
pieceIsBlocking2 = False
pieceIsBlocking3 = False
pieceIsBlocking4 = False
End Sub

Private Sub WhiteHorse1_Click()


    ' Find position in the array
    intWhiteHorse1 = FindItemIndex(astrSquare(), "WhiteHorse1")
    If isWhiteTurn = True Then
    ' Call the movepiece function
    Call MovePiece(intWhiteHorse1, "WhiteHorse1")
    Call resetCommandButtons
    intPreviousSquareNumber = intWhiteHorse1
    
    'Check if the new cordinates are on the board or inside (8,8)
    If x + 2 <= 8 And x + 2 >= 0 And y + 1 <= 8 And y + 1 >= 0 Then
    Call findSquare(x + 2, y + 1)
    cmdSquare(intSquareNumber).Visible = True
    End If
    
    If x - 2 <= 8 And x - 2 > 0 And y - 1 <= 8 And y - 1 > 0 Then
    Call findSquare(x - 2, y - 1)
    cmdSquare(intSquareNumber).Visible = True
    End If
    
    If x + 1 <= 8 And x + 1 > 0 And y - 2 <= 8 And y - 2 > 0 Then
    Call findSquare(x + 1, y - 2)
    cmdSquare(intSquareNumber).Visible = True
    End If
    
    If x - 1 <= 8 And x - 1 > 0 And y - 2 <= 8 And y - 2 > 0 Then
    Call findSquare(x - 1, y - 2)
    cmdSquare(intSquareNumber).Visible = True
    End If
    
    If x + 1 <= 8 And x + 1 > 0 And y + 2 <= 8 And y + 2 > 0 Then
    Call findSquare(x + 1, y + 2)
    cmdSquare(intSquareNumber).Visible = True
    End If
    
    If x - 1 <= 8 And x - 1 > 0 And y + 2 <= 8 And y + 2 > 0 Then
    Call findSquare(x - 1, y + 2)
    cmdSquare(intSquareNumber).Visible = True
    End If
    
    If x - 2 <= 8 And x - 2 > 0 And y + 1 <= 8 And y + 1 > 0 Then
    Call findSquare(x - 2, y + 1)
    cmdSquare(intSquareNumber).Visible = True
    End If
    
    If x + 2 <= 8 And x + 2 > 0 And y - 1 <= 8 And y - 1 > 0 Then
    Call findSquare(x + 2, y - 1)
    cmdSquare(intSquareNumber).Visible = True
    End If
    
ElseIf isWhiteTurn = False And cmdSquare(intWhiteHorse1).Visible = True Then
 WhiteHorse1.Visible = False
    astrSquare(intWhiteHorse1) = ActivePiece.Name
    astrSquare(intPreviousSquareNumber) = ""
    
    ActivePiece.Top = WhiteHorse1.Top
    ActivePiece.Left = WhiteHorse1.Left
    
    'Set whiteturn to True
    isWhiteTurn = True
Call resetCommandButtons
End If
    ' See if the button was already clicke
    Set ActivePiece = Me.Controls("WhiteHorse1")
    

End Sub


Private Sub WhiteKing_Click()


intWhiteKing = FindItemIndex(astrSquare(), "WhiteKing")

Call MovePiece(intWhiteKing, "WhiteKing")
Call resetCommandButtons
intPreviousSquareNumber = intWhiteKing

If x + 1 <= 8 Then
    Call findSquare(x + 1, y)
    cmdSquare(intSquareNumber).Visible = True
End If
If x - 1 > 0 Then
    Call findSquare(x - 1, y)
    cmdSquare(intSquareNumber).Visible = True
End If

If y + 1 <= 8 Then
    Call findSquare(x, y + 1)
    cmdSquare(intSquareNumber).Visible = True
End If

If y - 1 > 0 Then
    Call findSquare(x, y - 1)
    cmdSquare(intSquareNumber).Visible = True
End If

If y + 1 <= 8 And x + 1 <= 8 Then
    Call findSquare(x + 1, y + 1)
    cmdSquare(intSquareNumber).Visible = True
End If

If y - 1 > 0 And x - 1 > 0 Then
    Call findSquare(x - 1, y - 1)
    cmdSquare(intSquareNumber).Visible = True
End If

If y + 1 <= 8 And x - 1 > 0 Then
    Call findSquare(x - 1, y + 1)
    cmdSquare(intSquareNumber).Visible = True
End If

If y - 1 > 0 And x + 1 <= 8 Then
    Call findSquare(x + 1, y - 1)
    cmdSquare(intSquareNumber).Visible = True
End If
Set ActivePiece = Me.Controls("WhiteKing")
End Sub

Private Sub WhitePawn4_Click()



intWhitePawn4 = FindItemIndex(astrSquare(), "WhitePawn4")

If isWhiteTurn = True Then

Call MovePiece(intWhitePawn4, "WhitePawn4")
Call resetCommandButtons
intPreviousSquareNumber = intWhitePawn4

If intWhitePawn4 = 11 Then

    Call findSquare(x, y + 1)
    cmdSquare(intSquareNumber).Visible = True
    
    Call findSquare(x, y + 2)
    cmdSquare(intSquareNumber).Visible = True
    
ElseIf y < 8 Then
Call findSquare(x, y + 1)
cmdSquare(intSquareNumber).Visible = True
End If

'Make pawns able to take pieces
Call findSquare(x + 1, y + 1)
If astrSquare(intSquareNumber) <> "" Then
cmdSquare(intSquareNumber).Visible = True
End If
Call findSquare(x - 1, y + 1)
If astrSquare(intSquareNumber) <> "" Then
cmdSquare(intSquareNumber).Visible = True
End If
ElseIf isWhiteTurn = False And cmdSquare(intWhitePawn4).Visible = True Then
 WhitePawn4.Visible = False
    astrSquare(intWhitePawn4) = ActivePiece.Name
    astrSquare(intPreviousSquareNumber) = ""
    
    ActivePiece.Top = WhitePawn4.Top
    ActivePiece.Left = WhitePawn4.Left
    
    'Set whiteturn to True
    isWhiteTurn = True
Call resetCommandButtons
End If
Set ActivePiece = Me.Controls("WhitePawn4")
End Sub

Private Sub cmdSquare_Click(Index As Integer)

ActivePiece.Left = Square(Index).Left
ActivePiece.Top = Square(Index).Top
astrSquare(Index) = ActivePiece.Name
astrSquare(intPreviousSquareNumber) = ""

Label1.Caption = astrSquare(Index)
If isWhiteTurn = False Then
    isWhiteTurn = True
Else
    isWhiteTurn = False
End If
    Call resetCommandButtons
    
End Sub

Private Sub Form_Load()

'Reset all the positions of the pieces
Call resetPosition



pieceIsBlocking1 = False
pieceIsBlocking2 = False
pieceIsBlocking3 = False
pieceIsBlocking4 = False
pieceIsBlocking5 = False
pieceIsBlocking6 = False
pieceIsBlocking7 = False
pieceIsBlocking8 = False
End Sub
Sub MovePiece(ByRef intPiecePosition As Integer, ByRef strPiece As String)

'Find the position of the piece (cordinates)
Call FindPosition(intPiecePosition)


End Sub

Function FindPosition(ByRef intPiecePosition As Integer)

'Math equation to find the cordinates of any number square
y = (intPiecePosition \ 8) + 1
x = (intPiecePosition - ((y - 1) * 8)) + 1



End Function

Function resetPosition()
isWhiteTurn = True

'Reset all the positions in the array

'Insert for loop to make all 64 blank
Dim l As Integer
For l = 0 To 63
    astrSquare(l) = ""
Next l

'Assign each space of the board
astrSquare(0) = "WhiteRook1"
astrSquare(1) = "WhiteHorse1"
astrSquare(2) = ""
astrSquare(3) = "WhiteQueen"
astrSquare(4) = "WhiteKing"
astrSquare(5) = "WhiteBishop2"
astrSquare(6) = ""
astrSquare(7) = "WhiteRook2"
astrSquare(8) = ""
astrSquare(11) = "WhitePawn4"
astrSquare(56) = "BlackRook1"

Set aPiece(0) = Me.Controls("WhiteRook1")
Set aPiece(1) = Me.Controls("WhiteHorse1")
'Set aPiece(2) = Me.Controls("WhiteBishop1")
Set aPiece(3) = Me.Controls("WhiteQueen")
Set aPiece(4) = Me.Controls("WhiteKing")
Set aPiece(5) = Me.Controls("WhiteBishop2")
'Set aPiece(6) = Me.Controls("WhiteHorse2")
Set aPiece(7) = Me.Controls("WhiteRook2")
'Set aPiece(8) = Me.Controls("WhitePawn1")
'Set aPiece(9) = Me.Controls("WhitePawn2")
'Set aPiece(10) = Me.Controls("WhitePawn3")
Set aPiece(11) = Me.Controls("WhitePawn4")
'Set aPiece(12) = Me.Controls("WhitePawn5")
'Set aPiece(13) = Me.Controls("WhitePawn6")
'Set aPiece(14) = Me.Controls("WhitePawn7")
'Set aPiece(15) = Me.Controls("WhitePawn8")

'Set aPiece(48) = Me.Controls("BlackPawn1")
'Set aPiece(49) = Me.Controls("BlackPawn2")
'Set aPiece(50) = Me.Controls("BlackPawn3")
'Set aPiece(51) = Me.Controls("BlackPawn4")
'Set aPiece(52) = Me.Controls("BlackPawn5")
'Set aPiece(53) = Me.Controls("BlackPawn6")
'Set aPiece(54) = Me.Controls("BlackPawn7")
'Set aPiece(55) = Me.Controls("BlackPawn8")
Set aPiece(56) = Me.Controls("BlackRook1")
'Set aPiece(57) = Me.Controls("BlackHorse1")
'Set aPiece(58) = Me.Controls("BlackBishop1")
'Set aPiece(59) = Me.Controls("BlackQueen")
'Set aPiece(60) = Me.Controls("BlackKing")
'Set aPiece(61) = Me.Controls("BlackBishop1")
'Set aPiece(62) = Me.Controls("BlackHorse2")
'Set aPiece(63) = Me.Controls("BlackRook2")
Dim p As Integer
For p = 0 To 63
    If IsEmpty(aPiece(p)) = False Then
        aPiece(p).Left = cmdSquare(p).Left
        aPiece(p).Top = cmdSquare(p).Top
    
    End If
Next p
End Function

Function findSquare(x, y)

'FInd the square number of a cordinate
intSquareNumber = (y - 1) * 8 + (x - 1)



End Function

Function resetCommandButtons()
'Make all commandbuttons not visible
Dim I As Integer
For I = 0 To 63
    cmdSquare(I).Visible = False
Next I

End Function


Private Sub WhiteQueen_Click()

intWhiteQueen = FindItemIndex(astrSquare(), "WhiteQueen")

'See whos turn it is
If isWhiteTurn = True Then
Call resetCommandButtons
Call MovePiece(intWhiteQueen, "WhiteQueen")

intPreviousSquareNumber = intWhiteQueen


'Vertical / Horizontal Movement
Dim u As Integer
For u = 1 To 7
    If y + u <= 8 Then
        Call findSquare(x, y + u)
        If astrSquare(intSquareNumber) <> "" Then
            cmdSquare(intSquareNumber).Visible = True
        End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking1 = False Then
            cmdSquare(intSquareNumber).Visible = True
        Else
        pieceIsBlocking1 = True
        End If
    End If

    If y - u >= 1 Then
        Call findSquare(x, y - u)
        If astrSquare(intSquareNumber) <> "" Then
            cmdSquare(intSquareNumber).Visible = True
        End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking2 = False Then
            cmdSquare(intSquareNumber).Visible = True
        Else
        pieceIsBlocking2 = True
        End If
    End If

    If x + u <= 8 Then
        Call findSquare(x + u, y)
        If astrSquare(intSquareNumber) <> "" Then
            cmdSquare(intSquareNumber).Visible = True
        End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking3 = False Then
            cmdSquare(intSquareNumber).Visible = True
        Else
        pieceIsBlocking3 = True
        End If
    End If
    
    If x - u >= 1 Then
        Call findSquare(x - u, y)
        If astrSquare(intSquareNumber) <> "" Then
            cmdSquare(intSquareNumber).Visible = True
       End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking4 = False Then
            cmdSquare(intSquareNumber).Visible = True
        Else
        pieceIsBlocking4 = True
        End If
    End If
    
    ' Bishop Like movement
     If x + u <= 8 And y + u <= 8 And x + u > 0 And y + u > 0 Then
            Call findSquare(x + u, y + u)
            If astrSquare(intSquareNumber) <> "" Then
            cmdSquare(intSquareNumber).Visible = True
               End If
            If astrSquare(intSquareNumber) = "" And pieceIsBlocking5 = False Then
            cmdSquare(intSquareNumber).Visible = True
            Else
            pieceIsBlocking5 = True
            End If
        End If
        
    If x - u <= 8 And y + u <= 8 And x - u > 0 And y + u > 0 Then
        Call findSquare(x - u, y + u)
        If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
            End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking6 = False Then
            cmdSquare(intSquareNumber).Visible = True
        Else
        pieceIsBlocking6 = True
        End If
    End If
    
    If x + u <= 8 And y - u <= 8 And x + u > 0 And y - u > 0 Then
        Call findSquare(x + u, y - u)
            If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
            End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking7 = False Then
                cmdSquare(intSquareNumber).Visible = True
        Else
            pieceIsBlocking7 = True
        End If
    End If
    
    If x - u <= 8 And y - u <= 8 And x - u > 0 And y - u > 0 Then
        Call findSquare(x - u, y - u)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
        If astrSquare(intSquareNumber) = "" And pieceIsBlocking8 = False Then
            cmdSquare(intSquareNumber).Visible = True
        Else
        pieceIsBlocking8 = True
        End If
    End If
Next

ElseIf isWhiteTurn = False And cmdSquare(intWhiteQueen).Visible = True Then
    'Check if the piece can actualy take the White Queen
    WhiteQueen.Visible = False
    astrSquare(intWhiteQueen) = ActivePiece.Name
    astrSquare(intPreviousSquareNumber) = ""
    
    ActivePiece.Top = WhiteQueen.Top
    ActivePiece.Left = WhiteQueen.Left
    
    'Set whiteturn to true
    isWhiteTurn = True
    Call resetCommandButtons
End If

pieceIsBlocking1 = False
pieceIsBlocking2 = False
pieceIsBlocking3 = False
pieceIsBlocking4 = False
pieceIsBlocking5 = False
pieceIsBlocking6 = False
pieceIsBlocking7 = False
pieceIsBlocking8 = False
Set ActivePiece = Me.Controls("WhiteQueen")

End Sub
Private Sub WhiteRook1_Click()


intWhiteRook1 = FindItemIndex(astrSquare(), "WhiteRook1")

If isWhiteTurn = True Then

Call resetCommandButtons
Call MovePiece(intWhiteRook1, "WhiteRook1")

intPreviousSquareNumber = intWhiteRook1

Dim u As Integer
For u = 1 To 7
If y + u <= 8 Then
    Call findSquare(x, y + u)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking1 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking1 = True
    End If
End If

If y - u >= 1 Then
    Call findSquare(x, y - u)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking2 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking2 = True
    End If
End If

If x + u <= 8 Then
    Call findSquare(x + u, y)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking3 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking3 = True
    End If
End If

If x - u >= 1 Then
    Call findSquare(x - u, y)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking4 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking4 = True
    End If
End If
Next u
ElseIf isWhiteTurn = False And cmdSquare(intWhiteRook1).Visible = True Then

    WhiteRook1.Visible = False
    astrSquare(intWhiteRook1) = ActivePiece.Name
    astrSquare(intPreviousSquareNumber) = ""
    
    ActivePiece.Top = WhiteRook1.Top
    ActivePiece.Left = WhiteRook1.Left
    
    'Set whiteturn to True
    isWhiteTurn = True
Call resetCommandButtons

End If

pieceIsBlocking1 = False
pieceIsBlocking2 = False
pieceIsBlocking3 = False
pieceIsBlocking4 = False

Set ActivePiece = Me.Controls("WhiteRook1")
End Sub

Private Sub WhiteRook2_Click()


intWhiteRook2 = FindItemIndex(astrSquare(), "WhiteRook2")
If isWhiteTurn = True Then
Call MovePiece(intWhiteRook2, "WhiteRook2")
Call resetCommandButtons
intPreviousSquareNumber = intWhiteRook2

Dim u As Integer
For u = 1 To 7
If y + u <= 8 Then
    Call findSquare(x, y + u)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking1 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking1 = True
    End If
End If

If y - u >= 1 Then
    Call findSquare(x, y - u)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking2 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking2 = True
    End If
End If
    
If x + u <= 8 Then
    Call findSquare(x + u, y)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking3 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking3 = True
    End If
End If

If x - u >= 1 Then
    Call findSquare(x - u, y)
    If astrSquare(intSquareNumber) <> "" Then
        cmdSquare(intSquareNumber).Visible = True
    End If
    If astrSquare(intSquareNumber) = "" And pieceIsBlocking4 = False Then
        cmdSquare(intSquareNumber).Visible = True
    Else
    pieceIsBlocking4 = True
    End If
End If
Next u

ElseIf isWhiteTurn = False And cmdSquare(intWhiteRook2).Visible = True Then

WhiteRook2.Visible = False
    astrSquare(intWhiteRook2) = ActivePiece.Name
    astrSquare(intPreviousSquareNumber) = ""
    
    ActivePiece.Top = WhiteRook2.Top
    ActivePiece.Left = WhiteRook2.Left
    
    'Set whiteturn to True
    isWhiteTurn = True
Call resetCommandButtons

End If
pieceIsBlocking1 = False
pieceIsBlocking2 = False
pieceIsBlocking3 = False
pieceIsBlocking4 = False

Set ActivePiece = Me.Controls("WhiteRook2")

End Sub
