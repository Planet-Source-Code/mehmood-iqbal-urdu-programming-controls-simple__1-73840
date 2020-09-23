VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form UPC 
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame15 
      Caption         =   "Urdu Text Box :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4920
      TabIndex        =   37
      Top             =   840
      Width           =   3855
      Begin MSForms.TextBox TextBox1 
         Height          =   1335
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Textbox1"
         Top             =   240
         Width           =   3615
         VariousPropertyBits=   -1394587621
         ForeColor       =   12582912
         ScrollBars      =   3
         Size            =   "6376;2355"
         SpecialEffect   =   3
         FontName        =   "Nafees Nastaleeq"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Urdu Command Buttons :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   2520
      Width           =   8655
      Begin MSForms.CommandButton CommandButton1 
         Height          =   615
         Left            =   6480
         TabIndex        =   36
         ToolTipText     =   "CommandButton1"
         Top             =   240
         Width           =   2055
         ForeColor       =   12582912
         Size            =   "3625;1085"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CommandButton2 
         Height          =   615
         Left            =   4320
         TabIndex        =   35
         ToolTipText     =   "CommandButton2"
         Top             =   240
         Width           =   2175
         ForeColor       =   12582912
         Size            =   "3836;1085"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CommandButton3 
         Height          =   615
         Left            =   2160
         TabIndex        =   34
         ToolTipText     =   "CommandButton3"
         Top             =   240
         Width           =   2175
         ForeColor       =   12582912
         Size            =   "3836;1085"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CommandButton4 
         Height          =   615
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "CommansButton4"
         Top             =   240
         Width           =   2055
         ForeColor       =   12582912
         Size            =   "3625;1085"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Urdu List Box :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   8880
      TabIndex        =   31
      Top             =   840
      Width           =   2775
      Begin MSForms.ListBox ListBox1 
         Height          =   2175
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "Listbox1"
         Top             =   240
         Width           =   2535
         ForeColor       =   12582912
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "4471;3186"
         MatchEntry      =   0
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Urdu Frames :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   26
      Top             =   6480
      Width           =   11055
      Begin MSForms.Frame Frame4 
         Height          =   615
         Left            =   120
         OleObjectBlob   =   "Form1.frx":0A02
         TabIndex        =   27
         ToolTipText     =   "Frame4"
         Top             =   240
         Width           =   2535
      End
      Begin MSForms.Frame Frame3 
         Height          =   615
         Left            =   2760
         OleObjectBlob   =   "Form1.frx":141A
         TabIndex        =   28
         ToolTipText     =   "Frame3"
         Top             =   240
         Width           =   2655
      End
      Begin MSForms.Frame Frame2 
         Height          =   615
         Left            =   5520
         OleObjectBlob   =   "Form1.frx":1E32
         TabIndex        =   29
         ToolTipText     =   "Frame2"
         Top             =   240
         Width           =   2655
      End
      Begin MSForms.Frame Frame1 
         Height          =   615
         Left            =   8280
         OleObjectBlob   =   "Form1.frx":284A
         TabIndex        =   30
         ToolTipText     =   "Frame1"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Urdu MutiPage :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      TabIndex        =   24
      Top             =   4800
      Width           =   5775
      Begin MSForms.MultiPage MultiPage1 
         Height          =   855
         Left            =   120
         OleObjectBlob   =   "Form1.frx":3262
         TabIndex        =   25
         ToolTipText     =   "MultiPage1"
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Urdu Tab Strip :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5880
      TabIndex        =   22
      Top             =   3600
      Width           =   5775
      Begin MSForms.TabStrip TabStrip1 
         Height          =   735
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Tabstrip1"
         Top             =   240
         Width           =   5535
         ListIndex       =   0
         ForeColor       =   12582912
         Size            =   "9763;1296"
         Items           =   "Tab1;Tab2;"
         TipStrings      =   ";;"
         Names           =   "Tab1;Tab2;"
         NewVersion      =   -1  'True
         TabsAllocated   =   2
         Tags            =   ";;"
         TabData         =   2
         Accelerator     =   ";;"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
         TabState        =   "3;3"
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Urdu Option Boxes :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3000
      TabIndex        =   17
      Top             =   3480
      Width           =   2775
      Begin MSForms.OptionButton OptionButton1 
         Height          =   495
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "OptionButton1"
         Top             =   240
         Width           =   2535
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   5
         Size            =   "4471;873"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.OptionButton OptionButton2 
         Height          =   495
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "OptionButton2"
         Top             =   960
         Width           =   2535
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   5
         Size            =   "4471;873"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.OptionButton OptionButton3 
         Height          =   495
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "OptionButton3"
         Top             =   1680
         Width           =   2535
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   5
         Size            =   "4471;873"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.OptionButton OptionButton4 
         Height          =   495
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "OptionButton4"
         Top             =   2400
         Width           =   2535
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   5
         Size            =   "4471;873"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Urdu Chech Boxes :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   2775
      Begin MSForms.CheckBox CheckBox1 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Checkbox1"
         Top             =   240
         Width           =   2535
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "4471;873"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CheckBox CheckBox2 
         Height          =   495
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Checkbox2"
         Top             =   840
         Width           =   2535
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "4471;873"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CheckBox CheckBox3 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Checkbox3"
         Top             =   1560
         Width           =   2535
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "4471;873"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CheckBox CheckBox4 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Checkbox4"
         Top             =   2280
         Width           =   2535
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "4471;873"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Urdu Combo Box :"
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
      Left            =   7440
      TabIndex        =   10
      Top             =   0
      Width           =   2895
      Begin MSForms.ComboBox ComboBox1 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Combobox1"
         Top             =   240
         Width           =   2655
         VariousPropertyBits=   746604571
         ForeColor       =   12582912
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4683;873"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Urdu Toggle Buttons :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4695
      Begin MSForms.ToggleButton ToggleButton1 
         Height          =   615
         Left            =   2400
         TabIndex        =   9
         ToolTipText     =   "ToggleButton1"
         Top             =   240
         Width           =   2175
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   6
         Size            =   "3836;1085"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ToggleButton ToggleButton2 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "ToggleButton2"
         Top             =   240
         Width           =   2175
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   6
         Size            =   "3836;1085"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ToggleButton ToggleButton3 
         Height          =   615
         Left            =   2400
         TabIndex        =   7
         ToolTipText     =   "ToggleButton3"
         Top             =   960
         Width           =   2175
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   6
         Size            =   "3836;1085"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ToggleButton ToggleButton4 
         Height          =   615
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "ToggleButton4"
         Top             =   960
         Width           =   2175
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   6
         Size            =   "3836;1085"
         Value           =   "0"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Urdu Labels :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin MSForms.Label Label1 
         Height          =   375
         Left            =   5400
         TabIndex        =   4
         ToolTipText     =   "Label1"
         Top             =   240
         Width           =   1335
         ForeColor       =   12582912
         Size            =   "2355;661"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label2 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         ToolTipText     =   "Label2"
         Top             =   240
         Width           =   1815
         ForeColor       =   12582912
         Size            =   "3201;661"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label3 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         ToolTipText     =   "Label3"
         Top             =   240
         Width           =   1695
         ForeColor       =   12582912
         Size            =   "2990;661"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label4 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Label4"
         Top             =   240
         Width           =   1575
         ForeColor       =   12582912
         Size            =   "2778;661"
         FontName        =   "Nafees Naskh v2.01"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   10680
      Picture         =   "Form1.frx":4A7A
      ToolTipText     =   "MD's Developments Logo"
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "UPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''
'' First Time Uploaded on the PlanetSourceCode.  ''
'' There were not Urdu Programming Controls      ''
'' available on the whole Internet world, but    ''
'' first time i'm introducing these controls.    ''
'' These controls maybe need more attention but  ''
'' i think it is a great gift for Urdu Language  ''
'' Programmers. These controls with enhanced     ''
'' functions will be uploaded later on the PSC.  ''
'' From these Controls you can understand the    ''
'' basic concept of Urdu Controls that how can   ''
'' you make & use all programming controls in VB6''
'' applications. You can edit & use these        ''
'' controls as you can. Must remember to give    ''
'' Feedback at my E-Mail Address.                ''
'' Also give vote & rating for this code on PSC. ''
'' Thank You.                                    ''
''                                               ''
'' Regards : Muhammad Mehmood Iqbal              ''
'' Email Address : ME_IQ_TM@Yahoo.Com            ''
'''''''''''''''''''''''''''''''''''''''''''''''''''

'Global Veriables For Whole Project
Dim Conn As ADODB.Connection
Dim Rec As ADODB.Recordset

Private Sub Form_Load()

'Database Path & Connection
Set Conn = New Connection
With Conn
  .ConnectionString = App.Path & "\UDB.mdb"
  .Provider = "Microsoft.Jet.OLEDB.4.0"
  .Open
End With

'Database Perameter Set & Table Selection
Set Rec = New Recordset
  With Rec
   .ActiveConnection = Conn
   .LockType = adLockOptimistic
   .CursorType = adOpenKeyset
   .Open "UD"             ' Table Name
  End With

'Adding 2 More Tabs in TabStrip
TabStrip1.Tabs.Add
TabStrip1.Tabs.Add

'Adding 2 More Pages in MultiPage
MultiPage1.Pages.Add
MultiPage1.Pages.Add

'Initializing Controls
Call Set_Label_Caption
Call Text_Box_Text
Call Command_Button_Captions
Call Check_Box_Captions
Call Option_Box_Captions
Call List_Box_Items
Call Combo_Box_Items
Call Frame_Captions
Call Toggle_Button_Captions
Call Tab_Strip_Captions
Call Multi_Page_Captions
Call Form_Captions

'Closing Database Connections
Rec.Close
Conn.Close

'Removing Veriables from Memory
Set Rec = Nothing
Set Conn = Nothing

End Sub

Private Sub Set_Label_Caption()

'Setting Label's Captions
Label1.Caption = Rec(1)
Rec.MoveNext
Label2.Caption = Rec(1)
Rec.MoveNext
Label3.Caption = Rec(1)
Rec.MoveNext
Label4.Caption = Rec(1)

End Sub

Private Sub Text_Box_Text()

'Setting TextBox's Text
Rec.MoveFirst
TextBox1.Text = Rec(2)

End Sub

Private Sub Command_Button_Captions()

'Setting Command Button's Captions
CommandButton1.Caption = Rec(3)
Rec.MoveNext
CommandButton2.Caption = Rec(3)
Rec.MoveNext
CommandButton3.Caption = Rec(3)
Rec.MoveNext
CommandButton4.Caption = Rec(3)

End Sub

Private Sub Check_Box_Captions()

'Setting CheckBoxes Captions
Rec.MoveFirst
CheckBox1.Caption = Rec(4)
Rec.MoveNext
CheckBox2.Caption = Rec(4)
Rec.MoveNext
CheckBox3.Caption = Rec(4)
Rec.MoveNext
CheckBox4.Caption = Rec(4)

End Sub

Private Sub Option_Box_Captions()

'Setting Option Button's Captions
Rec.MoveFirst
OptionButton1.Caption = Rec(5)
Rec.MoveNext
OptionButton2.Caption = Rec(5)
Rec.MoveNext
OptionButton3.Caption = Rec(5)
Rec.MoveNext
OptionButton4.Caption = Rec(5)

End Sub

Private Sub List_Box_Items()

'Setting ListBox Items, (4 items)
Rec.MoveFirst
ListBox1.AddItem Rec(6)
Rec.MoveNext
ListBox1.AddItem Rec(6)
Rec.MoveNext
ListBox1.AddItem Rec(6)
Rec.MoveNext
ListBox1.AddItem Rec(6)

End Sub

Private Sub Combo_Box_Items()

'Setting ComboBox Items, (4 items)
Rec.MoveFirst
ComboBox1.AddItem Rec(7)
Rec.MoveNext
ComboBox1.AddItem Rec(7)
Rec.MoveNext
ComboBox1.AddItem Rec(7)
Rec.MoveNext
ComboBox1.AddItem Rec(7)

End Sub

Private Sub Frame_Captions()

'Setting Frame's Captions
Rec.MoveFirst
Frame1.Caption = Rec(8)
Rec.MoveNext
Frame2.Caption = Rec(8)
Rec.MoveNext
Frame3.Caption = Rec(8)
Rec.MoveNext
Frame4.Caption = Rec(8)

End Sub

Private Sub Toggle_Button_Captions()

'Setting Toggle Button's Caption
Rec.MoveFirst
ToggleButton1.Caption = Rec(10)
Rec.MoveNext
ToggleButton2.Caption = Rec(10)
Rec.MoveNext
ToggleButton3.Caption = Rec(10)
Rec.MoveNext
ToggleButton4.Caption = Rec(10)

End Sub

Private Sub Tab_Strip_Captions()

'Setting TabStrip Tab's Captions
Rec.MoveFirst
TabStrip1.Tabs(0).Caption = Rec(11)
Rec.MoveNext
TabStrip1.Tabs(1).Caption = Rec(11)
Rec.MoveNext
TabStrip1.Tabs(2).Caption = Rec(11)
Rec.MoveNext
TabStrip1.Tabs(3).Caption = Rec(11)

End Sub

Private Sub Multi_Page_Captions()

'Setting MutiPage Page's Captions
Rec.MoveFirst
MultiPage1.Pages(0).Caption = Rec(12)
Rec.MoveNext
MultiPage1.Pages(1).Caption = Rec(12)
Rec.MoveNext
MultiPage1.Pages(2).Caption = Rec(12)
Rec.MoveNext
MultiPage1.Pages(3).Caption = Rec(12)

End Sub

Private Sub Form_Captions()

'Setting Form's Caption
Me.Caption = " Yeh Genuine Urdu Controls Ki Misaal Hai ! "

End Sub

Private Sub TextBox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyDown-Event behaviours for  ''
        '' Enter, Space, Delete & Tab keys to set      ''
        '' Behavior in Textbox1.Text, keys will behave ''
        '' as Normal Text writing behavior.            ''
        ''             ME_IQ_TM@Yahoo.Com              ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''
     
        'Space Key Behavior
        If KeyCode = 32 Then
        UniCode = &H20
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        KeyCode = 0
        
        'Enter Key Behavior
        ElseIf KeyCode = 13 Then
        UniCode = &HA
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        KeyCode = 0
        
        'Horizontal Tab Behavior
        ElseIf KeyCode = 9 Then
        UniCode = &H9
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        KeyCode = 0
        
        'Delete Key Behavior
        ElseIf KeyCode = 127 Then
        UniCode = &H7F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        KeyCode = 0
        
        End If
        
        'This Function Got End There

End Sub

Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)

        ''''''''''''''''''''''''''''''''''''''''''''''''
        ''                                             ''
        '' There are the KeyPress-Event behaviours for ''
        '' Alfabatic, Numaric & Symbolic keys to write ''
        '' Urdu. I've tried to make it near with Urdu  ''
        '' Phonetic Keyboard Layout.                   ''
        ''                                             ''
        ''              ME_IQ_TM@Yahoo.Com             ''
        ''                                             ''
        '''''''''''''''''''''''''''''''''''''''''''''''''
       
If ModeValue = False Then

        'For Small Letter's Behaviors

        'a Key Behavior
        If KeyAscii = 97 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H627
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'b Key Behavior
        ElseIf KeyAscii = 98 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H628
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'c Key Behavior
        ElseIf KeyAscii = 99 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H686
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'd Key Behavior
        ElseIf KeyAscii = 100 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'e Key Behavior
        ElseIf KeyAscii = 101 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H639
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'f Key Behavior
        ElseIf KeyAscii = 102 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H641
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'g Key Behavior
        ElseIf KeyAscii = 103 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6AF
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'h Key Behavior
        ElseIf KeyAscii = 104 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6BE
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'i Key Behavior
        ElseIf KeyAscii = 105 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6CC
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'j Key Behavior
        ElseIf KeyAscii = 106 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'k Key Behavior
        ElseIf KeyAscii = 107 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6A9
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'l Key Behavior
        ElseIf KeyAscii = 108 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H644
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'm Key Behavior
        ElseIf KeyAscii = 109 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H645
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'n Key Behavior
        ElseIf KeyAscii = 110 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H646
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'o Key Behavior
        ElseIf KeyAscii = 111 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6C1
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'p Key Behavior
        ElseIf KeyAscii = 112 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H67E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'q Key Behavior
        ElseIf KeyAscii = 113 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H642
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'r Key Behavior
        ElseIf KeyAscii = 114 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H631
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        's Key Behavior
        ElseIf KeyAscii = 115 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H633
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        't Key Behavior
        ElseIf KeyAscii = 116 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'u Key Behavior
        ElseIf KeyAscii = 117 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H621
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'v Key Behavior
        ElseIf KeyAscii = 118 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H637
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'w Key Behavior
        ElseIf KeyAscii = 119 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H648
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'x Key Behavior
        ElseIf KeyAscii = 120 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H634
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'y Key Behavior
        ElseIf KeyAscii = 121 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6D2
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'z Key Behavior
        ElseIf KeyAscii = 122 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H632
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        
        ' For Capital Latter's Behaviors
        
        'A Key Behavior
        ElseIf KeyAscii = 65 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H622
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'B Key Behavior
        ElseIf KeyAscii = 66 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFBB0
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'C Key Behavior
        ElseIf KeyAscii = 67 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'D Key Behavior
        ElseIf KeyAscii = 68 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H688
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'E Key Behavior
        ElseIf KeyAscii = 69 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H650
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'F Key Behavior
        ElseIf KeyAscii = 70 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H652
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'G Key Behavior
        ElseIf KeyAscii = 71 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H63A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'H Key Behavior
        ElseIf KeyAscii = 72 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'I Key Behavior
        ElseIf KeyAscii = 73 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H649
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'J Key Behavior
        ElseIf KeyAscii = 74 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H636
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'K Key Behavior
        ElseIf KeyAscii = 75 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H62E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'L Key Behavior
        ElseIf KeyAscii = 76 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFEFB
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'M Key Behavior
        ElseIf KeyAscii = 77 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H66B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'N Key Behavior
        ElseIf KeyAscii = 78 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H6BA
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'O Key Behavior
        ElseIf KeyAscii = 79 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H629
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'P Key Behavior
        ElseIf KeyAscii = 80 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'Q Key Behavior
        ElseIf KeyAscii = 81 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H626
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'R Key Behavior
        ElseIf KeyAscii = 82 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H691
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'S Key Behavior
        ElseIf KeyAscii = 83 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H635
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'T Key Behavior
        ElseIf KeyAscii = 84 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H679
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'U Key Behavior
        ElseIf KeyAscii = 85 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H626
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'V Key Behavior
        ElseIf KeyAscii = 86 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H638
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'W Key Behavior
        ElseIf KeyAscii = 87 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H624
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'Z Key Behavior
        ElseIf KeyAscii = 88 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H698
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'Y Key Behavior
        ElseIf KeyAscii = 89 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFBAF
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'Z Key Behavior
        ElseIf KeyAscii = 90 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H630
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        
        'For Numaric Key's Behaviors
        
        '0 Key Behavior
        ElseIf KeyAscii = 48 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H660
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '1 Key Behavior
        ElseIf KeyAscii = 49 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H661
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '2 Key Behavior
        ElseIf KeyAscii = 50 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H662
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '3 Key Behavior
        ElseIf KeyAscii = 51 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H663
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '4 Key Behavior
        ElseIf KeyAscii = 52 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H664
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '5 Key Behavior
        ElseIf KeyAscii = 53 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H665
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '6 Key Behavior
        ElseIf KeyAscii = 54 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H666
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '7 Key Behavior
        ElseIf KeyAscii = 55 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H667
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '8 Key Behavior
        ElseIf KeyAscii = 56 Or TextBox1.SelText <> "" Then
        UniCode = &H668
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '9 Key Behavior
        ElseIf KeyAscii = 57 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H669
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        ' Numaric Keys with 'Shift' Behavior
        
        ') Key Behavior
        ElseIf KeyAscii = 41 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFD3F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '! Key Behavior
        ElseIf KeyAscii = 33 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H21
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '@ Key Behavior
        ElseIf KeyAscii = 64 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H40
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '# Key Behavior
        ElseIf KeyAscii = 35 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H23
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '$ Key Behavior
        ElseIf KeyAscii = 36 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H24
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '% Key Behavior
        ElseIf KeyAscii = 37 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H66A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '^ Key Behavior
        ElseIf KeyAscii = 94 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '& Key Behavior
        ElseIf KeyAscii = 38 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H26
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '* Key Behavior
        ElseIf KeyAscii = 42 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H66D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '( Key Behavior
        ElseIf KeyAscii = 40 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &HFD3E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        
        'For Special Characters
        
        'Symbols
        
        '? Key Behavior
        ElseIf KeyAscii = 63 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H61F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '/ Key Behavior
        ElseIf KeyAscii = 47 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        ', Key Behavior
        ElseIf KeyAscii = 44 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H60C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '. Key Behavior
        ElseIf KeyAscii = 46 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H640
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '_ Key Behavior
        ElseIf KeyAscii = 95 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '- Key Behavior
        ElseIf KeyAscii = 45 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '+ Key Behavior
        ElseIf KeyAscii = 43 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '= Key Behavior
        ElseIf KeyAscii = 61 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H3D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        ': Key Behavior
        ElseIf KeyAscii = 58 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H3A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '; Key Behavior
        ElseIf KeyAscii = 59 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H201C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '< Key Behavior
        ElseIf KeyAscii = 60 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '> Key Behavior
        ElseIf KeyAscii = 62 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H650
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '{ Key Behavior
        ElseIf KeyAscii = 123 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2018
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '} Key Behavior
        ElseIf KeyAscii = 125 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2019
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '[ Key Behavior
        ElseIf KeyAscii = 91 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '] Key Behavior
        ElseIf KeyAscii = 93 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '| Key Behavior
        ElseIf KeyAscii = 124 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H7C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '\ Key Behavior
        ElseIf KeyAscii = 92 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H5C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '~ Key Behavior
        ElseIf KeyAscii = 126 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '` Key Behavior
        ElseIf KeyAscii = 96 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H64D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '" Key Behavior
        ElseIf KeyAscii = 34 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H2190
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        '' Key Behavior
        ElseIf KeyAscii = 39 Or TextBox1.SelText <> "" Then
        TextBox1.SelText = ""
        UniCode = &H201D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        End If
        KeyAscii = 0
        End If

        'This Function Got End There

End Sub

'Whole Project Got End There
