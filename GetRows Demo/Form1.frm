VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Height          =   2010
      ItemData        =   "Form1.frx":0000
      Left            =   3180
      List            =   "Form1.frx":0002
      TabIndex        =   7
      Top             =   540
      Width           =   1350
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   525
      Left            =   3180
      TabIndex        =   6
      Top             =   2580
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   525
      Left            =   1650
      TabIndex        =   3
      Top             =   2580
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   2580
      Width           =   1350
   End
   Begin VB.ListBox List2 
      Height          =   2010
      ItemData        =   "Form1.frx":0004
      Left            =   1650
      List            =   "Form1.frx":0006
      TabIndex        =   1
      Top             =   540
      Width           =   1350
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "Form1.frx":0008
      Left            =   120
      List            =   "Form1.frx":000A
      TabIndex        =   0
      Top             =   540
      Width           =   1350
   End
   Begin MSAdodcLib.Adodc a 
      Height          =   330
      Left            =   3420
      Top             =   240
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fill Two Different Recordsets gotten by Getrows and Added "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3210
      TabIndex        =   11
      Top             =   30
      Width           =   2220
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fill with Recordset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1650
      TabIndex        =   10
      Top             =   30
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fill with Getrows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   150
      TabIndex        =   9
      Top             =   30
      Width           =   1305
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3180
      TabIndex        =   8
      Top             =   3150
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1650
      TabIndex        =   5
      Top             =   3150
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   3150
      Width           =   1350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'This simple demo demonstrates the speed of GetRows Method Compared to
'movenext method of ADO. The first and second listboxes show this while
'the third demonstates yet another example and use of getrows, joining
'arrays of data from a recordset. You can compare the speed of this method too.
'************** By Suresh Batta *******************************************
'==============================================================================
Private Sub Command1_Click()
'declare variables
Dim rsTemp As Variant, lngUBound1 As Long, lngUBound2 As Long
Dim iRow As Long, iCol As Long, lngCounter As Long
Dim tmrStart As Long

'clear listbox
List1.Clear

'start timer
tmrStart = Timer

'get the recordset
With a
    .RecordSource = "select * from countries"
    .Refresh
End With

'get into var
rsTemp = a.Recordset.GetRows
lngUBound1 = UBound(rsTemp, 2)
lngUBound2 = UBound(rsTemp, 1)

'loop for 5 times
For lngCounter = 0 To 4
    For iRow = 0 To lngUBound1
        For iCol = 0 To lngUBound2
            List1.AddItem rsTemp(iCol, iRow)
        Next iCol
    Next iRow
Next lngCounter

'calculate and show time taken
Label1 = FormatNumber((Timer - tmrStart), 2) & " secs"

End Sub

Private Sub Command2_Click()
'declare variables
Dim lngCounter1 As Long, lngCounter2 As Long, tmrStart As Long

'clear listbox
List2.Clear

'start timer
tmrStart = Timer

'loop for 5 times
For lngCounter1 = 0 To 4
    'get the recordset
    With a
        .RecordSource = "select * from countries"
        .Refresh
    End With
    With a
        For lngCounter2 = 1 To .Recordset.RecordCount
            List2.AddItem .Recordset.Fields(0)
            List2.AddItem .Recordset.Fields(1)
            .Recordset.MoveNext
        Next lngCounter2
    End With
Next lngCounter1

'calculate and show time taken
Label2 = FormatNumber((Timer - tmrStart), 2) & " secs"

End Sub

'===========================================
'the code below shows how u can join two arrays
Private Sub Command3_Click()
'declare variables
Dim rsTemp As Variant, rstemp1() As Variant, lngLBound As Long, lngUBound As Long
Dim iRow As Long, iCol As Long
Dim tmrStart As Long

'clear listbox
List3.Clear

'start timer
tmrStart = Timer

'get the first recordset
With a
    .RecordSource = "select * from countries"
    .Refresh
End With
'get into var
rsTemp = a.Recordset.GetRows
lngLBound = a.Recordset.RecordCount - 1

'get the next recordset
With a
    .RecordSource = "select * from buscat"
    .Refresh
End With
'get into var
rstemp1 = a.Recordset.GetRows

'expand var
ReDim Preserve rsTemp(0 To 1, 0 To UBound(rsTemp, 2) + UBound(rstemp1, 2))

'loop thru to include second recordset
For iRow = 0 To UBound(rstemp1, 2)
    For iCol = 0 To UBound(rstemp1, 1)
        rsTemp(iCol, iRow + lngLBound) = rstemp1(iCol, iRow)
    Next iCol
Next iRow

'fill listbox
For iRow = 0 To UBound(rsTemp, 2)
    For iCol = 0 To UBound(rsTemp, 1)
        List3.AddItem rsTemp(iCol, iRow)
    Next iCol
Next iRow

'calculate and show time taken
Label3 = FormatNumber((Timer - tmrStart), 2) & " secs"

'free memory
Erase rstemp1
Erase rsTemp

End Sub

Private Sub Form_Load()
With a
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & _
                        App.Path & ChrW$(92) & "db.mdb" & ";Persist Security Info=False"
    .RecordSource = "select * from countries"
    .Refresh
End With

End Sub
