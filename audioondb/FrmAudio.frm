VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmAudio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AudioFiles"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4140
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Audio files"
      Filter          =   "wav only(*.wav)|*.wav"
   End
   Begin VB.CommandButton btnplay 
      Caption         =   "Play"
      Height          =   390
      Left            =   3915
      TabIndex        =   4
      Top             =   90
      Width           =   690
   End
   Begin VB.CommandButton btndel 
      Caption         =   "delete"
      Height          =   495
      Left            =   3375
      TabIndex        =   3
      Top             =   525
      Width           =   1215
   End
   Begin VB.CommandButton btnadd 
      Caption         =   "add"
      Height          =   495
      Left            =   2115
      TabIndex        =   2
      Top             =   525
      Width           =   1215
   End
   Begin VB.TextBox TxtNam 
      DataField       =   "AudioName"
      DataSource      =   "adc1"
      Height          =   375
      Left            =   1335
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   105
      Width           =   2520
   End
   Begin MSAdodcLib.Adodc adc1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      Top             =   1110
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   847
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=audios.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=audios.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * From Audios"
      Caption         =   "Audios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Audio Name:"
      Height          =   240
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   1125
   End
End
Attribute VB_Name = "FrmAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Const SND_FILENAME = &H20000     '  name is a file name


Dim act_Fil As String
Private Sub btnadd_Click()
'add audio file
dlg.ShowOpen
If Len(Trim(dlg.filename)) > 0 Then
'save the file
act_Fil = dlg.filename
Dim splt
splt = Split(dlg.filename, "\")
adc1.Recordset.AddNew
TxtNam.Text = splt(UBound(splt))
GetAudio
adc1.Refresh
adc1.Refresh
adc1.Refresh
End If
End Sub

Private Sub btndel_Click()
If adc1.Recordset.RecordCount <> 0 Then
adc1.Recordset.Delete
adc1.Refresh
adc1.Refresh
adc1.Refresh
End If
End Sub

Private Sub btnplay_Click()
If adc1.Recordset.RecordCount > 0 Then
SetAudios
'moveto prevent error
If adc1.Recordset.EOF Then
    adc1.Recordset.MoveLast
Else
    adc1.Recordset.MoveNext
End If
    adc1.Recordset.MovePrevious
End If
End Sub

Private Sub Form_Load()
ChDir App.Path
adc1.ConnectionString = "Provider=microsoft.jet.oledb.4.0;persist security info=false;data source=" & App.Path & "\audios.mdb"
adc1.Refresh
End Sub

Private Sub SetAudios()
On Error GoTo Handler
Dim BitArray() As Byte
Dim file_name As String
Dim FNS As Integer
Dim FileLen As Long, num_blocks As Long, Excess As Long, block_num As Long, hgt As Single
    file_name = App.Path & "\pinoy.wav"
    FNS = FreeFile
    Open file_name For Binary As #FNS
    FileLen = adc1.Recordset.Fields("Size").Value
    num_blocks = FileLen / 1000
    Excess = FileLen Mod 1000
    For block_num = 1 To num_blocks
        BitArray() = adc1.Recordset.Fields("Audio").GetChunk(1000)
        Put #FNS, , BitArray()
    Next block_num
    If Excess > 0 Then
        BitArray() = adc1.Recordset.Fields("Audio").GetChunk(Excess)
        Put #FNS, , BitArray()
    End If
    Close #FNS
    Dim nullp As Long
    PlaySound file_name, nullp, SND_FILENAME
    Kill App.Path & "\pinoy.wav"
Exit Sub
Handler:
    MsgBox Err.Description
End Sub

Private Sub GetAudio()
On Error GoTo Handler
Dim FNS As String
Dim BitArray() As Byte
Dim num_blocks As Long, Excess As Long, block_num As Long, FileLen As Long
    FNS = FreeFile
    Open act_Fil For Binary Access Read As #FNS
    FileLen = LOF(FNS)
    If FileLen > 0 Then
        num_blocks = FileLen / 1000
        Excess = FileLen Mod 1000
        adc1.Recordset("Size").Value = FileLen
        ReDim BitArray(100000)
        For block_num = 1 To num_blocks
            Get #FNS, , BitArray()
            adc1.Recordset.Fields("Audio").AppendChunk BitArray()
        Next block_num
        If Excess > 0 Then
            ReDim BitArray(Excess)
            Get #FNS, , BitArray()
            adc1.Recordset.Fields("Audio").AppendChunk BitArray()
        End If
        Close #FNS
        adc1.Recordset.Update
    End If
Exit Sub

Handler:
    MsgBox Err.Description

End Sub


