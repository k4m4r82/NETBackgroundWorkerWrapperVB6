VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Demo Synchronous dan Asynchronous Method"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwBuku 
      Height          =   5895
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10398
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Isbn"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Judul"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Penerbit"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "click me if you can"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   7455
   End
   Begin VB.CommandButton btnAsynchronous 
      Caption         =   "Asynchronous"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton btnSynchronous 
      Caption         =   "Synchronous"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents background As NETBackgroundWorkerWrapper.BackgroundWorkerWrapper
Attribute background.VB_VarHelpID = -1

Private Sub Form_Load()
    Set background = New NETBackgroundWorkerWrapper.BackgroundWorkerWrapper
End Sub

Private Sub background_ProgressChanged(ByVal sender As Variant, ByVal e As NETBackgroundWorkerWrapper.ProgressChangedEventArgsWrapper)
    Dim objBuku As Buku

    Set objBuku = e.userState
    
    Call IsiListView(objBuku)
End Sub

Private Sub background_RunWorkerCompleted(ByVal sender As Variant, ByVal e As NETBackgroundWorkerWrapper.RunWorkerCompletedEventArgsWrapper)
    MsgBox "Done !!!"
End Sub

Private Sub btnAsynchronous_Click()
    modBackgroundWorker.StartBackground background, "Download data buku"
End Sub

Private Sub btnSynchronous_Click()
    Dim arrayJson   As Object
    Dim objJson     As Object
    Dim objBuku     As Buku
    
    Dim jsonResult  As String
    
    jsonResult = GetRequest(API_URL)
    
    Set arrayJson = ModJSON.parse(jsonResult)
    
    For Each objJson In arrayJson
        Set objBuku = New Buku
        
        With objBuku
            .isbn = objJson.Item("isbn")
            .judul = objJson.Item("judul")
            .penerbit = objJson.Item("penerbit")
            .pengarang = objJson.Item("pengarang")
        End With
        
        Call IsiListView(objBuku)
    Next objJson
    
    MsgBox "Done !!!"
End Sub

Private Sub IsiListView(ByVal objBuku As Buku)

    With lvwBuku.ListItems.Add(, , lvwBuku.ListItems.Count + 1)
        .SubItems(1) = objBuku.isbn
        .SubItems(2) = objBuku.judul
        .SubItems(3) = objBuku.penerbit
    End With
    
End Sub
