Attribute VB_Name = "modBackgroundWorker"
Option Explicit

Private m_background As NETBackgroundWorkerWrapper.BackgroundWorkerWrapper

Public Sub StartBackground(background As NETBackgroundWorkerWrapper.BackgroundWorkerWrapper, argument As Variant)
    Set m_background = background
    m_background.RunWorkerAsync AddressOf BackgroundWork, argument
End Sub

Public Sub BackgroundWork(ByRef argument As Variant, ByRef e As NETBackgroundWorkerWrapper.RunWorkerCompletedEventArgsWrapper)
    Dim arrayJson   As Object
    Dim objJson     As Object
    Dim objBuku     As Buku

    Dim jsonResult  As String

    On Error GoTo errorHandler
    
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

        m_background.ReportProgress 0, objBuku
        
        If m_background.CancellationPending Then
            e.Cancelled = True
            Exit Sub
        End If
    Next objJson
    
    Exit Sub

errorHandler:
    Debug.Print Err.Description
End Sub
