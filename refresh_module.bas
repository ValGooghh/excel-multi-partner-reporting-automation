Attribute VB_Name = "Module4"
Option Explicit

'====================================================
' MAIN BUTTON MACRO
'====================================================
Public Sub Refresh_All_Transaksi_Folder()

    Dim folderPath As String
    folderPath = PickFolder()
    If folderPath = "" Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim runDate As Date: runDate = Date
    Dim runID As String: runID = Format(Now, "yyyymmdd_hhnnss")

    Dim wsLog As Worksheet
    Set wsLog = EnsureLogSheet()

    Dim filesProcessed As Long
    Dim okCount As Long
    Dim failCount As Long
    Dim failedList As String

    Dim f As String
    f = Dir(folderPath & "\*.xl*")

    Do While f <> ""

        If Left(f, 2) <> "~$" Then

            Dim fullPath As String
            fullPath = folderPath & "\" & f

            Dim wb As Workbook
            Dim tStart As Date, tEnd As Date
            Dim status As String, msg As String

            tStart = Now
            status = "SUCCESS"
            msg = ""

            On Error GoTo FileFail

            Set wb = Workbooks.Open(fullPath, UpdateLinks:=0, ReadOnly:=False)

            ForceNoBackgroundQuery wb
            wb.RefreshAll

            On Error Resume Next
            Application.CalculateUntilAsyncQueriesDone
            On Error GoTo FileFail

            wb.Save
            wb.Close False

            tEnd = Now

            filesProcessed = filesProcessed + 1
            okCount = okCount + 1

            AppendLog wsLog, runDate, runID, tStart, tEnd, folderPath, f, status, msg

            On Error GoTo 0
        End If

NextFile:
        f = Dir()
    Loop

 
    Dim exactFilePath As String
    exactFilePath = "C:\Portofolio\1. Januari 2026.xlsx"

    Dim exactStatus As String, exactMsg As String
    Dim exactStart As Date, exactEnd As Date

    exactStart = Now
    exactStatus = "SUCCESS"
    exactMsg = ""

    If Len(Dir(exactFilePath)) = 0 Then
        exactEnd = Now
        exactStatus = "FAILED"
        exactMsg = "File tidak ditemukan: " & exactFilePath

        failCount = failCount + 1
        filesProcessed = filesProcessed + 1
        AppendLog wsLog, runDate, runID, exactStart, exactEnd, "", exactFilePath, exactStatus, exactMsg
        failedList = failedList & "- " & exactFilePath & " (" & exactMsg & ")" & vbCrLf
    Else
        On Error GoTo ExactFail

        Dim wbExact As Workbook
        Set wbExact = Workbooks.Open(exactFilePath, UpdateLinks:=0, ReadOnly:=False)

        ForceNoBackgroundQuery wbExact
        wbExact.RefreshAll

        On Error Resume Next
        Application.CalculateUntilAsyncQueriesDone
        On Error GoTo ExactFail

        wbExact.Save
        wbExact.Close False

        exactEnd = Now

        okCount = okCount + 1
        filesProcessed = filesProcessed + 1
        AppendLog wsLog, runDate, runID, exactStart, exactEnd, "", exactFilePath, exactStatus, exactMsg
    End If

CleanExit:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    Dim laporan As String
    laporan = "LAPORAN REFRESH TRANSAKSI" & vbCrLf & vbCrLf & _
              "Tanggal   : " & Format(runDate, "yyyy-mm-dd") & vbCrLf & _
              "Run ID    : " & runID & vbCrLf & _
              "Folder    : " & folderPath & vbCrLf & vbCrLf & _
              "Total File: " & filesProcessed & vbCrLf & _
              "Berhasil  : " & okCount & vbCrLf & _
              "Gagal     : " & failCount & vbCrLf & vbCrLf

    If failCount = 0 Then
        laporan = laporan & "Status       : SEMUA FILE BERHASIL DI-REFRESH ?" & vbCrLf
        MsgBox laporan, vbInformation, "REFRESH TRANSAKSI"
    Else
        laporan = laporan & "File Bermasalah:" & vbCrLf & failedList & vbCrLf & _
                  "Status       : SELESAI DENGAN ERROR." & vbCrLf
        MsgBox laporan, vbExclamation, "REFRESH TRANSAKSI"
    End If

    Exit Sub

FileFail:
    tEnd = Now
    filesProcessed = filesProcessed + 1
    failCount = failCount + 1

    status = "FAILED"
    msg = HumanErrorMessage(Err.Number, Err.Description)

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Set wb = Nothing
    On Error GoTo 0

    AppendLog wsLog, runDate, runID, tStart, tEnd, folderPath, f, status, msg
    failedList = failedList & "- " & f & " (" & msg & ")" & vbCrLf

    Resume NextFile

ExactFail:
    exactEnd = Now
    filesProcessed = filesProcessed + 1
    failCount = failCount + 1

    exactStatus = "FAILED"
    exactMsg = HumanErrorMessage(Err.Number, Err.Description)

    On Error Resume Next
    If Not wbExact Is Nothing Then wbExact.Close False
    Set wbExact = Nothing
    On Error GoTo 0

    AppendLog wsLog, runDate, runID, exactStart, exactEnd, "", exactFilePath, exactStatus, exactMsg
    failedList = failedList & "- " & exactFilePath & " (" & exactMsg & ")" & vbCrLf

    Resume CleanExit

End Sub

'====================================================
' POPUP FOLDER PICKER
'====================================================
Private Function PickFolder() As String

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "Pilih Folder File Transaksi"
        .AllowMultiSelect = False

        If .Show <> -1 Then
            PickFolder = ""
        Else
            PickFolder = .SelectedItems(1)
        End If
    End With

End Function

'====================================================
' CREATE / CHECK LOG SHEET
'====================================================
Private Function EnsureLogSheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("LOG_REFRESH")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "LOG_REFRESH"

        ws.Range("A1:I1").Value = Array("RunDate", "RunID", "StartTime", "EndTime", _
                                        "DurationSec", "Folder", "FileName", _
                                        "Status", "Message")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:I").AutoFit
    End If

    Set EnsureLogSheet = ws

End Function

'====================================================
' WRITE LOG ROW
'====================================================
Private Sub AppendLog(ws As Worksheet, runDate As Date, runID As String, _
                      tStart As Date, tEnd As Date, folderPath As String, _
                      fileName As String, status As String, msg As String)

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = runDate
    ws.Cells(r, 2).Value = runID
    ws.Cells(r, 3).Value = tStart
    ws.Cells(r, 4).Value = tEnd
    ws.Cells(r, 5).Value = Round((tEnd - tStart) * 86400, 2)
    ws.Cells(r, 6).Value = folderPath
    ws.Cells(r, 7).Value = fileName
    ws.Cells(r, 8).Value = status
    ws.Cells(r, 9).Value = msg

    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd"
    ws.Cells(r, 3).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(r, 4).NumberFormat = "yyyy-mm-dd hh:mm:ss"

End Sub

'====================================================
' FORCE REFRESH SYNCHRONOUS
'====================================================
Private Sub ForceNoBackgroundQuery(wb As Workbook)

    On Error Resume Next

    Dim cn As WorkbookConnection

    For Each cn In wb.Connections
        If cn.Type = xlConnectionTypeOLEDB Then
            cn.OLEDBConnection.BackgroundQuery = False
        End If
        If cn.Type = xlConnectionTypeODBC Then
            cn.ODBCConnection.BackgroundQuery = False
        End If
    Next cn

    On Error GoTo 0

End Sub

Private Function HumanErrorMessage(errNumber As Long, errDesc As String) As String

    Dim d As String
    d = LCase(errDesc)

    If errNumber = 1004 Then
        If InStr(d, "password") > 0 Or InStr(d, "protected") > 0 Then
            HumanErrorMessage = "File terkunci / butuh password."
        ElseIf InStr(d, "save") > 0 Then
            HumanErrorMessage = "File tidak bisa disimpan (mungkin sedang dibuka orang lain)."
        Else
            HumanErrorMessage = "Terjadi error saat proses Excel (1004)."
        End If

    ElseIf errNumber = 91 Then
        HumanErrorMessage = "Object tidak ditemukan (kemungkinan koneksi/pivot bermasalah)."

    ElseIf errNumber = 70 Then
        HumanErrorMessage = "Akses ditolak (file sedang digunakan / tidak punya izin)."

    ElseIf InStr(d, "cannot access") > 0 Or InStr(d, "not found") > 0 Then
        HumanErrorMessage = "File tidak ditemukan atau tidak bisa diakses."

    ElseIf InStr(d, "connection") > 0 Or InStr(d, "refresh") > 0 Then
        HumanErrorMessage = "Gagal refresh data (cek koneksi internet/server)."

    Else
        HumanErrorMessage = "Gagal diproses. (" & errNumber & ")"
    End If

End Function


