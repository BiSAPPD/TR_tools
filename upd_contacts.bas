Sub upd_contacts()

Const cRowStart = 2
Const cDateStat = 2
Const cEmplNameOld = 3
Const cEmplNameStatus = 4
Const cEmplNameNew = 17
Const cBrandColumnStart = 9
Const cBrandColumnEnd = 15

Const cRowStartTR = 2
Const cSrepNameTR = 3
Const cFlsmNameTR = 6

Const cShTRCntct = "Contacts"
Const cShLog = "Log"
Const cShData = "Data"

Dim ActWb As String, PreviousPath As String, BrandName As String, ActualPatch As String, PreviousPatch As String
Dim EmplNameStatus As String, EmplNameOld As String, EmplNameNew As String, ShLog As String
Dim TRName As String
Dim DateStatMonth As Integer, DateStatYear As Integer, ActualMonth As Integer, ActualYear As Integer
Dim Lrow As Long, f_clm As Long, f_rw As Long, LRowCntct As Long, f_clm_tr As Long
Dim clmEmp As Integer, i As Integer, n As Integer
Dim WbWrk As Workbook, WbTR As Workbook
Dim DataSh As Worksheet, LogSH As Worksheet, OutDataSh As Worksheet

ActualMonth = CInt(InputBox("Month"))
ActualYear = CInt(InputBox("YearEnd"))

myLib.VBA_Start

ActWb = ActiveWorkbook.name
myLib.CreateSh (cShLog)
myLib.sheetActivateCleer (cShLog)

Sheets(cShData).Select
Lrow = myLib.GetLastRow

PreviousPatch = Empty
TRName = Empty
For f_clm = cBrandColumnStart To cBrandColumnEnd
    
    For f_rw = cRowStart To Lrow
        BrandName = Cells(f_rw, f_clm)
        Workbooks(ActWb).Activate
        Sheets(cShData).Select
        DateStatMonth = month(Cells(f_rw, cDateStat))
        DateStatYear = Year(Cells(f_rw, cDateStat))
        ActualPatch = myLib.GetPatchHistTR(BrandName, ActualYear, DateStatYear, ActualMonth, DateStatMonth)
        EmplNameStatus = Cells(f_rw, cEmplNameStatus)
        EmplNameOld = Cells(f_rw, cEmplNameOld)
        EmplNameNew = Cells(f_rw, cEmplNameNew)
        If PreviousPatch <> ActualPatch Then
            If Len(TRName) <> 0 Then WbTR.Close
            TRName = myLib.OpenFile(ActualPatch, cShTRCntct)
            Sheets(cShTRCntct).Select
            LRowCntct = myLib.GetLastRow
        End If
        Workbooks(TRName).Activate
        Sheets(cShTRCntct).Select
        For f_clm_tr = cRowStartTR To LRowCntct
            Select Case EmplNameStatus
                Case "SREP": clmEmp = cSrepNameTR
                Case "FLSM": clmEmp = cFlsmNameTR
            End Select
            If EmplNameOld = Cells(f_clm_tr, clmEmp) Then
                Cells(f_clm_tr, clmEmp) = EmplNameNew
                Exit For
            End If
        Next
        Workbooks(ActWb).Activate
        Sheets(cShLog).Select
        i = i + 1
        n = n + 1: Cells(i, n) = ActualPatch
        n = n + 1: Cells(i, n) = EmplNameOld
        n = n + 1: Cells(i, n) = EmplNameStatus
        n = n + 1: Cells(i, n) = EmplNameNew
    Next
Next
myLib.VBA_End
End Sub

