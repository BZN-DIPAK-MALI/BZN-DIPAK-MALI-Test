Imports System
Imports System.IO
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel ' For XLSX
Imports NPOI.HSSF.UserModel ' For XLS

Module Payment

    Dim objLogCls As New ClsErrLog
    Dim objGetSetINI As ClsShared
    Dim objBaseClass As ClsBase
    Public queue As Queue
    Public Function GenerateRBIOutPutFile(ByRef _dtRBI As DataTable, ByVal strFileName As String) As Boolean


        Try
            GenerateRBIOutPutFile = False

            gstrOutputFile = Path.GetFileName(strFileName)
            Using fs As FileStream = New FileStream(strFileName, FileMode.Create, FileAccess.Write)

                Dim workbook As IWorkbook

                If (Path.GetExtension(strFileName).ToUpper() = ".xlsx".ToString().ToUpper()) Then
                    workbook = New XSSFWorkbook()
                Else
                    workbook = New HSSFWorkbook()
                End If

                Dim sheet As ISheet = workbook.CreateSheet("Sheet1")

                ' Create header row
                Dim headerRow As IRow = sheet.CreateRow(0)
                For Each col As DataColumn In _dtRBI.Columns
                    Dim cell As ICell = headerRow.CreateCell(col.Ordinal)
                    cell.SetCellValue(col.ColumnName)
                Next

                ' Create data rows
                For Each dr As DataRow In _dtRBI.Rows
                    Dim dataRow As IRow = sheet.CreateRow(sheet.LastRowNum + 1)
                    For Each col As DataColumn In _dtRBI.Columns
                        Dim cell As ICell = dataRow.CreateCell(col.Ordinal)
                        cell.SetCellValue(dr(col).ToString())
                    Next
                Next

                ' Auto-size columns (optional)
                For i As Integer = 0 To _dtRBI.Columns.Count - 1
                    sheet.AutoSizeColumn(i)
                Next

                workbook.Write(fs)
            End Using

        Catch ex As Exception
            blnErrorLog = True

            GenerateRBIOutPutFile = True
            Call objBaseClass.Handle_Error(ex, "Form", Err.Number, "GenerateRBIOutPutFile")
        Finally



        End Try
    End Function

    Public Function GenerateRBIOutPutFile(dt As DataTable, filePath As String, isXlsx As Boolean)

    End Function

    Private Function Check_Tab(ByVal strTemp) As String
        Try
            If InStr(strTemp, vbTab) > 0 Then
                Check_Tab = Chr(34) & strTemp & Chr(34) & vbTab
            Else
                Check_Tab = strTemp & vbTab
            End If
        Catch ex As Exception
            blnErrorLog = True
            objGetSetINI.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Check_Tab")
        End Try
    End Function
    Private Function Check_Comma(ByVal strTemp) As String
        Try
            If InStr(strTemp, ",") > 0 Then
                Check_Comma = Chr(34) & strTemp & Chr(34) & ","
            Else
                Check_Comma = strTemp & ","
            End If
        Catch ex As Exception
            blnErrorLog = True
            objGetSetINI.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Check_Comma")

        End Try
    End Function

    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen).Trim()

        Catch ex As Exception
            blnErrorLog = True

            Call objLogCls.Handle_Error(ex, "Form", Err.Number, "Pad_Length")

        End Try

    End Function

    Private Sub ClearArray(ByRef ArrRow() As String)

        Try
            For i As Integer = 0 To ArrRow.Length - 1
                ArrRow(i) = ""
            Next

        Catch ex As Exception

        End Try
    End Sub

End Module