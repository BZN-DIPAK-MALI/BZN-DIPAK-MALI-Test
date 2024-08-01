Imports System
Imports System.IO
Public Class Form1
    Dim objBaseClass As ClsBase
    Dim objFileValidate As ClsValidation
    Dim objGetSetINI As ClsShared
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            Timer1.Enabled = True
            Timer1.Interval = 1000

            Generate_SettingFile()

        Catch ex As Exception
            Call objBaseClass.Handle_Error(ex, "Form", Err.Number, "Form_Load")

        End Try
    End Sub
    Private Sub Generate_SettingFile()

        Dim strConverterCaption As String = ""
        Dim strSettingsFilePath As String = My.Application.Info.DirectoryPath & "\settings.ini"

        Try
            objGetSetINI = New ClsShared

            '-Genereate Settings.ini File-
            If Not File.Exists(strSettingsFilePath) Then

                '-General Section-
                Call objGetSetINI.SetINISettings("General", "Date", Format(Now, "dd/MM/yyyy"), strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Audit Log", My.Application.Info.DirectoryPath & "\Audit", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Error Log", My.Application.Info.DirectoryPath & "\Error", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Input Folder", My.Application.Info.DirectoryPath & "\INPUT", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderSuc", My.Application.Info.DirectoryPath & "\Archive\Success", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderUnSuc", My.Application.Info.DirectoryPath & "\Archive\Unsuccess", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Output Folder", My.Application.Info.DirectoryPath & "\Output", strSettingsFilePath)

                ''Call objGetSetINI.SetINISettings("General", "Temp Folder", My.Application.Info.DirectoryPath & "\Temp", strSettingsFilePath)
                ''Call objGetSetINI.SetINISettings("General", "Report Folder", My.Application.Info.DirectoryPath & "\Report", strSettingsFilePath)
                ''Call objGetSetINI.SetINISettings("General", "Master", My.Application.Info.DirectoryPath & "\Master\Master.xlsx", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Validation", My.Application.Info.DirectoryPath & "\Validation\RBI Generic Validation.xls", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("General", "Special Character Validation", My.Application.Info.DirectoryPath & "\Validation\Special Character Mapping.xls", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Converter Caption", "Maruti special character Convertor", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Process Output File Ignoring Invalid Transactions", "Y", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Input ExcelSheet No", "2", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "RBI File Counter", "0", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Start Input File Processing From Line Number", "5", strSettingsFilePath) '' change by dipak22-11-2016     6 to 2 
                Call objGetSetINI.SetINISettings("General", "==", "==", strSettingsFilePath) 'Separator



                'Call objGetSetINI.SetINISettings("Client Details", "Client Name", strSettingClientName, strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Client Details", "Client Code1", strSettingClientCode, strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Client Details", "Client Code2", strSettingClientCode2, strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Client Details", "Client Name", "JAVI INFRATECH PRIVATE LIMITED", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Client Details", "Client Code", "RBI004", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Client Details", "Domain ID", "TEST", strSettingsFilePath)

                'Call objGetSetINI.SetINISettings("Client Details", "Input Date Format", "dd/MM/yyyy", strSettingsFilePath)    'Blank By Default
                'Call objGetSetINI.SetINISettings("Client Details", "==", "==", strSettingsFilePath) 'Separator
                ''---


            End If

            '-Get Converter Caption from Settings-
            If File.Exists(strSettingsFilePath) Then
                strConverterCaption = objGetSetINI.GetINISettings("General", "Converter Caption", strSettingsFilePath)
                If strConverterCaption <> "" Then
                    Text = strConverterCaption.ToString() & " - Version " & Mid(Application.ProductVersion.ToString(), 1, 3)
                Else
                    MsgBox("Either settings.ini file does not contains the key as [ Converter Caption ] or the key value is blank" & vbCrLf & "Please refer to " & strSettingsFilePath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End
                End If
            End If

        Catch ex As Exception
            MsgBox("Error" & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error while Generating Settings File")
            End

        Finally
            If Not objGetSetINI Is Nothing Then
                objGetSetINI.Dispose()
                objGetSetINI = Nothing
            End If

        End Try

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Interval = 1000
        Timer1.Enabled = False

        Conversion_Process()

        Timer1.Enabled = True
    End Sub
    Private Function GetAllSettings() As Boolean

        Try
            GetAllSettings = False

            If Not File.Exists(My.Application.Info.DirectoryPath & "\settings.ini") Then
                GetAllSettings = True
                MsgBox("Either settings.ini file does not exists or invalid file path" & vbCrLf & My.Application.Info.DirectoryPath & "\settings.ini", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            End If

            '-Audit Folder Path-
            If strAuditFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strAuditFolderPath) Then
                    Directory.CreateDirectory(strAuditFolderPath)
                    If Not Directory.Exists(strAuditFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Audit Log folder. Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", True)
                        End If
                        MsgBox("Invalid path for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                        Exit Function
                    End If
                End If
            End If

            '-Error Folder Path-
            If strErrorFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Error Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strErrorFolderPath) Then
                    Directory.CreateDirectory(strErrorFolderPath)
                    If Not Directory.Exists(strErrorFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Error Log folder. Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Error Log folder." & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Input Folder Path-
            If strInputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Input folder" & vbCrLf & "Please check settings.ini file, the key as [ Input Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strInputFolderPath) Then
                    Directory.CreateDirectory(strInputFolderPath)
                    If Not Directory.Exists(strInputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Input Folder. Please check [ settings.ini ] file, the key as [ Input Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Input Folder." & vbCrLf & "Please check settings.ini file, the key as [ Input Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If


            '-Archive Successful-
            If strArchivedFolderSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archive Suc folder" & vbCrLf & "Please check settings.ini file, the key as [ Archive Suc Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderSuc) Then
                    Directory.CreateDirectory(strArchivedFolderSuc)
                    If Not Directory.Exists(strArchivedFolderSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archive Suc Folder. Please check settings.ini file, the key as [ Archive Suc Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archive Suc folder", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Settings Error")
                    End If
                End If
            End If

            '-Archive Unsuccessful-
            If strArchivedFolderUnSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archive UnSuc folder" & vbCrLf & "Please check settings.ini file, the key as [ Archive UnSuc Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderUnSuc) Then
                    Directory.CreateDirectory(strArchivedFolderUnSuc)
                    If Not Directory.Exists(strArchivedFolderUnSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archive UnSuc Folder. Please check settings.ini file, the key as [ Archive UnSuc Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archive UnSuc folder", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Settings Error")
                    End If
                End If
            End If

            '-Output Folder Path-
            If strOutputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Output folder" & vbCrLf & "Please check settings.ini file, the key as [ Output Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strOutputFolderPath) Then
                    Directory.CreateDirectory(strOutputFolderPath)
                    If Not Directory.Exists(strOutputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Output Folder. Please check [ settings.ini ] file, the key as [ Output Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Output Folder." & vbCrLf & "Please check settings.ini file, the key as [ Output Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            ''-Temp Folder Path-
            'If strTempFolderPath = "" Then
            '    GetAllSettings = True
            '    MsgBox("Path is blank for Temp folder" & vbCrLf & "Please check settings.ini file, the key as [ Temp Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    Exit Function
            'Else
            '    If Not Directory.Exists(strTempFolderPath) Then
            '        Directory.CreateDirectory(strTempFolderPath)
            '        If Not Directory.Exists(strTempFolderPath) Then
            '            GetAllSettings = True
            '            If Not objBaseClass Is Nothing Then
            '                objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Temp Folder. Please check settings.ini file, the key as [ Temp Folder ] contains invalid path specification.", True)
            '            End If
            '            MsgBox("Invalid path for Temp Folder." & vbCrLf & "Please check settings.ini file, the key as [ Temp Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '        End If
            '    End If
            'End If

            ''-Report Folder Path-
            'If strReportFolderPath = "" Then
            '    GetAllSettings = True
            '    MsgBox("Path is blank for Report folder" & vbCrLf & "Please check settings.ini file, the key as [ Report Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    Exit Function
            'Else
            '    If Not Directory.Exists(strReportFolderPath) Then
            '        Directory.CreateDirectory(strReportFolderPath)
            '        If Not Directory.Exists(strReportFolderPath) Then
            '            GetAllSettings = True
            '            If Not objBaseClass Is Nothing Then
            '                objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Report Folder. Please check settings.ini file, the key as [ Report Folder ] contains invalid path specification.", True)
            '            End If
            '            MsgBox("Invalid path for Report Folder." & vbCrLf & "Please check settings.ini file, the key as [ Report Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '        End If
            '    End If
            'End If

            ''-Validation File Path-
            'If strValidationPath = "" Then
            '    GetAllSettings = True
            '    MsgBox("Path is blank for Validation file." & vbCrLf & "Please check settings.ini file, the key as [ Validation ] is either does not exist or left blank.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    Exit Function
            'Else
            '    If Not File.Exists(strValidationPath) Then
            '        GetAllSettings = True
            '        If Not objBaseClass Is Nothing Then
            '            objBaseClass.LogEntry("Error in settings.ini file, Validation file does not exist or invalid file path", True)
            '        End If
            '        MsgBox("Validation file does not exist or invalid file path" & vbCrLf & strValidationPath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    End If
            'End If

            '-Master File Path-
            'If strMasterPath = "" Then
            '    GetAllSettings = True
            '    MsgBox("Path is blank for Master file." & vbCrLf & "Please check settings.ini file, the key as [ Master ] is either does not exist or left blank.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    Exit Function
            'Else
            '    If Not File.Exists(strMasterPath) Then
            '        GetAllSettings = True
            '        If Not objBaseClass Is Nothing Then
            '            objBaseClass.LogEntry("Error in settings.ini file, Master file does not exist or invalid file path", True)
            '        End If
            '        MsgBox("Validation file does not exist or invalid file path" & vbCrLf & strValidationPath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    End If
            'End If


            '-Special Character Validation File Path-
            If strSpCharValidation = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Special Character Mapping file." & vbCrLf & "Please check settings.ini file, the key [Special Character Validation] is either does not exist or left blank.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            Else
                If Not File.Exists(strSpCharValidation) Then
                    GetAllSettings = True
                    If Not objBaseClass Is Nothing Then
                        objBaseClass.LogEntry("Error in settings.ini file, Special Character Mapping file does not exist or invalid file path", True)
                    End If
                    MsgBox("Special Character Mapping file does not exist or invalid file path" & vbCrLf & strSpCharValidation, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                End If
            End If

        Catch ex As Exception
            GetAllSettings = True
            MsgBox("Error-" & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error While Getting Log Path from Settings.ini File")

        End Try

    End Function
    Private Sub Conversion_Process()
        Dim objFolderAll As DirectoryInfo

        Try
            objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")

            '-Get Settings-
            If GetAllSettings() = True Then
                MsgBox("Either file path is invalid or any key value is left blank in settings.ini file", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error In Settings")
                Exit Sub
            End If

            '-Process Input-
            'Process_Each(txtFilePath.Text)


            objFolderAll = New DirectoryInfo(strInputFolderPath)
            If objFolderAll.GetFiles.Length = 0 Then
                objFolderAll = Nothing
            Else
                objBaseClass.LogEntry("", False)
                objBaseClass.LogEntry("Process Started for INPUT Files")

                For Each objFileOne As FileInfo In objFolderAll.GetFiles("*.*")
                    objBaseClass.isCompleteFileAvailable(objFileOne.FullName)
                    If Mid(objFileOne.FullName, objFileOne.FullName.Length - 3, 4).ToString().ToUpper() <> ".BAK" Then
                        objBaseClass.LogEntry("", False)
                        objBaseClass.LogEntry("INPUT File [ " & objFileOne.Name & " ] -- Started At -- " & Format(Date.Now, "hh:mm:ss"), False)

                        Process_Each(objFileOne.FullName)

                        objFolderAll.Refresh()
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox("Error-" & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Conversion_Process")

        Finally
            If Not objBaseClass Is Nothing Then
                objBaseClass.Dispose()
                objBaseClass = Nothing
            End If

        End Try

    End Sub
    Private Sub Process_Each(ByVal StrInputFileName As String)

        Dim StrAns As Int32


        Try
            gstrInputFolder = StrInputFileName.Substring(0, StrInputFileName.LastIndexOf("\"))
            'gstrInputFile = StrInputFileName.Substring(StrInputFileName.LastIndexOf("\"))
            gstrInputFile = Path.GetFileName(StrInputFileName)



            '-Conversion Process-

            objBaseClass.LogEntry("", False)
            objBaseClass.LogEntry("Process Started")
            objBaseClass.LogEntry("Reading Input File " & gstrInputFile, False)


            objFileValidate = New ClsValidation(gstrInputFolder & "\" & gstrInputFile, objBaseClass.gstrIniPath)

            If objFileValidate.CheckValidateFile() = True Then
                objBaseClass.LogEntry("Input File Reading Completed Successfully")

                If objFileValidate.DtUnSucInput.Rows.Count = 0 Then
                    objBaseClass.LogEntry("Input File Validated Successfully")

                    If objFileValidate.DtInput.Rows.Count > 0 Then
                        objBaseClass.LogEntry("Output File Generation Process Started")




                        If GenerateRBIOutPutFile(objFileValidate.DtInput, strOutputFolderPath & "\" & gstrInputFile) = True Then       ''Generating  HLL Output
                            objBaseClass.LogEntry("Output File Generation process failed due to Error", True)
                            objBaseClass.LogEntry("Output File Generation process failed due to Error", False)
                            'MessageBox.Show("Output File Generation process failed due to Error", Client_Name, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            objBaseClass.LogEntry("Output Files  [ " & gstrOutputFile & " ]  is Generated Successfully", False)
                            'MessageBox.Show("Output File [ " & gstrOutputFile & " ] is Generated Successfully", Client_Name, MessageBoxButtons.OK, MessageBoxIcon.Information)

                            objBaseClass.FileMove(StrInputFileName, strArchivedFolderSuc & "\" & Path.GetFileName(StrInputFileName))
                            objBaseClass.LogEntry("[ " & gstrInputFile & " ] files moved to Archived Folder Successful")
                        End If

                    Else
                        objBaseClass.LogEntry("No Valid Record present in Input File")
                        'MessageBox.Show("No Valid Record present in Input File", Client_Name, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        objBaseClass.FileMove(StrInputFileName, strArchivedFolderUnSuc & "\" & Path.GetFileName(StrInputFileName))


                    End If



                Else
                    objBaseClass.LogEntry("No Valid Record present in Input File")
                    'MessageBox.Show("No Valid Record present in Input File", Client_Name, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    objBaseClass.FileMove(StrInputFileName, strArchivedFolderUnSuc & "\" & Path.GetFileName(StrInputFileName))
                End If

            Else
                objBaseClass.LogEntry(gstrInputFile & " is not Valid Input File", False)
                'MessageBox.Show(gstrInputFile & " is not Valid Input File", Client_Name, MessageBoxButtons.OK, MessageBoxIcon.Error)
                objBaseClass.FileMove(StrInputFileName, strArchivedFolderUnSuc & "\" & Path.GetFileName(StrInputFileName))


            End If

            '-Process Status-
            If StrAns <> 7 Then
                objBaseClass.LogEntry("Process Completed")
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "CmdProcess_Click")

        Finally


            'objBaseClass.ObjectDispose(objFileValidate.DtInput)
            'objBaseClass.ObjectDispose(objFileValidate.DtUnSucInput)

            If Not objFileValidate Is Nothing Then
                objBaseClass.ObjectDispose(objFileValidate.DtInput)
                objBaseClass.ObjectDispose(objFileValidate.DtUnSucInput)

                objFileValidate.Dispose()
                objFileValidate = Nothing
            End If

        End Try

    End Sub
End Class
