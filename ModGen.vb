Option Explicit On

Module ModGen

    Public gstrINIPath As String                    ' iniSetting path 

    Public gstrOutputFile As String
    Public gstrInputFile As String
    Public gstrInputFolder As String

    'Added by Jaiwant dtd 03-06-2011
    Public gstrOutputFileListing(0) As String
    Public gstrOutputFileCount As Integer

    Public strAuditFolderPath As String             ' Audit folder path
    Public strErrorFolderPath As String             ' Error folder path
    Public strInputFolderPath As String             ' Input folder path
    Public strOutputFolderPath As String            ' Output folder path
    Public strTempFolderPath As String              ' Temp folder path
    Public strReportFolderPath As String            ' Report folder path
    Public strValidationPath As String              ' Validation file path
    Public strSpCharValidation As String            ' Special Character Validation file path
    Public strMasterPath As String              ' Master file path


    Public strProceed As String
    Public NoOfRecords As Double
    Public StrOutputType As String
    Public strInputStartLineNo As String
    Public strFileCounterStart As String

    Public strArchivedFolderSuc As String
    Public strArchivedFolderUnSuc As String

    Public Client_Code As String = ""
    Public Client_Code_1 As String = ""
    Public Client_Code_2 As String = ""
    Public Client_Name As String = ""
    Public strInputDateFormat As String

    Public blnErrorLog As Boolean = False
    Public strSettingClientCode As String
    Public strSettingClientCode2 As String
    Public strSettingClientName As String

    ''Encryption
    Public strEncrypt As String
    Public strBatchFilePath As String
    Public strPICKDIRpath As String
    Public strDROPDIRPath As String

    ''Additional
    Public strBackDate As String
    Public strBeneCode As String
    Public strDomainId As String

    Public InputExcelSheetNo As String

End Module
