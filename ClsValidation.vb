Imports System
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions

Public Class ClsValidation

    Implements IDisposable

    Private ObjBaseClass As ClsBase         ''need to be dispose 

    Private DtValidation As DataTable             ''need to be dispose
    Private DtSpCharValidation As DataTable       ''need to be dispose
    ''---

    Public DtInput As DataTable             ''need to be dispose
    Public DtUnSucInput As DataTable        ''need to be dispose

    Private DtTemp As DataTable             ''need to be dispose

    Private StrFilePath As String
    Private ValidationPath As String
    Private SpCharValidationPath As String
    ''---

    Public StrSettingPath As String

    Public ErrorMessage As String


    Public Sub New(ByVal _strFilePath As String, ByVal _SettINIPath As String)

        StrFilePath = _strFilePath
        StrSettingPath = _SettINIPath

        Try
            ObjBaseClass = New ClsBase(_SettINIPath)
            'ValidationPath = ObjBaseClass.GetINISettings("General", "Validation", _SettINIPath)
            SpCharValidationPath = ObjBaseClass.GetINISettings("General", "Special Character Validation", _SettINIPath)
            'strMasterPath = ObjBaseClass.GetINISettings("General", "Master", _SettINIPath)
            ''---

            DtInput = New DataTable("RBIInput")
            DefineColumnForRBI(DtInput)
            DtUnSucInput = New DataTable("RBIUnSucInput")
            DefineColumnForRBI(DtUnSucInput)

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Constructor")
        End Try

    End Sub

    Private Sub DefineColumnForRBI(ByRef DtInput As DataTable)

        DtInput.Columns.Add(New DataColumn("Sl.No")) '0
        DtInput.Columns.Add(New DataColumn("Vendor /Supplier/Bene Code")) '1
        DtInput.Columns.Add(New DataColumn("Supplier/Bene Name")) '2
        DtInput.Columns.Add(New DataColumn("Supplier Address/Bene Address")) '3
        DtInput.Columns.Add(New DataColumn("Supplier Country/Bene Country")) '4
        DtInput.Columns.Add(New DataColumn("Supplier Account Number")) '5
        DtInput.Columns.Add(New DataColumn("Beneficiary Bank SWIFT Code")) '6
        DtInput.Columns.Add(New DataColumn("Routing No US/ Sort Code GBP/ IBAN Code EUR")) '7
        DtInput.Columns.Add(New DataColumn("Payment Date")) '8
        DtInput.Columns.Add(New DataColumn("Bill of Entry no.")) '9
        DtInput.Columns.Add(New DataColumn("Invoice No.")) '10
        DtInput.Columns.Add(New DataColumn("Invoice Date")) '11
        DtInput.Columns.Add(New DataColumn("Currency")) '12
        DtInput.Columns.Add(New DataColumn("Payment Amount")) '13
        DtInput.Columns.Add(New DataColumn("Other Charges-INR (Insurance + Freight+ Other charges)")) '14
        DtInput.Columns.Add(New DataColumn("Other Charges-FCY (Insurance + Freight+ Other charges)")) '15
        DtInput.Columns.Add(New DataColumn("EEFC A/C Amount")) '16
        DtInput.Columns.Add(New DataColumn("Operative A/C Amount")) '17
        DtInput.Columns.Add(New DataColumn("EEFC Account")) '18
        DtInput.Columns.Add(New DataColumn("Operative Account")) '19
        DtInput.Columns.Add(New DataColumn("Fee Account")) '20
        DtInput.Columns.Add(New DataColumn("Customer Rate details"))
        DtInput.Columns.Add(New DataColumn("Foreign Bank Charges"))
        DtInput.Columns.Add(New DataColumn("HS Code"))
        DtInput.Columns.Add(New DataColumn("Type of Goods"))
        DtInput.Columns.Add(New DataColumn("Commodity/Goods Description"))
        DtInput.Columns.Add(New DataColumn("Vessel Name"))
        DtInput.Columns.Add(New DataColumn("Port of Shipment Country"))
        DtInput.Columns.Add(New DataColumn("Goods Carrier (Air/Ship)"))
        DtInput.Columns.Add(New DataColumn("Origin Of Goods"))
        DtInput.Columns.Add(New DataColumn("Comments To Bank"))
        DtInput.Columns.Add(New DataColumn("Remarks to Beneficiary"))
        DtInput.Columns.Add(New DataColumn("Transaction Type"))
        DtInput.Columns.Add(New DataColumn("Remarks")) '33

        'DtInput.Columns.Add(New DataColumn("TXN_NO"))   ''34
        'DtInput.Columns.Add(New DataColumn("SUBTXN_NO"))   '35
        'DtInput.Columns.Add(New DataColumn("REASON"))   '36

    End Sub

    Public Function CheckValidateFile() As Boolean

        Try
            If Not File.Exists(StrFilePath) Then
                Call ObjBaseClass.Handle_Error(New ApplicationException("Input File path is incorrect or File not found. [" & StrFilePath & "]"), "ClsValidation", -123, "CheckValidateFile")
                CheckValidateFile = False
                Exit Function
            End If


            If Not File.Exists(SpCharValidationPath) Then
                Call ObjBaseClass.Handle_Error(New ApplicationException("Special Character Validation File path is incorrect or File not found. [" & SpCharValidationPath & "]"), "ClsValidation", -123, "CheckValidateFile")
                CheckValidateFile = False
                Exit Function
            Else
                CheckValidateFile = Validate()
            End If
            ''----

            'If File.Exists(ValidationPath) Then
            '    CheckValidateFile = Validate()
            'Else
            '    Call ObjBaseClass.Handle_Error(New ApplicationException("Validation File path is incorrect. [" & ValidationPath & "]"), "ClsValidation", -123, "CheckValidateFile")
            'End If

        Catch ex As Exception
            blnErrorLog = True

            CheckValidateFile = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "CheckValidateFile")
        End Try

    End Function

    Private Function GetInArrayByComma(ByVal pStrValue As String) As String()

        Try

            Dim Tmpstr As String = ""
            Dim Index_S, Index_E, TmpIndex As Integer


            Index_E = InStr(pStrValue, Chr(34))

            If Index_E > 0 Then

                Index_S = 0
                Tmpstr = ""
                While True

                    Index_E = InStr(Index_S + 1, pStrValue, Chr(34))

                    If Index_E > 0 Then

                        Tmpstr += pStrValue.Substring(Index_S, Index_E - Index_S - 1).Replace(",", "|")
                        Index_S = Index_E
                        Index_E = InStr(Index_E + 1, pStrValue, Chr(34))
                        Tmpstr += pStrValue.Substring(Index_S, (Index_E - Index_S) - 1)
                        Index_S = Index_E

                    Else
                        Tmpstr += pStrValue.Substring(Index_S, pStrValue.Length - Index_S).Replace(",", "|")
                        GetInArrayByComma = Tmpstr.Split("|")
                        Exit While
                    End If

                End While

            Else
                GetInArrayByComma = pStrValue.Split(",")

            End If

        Catch ex As Exception

        End Try

    End Function

    Private Function RemoveBlankRow(ByRef _DtTemp As DataTable)
        'Added By Jaiwant dtd 03-06-2011    'To Remove Blank Row Exists in DataTable
        Dim blnRowBlank As Boolean

        Try
            For Each vRow As DataRow In _DtTemp.Rows
                blnRowBlank = True

                For intCol As Int32 = 0 To _DtTemp.Columns.Count - 1
                    If vRow.Item(intCol).ToString().Trim() <> "" Then
                        blnRowBlank = False
                        Exit For
                    End If
                Next

                If blnRowBlank = True Then
                    _DtTemp.Rows(vRow.Table.Rows.IndexOf(vRow)).Delete()
                End If
            Next
            _DtTemp.AcceptChanges()

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "RemoveBlankRow")

        End Try
    End Function
    Private Sub DefineColVijayTanks(ByRef DtInput As DataTable)

        DtInput.Columns.Add(New DataColumn("CustomerReferenceNo"))  '0
        DtInput.Columns.Add(New DataColumn("TransactionTypeCode"))   '1
        DtInput.Columns.Add(New DataColumn("MessageType"))   '2
        DtInput.Columns.Add(New DataColumn("DebitAccountNo"))   '3
        DtInput.Columns.Add(New DataColumn("PaymentAmount"))   '4
        DtInput.Columns.Add(New DataColumn("TransactionCurrency"))   '5
        DtInput.Columns.Add(New DataColumn("Valuedate"))   '6
        DtInput.Columns.Add(New DataColumn("BeneficiaryAccountNo"))   '7
        DtInput.Columns.Add(New DataColumn("BeneficiaryName"))   '8
        DtInput.Columns.Add(New DataColumn("BeneficiaryBankIFSC"))   '9
        DtInput.Columns.Add(New DataColumn("BeneficiaryEmailID"))   '10
        DtInput.Columns.Add(New DataColumn("BeneficiaryID"))   '11
        DtInput.Columns.Add(New DataColumn("Remarks"))   '12
        DtInput.Columns.Add(New DataColumn("ClientCode"))   '13
        DtInput.Columns.Add(New DataColumn("InputFileName"))   '14
        DtInput.Columns.Add(New DataColumn("FileProcessDate"))   '15

    End Sub
    Private Function Validate() As Boolean

        Validate = False

        Dim DrValidOutputColumn() As DataRow = Nothing

        Dim StrDataRow(33) As String
        Dim InputLineNumber As Int32
        Dim ArrDataRow As Object

        Dim intPosition As Int32
        Dim intPosField As Int32 = 0
        Dim intPosMaxLen As Int32
        Dim intPosStaticVal As Int32
        Dim intPosSpCharValid As Int32 = 3

        Dim TXN_NO As Integer
        Dim SUBTXN_NO As Integer

        Dim StrTxnType As String = ""
        Dim strIFSCCode As String = ""
        Dim strBeneAcctNo As String = ""
        ''---
        Dim StrDataRow2(15) As String
        Try
            ErrorMessage = ""


            DtTemp = ObjBaseClass.ReadExcelIntoDataTable(gstrInputFolder & "\" & gstrInputFile, 0)


            Dim dtch As New DataTable

            DefineColVijayTanks(dtch)

            Dim linecountExcel As Integer = 0
            Dim strIFSCCode1 As String = ""

            For Each ROW1 As DataRow In DtTemp.Rows
                ClearArray(StrDataRow2)
                linecountExcel += 1
                strIFSCCode1 = ""
                ArrDataRow = ROW1.ItemArray()

                StrDataRow2(0) = System.DateTime.Now.ToString("ddMMyyHHmmss") + linecountExcel.ToString().PadLeft(4, "0")  ''CustomerRefNo

                strIFSCCode1 = Left(ArrDataRow(4).ToString, 4)
                If strIFSCCode1.ToString.ToUpper = "FDRL".ToString.ToUpper() Then
                    StrDataRow2(1) = "BT"
                    StrDataRow2(2) = ""
                Else
                    If (Val(ArrDataRow(7).ToString()) >= Val(200000)) Then
                        StrDataRow2(1) = "LBT"
                        StrDataRow2(2) = "RTGS"
                    Else
                        StrDataRow2(1) = "LBT"
                        StrDataRow2(2) = "NEFT"
                    End If
                End If

                StrDataRow2(3) = "13355500014785".ToString()     '' DebitAccountNo
                StrDataRow2(4) = ArrDataRow(5).ToString()     ''PaymentAmount
                StrDataRow2(5) = "INR".ToString()     ''TransactionCurrency


                ' Original date string in dd/M/yyyy format
                Dim originalDateString As String = ArrDataRow(1).ToString()
                ' Parse the original date string
                Dim parsedDate As DateTime = DateTime.ParseExact(originalDateString, "dd/M/yyyy", CultureInfo.InvariantCulture)

                ' Format the date to dd/MM/yyyy
                Dim formattedDateString As String = parsedDate.ToString("dd/MM/yyyy")


                StrDataRow2(6) = formattedDateString.ToString().Replace("-", "/").Replace("\", "/").Replace(".", "/") ''ArrDataRow(5).ToString()     ''Valuedate

                StrDataRow2(7) = ArrDataRow(3).ToString()     ''BeneficiaryAccountNo

                Dim pattern As String = "[^a-zA-Z0-9\s]"
                Dim replacement As String = ""
                Dim rgx As New Regex(pattern)

                StrDataRow2(8) = ArrDataRow(2).ToString()     ''BeneficiaryName
                Dim result As String = rgx.Replace(StrDataRow2(8).ToString(), replacement)


                StrDataRow2(9) = ArrDataRow(4).ToString()     ''BeneficiaryBankIFSC 
                StrDataRow2(10) = ArrDataRow(7).ToString()    ''BeneficiaryEmailID
                StrDataRow2(11) = ""    ''BeneficiaryID
                StrDataRow2(12) = ""    ''PurposeCode
                StrDataRow2(13) = "" 'ArrDataRow(2).ToString()    ''ClientCode

                StrDataRow2(14) = gstrInputFile
                StrDataRow2(15) = System.DateTime.Now.ToString("dd/MM/yyyy").Replace("-", "/").Replace(".", "/")

                dtch.Rows.Add(StrDataRow2)

            Next



            DtValidation = ObjBaseClass.ReadExcelIntoDataTable(SpCharValidationPath, 1)
            DrValidOutputColumn = DtValidation.Select()

            DtSpCharValidation = ObjBaseClass.ReadExcelIntoDataTable(SpCharValidationPath, 0)


            'DtTemp.Columns.RemoveAt(28)
            'DtTemp.AcceptChanges()

            RemoveBlankRow(DtTemp)

            If DtValidation.Rows.Count >= 34 Then

                InputLineNumber = 1
                TXN_NO = 0
                SUBTXN_NO = 0

                For Each ROW As DataRow In DtTemp.Rows

                    ArrDataRow = ROW.ItemArray()


                    ClearArray(StrDataRow)
                    TXN_NO += 1
                    SUBTXN_NO = 0

                    intPosSpCharValid = 3
                    'StrDataRow(36) = ""



                    For StrIndex As Int32 = 0 To DrValidOutputColumn.Length - 1
                        If Val(DrValidOutputColumn(StrIndex)(intPosField).ToString.Trim()) <> 0 Then
                            StrDataRow(StrIndex) = GetValueFormArray(ArrDataRow, Val(DrValidOutputColumn(StrIndex)(intPosField).ToString.Trim().ToString.Trim()))

                            'If StrDataRow(StrIndex) = "~ERROR~" Then
                            '    StrDataRow(36) = StrDataRow(36).ToString() & "Input Line No. :" & InputLineNumber & " Invalid input field position defined in validation file. [ Reference : Input data array length = " & ArrDataRow.Length & " , Field Position = " & Val(DrValidOutputColumn(StrIndex)(intPosField).ToString.Trim()) & "]" & "| "
                            'End If
                        Else
                            StrDataRow(StrIndex) = ""
                        End If


                        If StrDataRow(StrIndex) <> "" Then
                            StrDataRow(StrIndex) = RemoveJunk(StrDataRow(StrIndex).ToString)
                        End If

                    Next


                    Dim ColNo As Int32 = 0

                    For Each Vrow As DataRow In DtValidation.Select()

                        If Vrow(intPosSpCharValid).ToString.Trim.ToUpper() = "Y" Then
                            If StrDataRow(ColNo) <> "" Then
                                '''For Special Character Validation
                                'Dim strSpCharValid As String = ""
                                'strSpCharValid = SpCharValidation(StrDataRow(ColNo), DtSpCharValidation)
                                'If strSpCharValid <> "" Then
                                '    StrDataRow(36) = StrDataRow(36).ToString() & "Line NO: " & InputLineNumber & " Special Character Found [" & strSpCharValid & "] for Column [" & Vrow(1).ToString.Trim.ToUpper() & "| " ''Reason
                                'End If

                                ''To Replce Special Character mention in Special Character Mapping
                                For Each SVRow As DataRow In DtSpCharValidation.Rows
                                    If StrDataRow(ColNo) <> "" Then
                                        StrDataRow(ColNo) = Replace(StrDataRow(ColNo).ToString(), SVRow(0).ToString(), SVRow(1).ToString())
                                    End If
                                Next

                                Dim charactersToRemove As String = Vrow(4).ToString()
                                For Each c As Char In charactersToRemove
                                    StrDataRow(ColNo) = StrDataRow(ColNo).Replace(c, Vrow(5).ToString())
                                Next


                            End If
                        End If

                        ColNo += 1
                    Next


                    'StrDataRow(34) = TXN_NO
                    '    StrDataRow(35) = SUBTXN_NO

                    'If StrDataRow(36).ToString.Trim() = "" Then
                    DtInput.Rows.Add(StrDataRow)
                    'Else
                    '    DtUnSucInput.Rows.Add(StrDataRow)
                    'End If


                Next

                'RemoveBlankRow(DtInput)


                Validate = True

            Else
                Call ObjBaseClass.Handle_Error(New ApplicationException("Validation is not maintained properly in " & Path.GetFileName(ValidationPath) & " validation file. It must be atleast 28 columns defination."), "ClsValidation", -123, "Validate")
            End If

        Catch ex As Exception
            blnErrorLog = True

            Validate = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Validate")

        Finally

            DrValidOutputColumn = Nothing
            ObjBaseClass.ObjectDispose(DtValidation)
            ObjBaseClass.ObjectDispose(DtTemp)

        End Try

    End Function

    Private Sub ClearArray(ByRef pArr() As String)
        Try
            For I As Int16 = 0 To pArr.Length - 1
                pArr(I) = ""
            Next

        Catch ex As Exception

        End Try

    End Sub

    Private Function GetSubstring(ByVal pStrValue As String, ByVal pStartPos As Int16, ByVal pEndPos As Int16) As String

        Try
            If pStartPos = 0 And pEndPos = 0 Then
                GetSubstring = ""
            Else
                pStartPos = pStartPos - 1
                If pStartPos >= pEndPos Then
                    GetSubstring = "~Error~"
                Else
                    GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                End If
            End If

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            GetSubstring = "~Error~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetSubstring")
        End Try

    End Function

    Private Function GetValidateDate(ByRef pStrDate As String) As Boolean

        Try
            strInputDateFormat = strInputDateFormat.ToUpper()

            Dim TmpstrInputDateFormat() As String
            Dim TempStrDateValue() As String = pStrDate.Split(" ")

            If InStr(TempStrDateValue(0), "/") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("/")
                TmpstrInputDateFormat = strInputDateFormat.Split("/")
            ElseIf InStr(TempStrDateValue(0), "-") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("-")
                TmpstrInputDateFormat = strInputDateFormat.Split("-")
            End If

            Dim HsUserDate As New Hashtable
            Dim HsSystemDate As New Hashtable
            Dim StrFinalDate As String

            If TempStrDateValue.Length = 3 Then
                For IntStr As Integer = 0 To TempStrDateValue.Length - 1
                    HsUserDate.Add(GetShort(TmpstrInputDateFormat(IntStr)), TempStrDateValue(IntStr))
                Next
                Dim SysDate() As String
                Dim dtSys As String = System.Globalization.DateTimeFormatInfo.CurrentInfo.ShortDatePattern.ToUpper()
                If InStr(dtSys, "/") > 0 Then
                    SysDate = dtSys.Split("/")
                ElseIf InStr(dtSys, "-") > 0 Then
                    SysDate = dtSys.Split("-")
                End If

                StrFinalDate = ""
                For IntStr As Integer = 0 To SysDate.Length - 1
                    If StrFinalDate = "" Then
                        StrFinalDate += HsUserDate(GetShort(SysDate(IntStr))).ToString().Trim()
                    Else
                        StrFinalDate += "/" & HsUserDate(GetShort(SysDate(IntStr))).ToString().Trim()
                    End If
                Next

                Try
                    ''pStrDate = Format(CDate(StrFinalDate), "dd/MM/yyyy")
                    pStrDate = CDate(StrFinalDate)

                    GetValidateDate = True

                Catch ex As Exception
                    GetValidateDate = False

                End Try
            Else
                GetValidateDate = False
            End If

        Catch ex As Exception
            GetValidateDate = False

        End Try
    End Function

    Private Function GetShort(ByVal pStr As String) As String

        pStr = pStr.ToUpper

        If InStr(pStr, "D") > 0 Then
            GetShort = "D"
        ElseIf InStr(pStr, "M") > 0 Then
            GetShort = "M"
        ElseIf InStr(pStr, "Y") > 0 Then
            GetShort = "Y"
        End If

    End Function

    Private Sub AddRowsToDataTable(ByVal pNotValid As Boolean, ByVal Data() As String)
        Try
            If Data Is Nothing Then Exit Sub

            If pNotValid = True Then
                DtUnSucInput.Rows.Add(Data)
            Else
                DtInput.Rows.Add(Data)
            End If


        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "AddRowsToDataTable")
        End Try
    End Sub

    Private Function GetValueFormArray(ByRef pArray() As Object, ByVal pPosition As Int16) As String

        Try
            If pArray.Length >= pPosition Then
                GetValueFormArray = pArray(pPosition - 1).ToString()
            Else
                GetValueFormArray = "~ERROR~"
            End If

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            GetValueFormArray = "~ERROR~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetValueFormArray")

        End Try

    End Function

    Public Function RemoveJunk(ByVal sText As String) As String
        ''Added By Jaiwant dtd  03-Dec-2010  ''To remove Junk Characters
        Try
            ''PURPOSE: To return only the alpha chars A-Z or a-z or 0-9 and special chars in a string and ignore junk chars.
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""

            If sText <> "" Then
                For n = 1 To iTextLen
                    sChar = Mid(sText, n, 1)
                    If IsAlpha(sChar) Then
                        RemoveJunk = RemoveJunk + sChar
                    End If
                Next
            End If

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            'Call ObjBaseClass.Handle_Error(ex, "ClsValidation", "RemoveJunk")
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "RemoveJunk")

        End Try
    End Function

    Private Function IsAlpha(ByVal sChr As String) As Boolean
        ''Added By Jaiwant dtd  03-Dec-2010  ''To remove Junk Characters

        IsAlpha = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" _
        Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[;]" Or sChr Like "[:]" _
        Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[?]" Or sChr Like "[/]" _
        Or sChr Like "[']" Or sChr Like "[""]" Or sChr Like "[|]" Or sChr Like "[\]" _
        Or sChr Like "[{]" Or sChr Like "[[]" Or sChr Like "[}]" Or sChr Like "[]]" _
        Or sChr Like "[+]" Or sChr Like "[=]" Or sChr Like "[_]" Or sChr Like "[-]" _
        Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[*]" Or sChr Like "[&]" _
        Or sChr Like "[^]" Or sChr Like "[%]" Or sChr Like "[$]" Or sChr Like "[#]" _
        Or sChr Like "[@]" Or sChr Like "[!]" Or sChr Like "[`]" Or sChr Like "[~]" _
        Or sChr Like "[ ]" 'commented dtd 03-06-2011

    End Function

    Public Function SpCharValidation(ByVal StringValue As String, ByRef _dtSpChar As DataTable) As String

        ''Added by Jaiwant dtd  03-Dec-2010
        Dim ArrSpChar(0) As String
        Dim intSpCharRow As Integer
        ''---
        ClearArray(ArrSpChar)
        Array.Resize(ArrSpChar, _dtSpChar.Select.Length)
        intSpCharRow = 0

        For Each SVRow As DataRow In _dtSpChar.Rows
            ArrSpChar(intSpCharRow) = SVRow(0).ToString
            intSpCharRow += 1
        Next

        ''Added By Jaiwant dtd  03-Dec-2010 ''For All Special Characters
        Dim StrOriginalValue As String = ""
        'Dim arrSpecialChar() As String = {"'", ";", ".", ",", "<", ">", ":", "?", """", "/", "{", "[", "}", "]", "`", "~", "!", "@", "#", "$", "%", "^", "*", "(", ")", "_", "-", "+", "=", "|", "\", "&", " "} ''Commented by Lakshmi dtd 22-03-2012
        Dim arrSpecialChar() As String = {"'", ";", ".", ",", "<", ">", ":", "?", """", "/", "{", "[", "}", "]", "`", "~", "!", "@", "#", "$", "%", "^", "*", "(", ")", "_", "-", "+", "=", "|", "\", "&"} ''Added by Lakshmi dtd 22-03-2012

        Try
            ''To remove special chars from array which need to ignore.
            For iIChar As Int16 = 0 To ArrSpChar.Length - 1
                For iSChar As Int16 = 0 To arrSpecialChar.Length - 1
                    If ArrSpChar(iIChar) = arrSpecialChar(iSChar) Then
                        arrSpecialChar(iSChar) = Nothing
                    End If
                Next
            Next
            SpCharValidation = ""
            Dim i As Integer
            For i = 0 To arrSpecialChar.Length - 1
                If InStr(StringValue, arrSpecialChar(i), CompareMethod.Binary) <> 0 Then
                    SpCharValidation = SpCharValidation & arrSpecialChar(i)
                End If
            Next

            Return SpCharValidation

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "SpCharValidation")

        End Try

    End Function

#Region " IDisposable Support "

    Public Sub Dispose() Implements IDisposable.Dispose

        If Not ObjBaseClass Is Nothing Then ObjBaseClass.Dispose()
        If Not DtValidation Is Nothing Then DtValidation.Dispose()
        ''Added by Jaiwant dtd  03-Dec-2010
        If Not DtSpCharValidation Is Nothing Then DtSpCharValidation.Dispose()
        ''----
        If Not DtInput Is Nothing Then DtInput.Dispose()
        If Not DtUnSucInput Is Nothing Then DtUnSucInput.Dispose()
        If Not DtTemp Is Nothing Then DtTemp.Dispose()

        ObjBaseClass = Nothing
        DtValidation = Nothing
        ''Added by Jaiwant dtd  03-Dec-2010
        DtSpCharValidation = Nothing
        ''----
        DtInput = Nothing
        DtUnSucInput = Nothing
        DtTemp = Nothing

        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class
