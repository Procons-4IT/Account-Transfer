Public Class clsActTransfer
    Inherits clsBase
    Private oGrid As SAPbouiCOM.Grid
    Private oGridColumn As SAPbouiCOM.GridColumn
    Private oDtJEList As SAPbouiCOM.DataTable
    Private oDtAccountList As SAPbouiCOM.DataTable
    Private oDtTransList_S As SAPbouiCOM.DataTable
    Private oDtTransList_S1 As SAPbouiCOM.DataTable
    Private oDtTransList_P As SAPbouiCOM.DataTable
    Private oDtCurrList As SAPbouiCOM.DataTable
    Private strQuery As String
    Private oEditText As SAPbouiCOM.EditText
    Private oEditTextCol As SAPbouiCOM.EditTextColumn
    Private oComboBox As SAPbouiCOM.ComboBox
    Private oRecordSet As SAPbobsCOM.Recordset
    Private oComboBox1 As SAPbouiCOM.ComboBox

    Private Enum strType As Integer
        ACT = 0
        AOF = 1
        ACR = 2
        ALL = 3
    End Enum

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_OAPT, frm_OAPT)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initializeDataSource(oForm)
            filterChooseFromList(oForm)
            fillCombo(oForm)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, False)
            oForm.EnableMenu(mnu_DELETE_ROW, False)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_OAPT Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "_1" Then
                                    If Not validate(oForm) Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        If oApplication.SBO_Application.MessageBox("Are you sure you want to post journal entries?", 1, "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf (pVal.ItemUID = "18") Then
                                    If validateLoad(oForm) Then
                                        loadAccountDetail(oForm, "11", "3")
                                        loadAccountDetail(oForm, "11", "_3")
                                        loadAccountDetail(oForm, "13", "34")
                                        loadJEDetail(oForm)
                                        'loadCurrencyDetail(oForm)
                                    End If
                                ElseIf (pVal.ItemUID = "19") Then
                                    AddRow(oForm)
                                ElseIf (pVal.ItemUID = "20") Then
                                    Delete(oForm)
                                ElseIf (pVal.ItemUID = "35") Then
                                    oForm.PaneLevel = 1
                                ElseIf (pVal.ItemUID = "36") Then
                                    oForm.PaneLevel = 2
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                If pVal.ItemUID = "1" Then
                                    If CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value.Length = 0 Or _
                                    CType(oForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value.Length = 0 Then
                                        oApplication.Utilities.Message("Select From Date & To Date to Proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                If pVal.ItemUID = "30" Then
                                    filterChooseFromListByList(oForm)
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "_1" Then
                                    post_JournalEntry(oForm)
                                    'post_JournalVoucher(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ItemUID = "1" And pVal.ColUID = "FormatCode" And Not IsNothing(oDataTable) Then
                                        If pVal.Before_Action = False Then
                                            oGrid = oForm.Items.Item("1").Specific

                                            'Dim intAddRow As Integer = oGrid.DataTable.Rows.Count
                                            'If intAddRow < oDataTable.Rows.Count Then
                                            '    intAddRow = oDataTable.Rows.Count - intAddRow
                                            '    oGrid.DataTable.Rows.Add(intAddRow + 1)
                                            'Else
                                            '    intAddRow = intAddRow - oDataTable.Rows.Count
                                            '    oGrid.DataTable.Rows.Add(intAddRow)
                                            'End If

                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oGrid.DataTable.SetValue("FormatCode", pVal.Row + index, oDataTable.GetValue("FormatCode", index))
                                                oGrid.DataTable.SetValue("AcctCode", pVal.Row + index, oDataTable.GetValue("AcctCode", index))
                                                oGrid.DataTable.SetValue("AcctName", pVal.Row + index, oDataTable.GetValue("AcctName", index))
                                                oGrid.DataTable.Rows.Add()
                                            Next
                                            oApplication.Utilities.assignLineNo(oGrid, oForm)
                                        End If
                                    ElseIf (pVal.ItemUID = "17") And Not IsNothing(oDataTable) Then
                                        oForm.DataSources.UserDataSources.Item("_17").ValueEx = oDataTable.GetValue("FormatCode", 0)
                                    ElseIf (pVal.ItemUID = "23") And Not IsNothing(oDataTable) Then
                                        oForm.DataSources.UserDataSources.Item("_23").ValueEx = oDataTable.GetValue("FormatCode", 0)
                                        oForm.DataSources.UserDataSources.Item("_27").ValueEx = oDataTable.GetValue("AcctName", 0)
                                    ElseIf (pVal.ItemUID = "25") And Not IsNothing(oDataTable) Then
                                        oForm.DataSources.UserDataSources.Item("_25").ValueEx = oDataTable.GetValue("FormatCode", 0)
                                        oForm.DataSources.UserDataSources.Item("_28").ValueEx = oDataTable.GetValue("AcctName", 0)
                                    ElseIf (pVal.ItemUID = "30") And Not IsNothing(oDataTable) Then
                                        oForm.DataSources.UserDataSources.Item("_30").ValueEx = oDataTable.GetValue("CardCode", 0)
                                        oForm.DataSources.UserDataSources.Item("_32").ValueEx = oDataTable.GetValue("CardName", 0)
                                    End If
                                Catch ex As Exception
                                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    reDrawForm(oForm)
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_OAPT
                    LoadForm()
            End Select
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oForm.Items.Item("33").TextStyle = 7
            oForm.Items.Item("37").TextStyle = 7

            'From Account Selection
            oDtAccountList = oForm.DataSources.DataTables.Add("dtAccountList")
            oGrid = oForm.Items.Item("1").Specific
            strQuery = "Select FormatCode,AcctCode,AcctName From OACT Where 1 = 2"
            oDtAccountList.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtAccountList
            formatPostingFromGrid(oForm)

            'Summary Grid for Selection Grid 
            oDtTransList_S = oForm.DataSources.DataTables.Add("dtTransList_S")
            oGrid = oForm.Items.Item("3").Specific
            strQuery = "Select T1.FormatCode,T1.AcctCode,T1.AcctName,Sum(Debit) - Sum(Credit) As 'Amount','' As FCCurrency,Sum(Debit) - Sum(Credit) As 'Amount_LC'  From  JDT1 T0 Join OACT T1 On T0.Account = T1.AcctCode Where 1 = 2"
            strQuery += " Group By T1.FormatCode,T1.AcctCode,T1.AcctName "
            oDtTransList_S.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtTransList_S
            formatTransactionGrid(oForm, "3", "11")

            'Summary Grid for Selection Grid 'Visible True Grid
            oDtTransList_S1 = oForm.DataSources.DataTables.Add("dtTransList_S1")
            oGrid = oForm.Items.Item("_3").Specific
            strQuery = "Select T1.FormatCode,T1.AcctCode,T1.AcctName,Sum(Debit) - Sum(Credit) As 'Amount','' As FCCurrency,Sum(Debit) - Sum(Credit) As 'Amount_LC',T0.LicTradNum 'Federal Tax ID',Sum(T0.BaseSum) 'BaseAmount'  From  JDT1 T0 Join OACT T1 On T0.Account = T1.AcctCode Where 1 = 2"
            strQuery += " Group By T1.FormatCode,T1.AcctCode,T1.AcctName,T0.LicTradNum "
            oDtTransList_S1.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtTransList_S1
            formatTransactionGrid(oForm, "3", "11")

            'Summary Grid for Posting Grid
            oDtTransList_P = oForm.DataSources.DataTables.Add("dtTransList_P")
            oGrid = oForm.Items.Item("34").Specific
            strQuery = "Select T1.FormatCode,T1.AcctCode,T1.AcctName,Sum(Debit) - Sum(Credit) As 'Amount','' As FCCurrency,Sum(Debit) - Sum(Credit) As 'Amount_LC' From  JDT1 T0 Join OACT T1 On T0.Account = T1.AcctCode Where 1 = 2"
            strQuery += " Group By T1.FormatCode,T1.AcctCode,T1.AcctName "
            oDtTransList_P.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtTransList_P
            formatTransactionGrid(oForm, "34", "13")

            oDtJEList = oForm.DataSources.DataTables.Add("dtJEList")
            oDtCurrList = oForm.DataSources.DataTables.Add("dtCurrList")
            oForm.PaneLevel = 1

            oRecordSet.DoQuery("Select Convert(VarChar(8),U_ATF,112) As 'ATF' From [@OATF] Where ISNULL(U_ATF,'') <> ''")
            If Not oRecordSet.EoF Then
                Dim strTrnsFrom As String = oRecordSet.Fields.Item(0).Value
                If strTrnsFrom.Length > 0 Then
                    oForm.Items.Item("4").Visible = False
                    oForm.Items.Item("5").Specific.value = strTrnsFrom
                    oForm.Items.Item("5").Width = 1
                    oForm.Items.Item("5").Height = 1
                End If
            End If

            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub initializeDataSource(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.DataSources.UserDataSources.Add("_17", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 210)
            oForm.DataSources.UserDataSources.Add("_23", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 210)
            oForm.DataSources.UserDataSources.Add("_27", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200)
            oForm.DataSources.UserDataSources.Add("_25", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 210)
            oForm.DataSources.UserDataSources.Add("_28", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200)
            oForm.DataSources.UserDataSources.Add("_30", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            oForm.DataSources.UserDataSources.Add("_32", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            CType(oForm.Items.Item("17").Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "", "_17")
            CType(oForm.Items.Item("23").Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "", "_23")
            CType(oForm.Items.Item("27").Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "", "_27")
            CType(oForm.Items.Item("25").Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "", "_25")
            CType(oForm.Items.Item("28").Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "", "_28")
            CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "", "_30")
            CType(oForm.Items.Item("32").Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "", "_32")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function validateLoad(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        oGrid = oForm.Items.Item("1").Specific
        oComboBox = oForm.Items.Item("11").Specific
        Try
            'If oForm.Items.Item("5").Specific.value.ToString().Length = 0 Then
            '    'oApplication.Utilities.Message("Select Posting From Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    '_retVal = False
            'ElseIf oForm.Items.Item("7").Specific.value.ToString().Length = 0 Then
            '    'oApplication.Utilities.Message("Select Posting To Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    '_retVal = False
            'End If
            If oForm.Items.Item("23").Specific.value.ToString().Length = 0 Then
                oApplication.Utilities.Message("Select Posting From Account...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf oForm.Items.Item("25").Specific.value.ToString().Length = 0 Then
                oApplication.Utilities.Message("Select Posting To Account...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf (oForm.Items.Item("5").Specific.value > oForm.Items.Item("7").Specific.value) Then
                oApplication.Utilities.Message("Posting From Date Should be less than or equal to Posting To Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
                'ElseIf (oGrid.DataTable.Rows.Count = 0) Then
                '    oApplication.Utilities.Message("Select Posting Account to Filter...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
            ElseIf (oComboBox.Selected.Value.Length = 0) Then
                oApplication.Utilities.Message("Select Selection Group...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CType(oForm.Items.Item("13").Specific, SAPbouiCOM.ComboBox).Selected.Value.Length = 0 Then
                oApplication.Utilities.Message("Select Posting To Type...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Private Function validate(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        oGrid = oForm.Items.Item("3").Specific
        Try
            If oForm.Items.Item("9").Specific.value.ToString().Length = 0 Then
                oApplication.Utilities.Message("Select Posting Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf oForm.Items.Item("17").Specific.value.ToString().Length = 0 Then
                oApplication.Utilities.Message("Select Posting To Account...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf (oGrid.DataTable.Rows.Count = 0) Then
                oApplication.Utilities.Message("No Records for Posting...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CType(oForm.Items.Item("13").Specific, SAPbouiCOM.ComboBox).Selected.Value.Length = 0 Then
                oApplication.Utilities.Message("Select Posting To Type...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf (oForm.Items.Item("9").Specific.value.ToString().Length > 0) Then
                strQuery = "Select PeriodStat From OFPR Where Convert(VarChar(8),F_RefDate,112) <= '" + oForm.Items.Item("9").Specific.value.ToString() + "' And Convert(VarChar(8),T_RefDate,112) >= '" + oForm.Items.Item("9").Specific.value.ToString() + "'"
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If Not oRecordSet.EoF Then
                    If oRecordSet.Fields.Item(0).Value = "Y" Then
                        oApplication.Utilities.Message("Entered Posting To Period is Locked ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        _retVal = False
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Private Sub filterChooseFromList(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            oCFLs = oForm.ChooseFromLists

            oCFL = oCFLs.Item("CFL_1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_3")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_4")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'oCFL = oCFLs.Item("CFL_5")
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "Postable"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "Y"
            'oCFL.SetConditions(oCons)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub filterChooseFromListByList(ByVal oForm As SAPbouiCOM.Form)
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists
            oCFL = oCFLs.Item("CFL_5")
            oCons = oCFL.GetConditions()
            If oCons.Count = 0 Then
                oCon = oCons.Add()
                strQuery = "Select CardCode From OCRD Where CardCode IN (Select ContraAct From JDT1)"
                oRecordSet.DoQuery(strQuery)
                Dim intRecordCount As Integer = 0
                If Not oRecordSet.EoF Then
                    While Not oRecordSet.EoF
                        If intRecordCount > 0 Then
                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                            oCon = oCons.Add()
                        End If
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = oRecordSet.Fields.Item(0).Value
                        intRecordCount += 1
                        oRecordSet.MoveNext()
                    End While
                End If
                oCFL.SetConditions(oCons)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub fillCombo(ByVal aForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oComboBox = oForm.Items.Item("11").Specific
            oComboBox.ValidValues.Add("", "")
            oComboBox.ValidValues.Add("0", "Account")
            oComboBox.ValidValues.Add("1", "Account/Offset")
            oComboBox.ValidValues.Add("2", "Account/Currency")
            oComboBox.ValidValues.Add("3", "Account/Offset/Currency")
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oForm.Items.Item("11").DisplayDesc = True
            oComboBox.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue)

            oComboBox = oForm.Items.Item("13").Specific
            oComboBox.ValidValues.Add("", "")
            oComboBox.ValidValues.Add("0", "Account")
            oComboBox.ValidValues.Add("1", "Account/Offset")
            oComboBox.ValidValues.Add("2", "Account/Currency")
            oComboBox.ValidValues.Add("3", "Account/Offset/Currency")
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oComboBox.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("13").DisplayDesc = True
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub loadAccountDetail(ByVal aForm As SAPbouiCOM.Form, ByVal strGroupTypeID As String, ByVal strGridID As String)
        oForm.Freeze(True)
        Try
            Dim strFrmAcct, strToAcct, strOffset As String
            'strFrmAcct = oApplication.Utilities.getAccount(oForm.DataSources.UserDataSources.Item("_23").ValueEx)
            'strToAcct = oApplication.Utilities.getAccount(oForm.DataSources.UserDataSources.Item("_25").ValueEx)
            strFrmAcct = oForm.DataSources.UserDataSources.Item("_23").ValueEx
            strToAcct = oForm.DataSources.UserDataSources.Item("_25").ValueEx
            strOffset = oForm.DataSources.UserDataSources.Item("_30").ValueEx
            Dim strToAccount As String = oForm.Items.Item("17").Specific.value

            oComboBox = oForm.Items.Item(strGroupTypeID).Specific
            Select Case oComboBox.Selected.Value
                Case strType.ACT
                    strQuery = "Select T2.FormatCode,T2.AcctCode,T2.AcctName,Sum(Debit) - Sum(Credit) As 'Amount','' As FCCurrency,Sum(Debit) - Sum(Credit) As 'Amount_LC' "
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum 'Federal Tax ID',Sum(T0.BaseSum) 'BaseAmount' "
                    End If
                    strQuery += " From JDT1 T0 Join OJDT T1 On T1.TransID = T0.TransID Join OACT T2 On T0.Account = T2.AcctCode "
                    If strGridID = "_3" Then
                        strQuery += " Left Outer Join OCRD T3 On T3.CardCode = T0.ContraAct "
                    End If
                    'strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "' And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    If oForm.Items.Item("5").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "'"
                    End If
                    If oForm.Items.Item("7").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    End If
                    strQuery += " And ISNULL(T0.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T1.U_ActTra,'N') = 'N' "
                    strQuery += " Group By T2.FormatCode,T2.AcctCode,T2.AcctName "
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum  "
                    End If
                    If strOffset.Length > 0 Then
                        strQuery += ",T0.ContraAct "
                    End If
                    strQuery += " Having T2.FormatCode BetWeen '" + strFrmAcct + "' And '" + strToAcct + "'"
                    If strOffset.Length > 0 Then
                        strQuery += " And T0.ContraAct = '" + strOffset + "'"
                    End If
                    strQuery += " And Sum(Debit) - Sum(Credit) <> 0 "
                    strQuery += " Order By T2.FormatCode "
                Case strType.AOF
                    strQuery = "Select T2.FormatCode,T2.AcctCode,T2.AcctName,T0.ContraAct,Sum(Debit) - Sum(Credit) As 'Amount','' As FCCurrency,Sum(Debit) - Sum(Credit) As 'Amount_LC' "
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum 'Federal Tax ID',Sum(T0.BaseSum) 'BaseAmount' "
                    End If
                    strQuery += " From JDT1 T0 Join OJDT T1 On T1.TransID = T0.TransID Join OACT T2 On T0.Account = T2.AcctCode "
                    If strGridID = "_3" Then
                        strQuery += " Left Outer Join OCRD T3 On T3.CardCode = T0.ContraAct "
                    End If
                    'strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "' And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    If oForm.Items.Item("5").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "'"
                    End If
                    If oForm.Items.Item("7").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    End If
                    strQuery += " And ISNULL(T0.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T1.U_ActTra,'N') = 'N' "
                    strQuery += " Group By T2.FormatCode,T2.AcctCode,T2.AcctName,T0.ContraAct "
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum  "
                    End If
                    strQuery += " Having T2.FormatCode BetWeen '" + strFrmAcct + "' And '" + strToAcct + "'"
                    If strOffset.Length > 0 Then
                        strQuery += " And T0.ContraAct = '" + strOffset + "'"
                    End If
                    strQuery += " And Sum(Debit) - Sum(Credit) <> 0 "
                    strQuery += " Order By T2.FormatCode "
                Case strType.ACR
                    strQuery = "Select T0.* From ( "
                    strQuery += "Select T2.FormatCode,T2.AcctCode,T2.AcctName,ISNULL(T0.FCCurrency,'') As 'FCCurrency',Sum(Debit) - Sum(Credit) As 'Amount',Sum(Debit) - Sum(Credit) As 'Amount_LC' "
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum 'Federal Tax ID',Sum(T0.BaseSum) 'BaseAmount' "
                    End If
                    strQuery += " From JDT1 T0 Join OJDT T1 On T1.TransID = T0.TransID Join OACT T2 On T0.Account = T2.AcctCode "
                    If strGridID = "_3" Then
                        strQuery += " Left Outer Join OCRD T3 On T3.CardCode = T0.ContraAct "
                    End If
                    'strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "' And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    If oForm.Items.Item("5").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "'"
                    End If
                    If oForm.Items.Item("7").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    End If
                    strQuery += " And ISNULL(T0.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T1.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T0.FCCurrency,'') = '' "
                    strQuery += " Group By T2.FormatCode,T2.AcctCode,T2.AcctName,ISNULL(T0.FCCurrency,'') "
                    If strOffset.Length > 0 Then
                        strQuery += ",T0.ContraAct "
                    End If
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum  "
                    End If
                    strQuery += " Having T2.FormatCode BetWeen '" + strFrmAcct + "' And '" + strToAcct + "'"
                    If strOffset.Length > 0 Then
                        strQuery += " And T0.ContraAct = '" + strOffset + "'"
                    End If
                    strQuery += " And Sum(Debit) - Sum(Credit) <> 0 "
                    strQuery += " Union All "
                    strQuery += " Select T2.FormatCode,T2.AcctCode,T2.AcctName,ISNULL(T0.FCCurrency,'') As 'FCCurrency',Sum(FCDebit) - Sum(FCCredit) As 'Amount',Sum(Debit) - Sum(Credit) As 'Amount_LC' "
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum 'Federal Tax ID',Sum(T0.BaseSum) 'BaseAmount' "
                    End If
                    strQuery += " From JDT1 T0 Join OJDT T1 On T1.TransID = T0.TransID Join OACT T2 On T0.Account = T2.AcctCode "
                    If strGridID = "_3" Then
                        strQuery += " Left Outer Join OCRD T3 On T3.CardCode = T0.ContraAct "
                    End If
                    'strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "' And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    If oForm.Items.Item("5").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "'"
                    End If
                    If oForm.Items.Item("7").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    End If
                    strQuery += " And ISNULL(T0.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T1.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T0.FCCurrency,'') <> '' "
                    strQuery += " Group By T2.FormatCode,T2.AcctCode,T2.AcctName,ISNULL(T0.FCCurrency,'') "
                    If strOffset.Length > 0 Then
                        strQuery += ",T0.ContraAct "
                    End If
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum  "
                    End If
                    strQuery += " Having T2.FormatCode BetWeen '" + strFrmAcct + "' And '" + strToAcct + "'"
                    If strOffset.Length > 0 Then
                        strQuery += " And T0.ContraAct = '" + strOffset + "'"
                    End If
                    strQuery += " And Sum(FCDebit) - Sum(FCCredit) <> 0 "
                    strQuery += " ) T0 Order By FormatCode"
                Case strType.ALL
                    strQuery = "Select T0.* From ( "
                    strQuery += "Select T2.FormatCode,T2.AcctCode,T2.AcctName,T0.ContraAct,ISNULL(T0.FCCurrency,'') As 'FCCurrency',Sum(Debit) - Sum(Credit) As 'Amount',Sum(Debit) - Sum(Credit) As 'Amount_LC' "
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum 'Federal Tax ID',Sum(T0.BaseSum) 'BaseAmount' "
                    End If
                    strQuery += " From JDT1 T0 Join OJDT T1 On T1.TransID = T0.TransID Join OACT T2 On T0.Account = T2.AcctCode "
                    If strGridID = "_3" Then
                        strQuery += " Left Outer Join OCRD T3 On T3.CardCode = T0.ContraAct "
                    End If
                    'strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "' And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    If oForm.Items.Item("5").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "'"
                    End If
                    If oForm.Items.Item("7").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    End If
                    strQuery += " And ISNULL(T0.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T1.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T0.FCCurrency,'') = '' "
                    strQuery += " Group By T2.FormatCode,T2.AcctCode,T2.AcctName,T0.ContraAct,ISNULL(T0.FCCurrency,'') "
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum  "
                    End If
                    strQuery += " Having T2.FormatCode BetWeen '" + strFrmAcct + "' And '" + strToAcct + "'"
                    If strOffset.Length > 0 Then
                        strQuery += " And T0.ContraAct = '" + strOffset + "'"
                    End If
                    strQuery += " And Sum(Debit) - Sum(Credit) <> 0 "
                    strQuery += " Union All "
                    strQuery += " Select T2.FormatCode,T2.AcctCode,T2.AcctName,T0.ContraAct,ISNULL(T0.FCCurrency,'') As 'FCCurrency',Sum(FCDebit) - Sum(FCCredit) As 'Amount',Sum(Debit) - Sum(Credit) As 'Amount_LC' "
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum 'Federal Tax ID',Sum(T0.BaseSum) 'BaseAmount' "
                    End If
                    strQuery += " From JDT1 T0 Join OJDT T1 On T1.TransID = T0.TransID Join OACT T2 On T0.Account = T2.AcctCode "
                    If strGridID = "_3" Then
                        strQuery += " Left Outer Join OCRD T3 On T3.CardCode = T0.ContraAct "
                    End If
                    'strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "' And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    If oForm.Items.Item("5").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "'"
                    End If
                    If oForm.Items.Item("7").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    End If
                    strQuery += " And ISNULL(T0.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T1.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T0.FCCurrency,'') <> '' "
                    strQuery += " Group By T2.FormatCode,T2.AcctCode,T2.AcctName,T0.ContraAct,ISNULL(T0.FCCurrency,'') "
                    If strGridID = "_3" Then
                        strQuery += " ,T3.LicTradNum  "
                    End If
                    strQuery += " Having T2.FormatCode BetWeen '" + strFrmAcct + "' And '" + strToAcct + "'"
                    If strOffset.Length > 0 Then
                        strQuery += " And T0.ContraAct = '" + strOffset + "'"
                    End If
                    strQuery += " And Sum(FCDebit) - Sum(FCCredit) <> 0 "
                    strQuery += " ) T0 Order By FormatCode"
            End Select
            oGrid = oForm.Items.Item(strGridID).Specific

            If strGridID = "34" Then
                Dim oInRecordSet As SAPbobsCOM.Recordset
                oInRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim strPostingType As String = String.Empty
                Dim strFromDate As String = String.Empty
                Dim strToDate As String = String.Empty
                Dim strFromFormat As String = String.Empty
                Dim strToFormat As String = String.Empty
                Dim strPostFormat As String = String.Empty
                Dim strOffset1 As String = String.Empty
                Dim strMainQuery As String = String.Empty

                strPostingType = oComboBox.Selected.Value
                strFromDate = oForm.Items.Item("5").Specific.value
                strToDate = oForm.Items.Item("7").Specific.value
                strFromFormat = oForm.DataSources.UserDataSources.Item("_23").ValueEx
                strToFormat = oForm.DataSources.UserDataSources.Item("_25").ValueEx
                strPostFormat = oForm.Items.Item("17").Specific.value
                strOffset1 = oForm.DataSources.UserDataSources.Item("_30").ValueEx
                Dim strQuery = "Exec Procon_AccountTransfer_Posting '" + strPostingType + "','" + strFromDate + "','" + strToDate + "','" + strFromFormat + "','" + strToFormat + "','" + strPostFormat + "','" + strOffset1 + "'"
                oInRecordSet.DoQuery(strQuery)
                Select Case oComboBox.Selected.Value
                    Case strType.ACT
                        strMainQuery = "Select T2.FormatCode,T2.AcctCode,T2.AcctName,Sum(T2.Amount) As 'Amount','' As FCCurrency,Sum(T2.Amount_LC) As 'Amount_LC' From Z_POSD T2"
                        strMainQuery += " Group By T2.FormatCode,T2.AcctCode,T2.AcctName "
                    Case strType.AOF
                        strMainQuery = "Select T2.FormatCode,T2.AcctCode,T2.AcctName,T2.Offset As 'ContraAct',Sum(T2.Amount) As 'Amount','' As FCCurrency,Sum(T2.Amount_LC) As 'Amount_LC' From Z_POSD T2"
                        strMainQuery += " Group By T2.FormatCode,T2.AcctCode,T2.AcctName,T2.Offset "
                    Case strType.ACR
                        strMainQuery = "Select T2.FormatCode,T2.AcctCode,T2.AcctName,Sum(T2.Amount) As 'Amount',ISNULL(T2.FCCurrency,'') As FCCurrency,Sum(T2.Amount_LC) As 'Amount_LC' From Z_POSD T2"
                        strMainQuery += " Group By T2.FormatCode,T2.AcctCode,T2.AcctName,ISNULL(T2.FCCurrency,'') "
                    Case strType.ALL
                        strMainQuery = "Select T2.FormatCode,T2.AcctCode,T2.AcctName,T2.Offset AS 'ContraAct',Sum(T2.Amount) As 'Amount',ISNULL(T2.FCCurrency,'') As FCCurrency,Sum(T2.Amount_LC) As 'Amount_LC' From Z_POSD T2"
                        strMainQuery += " Group By T2.FormatCode,T2.AcctCode,T2.Offset,T2.AcctName,ISNULL(T2.FCCurrency,'') "
                End Select
                oGrid.DataTable.ExecuteQuery(strMainQuery)
            Else
                oGrid.DataTable.ExecuteQuery(strQuery)
            End If

            formatTransactionGrid(oForm, strGridID, strGroupTypeID)
            If oGrid.Rows.Count < 2 Then
                If (oGrid.DataTable.GetValue("AcctCode", 0) = "") Then
                    oApplication.Utilities.Message("No Record Found", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub loadJEDetail(ByVal aForm As SAPbouiCOM.Form)
        oForm.Freeze(True)
        Try
            Dim strFrmAcct, strToAcct, strOffset As String
            'strFrmAcct = oApplication.Utilities.getAccount(oForm.DataSources.UserDataSources.Item("_23").ValueEx)
            'strToAcct = oApplication.Utilities.getAccount(oForm.DataSources.UserDataSources.Item("_25").ValueEx)
            strFrmAcct = oForm.DataSources.UserDataSources.Item("_23").ValueEx
            strToAcct = oForm.DataSources.UserDataSources.Item("_25").ValueEx
            strOffset = oForm.DataSources.UserDataSources.Item("_30").ValueEx

            strQuery = "Select T0.TransID,T0.Line_ID "
            strQuery += " From JDT1 T0 Join OJDT T1 On T1.TransID = T0.TransID Join OACT T2 On T0.Account = T2.AcctCode "
            'strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "' And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
            If oForm.Items.Item("5").Specific.value.ToString().Length > 0 Then
                strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "'"
            End If
            If oForm.Items.Item("7").Specific.value.ToString().Length > 0 Then
                strQuery += " And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
            End If
            strQuery += " And ISNULL(T0.U_ActTra,'N') = 'N' "
            strQuery += " And ISNULL(T1.U_ActTra,'N') = 'N' "

            oGrid = oForm.Items.Item("1").Specific
            oDtJEList = oForm.DataSources.DataTables.Item("dtJEList")
            strQuery += " Where T2.FormatCode BetWeen '" + strFrmAcct + "' And '" + strToAcct + "'"
            If strOffset.Length > 0 Then
                strQuery += " And T0.ContraAct = '" + strOffset + "'"
            End If
            oDtJEList.ExecuteQuery(strQuery)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub loadCurrencyDetail(ByVal aForm As SAPbouiCOM.Form)
        oForm.Freeze(True)
        Try
            Dim strFrmAcct, strToAcct, strOffset As String
            'strFrmAcct = oApplication.Utilities.getAccount(oForm.DataSources.UserDataSources.Item("_23").ValueEx)
            'strToAcct = oApplication.Utilities.getAccount(oForm.DataSources.UserDataSources.Item("_25").ValueEx)
            strFrmAcct = oForm.DataSources.UserDataSources.Item("_23").ValueEx
            strToAcct = oForm.DataSources.UserDataSources.Item("_25").ValueEx
            strOffset = oForm.DataSources.UserDataSources.Item("_30").ValueEx

            oComboBox = oForm.Items.Item("13").Specific
            oDtCurrList = oForm.DataSources.DataTables.Item("dtCurrList")

            Select Case oComboBox.Selected.Value
                Case strType.ACT, strType.AOF
                    strQuery = "Select '' As FCCurrency  From OADM"
                Case strType.ACR, strType.ALL
                    strQuery = "Select Distinct T0.FCCurrency "
                    strQuery += " From JDT1 T0 Join OJDT T1 On T1.TransID = T0.TransID Join OACT T2 On T0.Account = T2.AcctCode "
                    'strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "' And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    If oForm.Items.Item("5").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) >= '" + oForm.Items.Item("5").Specific.value + "'"
                    End If
                    If oForm.Items.Item("7").Specific.value.ToString().Length > 0 Then
                        strQuery += " And Convert(VarChar(8),T0.RefDate,112) <= '" + oForm.Items.Item("7").Specific.value + "'"
                    End If
                    strQuery += " And ISNULL(T0.U_ActTra,'N') = 'N' "
                    strQuery += " And ISNULL(T1.U_ActTra,'N') = 'N' "

                    oGrid = oForm.Items.Item("1").Specific

                    strQuery += " Where T2.FormatCode BetWeen '" + strFrmAcct + "' And '" + strToAcct + "'"
                    strQuery += " AND ISNULL(FCCurrency,'') <> '' "
                    If strOffset.Length > 0 Then
                        strQuery += " And T0.ContraAct = '" + strOffset + "'"
                    End If
                    strQuery += "  Union All  Select '' As FCCurrency  From OADM "
            End Select

            oDtCurrList.ExecuteQuery(strQuery)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Public Function post_JournalEntry(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        Dim strTable As String = String.Empty

        Try
            Dim dblCredit_S As Double
            Dim dblDebit_S As Double

            Dim oJE As SAPbobsCOM.JournalEntries
            Dim strToAccount As String = oForm.Items.Item("17").Specific.value
            Dim dtPostingDt As DateTime = oApplication.Utilities.GetDateTimeValue(oForm.Items.Item("9").Specific.value)

            Dim dblSummaryAmt_LC As Double = 0
            Dim dblSummaryAmt_FC As Double = 0
            Dim dblSummaryAmt_SC As Double = 0

            Dim intCurrentLine As Integer = 0
            Dim blnRecordExist As Boolean = False
            oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            oJE.ReferenceDate = dtPostingDt
            oJE.TaxDate = dtPostingDt
            oJE.DueDate = dtPostingDt
            oJE.UserFields.Fields.Item("U_ActTra").Value = "Y"

            oJE.Memo = "Account Posting"
            oJE.Reference = "Account Posting"
            oJE.Reference2 = "Account Posting"
            oJE.Reference3 = "Account Posting"

            oGrid = oForm.Items.Item("3").Specific
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                blnRecordExist = True
                Dim strFcCurrency As String = String.Empty
                Dim strOffset As String = String.Empty

                'Set Offset Account.
                oComboBox = oForm.Items.Item("11").Specific
                Select Case oComboBox.Selected.Value
                    Case strType.ACT
                        strOffset = String.Empty
                    Case strType.AOF
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                    Case strType.ACR
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                    Case strType.ALL
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                End Select

                dblSummaryAmt_FC = oGrid.DataTable.GetValue("Amount", index)
                dblSummaryAmt_LC = oGrid.DataTable.GetValue("Amount_LC", index)

                If dblSummaryAmt_FC < 0 Then
                    If intCurrentLine > 0 Then
                        oJE.Lines.Add()
                    End If
                    oJE.Lines.SetCurrentLine(intCurrentLine)
                    If strOffset.Length > 0 Then
                        oJE.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    oJE.Lines.AccountCode = oGrid.DataTable.GetValue("AcctCode", index)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.Lines.FCCurrency = strFcCurrency
                        oJE.Lines.FCDebit = (dblSummaryAmt_FC * -1)
                        oJE.Lines.Debit = (dblSummaryAmt_LC * -1)
                    Else
                        'oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.Lines.Debit = (dblSummaryAmt_FC * -1)
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblDebit_S += oGrid.DataTable.GetValue("Amount_LC", index) * -1
                Else
                    If intCurrentLine > 0 Then
                        oJE.Lines.Add()
                    End If
                    oJE.Lines.SetCurrentLine(intCurrentLine)
                    If strOffset.Length > 0 Then
                        oJE.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    oJE.Lines.AccountCode = oGrid.DataTable.GetValue("AcctCode", index)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.Lines.FCCurrency = strFcCurrency
                        oJE.Lines.FCCredit = dblSummaryAmt_FC
                        oJE.Lines.Credit = dblSummaryAmt_LC
                    Else
                        'oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.Lines.Credit = dblSummaryAmt_FC
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblCredit_S += oGrid.DataTable.GetValue("Amount_LC", index)
                End If
            Next

            oGrid = oForm.Items.Item("34").Specific
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                Dim strFcCurrency As String = String.Empty
                Dim strOffset As String = String.Empty

                'Set Offset Account.
                oComboBox = oForm.Items.Item("13").Specific
                Select Case oComboBox.Selected.Value
                    Case strType.ACT
                        strOffset = String.Empty
                    Case strType.AOF
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                    Case strType.ACR
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                    Case strType.ALL
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                End Select

                dblSummaryAmt_FC = oGrid.DataTable.GetValue("Amount", index)
                dblSummaryAmt_LC = oGrid.DataTable.GetValue("Amount_LC", index)

                If dblSummaryAmt_FC < 0 Then
                    If intCurrentLine > 0 Then
                        oJE.Lines.Add()
                    End If
                    oJE.Lines.SetCurrentLine(intCurrentLine)
                    oJE.Lines.AccountCode = oApplication.Utilities.getAccount(strToAccount)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.Lines.FCCurrency = strFcCurrency
                        oJE.Lines.FCCredit = (dblSummaryAmt_FC * -1)
                        oJE.Lines.Credit = (dblSummaryAmt_LC * -1)
                    Else
                        'oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.Lines.Credit = (dblSummaryAmt_FC * -1)
                    End If
                    If strOffset.Length > 0 Then
                        oJE.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblCredit_S += oGrid.DataTable.GetValue("Amount_LC", index) * -1
                Else
                    If intCurrentLine > 0 Then
                        oJE.Lines.Add()
                    End If
                    oJE.Lines.SetCurrentLine(intCurrentLine)
                    oJE.Lines.AccountCode = oApplication.Utilities.getAccount(strToAccount)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.Lines.FCCurrency = strFcCurrency
                        oJE.Lines.FCDebit = dblSummaryAmt_FC
                        oJE.Lines.Debit = dblSummaryAmt_LC
                    Else
                        'oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.Lines.Debit = dblSummaryAmt_FC
                    End If
                    If strOffset.Length > 0 Then
                        oJE.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblDebit_S += oGrid.DataTable.GetValue("Amount_LC", index)
                End If
            Next

            'If dblCredit_S <> dblDebit_S Then
            '    Dim dblAdjust As Double
            '    If intCurrentLine > 0 Then
            '        oJE.Lines.Add()
            '    End If
            '    oJE.Lines.SetCurrentLine(intCurrentLine)
            '    oJE.Lines.AccountCode = oApplication.Utilities.getAccount(strToAccount)
            '    If dblCredit_S > dblDebit_S Then
            '        dblAdjust = dblCredit_S - dblDebit_S
            '        oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
            '        oJE.Lines.Debit = dblAdjust
            '        dblDebit_S += dblAdjust
            '    ElseIf (dblDebit_S > dblCredit_S) Then
            '        dblAdjust = dblDebit_S - dblCredit_S
            '        oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
            '        oJE.Lines.Credit = dblAdjust
            '        dblCredit_S += dblAdjust
            '    End If
            'End If

            'MessageBox.Show(dblCredit_S.ToString())
            'MessageBox.Show(dblDebit_S.ToString())
            'Throw New Exception("Error")

            'if Record Exist...
            If blnRecordExist Then
                Dim intCode As Integer = oJE.Add()
                If intCode = 0 Then
                    _retVal = True
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    For index As Integer = 0 To oDtJEList.Rows.Count - 1
                        strQuery = "Update JDT1 Set U_ActTra = 'Y' Where TransID = '" + oDtJEList.GetValue("TransID", index).ToString() + "' And Line_ID = '" + oDtJEList.GetValue("Line_ID", index).ToString() + "'"
                        oRecordSet.DoQuery(strQuery)
                    Next
                    Dim strTransID As String = oApplication.Company.GetNewObjectKey()
                    If strTransID > 0 Then
                        strQuery = "Update JDT1 Set ContraAct = U_Offset Where TransID = '" + strTransID + "' And ISNULL(U_OffSet,'') <> ''"
                        oRecordSet.DoQuery(strQuery)
                    End If
                    oApplication.SBO_Application.MessageBox("Journal Entry Posted Successfully...", 1, "Ok", "", "")
                    clearSource(oForm)
                Else
                    _retVal = False
                    Throw New Exception(oApplication.Company.GetLastErrorDescription())
                End If
            End If
        Catch ex As Exception
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        Return _retVal
    End Function
    Public Function post_JournalEntry_Old(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        Dim strTable As String = String.Empty

        Try
            Dim dblCredit_S As Double
            Dim dblDebit_S As Double

            Dim oJE As SAPbobsCOM.JournalEntries
            Dim strToAccount As String = oForm.Items.Item("17").Specific.value
            Dim dtPostingDt As DateTime = oApplication.Utilities.GetDateTimeValue(oForm.Items.Item("9").Specific.value)

            Dim dblSummaryAmt_LC As Double = 0
            Dim dblSummaryAmt_FC As Double = 0
            Dim dblSummaryAmt_SC As Double = 0

            Dim intCurrentLine As Integer = 0
            Dim blnRecordExist As Boolean = False
            oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            oJE.ReferenceDate = dtPostingDt
            oJE.TaxDate = dtPostingDt
            oJE.DueDate = dtPostingDt
            oJE.UserFields.Fields.Item("U_ActTra").Value = "Y"

            oJE.Memo = "Account Posting"
            oJE.Reference = "Account Posting"
            oJE.Reference2 = "Account Posting"
            oJE.Reference3 = "Account Posting"

            oGrid = oForm.Items.Item("3").Specific
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                blnRecordExist = True
                Dim strFcCurrency As String = String.Empty
                Dim strOffset As String = String.Empty

                'Set Offset Account.
                oComboBox = oForm.Items.Item("11").Specific
                Select Case oComboBox.Selected.Value
                    Case strType.ACT
                        strOffset = String.Empty
                    Case strType.AOF
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                    Case strType.ACR
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                    Case strType.ALL
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                End Select

                dblSummaryAmt_FC = oGrid.DataTable.GetValue("Amount", index)
                dblSummaryAmt_LC = oGrid.DataTable.GetValue("Amount_LC", index)

                If dblSummaryAmt_FC < 0 Then
                    If intCurrentLine > 0 Then
                        oJE.Lines.Add()
                    End If
                    oJE.Lines.SetCurrentLine(intCurrentLine)
                    If strOffset.Length > 0 Then
                        oJE.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    oJE.Lines.AccountCode = oGrid.DataTable.GetValue("AcctCode", index)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.Lines.FCCurrency = strFcCurrency
                        oJE.Lines.FCDebit = (dblSummaryAmt_FC * -1)
                        oJE.Lines.Debit = (dblSummaryAmt_LC * -1)
                    Else
                        oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.Lines.Debit = (dblSummaryAmt_FC * -1)
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblDebit_S += oGrid.DataTable.GetValue("Amount_LC", index) * -1
                Else
                    If intCurrentLine > 0 Then
                        oJE.Lines.Add()
                    End If
                    oJE.Lines.SetCurrentLine(intCurrentLine)
                    If strOffset.Length > 0 Then
                        oJE.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    oJE.Lines.AccountCode = oGrid.DataTable.GetValue("AcctCode", index)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.Lines.FCCurrency = strFcCurrency
                        oJE.Lines.FCCredit = dblSummaryAmt_FC
                        oJE.Lines.Credit = dblSummaryAmt_LC
                    Else
                        oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.Lines.Credit = dblSummaryAmt_FC
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblCredit_S += oGrid.DataTable.GetValue("Amount_LC", index)
                End If
            Next

            oGrid = oForm.Items.Item("34").Specific
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                Dim strFcCurrency As String = String.Empty
                Dim strOffset As String = String.Empty

                'Set Offset Account.
                oComboBox = oForm.Items.Item("13").Specific
                Select Case oComboBox.Selected.Value
                    Case strType.ACT
                        strOffset = String.Empty
                    Case strType.AOF
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                    Case strType.ACR
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                    Case strType.ALL
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                End Select

                dblSummaryAmt_FC = oGrid.DataTable.GetValue("Amount", index)
                dblSummaryAmt_LC = oGrid.DataTable.GetValue("Amount_LC", index)

                If dblSummaryAmt_FC < 0 Then
                    If intCurrentLine > 0 Then
                        oJE.Lines.Add()
                    End If
                    oJE.Lines.SetCurrentLine(intCurrentLine)
                    oJE.Lines.AccountCode = oApplication.Utilities.getAccount(strToAccount)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.Lines.FCCurrency = strFcCurrency
                        oJE.Lines.FCCredit = (dblSummaryAmt_FC * -1)
                        oJE.Lines.Credit = (dblSummaryAmt_LC * -1)
                    Else
                        oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.Lines.Credit = (dblSummaryAmt_FC * -1)
                    End If
                    If strOffset.Length > 0 Then
                        oJE.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblCredit_S += oGrid.DataTable.GetValue("Amount_LC", index) * -1
                Else
                    If intCurrentLine > 0 Then
                        oJE.Lines.Add()
                    End If
                    oJE.Lines.SetCurrentLine(intCurrentLine)
                    oJE.Lines.AccountCode = oApplication.Utilities.getAccount(strToAccount)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.Lines.FCCurrency = strFcCurrency
                        oJE.Lines.FCDebit = dblSummaryAmt_FC
                        oJE.Lines.Debit = dblSummaryAmt_LC
                    Else
                        oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.Lines.Debit = dblSummaryAmt_FC
                    End If
                    If strOffset.Length > 0 Then
                        oJE.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblDebit_S += oGrid.DataTable.GetValue("Amount_LC", index)
                End If
            Next

            'If dblCredit_S <> dblDebit_S Then
            '    Dim dblAdjust As Double
            '    If intCurrentLine > 0 Then
            '        oJE.Lines.Add()
            '    End If
            '    oJE.Lines.SetCurrentLine(intCurrentLine)
            '    oJE.Lines.AccountCode = oApplication.Utilities.getAccount(strToAccount)
            '    If dblCredit_S > dblDebit_S Then
            '        dblAdjust = dblCredit_S - dblDebit_S
            '        oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
            '        oJE.Lines.Debit = dblAdjust
            '        dblDebit_S += dblAdjust
            '    ElseIf (dblDebit_S > dblCredit_S) Then
            '        dblAdjust = dblDebit_S - dblCredit_S
            '        oJE.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
            '        oJE.Lines.Credit = dblAdjust
            '        dblCredit_S += dblAdjust
            '    End If
            'End If

            'MessageBox.Show(dblCredit_S.ToString())
            'MessageBox.Show(dblDebit_S.ToString())
            'Throw New Exception("Error")

            'if Record Exist...
            If blnRecordExist Then
                Dim intCode As Integer = oJE.Add()
                If intCode = 0 Then
                    _retVal = True
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    For index As Integer = 0 To oDtJEList.Rows.Count - 1
                        strQuery = "Update JDT1 Set U_ActTra = 'Y' Where TransID = '" + oDtJEList.GetValue("TransID", index).ToString() + "' And Line_ID = '" + oDtJEList.GetValue("Line_ID", index).ToString() + "'"
                        oRecordSet.DoQuery(strQuery)
                    Next
                    Dim strTransID As String = oApplication.Company.GetNewObjectKey()
                    If strTransID > 0 Then
                        strQuery = "Update JDT1 Set ContraAct = U_Offset Where TransID = '" + strTransID + "' And ISNULL(U_OffSet,'') <> ''"
                        oRecordSet.DoQuery(strQuery)
                    End If
                    oApplication.SBO_Application.MessageBox("Journal Entry Posted Successfully...", 1, "Ok", "", "")
                    clearSource(oForm)
                Else
                    _retVal = False
                    Throw New Exception(oApplication.Company.GetLastErrorDescription())
                End If
            End If
        Catch ex As Exception
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        Return _retVal
    End Function

    Public Function post_JournalVoucher(ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        Dim strTable As String = String.Empty

        Try
            Dim dblCredit_S As Double
            Dim dblDebit_S As Double

            Dim oJE As SAPbobsCOM.JournalVouchers
            Dim strToAccount As String = oForm.Items.Item("17").Specific.value
            Dim dtPostingDt As DateTime = oApplication.Utilities.GetDateTimeValue(oForm.Items.Item("9").Specific.value)

            Dim dblSummaryAmt_LC As Double = 0
            Dim dblSummaryAmt_FC As Double = 0
            Dim dblSummaryAmt_SC As Double = 0

            Dim intCurrentLine As Integer = 0
            Dim blnRecordExist As Boolean = False
            oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)

            oJE.JournalEntries.ReferenceDate = dtPostingDt
            oJE.JournalEntries.TaxDate = dtPostingDt
            oJE.JournalEntries.DueDate = dtPostingDt
            oJE.JournalEntries.UserFields.Fields.Item("U_ActTra").Value = "Y"

            oJE.JournalEntries.Memo = "Account Posting"
            oJE.JournalEntries.Reference = "Account Posting"
            oJE.JournalEntries.Reference2 = "Account Posting"
            oJE.JournalEntries.Reference3 = "Account Posting"

            oGrid = oForm.Items.Item("3").Specific
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                blnRecordExist = True
                Dim strFcCurrency As String = String.Empty
                Dim strOffset As String = String.Empty

                'Set Offset Account.
                oComboBox = oForm.Items.Item("11").Specific
                Select Case oComboBox.Selected.Value
                    Case strType.ACT
                        strOffset = String.Empty
                    Case strType.AOF
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                    Case strType.ACR
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                    Case strType.ALL
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                End Select

                dblSummaryAmt_FC = oGrid.DataTable.GetValue("Amount", index)
                dblSummaryAmt_LC = oGrid.DataTable.GetValue("Amount_LC", index)

                If dblSummaryAmt_FC < 0 Then
                    If intCurrentLine > 0 Then
                        oJE.JournalEntries.Lines.Add()
                    End If
                    oJE.JournalEntries.Lines.SetCurrentLine(intCurrentLine)
                    If strOffset.Length > 0 Then
                        oJE.JournalEntries.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    oJE.JournalEntries.Lines.AccountCode = oGrid.DataTable.GetValue("AcctCode", index)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.JournalEntries.Lines.FCCurrency = strFcCurrency
                        oJE.JournalEntries.Lines.FCDebit = (dblSummaryAmt_FC * -1)
                        oJE.JournalEntries.Lines.Debit = (dblSummaryAmt_LC * -1)
                    Else
                        oJE.JournalEntries.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.JournalEntries.Lines.Debit = (dblSummaryAmt_FC * -1)
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblDebit_S += oGrid.DataTable.GetValue("Amount_LC", index) * -1
                Else
                    If intCurrentLine > 0 Then
                        oJE.JournalEntries.Lines.Add()
                    End If
                    oJE.JournalEntries.Lines.SetCurrentLine(intCurrentLine)
                    If strOffset.Length > 0 Then
                        oJE.JournalEntries.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    oJE.JournalEntries.Lines.AccountCode = oGrid.DataTable.GetValue("AcctCode", index)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.JournalEntries.Lines.FCCurrency = strFcCurrency
                        oJE.JournalEntries.Lines.FCCredit = dblSummaryAmt_FC
                        oJE.JournalEntries.Lines.Credit = dblSummaryAmt_LC
                    Else
                        oJE.JournalEntries.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.JournalEntries.Lines.Credit = dblSummaryAmt_FC
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblCredit_S += oGrid.DataTable.GetValue("Amount_LC", index)
                End If
            Next

            oGrid = oForm.Items.Item("34").Specific
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                Dim strFcCurrency As String = String.Empty
                Dim strOffset As String = String.Empty

                'Set Offset Account.
                oComboBox = oForm.Items.Item("13").Specific
                Select Case oComboBox.Selected.Value
                    Case strType.ACT
                        strOffset = String.Empty
                    Case strType.AOF
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                    Case strType.ACR
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                    Case strType.ALL
                        strOffset = oGrid.DataTable.GetValue("ContraAct", index)
                        strFcCurrency = oGrid.DataTable.GetValue("FCCurrency", index)
                End Select

                dblSummaryAmt_FC = oGrid.DataTable.GetValue("Amount", index)
                dblSummaryAmt_LC = oGrid.DataTable.GetValue("Amount_LC", index)

                If dblSummaryAmt_FC < 0 Then
                    If intCurrentLine > 0 Then
                        oJE.JournalEntries.Lines.Add()
                    End If
                    oJE.JournalEntries.Lines.SetCurrentLine(intCurrentLine)
                    oJE.JournalEntries.Lines.AccountCode = oApplication.Utilities.getAccount(strToAccount)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.JournalEntries.Lines.FCCurrency = strFcCurrency
                        oJE.JournalEntries.Lines.FCCredit = (dblSummaryAmt_FC * -1)
                        oJE.JournalEntries.Lines.Credit = (dblSummaryAmt_LC * -1)
                    Else
                        oJE.JournalEntries.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.JournalEntries.Lines.Credit = (dblSummaryAmt_FC * -1)
                    End If
                    If strOffset.Length > 0 Then
                        oJE.JournalEntries.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblCredit_S += oGrid.DataTable.GetValue("Amount_LC", index) * -1
                Else
                    If intCurrentLine > 0 Then
                        oJE.JournalEntries.Lines.Add()
                    End If
                    oJE.JournalEntries.Lines.SetCurrentLine(intCurrentLine)
                    oJE.JournalEntries.Lines.AccountCode = oApplication.Utilities.getAccount(strToAccount)
                    If strFcCurrency <> "" And strFcCurrency.Length > 0 Then
                        oJE.JournalEntries.Lines.FCCurrency = strFcCurrency
                        oJE.JournalEntries.Lines.FCDebit = dblSummaryAmt_FC
                        oJE.JournalEntries.Lines.Debit = dblSummaryAmt_LC
                    Else
                        oJE.JournalEntries.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
                        oJE.JournalEntries.Lines.Debit = dblSummaryAmt_FC
                    End If
                    If strOffset.Length > 0 Then
                        oJE.JournalEntries.Lines.UserFields.Fields.Item("U_Offset").Value = strOffset
                    End If
                    If intCurrentLine = 0 Then
                        intCurrentLine += 1
                    ElseIf (intCurrentLine > 0) Then
                        intCurrentLine += 1
                    End If
                    dblDebit_S += oGrid.DataTable.GetValue("Amount_LC", index)
                End If
            Next


            'If dblCredit_S <> dblDebit_S Then
            '    Dim dblAdjust As Double
            '    If intCurrentLine > 0 Then
            '        oJE.JournalEntries.Lines.Add()
            '    End If
            '    oJE.JournalEntries.Lines.SetCurrentLine(intCurrentLine)
            '    oJE.JournalEntries.Lines.AccountCode = oApplication.Utilities.getAccount(strToAccount)
            '    If dblCredit_S > dblDebit_S Then
            '        dblAdjust = dblCredit_S - dblDebit_S
            '        oJE.JournalEntries.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
            '        oJE.JournalEntries.Lines.Debit = dblAdjust
            '        dblDebit_S += dblAdjust
            '    ElseIf (dblDebit_S > dblCredit_S) Then
            '        dblAdjust = dblDebit_S - dblCredit_S
            '        oJE.JournalEntries.Lines.FCCurrency = oApplication.Company.GetCompanyService().GetAdminInfo().LocalCurrency
            '        oJE.JournalEntries.Lines.Credit = dblAdjust
            '        dblCredit_S += dblAdjust
            '    End If
            'End If

            'MessageBox.Show(dblCredit_S.ToString())
            'MessageBox.Show(dblDebit_S.ToString())
            'Throw New Exception("Error")

            'if Record Exist...
            If blnRecordExist Then
                Dim intCode As Integer = oJE.Add()
                If intCode = 0 Then
                    oApplication.SBO_Application.MessageBox("Journal Voucher Created Successfully...", 1, "Ok", "", "")
                    _retVal = True

                    'oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'For index As Integer = 0 To oDtJEList.Rows.Count - 1
                    '    strQuery = "Update JDT1 Set U_ActTra = 'Y' Where TransID = '" + oDtJEList.GetValue("TransID", index).ToString() + "' And Line_ID = '" + oDtJEList.GetValue("Line_ID", index).ToString() + "'"
                    '    oRecordSet.DoQuery(strQuery)
                    'Next
                    'Dim strTransID As String = oApplication.Company.GetNewObjectKey()
                    'If strTransID > 0 Then
                    '    strQuery = "Update JDT1 Set ContraAct = U_Offset Where TransID = '" + strTransID + "' And ISNULL(U_OffSet,'') <> ''"
                    '    oRecordSet.DoQuery(strQuery)
                    'End If
                    'oApplication.SBO_Application.MessageBox("Journal Entry Posted Successfully...", 1, "Ok", "", "")

                    clearSource(oForm)

                Else
                    _retVal = False
                    Throw New Exception(oApplication.Company.GetLastErrorDescription())
                End If
            End If
        Catch ex As Exception
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        Return _retVal
    End Function

    Public Sub formatPostingFromGrid(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oGrid = oForm.Items.Item("1").Specific
            oGrid.Columns.Item("FormatCode").TitleObject.Caption = "Format Code"
            oGrid.Columns.Item("AcctCode").TitleObject.Caption = "Account Code"
            oGrid.Columns.Item("AcctCode").Visible = False
            oGrid.Columns.Item("AcctName").TitleObject.Caption = "Account Name"
            oGrid.Columns.Item("AcctName").Editable = False
            oEditTextCol = oGrid.Columns.Item("FormatCode")
            oEditTextCol.LinkedObjectType = "1"
            oEditTextCol = oGrid.Columns.Item("FormatCode")
            oEditTextCol.ChooseFromListUID = "CFL_1"
            oEditTextCol.ChooseFromListAlias = "FormatCode"
            oEditTextCol.LinkedObjectType = "1"
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
            oApplication.Utilities.assignLineNo(oGrid, oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Public Sub formatTransactionGrid(ByVal oForm As SAPbouiCOM.Form, ByVal strGridID As String, ByVal strGroupID As String)
        Try
            oForm.Freeze(True)
            oGrid = oForm.Items.Item(strGridID).Specific

            oGrid.Columns.Item("FormatCode").TitleObject.Caption = "Format Code"
            oGrid.Columns.Item("FormatCode").Editable = False
            oGrid.Columns.Item("AcctCode").TitleObject.Caption = "Account Code"
            oGrid.Columns.Item("AcctCode").Visible = False
            oGrid.Columns.Item("AcctName").TitleObject.Caption = "Account Name"
            oGrid.Columns.Item("AcctName").Editable = False
            oGrid.Columns.Item("Amount").TitleObject.Caption = "Transaction Amount"
            oGrid.Columns.Item("Amount").Editable = False
            oGrid.Columns.Item("Amount").RightJustified = True
            oGrid.Columns.Item("FCCurrency").Editable = False
            oComboBox = oForm.Items.Item(strGroupID).Specific
            Select Case oComboBox.Selected.Value
                Case strType.AOF
                    oGrid.Columns.Item("ContraAct").TitleObject.Caption = "Offset"
                    oGrid.Columns.Item("ContraAct").Editable = False
                Case strType.ACR
                    oGrid.Columns.Item("FCCurrency").TitleObject.Caption = "Currency"
                    oGrid.Columns.Item("FCCurrency").Visible = True
                    oGrid.Columns.Item("FCCurrency").Editable = False
                Case strType.ALL
                    oGrid.Columns.Item("FCCurrency").TitleObject.Caption = "Currency"
                    oGrid.Columns.Item("FCCurrency").Visible = True
                    oGrid.Columns.Item("ContraAct").Editable = False
                    oGrid.Columns.Item("FCCurrency").Editable = False
            End Select

            oGrid.Columns.Item("Amount_LC").TitleObject.Caption = "Transaction Amount(LC)"
            oGrid.Columns.Item("Amount_LC").Editable = False
            oGrid.Columns.Item("Amount_LC").RightJustified = True

            oEditTextCol = oGrid.Columns.Item("FormatCode")
            oEditTextCol.LinkedObjectType = "1"
            oEditTextCol = oGrid.Columns.Item("AcctCode")
            oEditTextCol.LinkedObjectType = "1"

            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
            oApplication.Utilities.assignLineNo(oGrid, oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oGrid = oForm.Items.Item("1").Specific
            oForm.Items.Item("1").Width = 380
            oForm.Items.Item("1").Height = 110
            oForm.Items.Item("20").Left = oForm.Items.Item("1").Left + oForm.Items.Item("1").Width - 60
            oForm.Items.Item("19").Left = oForm.Items.Item("20").Left - 65
            oGrid.AutoResizeColumns()
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oGrid = aForm.Items.Item("1").Specific
            If oGrid.DataTable.Rows.Count - 1 < 0 Then
                oGrid.DataTable.Rows.Add()
            End If
            oApplication.Utilities.assignLineNo(oGrid, aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Delete(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            For intRow As Integer = oGrid.DataTable.Rows.Count - 1 To 0 Step -1
                If oGrid.Rows.IsSelected(intRow) Then
                    oGrid.DataTable.Rows.Remove(intRow)
                End If
            Next
            oApplication.Utilities.assignLineNo(oGrid, aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub clearSource(ByVal oForm As SAPbouiCOM.Form)
        Try
            CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("11").Specific, SAPbouiCOM.ComboBox).Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            CType(oForm.Items.Item("13").Specific, SAPbouiCOM.ComboBox).Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            CType(oForm.Items.Item("9").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("17").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("23").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("27").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("25").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("28").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("32").Specific, SAPbouiCOM.EditText).Value = ""

            oGrid = oForm.Items.Item("1").Specific
            strQuery = "Select FormatCode,AcctCode,AcctName From OACT Where 1 = 2"
            oDtAccountList = oForm.DataSources.DataTables.Item("dtAccountList")
            oDtAccountList.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtAccountList
            formatPostingFromGrid(oForm)

            'Summary by Selection
            oGrid = oForm.Items.Item("3").Specific
            strQuery = "Select T1.FormatCode,T1.AcctCode,T1.AcctName,Sum(Debit) - Sum(Credit) As 'Amount','' As FCCurrency,Sum(Debit) - Sum(Credit) As 'Amount_LC' From  JDT1 T0 Join OACT T1 On T0.Account = T1.AcctCode Where 1 = 2"
            strQuery += " Group By T1.AcctCode,T1.AcctName,T1.FormatCode "
            oDtTransList_S = oForm.DataSources.DataTables.Item("dtTransList_S")
            oDtTransList_S.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtTransList_S
            formatTransactionGrid(oForm, "3", "11")

            oGrid = oForm.Items.Item("_3").Specific 'Visible True Grid
            strQuery = "Select T1.FormatCode,T1.AcctCode,T1.AcctName,Sum(Debit) - Sum(Credit) As 'Amount','' As FCCurrency,Sum(Debit) - Sum(Credit) As 'Amount_LC',T0.LicTradNum 'Federal Tax ID',Sum(T0.BaseSum) 'BaseAmount' From  JDT1 T0 Join OACT T1 On T0.Account = T1.AcctCode Where 1 = 2"
            strQuery += " Group By T1.AcctCode,T1.AcctName,T1.FormatCode,T0.LicTradNum "
            oDtTransList_S1 = oForm.DataSources.DataTables.Item("dtTransList_S1")
            oDtTransList_S1.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtTransList_S1
            formatTransactionGrid(oForm, "3", "11")

            'Summary by Posting
            oGrid = oForm.Items.Item("34").Specific
            strQuery = "Select T1.FormatCode,T1.AcctCode,T1.AcctName,Sum(Debit) - Sum(Credit) As 'Amount','' As FCCurrency,Sum(Debit) - Sum(Credit) As 'Amount_LC' From  JDT1 T0 Join OACT T1 On T0.Account = T1.AcctCode Where 1 = 2"
            strQuery += " Group By T1.AcctCode,T1.AcctName,T1.FormatCode "
            oDtTransList_P = oForm.DataSources.DataTables.Item("dtTransList_P")
            oDtTransList_P.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtTransList_P
            formatTransactionGrid(oForm, "34", "13")

            oDtJEList.Rows.Clear()
            'oDtCurrList.Rows.Clear()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
