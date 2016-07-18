Public Class clsSystemForms

    Private oForm As SAPbouiCOM.Form
    Private oExistingItem As SAPbouiCOM.Item
    Private oItem As SAPbouiCOM.Item
    Private oChkBox As SAPbouiCOM.CheckBox
    Private oButton As SAPbouiCOM.Button

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub AddItems(ByVal FormType As Integer, ByVal FormUID As String)
        Select Case FormType

            Case frm_WAREHOUSES
                AddItemsToWareHouses(FormUID)

            Case frm_ITEM_MASTER
                AddItemsToItemMaster(FormUID)

            Case frm_INVOICES
                AddItemsToInvoices(FormUID)

        End Select
    End Sub

    Private Sub AddItemsToWareHouses(ByVal sFormUID As String)
        oForm = oApplication.SBO_Application.Forms.Item(sFormUID)

        oExistingItem = oForm.Items.Item("89")
        oItem = oForm.Items.Add("1001", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        oItem.Top = oExistingItem.Top + oExistingItem.Height + 2
        oItem.Left = oExistingItem.Left
        oItem.Height = oExistingItem.Height
        oItem.Width = oExistingItem.Width + 10
        oItem.FromPane = oExistingItem.FromPane
        oItem.ToPane = oExistingItem.ToPane
        oChkBox = oItem.Specific
        oChkBox.Caption = "Rental"
        oChkBox.DataBind.SetBound(True, "OWHS", "U_Rental")

        'oExistingItem = oForm.Items.Item("1001")
        'oItem = oForm.Items.Add("1002", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        'oItem.Top = oExistingItem.Top + oExistingItem.Height + 2
        'oItem.Left = oExistingItem.Left
        'oItem.Height = oExistingItem.Height
        'oItem.Width = oExistingItem.Width + 10
        'oItem.FromPane = oExistingItem.FromPane
        'oItem.ToPane = oExistingItem.ToPane
        'oChkBox = oItem.Specific
        'oChkBox.Caption = "Rental"
        'oChkBox.DataBind.SetBound(True, "OWHS", "U_Rental")

    End Sub

    Private Sub AddItemsToItemMaster(ByVal sFormUID As String)
        oForm = oApplication.SBO_Application.Forms.Item(sFormUID)

        oExistingItem = oForm.Items.Item("42")
        oItem = oForm.Items.Add("1001", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        oItem.Top = oExistingItem.Top + oExistingItem.Height + 1
        oItem.Left = oExistingItem.Left
        oItem.Height = oExistingItem.Height
        oItem.Width = oExistingItem.Width + 10
        oChkBox = oItem.Specific
        oChkBox.Caption = "Rental Item"
        oChkBox.DataBind.SetBound(True, "OITM", "U_Rental")

    End Sub

    Private Sub AddItemsToInvoices(ByVal sFromUID As String)
        oForm = oApplication.SBO_Application.Forms.Item(sFromUID)

        'oExistingItem = oForm.Items.Item("37")
        'oExistingItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 10

        'oExistingItem = oForm.Items.Item("36")
        'oExistingItem.Width = oExistingItem.Width / 1.3
        'oExistingItem.Left = oForm.Items.Item("37").Left + oForm.Items.Item("37").Width + 5

        'oExistingItem = oForm.Items.Item("35")
        'oExistingItem.Width = oExistingItem.Width / 2.4
        'oExistingItem.Left = oForm.Items.Item("36").Left + oForm.Items.Item("36").Width + 5

        oExistingItem = oForm.Items.Item("2")
        oItem = oForm.Items.Add("1001", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        oItem.Top = oExistingItem.Top
        oItem.Height = oExistingItem.Height
        oItem.Left = oExistingItem.Left + oExistingItem.Width + 5
        oItem.Width = oExistingItem.Width + 18
        oItem.Enabled = False
        oButton = oItem.Specific
        oButton.Caption = "Rental Order"

        'oExistingItem = oForm.Items.Item("1001")
        'oItem = oForm.Items.Add("1002", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        'oItem.Top = oExistingItem.Top
        'oItem.Height = oExistingItem.Height
        'oItem.Left = oExistingItem.Left + oExistingItem.Width + 5
        'oItem.Width = oExistingItem.Width
        'oButton = oItem.Specific
        'oItem.Enabled = False
        'oButton.Caption = "Rental Return"

    End Sub

End Class
