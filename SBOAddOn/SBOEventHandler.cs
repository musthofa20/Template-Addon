using System;
using System.Xml;
using System.Windows.Forms;
using System.Collections.Generic;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using CrystalDecisions.Windows.Forms;
using System.IO;

public class SBOEventHandler
{

    #region " Private Attribute "

    private SAPbouiCOM.Application oSBOApplication;
    private SAPbobsCOM.Company oSBOCompany;

    #endregion

    #region " Public Constructor "

    public SBOEventHandler() { }

    public SBOEventHandler(SAPbouiCOM.Application oSBOApplication)
    {
        this.oSBOApplication = oSBOApplication;
    }

    public SBOEventHandler(SAPbouiCOM.Application oSBOApplication, SAPbobsCOM.Company oSBOCompany)
    {
        this.oSBOApplication = oSBOApplication;
        this.oSBOCompany = oSBOCompany;
    }

    #endregion

    public void HandleAppEvent(SAPbouiCOM.BoAppEventTypes EventType)
    {
        switch (EventType)
        {
            case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                System.Windows.Forms.Application.Exit();
                break;
            case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                System.Windows.Forms.Application.Exit();
                break;
            case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                System.Windows.Forms.Application.Exit();
                break;
            case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                System.Windows.Forms.Application.Exit();
                break;
        }
    }
    public void HandleMenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

        try
        {
            if (pVal.MenuUID == "")
            {
                MenuEvent_Handler(ref pVal, out BubbleEvent);
            }
            else if (pVal.MenuUID == "FORM1")
            {
                MenuEvent_Handler_FORM1(ref pVal, out BubbleEvent);
            }
            else if (pVal.MenuUID == "Add")
            {
                MenuEvent_Handler_AddRow(ref pVal, out BubbleEvent);
            }
            else if (pVal.MenuUID == "Del")
            {
                MenuEvent_Handler_DeleteRow(ref pVal, out BubbleEvent);
            }
        }
        catch (Exception ex)
        {
            oSBOApplication.MessageBox(ex.Message);
        }
    }
    public void HandleItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

        try
        {
            if (pVal.FormTypeEx == "")
            {
                ItemEvent_Handler(FormUID, ref pVal, out BubbleEvent);
            }
            else if (pVal.FormUID == "frm_FORM1")
            {
                ItemEvent_Handler_FORM1(FormUID, ref pVal, out BubbleEvent);
            }
            else if (pVal.FormTypeEx == "940")
            {
                ItemEvent_Handler_TEST1(FormUID, ref pVal, out BubbleEvent);
            }
            else if (pVal.FormType == 139)
            {
                ItemEvent_Handler_SalesOrder(FormUID, ref pVal, out BubbleEvent);
            }
        }
        catch (Exception ex)
        {
            oSBOApplication.MessageBox(ex.Message);
        }
    }
    public void HandleFormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;

        try
        {
            if (BusinessObjectInfo.FormTypeEx == "")
            {
                FormDataEvent_Handler(ref BusinessObjectInfo, out BubbleEvent);
            }
        }
        catch (Exception ex)
        {
            oSBOApplication.MessageBox(ex.Message);
        }
    }
    public void HandleRightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;

        try
        {
            if (eventInfo.FormUID == "")
            {
                RightClickEvent_Handler(ref eventInfo, out BubbleEvent);
            }
            else if (eventInfo.FormUID == "frm_FORM1")
            {
                RightClickEvent_Handler_FORM1(ref eventInfo, out BubbleEvent);
            }
        }
        catch (Exception ex)
        {
            oSBOApplication.MessageBox(ex.Message);
        }
    }
    public void HandleLayoutKeyEvent(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;
        try
        {
            if (eventInfo.FormUID == "frm_FORM1")
                LayoutKeyEvent_Handler_frm_FORM1(ref eventInfo, out BubbleEvent);
        }
        catch (Exception ex)
        {
            oSBOApplication.MessageBox(ex.Message);
        }
    }

    #region " MenuEvent Handler "

    public void MenuEvent_Handler(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;
    }

    public void MenuEvent_Handler_FORM1(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

        if (pVal.BeforeAction == false)
        {
            XMLLoader oXMLLoader = new XMLLoader();
            XmlDocument oXMLDoc = new XmlDocument();
            SAPbobsCOM.Recordset oRec;
            try
            {
                oXMLLoader.LoadFromXML(oSBOApplication, oXMLDoc, "FORM1.srf");
                SAPbouiCOM.Form oForm = oSBOApplication.Forms.Item("frm_FORM1");
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("Mtx").Specific;
                GeneralVariables.oDBSOL_SOL_DS_H = oSBOApplication.Forms.Item("frm_FORM1").DataSources.DBDataSources.Item("@SOL_FORM_H");
                GeneralVariables.oDBSOL_SOL_DS_D = oSBOApplication.Forms.Item("frm_FORM1").DataSources.DBDataSources.Item("@SOL_FORM_D");

                GeneralVariables.oDBSOL_SOL_DS_H.SetValue("U_SOL_REF_NUM", 0, "123321");

                oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery(@"SELECT ""CODE"" FROM ""RTYP"" WHERE ""NAME"" = 'FORM1'"); //TypeCode
                if (oRec.RecordCount > 0)
                {
                    oForm.ReportType = oRec.Fields.Item("CODE").Value;
                }
                else
                {
                    string Layout = "FORM1.rpt"; //Should to define
                    string TypeCode = "FORM1"; //Should to define
                    oForm.ReportType = CRSetup(TypeCode, Layout, oForm);
                }

                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                return;
            }
        }
    }

    public void MenuEvent_Handler_AddRow(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;
        if (pVal.BeforeAction == false)
        {
            SAPbouiCOM.Form oForm = oSBOApplication.Forms.ActiveForm;
            if (oForm.UniqueID == "frm_FORM1")
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("Mtx").Specific;
                oMatrix.AddRow();
                oMatrix.FlushToDataSource();
            }
        }
    }

    public void MenuEvent_Handler_DeleteRow(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;
        if (pVal.BeforeAction == false)
        {
            SAPbouiCOM.Form oForm = oSBOApplication.Forms.ActiveForm;
            if (oForm.UniqueID == "frm_FORM1")
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("Mtx").Specific;
                oMatrix.DeleteRow(GeneralVariables.SelectedRow);
                oMatrix.FlushToDataSource();
            }
        }
    }

    #endregion

    #region " ItemEvent Handler "

    public void ItemEvent_Handler(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

    }

    public void ItemEvent_Handler_FORM1(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

        if ((pVal.ItemUID == "1") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
        {
            if ((pVal.ActionSuccess == false) && (pVal.BeforeAction == true))
            {
            }
            if ((pVal.ActionSuccess == true) && (pVal.BeforeAction == false))
            {
                // GENERATE OUTGOING PAYMENT 
                SAPbobsCOM.Payments vPay = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments;
                vPay.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;

                int iErrCode = 0;
                string sErrMsg = "";
                string errormessage = "";
                vPay.CardCode = "V23000";
                DateTime hariini = DateTime.Now;
                vPay.DocDate = hariini.Date;
                vPay.DueDate = hariini.Date;
                vPay.TaxDate = hariini.Date;
                vPay.TransferAccount = "161010";
                vPay.TransferDate = hariini.Date;
                vPay.Remarks = "TESTDI";

                vPay.TransferSum = 100;

                vPay.Invoices.SetCurrentLine(0);
                vPay.Invoices.DocEntry = 575;
                vPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice;

                vPay.Invoices.AppliedFC = 10;
                vPay.Invoices.Add();

                //vPay.PrimaryFormItems.AmountFC = 10;
                //vPay.PrimaryFormItems.CashFlowLineItemID = 18;
                //vPay.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;

                if (vPay.Add() != 0)
                {
                    oSBOCompany.GetLastError(out iErrCode, out sErrMsg);                    // Jika terjadi gagal generate outgoing payment maka simpan error message
                    errormessage = Convert.ToString(iErrCode) + " - " + sErrMsg;
                }
                else
                {
                    // update submission status dan nomor outgoing payment dan external ID
                    errormessage = "Generated";
                }
            }
        }
        if (pVal.ItemUID == "CardCode")
        {
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            {
                SAPbouiCOM.IChooseFromListEvent oCFLEvent;
                SAPbouiCOM.DataTable oDataTable;
                SAPbouiCOM.Form oForm;
                SAPbobsCOM.Recordset oRec;

                string sQuery = "";

                oForm = oSBOApplication.Forms.Item(FormUID);
                GeneralVariables.oDBSOL_SOL_DS_H = oForm.DataSources.DBDataSources.Item("@SOL_FORM_H");
                GeneralVariables.oDBSOL_SOL_DS_D = oForm.DataSources.DBDataSources.Item("@SOL_FORM_D");

                oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oCFLEvent = pVal as SAPbouiCOM.ChooseFromListEvent;

                if (oCFLEvent.BeforeAction == true)
                {
                    SAPbouiCOM.ChooseFromListCollection oCFLs;
                    SAPbouiCOM.Conditions oCons;
                    SAPbouiCOM.Condition oCon;
                    SAPbouiCOM.ChooseFromList oCFL;

                    oCFLs = oForm.ChooseFromLists;
                    oCFL = oCFLs.Item("CFL_BP");
                    oCons = oCFL.GetConditions();

                    if (oCons.Count == 0)
                    {
                        oCon = oCons.Add();
                        oCon = oCons.Add();
                        oCon = oCons.Add();
                        oCon = oCons.Add();
                    }

                    ////===========Filter Munculin Item FG===============
                    //oCon = oCons.Item(0);
                    //oCon.BracketOpenNum = 1; ;
                    //oCon.Alias = "ItmsGrpCod";
                    //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //oCon.CondVal = "120";       //<===KITOSHINDO       //"101"; <===LOCAL       //============DI UBAH SESUAI DENGAN GROUP CODE KITOSHINDO============
                    //oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                    //oCFL.SetConditions(oCons);

                    ////===========Filter Munculin Item SFG===============
                    //oCon = oCons.Item(1);
                    //oCon.Alias = "ItmsGrpCod";
                    //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //oCon.CondVal = "122";       //<===KITOSHINDO       //"104"; <===LOCAL       //============DI UBAH SESUAI DENGAN GROUP CODE KITOSHINDO============
                    //oCon.BracketCloseNum = 1;
                    //oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    //oCFL.SetConditions(oCons);

                    ////============Untuk Filter buang Asset Item============
                    //oCon = oCons.Item(2);
                    //oCon.Alias = "ItemType";
                    //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                    //oCon.CondVal = "F";
                    //oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    //oCFL.SetConditions(oCons);

                    ////============Untuk Filter buang Non-Inventoty Item============
                    //oCon = oCons.Item(3);
                    //oCon.Alias = "InvntItem";
                    //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                    //oCon.CondVal = "N";
                    //oCFL.SetConditions(oCons);
                }
                else if (oCFLEvent.BeforeAction == false)
                {
                    oDataTable = oCFLEvent.SelectedObjects;
                    DateTime current = DateTime.Today;
                    SAPbouiCOM.Matrix oMtx = oForm.Items.Item("Mtx").Specific;

                    //if (oDataTable.IsEmpty == true)
                    //{
                    try
                    {
                        GeneralVariables.oDBSOL_SOL_DS_H.SetValue("U_SOL_CARDCODE", 0, oDataTable.GetValue("CardCode", 0));
                        string cardname = oDataTable.GetValue("CardName", 0);
                    }
                    catch (Exception ex)
                    {
                        oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    // }
                }
            }
        }
        if (pVal.ItemUID == "Mtx" && pVal.ColUID == "ITCD")
        {
            SAPbouiCOM.IChooseFromListEvent oCFLEvent = pVal as SAPbouiCOM.ChooseFromListEvent;
            SAPbouiCOM.DataTable oDataTable;
            SAPbouiCOM.Form oForm = oSBOApplication.Forms.Item(FormUID);
            SAPbouiCOM.Matrix oMtx = oForm.Items.Item("Mtx").Specific;
            SAPbouiCOM.EditText oFlag;

            GeneralVariables.oDBSOL_SOL_DS_D = oForm.DataSources.DBDataSources.Item("@SOL_FORM_D");

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            {
                if (oCFLEvent.BeforeAction == true)
                {
                    SAPbouiCOM.ChooseFromListCollection oCFLs;
                    SAPbouiCOM.ChooseFromList oCFL;

                    oCFLs = oForm.ChooseFromLists;
                    oCFL = oCFLs.Item("CFL_ITEM");
                }
                else if (oCFLEvent.BeforeAction == false)
                {
                    oDataTable = oCFLEvent.SelectedObjects;
                    
                    try
                    {
                        oMtx.FlushToDataSource();
                        GeneralVariables.oDBSOL_SOL_DS_D.SetValue("U_SOL_ITEMCODE", pVal.Row - 1, oDataTable.GetValue("ItemCode", 0));
                        oMtx.LoadFromDataSource();
                    }
                    catch (Exception ex)
                    {
                        oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
            }
        }
    }

    public void ItemEvent_Handler_TEST1(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

        if ((pVal.ItemUID == "256000001") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) && (pVal.ActionSuccess == true) && (pVal.BeforeAction == false))
        {
            SAPbouiCOM.Form oForm = oSBOApplication.Forms.ActiveForm;
            string dataxml = oForm.GetAsXML();
            oSBOApplication.MessageBox(dataxml);
            Console.Write(dataxml);
        }
    }

    public void ItemEvent_Handler_SalesOrder(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        BubbleEvent = true;

        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        {
            if ((pVal.ActionSuccess == true) && (pVal.BeforeAction == false))
            {
                SAPbouiCOM.Form oForm = oSBOApplication.Forms.Item(FormUID); //From Sales Order
                //SAPbouiCOM.Form UDFSalesOrder = oSBOApplication.Forms.Item(oForm.UDFFormUID); //UDF Form Sales Order
                SAPbouiCOM.Button oButtonCancel = oForm.Items.Item("2").Specific;
                SAPbouiCOM.EditText oRemarks = oForm.Items.Item("16").Specific;
                SAPbouiCOM.ComboBox oCurrency = oForm.Items.Item("70").Specific;
                
                int top = oButtonCancel.Item.Top;
                int left = oButtonCancel.Item.Left;
                int height = oButtonCancel.Item.Height;

                SAPbouiCOM.Item oItemButton;
                SAPbouiCOM.Button NewButton;

                oItemButton = oForm.Items.Add("CSI", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItemButton.Top = top;
                oItemButton.Left = left + 70;
                oItemButton.Width = 120;
                oItemButton.Height = height;
                oItemButton.DisplayDesc = true;
                oItemButton.Enabled = true;
                NewButton = oForm.Items.Item("CSI").Specific;
                NewButton.Caption = "Check Supporting Item";
            }
        }
        if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) && (pVal.ItemUID == "ODO"))
        {
            if ((pVal.BeforeAction == false) && (pVal.ActionSuccess == true))
            {
                oSBOApplication.ActivateMenuItem("2051");
            }
        }
        if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK) && (pVal.ItemUID == "17"))
        {
            if ((pVal.BeforeAction == false) && (pVal.ActionSuccess == true))
            {
                SAPbouiCOM.Form oForm = oSBOApplication.Forms.ActiveForm;
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("38").Specific;
                SAPbouiCOM.EditText oItemNo = oMatrix.Columns.Item("1").Cells.Item(1).Specific;
                //oSBOApplication.MessageBox(oItemNo.Value);

                var Item = new List<string>();
                int i;
                for (i = 1; i < oMatrix.RowCount; i++)
                {
                    //SAPbouiCOM.EditText oSelectedItem = oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                    //Item.Add(oSelectedItem.Value);
                    SAPbouiCOM.EditText oSelectedItem = oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                    Item.Add(oSelectedItem.Value);
                }

                for (i = 0; i < Item.Count; i++)
                {
                    oSBOApplication.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Items, "4", Item[i]);
                }


            }
        }

        if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) && (pVal.ItemUID == "CSI"))
        {
            if ((pVal.BeforeAction == false) && (pVal.ActionSuccess == true))
            {
                SAPbouiCOM.Form oForm = oSBOApplication.Forms.ActiveForm;
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("38").Specific;
                SAPbobsCOM.Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbouiCOM.EditText oRemarks = oForm.Items.Item("16").Specific;
                SAPbouiCOM.Form UDFSalesOrder = oSBOApplication.Forms.Item(oForm.UDFFormUID); //UDF Form Sales Order
                SAPbouiCOM.EditText oNomorFakturPajak = UDFSalesOrder.Items.Item("U_SOL_FP_RUNNING_NO").Specific;
                oNomorFakturPajak.Value = "00001";


                
                //    oRec.DoQuery(@"SELECT IFNULL(""U_SOL_SUPPORT_ITEM"",'') AS ""SUPPORT_ITEM"" FROM OITM WHERE ""ItemCode"" = '" + oItemCode.Value + "'");
                //    if (oRec.RecordCount > 0)
                //    {
                //        string SupportingItemCode = oRec.Fields.Item("SUPPORT_ITEM").Value;
                //        if (SupportingItemCode != "")
                //        {
                //            oMatrix.AddRow(1, i);
                //            SAPbouiCOM.EditText oSupportingItem = oMatrix.Columns.Item("1").Cells.Item(i+1).Specific;
                //            oSupportingItem.Value = SupportingItemCode;
                //            oForm.Items.Item("16").Click();
                //        }
                //    }

                //}
            }
        }
    }

    #endregion

    #region " FormDataEvent Handler "

    public void FormDataEvent_Handler(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;
    }

    #endregion

    #region " RighClickEvent Handler "

    public void RightClickEvent_Handler(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;
    }

    public void RightClickEvent_Handler_FORM1(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;

        if (eventInfo.BeforeAction == true && eventInfo.ActionSuccess == false)
        {
            if (eventInfo.ItemUID == "Mtx")
            {
                GeneralVariables.SelectedRow = eventInfo.Row;
            }
            try
            {
                SAPbouiCOM.Form oForm = oSBOApplication.Forms.ActiveForm;
                if (oForm.TypeEx == "frm_FORM1" && eventInfo.ItemUID == "Mtx")
                {
                    SAPbouiCOM.Menus oMenus;
                    SAPbouiCOM.MenuItem oMenuItem;
                    SAPbouiCOM.MenuCreationParams oCreationPackage;

                    oCreationPackage = oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    oMenuItem = oSBOApplication.Menus.Item("1280");
                    oMenus = oMenuItem.SubMenus;

                    if (!oMenus.Exists("Add"))
                    {
                        oCreationPackage.UniqueID = "Add";
                        oCreationPackage.String = "Add Line";
                        oCreationPackage.Enabled = true;
                        oMenus.AddEx(oCreationPackage);
                    }

                    if (!oMenus.Exists("Del"))
                    {
                        oCreationPackage.UniqueID = "Del";
                        oCreationPackage.String = "Delete Line";
                        oCreationPackage.Enabled = true;
                        oMenus.AddEx(oCreationPackage);
                    }
                }
                else
                {
                    try
                    {
                        oSBOApplication.Menus.RemoveEx("Add");
                    }
                    catch { Exception ex; }

                    try
                    {
                        oSBOApplication.Menus.RemoveEx("Del");
                    }
                    catch { Exception ex; }
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
        }
    }

    public void RightClickEvent_Handler_SalesOrder(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;

        if (eventInfo.BeforeAction == true && eventInfo.ActionSuccess == false)
        {
            try
            {
                SAPbouiCOM.Menus oMenus;
                SAPbouiCOM.MenuItem oMenuItem;
                SAPbouiCOM.MenuCreationParams oCreationPackage;

                oCreationPackage = oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                oMenuItem = oSBOApplication.Menus.Item("1280");
                oMenus = oMenuItem.SubMenus;

                if (!oMenus.Exists("ODO"))
                {
                    oCreationPackage.UniqueID = "ODO";
                    oCreationPackage.String = "Open Delivery Order";
                    oCreationPackage.Enabled = true;
                    oMenus.AddEx(oCreationPackage);
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
        }
    }

    #endregion

    #region " LayoutKeyEvent Handler "
    public void LayoutKeyEvent_Handler_frm_FORM1(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
    {
        BubbleEvent = true;
        if (eventInfo.BeforeAction == true)
        {
            try
            {
                SAPbouiCOM.Form oForm = oSBOApplication.Forms.Item(eventInfo.FormUID);
                SAPbobsCOM.Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string Code = Convert.ToString(oForm.Items.Item("DocNum").Specific.value); //Document Number Column
                string DocEntry;
                oRec.DoQuery(@"SELECT ""DocEntry"" FROM ""@SOL_FORM_H"" WHERE ""DocNum"" = '" + Code + "'"); //Get DocEntry
                if (oRec.RecordCount > 0)
                {
                    BubbleEvent = true;
                    eventInfo.LayoutKey = Convert.ToString(oRec.Fields.Item("DocEntry").Value);
                }
                else
                    BubbleEvent = false;
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
        }
    }
    public string CRSetup(string TypeCode, string RPTFile, SAPbouiCOM.Form oForm)
    {
        try
        {
            SAPbobsCOM.ReportLayoutsService oLayoutService = (SAPbobsCOM.ReportLayoutsService)oSBOCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
            SAPbobsCOM.ReportLayout oReport = (SAPbobsCOM.ReportLayout)oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);
            SAPbobsCOM.ReportTypesService rptTypeService = (SAPbobsCOM.ReportTypesService)oSBOCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);
            SAPbobsCOM.ReportType newType = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType);
            newType.TypeName = TypeCode;
            newType.AddonName = TypeCode;
            newType.AddonFormType = TypeCode;
            newType.MenuID = TypeCode;

            SAPbobsCOM.ReportTypeParams newTypeParam = rptTypeService.AddReportType(newType);

            // Use TypeCode "RCRI" to specify a Crystal Report.
            // Use other TypeCode to specify a layout for a document type.
            // List of TypeCode types are in table RTYP.

            oReport.Name = RPTFile.Replace(".rpt", "");
            oReport.TypeCode = newTypeParam.TypeCode;
            oReport.Author = oSBOCompany.UserName;
            oReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal;
            string newReportCode;

            try
            {
                // Add new object
                SAPbobsCOM.ReportLayoutParams oNewReportParams = oLayoutService.AddReportLayout(oReport);
                // Get code of the added ReportLayout object
                newReportCode = oNewReportParams.LayoutCode;
            }

            catch (System.Exception err)
            {
                string errMessage = err.Message;
                throw new Exception(errMessage);
            }

            // Wpload .rpt file using SetBlob interface
            string rptFilePath = RPTFile;
            SAPbobsCOM.CompanyService oCompanyService = oSBOCompany.GetCompanyService();
            // Specify the table and field to update
            SAPbobsCOM.BlobParams oBlobParams = (SAPbobsCOM.BlobParams)oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams);
            oBlobParams.Table = "RDOC";
            oBlobParams.Field = "Template";

            // Specify the record whose blob field is to be set
            SAPbobsCOM.BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add();
            oKeySegment.Name = "DocCode";
            oKeySegment.Value = newReportCode;
            SAPbobsCOM.Blob oBlob = (SAPbobsCOM.Blob)oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob);

            // Put the rpt file into buffer
            FileStream oFile = new FileStream(rptFilePath, System.IO.FileMode.Open);
            int fileSize = (int)oFile.Length;
            byte[] buf = new byte[fileSize];
            oFile.Read(buf, 0, fileSize);
            oFile.Close();

            // Convert memory buffer to Base64 string
            oBlob.Content = Convert.ToBase64String(buf, 0, fileSize);
            try
            {
                //Upload Blob to database
                oCompanyService.SetBlob(oBlobParams, oBlob);
            }

            catch (System.Exception ex)
            {
                string errmsg = ex.Message;
            }
            return newType.TypeCode;
        }
        catch (Exception ex)
        {
            return "";
            oSBOApplication.MessageBox(ex.Message);
        }
    }

    #endregion

}
