using System;
using System.Windows.Forms;

public class Manager
{

    public const string ADDON_NAME = "Extra SBO Add On";
    private SAPbouiCOM.Application oSBOApplication = null;
    private SAPbobsCOM.Company oSBOCompany = null;

    public Manager()
    {
        StartUp();
    }
    private void StartUp()
    {
        try
        {
            SetupApplication();

            NumericSeparators();
            CatchingEvents();

            Utils.CreateMenu(oSBOCompany, oSBOApplication, "43520", "EXTRA_ADDON", SAPbouiCOM.BoMenuType.mt_POPUP, "Extra Add On", @"logo.png", 0, true);
            Utils.CreateMenu(oSBOCompany, oSBOApplication, "EXTRA_ADDON", "FORM1", SAPbouiCOM.BoMenuType.mt_STRING, "Form 1", "", 1, true);

            //Utils.CreateMenu(oSBOCompany, oSBOApplication, "43520", "DS_TEST", SAPbouiCOM.BoMenuType.mt_POPUP, "DS Add-On", @"logo.png", 0, true);
            //Utils.CreateMenu(oSBOCompany, oSBOApplication, "DS_TEST", "DS_FORM", SAPbouiCOM.BoMenuType.mt_STRING, "DS Form", "", 1, true);

            Utils.CreateUDT(oSBOCompany, "SOL_FORM_H", "FORM Header", SAPbobsCOM.BoUTBTableType.bott_Document);
            Utils.CreateUDT(oSBOCompany, "SOL_FORM_D", "FORM Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

            Utils.CreateUDF(oSBOCompany, "@SOL_FORM_H", "SOL_CARDCODE", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            Utils.CreateUDF(oSBOCompany, "@SOL_FORM_H", "SOL_REF_NUM", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 200);
            Utils.CreateUDF(oSBOCompany, "@SOL_FORM_H", "SOL_TOTAL", "Document Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum);

            Utils.CreateUDF(oSBOCompany, "@SOL_FORM_D", "SOL_ITEMCODE", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            Utils.CreateUDF(oSBOCompany, "@SOL_FORM_D", "SOL_QTY", "Qty", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity);
            Utils.CreateUDF(oSBOCompany, "@SOL_FORM_H", "SOL_PRICE", "Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price);

            String[] FindColumnAlias = { "DocNum", "U_SOL_CARDCODE", "U_SOL_REF_NUM", "U_SOL_TOTAL"};
            Utils.CreateUDO(oSBOCompany, "OBJFORM1", SAPbobsCOM.BoUDOObjType.boud_Document, "SOL_FORM_H",
                           SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES,
                           SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO,
                           SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES,
                           FindColumnAlias,
                           new string[] { "SOL_FORM_D" });

            oSBOApplication.StatusBar.SetText(ADDON_NAME + " Add-On Connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

        }
        catch (Exception ex)
        {
            if (oSBOApplication != null)
                oSBOApplication.MessageBox(ex.Message);
            else
                MessageBox.Show(ex.Message);

            System.Windows.Forms.Application.Exit();
        }
    }
    private void SetupApplication()
    {
        SAPbouiCOM.SboGuiApi oSboGuiApi = null;
        string sConnectionString = null;

        oSboGuiApi = new SAPbouiCOM.SboGuiApi();
        sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

        //oSboGuiApi.AddonIdentifier = "56455230354241534953303030303030323537363A4E3034343935353039383594C39644338665F389E758409963DC8B0A02D1DF";
        oSboGuiApi.Connect(sConnectionString);

        oSBOApplication = oSboGuiApi.GetApplication();
        oSBOCompany = oSBOApplication.Company.GetDICompany();

        if (oSBOCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            GeneralVariables.SQLHandler = new HANAQueries();
        else
            GeneralVariables.SQLHandler = new SQLQueries();
    }
    private void NumericSeparators()
    {
        SAPbobsCOM.Recordset oRec;

        oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        oRec.DoQuery(GeneralVariables.SQLHandler.SeparatorSQL());

        GeneralVariables.WinDecSep = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        GeneralVariables.SBODecSep = oRec.Fields.Item("DecSep").Value.ToString();
        GeneralVariables.SBOThousSep = oRec.Fields.Item("ThousSep").Value.ToString();

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
        oRec = null;
        GC.Collect();
    }
    public void CatchingEvents()
    {
        // events handled by SBO_Application_AppEvent 
        oSBOApplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBOApplication_AppEvent);
        // events handled by SBO_Application_MenuEvent 
        oSBOApplication.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBOApplication_MenuEvent);
        // events handled by SBO_Application_ItemEvent
        oSBOApplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBOApplication_ItemEvent);
        // events handled by SBO_Application_ProgressBarEvent
        oSBOApplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBOApplication_FormDataEvent);
        // events handled by SBO_Application_StatusBarEvent
        oSBOApplication.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBOApplication_RightClickEvent);
        // events handled by SBO_Application_Printing
        oSBOApplication.LayoutKeyEvent += new SAPbouiCOM._IApplicationEvents_LayoutKeyEventEventHandler(SBOApplication_LayoutKeyEvent);
    }

    #region " SBO Event Handler "
    private void SBOApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
    {
        SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
        oSBOEventHandler.HandleAppEvent(EventType);
    }
    private void SBOApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
    {
        SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
        oSBOEventHandler.HandleMenuEvent(ref pVal, out BubbleEvent);
    }
    private void SBOApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
    {
        SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
        oSBOEventHandler.HandleItemEvent(FormUID, ref pVal, out BubbleEvent);
    }
    private void SBOApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
    {
        SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
        oSBOEventHandler.HandleFormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
    }
    private void SBOApplication_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
    {
        SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
        oSBOEventHandler.HandleRightClickEvent(ref eventInfo, out BubbleEvent);
    }
    private void SBOApplication_LayoutKeyEvent(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
    {
        SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
        oSBOEventHandler.HandleLayoutKeyEvent(ref eventInfo, out BubbleEvent);
    }
    #endregion


    private void addOutgoing()
    {
        // GENERATE OUTGOING PAYMENT 
        SAPbobsCOM.Payments vPay = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
        vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments;
        vPay.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;

        int iErrCode = 0;
        string sErrMsg = "";
        string errormessage = "";
        vPay.CardCode = "V1010";
        DateTime hariini = DateTime.Now;
        vPay.DocDate = hariini.Date;
        vPay.DueDate = hariini.Date;
        vPay.TaxDate = hariini.Date;
        vPay.TransferAccount = "161010";
        vPay.TransferDate = hariini.Date;
        vPay.Remarks = "TESTDI";

        vPay.TransferSum = 100;

        vPay.Invoices.SetCurrentLine(0);
        vPay.Invoices.DocEntry = 587;
        vPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice;

        vPay.Invoices.AppliedFC = 2;
        vPay.Invoices.Add();

        vPay.PrimaryFormItems.AmountFC = 2;
        vPay.PrimaryFormItems.CashFlowLineItemID = 18;
        vPay.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;

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
