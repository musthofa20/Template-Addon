using System;

public class XMLLoader
{

    public void LoadFromXML(SAPbouiCOM.Application oSBOApplication, string FileName)
    {
        System.Xml.XmlDocument oXmlDoc;
        oXmlDoc = new System.Xml.XmlDocument();

        LoadFromXML(oSBOApplication, oXmlDoc, FileName);
    }

    public void LoadFromXML(SAPbouiCOM.Application oSBOApplication, System.Xml.XmlDocument oXMLDoc, string FileName)
    {
        string sPath;
        // sPath = IO.Directory.GetParent(Application.StartupPath).ToString

        try
        {
            oXMLDoc.Load(FileName);
            oSBOApplication.LoadBatchActions(oXMLDoc.InnerXml);
            sPath = oSBOApplication.GetLastBatchResults();
        }
        catch (Exception ex)
        {
            oSBOApplication.MessageBox(ex.Message);
        }
    }

    public void LoadFromXML(SAPbouiCOM.Application oSBOApplication, System.Xml.XmlDocument oXMLDoc, string FileName, string sFormUID)
    {
        string sPath;
        // sPath = IO.Directory.GetParent(Application.StartupPath).ToString

        try
        {
            oXMLDoc.Load(FileName);
            oXMLDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = oXMLDoc.SelectSingleNode("Application/forms/action/form/@uid").Value + sFormUID;
            oSBOApplication.LoadBatchActions(oXMLDoc.InnerXml);
            sPath = oSBOApplication.GetLastBatchResults();
        }
        catch (Exception ex)
        {
            oSBOApplication.MessageBox(ex.Message);
        }
    }

}
