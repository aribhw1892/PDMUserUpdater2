using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Xml.Serialization;
using System.Windows.Forms;
using EPDM.Interop.epdm;
using System.Xml;
using System.Runtime.InteropServices;


namespace UserUpdaterAddIn
{
    public static class XMLHelper
    {
        public static ArrayList XmlUserReader(string xmlPath)
        {
            ArrayList xmlUserData = null;
            StreamReader StrReader = null;

            try
            {
                //Deserialize an XML file
                Type[] ExtraTypes = { Type.GetType("UserUpdaterAddIn.User") };
                XmlSerializer XmlSer = new XmlSerializer(Type.GetType("System.Collections.ArrayList"), ExtraTypes);
                StrReader = new StreamReader(xmlPath);
                xmlUserData = (ArrayList)XmlSer.Deserialize(StrReader);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return xmlUserData;
        }
        

    }
   
}
