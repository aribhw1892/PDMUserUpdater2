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

        public static ArrayList AdXmlReader(string xmlPath)
        {
            ArrayList xmlUserData = new ArrayList();
            XmlDocument xdc = new XmlDocument();
            xdc.Load(xmlPath);
            XmlElement root = xdc.DocumentElement;
            //CreatexmlPath Expression as per the xml
            var nodes = root.SelectNodes("/Objs/Obj/MS");
            //var nodes1 = root.ChildNodes;
            //var nodes = xdc.GetElementsByTagName("\"MS\"");
            foreach (XmlNode node in nodes)
            {
                User user = new User();
                
                //Chevk if UserName is null
                if (!String.IsNullOrEmpty(node.FirstChild.InnerXml.ToString()))
                {
                    XmlNodeList childs = node.ChildNodes;
                    foreach (XmlNode childnode in childs)
                    {
                        if (childnode.Attributes["N"].Value == "UserName")
                        {
                            string userName = childnode.InnerText;
                            user.login = userName;
                            user.cn = userName;
                            user.givenName = "";
                            user.sn = "";
                            user.title = "";
                            Debug.Print(userName);
                        }
                        if (childnode.Attributes["N"].Value == "Groups")
                        {
                            if (!String.IsNullOrEmpty(childnode.InnerText.ToString()))
                            {
                                string group = childnode.InnerText;
                                user.group = group;
                                Debug.Print(group);
                            }
                            else
                            {
                                user.group = "";
                                Debug.Print("Nogroup");
                            }
                        }
                        if (childnode.Attributes["N"].Value == "DesktopProfile")
                        {
                            if (!String.IsNullOrEmpty(childnode.InnerText.ToString()))
                            {
                                if (childnode.InnerText.ToString() == "EPDM")
                                {
                                    string desktopProfile = childnode.InnerText;
                                    user.operation = desktopProfile;
                                    Debug.Print(desktopProfile);
                                    //Add to ArrayList if desktop Profile is EPDM
                                    xmlUserData.Add((User)user);
                                }
                                else
                                {
                                    user.operation = "";
                                    Debug.Print("Nogroup");
                                }
                            }
                            
                        }
                    }
                   
                }
            }

            return xmlUserData;
        }
    }
   
}
