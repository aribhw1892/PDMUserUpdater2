using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UserUpdaterAddIn;
using EPDM.Interop.epdm;
using System.Diagnostics;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using System.Windows.Forms;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using System.Xml.Serialization;
using System.Xml;


namespace UnitTestProject
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestEntireCmd()
        {
            //IEdmFile7 file; //Parent DrawingFile
            //IEdmFolder5 folder;
            //IEdmVault7 vault;
            //Fileobject to find child drawings
            //vault = new EdmVault5();
            //Obtain the only instance of the IEdmVaultObject
            IEdmVault5 vault = EdmVaultSingleton.Instance;
            vault.Login("admin", "", "UserUpdater");

            EdmCmd cmd = new EdmCmd();
            cmd.mbsComment = "UserUpdaterAddIn";
            cmd.mpoVault = vault;

            EdmCmdData[] cmdDatas = new EdmCmdData[1];
            //cmdDatas[0].mlObjectID1 = 82103;                // 0000046460-001 APPROVAL.SLDDRW
            //cmdDatas[0].mlObjectID2 = 1756;
            cmdDatas[0].mlObjectID1 = 236668; //EB-0000087549-002.SLDDRW //236640; //0000087727
            cmdDatas[0].mlObjectID2 = 2703;

            UserUpdaterAddIn.UserUpdaterAddIn addin = new UserUpdaterAddIn.UserUpdaterAddIn();
            addin.OnCmd(ref cmd, ref cmdDatas);
        }
        

        [TestMethod]
        public void XmlUserReaderTest()
        {
            
            IEdmVault7 vault = EdmVaultSingleton.Instance;
            vault.Login("admin", "", "UserUpdater");

            string xmlPath = "G:\\WORK\\SOURCE_REPS\\PDMUserUpdater\\UserUpdaterAddIn\\UserUpdaterAddIn\\XML\\XMLUser.xml";
            //string xmlPath = Path.Combine(Environment.CurrentDirectory, @"XML\XMLUser.xml");
            ArrayList xmlData = UserUpdaterAddIn.XMLHelper.XmlUserReader(xmlPath);
        }

        [TestMethod]
        public void AddUserTest()
        {
            try
            {
                
                IEdmVault7 vault = EdmVaultSingleton.Instance;
                vault.Login("admin", "", "UserUpdater");

                string xmlPath = "G:\\WORK\\SOURCE_REPS\\PDMUserUpdater\\UserUpdaterAddIn\\UserUpdaterAddIn\\XML\\XMLUser.xml";
                //string xmlPath = Path.Combine(Environment.CurrentDirectory, @"XML\XMLUser.xml");
                ArrayList xmlData = UserUpdaterAddIn.XMLHelper.XmlUserReader(xmlPath);
                var addUserList = from User e in xmlData where e.operation.ToLower().Equals("add") select e;
                ArrayList addUser = new ArrayList();
                //for (int i = 0; i <= addUserList.Count() - 1; i++)
                foreach (var user in addUserList)
                {
                    addUser.Add(user);
                }

                UserUpdaterAddIn.UserHelper.AddUser(vault, xmlData);
            }
            catch(System.Runtime.InteropServices.COMException ex)
            {

            }
            catch
            {

            }
            

            // Can be passed the whole array without filtering ass existing User will not add.

        }

        [TestMethod]
        public void RemoveUserTest()
        {
            
            IEdmVault7 vault = EdmVaultSingleton.Instance;
            vault.Login("admin", "", "UserUpdater");

            string userName = "jsmith";
            IEdmUser9 user = UserHelper.GetUserObject(vault, userName);
            UserUpdaterAddIn.UserHelper.RemoveUser(vault, user);
        }

        //[TestMethod]
        public void GetUserObjectUpdateTest()
        {
            
            IEdmVault7 vault = EdmVaultSingleton.Instance;
            vault.Login("admin", "", "UserUpdater");

            string userName = "admin";
            IEdmUser9 user = UserHelper.GetUserObject(vault, userName);
            //UserHelper.UpdateUser(vault, user);
        }
    }
}
