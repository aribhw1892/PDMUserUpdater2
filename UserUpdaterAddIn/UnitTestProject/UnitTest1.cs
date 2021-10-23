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

namespace UnitTestProject
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestEntireCmd()
        {
            IEdmFile7 file; //Parent DrawingFile
            IEdmFolder5 folder;
            //IEdmVault7 vault;
            //Fileobject to find child drawings
            //vault = new EdmVault5();
            //Obtain the only instance of the IEdmVaultObject
            IEdmVault5 vault = EdmVaultSingleton.Instance;
            //vault.Login("admin", "", "CBFTest");
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
    }
}
