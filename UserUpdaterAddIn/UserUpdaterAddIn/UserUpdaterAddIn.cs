using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using EPDM.Interop.epdm;
using System.Windows.Forms;
using System.IO;
using Microsoft.VisualBasic;
using System.Diagnostics;

namespace UserUpdaterAddIn
{
    [Guid("FD1CE079-A768-4CC4-99BB-5A5E287410D5"), ComVisible(true)]
    public class UserUpdaterAddIn : IEdmAddIn5
    {
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            // Set the add in's properties
            poInfo.mbsAddInName = "PDMUserGroupsUpdater";
            poInfo.mbsCompany = "CAD-AI";
            poInfo.mbsDescription = "Provides users and Groups update from XML capabilities.";
            poInfo.mlAddInVersion = 1;

            // Set the required versionh
            poInfo.mlRequiredVersionMajor = 20;
            poInfo.mlRequiredVersionMinor = 0;

            // Register any commands
            //// Register a menu command
            //poCmdMgr.AddCmd(1, "C# Add-in", (int)EdmMenuFlags.EdmMenu_Nothing);
            //Activate the Addin by clicking Button
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_CardButton);
        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData) //EdmCmdData[]
        {
            IEdmFile8 drawing;
            IEdmFolder5 folder;
            IEdmVault7 vault;

            // Check Command
            //// Handle the menu command
            //if (poCmd.meCmdType == EdmCmdType.EdmCmd_Menu)
            //{
            //    if (poCmd.mlCmdID == 1)
            //    {
            //        System.Windows.Forms.MessageBox.Show("C# Add-in");
            //    }
            //}
            if ((poCmd.mbsComment) == "UserUpdater")
            {
                int winHandle = poCmd.mlParentWnd;
                // Get Vault Object
                // Obtain the only instance of the IEdmVaultObject
                vault = EdmVaultSingleton.Instance;


            }

        }
    }
}
