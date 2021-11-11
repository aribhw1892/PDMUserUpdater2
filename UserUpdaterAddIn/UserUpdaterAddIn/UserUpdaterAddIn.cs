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
using System.Xml.Serialization;
using System.Xml;
using System.Collections;

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
            //IEdmFile8 drawing;
            //IEdmFolder5 folder;
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
            if ((poCmd.mbsComment) == "UserUpdaterAddIn")
            {
                int winHandle = poCmd.mlParentWnd;
                IEdmUser9 user;
                IEdmUserGroup8 group;
                // Get Vault Object
                // Obtain the only instance of the IEdmVaultObject
                vault = EdmVaultSingleton.Instance;

                
                //Get All User Data from XML
                    //*8Change before Release
                string xmlPath = "G:\\WORK\\SOURCE_REPS\\PDMUserUpdater\\UserUpdaterAddIn\\UserUpdaterAddIn\\XML\\UserGroupV1.xml";
                //string xmlPath = Path.Combine(Environment.CurrentDirectory, @"XML\XMLUser.xml");
                //ArrayList xmlUserData = XMLHelper.XmlUserReader(xmlPath);
                ArrayList xmlUserData = XMLHelper.AdXmlReader(xmlPath);

                ////Add USER
                ////Get users To Add
                //ArrayList addUserList = GetUserToAdd(xmlUserData);
                ////Add Users -Removed as it is not a requirement
                //UserHelper.AddUser(vault, addUserList);
                
                //Show Users Not In PDM.
                List<string> vaultUsers = UserHelper.GetAllUserFromVault(vault);
                ArrayList userNotInPDM = UserNotInPDM(xmlUserData, vaultUsers);
                string userNotInPDMstr = "";
                for (int i = 0; i <= userNotInPDM.Count - 1; i++)
                {
                    //Get User Object
                    userNotInPDMstr = userNotInPDMstr + "\n" + (userNotInPDM[i] as User).login.ToString();
                    //UserHelper.UpdateUser(vault, user, (userNotInPDM[i] as User));
                }
                MessageBox.Show("Below Users are not in PDM:" + "\n" + userNotInPDMstr);

                //// Add User To Group
                ////Get users To Remove
                //ArrayList groupUserList = GetUserToAdd(xmlUserData);
                ////Remove Users
                //for (int i = 0; i <= groupUserList.Count - 1; i++)
                //{
                //    //Get User Object
                //    user = UserHelper.GetUserObject(vault, (groupUserList[i] as User).username.Split('@')[0]);
                //    //Get Group Object
                //    group = GroupHelper.GetGroupObject(vault, (groupUserList[i] as User).group.ToString());
                //    UserHelper.AddUserToGroup(vault, group, user);
                //}

                // Remove Users
                //Get users To Remove
                ArrayList removeUserList = GetUserToRemove(xmlUserData, vaultUsers);
                //Remove Users
                string userRemovedstr = "";
                for (int i = 0; i <= removeUserList.Count - 1; i++)
                {
                    //Get User Object
                    user = UserHelper.GetUserObject(vault, (removeUserList[i] as User).login);
                    if ((removeUserList[i] as User).login != "Admin")
                    {
                        UserHelper.RemoveUser(vault, user);
                        userRemovedstr = userRemovedstr + "\n" + (removeUserList[i] as User).login.ToString();
                    }
                    
                }
                MessageBox.Show("Removed Users Not Available in Active Directory:" + "\n" + userRemovedstr);
            }
        }
        // Get All User To Add 
        public static ArrayList GetUserToAdd(ArrayList allUserData)
        {
            ArrayList addUser = new ArrayList();
            var addUserList = from User e in allUserData where e.operation.ToLower().Equals("add") select e;
            foreach (var user in addUserList)
            {
                addUser.Add(user);
            }

            return addUser;
        }
        // Get All User To Update
        public static ArrayList UserNotInPDM(ArrayList allUserData, List<string> vaultUsers)
        {
            ArrayList userNotInPDM = new ArrayList();
            for (int i = 0; i <= allUserData.Count - 1; i++)
            {
                if (vaultUsers.Any(e => e == (allUserData[i] as User).login))
                //if (vaultUsers.Any(e => e == (allUserData[i] as User).username.ToLower().Split('@')[0]))
                {
                    
                }
                else
                {
                    userNotInPDM.Add(allUserData[i] as User);
                }
                
            }
            return userNotInPDM;
        }
        public static ArrayList GetUserToRemove(ArrayList allUserData, List<string> vaultUsers)
        {
            ArrayList removeUser = new ArrayList();
            foreach (string user in vaultUsers)
            {
                var test = from User e in allUserData where e.login.Equals(user) select e;
                if (test.Count() == 0)
                {
                    User removeUserObj = new User();
                    removeUserObj.login = user;
                    removeUser.Add(removeUserObj);
                }
            }
            return removeUser;
        }
        // Get All User To Remove
        //public static ArrayList GetUserToRemove(ArrayList allUserData)
        //{
        //    ArrayList removeUser = new ArrayList();
        //    var addUserList = from User e in allUserData where e.operation.ToLower().Equals("remove") select e;
        //    foreach (var user in addUserList)
        //    {
        //        removeUser.Add(user);
        //    }

        //    return removeUser;
        //}



    }
}
