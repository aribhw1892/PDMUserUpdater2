using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPDM.Interop.epdm;
using System.Collections;
using System.Windows.Forms;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using Microsoft.VisualBasic;
using System.Diagnostics;

namespace UserUpdaterAddIn
{
    public static class UserHelper
    {
        //Add User To PDM
        public static void AddUser(IEdmVault7 vault, ArrayList newUsers)
        {
            try
            {
				//Assign IEdmVault object to the IEdmUserMgr7 object
				IEdmUserMgr7 UsrMgr = (IEdmUserMgr7)vault;

				//Declare EdmUserData array to hold new user data
				EdmUserData2[] UserData = new EdmUserData2[newUsers.Count];

				//Set the EdmUserData members for each new user
				for (int i = 0; i <= newUsers.Count - 1; i++)
				{
					if (newUsers[i] != null)
					{
						UserData[i].mbsCompleteName = (newUsers[i] as User).cn;
						UserData[i].mbsEmail = (newUsers[i] as User).username;
						UserData[i].mbsInitials = (newUsers[i] as User).givenName.Substring(0, 1) + (newUsers[i] as User).sn.Substring(0, 1);
						UserData[i].mbsUserName = (newUsers[i] as User).username.Split('@')[0];
						//TODO: More Attributes To Add
						//TODO: Add group Attribute


						//Return user's IEdmUser6 interface in mpoUser
						UserData[i].mlFlags = (int)EdmUserDataFlags.Edmudf_GetInterface;
						//Add this user even if others cannot be added
						UserData[i].mlFlags += (int)EdmUserDataFlags.Edmudf_ForceAdd;

						//Set permissions 
						EdmSysPerm[] perms = new EdmSysPerm[3];
						perms[0] = EdmSysPerm.EdmSysPerm_EditUserMgr;
						perms[1] = EdmSysPerm.EdmSysPerm_EditReportQuery;
						perms[2] = EdmSysPerm.EdmSysPerm_MandatoryVersionComments;
						UserData[i].moSysPerms = perms;
					}
				}

				//Add the users to the vault
				UsrMgr.AddUsers2(ref UserData);

				//Logging User Add data 
				string addUserlog = "";
				foreach (EdmUserData2 usr in UserData)
				{
					if (usr.mhStatus == 0)
					{
						addUserlog += "Created user \"" + usr.mpoUser.Name + "\" successfully. ID = " + usr.mpoUser.ID.ToString() + "\n";
					}
					else
					{
						IEdmVault11 vault2 = (IEdmVault11)vault;
						addUserlog += "Error creating user \"" + usr.mbsUserName + "\" - " + vault2.GetErrorMessage(usr.mhStatus) + "\n";
					}
				}
				//TODO: Need to check as log is not writing
				LogWriter.WriteLog(addUserlog);
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

        //Remove User From PDM
        public static void RemoveUser(IEdmVault7 vault, IEdmUser9 user)
        {
			IEdmUserMgr10 UsrMgr;
			try
			{
				//Remove the user from the vault
				UsrMgr = (IEdmUserMgr10)vault.CreateUtility(EdmUtility.EdmUtil_UserMgr);
				//user = (IEdmUser9)UsrMgr.GetUser(userName);
				if ((user == null))
				{
					Debug.Print("No more user to remove");
					return;
				}

				int[] users = new int[1];
				users[0] = user.ID;
				UsrMgr.RemoveUsers(users);

				Debug.Print("User " + user.Name + " removed.");
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + "\n" + ex.Message);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);

			}
		}

		//GetSingleUserObject
		public static IEdmUser9 GetUserObject(IEdmVault7 vault, string userName)
        {
			IEdmUserMgr10 UsrMgr;
			IEdmUser9 user;
			user = null;
			try
			{
				UsrMgr = (IEdmUserMgr10)vault.CreateUtility(EdmUtility.EdmUtil_UserMgr);
				user = (IEdmUser9)UsrMgr.GetUser(userName);
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + "\n" + ex.Message);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			return user;
		}

		//Update User To PDM
		public static void UpdateUser(IEdmVault7 vault, IEdmUser9 user, User userValues)
		{
			try
			{
				//Get the user search interface 
				IEdmFindUser poFind = default(IEdmFindUser);
				poFind = (IEdmFindUser)vault.CreateUtility(EdmUtility.EdmUtil_FindUser);

				//Search for a user with LoginName
				poFind.SetPropt(EdmFindUserProp.Edmfup_LoginName, user.Name);
				string val = null;
				val = (string)poFind.GetPropt(EdmFindUserProp.Edmfup_LoginName);
				poFind.SilentFind();
				IEdmEnum poResult = default(IEdmEnum);
				poResult = poFind.Result;
                IEdmUser10 poUser = default(IEdmUser10);
				EdmUserDataEx UserInfo = new EdmUserDataEx();

				//Specify which user data fields are valid
				UserInfo.mlEdmUserDataExFlags = (int)EdmUserDataExFlag.Edmudex_All;

				foreach (object foundUser in poResult)
				{
					poUser = (IEdmUser10)foundUser;

					//Get user's information
					poUser.GetUserDataEx(ref UserInfo);

					//TODO: Update user's information - Add Data Model
					UserInfo.mbsCompleteName = userValues.cn.ToString();
					UserInfo.mbsEmail = userValues.username.ToString();
					UserInfo.mbsInitials = userValues.givenName.Substring(0, 1) + userValues.sn.Substring(0, 1); ;
					poUser.SetUserDataEx(ref UserInfo);
				}
				//Update UserData
				user.UserData = "Changed System Addmin User Data 2";
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + "\n" + ex.Message);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
		
		//Get All Users From Vault
		public static List<string> GetAllUserFromVault(IEdmVault7 vault)
        {

			List<string> vaultUsers = new List<string>();
			IEdmUserMgr9 UserMgr = (IEdmUserMgr9)vault.CreateUtility(EdmUtility.EdmUtil_UserMgr);

            EdmStatePermission[] ppoPermissions = {};
            EdmTransitionPermission[] ppoTransitionPermissions = {};

            string Users = null;
            IEdmPos5 UserPos = default(IEdmPos5);
            IEdmUser5 User = default(IEdmUser5);
            UserPos = UserMgr.GetFirstUserPosition();
            //Collecting Users
            while (!UserPos.IsNull)
            {
                User = UserMgr.GetNextUser(UserPos);
                Users = Users + User.Name + " ID: " + User.ID + "\n";
				vaultUsers.Add(User.Name);
			}
			return vaultUsers;
		}

        //Add User To Group
		public static void AddUserToGroup(IEdmVault7 vault, IEdmUserGroup8 group, IEdmUser9 user)
        {
			IEdmUserMgr10 usrMgr;
			//IEdmUser9 user;
			
			try
			{
				usrMgr = (IEdmUserMgr10)vault;

				//Add efudd to Management group
				//mngmtGroup = (IEdmUserGroup8)usrMgr.GetUserGroup(groupName);

				if ((group == null))
				{
					//MessageBox.Show("Group does not exist. Create a Management group.");
					return;
				}

				//user = (IEdmUser9)UsrMgr.GetUser("efudd");

				if ((user == null))
				{
					//MessageBox.Show("No user set to add to group. Click Add users.");
					return;
				}

				int[] groupMbrIDs = new int[1];
				groupMbrIDs[0] = user.ID;
				group.AddMembers(groupMbrIDs);
				

				Debug.Print("User " + user.Name + " added to " + group.Name + " group.");

			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + "\n" + ex.Message);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);

			}
		}

        //Remove User From Group
		public static void RemoveUserFromGroup(IEdmVault7 vault, IEdmUserGroup8 group, IEdmUser9 user)
        {
			//Array Of Folders
			EdmMemberFolder[] folderMembers = new EdmMemberFolder[1]; 
			
			try
			{
				if ((group == null))
				{
					MessageBox.Show("No group set from which to remove group member. Click Add group member.");
					return;
				}

				//Remove user from Test folder, Management group, and vault -**Need to check
				group.RemoveMembers(folderMembers);
				//user = (IEdmUser9)UserMgr.GetUser("efudd");

				if ((user == null))
				{
					MessageBox.Show("No user set to remove from group. Click Add users.");
					return;
				}

				MessageBox.Show("User " + user.Name + " removed from group and vault.");


			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + "\n" + ex.Message);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);

			}
		}
    }
}
