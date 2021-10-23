using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPDM.Interop.epdm;
using System.Collections;
using System.Windows.Forms;

namespace UserUpdaterAddIn
{
    class UserHelper
    {
        //Add User To PDM
        public static void AddUser(IEdmVault7 vault, ArrayList NewUsers)
        {
            try
            {
				//Assign IEdmVault object to the IEdmUserMgr7 object
				IEdmUserMgr7 UsrMgr = (IEdmUserMgr7)vault;

				//Declare EdmUserData array to hold new user data
				EdmUserData2[] UserData = new EdmUserData2[NewUsers.Count];

				//Set the EdmUserData members for each new user
				for (int i = 0; i <= NewUsers.Count - 1; i++)
				{
					if (NewUsers[i] != null)
					{
						UserData[i].mbsCompleteName = (NewUsers[i] as User).cn;
						UserData[i].mbsEmail = (NewUsers[i] as User).username;
						UserData[i].mbsInitials = (NewUsers[i] as User).givenName.Substring(0, 1) + (NewUsers[i] as User).sn.Substring(0, 1);
						UserData[i].mbsUserName = (NewUsers[i] as User).username.Split('@')[0];
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

				string msg = "";
				foreach (EdmUserData2 usr in UserData)
				{
					if (usr.mhStatus == 0)

					{
						msg += "Created user \"" + usr.mpoUser.Name + "\" successfully. ID = " + usr.mpoUser.ID.ToString() + "\n";

					}
					else
					{
						IEdmVault11 vault2 = (IEdmVault11)vault;
						msg += "Error creating user \"" + usr.mbsUserName + "\" - " + vault2.GetErrorMessage(usr.mhStatus) + "\n";
					}
				}
				//MessageBox.Show(msg);
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
					MessageBox.Show("No user set to remove. Click Add users.");
					return;
				}

				int[] users = new int[1];
				users[0] = user.ID;
				UsrMgr.RemoveUsers(users);

				MessageBox.Show("User " + user.Name + " removed.");

				////Send message to all users with permission
				////to update users and groups 
				//IEdmPos5 UserPos = default(IEdmPos5);
				//UserPos = UsrMgr.GetFirstUserPosition();
				//while (!UserPos.IsNull)
				//{
				//	IEdmUser9 userWithPerm = default(IEdmUser9);
				//	userWithPerm = (IEdmUser9)UsrMgr.GetNextUser(UserPos);
				//	if (userWithPerm.IsLoggedIn)
				//	{
				//		if (userWithPerm.HasSysRightEx(EdmSysPerm.EdmSysPerm_EditUserMgr))
				//		{
				//			userWithPerm.SendMsg("ALERT: user removed", "User " + user.Name + " removed.");
				//		}
				//	}
				//}

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
		public static void GetAllUserFromVault(IEdmVault7 vault)
        {
            //Declare an IEdmUserMgr9 object
            //IEdmUserMgr9 UserMgr = default(IEdmUserMgr9);
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
            }

            MessageBox.Show(Users, vault.Name + " Vault Users", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //Get permissions for all states for a user 
            UserMgr.GetStatePermissions(User.ID, EdmObjectType.EdmObject_User, 0, out ppoPermissions);
            string str = null;
            str = "EdmStatePermissions for a user in the vault" + "\n";
            str = str + "\n";
            foreach (EdmStatePermission item in ppoPermissions)
            {

                str = str + "mlOwnerID: " + item.mlOwnerID + "\n";
                str = str + "meOwnerType (EdmObjectType.EdmObject_User (7) or EdmObjectType.EdmObject_UserGroup (8)): " + item.meOwnerType + "\n";
                str = str + "mlStateID: " + item.mlStateID + "\n";
                str = str + "mlEdmRightFlag as defined in EdmRightFlags: " + item.mlEdmRightFlag + "\n";
                str = str + "\n";

            }
            MessageBox.Show(str);

            //Get permissions for all transitions for a user 
            UserMgr.GetTransitionPermissions(User.ID, EdmObjectType.EdmObject_User, 0, out ppoTransitionPermissions);
            str = "EdmTransitionPermissions for a user in the vault" + "\n";
            str = str + "\n";
            foreach (EdmTransitionPermission item in ppoTransitionPermissions)
            {

                str = str + "mlOwnerID: " + item.mlOwnerID + "\n";
                str = str + "meOwnerType (EdmObjectType.EdmObject_User (7) or EdmObjectType.EdmObject_UserGroup (8)): " + item.meOwnerType + "\n";
                str = str + "mlTransitionID: " + item.mlTransitionID + "\n";
                str = str + "mlEdmRightFlag as defined in EdmTransitionRightFlags: " + item.mlEdmRightFlag + "\n";
                str = str + "\n";

            }
            MessageBox.Show(str);
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
					MessageBox.Show("Group does not exist. Create a Management group.");
					return;
				}

				//user = (IEdmUser9)UsrMgr.GetUser("efudd");

				if ((user == null))
				{
					MessageBox.Show("No user set to add to group. Click Add users.");
					return;
				}

				int[] groupMbrIDs = new int[1];
				groupMbrIDs[0] = user.ID;
				group.AddMembers(groupMbrIDs);
				

				MessageBox.Show("User " + user.Name + " added to " + group.Name + " group.");

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
