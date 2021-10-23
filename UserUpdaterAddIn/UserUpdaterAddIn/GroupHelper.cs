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
    class GroupHelper
    {
		//Add Group To Vault

		//Remove Group From Vault

		//Update Group 

		//GetSingleGroupObject
		public static IEdmUserGroup8 GetGroupObject(IEdmVault7 vault, string groupName)
		{
			IEdmUserMgr10 usrMgr;
			//IEdmUser9 user;
			IEdmUserGroup8 group;
			group = null;
			try
			{
				usrMgr = (IEdmUserMgr10)vault;
				//Add efudd to Management group
				group = (IEdmUserGroup8)usrMgr.GetUserGroup(groupName);
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + "\n" + ex.Message);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			return group;
		}
	}
}
