using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Diagnostics;

namespace UserUpdaterAddIn
{
    public class User
    {
        //First name
        private string mSn;
        //Last name
        private string mGivenName;
        //Title
        private string mTitle;
        //Complete name
        private string mCn;
        //Email address
        private string mUsername;
        //Operation
        private string mOperation;
        //Login
        private string mLogin;
        //GroupName
        private string mGroup;

        public User()
        {

        }

        public string sn
        {
            get { return mSn; }
            set { mSn = value; }
        }

        public string givenName
        {
            get { return mGivenName; }
            set { mGivenName = value; }
        }

        public string title
        {
            get { return mTitle; }
            set { mTitle = value; }
        }

        public string cn
        {
            get { return mCn; }
            set { mCn = value; }
        }

        public string username
        {
            get { return mUsername; }
            set { mUsername = value; }
        }
        public string operation
        {
            get { return mOperation; }
            set { mOperation = value; }
        }

        public string login
        {
            get { return mLogin; }
            set { mLogin = value; }
        }
        public string group
        {
            get { return mGroup; }
            set { mGroup = value; }
        }

    }
}
