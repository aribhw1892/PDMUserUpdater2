using System;
using System.Collections;
using System.Diagnostics;
using System.Threading;
using EPDM.Interop.epdm;

namespace UserUpdaterAddIn
{
    public class EdmVaultSingleton
    {
        private static EdmVault5 mInstance = null;
        private static object mLockObj = new object();
        private EdmVaultSingleton()
        {

        }
        public static EdmVault5 Instance
        {
            get
            {
                try
                {
                    if (mInstance == null)
                    {
                        Monitor.Enter(mLockObj);
                        if (mInstance == null)
                        {
                            mInstance = new EdmVault5();
                        }
                        Monitor.Exit(mLockObj);
                    }
                }
                catch (Exception ex)
                {
                    Monitor.Exit(mLockObj);
                }

                return mInstance;
            }
        }
    }
}
