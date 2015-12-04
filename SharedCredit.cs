 using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;


namespace AUTORIVET_KAOHE
{
internal class NativeCredMan
        {






       /// <summary>
        /// 凭据类型
        /// </summary>
        public enum CRED_TYPE : uint
        {
            //普通凭据
            GENERIC = 1, 
            //域密码
            DOMAIN_PASSWORD = 2,
            //域证书
            DOMAIN_CERTIFICATE = 3,
            //域可见密码
            DOMAIN_VISIBLE_PASSWORD = 4,
            //一般证书
            GENERIC_CERTIFICATE = 5,
            //域扩展
            DOMAIN_EXTENDED = 6,
            //最大
            MAXIMUM = 7,      // Maximum supported cred type
            MAXIMUM_EX = (MAXIMUM + 1000),  // Allow new applications to run on old OSes
        }

        //永久性
        public enum CRED_PERSIST : uint
        {
            SESSION = 1,
            //本地计算机
            LOCAL_MACHINE = 2,
            //企业
            ENTERPRISE = 3,
        }


            [DllImport("Advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
            //读取凭据信息
            static extern bool CredRead(string target, CRED_TYPE type, int reservedFlag, out IntPtr CredentialPtr);

            [DllImport("Advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
            //增加凭据
            static extern bool CredWrite([In] ref NativeCredential userCredential, [In] UInt32 flags);

            [DllImport("Advapi32.dll", EntryPoint = "CredFree", SetLastError = true)]
            static extern bool CredFree([In] IntPtr cred);

            [DllImport("Advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode)]
            //删除凭据
            static extern bool CredDelete(string target, CRED_TYPE type, int flags);

            //[DllImport("advapi32", SetLastError = true, CharSet = CharSet.Unicode)]
            //static extern bool CredEnumerateold(string filter, int flag, out int count, out IntPtr pCredentials);

            [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            public static extern bool CredEnumerate(string filter, uint flag, out uint count, out IntPtr pCredentials);

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            private struct NativeCredential
            {
                public UInt32 Flags;
                public CRED_TYPE Type;
                public IntPtr TargetName;
                public IntPtr Comment;
                public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
                public UInt32 CredentialBlobSize;
                public IntPtr CredentialBlob;
                public UInt32 Persist;
                public UInt32 AttributeCount;
                public IntPtr Attributes;
                public IntPtr TargetAlias;
                public IntPtr UserName;

                internal static NativeCredential GetNativeCredential(Credential cred)
                {
                    var ncred = new NativeCredential
                                    {
                                        AttributeCount = 0,
                                        Attributes = IntPtr.Zero,
                                        Comment = IntPtr.Zero,
                                        TargetAlias = IntPtr.Zero,
                                        //Type = CRED_TYPE.DOMAIN_PASSWORD,
                                        Type=cred.Type,
                                        Persist = (UInt32)cred.Persist,
                                        CredentialBlobSize = (UInt32)cred.CredentialBlobSize,
                                        TargetName = Marshal.StringToCoTaskMemUni(cred.TargetName),
                                        CredentialBlob = Marshal.StringToCoTaskMemUni(cred.CredentialBlob),
                                        UserName = Marshal.StringToCoTaskMemUni(cred.UserName)
                                    };
                    return ncred;
                }
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            public struct Credential
            {
                public UInt32 Flags;
                public CRED_TYPE Type;
                public string TargetName;
                public string Comment;
                public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
                public UInt32 CredentialBlobSize;
                public string CredentialBlob;
                public CRED_PERSIST Persist;
                public UInt32 AttributeCount;
                public IntPtr Attributes;
                public string TargetAlias;
                public string UserName;
            }

            /// <summary>
            /// 向添加计算机的凭据管理其中添加凭据
            /// </summary>
            /// <param name="key">internet地址或者网络地址</param>
            /// <param name="userName">用户名</param>
            /// <param name="secret">密码</param>
            /// <param name="type">密码类型</param>
            /// <param name="credPersist"></param>
            /// <returns></returns>
            public static int WriteCred(string key, string userName, string secret, CRED_TYPE type, CRED_PERSIST credPersist)
            {

                var byteArray = Encoding.Unicode.GetBytes(secret);
                if (byteArray.Length > 512)
                    throw new ArgumentOutOfRangeException("The secret message has exceeded 512 bytes.");

                var cred = new Credential
                               {
                                   TargetName = key,
                                   CredentialBlob = secret,
                                   CredentialBlobSize = (UInt32)Encoding.Unicode.GetBytes(secret).Length,
                                   AttributeCount = 0,
                                   Attributes = IntPtr.Zero,
                                   UserName = userName,
                                   Comment = null,
                                   TargetAlias = null,
                                   Type = type,
                                   Persist = credPersist
                               };
                var ncred = NativeCredential.GetNativeCredential(cred);

                var written = CredWrite(ref ncred, 0);
                var lastError = Marshal.GetLastWin32Error();
                if (written)
                {
                    return 0;
                }
                var message = "";
                if (lastError == 1312)
                {
                    message = (string.Format("Failed to save " + key + " with error code {0}.", lastError) + "  This error typically occurrs on home editions of Windows XP and Vista.  Verify the version of Windows is Pro/Business or higher.");
                }
                else
                {
                    message = string.Format("Failed to save " + key + " with error code {0}.", lastError);
                }
                MessageBox.Show(message);
                return 1;
            }

           /// <summary>
           /// 读取凭据
           /// </summary>
           /// <param name="targetName"></param>
           /// <param name="credType"></param>
           /// <param name="reservedFlag"></param>
           /// <param name="intPtr"></param>
           /// <returns></returns>
            public static bool WReadCred(string targetName,CRED_TYPE credType,int reservedFlag,out IntPtr intPtr)
            {
                return CredRead(targetName, CRED_TYPE.DOMAIN_PASSWORD, reservedFlag, out intPtr);
               
            }

            /// <summary>
            /// 删除凭据
            /// </summary>
            /// <param name="target"></param>
            /// <param name="type"></param>
            /// <param name="flags"></param>
            /// <returns></returns>
            public static bool DeleteCred(string target, CRED_TYPE type, int flags)
            {
                return CredDelete(target, type, flags);

            }
        }

        /// <summary>
        /// 查询凭据是否存在
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //private void btnQuery_Click(object sender, EventArgs e)
        //{
        //    string targetName = txtIP.Text.Trim();
        //    IntPtr intPtr=new IntPtr ();

        //    bool flag = false;
        //    try
        //    {
        //        flag = NativeCredMan.WReadCred(targetName, CRED_TYPE.DOMAIN_PASSWORD, 1, out intPtr);
        //    }
        //    catch
        //    {
        //        flag = false;
        //    }
        //    if (flag)
        //        txtMsg.Text = "该凭据已存在";
        //    else
        //        txtMsg.Text = "该凭据目前不存在";
        //}


        //private void btnDelete_Click(object sender, EventArgs e)
        //{
        //    string targetName = txtIP.Text.Trim();

        //    bool flag = false;
        //    try
        //    {
        //        flag = NativeCredMan.DeleteCred(targetName, CRED_TYPE.DOMAIN_PASSWORD,0);
        //    }
        //    catch
        //    {
        //        flag = false;
        //    }
        //    if (flag)
        //        txtMsg.Text = "该凭据已删除";
        //    else
        //        txtMsg.Text = "删除失败";
        //}
 }
