using System;
using System.Net;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using OpcRcw.Comn;

namespace Siemens.Opc
{
    public class ServerEnum
    {
        /// <summary>
        /// The CLSID for the OPC Server Enumerator.
        /// </summary>
        private static readonly Guid CLSID = new Guid("13486D51-4821-11D2-A494-3CB306C10000");

        /// <summary>
        /// Return a list of local OPC A&E servers URLs.
        /// </summary>
        public static string[] GetLocalAeServers()
        {
            IOPCServerList2 m_server = null;

            // Establish connection to the OPC Server
            m_server = (IOPCServerList2)Helper.CreateInstance(CLSID, "localhost", null);

            try
            {
                ArrayList servers = new ArrayList();
                string[] urls = null;

                // convert the interface version to a guid.
                Guid catid = new Guid("58E13251-AC87-11d1-84D5-00608CB8A7E9");

                // get list of servers in the specified specification.
                IOPCEnumGUID enumerator = null;

                m_server.EnumClassesOfCategories(
                    1,
                    new Guid[] { catid },
                    0,
                    null,
                    out enumerator);

                // read clsids.
                Guid[] clsids = ReadClasses(enumerator);

                enumerator = null;

                urls = new string[clsids.Length];
                int i = 0;

                // fetch class descriptions.
                foreach (Guid clsid in clsids)
                {
                    string progId;
                    string sTemp1;
                    string sTemp2;
                    Guid tempClsid = clsid;

                    try
                    {
                        m_server.GetClassDetails(
                            ref tempClsid,
                            out progId,
                            out sTemp1,
                            out sTemp2);

                        sTemp1 = "opcae://localhost/" + progId;
                        urls[i] = sTemp1;
                        i++;
                    }
                    catch (Exception)
                    {
                        // ignore bad clsids.
                    }
                }

                // free the server.
                Helper.ReleaseServer(m_server);
                m_server = null;

                return urls;
            }
            finally
            {
                // free the server.
                Helper.ReleaseServer(m_server);
                m_server = null;
            }
        }

        /// <summary>
        /// Reads the guids from the enumerator.
        /// </summary>
        private static Guid[] ReadClasses(IOPCEnumGUID enumerator)
        {
            ArrayList guids = new ArrayList();

            int fetched = 0;
            Guid[] buffer = new Guid[10];
            //IntPtr rgelt = new IntPtr();

            do
            {
                try
                {
                    //enumerator.Next(buffer.Length, rgelt, out fetched);
                    enumerator.Next(buffer.Length, buffer, out fetched);

                    for (int ii = 0; ii < fetched; ii++)
                    {
                        guids.Add(buffer[ii]);
                    }
                }
                catch
                {
                    break;
                }
            }
            while (fetched > 0);

            return (Guid[])guids.ToArray(typeof(Guid));
        }
    }

    class Helper
    {
        #region Type Conversion
        /// <summary>
        /// Convert a WIN32 FILETIME to DateTime.
        /// </summary>
        public static DateTime FiletimeToDateTime(int filetimeLow, int filetimeHigh)
        {
            // convert FILETIME structure to a 64 bit integer.
            long buffer = (long)filetimeHigh;

            if (buffer < 0)
            {
                buffer += ((long)UInt32.MaxValue + 1);
            }

            long ticks = (buffer << 32);

            buffer = (long)filetimeLow;

            if (buffer < 0)
            {
                buffer += ((long)UInt32.MaxValue + 1);
            }

            ticks += buffer;

            // check for invalid value.
            if (ticks == 0)
            {
                return DateTime.MinValue;
            }

            DateTime output = new DateTime(1601, 1, 1);
            return output.Add(new TimeSpan(ticks));
        }

        /// <summary>
        /// Convert a DateTime to WIN32 FILETIME.
        /// </summary>
        public static void DateTimeToFiletime(DateTime datetime, out int filetimeLow, out int filetimeHigh)
        {
            DateTime minVal = new DateTime(1601, 1, 1);

            if (datetime <= minVal)
            {
                filetimeHigh = 0;
                filetimeLow = 0;
                return;
            }

            // adjust for WIN32 FILETIME base.
            long ticks = 0;

            ticks = datetime.Subtract(new TimeSpan(minVal.Ticks)).Ticks;

            filetimeHigh = (int)((ticks >> 32) & 0xFFFFFFFF);
            filetimeLow = (int)(ticks & 0xFFFFFFFF);
        }

        /// <summary>
        /// Converts the VARTYPE to a system type.
        /// </summary>
        internal static System.Type GetType(VarEnum input)
        {
            switch (input)
            {
                case VarEnum.VT_EMPTY: return null;
                case VarEnum.VT_I1: return typeof(sbyte);
                case VarEnum.VT_UI1: return typeof(byte);
                case VarEnum.VT_I2: return typeof(short);
                case VarEnum.VT_UI2: return typeof(ushort);
                case VarEnum.VT_I4: return typeof(int);
                case VarEnum.VT_UI4: return typeof(uint);
                case VarEnum.VT_I8: return typeof(long);
                case VarEnum.VT_UI8: return typeof(ulong);
                case VarEnum.VT_R4: return typeof(float);
                case VarEnum.VT_R8: return typeof(double);
                case VarEnum.VT_CY: return typeof(decimal);
                case VarEnum.VT_BOOL: return typeof(bool);
                case VarEnum.VT_DATE: return typeof(DateTime);
                case VarEnum.VT_BSTR: return typeof(string);
                case VarEnum.VT_ARRAY | VarEnum.VT_I1: return typeof(sbyte[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_UI1: return typeof(byte[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_I2: return typeof(short[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_UI2: return typeof(ushort[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_I4: return typeof(int[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_UI4: return typeof(uint[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_I8: return typeof(long[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_UI8: return typeof(ulong[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_R4: return typeof(float[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_R8: return typeof(double[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_CY: return typeof(decimal[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_BOOL: return typeof(bool[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_DATE: return typeof(DateTime[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_BSTR: return typeof(string[]);
                case VarEnum.VT_ARRAY | VarEnum.VT_VARIANT: return typeof(object[]);
                default: return null;
            }
        }

        /// <summary>
        /// The size, in bytes, of a VARIANT structure.
        /// </summary>
        private const int VARIANT_SIZE = 0x10;

        /// <summary>
        /// Frees all memory referenced by a VARIANT stored in unmanaged memory.
        /// </summary>
        [DllImport("oleaut32.dll")]
        public static extern void VariantClear(IntPtr pVariant);

        /// <summary>
        /// Unmarshals an array of VARIANTs as objects.
        /// </summary>
        public static object[] GetVARIANTs(ref IntPtr pArray, int size, bool deallocate)
        {
            // this method unmarshals VARIANTs one at a time because a single bad value throws 
            // an exception with GetObjectsForNativeVariants(). This approach simply sets the 
            // offending value to null.

            if (pArray == IntPtr.Zero || size <= 0)
            {
                return null;
            }

            object[] values = new object[size];

            IntPtr pos = pArray;

            for (int ii = 0; ii < size; ii++)
            {
                try
                {
                    values[ii] = Marshal.GetObjectForNativeVariant(pos);
                    if (deallocate) VariantClear(pos);
                }
                catch (Exception)
                {
                    values[ii] = null;
                }

                pos = (IntPtr)(pos.ToInt32() + VARIANT_SIZE);
            }

            if (deallocate)
            {
                Marshal.FreeCoTaskMem(pArray);
                pArray = IntPtr.Zero;
            }

            return values;
        }

        /// <summary>
        /// The constant used to selected the default locale.
        /// </summary>
        internal const int LOCALE_SYSTEM_DEFAULT = 0x800;
        internal const int LOCALE_USER_DEFAULT = 0x400; 

        /// <summary>
        /// Converts a LCID to a Locale string.
        /// </summary>
        internal static string GetLocale(int input)
        {
            try
            {
                if (input == LOCALE_SYSTEM_DEFAULT || input == LOCALE_USER_DEFAULT || input == 0)
                {
                    return CultureInfo.InvariantCulture.Name;
                }

                return new CultureInfo(input).Name;
            }
            catch
            {
                throw new Exception("Invalid LCID");
            }
        }

        /// <summary>
        /// Converts a Locale string to a LCID.
        /// </summary>
        internal static int GetLocale(string input)
        {
            // check for the default culture.
            if (input == null || input == "")
            {
                return 0;
            }

            CultureInfo locale = null;

            try { locale = new CultureInfo(input); }
            catch { locale = CultureInfo.CurrentCulture; }

            return locale.LCID;
        }
        #endregion

        #region COM Server Handling
        /// <summary>
        /// Creates an instance of a COM server.
        /// </summary>
        public static object CreateInstance(Guid clsid, string hostName, NetworkCredential credential)
        {
            ServerInfo serverInfo = new ServerInfo();
            COSERVERINFO coserverInfo = serverInfo.Allocate(hostName, null);

            GCHandle hIID = GCHandle.Alloc(IID_IUnknown, GCHandleType.Pinned);

            MULTI_QI[] results = new MULTI_QI[1];

            results[0].iid = hIID.AddrOfPinnedObject();
            results[0].pItf = null;
            results[0].hr = 0;

            try
            {
                // check whether connecting locally or remotely.
                uint clsctx = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER;

                if (hostName != null && hostName.Length > 0 && hostName != "localhost")
                {
                    clsctx = CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER;
                }

                // create an instance.
                CoCreateInstanceEx(
                    ref clsid,
                    null,
                    clsctx,
                    ref coserverInfo,
                    1,
                    results);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (hIID.IsAllocated) hIID.Free();
                serverInfo.Deallocate();
            }

            if (results[0].hr != 0)
            {
                throw new ExternalException("CoCreateInstanceEx: " + GetSystemMessage((int)results[0].hr));
            }

            return results[0].pItf;
        }

        /// <summary>
        /// Releases the server if it is a true COM server.
        /// </summary>
        public static void ReleaseServer(object server)
        {
            if (server != null && server.GetType().IsCOMObject)
            {
                Marshal.ReleaseComObject(server);
            }
        }

        /// <summary>
        /// Parse the URL of an OPC Server.
        /// </summary>
        public static void ParseUrl(string url, out Guid clsid, out string hostName)
        {
            hostName = "localhost";

            string scheme = string.Empty;
            string progID = string.Empty;
            string sClsid = null;
            string strDummy = string.Empty;

            string buffer = url;

            // extract the scheme (default is http).
            int index = buffer.IndexOf("://");

            if (index >= 0)
            {
                scheme = buffer.Substring(0, index);
                buffer = buffer.Substring(index + 3);
            }

            // extract the hostname (default is localhost).
            index = buffer.IndexOfAny(new char[] { ':', '/' });

            if (index < 0)
            {
                throw new ApplicationException("Invalid url");
            }
            else
            {
                hostName = buffer.Substring(0, index);
                buffer = buffer.Substring(index + 1);
            }

            // extract the progID (default is localhost).
            index = buffer.IndexOf('/');

            // look up prog id if sClsid not specified in the url.
            if (index < 0)
            {
                progID = buffer;
            }
            else
            {
                progID = buffer.Substring(0, index);
                sClsid = buffer.Substring(index + 1);
            }
            if (sClsid == null)
            {
                System.Type type = System.Type.GetTypeFromProgID(progID, hostName, true);

                // try converting prog id to a guid if type is not found.
                if (type == null)
                {
                    clsid = new Guid(progID);
                }

                    // fetch the guid of the system type.
                else
                {
                    clsid = type.GUID;
                }
            }

            // convert clsid string to a guid.
            else
            {
                clsid = new Guid(sClsid);
            }
        }
        #endregion

        #region OLE32 Function/Interface Declarations
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SOLE_AUTHENTICATION_SERVICE
        {
            public uint dwAuthnSvc;
            public uint dwAuthzSvc;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pPrincipalName;
            public int hr;
        }

        private const uint RPC_C_AUTHN_NONE = 0;
        private const uint RPC_C_AUTHN_DCE_PRIVATE = 1;
        private const uint RPC_C_AUTHN_DCE_PUBLIC = 2;
        private const uint RPC_C_AUTHN_DEC_PUBLIC = 4;
        private const uint RPC_C_AUTHN_GSS_NEGOTIATE = 9;
        private const uint RPC_C_AUTHN_WINNT = 10;
        private const uint RPC_C_AUTHN_GSS_SCHANNEL = 14;
        private const uint RPC_C_AUTHN_GSS_KERBEROS = 16;
        private const uint RPC_C_AUTHN_DPA = 17;
        private const uint RPC_C_AUTHN_MSN = 18;
        private const uint RPC_C_AUTHN_DIGEST = 21;
        private const uint RPC_C_AUTHN_MQ = 100;
        private const uint RPC_C_AUTHN_DEFAULT = 0xFFFFFFFF;

        private const uint RPC_C_AUTHZ_NONE = 0;
        private const uint RPC_C_AUTHZ_NAME = 1;
        private const uint RPC_C_AUTHZ_DCE = 2;
        private const uint RPC_C_AUTHZ_DEFAULT = 0xffffffff;

        private const uint RPC_C_AUTHN_LEVEL_DEFAULT = 0;
        private const uint RPC_C_AUTHN_LEVEL_NONE = 1;
        private const uint RPC_C_AUTHN_LEVEL_CONNECT = 2;
        private const uint RPC_C_AUTHN_LEVEL_CALL = 3;
        private const uint RPC_C_AUTHN_LEVEL_PKT = 4;
        private const uint RPC_C_AUTHN_LEVEL_PKT_INTEGRITY = 5;
        private const uint RPC_C_AUTHN_LEVEL_PKT_PRIVACY = 6;

        private const uint RPC_C_IMP_LEVEL_ANONYMOUS = 1;
        private const uint RPC_C_IMP_LEVEL_IDENTIFY = 2;
        private const uint RPC_C_IMP_LEVEL_IMPERSONATE = 3;
        private const uint RPC_C_IMP_LEVEL_DELEGATE = 4;

        private const uint EOAC_NONE = 0x00;
        private const uint EOAC_MUTUAL_AUTH = 0x01;
        private const uint EOAC_CLOAKING = 0x10;
        private const uint EOAC_SECURE_REFS = 0x02;
        private const uint EOAC_ACCESS_CONTROL = 0x04;
        private const uint EOAC_APPID = 0x08;

        [DllImport("ole32.dll")]
        private static extern int CoInitializeSecurity(
            IntPtr pSecDesc,
            int cAuthSvc,
            SOLE_AUTHENTICATION_SERVICE[] asAuthSvc,
            IntPtr pReserved1,
            uint dwAuthnLevel,
            uint dwImpLevel,
            IntPtr pAuthList,
            uint dwCapabilities,
            IntPtr pReserved3);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct COSERVERINFO
        {
            public uint dwReserved1;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pwszName;
            public IntPtr pAuthInfo;
            public uint dwReserved2;
        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct COAUTHINFO
        {
            public uint dwAuthnSvc;
            public uint dwAuthzSvc;
            public IntPtr pwszServerPrincName;
            public uint dwAuthnLevel;
            public uint dwImpersonationLevel;
            public IntPtr pAuthIdentityData;
            public uint dwCapabilities;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct COAUTHIDENTITY
        {
            public IntPtr User;
            public uint UserLength;
            public IntPtr Domain;
            public uint DomainLength;
            public IntPtr Password;
            public uint PasswordLength;
            public uint Flags;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct MULTI_QI
        {
            public IntPtr iid;
            [MarshalAs(UnmanagedType.IUnknown)]
            public object pItf;
            public uint hr;
        }

        private const uint CLSCTX_INPROC_SERVER = 0x1;
        private const uint CLSCTX_INPROC_HANDLER = 0x2;
        private const uint CLSCTX_LOCAL_SERVER = 0x4;
        private const uint CLSCTX_REMOTE_SERVER = 0x10;

        private static readonly Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        private const uint SEC_WINNT_AUTH_IDENTITY_ANSI = 0x1;
        private const uint SEC_WINNT_AUTH_IDENTITY_UNICODE = 0x2;

        [DllImport("ole32.dll")]
        private static extern void CoCreateInstanceEx(
            ref Guid clsid,
            [MarshalAs(UnmanagedType.IUnknown)]
            object punkOuter,
            uint dwClsCtx,
            [In]
            ref COSERVERINFO pServerInfo,
            uint dwCount,
            [In, Out]
            MULTI_QI[] pResults);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct LICINFO
        {
            public int cbLicInfo;
            [MarshalAs(UnmanagedType.Bool)]
            public bool fRuntimeKeyAvail;
            [MarshalAs(UnmanagedType.Bool)]
            public bool fLicVerified;
        }

        [ComImport]
        [GuidAttribute("B196B28F-BAB4-101A-B69C-00AA00341D07")]
        [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IClassFactory2
        {
            void CreateInstance(
                [MarshalAs(UnmanagedType.IUnknown)]
                object punkOuter,
                [MarshalAs(UnmanagedType.LPStruct)] 
                Guid riid,
                [MarshalAs(UnmanagedType.Interface)]
                [Out] out object ppvObject);

            void LockServer(
                [MarshalAs(UnmanagedType.Bool)]
                bool fLock);

            void GetLicInfo(
                [In, Out] ref LICINFO pLicInfo);

            void RequestLicKey(
                int dwReserved,
                [MarshalAs(UnmanagedType.BStr)]
                string pbstrKey);

            void CreateInstanceLic(
                [MarshalAs(UnmanagedType.IUnknown)]
                object punkOuter,
                [MarshalAs(UnmanagedType.IUnknown)]
                object punkReserved,
                [MarshalAs(UnmanagedType.LPStruct)]  
                Guid riid,
                [MarshalAs(UnmanagedType.BStr)]
                string bstrKey,
                [MarshalAs(UnmanagedType.IUnknown)]
                [Out] out object ppvObject);
        }

        [ComImport]
        [GuidAttribute("0000013D-0000-0000-C000-000000000046")]
        [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IClientSecurity
        {
            void QueryBlanket(
                [MarshalAs(UnmanagedType.IUnknown)]
                object pProxy,
                ref uint pAuthnSvc,
                ref uint pAuthzSvc,
                [MarshalAs(UnmanagedType.LPWStr)]
                ref string pServerPrincName,
                ref uint pAuthnLevel,
                ref uint pImpLevel,
                ref IntPtr pAuthInfo,
                ref uint pCapabilities);

            void SetBlanket(
                [MarshalAs(UnmanagedType.IUnknown)]
                object pProxy,
                uint pAuthnSvc,
                uint pAuthzSvc,
                [MarshalAs(UnmanagedType.LPWStr)]
                string pServerPrincName,
                uint pAuthnLevel,
                uint pImpLevel,
                IntPtr pAuthInfo,
                uint pCapabilities);

            void CopyProxy(
                [MarshalAs(UnmanagedType.IUnknown)]
                object pProxy,
                [MarshalAs(UnmanagedType.IUnknown)]
                [Out] out object ppCopy);
        }

        [DllImport("ole32.dll")]
        private static extern void CoGetClassObject(
            [MarshalAs(UnmanagedType.LPStruct)] 
            Guid clsid,
            uint dwClsContext,
            [In] ref COSERVERINFO pServerInfo,
            [MarshalAs(UnmanagedType.LPStruct)] 
            Guid riid,
            [MarshalAs(UnmanagedType.IUnknown)]
            [Out] out object ppv);

        private const int MAX_MESSAGE_LENGTH = 1024;

        private const uint FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200;
        private const uint FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;

        [DllImport("Kernel32.dll")]
        private static extern int FormatMessageW(
            int dwFlags,
            IntPtr lpSource,
            int dwMessageId,
            int dwLanguageId,
            IntPtr lpBuffer,
            int nSize,
            IntPtr Arguments);

        /// <summary>
        /// Retrieves the system message text for the specified error.
        /// </summary>
        public static string GetSystemMessage(int error)
        {
            IntPtr buffer = Marshal.AllocCoTaskMem(MAX_MESSAGE_LENGTH);

            int result = FormatMessageW(
                (int)(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_FROM_SYSTEM),
                IntPtr.Zero,
                error,
                0,
                buffer,
                MAX_MESSAGE_LENGTH - 1,
                IntPtr.Zero);

            string msg = Marshal.PtrToStringUni(buffer);
            Marshal.FreeCoTaskMem(buffer);

            if (msg != null && msg.Length > 0)
            {
                return msg;
            }

            return String.Format("0x{0,0:X}", error);
        }
        #endregion

        #region ServerInfo Class
        /// <summary>
        /// A class used to allocate and deallocate the elements of a COSERVERINFO structure.
        /// </summary>
        class ServerInfo
        {
            #region Public Interface
            /// <summary>
            /// Allocates a COSERVERINFO structure.
            /// </summary>
            public COSERVERINFO Allocate(string hostName, NetworkCredential credential)
            {
                string userName = null;
                string password = null;
                string domain = null;

                if (credential != null)
                {
                    userName = credential.UserName;
                    password = credential.Password;
                    domain = credential.Domain;
                }

                m_hUserName = GCHandle.Alloc(userName, GCHandleType.Pinned);
                m_hPassword = GCHandle.Alloc(password, GCHandleType.Pinned);
                m_hDomain = GCHandle.Alloc(domain, GCHandleType.Pinned);

                m_hIdentity = new GCHandle();

                if (userName != null && userName != String.Empty)
                {
                    COAUTHIDENTITY identity = new COAUTHIDENTITY();

                    identity.User = m_hUserName.AddrOfPinnedObject();
                    identity.UserLength = (uint)((userName != null) ? userName.Length : 0);
                    identity.Password = m_hPassword.AddrOfPinnedObject();
                    identity.PasswordLength = (uint)((password != null) ? password.Length : 0);
                    identity.Domain = m_hDomain.AddrOfPinnedObject();
                    identity.DomainLength = (uint)((domain != null) ? domain.Length : 0);
                    identity.Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;

                    m_hIdentity = GCHandle.Alloc(identity, GCHandleType.Pinned);
                }

                COAUTHINFO authInfo = new COAUTHINFO();
                authInfo.dwAuthnSvc = RPC_C_AUTHN_WINNT;
                authInfo.dwAuthzSvc = RPC_C_AUTHZ_NONE;
                authInfo.pwszServerPrincName = IntPtr.Zero;
                authInfo.dwAuthnLevel = RPC_C_AUTHN_LEVEL_CONNECT;
                authInfo.dwImpersonationLevel = RPC_C_IMP_LEVEL_IMPERSONATE;
                authInfo.pAuthIdentityData = (m_hIdentity.IsAllocated) ? m_hIdentity.AddrOfPinnedObject() : IntPtr.Zero;
                authInfo.dwCapabilities = EOAC_NONE;

                m_hAuthInfo = GCHandle.Alloc(authInfo, GCHandleType.Pinned);

                COSERVERINFO serverInfo = new COSERVERINFO();

                serverInfo.pwszName = hostName;
                serverInfo.pAuthInfo = (credential != null) ? m_hAuthInfo.AddrOfPinnedObject() : IntPtr.Zero;
                serverInfo.dwReserved1 = 0;
                serverInfo.dwReserved2 = 0;

                return serverInfo;
            }

            /// <summary>
            /// Deallocated memory allocated when the COSERVERINFO structure was created.
            /// </summary>
            public void Deallocate()
            {
                if (m_hUserName.IsAllocated) m_hUserName.Free();
                if (m_hPassword.IsAllocated) m_hPassword.Free();
                if (m_hDomain.IsAllocated) m_hDomain.Free();
                if (m_hIdentity.IsAllocated) m_hIdentity.Free();
                if (m_hAuthInfo.IsAllocated) m_hAuthInfo.Free();
            }
            #endregion

            #region Private Members
            GCHandle m_hUserName;
            GCHandle m_hPassword;
            GCHandle m_hDomain;
            GCHandle m_hIdentity;
            GCHandle m_hAuthInfo;
            #endregion
        }
        #endregion

        #region Copy COM arrays
        /// <summary>
        /// Unmarshals and frees an array of 32 bit integers.
        /// </summary>
        public static int[] GetInt32s(ref IntPtr pArray, int size, bool deallocate)
        {
            if (pArray == IntPtr.Zero || size <= 0)
            {
                return null;
            }

            int[] array = new int[size];
            Marshal.Copy(pArray, array, 0, size);

            if (deallocate)
            {
                Marshal.FreeCoTaskMem(pArray);
                pArray = IntPtr.Zero;
            }

            return array;
        }

        /// <summary>
        /// Unmarshals and frees a array of unicode strings.
        /// </summary>
        public static string[] GetUnicodeStrings(ref IntPtr pArray, int size, bool deallocate)
        {
            if (pArray == IntPtr.Zero || size <= 0)
            {
                return null;
            }

            int[] pointers = new int[size];
            Marshal.Copy(pArray, pointers, 0, size);

            string[] strings = new string[size];

            for (int ii = 0; ii < size; ii++)
            {
                IntPtr pString = (IntPtr)pointers[ii];
                strings[ii] = Marshal.PtrToStringUni(pString);
                if (deallocate) Marshal.FreeCoTaskMem(pString);
            }

            if (deallocate)
            {
                Marshal.FreeCoTaskMem(pArray);
                pArray = IntPtr.Zero;
            }

            return strings;
        }

        /// <summary>
        /// Unmarshals and frees a array of 16 bit integers.
        /// </summary>
        public static short[] GetInt16s(ref IntPtr pArray, int size, bool deallocate)
        {
            if (pArray == IntPtr.Zero || size <= 0)
            {
                return null;
            }

            short[] array = new short[size];
            Marshal.Copy(pArray, array, 0, size);

            if (deallocate)
            {
                Marshal.FreeCoTaskMem(pArray);
                pArray = IntPtr.Zero;
            }

            return array;
        }
        #endregion
    }

    #region ConnectionPoint
    /// <summary>
    /// Adds and removes a connection point to a server.
    /// </summary>
    public class ConnectionPoint : IDisposable
    {
        /// <summary>
        /// The COM server that supports connection points.
        /// </summary>
        private IConnectionPoint m_server = null;

        /// <summary>
        /// The id assigned to the connection by the COM server.
        /// </summary>
        private int m_cookie = 0;

        /// <summary>
        /// The number of times Advise() has been called without a matching Unadvise(). 
        /// </summary>
        private int m_refs = 0;

        /// <summary>
        /// Initializes the object by finding the specified connection point.
        /// </summary>
        public ConnectionPoint(object server, Guid iid)
        {
            ((IConnectionPointContainer)server).FindConnectionPoint(ref iid, out m_server);
        }

        /// <summary>
        /// Releases the COM server.
        /// </summary>
        public void Dispose()
        {
            if (m_server != null)
            {
                while (Unadvise() > 0) ;
                Helper.ReleaseServer(m_server);
                m_server = null;
            }
        }

        //=====================================================================
        // IConnectionPoint

        /// <summary>
        /// Establishes a connection, if necessary and increments the reference count.
        /// </summary>
        public int Advise(object callback)
        {
            if (m_refs++ == 0) m_server.Advise(callback, out m_cookie);
            return m_refs;
        }

        /// <summary>
        /// Decrements the reference count and closes the connection if no more references.
        /// </summary>
        public int Unadvise()
        {
            if (--m_refs == 0) m_server.Unadvise(m_cookie);
            return m_refs;
        }
    }
    #endregion
}
