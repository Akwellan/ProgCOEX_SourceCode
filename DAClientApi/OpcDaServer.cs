using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using OpcRcw.Da;
using OpcRcw.Comn;
using Siemens.Opc;

namespace Siemens.Opc.Da
{
    public delegate void ServerShutdownEventHandler(string reason);

    public delegate void ReadHandler(IList<ReadResult> ReadResults);

    /// <summary>
    /// This class encapsulates a connection to an OPC Alarm & Events server.
    /// </summary>
    public class Server: IDisposable
    {
        readonly int ReadWriteGroup = 0;

        #region Properties

        /// <summary>
        /// Indicates if the server is currently connected
        /// </summary>
        public bool IsConnected
        {
            get{ return m_connected; }
        }
        /// <summary>
        /// Provides the URL of the currently connected OPC server
        /// </summary>
        public string Url
        {
            get { return m_Url; }
        }

        #endregion

        #region OPC-Connect

        /// <summary>Establishes the connection to an OPC DA server.</summary>
        /// <param name="url"> The Url of the OPC Server </param>
        /// <exception cref="InvalidOperationException"> if the opcserver is already connected </exception>
        /// <exception cref="Exception">throws and forwards any exception (with short error description)</exception>
        public void Connect(string url)
        {
            lock (this)
            {
                if (!m_connected)
                {
                    try
                    {
                        Guid clsid;
                        string hostName;
                        // Parse the URL for the OPC DA Server and get the host name and clsid of the server
                        // URL has the syntax opcae://localhost/OPCSample.OPCEventServer/{65168852-5783-11d1-84a0-00608cb8a7e9}
                        Helper.ParseUrl(url, out clsid, out hostName);

                        // Establish connection to the OPC Server
                        m_server = (OpcRcw.Da.IOPCServer ) Helper.CreateInstance(clsid, hostName, null);
                        if (m_server == null)
                        {
                            throw new Exception("CreateInstance failed");
                        }

                        // get interfaces to browse
                        try
                        {
                            m_browserDA3 = (OpcRcw.Da.IOPCBrowse)m_server;
                        }
                        catch (Exception)
                        {
                            m_browserDA2 = (OpcRcw.Da.IOPCBrowseServerAddressSpace)m_server;
                        }

                        int serverGroupHandle;
                        int RevisedUpdateRate;
                        Guid iid = Guid.Empty;
                        iid = typeof(IOPCGroupStateMgt).GUID;
                        if (iid == Guid.Empty)
                        {
                            throw new Exception("Creating iid failed");
                        }

                        object objGroup = null;
                        m_server.AddGroup(
                            "ReadWriteDefault", //name
                            0,                  //active
                            0,                  //requestedUpdateRate
                            ReadWriteGroup,     //ClientGroupHandle
                            IntPtr.Zero,        //TimeBias
                            IntPtr.Zero,        //PercentDeadband
                            0,                  //LCID
                            out serverGroupHandle,
                            out RevisedUpdateRate,
                            ref iid,
                            out objGroup);
                        if (objGroup == null)
                        {
                            throw new Exception("Creating default group failed");
                        }

                        m_groupStateManagement = (OpcRcw.Da.IOPCGroupStateMgt)objGroup;
                        if (m_groupStateManagement == null)
                        {
                            throw new Exception("Creating IOPCGroupStateMgt2 failed");
                        }

                        m_itemManagement = (OpcRcw.Da.IOPCItemMgt)m_groupStateManagement;
                        if (m_itemManagement == null)
                        {
                            throw new Exception("Creating ItemManagement failed");
                        }

                        m_syncIO = (OpcRcw.Da.IOPCSyncIO)objGroup;
                        if (m_syncIO == null)
                        {
                            throw new Exception("Creating IOPCSyncIO failed");
                        }

                        m_asyncIO2 = (OpcRcw.Da.IOPCAsyncIO2)objGroup;
                        if (m_asyncIO2 == null)
                        {
                            throw new Exception("Creating IOPCASyncIO2 failed");
                        }

                        m_serverPrivate = new ServerPrivate(objGroup);

                        m_connected = true;
                        m_Url = url;
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
                else
                {
                    throw new InvalidOperationException("Already connected to OPC-Server!");
                }
            }
        }

        /// <summary>Release the connection to an OPC Alarm & Events server.</summary>
        /// <exception cref="InvalidOperationException"> if the opcserver is not connected </exception>
        /// <exception cref="Exception">throws and forwards any exception</exception>
        public void Disconnect()
        {
            lock (this)
            {
                if (m_connected)
                {
                    m_mapServerHandles.Clear();
                    try
                    {
                        // close callback connections.
                        // close connections.
                        if (m_connection != null)
                        {
                            m_connection.Dispose();
                            m_connection = null;
                        }

                        Helper.ReleaseServer(m_server);
                        m_server = null;
                        m_connected = false;
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
                else
                {
                    throw new InvalidOperationException("Not connected to OPC-Server!");
                }
            }
        }

        #endregion

        #region GetErrorString

        public void GetErrorString(int error, int locale, out string errorString)
        {
            if (!m_connected)
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }
            m_server.GetErrorString(error, locale, out errorString);
        }

        #endregion

        #region SetAsynIOInterface

        public void SetAsyncIOInterface(ReadHandler EndRead)
        {
            m_serverPrivate.EndRead = EndRead;
        }

        #endregion

        #region Read

        /// <summary>
        /// Read a single item
        /// </summary>
        /// <param name="itemId"></param>
        /// <returns>the item value or null is reading failed</returns>
        public object Read(string itemId)
        {
            if (!m_connected)
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }

            object result = null;

            StringCollection itemIds = new StringCollection();
            object[] values;
            int[] pErrors;

            itemIds.Add(itemId);

            Read(itemIds, out values, out pErrors);

            if (pErrors[0] == 0)
            {
                result = values[0];
            }
            else
            {
                throw new Exception("Read failed");
            }

            return result;
        }

        /// <summary>
        /// Read the list of items.
        /// </summary>
        /// <param name="itemIds">Items to read.</param>
        /// <param name="values">The values that were read.</param>
        /// <param name="pErrors">Individual error code for each item</param>
        /// <returns>Returns true if read succeeded for all items. Otherwise returns false.</returns>
        public bool Read(StringCollection itemIds, out object[] values, out int[] pErrors)
        {
            bool bResult = true;
            values = new object[itemIds.Count];
            pErrors = new int[itemIds.Count];

            if (!m_connected)
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }

            int[] serverHandles;

            try
            {
                bResult = GetServerHandles(itemIds, out serverHandles, out pErrors);
            }
            catch (Exception e)
            {
                throw e;
            }

            // build in parameters to read
            int[] serverHandlesToRead;
            ArrayList indexesToRead = new ArrayList();
            ArrayList indexesNotToRead = new ArrayList();
            ArrayList handlesTmp = new ArrayList();


            // check each handle
            for (int i = 0; i < itemIds.Count; i++)
            {
                if (pErrors[i] == 0)
                {
                    handlesTmp.Add(serverHandles[i]);
                    indexesToRead.Add(i);
                }
                else
                {
                    indexesNotToRead.Add(i);
                }
            }
            serverHandlesToRead = (int[])handlesTmp.ToArray(typeof(int));

            if (serverHandlesToRead.Length == 0)
            {
                return false;
            }

            IntPtr pItemValues;
            IntPtr ppErrors;

            try
            {
                m_syncIO.Read(
                    OPCDATASOURCE.OPC_DS_DEVICE,    // DataSource is Device or Cache
                    serverHandlesToRead.Length,     // Number of items to read
                    serverHandlesToRead,            //
                    out pItemValues,                //
                    out ppErrors);                  //
            }
            catch (Exception e)
            {
                throw e;
            }

            if (pItemValues == IntPtr.Zero)
            {
                throw new Exception("Read failed. No Results.");
            }
            if (ppErrors == IntPtr.Zero)
            {
                throw new Exception("Read failed. No ErrorCodes returned.");
            }

            int[] errorsTmp = Helper.GetInt32s(ref ppErrors, serverHandlesToRead.Length, true);

            int indexTmp = 0;
            IntPtr pos = pItemValues;
            // Fill result of items that were read
            foreach(int index in indexesToRead)
            {
                // set error code
                pErrors[index] = errorsTmp[indexTmp];

                OPCITEMSTATE itemState = (OpcRcw.Da.OPCITEMSTATE)Marshal.PtrToStructure(pos, typeof(OPCITEMSTATE));
                Marshal.DestroyStructure(pos, typeof(OpcRcw.Da.OPCITEMSTATE));
                pos = (IntPtr)(pos.ToInt64() + Marshal.SizeOf(typeof(OPCITEMSTATE)));

                // set value
                values[index] = itemState.vDataValue;

                indexTmp++;
            }

            Marshal.FreeCoTaskMem(pItemValues);
            pItemValues = IntPtr.Zero;

            // Fill results of items that were not read.
            foreach (int index in indexesNotToRead)
            {
                // error code was alread set above - just initialize value element
                values[index] = null;
            }

            return bResult;
        }

        public int ReadAsync(string ItemId, int ClientHandle)
        {
            if (!m_connected)
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }
            if (m_serverPrivate.EndRead == null)
            {
                throw new InvalidOperationException("EndRead is null");
            }
            int[] ServerHandles = new int[1];
            ServerHandles[0] = GetServerHandle(ItemId);
            IntPtr pErrors;
            
            int cancelId;
            m_asyncIO2.Read(1, ServerHandles, ClientHandle, out cancelId, out pErrors);
            if (pErrors == IntPtr.Zero)
            {
                throw new Exception("Read failed. No ErrorCodes returned.");
            }
            int[] errors = new int[1];
            Marshal.Copy(pErrors, errors, 0, 1);
            if (errors[0] != 0)
            {
                string ErrorString;
                m_server.GetErrorString(errors[0], 0, out ErrorString);
                throw new Exception(ErrorString);
            }
            return cancelId;
        }

        #endregion

        #region Write

        public void Write(string ItemId, object value)
        {
            if (!m_connected)
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }
            int[] ServerHandles = new int[1];
            ServerHandles[0] = GetServerHandle(ItemId);
            object[] Values = new object[1];
            Values[0] = value;
            IntPtr pErrors;

            try
            {
                m_syncIO.Write(1, ServerHandles, Values, out pErrors);
            }
            catch (Exception)
            {
                throw;
            }
            if (pErrors == IntPtr.Zero)
            {
                throw new Exception("Write failed. No ErrorCodes returned.");
            }
            int[] errors = new int[1];
            Marshal.Copy(pErrors, errors, 0, 1);
            if (errors[0] != 0)
            {
                string ErrorString;
                m_server.GetErrorString(errors[0], 0, out ErrorString);
                throw new Exception(ErrorString);
            }
        }

        #endregion

        #region Browse

        public BrowseResult[] Browse(string ItemId)
        {
            if (!m_connected)
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }

            BrowseResult[] ret = null;
            IntPtr pContinuationPoint = IntPtr.Zero;
            int MaxElementsReturned = 10; // There is no client side restriction on the number of returned elements.
            int MoreElements = 0;
            int BrowseElementCount = 0;
            IntPtr BrowseElements = new IntPtr(1);
            int[] PropertyIds = new int[1];
            try
            {
                m_browserDA3.Browse(
                    ItemId,
                    ref pContinuationPoint,
                    MaxElementsReturned,
                    OPCBROWSEFILTER.OPC_BROWSE_FILTER_ALL,
                    "", // no name filter
                    "", // no vendow filter
                    1, // return all properties
                    1, // return property values
                    0, // Property count, is ignored since all properties are returned
                    PropertyIds, //array of property ids, is ignored since all properties are returned
                    out MoreElements,
                    out BrowseElementCount,
                    out BrowseElements);

                OpcRcw.Da.OPCBROWSEELEMENT[] elements = GetBrowseElements(BrowseElements, BrowseElementCount);

                ret = new BrowseResult[BrowseElementCount];
                int i = 0;
                foreach (OpcRcw.Da.OPCBROWSEELEMENT element in elements)
                {
                    OpcRcw.Da.OPCITEMPROPERTIES ItemProperties = element.ItemProperties;
                    IntPtr Properties = ItemProperties.pItemProperties;
                    int PropertyCount = ItemProperties.dwNumProperties;
                    ret[i++] = new BrowseResult
                    {
                        ItemId = element.szItemID,
                        ItemName = element.szName,
                        ItemProperties = GetItemProperties(Properties, PropertyCount)
                    };
                }

            }
            catch (Exception e)
            {
                throw e;
            }
            return ret;
        }

        #endregion

        #region Subscribe

        public Subscription CreateSubscription(string groupName, DataChange callback)
        {
            if (!m_connected)
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }
            Subscription subscription = new Subscription(this);
            subscription.Create(groupName, callback);
            return subscription;
        }

        public void DeleteSubscription(Subscription subscription)
        {
            if (!m_connected)
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }
            subscription.Delete();
        }

        #endregion

        #region Getters

        public bool isConnected()
        {
            return m_connected;
        }

        public IOPCServer getIOPCServer()
        {
            return m_server;
        }

        #endregion

        #region Helper

        internal static OpcRcw.Da.OPCBROWSEELEMENT[] GetBrowseElements(IntPtr BrowseElements, int BrowseElementCount)
        {
            if (BrowseElements != IntPtr.Zero && BrowseElementCount > 0)
            {
                OpcRcw.Da.OPCBROWSEELEMENT[] ret = new OpcRcw.Da.OPCBROWSEELEMENT[BrowseElementCount];
                IntPtr pos = BrowseElements;
                for (int i = 0; i < BrowseElementCount; i++)
                {
                    OpcRcw.Da.OPCBROWSEELEMENT element = (OpcRcw.Da.OPCBROWSEELEMENT)Marshal.PtrToStructure(pos, typeof(OpcRcw.Da.OPCBROWSEELEMENT));
                    Marshal.StructureToPtr(element, pos, false);
                    pos = (IntPtr)(pos.ToInt32() + Marshal.SizeOf(typeof(OpcRcw.Da.OPCBROWSEELEMENT)));
                    ret[i] = element;
                }
                return ret;
            }
            return new OpcRcw.Da.OPCBROWSEELEMENT[0];
        }

        internal static ItemProperty[] GetItemProperties(IntPtr ItemProperties, int ItemPropertyCount)
        {
            ItemProperty[] result = null;
            if (ItemPropertyCount > 0 && ItemProperties != IntPtr.Zero)
            {
                result = new ItemProperty[ItemPropertyCount];
                IntPtr pos = ItemProperties;
                for (int i = 0; i < ItemPropertyCount; i++)
                {
                    ItemProperty itemProperty = new ItemProperty();
                    OpcRcw.Da.OPCITEMPROPERTY property = (OpcRcw.Da.OPCITEMPROPERTY)Marshal.PtrToStructure(pos, typeof(OpcRcw.Da.OPCITEMPROPERTY));
                    itemProperty.PropertyId = (PropertiyId)property.dwPropertyID;
                    int ErrorId = property.hrErrorID;
                    itemProperty.Description = property.szDescription;
                    itemProperty.ItemId = property.szItemID;
                    itemProperty.DataType = Helper.GetType((VarEnum)property.vtDataType);
                    itemProperty.Value = ConvertValue(itemProperty.PropertyId, property.vValue);

                    result[i] = itemProperty;

                    pos = (IntPtr)(pos.ToInt32() + Marshal.SizeOf(typeof(OpcRcw.Da.OPCITEMPROPERTY)));
                }
            }
            return result;
        }

        internal static object ConvertValue(PropertiyId propertyId, object Value)
        {
            ///ToDo Other Property Ids
            switch(propertyId)
            {
                case PropertiyId.AccessRights:
                    return (AccessRights)System.Convert.ToInt32(Value);
                case PropertiyId.DataType:
                    return Helper.GetType((VarEnum)System.Convert.ToUInt16(Value));
                default:
                    return Value;
            }
        }

        internal int GetServerHandle(string ItemId)
        {
            int ret;
            if (!m_mapServerHandles.TryGetValue(ItemId, out ret))
            {
                OPCITEMDEF[] ItemIds = new OPCITEMDEF[1];
                ItemIds[0] = new OPCITEMDEF()
                {
                    bActive = 0,
                    szItemID = ItemId
                };
                IntPtr pAddResults;
                IntPtr pErrors;
                m_itemManagement.AddItems(1, ItemIds, out pAddResults, out pErrors);
                if (pAddResults == IntPtr.Zero)
                {
                    throw new Exception("GetServerHandle failed. No Results.");
                }
                if (pErrors == IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pAddResults);
                    throw new Exception("GetServerHandle failed. No ErrorCodes.");
                }
                int[] errors = new int[1];
                Marshal.Copy(pErrors, errors, 0, 1);
                Marshal.FreeCoTaskMem(pErrors);

                if (errors[0] != 0)
                {
                    Marshal.DestroyStructure(pAddResults, typeof(OpcRcw.Da.OPCITEMRESULT));
                    Marshal.FreeCoTaskMem(pAddResults);
                    string ErrorString;
                    m_server.GetErrorString(errors[0], 0, out ErrorString);
                    throw new Exception(ErrorString);
                }

                OPCITEMRESULT result = (OPCITEMRESULT)Marshal.PtrToStructure(pAddResults, typeof(OPCITEMRESULT));
                ret = result.hServer;
                m_mapServerHandles.Add(ItemId, ret);

                Marshal.FreeCoTaskMem(result.pBlob);
                result.pBlob = IntPtr.Zero;
                Marshal.DestroyStructure(pAddResults, typeof(OpcRcw.Da.OPCITEMRESULT));
                Marshal.FreeCoTaskMem(pAddResults);
            }
            return ret;
        }

        /// <summary>
        /// For each item check if is has a handle already. If not add the item.
        /// </summary>
        /// <param name="itemIds"></param>
        /// <param name="handles"></param>
        /// <returns>true if all handles could be retreived. Otherwise false.</returns>
        internal bool GetServerHandles(StringCollection itemIds, out int[] handles, out int[] pErrors)
        {
            bool bResult = true;
            int currentHandle;
            handles = new int[itemIds.Count];
            pErrors = new int[itemIds.Count];

            ArrayList indexesToAdd = new ArrayList();

            // collect data to AddItems
            for (int i = 0; i < itemIds.Count; i++)
            {
                // Item exists already
                if (m_mapServerHandles.TryGetValue(itemIds[i], out currentHandle))
                {
                    handles[i] = currentHandle;
                    pErrors[i] = 0;
                }
                // Need to add item
                else
                {
                    indexesToAdd.Add(i);
                }
            }

            if(indexesToAdd.Count > 0)
            {
                OPCITEMDEF[] itemsToAdd = new OPCITEMDEF[indexesToAdd.Count];

                for (int i = 0; i < indexesToAdd.Count; i++)
                {
                    //itemsToAdd[i] = new OPCITEMDEF();

                    itemsToAdd[i].szItemID              = itemIds[(int)indexesToAdd[i]];
                    itemsToAdd[i].szAccessPath          = "";
                    itemsToAdd[i].bActive               = 0;
                    itemsToAdd[i].vtRequestedDataType   = 0;
                    itemsToAdd[i].hClient               = 0;
                    itemsToAdd[i].dwBlobSize            = 0;
                    itemsToAdd[i].pBlob                 = IntPtr.Zero;
                }

                IntPtr pAddResults;
                IntPtr ppErrors;

                m_itemManagement.AddItems(
                    indexesToAdd.Count,
                    itemsToAdd,
                    out pAddResults,
                    out ppErrors);

                if (pAddResults == IntPtr.Zero)
                {
                    throw new Exception("GetServerHandle failed. No Results.");
                }
                if (ppErrors == IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pAddResults);
                    throw new Exception("GetServerHandle failed. No ErrorCodes.");
                }

                IntPtr posResults = pAddResults;
                int[] pErrorsTmp = Helper.GetInt32s(ref ppErrors, indexesToAdd.Count, true);

                for (int i = 0; i < indexesToAdd.Count; i++)
                {
                    int currentIndex = (int)indexesToAdd[i];
                    if (pErrorsTmp[i] != 0)
                    {
                        // caller to check pErrors
                        bResult = false;
                        pErrors[currentIndex] = pErrorsTmp[i];
                        continue;
                    }

                    OPCITEMRESULT itemResult = (OPCITEMRESULT)Marshal.PtrToStructure(posResults, typeof(OPCITEMRESULT));

                    m_mapServerHandles.Add(itemsToAdd[i].szItemID, itemResult.hServer);
                    handles[currentIndex] = itemResult.hServer;

                    Marshal.FreeCoTaskMem(itemResult.pBlob);
                    itemResult.pBlob = IntPtr.Zero;

                    Marshal.DestroyStructure(posResults, typeof(OpcRcw.Da.OPCITEMRESULT));

                    posResults = (IntPtr)(posResults.ToInt64() + Marshal.SizeOf(typeof(OPCITEMRESULT)));
                }

                Marshal.FreeCoTaskMem(pAddResults);
                pAddResults = IntPtr.Zero;
            }

            return bResult;
        }

        #endregion

        #region IOPCShutdown Implementation

        /// <summary>
        /// A class that implements the IOPCShutdown interface.
        /// </summary>
        private class Callback : OpcRcw.Comn.IOPCShutdown
        {
            /// <summary>
            /// Initializes the object with the containing server object.
            /// </summary>
            public Callback(Server server)
            {
                m_server = server;
            }

            /// <summary>
            /// An event to receive server shutdown notificiations.
            /// </summary>
            public event ServerShutdownEventHandler ServerShutdown
            {
                add { lock (this) { m_serverShutdown += value; } }
                remove { lock (this) { m_serverShutdown -= value; } }
            }

            /// <summary>
            /// The server object handling the server sending shutdown messages.
            /// </summary>
            private Server m_server = null;

            /// <summary>
            /// Raised when data changed callbacks arrive.
            /// </summary>
            private event ServerShutdownEventHandler m_serverShutdown = null;

            /// <summary>
            /// Called when a shutdown event is received.
            /// </summary>
            public void ShutdownRequest(string reason)
            {
                try
                {
                    lock (this)
                    {
                        if (m_serverShutdown != null)
                        {
                            m_serverShutdown(reason);
                        }
                    }
                }
                catch (Exception e)
                {
                    string stack = e.StackTrace;
                }
            }
        }

        #endregion

        #region Fields

        /// <summary> 
        /// Main interface of the OPC Server 
        /// </summary>
        private OpcRcw.Da.IOPCServer m_server;

        // Interfaces to Browse
        private OpcRcw.Da.IOPCBrowse m_browserDA3;
        private OpcRcw.Da.IOPCBrowseServerAddressSpace m_browserDA2;

        // Interfaces to Read and Write
        private OpcRcw.Da.IOPCSyncIO m_syncIO;
        private OpcRcw.Da.IOPCAsyncIO2 m_asyncIO2;

        // Interface to setup Group parameters
        private OpcRcw.Da.IOPCGroupStateMgt m_groupStateManagement;

        // Interface to Add and Remove Items
        private OpcRcw.Da.IOPCItemMgt m_itemManagement;

        private Dictionary<string, int> m_mapServerHandles;

        private ServerPrivate m_serverPrivate;

        /// <summary> 
        /// URL used for the connected OPC server 
        /// </summary>
        private string m_Url = string.Empty;

        /// <summary>
        /// Determines whether this object has already been connected.
        /// this object is only allowed to connect exactly one time! 
        /// </summary>
        bool m_connected = false;

        /// <summary>
        /// A connect point with the COM server.
        /// </summary>
        private ConnectionPoint m_connection = null;

        /// <summary>
        /// The internal object that implements the IOPCShutdown interface.
        /// </summary>
        private Callback m_callback = null;

        #endregion

        #region Construction / Destruction

        /// <summary>
        /// Initializes the server object with the state and the interface of the subscription.
        /// </summary>
        public Server()
        {
            m_callback     = new Callback(this);
            m_mapServerHandles = new Dictionary<string,int>();
        }
        /// <summary>
        /// This must be called explicitly by clients to ensure the COM server is released.
        /// </summary>
        public virtual void Dispose()
        {
            lock (this)
            {
                if (m_connected)
                {
                    Disconnect();
                }
            }
        }

        #endregion

        #region IDisposable Members

        void IDisposable.Dispose()
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
