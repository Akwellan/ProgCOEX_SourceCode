using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Siemens.Opc;
using OpcRcw.Da;
using OpcRcw.Comn;

namespace Siemens.Opc.Da
{
    public delegate void DataChange(IList<DataValue> DataValues);

    public class Subscription : IDisposable
    {
        #region Static members
        static int s_clientGroupHandle = 0;
        #endregion

        #region Construction
        internal Subscription(Server server)
        {
            m_server = server;
            m_mapServerHandles = new Dictionary<string,int>();
        }
        #endregion

        #region Create
        internal void Create(string groupName, DataChange fct)
        {
            if (!m_server.isConnected())
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }
            Guid iid = Guid.Empty;
            iid = typeof(IOPCGroupStateMgt).GUID;
            if (iid == Guid.Empty)
            {
                throw new Exception("Creating iid failed");
            }
            m_hClient = s_clientGroupHandle++;
            m_server.getIOPCServer().AddGroup(
                groupName,          //GroupName
                1,                  //active
                0,                  //RequestedUpdateRate
                m_hClient,          //GroupHandle
                IntPtr.Zero,        //TimeBias
                IntPtr.Zero,        //PercentDeadband
                0,                  //LCID
                out m_hServer,
                out m_revisedUpdateRate,
                ref iid,
                out m_group);
            if (m_group == null)
            {
                throw new Exception("Creating default group failed");
            }
            m_itemManagement = (OpcRcw.Da.IOPCItemMgt)m_group;
            if (m_itemManagement == null)
            {
                throw new Exception("Creating ItemManagement failed");
            }
            m_subscriptionPrivate = new SubscriptionPrivate(fct, m_hClient, m_group);
        }
        #endregion

        #region Delete
        internal void Delete()
        {
            m_server.getIOPCServer().RemoveGroup(m_hServer, 1);
            m_mapServerHandles.Clear();
        }
        #endregion

        #region AddItems
        public void AddItem(string itemId, int clientHandle)
        {
            if (!m_server.isConnected())
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }
            OPCITEMDEF[] ItemDefs = new OPCITEMDEF[1];
            ItemDefs[0] = new OPCITEMDEF()
            {
                bActive = 1,
                szItemID = itemId,
                hClient = clientHandle
            };
            IntPtr pAddResults;
            IntPtr pErrors;
            m_itemManagement.AddItems(1, ItemDefs, out pAddResults, out pErrors);
            if (pAddResults == IntPtr.Zero)
            {
                throw new Exception("AddItems failed. No Results.");
            }
            if (pErrors == IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(pAddResults);
                throw new Exception("AddItems failed. No ErrorCodes.");
            }
            int[] errors = new int[1];
            Marshal.Copy(pErrors, errors, 0, 1);
            if (errors[0] != 0)
            {
                Marshal.FreeCoTaskMem(pAddResults);
                Marshal.FreeCoTaskMem(pErrors);
                string ErrorString;
                m_server.GetErrorString(errors[0], 0, out ErrorString);
                throw new Exception(ErrorString);
            }
            OPCITEMRESULT result = (OPCITEMRESULT)Marshal.PtrToStructure(pAddResults, typeof(OPCITEMRESULT));
            int hServer = result.hServer;
            m_mapServerHandles.Add(itemId, hServer);
            Marshal.FreeCoTaskMem(pAddResults);
            Marshal.FreeCoTaskMem(pErrors);
        }
        #endregion

        #region RemoveItems
        public void RemoveItem(string ItemId)
        {
            if (!m_server.isConnected())
            {
                throw new InvalidOperationException("Not connected to OPC-Server!");
            }
            int[] ServerHandles = new int[1];
            if (! m_mapServerHandles.TryGetValue(ItemId, out ServerHandles[0]))
            {
                throw new Exception("Error in RemoveItem. Item " + ItemId + " is not monitored.");
            }
            IntPtr pErrors;
            m_itemManagement.RemoveItems(1, ServerHandles, out pErrors);
            if (pErrors == IntPtr.Zero)
            {
                throw new Exception("RemoveItem failed. No ErrorCodes.");
            }
            int[] errors = new int[1];
            Marshal.Copy(pErrors, errors, 0, 1);
            if (errors[0] != 0)
            {
                Marshal.FreeCoTaskMem(pErrors);
                throw new Exception("RemoveItem failed. Error occured");
            }
            m_mapServerHandles.Remove(ItemId);
        }
        #endregion

        #region Properties
        public int UpdateRate
        {
            get { return m_revisedUpdateRate; }
        }
        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Helper
        #endregion

        #region Fields
        Server m_server;
        int m_hServer; // Server group handle
        int m_hClient; // Client group handle
        int m_revisedUpdateRate;
        object m_group = null;
        OpcRcw.Da.IOPCItemMgt m_itemManagement = null;
        Dictionary<string, int> m_mapServerHandles;
        SubscriptionPrivate m_subscriptionPrivate = null;
        #endregion
    }
}
