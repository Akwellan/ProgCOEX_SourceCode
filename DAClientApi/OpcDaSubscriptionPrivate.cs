using System;
using System.Collections.Generic;
using System.Text;
using OpcRcw.Da;

namespace Siemens.Opc.Da
{
    internal class SubscriptionPrivate : IOPCDataCallback
    {
        #region Construction
        public SubscriptionPrivate(DataChange fct, int ClientHandle, object group)
        {
            m_hClient = ClientHandle;
            m_fct = fct;

            Guid iid = Guid.Empty;
            iid = typeof(IOPCDataCallback).GUID;
            OpcRcw.Comn.IConnectionPointContainer ConnectionPointContainer = (OpcRcw.Comn.IConnectionPointContainer)group;
            OpcRcw.Comn.IConnectionPoint ConnectionPoint;
            ConnectionPointContainer.FindConnectionPoint(ref iid, out ConnectionPoint);
            ConnectionPoint.Advise(this, out m_ConnectionPointCookie);
        }
        #endregion

        #region IOPCDataCallback Members

        public void OnCancelComplete(int dwTransid, int hGroup)
        {
            throw new NotImplementedException();
        }

        public void OnDataChange(int dwTransid, int hGroup, int hrMasterquality, int hrMastererror, int dwCount, int[] phClientItems, object[] pvValues, short[] pwQualities, OpcRcw.Da.FILETIME[] pftTimeStamps, int[] pErrors)
        {
            if (hGroup == m_hClient)
            {
                IList<DataValue> DataValues = new List<DataValue>();
                for (int i = 0; i < dwCount; i++)
                {
                    DataValue dataValue = new DataValue(phClientItems[i], pvValues[i], pwQualities[i], pftTimeStamps[i], pErrors[i]);
                    DataValues.Add(dataValue);
                }
                m_fct(DataValues);
            }
        }

        public void OnReadComplete(int dwTransid, int hGroup, int hrMasterquality, int hrMastererror, int dwCount, int[] phClientItems, object[] pvValues, short[] pwQualities, OpcRcw.Da.FILETIME[] pftTimeStamps, int[] pErrors)
        {
            throw new NotImplementedException();
        }

        public void OnWriteComplete(int dwTransid, int hGroup, int hrMastererr, int dwCount, int[] pClienthandles, int[] pErrors)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Private fields
        int m_hClient;
        DataChange m_fct;
        int m_ConnectionPointCookie;
        #endregion
    }
}
