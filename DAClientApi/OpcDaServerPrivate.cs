using System;
using System.Collections.Generic;
using System.Text;
using OpcRcw.Da;

namespace Siemens.Opc.Da
{
    internal class ServerPrivate : IOPCDataCallback
    {
        #region Construction
        internal ServerPrivate(object group)
        {
            Guid iid = Guid.Empty;
            iid = typeof(IOPCDataCallback).GUID;
            OpcRcw.Comn.IConnectionPointContainer ConnectionPointContainer = (OpcRcw.Comn.IConnectionPointContainer)group;
            OpcRcw.Comn.IConnectionPoint ConnectionPoint;
            ConnectionPointContainer.FindConnectionPoint(ref iid, out ConnectionPoint);
            ConnectionPoint.Advise(this, out m_ConnectionPointCookie);
        }
        #endregion

        #region PublicProperties
        public ReadHandler EndRead
        {
            get { return m_endRead; }
            set { m_endRead = value; }
        }
        #endregion

        #region IOPCDataCallback Members

        public void OnCancelComplete(int dwTransid, int hGroup)
        {
            throw new NotImplementedException();
        }

        public void OnDataChange(int dwTransid, int hGroup, int hrMasterquality, int hrMastererror, int dwCount, int[] phClientItems, object[] pvValues, short[] pwQualities, OpcRcw.Da.FILETIME[] pftTimeStamps, int[] pErrors)
        {
            throw new NotImplementedException();
        }

        public void OnReadComplete(int dwTransid, int hGroup, int hrMasterquality, int hrMastererror, int dwCount, int[] phClientItems, object[] pvValues, short[] pwQualities, OpcRcw.Da.FILETIME[] pftTimeStamps, int[] pErrors)
        {
            Console.WriteLine("End Read");
        }

        public void OnWriteComplete(int dwTransid, int hGroup, int hrMastererr, int dwCount, int[] pClienthandles, int[] pErrors)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Private fields
        int m_ConnectionPointCookie;
        ReadHandler m_endRead = null;
        #endregion
    }
}
