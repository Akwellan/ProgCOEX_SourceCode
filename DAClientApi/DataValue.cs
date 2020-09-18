using System;
using System.Collections.Generic;
using System.Text;

namespace Siemens.Opc.Da
{
    public class DataValue
    {
        #region Construction
        public DataValue(
            int ClientHandle,
            object Value,
            short Quality,
            OpcRcw.Da.FILETIME TimeStamp,
            int Error)
        {
            m_clientHandle = ClientHandle;
            m_value = Value;
            m_quality = Quality;
            m_timeStamp = TimeStamp;
            m_error = Error;
        }
        #endregion

        #region Public Methods
        bool HasGoodQuality()
        {
            return (m_quality & 0X11000000) > 0;
        }
        bool IsGood()
        {
            return m_error == 0;
        }
        #endregion

        #region Public Properties
        public int ClientHandle
        {
            get { return m_clientHandle; }
        }
        public object Value
        {
            get { return m_value; }
        }
        public short Quality
        {
            get { return m_quality; }
        }
        public OpcRcw.Da.FILETIME TimeStamp
        {
            get { return m_timeStamp; }
        }
        public int Error
        {
            get { return m_error; }
        }
        #endregion

        #region toString
        public override string ToString()
        {
            return ClientHandle + "|" + Value.ToString();
        }
        #endregion

        #region Private fields
        int m_clientHandle;
        object m_value;
        short m_quality;
        OpcRcw.Da.FILETIME m_timeStamp;
        int m_error;
        #endregion
    }
}
