using System;
using System.Collections.Generic;
using System.Text;

namespace Siemens.Opc.Da
{
    public class Category
    {
        #region Public Interface
        /// <summary>
        /// A unique identifier for the category.
        /// </summary>
        public int ID
        {
            get { return m_id; }
            set { m_id = value; }
        }

        /// <summary>
        /// The unique name for the category.
        /// </summary>
        public string Name
        {
            get { return m_name; }
            set { m_name = value; }
        }

        /// <summary>
        /// Returns a string that represents the current object.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Name;
        }
        #endregion

        #region Fields
        private int m_id = 0;
        private string m_name = null;
        #endregion
    }

    public class Attribute
    {
        #region Public Interface
        /// <summary>
        /// A unique identifier for the attribute.
        /// </summary>
        public int ID
        {
            get { return m_id; }
            set { m_id = value; }
        }

        /// <summary>
        /// The unique name for the attribute.
        /// </summary>
        public string Name
        {
            get { return m_name; }
            set { m_name = value; }
        }

        /// <summary>
        /// The data type of the attribute.
        /// </summary>
        public System.Type DataType
        {
            get { return m_datatype; }
            set { m_datatype = value; }
        }

        /// <summary>
        /// Returns a string that represents the current object.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Name;
        }
        #endregion

        #region Fields
        private int m_id = 0;
        private string m_name = null;
        private System.Type m_datatype = null;
        #endregion
    }

    public class EventAcknowledgement
    {
        #region Public Interface
        /// <summary>
        /// The name of the source that generated the event.
        /// </summary>
        public string SourceName
        {
            get { return m_sourceName; }
            set { m_sourceName = value; }
        }

        /// <summary>
        /// The name of the condition that is being acknowledged.
        /// </summary>
        public string ConditionName
        {
            get { return m_conditionName; }
            set { m_conditionName = value; }
        }

        /// <summary>
        /// The time that the condition or sub-condition became active.
        /// </summary>
        public DateTime ActiveTime
        {
            get { return m_activeTime; }
            set { m_activeTime = value; }
        }

        /// <summary>
        /// The cookie for the condition passed to client during the event notification.
        /// </summary>
        public int Cookie
        {
            get { return m_cookie; }
            set { m_cookie = value; }
        }

        /// <summary>
        /// Constructs an acknowledgment with its default values.
        /// </summary>
        public EventAcknowledgement() { }

        /// <summary>
        /// Constructs an acknowledgment from an event notification.
        /// </summary>
        public EventAcknowledgement(EventNotification notification)
        {
            m_sourceName = notification.SourceID;
            m_conditionName = notification.ConditionName;
            m_activeTime = notification.ActiveTime;
            m_cookie = notification.Cookie;
        }
        #endregion

        #region Fields
        private string m_sourceName = null;
        private string m_conditionName = null;
        private DateTime m_activeTime = DateTime.MinValue;
        private int m_cookie = 0;
        #endregion
    }
}
