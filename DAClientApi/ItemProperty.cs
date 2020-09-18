using System;
using System.Collections.Generic;
using System.Text;

namespace Siemens.Opc.Da
{
    public class ItemProperty
    {
        #region Properties
        public PropertiyId PropertyId
        {
            get { return m_propertyId; }
            set { m_propertyId = value; }
        }

        public string Description
        {
            get { return m_description; }
            set { m_description = value; }
        }

        public string ItemId
        {
            get { return m_itemId; }
            set { m_itemId = value; }
        }

        public System.Type DataType
        {
            get { return m_datatype; }
            set { m_datatype = value; }
        }

        public Object Value
        {
            get { return m_value; }
            set { m_value = value; }
        }
        #endregion

        public override string ToString()
        {
            return PropertyId.ToString() + "|" + DataType + "|" + Value + "|" + ItemId + "|" + Description;
        }

        #region Private Fields
        private PropertiyId m_propertyId = PropertiyId.Invalid;
        private string m_description = null;
        private string m_itemId = null;
        private System.Type m_datatype = null;
        private Object m_value = null;
        #endregion
    }
}
