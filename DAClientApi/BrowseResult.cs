using System;
using System.Collections.Generic;
using System.Text;

namespace Siemens.Opc.Da
{
    public class BrowseResult
    {
        #region Properties
        public string ItemId
        {
            get { return m_itemId; }
            set { m_itemId = value; }
        }
        public string ItemName
        {
            get { return m_itemName; }
            set { m_itemName = value; }
        }

        public ItemProperty[] ItemProperties
        {
            get { return m_itemProperties; }
            set { m_itemProperties = value; }
        }
        #endregion

        #region Private Fields
        private string m_itemId = null;
        private string m_itemName = null;
        private ItemProperty[] m_itemProperties = null;
        #endregion
    }
}
