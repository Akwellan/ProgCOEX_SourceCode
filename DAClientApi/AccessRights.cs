using System;
using System.Collections.Generic;
using System.Text;

using OpcRcw.Da;

namespace Siemens.Opc.Da
{
    [Flags]
    public enum AccessRights
    {
        Readable = OpcRcw.Da.Constants.OPC_READABLE,
        Writable = OpcRcw.Da.Constants.OPC_WRITEABLE
    }
}
