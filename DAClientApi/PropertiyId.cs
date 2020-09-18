using System;
using System.Collections.Generic;
using System.Text;

namespace Siemens.Opc.Da
{
    public enum PropertiyId
    {
        Invalid = 0,
        DataType = 1,
        Value = 2,
        Quality = 3,
        Timestamp = 4,
        AccessRights = 5,
        ScanRate = 6,
        EUType = 7,
        EUInfo = 8,
        EUUnits = 100,
        Description = 101,
        EUHigh = 102,
        EULow = 103,
        IRHigh = 104,
        IRLow = 105,
        LabelClose = 106,
        LabelOpen = 107,
        Timezone = 108,
        ConditionStatus = 300,
        AlarmQuickHelp = 301,
        AlarmAreasList = 302,
        PrimaryAlarmArea = 303,
        ConditionLogic = 304,
        LimitExceeded = 305,
        Deadband = 306,
        LimitHiHi = 307,
        LimitHi = 308,
        LimitLo = 309,
        LimitLoLo = 310,
        LimitChangeRate = 311,
        LimitDeviation = 312,
        SoundFile = 313
    }
}
