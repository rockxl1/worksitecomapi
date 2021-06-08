using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects.Enums
{
    public enum XTR_EnumHistEvent
    {
        imHistoryLaunch = 0,
        imHistoryView = 1,
        imHistoryCheckout = 2,
        imHistoryCheckin = 3,
        imHistoryProfileEdit = 4,
        imHistoryClose = 5,
        imHistoryNew = 6,
        imHistoryVersionNew = 7,
        imHistorySecurityChange = 8,
        imHistoryCopy = 9,
        imHistoryPrint = 10,
        imHistoryMail = 11,
        imHistoryEchoSync = 12,
        imHistoryDelete = 13,
        imHistoryPurge = 14,
        imHistoryArchive = 15,
        imHistoryRestore = 16,
        imHistoryRelease = 17,
        imHistoryExport = 18,
        imHistoryModify = 19,
        imHistoryEditTime = 20,
        imHistoryNotLogged = 21,
        imHistoryFrozen = 22,
        imHistoryMigrated = 23,
        imHistoryUndeclared = 24,
        imHistoryReconciled = 25,
        imHistoryRemoveFromFolder = 26,
        imHistoryDeleteFolder = 27,
        imHistoryWorkflowEvent = 28,
        imHistoryWorkflowAttach = 29,
        imHistoryWorkflowDetach = 30,
        imHistoryShred = 1025,
    }
}
