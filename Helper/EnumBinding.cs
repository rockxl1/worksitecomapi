using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.Helper
{
    public class EnumBinding
    {
        public static IManage.imProfileAttributeID GetImProfileAttributeID(XTRWorkSite.WSObjects.Enums.XTR_EnumProfileAttributeID columnId)
        {
            return (IManage.imProfileAttributeID)Enum.Parse(typeof(IManage.imProfileAttributeID), ((int)columnId).ToString());
        }

        //public static IManage.imProfileAttributeID GetImProfileAttributeID(XTRWorkSite.WSObjects.Enums.XTR_EnumWorkSpaceAttributes columnId)
        //{
        //    return (IManage.imProfileAttributeID)Enum.Parse(typeof(IManage.imProfileAttributeID), ((int)columnId).ToString());
        //}

        public static IMANADMIN.AttributeID GetAttributeID(XTRWorkSite.WSObjects.Enums.XTR_EnumAttributeID tableId)
        {
            return (IMANADMIN.AttributeID)Enum.Parse(typeof(IMANADMIN.AttributeID), ((int)tableId).ToString());
        }

        public static IManage.imFolderAttributeID GetImFolderAttributeID(XTRWorkSite.WSObjects.Enums.XTR_EnumFolderAttributeId attId)
        {
            return (IManage.imFolderAttributeID)Enum.Parse(typeof(IManage.imFolderAttributeID), ((int)attId).ToString());
        }

        public static IManage.imHistEvent GetHistEvent(XTRWorkSite.WSObjects.Enums.XTR_EnumHistEvent hist)
        {
            return (IManage.imHistEvent)Enum.Parse(typeof(IManage.imHistEvent), ((int)hist).ToString());
        }

        public static IManage.imFolderAttributeID GetWorkSpaceSearchParameteres(WSObjects.Enums.XTR_EnumWorkSpaceAttributes attId)
        {
            return (IManage.imFolderAttributeID)Enum.Parse(typeof(IManage.imFolderAttributeID), ((int)attId).ToString());
        }

        public static IManage.imCheckinDisposition GetCheckInDisposition(WSObjects.Enums.XTR_CheckInTypes type)
        {
            return (IManage.imCheckinDisposition)Enum.Parse(typeof(IManage.imCheckinDisposition), ((int)type).ToString());
        }

        public static IManage.imCheckinOptions GetCheckinOpt(WSObjects.Enums.XTR_CheckInOption opt)
        {
            return (IManage.imCheckinOptions)Enum.Parse(typeof(IManage.imCheckinOptions), ((int)opt).ToString());
        }

        

    }
}
