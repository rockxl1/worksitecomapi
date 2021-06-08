using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IManage;


namespace XTRWorkSite.WSObjects
{
    public class XTR_Data
    {
        private IManDateRange date;
        //public DateTime AbsoluteEndDate { get; set; }
        //public DateTime AbsoluteStartDate { get; set; }

        public XTR_Data(WorkSiteAccess db, DateTime? start, DateTime? end, int? daysBack)
        {
            if (db._m_dms == null)
            {
                db.LoadIManageLibrary();
            }
           date = db._m_dms.CreateDateRange();

           if (start.HasValue)
           {
               date.AbsoluteStartDate.Value = start.Value;
           }

           if (end.HasValue)
           {
               date.AbsoluteEndDate.Value = end.Value;
           }

           if (daysBack.HasValue)
           {
               date.DateRangeType = imDateRangeType.imRelativeDateRangeType;
               date.RelativeGranularity = imGranularity.eDay;
               date.RelativeStartDate.Value = daysBack.Value;
               date.RelativeEndDate.Value = 0;
           }
        }

        public string GetDate()
        {
            return date.Value;
        }

    }
}
