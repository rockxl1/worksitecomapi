 using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects
{
    public class XTR_SearchParameters<T>
    {
        public List<XTR_Parameter<T>> Parameters { get; set; }
        /// <summary>
        /// Max size 9999
        /// </summary>
        public int? MaxRowCount { get; set; }
        public bool SearchDown {get;set;}

        public XTR_SearchParameters()
        {
            Parameters = new List<XTR_Parameter<T>>();
        }
    }

   
}
