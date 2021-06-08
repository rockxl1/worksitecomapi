using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects
{
    public class XTR_Parameter<T>
    {
        private T _Column;

        public T Column
        {
            get { return (T)(object)_Column; }
            set { _Column = value; }
        }
        public string Value { get; set; }
        public string[] ValueOr { get; set; }
    }
}
