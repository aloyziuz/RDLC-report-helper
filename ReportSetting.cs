using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace RDLCReportHelper
{
    public class ReportSetting
    {
        public string FontFamily { get; set; }
        public int FontSize { get; set; }
        public float PageWidth { get; set; }
        public float PageHeight { get; set; }
        public float MarginLeft { get; set; }
        public float MarginRight { get; set; }
        public float MarginTop { get; set; }
        public float MarginBottom { get; set; }

        public static ReportSetting ConvertToReportSetting(DataSet ds, string modulename)
        {
            ReportSetting res = new ReportSetting();
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                string attribute = res.RemoveModuleName((string)dr[0], modulename);
                PropertyInfo prop = res.GetType().GetProperty(attribute);
                var setter = prop.GetSetMethod();
                setter.Invoke(res, new object[] { Convert.ChangeType(dr[1], setter.GetParameters()[0].ParameterType) });
            }
            return res;
        }

        private string RemoveModuleName(string key, string modulename)
        {
            int index = key.IndexOf(modulename, StringComparison.Ordinal);
            return (index < 0) ? key : key.Remove(index, modulename.Length);
        }
    }
}
