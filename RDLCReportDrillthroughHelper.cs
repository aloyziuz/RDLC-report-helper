using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RDLCReportHelper
{
    public class DrillthroughHelper
    {
        /// <summary>
        /// get the report parameter (usually clicked value) of a named parameter in the report
        /// </summary>
        /// <param name="e">event args</param>
        /// <param name="parametername">name of the report parameter containing value</param>
        /// <returns>content string of the parameter</returns>
        public static string GetDrillthroughArgs(DrillthroughEventArgs e, string parametername)
        {
            var param = e.Report.GetParameters();
            var match = from p in param where p.Name == parametername select p;
            if(match.Count() > 0)
            {
                return match.First().Values[0];
            }
            return "";
        }
    }
}
