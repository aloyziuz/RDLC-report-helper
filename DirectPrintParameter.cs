using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RDLCReportHelper
{
    public class DirectPrintParameter
    {
        public string DatasetName { get; set; }
        public string RDLCPath { get; set; }
        public IEnumerable<ReportParameter> Parameters { get; set; }
        public IEnumerable<object> DataSource { get; set; }
    }
}
