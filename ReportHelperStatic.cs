using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RDLCReportHelper
{
    public class ReportHelperStatic
    {
        public static void SetReportViewer(ReportViewer rv, IEnumerable<ReportDataSource> dataSources, string rdlcName,
            RDLCBuildType t = RDLCBuildType.Content, IEnumerable<ReportParameter> parameter = null)
        {
            if (dataSources != null)
            {
                foreach (var source in dataSources)
                    rv.LocalReport.DataSources.Add(source);

                switch (t)
                {
                    case RDLCBuildType.Embedded:
                        rv.LocalReport.ReportEmbeddedResource = rdlcName;
                        break;
                    case RDLCBuildType.Content:
                        rv.LocalReport.ReportPath = rdlcName;
                        break;
                }

                if (parameter != null)
                {
                    rv.LocalReport.SetParameters(parameter);
                }
                rv.ProcessingMode = ProcessingMode.Local;
                rv.RefreshReport();
            }
        }
    }
}
