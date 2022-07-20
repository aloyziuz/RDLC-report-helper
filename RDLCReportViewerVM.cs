using Microsoft.Reporting.WinForms;
using MVVMHelper;
using System.Collections.Generic;
using System.Data;

namespace RDLCReportHelper
{
    public enum RDLCBuildType
    {
        Content, Embedded
    }

    public class RDLCReportViewerVM:ObservableObject
    {
        public RDLCReportViewerVM(ReportViewer rv, string datasetName, string rdlcName, DataTable dt, 
            RDLCBuildType t = RDLCBuildType.Content, IEnumerable<ReportParameter> parameter = null)
        {
            if (!string.IsNullOrWhiteSpace(datasetName) || dt != null)
                rv.LocalReport.DataSources.Add(new ReportDataSource(datasetName, dt));
            Init(rv, rdlcName, t, parameter);
        }

        public RDLCReportViewerVM(ReportViewer rv, string datasetName, string rdlcName, IEnumerable<object> source,
            RDLCBuildType t = RDLCBuildType.Content, IEnumerable<ReportParameter> parameter = null)
        {
            if (!string.IsNullOrWhiteSpace(datasetName) || source != null)
                rv.LocalReport.DataSources.Add(new ReportDataSource(datasetName, source));
            Init(rv, rdlcName, t, parameter);
        }

        public RDLCReportViewerVM(ReportViewer rv, IEnumerable<ReportDataSource> dataSources, string rdlcName,
            RDLCBuildType t = RDLCBuildType.Content, IEnumerable<ReportParameter> parameter = null)
        {
            if (dataSources != null)
            {
                foreach (var source in dataSources)
                    rv.LocalReport.DataSources.Add(source);
                Init(rv, rdlcName, t, parameter);
            }
        }

        private void Init(ReportViewer rv, string rdlcName, RDLCBuildType t, IEnumerable<ReportParameter> parameter)
        {
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
