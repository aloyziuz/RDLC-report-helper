using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace RDLCReportHelper
{
    /// <summary>
    /// Interaction logic for RDLCReportViewer.xaml
    /// </summary>
    public partial class RDLCReportViewerUC : UserControl
    {
        public ReportViewer ReportViewer { get => reportViewer; }
        public DrillthroughEventHandler DrillthroughHandler { get; set; }

        public RDLCReportViewerUC()
        {
            InitializeComponent();

            ReportViewer.Drillthrough += ReportViewer_Drillthrough;
        }

        private void ReportViewer_Drillthrough(object sender, DrillthroughEventArgs e)
        {
            if (DrillthroughHandler != null) 
                DrillthroughHandler(sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="setting">measurement sizes in cm</param>
        public void SetPageSettings(ReportSetting setting)
        {
            RDLCReportPrintHelper h = new RDLCReportPrintHelper(setting);
            var page = h.GetPageSettings();
            reportViewer.SetPageSettings(page);
            reportViewer.RefreshReport();
        }
    }
}
