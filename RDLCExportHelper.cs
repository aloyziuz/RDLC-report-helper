using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace RDLCReportHelper
{
    public class RDLCExportHelper
    {
        private ReportSetting setting;
        private const string DEVICE_INFO = @"<DeviceInfo>
                <OutputFormat>{6}</OutputFormat>
                <PageWidth>{0}cm</PageWidth>
                <PageHeight>{1}cm</PageHeight>
                <MarginTop>{2}cm</MarginTop>
                <MarginLeft>{3}cm</MarginLeft>
                <MarginRight>{4}cm</MarginRight>
                <MarginBottom>{5}cm</MarginBottom>
                <PrintDpiX>{7}</PrintDpiX>
                <PrintDpiY>{7}</PrintDpiY>
            </DeviceInfo>";

        public int DPI { get; set; } = 96;




        public RDLCExportHelper(ReportSetting setting)
        {
            this.setting = setting;
        }

        public void ExportReport(DirectPrintParameter pm, string savePath, string rendertype, 
            RDLCBuildType t = RDLCBuildType.Content, string exportFormat = "EMF")
        {
            LocalReport r = new LocalReport();

            switch (t)
            {
                case RDLCBuildType.Content:
                    r.ReportPath = pm.RDLCPath;
                    break;
                case RDLCBuildType.Embedded:
                    r.ReportEmbeddedResource = pm.RDLCPath;
                    break;
            }

            r.DataSources.Add(new ReportDataSource(pm.DatasetName, pm.DataSource));
            if (pm.Parameters != null)
                r.SetParameters(pm.Parameters);

            RenderAndSave(r, rendertype, savePath);
        }

        public void RenderAndSave(LocalReport r, string type, string savePath)
        {
            string mimeType, encoding, fileNameExtension;
            Warning[] warnings;
            string[] streams;

            var info = string.Format(DEVICE_INFO, setting.PageWidth, setting.PageHeight, setting.MarginTop,
                setting.MarginLeft, setting.MarginRight, setting.MarginBottom, this.DPI);
            
            var res = r.Render(type, info, out mimeType, out encoding, 
                out fileNameExtension, out streams, out warnings);
            //
            using (var f = new FileStream(savePath, FileMode.OpenOrCreate,
                               FileAccess.Write, FileShare.None))
            {
                f.Write(res, 0, res.Length);
            }
        }
    }
}
