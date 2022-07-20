using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace RDLCReportHelper
{
    public class RDLCReportPrintHelper : IDisposable
    {
        private int m_currentPageIndex;
        private IList<MemoryStream> m_streams;
        private float pWidth;
        private float pHeight;
        private float margintop;
        private float marginbottom;
        private float marginleft;
        private float marginright;
        private readonly float INCH_PER_CM = 0.3937f;

        private const string DEVICE_INFO = @"<DeviceInfo>
                <OutputFormat>{6}</OutputFormat>
                <PageWidth>{0}cm</PageWidth>
                <PageHeight>{1}cm</PageHeight>
                <MarginTop>{2}cm</MarginTop>
                <MarginLeft>{3}cm</MarginLeft>
                <MarginRight>{4}cm</MarginRight>
                <MarginBottom>{5}cm</MarginBottom>
                {7}
            </DeviceInfo>";
        private const string PRINT_DPI = @"
                <PrintDpiX>{0}</PrintDpiX>
                <PrintDpiY>{0}</PrintDpiY>";

        public int? DPI { get; set; }

        /// <summary>
        /// all the units are in cm
        /// </summary>
        /// <param name="pagewidth"></param>
        /// <param name="pageheight"></param>
        /// <param name="pagemargin"></param>
        public RDLCReportPrintHelper(ReportSetting setting)
        {
            pWidth = setting.PageWidth;
            pHeight = setting.PageHeight;
            this.marginbottom = setting.MarginBottom;
            this.margintop = setting.MarginTop;
            this.marginleft = setting.MarginLeft;
            this.marginright = setting.MarginRight;
        }

        public ReportSetting GetCurrentSettingInInch()
        {
            return new ReportSetting
            {
                PageHeight = pHeight,
                PageWidth = pWidth,
                MarginBottom = marginbottom,
                MarginLeft = marginleft,
                MarginRight = marginright,
                MarginTop = margintop
            };
        }

        public void DirectPrint(DirectPrintParameter pm, RDLCBuildType t = RDLCBuildType.Content, 
            string exportFormat = "EMF")
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
            Export(r, exportFormat);
            Print();
        }

        private float CmToInch(float unitInCm)
        {
            return unitInCm * INCH_PER_CM;
        }

        public PageSettings GetPageSettings()
        {
            var res = new PageSettings();
            SetPageSettings(res, Enumerable.Empty<PaperSize>());
            return res;
        }

        private void SetPageSettings(PageSettings s, IEnumerable<PaperSize> availablesizes)
        {
            int w, h;
            w = Convert.ToInt32(CmToInch(pWidth) * 100);
            h = Convert.ToInt32(CmToInch(pHeight) * 100);

            var match = from size in availablesizes where size.Width == w && size.Height == h select size;
            if (match.Count() > 0)
            {
                s.PaperSize = match.First();
            }
            else
            {
                s.PaperSize = new PaperSize("Custom", w, h);
                s.PaperSize.RawKind = 139;
            }
            s.Margins = new Margins(Convert.ToInt32(CmToInch(marginleft) * 100),
                Convert.ToInt32(CmToInch(marginright) * 100),
                Convert.ToInt32(CmToInch(margintop) * 100),
                Convert.ToInt32(CmToInch(marginbottom) * 100));
        }

        // Routine to provide to the report renderer, in order to
        //    save an image for each page of the report.
        private Stream CreateStream(string name,
          string fileNameExtension, Encoding encoding,
          string mimeType, bool willSeek)
        {
            var stream = new MemoryStream();
            m_streams.Add(stream);
            return stream;
        }

        // Export the given report as an EMF (Enhanced Metafile) file.
        private void Export(LocalReport report, string outFormat)
        {
            var printdpi = "";
            if (this.DPI.HasValue)
                printdpi = string.Format(PRINT_DPI, this.DPI.Value);

            var deviceInfo = string.Format(DEVICE_INFO, pWidth, pHeight, margintop, marginleft, marginright, 
                marginbottom, outFormat, printdpi);
            Warning[] warnings;
            m_streams = new List<MemoryStream>();
            report.Render("Image", deviceInfo, CreateStream, out warnings);
            foreach (Stream stream in m_streams)
                stream.Position = 0;

            Console.WriteLine(m_streams.Count);
        }

        // Handler for PrintPageEvents\
        private void PrintPage(object sender, PrintPageEventArgs ev)
        {
            Metafile pageImage = new
               Metafile(m_streams[m_currentPageIndex]);

            // Adjust rectangular area with printer margins.
            Rectangle adjustedRect = new Rectangle();
            adjustedRect = new Rectangle(
                    ev.PageBounds.Left - (int)ev.PageSettings.HardMarginX,
                    ev.PageBounds.Top - (int)ev.PageSettings.HardMarginY,
                    ev.PageSettings.PaperSize.Width,
                    ev.PageSettings.PaperSize.Height);

            // Draw a white background for the report
            ev.Graphics.FillRectangle(Brushes.White, adjustedRect);

            // Draw the report content
            ev.Graphics.DrawImage(pageImage, adjustedRect);

            // Prepare for the next page. Make sure we haven't hit the end.
            m_currentPageIndex++;
            ev.HasMorePages = (m_currentPageIndex < m_streams.Count);
        }

        private void Print()
        {
            if (m_streams == null || m_streams.Count == 0)
                throw new Exception("Error: no stream to print.");
            PrintDocument printDoc = new PrintDocument();
            if (!printDoc.PrinterSettings.IsValid)
            {
                throw new Exception("Error: cannot find the default printer.");
            }
            else
            {
                var sizes = Enumerate<PaperSize>(printDoc.PrinterSettings.PaperSizes.GetEnumerator());
                SetPageSettings(printDoc.DefaultPageSettings, sizes);
                SetPageSettings(printDoc.PrinterSettings.DefaultPageSettings, sizes);
                printDoc.PrintPage += new PrintPageEventHandler(PrintPage);
                m_currentPageIndex = 0;
                printDoc.Print();
            }
        }

        public IEnumerable<T> Enumerate<T>(System.Collections.IEnumerator enumerator)
        {
            List<T> res = new List<T>();
            while (enumerator.MoveNext())
            {
                res.Add((T)enumerator.Current);
            }
            return res;
        }

        public void Dispose()
        {
            if (m_streams != null)
            {
                foreach (Stream stream in m_streams)
                    stream.Close();
                m_streams = null;
            }
        }
    }
}
