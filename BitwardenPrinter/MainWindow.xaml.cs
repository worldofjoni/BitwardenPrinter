using CsvHelper;
using Microsoft.Win32;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;

namespace BitwardenPrinter
{
    // TODO:
    // - UI Refinement

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        }

        const String filePrefix = @"BitwardenPrinter-";

        private void buttonFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = GetDownloadFolderPath();
            if (openFileDialog.ShowDialog() == true)
            {
                labelFile.Text = "Opened file: " + openFileDialog.FileName;
                var path = openFileDialog.FileName;

                using (var reader = new StreamReader(path))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<CsvEntry>();

                    // create PDF
                    var pdfFilename = createPDF(records, extractDate(path));
                    openPDF(pdfFilename);
                }
            }
        }

        String createPDF(IEnumerable<CsvEntry> records, DateTime rawDate)
        {

            Document document = new Document();
            document.Info.Title = "BitwardenPrinter";

            // extract date From Filename
            var date = rawDate.ToString("d");
            var user = textAccount.Text;
            int cnt = 0;

            var section = document.AddSection();
            section.PageSetup = document.DefaultPageSetup.Clone();
            float sectionWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;

            var header = section.Headers.Primary.AddParagraph();
            header.Format.Alignment = ParagraphAlignment.Left;
            header.AddText("Bitwarden Export");
            header.AddTab();
            header.AddText(user);
            header.AddTab();
            header.AddText(date);
            header.Format.ClearAll();
            header.Format.AddTabStop(sectionWidth / 2, TabAlignment.Center);
            header.Format.AddTabStop(sectionWidth, TabAlignment.Right);

            var footer = section.Footers.Primary.AddParagraph();
            footer.Format.Alignment = ParagraphAlignment.Center;
            footer.AddPageField();
            footer.AddText(" / ");
            footer.AddNumPagesField();

            // Title
            var titlePar = section.AddParagraph();
            titlePar.Format.Alignment = ParagraphAlignment.Center;
            titlePar.Format.Font.Size = 24;
            titlePar.AddFormattedText("Bitwarden Vault Export", TextFormat.Bold);
            
            section.AddParagraph();

            var infoPar = section.AddParagraph();
            // add text later

            section.AddParagraph();

            //table
            float columnWidth = sectionWidth / 3;

            var table = section.AddTable();
            table.Borders.Color = Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            // collumns:
            table.AddColumn(columnWidth);
            table.AddColumn(columnWidth);
            table.AddColumn(columnWidth);

            //rows
            var row = table.AddRow();
            row.HeadingFormat = true; 
            row.Format.Font.Bold = true;

            row.Cells[0].AddParagraph("Credential Name");
            row.Cells[1].AddParagraph("Username");
            row.Cells[2].AddParagraph("Password");

            //table.SetEdge()

            // data rows

            foreach ( var r in records)
            {
                row = table.AddRow();
                row.Cells[0].AddParagraph(zwnj2(r.name));
                row.Cells[1].AddParagraph(zwnj2(r.login_username));
                row.Cells[2].AddParagraph(zwnj(r.login_password));
                cnt++;
            }
            // add info text (row count is now calculated)
            var infos =
                "Account: " + user + Environment.NewLine +
                "Date: " + date + Environment.NewLine + // TODO: use export date from filename
                "Entries: " + cnt + Environment.NewLine +
                "";
            infoPar.AddText(infos);



            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(false);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();

            // save PDF
            string filename = System.IO.Path.GetTempPath() + filePrefix + Guid.NewGuid().ToString() + @".pdf";
            pdfRenderer.PdfDocument.Save(filename);


            return filename;

        }

        List<Process> processes = new List<Process>();

        void openPDF(String filename)
        {
            var info = new ProcessStartInfo()
            {
                UseShellExecute = true,
                FileName = filename,
                Verb = "open",
            };
            processes.Add(Process.Start(info));
        }


        static string GetDownloadFolderPath()
        {
            return Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();
        }

        void cleanup()
        {
            // stop open processes
            processes.ForEach(p => p.Kill());
            processes.ForEach(p => p.WaitForExit());
            

            // delete files (should be unnecessary)
            var files = Directory.GetFiles(System.IO.Path.GetTempPath(), filePrefix + "*.*", SearchOption.TopDirectoryOnly);
            foreach (var f in files)
            {
                try { File.Delete(f); }
                catch {}
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            cleanup();
        }

        String zwnj(String str)
        {
            String ret = "";
            foreach (var c in str)
            {
                ret += c + "\x200c";
            }
            return ret;
        }

        String zwnj2(String str)
        {
            String ret = "";
            foreach (var c in str)
            {
                ret += c + (@"@.:-_/\|,;".IndexOf(c) != -1 ? "\x200c" : "");
            }
            return ret;
        }

        DateTime extractDate(String path)
        {
            path = Path.GetFileName(path);
            path = path.Replace("bitwarden_export_","").Replace(".csv","");
            DateTime time;
            try { 
                time = DateTime.ParseExact(path,"yyyyMMddHHmmss",CultureInfo.InvariantCulture);
            } catch
            {
                time = DateTime.Now;
            }

            return time;
        }
    }

    
}
