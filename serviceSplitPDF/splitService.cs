using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace serviceSplitPDF
{
    public partial class splitService : ServiceBase
    {
        public splitService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                FileSystemWatcher watcher = new FileSystemWatcher("C:\\RPA\\PP1450_Documents")
                {
                    Filter = "*.pdf",
                    EnableRaisingEvents = true
                };
                watcher.Created += OnCreated;
            }
            catch(Exception e)
            {
                File.AppendAllText("C:\\RPA\\ErrorLog.txt", e.Message.ToString());
            }

        }
        private static void OnCreated(Object source, FileSystemEventArgs e)
        {
            try
            {
                PdfDocument fullDoc = PdfReader.Open(e.FullPath, PdfDocumentOpenMode.Import);
                Console.WriteLine(e.ChangeType);

                PdfDocument newdoc = new PdfDocument();
                Random random = new Random();
                string fileNameGuid = DateTime.Now.ToString("yyyyMMdd_HHmmss_") + random.Next(1, 100000).ToString();
                for (int i = 0; i < fullDoc.PageCount; i++)
                {

                    newdoc.AddPage(fullDoc.Pages[i]);
                    if (i == 0)
                    {
                        newdoc.Save("C:\\RPA\\Prva\\" + fileNameGuid + ".pdf");
                        newdoc = new PdfDocument();
                    }

                }
                newdoc.Save("C:\\RPA\\Ostanato\\" + fileNameGuid + ".pdf");
            }
            catch(Exception e1)
            {
                File.AppendAllText("C:\\RPA\\ErrorLog.txt", e1.Message.ToString());
            }
        }
        protected override void OnStop()
        {
        }
    }
}
