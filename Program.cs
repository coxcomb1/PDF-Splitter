using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;

namespace PDFSplitter
{
    public class SplitPdfByPageNumber
    {

        [STAThread] // Add STAThreadAttribute to ensure single-threaded apartment (STA) mode
        public static void Main(String[] args)
        {
            try
            {
                Console.WriteLine("Hello!\nWelcome to the PDF-splitter!\n");
                Console.WriteLine("Start by choosing the PDF document that you want to split:\n");
                
                string source = new SplitPdfByPageNumber().SelectFile();
                
                if (source == "Cancel")
                {
                    Console.WriteLine("You did not choose any file. Ending process");
                    Environment.Exit(0);
                }

                FileInfo sourceFile = new FileInfo(source);
                // placed in the same directory as source with the changed name DEST
                //                                                    remove the ".pdf" extension substring from the name
                string DEST = sourceFile.DirectoryName + @"\\" + sourceFile.Name.Remove(sourceFile.Name.Length - 4) + "-split_{0}.pdf";
                FileInfo file = new FileInfo(DEST);
                file.Directory.Create();
                
                Console.WriteLine("You chose to split the document: " + sourceFile.Name + ".\n");
                Console.WriteLine("Next, scroll through your document and choose at what pages you want to split the document.\nYou may choose multiple pages to split at.");
                Console.WriteLine("Document length: " + new SplitPdfByPageNumber().GetNumberOfPdfPages(source) + " pages\n");

                Console.WriteLine("Opening file...\n");
                var p = new Process();
                p.StartInfo = new ProcessStartInfo(source)
                {
                    UseShellExecute = true
                };
                p.Start();

                Console.WriteLine("At what pages do you want to split? Use commas (,) as separators (Example input: \"3,8,15\")");
                string in_pageSplitInputString = Console.ReadLine();
                List<int> splitAtPageNumbers = in_pageSplitInputString.Split(',').Select(int.Parse).ToList();

                bool isFileLocked = new SplitPdfByPageNumber().IsFileLocked(DEST);

                if (isFileLocked == false)
                {
                    Console.WriteLine("Starting PDF splitting\n");
                    int numberOfSplitFiles = new SplitPdfByPageNumber().ManipulatePdf(DEST, splitAtPageNumbers, source);

                    Console.WriteLine("The PDF was successfully split into " + numberOfSplitFiles + " files!");
                }
                else
                {
                    Console.WriteLine("The file was in use or another error occured. Please close the file and retry.");
                }

                Console.WriteLine("Ending execution");
                Console.WriteLine("Press Enter to exit");
                Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine("\n********* ERROR *********\n" + e.Message.ToString() + "\n********* ERROR *********\n");
                Console.WriteLine("Press Enter to exit");
                Console.ReadLine();
            }
        }

        [STAThread] // Add STAThreadAttribute to ensure single-threaded apartment (STA) mode
        private string SelectFile()
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "This PC";
                openFileDialog.Filter = "PDF Files(*.pdf)|*.pdf";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of the selected file
                    filePath = openFileDialog.FileName;
                    MessageBox.Show(filePath, "You Chose the File:", MessageBoxButtons.OKCancel);
                }
                else
                {
                    filePath = "Cancel";
                }
            }

            return filePath;
        }

        protected int ManipulatePdf(String dest, List<int> splitAtPageNumbers, String source)
        {
            PdfDocument pdfDoc = new PdfDocument(new PdfReader(source));
            IList<PdfDocument> splitDocuments = new SplitPdf(pdfDoc, dest).SplitByPageNumbers(splitAtPageNumbers);

            foreach (PdfDocument doc in splitDocuments)
            {
                doc.Close();
            }

            pdfDoc.Close();

            return splitDocuments.Count();
        }

        private class SplitPdf : PdfSplitter
        {
            private String dest;
            private int partNumber = 1;

            public SplitPdf(PdfDocument pdfDocument, String dest) : base(pdfDocument)
            {
                this.dest = dest;
            }

            protected override PdfWriter GetNextPdfWriter(PageRange documentPageRange)
            {
                return new PdfWriter(String.Format(dest, partNumber++));
            }
        }
        private bool IsFileLocked(string file)
        {
            //check that problem is not in destination file
            if (File.Exists(file) == true)
            {
                FileStream stream = null;
                try
                {
                    stream = new FileInfo(file).Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                    //stream = File.Open(file, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                }
                catch (Exception ex2)
                {
                    //_log.WriteLog(ex2, "Error in checking whether file is locked " + file);
                    int errorCode = Marshal.GetHRForException(ex2) & ((1 << 16) - 1);
                    if ((ex2 is IOException)) //&& (errorCode == ERROR_SHARING_VIOLATION || errorCode == ERROR_LOCK_VIOLATION))
                    {
                        return true;
                    }
                }
                finally
                {
                    if (stream != null)
                        stream.Close();
                }
            }
            return false;
        }

        public int GetNumberOfPdfPages(string filePath)
        {
            using (StreamReader sr = new StreamReader(File.OpenRead(filePath)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());

                return matches.Count;
            }
        }
    }
}