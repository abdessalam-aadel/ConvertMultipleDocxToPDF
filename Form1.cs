using System;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace ConvertMultipleDocxToPDF
{
    public partial class Form1 : Form
    {
        // Array of DOCX Files found in Folder
        string[] DOCXfiles;

        // Store slected path of Folder browser dialog in variable
        string selected_path;

        // Create fileCount to counting number of DOCX files found
        int fileCount = 0;

        public Form1() => InitializeComponent();

        // Handle Event Click of Buttton Start
        private void btnStart_Click(object sender, EventArgs e)
        {
            if (DOCXfiles == null || string.IsNullOrEmpty(TxtBoxLoad.Text))
            {
                labelErrorMessage.Text = "No source folder was selected, Please select one.";
                return;
            }

            else if (DOCXfiles.Length == 0)
            {
                labelErrorMessage.Text = "No DOCX file was found in the selected folder";
                return;
            }

            labelErrorMessage.Text = "";
            Cursor = Cursors.WaitCursor;
            labelInfo.Text = "Processing ...";
            labelErrorMessage.Text = "";

            // Create a new instance of Microsoft Word through the Interop library
            Word.Application wordApp = new Word.Application();

            // Log file
            string logFilePath = selected_path + @"\exceptions.log";
            // Delete the log file if it exists
            if (File.Exists(logFilePath))
            {
                File.Delete(logFilePath);
            }

            foreach (string file in DOCXfiles)
            {
                try
                {
                    // Open the document as read-only
                    Word.Document document = wordApp.Documents.Open(file, ReadOnly: true);
                    // Get Directory name
                    string targetFolder = Path.GetDirectoryName(file);
                    // Define the output PDF path
                    string pdfFileName = Path.GetFileNameWithoutExtension(file) + ".pdf";
                    string pdfFilePath = Path.Combine(targetFolder, pdfFileName);

                    // Save the document as PDF
                    document.SaveAs2(pdfFilePath, Word.WdSaveFormat.wdFormatPDF);

                    // Close the document without saving changes
                    document.Close(SaveChanges: false);
                }
                catch (Exception ex)
                {
                    // Write Exception into exceptions.log
                    LogException(logFilePath,file, ex);
                    // Continue to the next iteration
                    continue;
                }
            }

            // Quit the Word application
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);

            // Clear string array
            DOCXfiles = null;
            Cursor = Cursors.Default;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "Done.";
        }

        // Methode Write exceptions into log file
        static void LogException(string logFilePath,string filePath, Exception ex)
        {
            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                string filename = Path.GetFileNameWithoutExtension(filePath);
                writer.WriteLine($"{filename} : {ex.Message}");
            }
        }

        // Handle Event Click of Buttton Load Folder
        private void btnLoad_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FD = new FolderBrowserDialog();
            if (selected_path != null)
                FD.SelectedPath = selected_path;
            if (FD.ShowDialog() == DialogResult.OK)
            {
                string path = FD.SelectedPath;
                selected_path = path;
                TxtBoxLoad.Text = path;
                fileCount = SearchDOCXFiles(path, out DOCXfiles);
                // Check the Empty Folder
                labelInfo.Text = fileCount == 0 ? "Your Folder is Empty." : fileCount + " DOCX files found.";
                labelErrorMessage.Text = "";
            }
        }

        // Handle Methode Search in all Sub-Directory and Get all DOCX files found,
        // and bring out to the string array
        private int SearchDOCXFiles(string path, out string[] DOCXfiles)
        {
            DOCXfiles = Directory
                        .GetFiles(path, "*.*", SearchOption.AllDirectories)
                        .Where(s => s.ToLower().EndsWith(".doc") || s.ToLower().EndsWith(".docx"))
                        .ToArray();
            return DOCXfiles.Length;
        }

        private void GitLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Go to Github repository
            string url = "https://github.com/abdessalam-aadel/ConvertMultipleDocxToPDF";

            // Open the URL in the default web browser
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true // Ensures the URL is opened in the default web browser
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
            // Condition >> Drag Folder
            if (Directory.Exists(path))
            {
                TxtBoxLoad.Text = path;
                fileCount = SearchDOCXFiles(path, out DOCXfiles);
                selected_path = path;
                // Check the Empty Folder
                labelInfo.Text = fileCount == 0 ? "Your Folder is Empty." : fileCount + " DOCX files found.";
                labelErrorMessage.Text = "";
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "...";
            labelErrorMessage.Text = "";
            DOCXfiles = null;
        }
    }
}
