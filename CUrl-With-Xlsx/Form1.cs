using System;
using System.Windows.Forms;
using System.Net;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace CUrl_With_Xlsx
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (openExcel.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                textBox1.Text = openExcel.FileName;
                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(@"" + textBox1.Text)))
                {
                    foreach (var worksheet in xlPackage.Workbook.Worksheets)
                    {
                        comboBoxExcelSheetNames.Items.Add(worksheet.Name);
                    }
                }
                comboBoxExcelSheetNames.SelectedIndex = 0;
                
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result.ToString() == "Ok")
            {

                textBox2.Text = dialog.FileName;
            }
            dialog = null;

        }
        public void WriteLog(string log,TextBox tb)
        {
                tb.AppendText(log + Environment.NewLine);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox3.Text = String.Empty;
            GetxmlsxAsync();

        }

        private async void GetxmlsxAsync() {
            
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(@"" + textBox1.Text)))
            {

                var myWorksheet = xlPackage.Workbook.Worksheets[comboBoxExcelSheetNames.SelectedIndex]; //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                var ids = new List<string>();
                var uris = new List<Uri>();
                for (int rowNum = (checkBox1.Checked ? 2: 1); rowNum <= totalRows; rowNum++) //select starting row here
                {
                    if (Convert.ToString(myWorksheet.Cells[rowNum, Int32.Parse(textBox5.Text)].Value) != "-" && myWorksheet.Cells[rowNum, Int32.Parse(textBox5.Text)].Value != null && Convert.ToString(myWorksheet.Cells[rowNum, Int32.Parse(textBox5.Text)].Value) != "#N/D" && Uri.IsWellFormedUriString(Convert.ToString(myWorksheet.Cells[rowNum, Int32.Parse(textBox5.Text)].Value), UriKind.Absolute))
                    {

                        ids.Add(Convert.ToString(myWorksheet.Cells[rowNum, Int32.Parse(textBox4.Text)].Value));

                        uris.Add(new Uri(Convert.ToString(myWorksheet.Cells[rowNum, Int32.Parse(textBox5.Text)].Value)));
                    }
                }
                myWorksheet = null;
                progressBar2.Maximum = ids.Count;
                progressBar2.Value = 0;
                xlPackage.Dispose();
                label2.Text = "Arquivos: 0 de " + ids.Count();
                foreach (var doc in ids.Zip(uris,Tuple.Create))
                {
                   await Download(doc);
                    progressBar2.Value += 1;
                    label2.Text = "Arquivos: "+ progressBar2.Value + " de " + ids.Count();
                }
                if (checkBox2.Checked)
                {
                    System.Diagnostics.Process.Start("explorer.exe", textBox2.Text);
                }

            }
        }


        public async Task Download(Tuple<String, Uri> doc)
        {
            // Create a new WebClient instance.

            WebClient webClient = new WebClient
            {
                Proxy = null
            };
            progressBar1.Value = 0;
            progressBar1.Maximum = 100;
            webClient.DownloadProgressChanged += (s, e) =>
            {
                if (e.ProgressPercentage > 0 && label1.Text != "Completed:" + e.ProgressPercentage.ToString() + "%")
                {
                    progressBar1.Value = e.ProgressPercentage;
                    label1.Text = "Completed:" + progressBar1.Value.ToString() + "%";
                    textBox3.Text = textBox3.Text.Remove(textBox3.Text.LastIndexOf(Environment.NewLine));
                    textBox3.AppendText(Environment.NewLine + "Completed:" + progressBar1.Value.ToString() + "% ");
                }
            };
            webClient.DownloadFileCompleted += (s, e) =>
            {
                WriteLog(Environment.NewLine + "Successfully Downloaded File \"" + doc.Item1 + "\" from \"" + doc.Item2 + "\" " + Environment.NewLine, textBox3);
                WriteLog("Downloaded file saved in the following file system folder:" + Environment.NewLine + textBox2.Text, textBox3);
            };
            // Concatenate the domain with the Web resource filename.
            WriteLog(Environment.NewLine + "Downloading File \"" + doc.Item1 + "\" from \"" + doc.Item2 + "\" ......." + Environment.NewLine, textBox3);
            // Download the Web resource and save it into the current filesystem folder.
            await webClient.DownloadFileTaskAsync(doc.Item2, textBox2.Text +"\\"+ doc.Item1 + (textBox6.Text != "" ? "-"+ textBox6.Text : "" ) + textBox7.Text);
            webClient.Dispose();
            label1.Text = "";
            
        }

    }
}
