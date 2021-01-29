using AtexisTool.pdfExtract.Models;
using AtexisTool.pdfExtract.ViewModels;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Filter;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Org.BouncyCastle.Utilities;
using System.Diagnostics;

namespace AtexisTool.pdfExtract
{
    public partial class MainForm : Form
    {

        public int totalPages=0;

        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        TextBoxTraceListener _textBoxListener;

        FilesViewModel _filesViewModel;
        string[] lst;
        public MainForm()
        {
            InitializeComponent();
            _textBoxListener = new TextBoxTraceListener(logTextBox);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

            logger.Info("Setting background Colour");
            this.BackColor = Color.FromArgb(31, 51, 105);
            //this.dataGridView1.AutoGenerateColumns = false;            

            //for (int i = 1; i < 1000; i++)
            //{
            //    _textBoxListener.WriteLine($"Numero {i}: {Convert.ToChar(i).ToString()}");
            //    logTextBox.SelectionStart = logTextBox.TextLength;
            //    logTextBox.ScrollToCaret();
            //}

            logger.Info("Requested folder to extract info from pdf");
            _textBoxListener.WriteLine("REQUEST: Please, select folder to extract info from pdf");
            logTextBox.SelectionStart = logTextBox.TextLength;
            logTextBox.ScrollToCaret();
            //this.dataGridView1.DataSource = _filesViewModel;
        }

        private void SelectFolderButton_Click(object sender, EventArgs e)
        {
            _filesViewModel = new FilesViewModel();
            this.logTextBox.Text = "";
            this.dataGridView1.DataSource = null;

            this.folderBrowserDialog1.ShowNewFolderButton = true;
            DialogResult result = folderBrowserDialog1.ShowDialog();

            logger.Info("Select Folder button Clicked");
            _textBoxListener.WriteLine("INFO: Select Folder button Clicked");
            logTextBox.SelectionStart = logTextBox.TextLength;
            logTextBox.ScrollToCaret();

            if (result == DialogResult.OK)
            {


                txtPath.Text = folderBrowserDialog1.SelectedPath;
                Environment.SpecialFolder root = folderBrowserDialog1.RootFolder;

                lst = Directory.GetFiles(txtPath.Text, "*.pdf", SearchOption.TopDirectoryOnly);
                if (lst.Length > 0)
                {
                    createExcel();
                }
                int i = 1;
                foreach (string file in lst)
                {
                    _filesViewModel.FileListModel.Add(new FileModel
                    {
                        id = i,
                        filename = System.IO.Path.GetFileName(file),
                        status = "?",
                        remarks = ""
                    });
                    i++;
                    logger.Info($"New file have been found and added to list: {System.IO.Path.GetFileName(file)}");
                    _textBoxListener.WriteLine($"INFO: New file have been found and added to list: {System.IO.Path.GetFileName(file)}");
                    logTextBox.SelectionStart = logTextBox.TextLength;
                    logTextBox.ScrollToCaret();

                }

                logger.Info($"Total files found:{lst.Count().ToString()}");
                _textBoxListener.WriteLine($"INFO: Total files found:{lst.Count().ToString()}");
                logTextBox.SelectionStart = logTextBox.TextLength;
                logTextBox.ScrollToCaret();

                if (lst.Count() > 0)
                {
                    this.dataGridView1.DataSource = _filesViewModel.FileListModel;
                    this.dataGridView1.Columns[0].DataPropertyName = "id";
                    this.dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    this.dataGridView1.Columns[1].DataPropertyName = "filename";
                    this.dataGridView1.Columns[2].DataPropertyName = "status";
                    this.dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    this.dataGridView1.Columns[3].DataPropertyName = "remarks";                    
                    this.dataGridView1.Refresh();

                    logger.Info($"Analizing each file added");
                    _textBoxListener.WriteLine($"INFO: Analizing each file added");
                    logTextBox.SelectionStart = logTextBox.TextLength;
                    logTextBox.ScrollToCaret();

                    for (int j = 0; j <= dataGridView1.RowCount - 1; j++)
                    {
                        logTextBox.SelectionStart = logTextBox.TextLength;
                        logTextBox.ScrollToCaret();

                        pdfAnalize(System.IO.Path.Combine(txtPath.Text, dataGridView1.Rows[j].Cells[1].Value.ToString()), j);
                        if (dataGridView1.Rows[j].Cells[2].Value.ToString() != "KO")
                        {
                            dataGridView1.Rows[j].Cells[2].Value = "OK";
                            dataGridView1.Rows[j].Cells[2].Style.BackColor = Color.GreenYellow;
                            dataGridView1.Refresh();
                        }

                    }

                    //Total pages analized
                    workSheet.Cells[leftNumCell, 2].NumberFormat = "@";
                    workSheet.Cells[leftNumCell, 2] = $"Total pages analized: {totalPages} pages";
                }
            }
            else
            {
                logger.Warn("No folder selectec or cancel button have been clicked");
                _textBoxListener.WriteLine("WARNING: No folder selectec or cancel button have been clicked");
                logTextBox.SelectionStart = logTextBox.TextLength;
                logTextBox.ScrollToCaret();
            }

        }

        private void pdfAnalize(string src, int j)
        {
            iText.Kernel.Pdf.PdfDocument pdfDoc = new PdfDocument(new PdfReader(src));
            int n = pdfDoc.GetNumberOfPages();

            logger.Info($"********** Document:{src} --> Total Pages:{n} **********");
            _textBoxListener.WriteLine($"INFO: ********** Document:{src} --> Total Pages:{n} **********");
            logTextBox.SelectionStart = logTextBox.TextLength;
            logTextBox.ScrollToCaret();

            //Rectangle rect = new Rectangle(0, 0,595, 842);
            iText.Kernel.Geom.Rectangle rect = new iText.Kernel.Geom.Rectangle(0, 0, 595, 50);
            TextRegionEventFilter regionFilter = new TextRegionEventFilter(rect);

            for (int i = 1; i <= n; i++)
            {
                totalPages++;
                PdfPage page = pdfDoc.GetPage(i);
                ITextExtractionStrategy strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);
                string text = PdfTextExtractor.GetTextFromPage(page, strategy);
                //var result = Regex.Split(text, "\r\n|\r|\n");                
                if (text.Length > 0)
                {
                    analizeText(i, text, j);
                }
                else
                {
                    logger.Error($"Page: {i} --NO TEXT HAVE BEEN FOUND");
                    _textBoxListener.WriteLine($"ERROR: Page: {i} --NO TEXT HAVE BEEN FOUND");
                    logTextBox.SelectionStart = logTextBox.TextLength;
                    logTextBox.ScrollToCaret();

                    dataGridView1.Rows[j].Cells[2].Value = "KO";
                    dataGridView1.Rows[j].Cells[2].Style.BackColor = Color.Red;
                    if (dataGridView1.Rows[j].Cells[3].Value.ToString() == "" || dataGridView1.Rows[j].Cells[3].Value.ToString().Equals(""))
                    {
                        dataGridView1.Rows[j].Cells[3].Value = $" Check=> Page:{i}";
                    }
                    else
                    {
                        dataGridView1.Rows[j].Cells[3].Value = dataGridView1.Rows[j].Cells[3].Value + " | " + $"Page:{i}";
                    }
                    dataGridView1.Refresh();
                }
            }

            logger.Info($"********** Document:{src} --> Extract Pages finished **********");
            _textBoxListener.WriteLine($"INFO: ********** Document:{src} --> Extract Pages finished **********");
            logTextBox.SelectionStart = logTextBox.TextLength;
            logTextBox.ScrollToCaret();
            pdfDoc.Close();
        }

        private void analizeText(int page, string text, int j)
        {
            string auxtext = "";
            if (text.Trim().StartsWith("Revis"))
            {
                if (text.IndexOf("Revision", 0) >= 0)
                {
                    auxtext = text.Replace("Revision", "").Trim();
                    auxtext = auxtext.Replace("  ", " ").Trim();
                    auxtext = auxtext.Replace("   ", " ").Trim();
                }
                else if (text.IndexOf("Revisión", 0) >= 0)
                {
                    auxtext = text.Replace("Revisión", "").Trim();
                    auxtext = auxtext.Replace("  ", " ").Trim();
                    auxtext = auxtext.Replace("   ", " ").Trim();
                }
                else if (text.IndexOf("revision", 0) >= 0)
                {
                    auxtext = text.Replace("revision", "").Trim();
                    auxtext = auxtext.Replace("  ", " ").Trim();
                    auxtext = auxtext.Replace("   ", " ").Trim();
                }
                else if (text.IndexOf("revisión", 0) >= 0)
                {
                    auxtext = text.Replace("revisión", "").Trim();
                    auxtext = auxtext.Replace("  ", " ").Trim();
                    auxtext = auxtext.Replace("   ", " ").Trim();
                }

                string[] result = auxtext.Split(new char[] { ' ' });
                if (result.Length == 2)
                {
                    logger.Info($"Page: {page} -- found text: {text} => Result: {result[1]} {result[0]}");
                    _textBoxListener.WriteLine($"INFO: Page: {page} -- found text: {text} => Result: {result[1]} {result[0]}");
                    logTextBox.SelectionStart = logTextBox.TextLength;
                    logTextBox.ScrollToCaret();
                    updateExcel(result[1], int.Parse(result[0]));
                }
                else
                {
                    dataGridView1.Rows[j].Cells[2].Value = "KO";
                    dataGridView1.Rows[j].Cells[2].Style.BackColor = Color.Red;
                    dataGridView1.Rows[j].Cells[3].Value = $"Error Page {page}: {auxtext}";
                    dataGridView1.Refresh();
                }
            }
            else if (text.Trim().IndexOf("Revis") > 0)
            {
                if (text.IndexOf("Revision", 0) >= 0)
                {
                    auxtext = text.Replace("Revision", "").Trim();
                    auxtext = auxtext.Replace("  ", " ").Trim();
                    auxtext = auxtext.Replace("   ", " ").Trim();
                }
                else if (text.IndexOf("Revisión", 0) >= 0)
                {
                    auxtext = text.Replace("Revisión", "").Trim();
                    auxtext = auxtext.Replace("  ", " ").Trim();
                    auxtext = auxtext.Replace("   ", " ").Trim();
                }
                else if (text.IndexOf("revision", 0) >= 0)
                {
                    auxtext = text.Replace("revision", "").Trim();
                    auxtext = auxtext.Replace("  ", " ").Trim();
                    auxtext = auxtext.Replace("   ", " ").Trim();
                }
                else if (text.IndexOf("revisión", 0) >= 0)
                {
                    auxtext = text.Replace("revisión", "").Trim();
                    auxtext = auxtext.Replace("  ", " ").Trim();
                    auxtext = auxtext.Replace("   ", " ").Trim();
                }

                string[] result = auxtext.Split(new char[] { ' ' });
                if (result.Length == 2)
                {
                    logger.Info($"Page: {page} -- found text: {text} => Result: {result[0]} {result[1]}");
                    _textBoxListener.WriteLine($"INFO: Page: {page} -- found text: {text} => Result: {result[0]} {result[1]}");
                    logTextBox.SelectionStart = logTextBox.TextLength;
                    logTextBox.ScrollToCaret();

                    updateExcel(result[0], int.Parse(result[1]));

                }
                else
                {
                    dataGridView1.Rows[j].Cells[2].Value = "KO";
                    dataGridView1.Rows[j].Cells[2].Style.BackColor = Color.Red;
                    dataGridView1.Rows[j].Cells[2].Value = $"Error Page {page}: {auxtext}";
                    dataGridView1.Refresh();
                }
            }
            else
            {
                logger.Info($"Page: {page} -- found text: {text} => Result: {text} 0");
                _textBoxListener.WriteLine($"INFO: Page: {page} -- found text: {text} => Result: {text} 0");
                logTextBox.SelectionStart = logTextBox.TextLength;
                logTextBox.ScrollToCaret();
                updateExcel(text, 0);
            }
        }

        private Microsoft.Office.Interop.Excel.Application excel;
        private Workbook worKBooK;
        private Worksheet workSheet;
        private Range cellRange;

        private int leftNumCell = 3;
        private int rightNumCell = 3;
        private void createExcel()
        {
            leftNumCell = 3;
            rightNumCell = 3;

            try
            {
                logger.Info("Creanting Excel Workbook");
                _textBoxListener.WriteLine("INFO: Creanting Excel Workbook");
                logTextBox.SelectionStart = logTextBox.TextLength;
                logTextBox.ScrollToCaret();
                //this.BeginInvoke((System.Action)(() => this.logTextBox.Text= this.logTextBox.Text + Environment.NewLine + "Creanting Excel Workbook"));

                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                excel.DisplayAlerts = false;
                worKBooK = excel.Workbooks.Add(Type.Missing);

                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)worKBooK.ActiveSheet;
                workSheet.Name = "strFolder";

                logger.Info("Configuring excel columns and styles");
                _textBoxListener.WriteLine("INFO: Configuring excel columns and styles");
                logTextBox.SelectionStart = logTextBox.TextLength;
                logTextBox.ScrollToCaret();
                //workSheet.Range["A2:E5"].Style.Font.Underline = FontStyle.Underline;
                workSheet.Columns[1].ColumnWidth = 2;
                workSheet.Cells[1, 2] = "Page";
                workSheet.Cells[1, 2].Font.Underline = true;
                //workSheet.Cells[1, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;                
                workSheet.Columns[2].ColumnWidth = 25;
                workSheet.Columns[3].ColumnWidth = 8;
                workSheet.Cells[1, 3] = "Revision";
                workSheet.Cells[1, 3].Font.Underline = true;
                workSheet.Cells[1, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Columns[4].ColumnWidth = 2;
                workSheet.Cells[1, 5] = "Page";
                workSheet.Cells[1, 5].Font.Underline = true;
                //workSheet.Cells[1, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Columns[5].ColumnWidth = 25;
                workSheet.Cells[1, 6] = "Revision";
                workSheet.Cells[1, 6].Font.Underline = true;
                workSheet.Cells[1, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Columns[6].ColumnWidth = 8;
                workSheet.Columns[7].ColumnWidth = 2;
                logger.Info("Excel configured, but hidden");
                _textBoxListener.WriteLine("INFO: Excel configured, but hidden");
                logTextBox.SelectionStart = logTextBox.TextLength;
                logTextBox.ScrollToCaret();
            }

            catch (Exception ex)
            {
                logger.Error("An error have happened during excel creation or configurating");
                _textBoxListener.WriteLine("ERROR: An error have happened during excel creation or configurating");
                logTextBox.SelectionStart = logTextBox.TextLength;
                logTextBox.ScrollToCaret();
            }
        }

        string previousAta = "";
        string previousAtaPage = "";
        int previousPage = 9999;
        int previousRevision = 999;
        Boolean PreviousIsRoman;

        private void getValues(string ataPage, int revision)
        {
            previousAtaPage = ataPage;
            int numeric;

            string[] svalue = ataPage.Split(new char[] { '-' });
            if (svalue.Length == 2)
            {
                previousAta = svalue[0];

            }

            Boolean isNumeric = int.TryParse(svalue[1], out numeric);

            //if (Regex.IsMatch(svalue[1], @"^\d+$"))
            if (isNumeric)
            {
                previousPage = Convert.ToInt32(svalue[1]);
                PreviousIsRoman = false;
            }
            else
            {
                PreviousIsRoman = true;
                previousPage = RomanoToInt(svalue[1]);
            }

            previousRevision = revision;

        }
        private void updateExcel(string ataPage, int revision)
        {
            Boolean IsRoman = false;
            int numeric;
            try
            {
                if (previousAta == "" && previousPage == 9999 && previousRevision == 999)
                {
                    getValues(ataPage.Trim(), revision);

                    workSheet.Cells[leftNumCell, 2].NumberFormat = "@";
                    workSheet.Cells[leftNumCell, 2] = ataPage.Trim();
                    workSheet.Cells[leftNumCell, 3] = revision;

                    workSheet.Cells[leftNumCell, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    workSheet.Cells[leftNumCell, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    leftNumCell = leftNumCell + 1;
                }
                else
                {

                    string[] sValues = ataPage.Split(new char[] { '-' });
           
                    Boolean isNumeric = int.TryParse(sValues[1], out numeric);
                    if (!isNumeric)
                    {
                        IsRoman = true;
                    }

                    if (PreviousIsRoman == IsRoman)
                    {
                        if (previousAta == sValues[0].Trim() && previousRevision == revision)
                        {
                            leftNumCell = leftNumCell - 1;
                            workSheet.Cells[leftNumCell, 2].NumberFormat = "@";

                            if (previousAtaPage.IndexOf("─") >= 0)
                            {
                                var sNewValues = previousAtaPage.Trim('–');
                                if (sNewValues.Length == 2)
                                {
                                    workSheet.Cells[leftNumCell, 2] = sNewValues[0] + " – " + ataPage.Trim();
                                }
                            }
                            else
                            {
                                workSheet.Cells[leftNumCell, 2] = previousAtaPage + " – " + ataPage.Trim();
                                leftNumCell = leftNumCell + 1;
                            }
                        }
                        else
                        {
                            workSheet.Cells[leftNumCell, 2].NumberFormat = "@";
                            workSheet.Cells[leftNumCell, 2] = ataPage.Trim();
                            workSheet.Cells[leftNumCell, 3] = revision;

                            getValues(ataPage.Trim(), revision);

                            workSheet.Cells[leftNumCell, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                            workSheet.Cells[leftNumCell, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                            leftNumCell = leftNumCell + 1;
                        }
                    }
                    else
                    {
                        //leftNumCell = leftNumCell + 1;
                        workSheet.Cells[leftNumCell, 2].NumberFormat = "@";
                        workSheet.Cells[leftNumCell, 2] = ataPage.Trim();
                        workSheet.Cells[leftNumCell, 3] = revision;

                        getValues(ataPage.Trim(), revision);

                        workSheet.Cells[leftNumCell, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                        workSheet.Cells[leftNumCell, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        leftNumCell = leftNumCell + 1;
                    }
                }
                //}
            }
            catch (Exception ex)
            {

            }
        }

        private int RomanoToInt(string nRomano)
        {
            switch (nRomano.ToUpper())
            {
                case "I":
                    return 1;

                case "II":
                    return 2;

                case "III":
                    return 3;

                case "IV":
                    return 4;

                case "V":
                    return 5;

                case "VI":
                    return 6;

                case "VII":
                    return 7;

                case "VIII":
                    return 8;

                case "IX":
                    return 9;

                case "X":
                    return 10;

                case "XI":
                    return 11;

                case "XII":
                    return 12;

                case "XIII":
                    return 13;

                case "XIV":
                    return 14;

                case "XV":
                    return 15;
                case "XVI":
                    return 16;

                case "XVII":
                    return 17;

                case "XVIII":
                    return 18;

                case "XIX":
                    return 19;

                case "XX":
                    return 20;
                default:
                    return 9999;
            }

        }
    }
    public class TextBoxTraceListener : TraceListener
    {
        private System.Windows.Forms.TextBox _target;
        private StringSendDelegate _invokeWrite;

        public TextBoxTraceListener(System.Windows.Forms.TextBox target)
        {
            _target = target;
            _invokeWrite = new StringSendDelegate(SendString);
        }

        public override void Write(string message)
        {
            _target.Invoke(_invokeWrite, new object[] { message });
        }

        public override void WriteLine(string message)
        {
            _target.Invoke(_invokeWrite, new object[]
                { message + Environment.NewLine });
        }

        private delegate void StringSendDelegate(string message);
        private void SendString(string message)
        {
            // No need to lock text box as this function will only 
            // ever be executed from the UI thread
            _target.Text += message;
        }
    }
}
