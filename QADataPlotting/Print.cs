using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.SqlClient;
using Aspose.Cells;

namespace QADataPlotting
{
    public partial class Form2 : Form
    {
        private QAPlotting _QAplot;
        public Form2(QAPlotting QAplot)
        {
            InitializeComponent();
            _QAplot = QAplot;
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Fixed3D;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            ConvertAndShowExcelToPdf();
            /*
            string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Form2.xlsx");
            Excel.Application excelApp = new Excel.Application();

            // Open Excel workbook
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            string pngFolderPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            
            string QANumber = _QAplot.TextBoxValue;
            if (QANumber == null || QANumber == "")
            {
                QANumber = "null";
            }
            string pdfFileName = @""+ QANumber +".pdf";
            string pdfFilePath = Path.Combine(pngFolderPath, pdfFileName);

            // Get the first worksheet
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            
            // Set PageSetup properties for landscape orientation
            worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4; // Set ukuran kertas ke A4
            worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape; // Set orientasi ke Landscape
            worksheet.PageSetup.LeftMargin = excelApp.InchesToPoints(0.0); // Set margin kiri
            worksheet.PageSetup.RightMargin = excelApp.InchesToPoints(0.0); // Set margin kanan
            worksheet.PageSetup.TopMargin = excelApp.InchesToPoints(0.0); // Set margin atas
            worksheet.PageSetup.BottomMargin = excelApp.InchesToPoints(0.0); // Set margin bawah
            worksheet.PageSetup.FitToPagesWide = 1; // Fit to one page wide
            worksheet.PageSetup.FitToPagesTall = 1; // Fit to one page tall

            worksheet.PageSetup.PrintArea = worksheet.UsedRange.Address;
            foreach (Excel.Shape shape in worksheet.Shapes)
            {
                // Ensure the shape is visible and within the print area
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                shape.Placement = Excel.XlPlacement.xlMoveAndSize;
            }
            worksheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFilePath);
            
            workbook.Save();
            workbook.Close();
            
            webBrowser1.Navigate(pdfFilePath);*/
        }

        private string pdfFilePath;
        private void ConvertAndShowExcelToPdf()
        {
            string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FormQA.xlsx");
            Excel.Application excelApp = new Excel.Application();

            // Open Excel workbook
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            string QANumber = "null"; // Assign a default value for demonstration
            pdfFilePath = Path.Combine(Path.GetTempPath(), QANumber + ".pdf"); // Use temp path for temporary file

            // Get the first worksheet
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            // Set PageSetup properties for landscape orientation
            worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4; // Set ukuran kertas ke A4
            worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape; // Set orientasi ke Landscape
            worksheet.PageSetup.LeftMargin = excelApp.InchesToPoints(0.0); // Set margin kiri
            worksheet.PageSetup.RightMargin = excelApp.InchesToPoints(0.0); // Set margin kanan
            worksheet.PageSetup.TopMargin = excelApp.InchesToPoints(0.0); // Set margin atas
            worksheet.PageSetup.BottomMargin = excelApp.InchesToPoints(0.0); // Set margin bawah
            worksheet.PageSetup.FitToPagesWide = 1; // Fit to one page wide
            worksheet.PageSetup.FitToPagesTall = 1; // Fit to one page tall

            worksheet.PageSetup.PrintArea = worksheet.UsedRange.Address;
            foreach (Excel.Shape shape in worksheet.Shapes)
            {
                // Ensure the shape is visible and within the print area
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                shape.Placement = Excel.XlPlacement.xlMoveAndSize;
            }

            // Export to PDF
            worksheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFilePath);

            workbook.Close(false);
            excelApp.Quit();

            // Display the PDF in WebBrowser control
            webBrowser1.Navigate(pdfFilePath);
        }

        private void bt_Save_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                string QANumber = _QAplot.ValueQANumber;
                saveFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
                saveFileDialog.Title = "Save PDF File";
                //saveFileDialog.FileName = QANumber + ".pdf"; // Set default file name
                if (_QAplot.ValueRejectRes == 0 && _QAplot.ValueRejectOff == 0 && _QAplot.ValueRejectMat == 0 && _QAplot.ValueRejectNoi == 0)
                {
                    saveFileDialog.FileName = QANumber + ".pdf"; // Set default file name
                }
                else
                {
                    saveFileDialog.FileName = QANumber + "_reject.pdf"; // Set default file name
                }

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = saveFileDialog.FileName;
                    File.Copy(pdfFilePath, selectedPath, true); // Save the temporary PDF to the selected location
                    MessageBox.Show("File saved successfully!", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }
}
