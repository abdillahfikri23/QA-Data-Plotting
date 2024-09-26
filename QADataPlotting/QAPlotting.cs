using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace QADataPlotting
{
    public partial class QAPlotting : Form
    {
        string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FormQA.xlsx");
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NA.png");
        string imageacc = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "acc.png");
        string stampacc = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pass.png");
        string stamprej = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "reject.png");

        string mechanicalInspectStatus;

        // Deklarasikan jumlah data yang ditolak
        int rejectedResponse = 0;
        int rejectedOffset = 0;
        int rejectedMatch = 0;
        int rejectedNoise = 0;
        public double totalReject;
        int reTest = 0;
        int reTestnoi = 0;

        public double specMinRespone;
        public double specMinOffset;
        public double specMinMatch;
        public double specMinNoise;
        public double specMaxRespone;
        public double specMaxOffset;
        public double specMaxMatch;
        public double specMaxNoise;

        public double earlyRespone;
        public double earlyOffset;
        public double earlyMatch;
        public double earlyNoise;
        public bool earlySpec = true;

        public string tester;
        public string type;
        public string quality;
        public string serialNumber;
        public string QANumber;
        public string DeviceCode;
        public string inspectedBy;
        public string Res;
        public string Off;
        public string Mat;
        public string Noi;
        public string signNoi;

        public double valueRes;
        public double valueOff;
        public double valueMat;
        public double valueNoi;
        public int samplingSize;
        public int samplingSizeMax = 125;
        public int goodValueRes;
        public int goodValueOff;
        public int goodValueMat;
        public int goodValueNoi;

        //Response
        double cellValueRes1;
        double cellValueRes2;
        double cellValueRes3;
        double cellValueRes4;
        double cellValueRes5;
        double cellValueRes6;
        double cellValueRes7;
        double cellValueRes8;
        double cellValueRes9;
        double cellValueRes10;
        double cellValueRes11;
        double cellValueRes12;
        double cellValueRes13;
        double cellValueRes14;
        double cellValueRes15;
        double cellValueRes16;
        double cellValueRes17;
        double cellValueRes18;
        double cellValueRes19;
        double cellValueRes20;
        double cellValueRes21;
        double cellValueRes22;
        double cellValueRes23;
        double cellValueRes24;
        double cellValueRes25;
        double cellValueRes26;

        //Match
        double cellValueMat1;
        double cellValueMat2;
        double cellValueMat3;
        double cellValueMat4;
        double cellValueMat5;
        double cellValueMat6;
        double cellValueMat7;
        double cellValueMat8;
        double cellValueMat9;
        double cellValueMat10;
        double cellValueMat11;
        double cellValueMat12;
        double cellValueMat13;
        double cellValueMat14;
        double cellValueMat15;
        double cellValueMat16;
        double cellValueMat17;
        double cellValueMat18;
        double cellValueMat19;
        double cellValueMat20;
        double cellValueMat21;
        double cellValueMat22;

        //Offset
        double cellValueOff1;
        double cellValueOff2;
        double cellValueOff3;
        double cellValueOff4;
        double cellValueOff5;
        double cellValueOff6;
        double cellValueOff7;
        double cellValueOff8;
        double cellValueOff9;
        double cellValueOff10;
        double cellValueOff11;
        double cellValueOff12;
        double cellValueOff13;
        double cellValueOff14;
        double cellValueOff15;
        double cellValueOff16;
        double cellValueOff17;
        double cellValueOff18;
        double cellValueOff19;
        double cellValueOff20;
        double cellValueOff21;
        double cellValueOff22;

        double cellValueNoi1;
        double cellValueNoi2;
        double cellValueNoi3;
        double cellValueNoi4;
        double cellValueNoi5;
        double cellValueNoi6;
        double cellValueNoi7;
        double cellValueNoi8;
        double cellValueNoi9;
        double cellValueNoi10;
        double cellValueNoi11;
        double cellValueNoi12;
        double cellValueNoi13;
        double cellValueNoi14;
        double cellValueNoi15;
        double cellValueNoi16;
        double cellValueNoi17;
        double cellValueNoi18;
        double cellValueNoi19;
        double cellValueNoi20;

        public QAPlotting()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Fixed3D;
            cb_QANo.SelectedIndexChanged += new EventHandler(cb_QANo_SelectedIndexChanged);
        }

        private void QAPlotting_Load(object sender, EventArgs e)
        {
            clearDataEarly();
            LoadQANumberData();
            DisplayExcel("FORM");
            this.Menu = new MainMenu();

            //MENU
            MenuItem item1 = new MenuItem("MENU");
            this.Menu.MenuItems.Add(item1);
            item1.MenuItems.Add("DATA_SQL", new EventHandler(DataSql_Click));

            Tester();
        }

        private void LoadQANumberData()
        {
            DateTime selectedDate = dateTimePicker1.Value.Date;
            string formattedDate = selectedDate.ToString("M/d/yyyy");
            string connectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123";
            //string query = "SELECT QANo FROM dmQABLDetail WHERE ProdSubmitDateActual = '" + formattedDate + "' AND VisualInspectStatus = 'CLOSE' AND ElectricalInspectStatus = 'PASS' AND MechanicalInspectStatus = 'PASS'"; // Ganti ColumnName dan TableName dengan nama kolom dan tabel yang sesuai
            string query = "SELECT QANo FROM dmQABLDetail WHERE ProdSubmitDateActual = '" + formattedDate + "' AND ElectricalInspectStatus = 'PASS'"; // Ganti ColumnName dan TableName dengan nama kolom dan tabel yang sesuai
            //string query = "SELECT QANo FROM dmQABLDetail WHERE ProdSubmitDateActual = '" + formattedDate + "' AND MechanicalInspectStatus = 'PASS'"; // Ganti ColumnName dan TableName dengan nama kolom dan tabel yang sesuai

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    cb_QANo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    cb_QANo.AutoCompleteSource = AutoCompleteSource.ListItems;
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            cb_QANo.Items.Add(reader["QANo"].ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void DataSql_Click(object sender, EventArgs e)
        {
            Form1 dataValue = new Form1();
            dataValue.Show();
        }

        public string ValueQANumber
        {
            get { return cb_QANo.Text; }
            set { cb_QANo.Text = value; }
        }

        public int ValueRejectRes
        {
            get { return rejectedResponse; }
            set { rejectedResponse = value; }
        }

        public int ValueRejectOff
        {
            get { return rejectedOffset; }
            set { rejectedResponse = value; }
        }

        public int ValueRejectMat
        {
            get { return rejectedMatch; }
            set { rejectedResponse = value; }
        }

        public int ValueRejectNoi
        {
            get { return rejectedNoise; }
            set { rejectedNoise = value; }
        }

        public Button Button
        {
            get { return this.BT_SelectFile; }
        }

        public void DisplayExcel(string sheetName)
        {
            // Buka aplikasi Excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            // Buka file Excel
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(excelFilePath);

            // Cari sheet dengan nama yang sesuai
            Excel.Worksheet excelWorksheet = null;
            foreach (Excel.Worksheet sheet in excelWorkbook.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    excelWorksheet = sheet;
                    break;
                }
            }

            // Jika sheet tidak ditemukan, keluarkan pesan kesalahan
            if (excelWorksheet == null)
            {
                MessageBox.Show("Sheet dengan nama " + sheetName + " tidak ditemukan!");
                excelWorkbook.Close(false);
                excelApp.Quit();
                releaseObject(excelWorkbook);
                releaseObject(excelApp);
                return;
            }

            // Simpan file Excel ke tempat sementara (bisa juga disimpan di temp path)
            string tempFilePath = System.IO.Path.GetTempFileName() + ".html";
            excelWorksheet.SaveAs(tempFilePath, Excel.XlFileFormat.xlHtml);

            // Tampilkan file HTML di WebBrowser
            webBrowser1.Navigate(tempFilePath);

            excelWorkbook.Close(false);
            excelApp.Quit();

            // Bersihkan objek Excel dari memori
            releaseObject(excelWorksheet);
            releaseObject(excelWorkbook);
            releaseObject(excelApp);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        static void PlaceImageInCell(Excel.Worksheet worksheet, string imagePath, int row, int column, float width, float height)
        {
            // Konversi koordinat sel ke piksel
            float left = (float)((double)worksheet.Cells[1, column].Left);
            float top = (float)((double)worksheet.Cells[row, 1].Top);

            // Tambahkan gambar
            Excel.Pictures pictures = worksheet.Pictures() as Excel.Pictures;
            Excel.Picture picture = pictures.Insert(imagePath) as Excel.Picture;
            
            // Atur posisi dan ukuran gambar
            picture.Left = left;
            picture.Top = top;
            picture.Width = width;
            picture.Height = height;

            // Pastikan gambar berada di belakang teks
            picture.Placement = Excel.XlPlacement.xlMoveAndSize;
        }

        static void PlaceAcc(Excel.Worksheet worksheet, string imageacc, int left, int top, int width, int height)
        {
            // Menambahkan gambar ke worksheet
            Excel.Pictures pictures = (Excel.Pictures)worksheet.Pictures();
            Excel.Picture picture = pictures.Insert(imageacc) as Excel.Picture;

            // Mengatur posisi gambar menggunakan piksel
            picture.Left = left;
            picture.Top = top;

            // Mengatur ukuran gambar menggunakan piksel
            picture.Width = width;
            picture.Height = height;

            // Pastikan gambar berada di belakang teks
            picture.Placement = Excel.XlPlacement.xlMoveAndSize;
        }

        static void PlaceinCell(Excel.Worksheet worksheet, string imageacc, int leftColumn, int topRow, int widthColumns, int heightRows)
        {
            // Menambahkan gambar ke worksheet
            Excel.Pictures pictures = (Excel.Pictures)worksheet.Pictures();
            Excel.Picture picture = pictures.Insert(imageacc) as Excel.Picture;

            // Mengatur posisi gambar menggunakan koordinat sel
            Excel.Range cell = worksheet.Cells[topRow, leftColumn];
            picture.Left = (float)cell.Left;
            picture.Top = (float)cell.Top;

            // Mengatur ukuran gambar menggunakan lebar dan tinggi sel
            picture.Width = widthColumns;
            picture.Height = heightRows;

            // Pastikan gambar berada di belakang teks
            picture.Placement = Excel.XlPlacement.xlMoveAndSize;
        }

        static void PlaceAcc2(Excel.Worksheet worksheet, string imageacc, int startRow, int startColumn, int colSpan)
        {
            // Mendapatkan referensi ke sel-sel yang dituju
            Excel.Range startCell = worksheet.Cells[startRow, startColumn] as Excel.Range;
            Excel.Range endCell = worksheet.Cells[startRow, startColumn + colSpan - 1] as Excel.Range;

            // Menggabungkan sel untuk mendapatkan area yang mencakup beberapa kolom
            Excel.Range targetRange = worksheet.Range[startCell, endCell];

            // Menghitung posisi dan ukuran area gabungan
            float cellLeft = (float)targetRange.Left;
            float cellTop = (float)targetRange.Top;
            float cellWidth = (float)targetRange.Width;
            float cellHeight = (float)targetRange.Height;

            // Ukuran gambar yang diinginkan (sesuaikan ini dengan ukuran gambar yang diinginkan)
            float imageWidth = cellWidth;  // Misalnya, ukuran gambar sesuai dengan lebar area gabungan
            float imageHeight = cellHeight; // Misalnya, ukuran gambar sesuai dengan tinggi sel

            // Menghitung posisi gambar agar berada di tengah area gabungan
            float imageLeft = cellLeft + (cellWidth - imageWidth) / 2;
            float imageTop = cellTop + (cellHeight - imageHeight) / 2;

            // Menambahkan gambar dan mengatur posisinya sesuai area gabungan
            Excel.Shape picture = worksheet.Shapes.AddPicture(
                imageacc,
                Microsoft.Office.Core.MsoTriState.msoFalse,  // LinkToFile
                Microsoft.Office.Core.MsoTriState.msoCTrue,  // SaveWithDocument
                imageLeft, imageTop, imageWidth, imageHeight);

            // Mengatur gambar berada di belakang teks
            picture.Placement = Excel.XlPlacement.xlMoveAndSize;
        }

        void clearDataEarly()
        {
            // Inisialisasi aplikasi Excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = workbook.Sheets[1]; // Menggunakan worksheet pertama

            try
            {
                DeleteImages(worksheet);
                DeleteCell(worksheet, 43, 72);
                DeleteCell(worksheet, 44, 72);
                DeleteCell(worksheet, 43, 60);
                DeleteCell(worksheet, 44, 60);
                DeleteCell(worksheet, 36, 23);
                DeleteCell(worksheet, 37, 23);
                DeleteCell(worksheet, 38, 23);
                DeleteCell(worksheet, 39, 23);
                DeleteCell(worksheet, 40, 23);
                DeleteCell(worksheet, 36, 11);
                DeleteCell(worksheet, 37, 11);
                DeleteCell(worksheet, 38, 11);
                DeleteCell(worksheet, 39, 11);
                DeleteCell(worksheet, 40, 11);
                DeleteCell(worksheet, 41, 23);
                DeleteCell(worksheet, 36, 2);
                DeleteCell(worksheet, 36, 11);
                DeleteCell(worksheet, 37, 2);
                DeleteCell(worksheet, 37, 11);

                DeleteCell(worksheet, 2, 83);
                DeleteCell(worksheet, 3, 83);
                DeleteCell(worksheet, 5, 10);
                DeleteCell(worksheet, 6, 10);
                DeleteCell(worksheet, 7, 10);
                DeleteCell(worksheet, 8, 10);
                DeleteCell(worksheet, 5, 42);
                DeleteCell(worksheet, 7, 42);
                DeleteCell(worksheet, 8, 42);
                DeleteCell(worksheet, 5, 79);
                DeleteCell(worksheet, 6, 79);
                DeleteCell(worksheet, 7, 79);
                DeleteCell(worksheet, 8, 79);
                DeleteCell(worksheet, 32, 45);
                DeleteCell(worksheet, 6, 42);
                DeleteCell(worksheet, 36, 65);
                DeleteCell(worksheet, 36, 87);
                DeleteCell(worksheet, 37, 87);
                DeleteCell(worksheet, 32, 18);

                //Menghapus data sesuai range
                DeleteRangeData(worksheet, 15, 3, 15, 150);
                DeleteRangeData(worksheet, 16, 3, 16, 150);
                DeleteRangeData(worksheet, 17, 3, 17, 150);
                DeleteRangeData(worksheet, 18, 3, 18, 150);
                DeleteRangeData(worksheet, 19, 3, 19, 150);
                DeleteRangeData(worksheet, 20, 3, 20, 150);
                DeleteRangeData(worksheet, 21, 3, 21, 150);
                DeleteRangeData(worksheet, 22, 3, 22, 150);
                DeleteRangeData(worksheet, 23, 3, 23, 150);
                DeleteRangeData(worksheet, 24, 3, 24, 150);
                DeleteRangeData(worksheet, 25, 3, 25, 150);
                DeleteRangeData(worksheet, 26, 3, 26, 150);
                DeleteRangeData(worksheet, 27, 3, 27, 150);
                DeleteRangeData(worksheet, 28, 3, 28, 150);
                DeleteRangeData(worksheet, 29, 3, 29, 150);

                // Simpan perubahan
                workbook.Save();

                DisplayExcel("FORM");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // Tutup aplikasi Excel
                workbook.Close();
                excelApp.Quit();
            }
        }

        static void DeleteImages(Excel.Worksheet worksheet)
        {
            // Dapatkan koleksi gambar di lembar kerja
            Excel.Pictures pictures = worksheet.Pictures() as Excel.Pictures;

            // Hapus semua gambar satu per satu
            for (int i = pictures.Count; i >= 1; i--)
            {
                pictures.Item(i).Delete();
            }
        }

        static void DeleteImagesName(Excel.Worksheet worksheet, string imageName)
        {
            try
            {
                // Iterasi melalui semua shape di worksheet
                foreach (Excel.Shape shape in worksheet.Shapes)
                {
                    if (shape.Name == imageName)
                    {
                        shape.Delete();
                        break; // Hentikan iterasi setelah gambar ditemukan dan dihapus
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Terjadi kesalahan saat menghapus gambar: " + ex.Message);
            }
        }

        static void DeleteImagesRange(Excel.Worksheet worksheet, string rangeToDelete)
        {
            try
            {
                // Tentukan range yang ingin Anda gunakan
                Excel.Range range = worksheet.Range[rangeToDelete];

                // Dapatkan semua gambar dalam worksheet
                Excel.Pictures pictures = worksheet.Pictures() as Excel.Pictures;

                foreach (Excel.Picture picture in pictures)
                {
                    // Cek apakah gambar ada di dalam range yang ditentukan
                    Excel.Range pictureRange = worksheet.Range[picture.TopLeftCell, picture.BottomRightCell];
                    if (IsRangeIntersect(pictureRange, range))
                    {
                        // Hapus gambar
                        picture.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to delete pictures in range: " + ex.Message);
            }
        }

        static bool IsRangeIntersect(Excel.Range range1, Excel.Range range2)
        {
            return !(range1.Row > range2.Row + range2.Rows.Count - 1 ||
                     range1.Row + range1.Rows.Count - 1 < range2.Row ||
                     range1.Column > range2.Column + range2.Columns.Count - 1 ||
                     range1.Column + range1.Columns.Count - 1 < range2.Column);
        }

        void clearDataAll()
        {
            // Inisialisasi aplikasi Excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = workbook.Sheets[1]; // Menggunakan worksheet pertama

            try
            {
                DeleteImages(worksheet);
                DeleteCell(worksheet, 43, 72);
                DeleteCell(worksheet, 44, 72);
                DeleteCell(worksheet, 43, 60);
                DeleteCell(worksheet, 44, 60);
                DeleteCell(worksheet, 36, 23);
                DeleteCell(worksheet, 37, 23);
                DeleteCell(worksheet, 38, 23);
                DeleteCell(worksheet, 39, 23);
                DeleteCell(worksheet, 40, 23);
                DeleteCell(worksheet, 36, 11);
                DeleteCell(worksheet, 37, 11);
                DeleteCell(worksheet, 38, 11);
                DeleteCell(worksheet, 39, 11);
                DeleteCell(worksheet, 40, 11);
                DeleteCell(worksheet, 41, 23);
                DeleteCell(worksheet, 36, 2);
                DeleteCell(worksheet, 36, 11);
                DeleteCell(worksheet, 37, 2);
                DeleteCell(worksheet, 37, 11);

                DeleteCell(worksheet, 2, 83);
                DeleteCell(worksheet, 3, 83);
                DeleteCell(worksheet, 5, 10);
                DeleteCell(worksheet, 6, 10);
                DeleteCell(worksheet, 7, 10);
                DeleteCell(worksheet, 8, 10);
                DeleteCell(worksheet, 5, 42);
                DeleteCell(worksheet, 7, 42);
                DeleteCell(worksheet, 8, 42);
                DeleteCell(worksheet, 5, 79);
                DeleteCell(worksheet, 6, 79);
                DeleteCell(worksheet, 7, 79);
                DeleteCell(worksheet, 8, 79);
                DeleteCell(worksheet, 32, 45);
                DeleteCell(worksheet, 6, 42);
                DeleteCell(worksheet, 36, 65);
                DeleteCell(worksheet, 36, 87);
                DeleteCell(worksheet, 37, 87);
                DeleteCell(worksheet, 32, 18);

                //Menghapus data sesuai range
                DeleteRangeData(worksheet, 15, 3, 15, 150);
                DeleteRangeData(worksheet, 16, 3, 16, 150);
                DeleteRangeData(worksheet, 17, 3, 17, 150);
                DeleteRangeData(worksheet, 18, 3, 18, 150);
                DeleteRangeData(worksheet, 19, 3, 19, 150);
                DeleteRangeData(worksheet, 20, 3, 20, 150);
                DeleteRangeData(worksheet, 21, 3, 21, 150);
                DeleteRangeData(worksheet, 22, 3, 22, 150);
                DeleteRangeData(worksheet, 23, 3, 23, 150);
                DeleteRangeData(worksheet, 24, 3, 24, 150);
                DeleteRangeData(worksheet, 25, 3, 25, 150);
                DeleteRangeData(worksheet, 26, 3, 26, 150);
                DeleteRangeData(worksheet, 27, 3, 27, 150);
                DeleteRangeData(worksheet, 28, 3, 28, 150);
                DeleteRangeData(worksheet, 29, 3, 29, 150);

                // Simpan perubahan
                workbook.Save();

                MessageBox.Show("Data is successfull deleted");
                DisplayExcel("FORM");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // Tutup aplikasi Excel
                workbook.Close();
                excelApp.Quit();
            }
        }

        void clearData()
        {
            // Inisialisasi aplikasi Excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = workbook.Sheets[1];

            try
            {
                DeleteImages(worksheet);
                //Menghapus data sesuai range
                DeleteRangeData(worksheet, 6, 42, 6, 60);
                DeleteRangeData(worksheet, 15, 3, 15, 74);
                DeleteRangeData(worksheet, 16, 3, 16, 74);
                DeleteRangeData(worksheet, 17, 3, 17, 74);
                DeleteRangeData(worksheet, 18, 3, 18, 74);
                DeleteRangeData(worksheet, 19, 3, 19, 74);
                DeleteRangeData(worksheet, 20, 3, 20, 74);
                DeleteRangeData(worksheet, 21, 3, 21, 74);
                DeleteRangeData(worksheet, 22, 3, 22, 74);
                DeleteRangeData(worksheet, 23, 3, 23, 74);
                DeleteRangeData(worksheet, 24, 3, 24, 74);
                DeleteRangeData(worksheet, 25, 3, 25, 74);
                DeleteRangeData(worksheet, 26, 3, 26, 74);
                DeleteRangeData(worksheet, 27, 3, 27, 74);
                DeleteRangeData(worksheet, 28, 3, 28, 74);
                DeleteCell(worksheet, 36, 11);
                DeleteCell(worksheet, 37, 11);
                DeleteCell(worksheet, 38, 11);
                DeleteCell(worksheet, 39, 11);
                //DeleteCell(worksheet, 40, 11);
                DeleteCell(worksheet, 36, 23);
                DeleteCell(worksheet, 37, 23);
                DeleteCell(worksheet, 38, 23);
                DeleteCell(worksheet, 39, 23);
                DeleteCell(worksheet, 41, 23);

                // Simpan perubahan
                workbook.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // Tutup aplikasi Excel
                workbook.Close();
                excelApp.Quit();
            }
        }

        void clearDataNoise()
        {
            // Inisialisasi aplikasi Excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = workbook.Sheets[1]; // Menggunakan worksheet pertama

            try
            {
                //string imagename1 = "Picture 3";
                //DeleteImagesName(worksheet, imagename1);
                //string imagename2 = "Picture 4";
                //DeleteImagesName(worksheet, imagename2);

                string rangeToDelete = "A43:CS43";
                DeleteImagesRange(worksheet, rangeToDelete);

                string rangeToDelete1 = "BX14:CQ14";
                DeleteImagesRange(worksheet, rangeToDelete1);

                //Menghapus data sesuai range
                DeleteRangeData(worksheet, 15, 76, 15, 150);
                DeleteRangeData(worksheet, 16, 76, 16, 150);
                DeleteRangeData(worksheet, 17, 76, 17, 150);
                DeleteRangeData(worksheet, 18, 76, 18, 150);
                DeleteRangeData(worksheet, 19, 76, 19, 150);
                DeleteRangeData(worksheet, 20, 76, 20, 150);
                DeleteRangeData(worksheet, 21, 76, 21, 150);
                DeleteRangeData(worksheet, 22, 76, 22, 150);
                DeleteRangeData(worksheet, 23, 76, 23, 150);
                DeleteRangeData(worksheet, 24, 76, 24, 150);
                DeleteRangeData(worksheet, 25, 76, 25, 150);
                DeleteRangeData(worksheet, 26, 76, 26, 150);
                DeleteRangeData(worksheet, 27, 76, 27, 150);
                DeleteRangeData(worksheet, 28, 76, 28, 150);
                DeleteRangeData(worksheet, 29, 76, 29, 150);
                DeleteCell(worksheet, 40, 11);
                DeleteCell(worksheet, 40, 23);
                DeleteCell(worksheet, 41, 23);

                // Simpan perubahan
                workbook.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // Tutup aplikasi Excel
                workbook.Close();
                excelApp.Quit();
            }
        }

        static void DeleteCell(Excel.Worksheet worksheet, int row, int column)
        {
            Excel.Range range = (Excel.Range)worksheet.Cells[row, column];
            range.Value = "";
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
        }

        static void DeleteRangeData(Excel.Worksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            Excel.Range range = worksheet.Range[worksheet.Cells[startRow, startColumn], worksheet.Cells[endRow, endColumn]];
            range.ClearContents(); // Hapus nilai di dalam rentang sel
        }

        void plottingData()
        {
            // Koneksi ke SQL Server
            string connectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123";
            string query1 = "SELECT * FROM dmQABLHeader WHERE QANo = @QA_Number";
            string query2 = "SELECT * FROM dmQABLSerialPaper WHERE QANo = @QA_Number";
            string query3 = "SELECT * FROM TESTER_COINTRAY_DEVSPECQA WHERE DEVICE_CODE = @Device_Code";
            string query4 = "SELECT * FROM dmQABLDetail WHERE QANo = @QA_Number";
            string query5 = "SELECT * FROM dmQABLVisualMechanical WHERE QANo = @QA_Number";

            QANumber = cb_QANo.Text;
            DataTable dataTable1 = new DataTable();
            DataTable dataTable2 = new DataTable();
            DataTable dataTable3 = new DataTable();
            DataTable dataTable4 = new DataTable();
            DataTable dataTable5 = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query1, connection))
                {
                    // Parameter
                    command.Parameters.AddWithValue("@QA_Number", QANumber);
                    connection.Open();
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable1);

                        if (dataTable1.Rows.Count > 0)
                        {
                            for (int dc = 0; dc < dataTable1.Rows.Count; dc++)
                            {
                                // Pastikan kolom "DeviceCode" ada dan tidak null
                                if (dataTable1.Rows[dc]["DeviceCode"] != DBNull.Value)
                                {
                                    DeviceCode = dataTable1.Rows[dc]["DeviceCode"].ToString();
                                }
                                else
                                {
                                    MessageBox.Show("DeviceCode is null for QANumber: " + QANumber);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("No data found for QANumber: " + QANumber);
                        }
                    }
                }

                using (SqlCommand command = new SqlCommand(query2, connection))
                {
                    // Parameter
                    command.Parameters.AddWithValue("@QA_Number", QANumber);
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable2);
                    }
                }
                
                using (SqlCommand command = new SqlCommand(query3, connection))
                {
                    // Parameter
                    command.Parameters.AddWithValue("@Device_Code", DeviceCode);
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable3);

                        if (dataTable3.Rows.Count > 0)
                        {
                            specMaxRespone = Convert.ToDouble(dataTable3.Rows[0]["RESPONSE_A_MAX"]);
                            specMinRespone = Convert.ToDouble(dataTable3.Rows[0]["RESPONSE_A_MIN"]);
                            specMaxOffset = Convert.ToDouble(dataTable3.Rows[0]["OFFSET_MAX"]);
                            specMinOffset = Convert.ToDouble(dataTable3.Rows[0]["OFFSET_MIN"]);
                            //specMaxMatch = Convert.ToDouble(dataTable3.Rows[0]["MATCH_MAX"]) * 100;
                            specMaxMatch = Convert.ToDouble(dataTable3.Rows[0]["RESPONSE_A_MIN"]) * 10/100;
                            specMaxNoise = Convert.ToDouble(dataTable3.Rows[0]["NOISE_MAX"]) * 6.5;
                        }
                    }
                }

                using (SqlCommand command = new SqlCommand(query4, connection))
                {
                    // Parameter
                    command.Parameters.AddWithValue("@QA_Number", QANumber);
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable4);
                    }
                }

                using (SqlCommand command = new SqlCommand(query5, connection))
                {
                    // Parameter
                    command.Parameters.AddWithValue("@QA_Number", QANumber);
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable5);
                    }
                }
            }

            // Mengecek apakah ada baris dalam DataTable
            if (dataTable1.Rows.Count == 0)
            {
                MessageBox.Show("QA Number Not Found in dmQABLHeader Tabel!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (dataTable2.Rows.Count == 0)
            {
                MessageBox.Show("QA Number Not Found in dmQABLSerialPaper!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (dataTable3.Rows.Count == 0)
            {
                MessageBox.Show("Device Code Not Found in TESTER_COINTRAY_DEVSPECQA!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            clearData();

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            // Buka workbook yang sudah ada
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = workbook.Sheets["FORM"];

            //Show Form to Input Data of Spec
            Spec spec = new Spec(this);
            spec.AddValueRes( specMinRespone.ToString() + " < x <" + specMaxRespone.ToString());
            spec.AddValueOff(specMinOffset.ToString() + " < x < " + specMaxOffset.ToString());
            //spec.AddValueMat(" x < " + specMaxMatch.ToString()+ " %");
            spec.AddValueMat(" x < " + specMaxMatch.ToString("N2"));
            spec.AddValueNoi(" x < " + specMaxNoise.ToString());

            if (earlySpec == true)
            {
                spec.ShowDialog();

                for (int l = 0; l < dataTable1.Rows.Count; l++)
                {
                    worksheet.Cells[32, 18] = "0/1"; 
                    worksheet.Cells[2, 83] = dataTable1.Rows[l]["QAWeek"].ToString();
                    worksheet.Cells[3, 83] = dataTable1.Rows[l]["QADate"].ToString();
                    worksheet.Cells[5, 10] = CB_TesterName.Text;
                    worksheet.Cells[7, 10] = dataTable1.Rows[l]["DeviceName"].ToString();
                    worksheet.Cells[8, 10] = dataTable2.Rows[l]["Quantity"].ToString();
                    if (dataTable2.Rows.Count > 3)
                    {
                        worksheet.Cells[5, 42] = dataTable2.Rows[0]["SerialPaperNo"].ToString() + ", " + dataTable2.Rows[1]["SerialPaperNo"].ToString();
                        worksheet.Cells[6, 42] = dataTable2.Rows[2]["SerialPaperNo"].ToString() + ", " + dataTable2.Rows[3]["SerialPaperNo"].ToString();
                        int quantity1 =  Convert.ToInt16(dataTable2.Rows[0]["Quantity"].ToString());
                        int quantity2 = Convert.ToInt16(dataTable2.Rows[1]["Quantity"].ToString());
                        int quantity3 = Convert.ToInt16(dataTable2.Rows[2]["Quantity"].ToString());
                        int quantity4 = Convert.ToInt16(dataTable2.Rows[3]["Quantity"].ToString());
                        worksheet.Cells[8, 10] = quantity1 + quantity2 + quantity3 + quantity4;
                    }
                    else if (dataTable2.Rows.Count > 2)
                    {
                        worksheet.Cells[5, 42] = dataTable2.Rows[0]["SerialPaperNo"].ToString() + ", " + dataTable2.Rows[1]["SerialPaperNo"].ToString();
                        worksheet.Cells[6, 42] = dataTable2.Rows[2]["SerialPaperNo"].ToString();
                        int quantity1 = Convert.ToInt16(dataTable2.Rows[0]["Quantity"].ToString());
                        int quantity2 = Convert.ToInt16(dataTable2.Rows[1]["Quantity"].ToString());
                        int quantity3 = Convert.ToInt16(dataTable2.Rows[2]["Quantity"].ToString());
                        worksheet.Cells[8, 10] = quantity1 + quantity2 + quantity3;
                    }
                    else if (dataTable2.Rows.Count > 1)
                    {
                        worksheet.Cells[5, 42] = dataTable2.Rows[0]["SerialPaperNo"].ToString() + ", " + dataTable2.Rows[1]["SerialPaperNo"].ToString();
                        int quantity1 = Convert.ToInt16(dataTable2.Rows[0]["Quantity"].ToString());
                        int quantity2 = Convert.ToInt16(dataTable2.Rows[1]["Quantity"].ToString());
                        worksheet.Cells[8, 10] = quantity1 + quantity2;
                    }
                    else if (dataTable2.Rows.Count == 1)
                    {
                        worksheet.Cells[5, 42] = dataTable2.Rows[0]["SerialPaperNo"].ToString();
                        worksheet.Cells[8, 10] = dataTable2.Rows[0]["Quantity"].ToString();
                    }
                    worksheet.Cells[7, 42] = dataTable1.Rows[l]["QANo"].ToString();
                    worksheet.Cells[8, 42] = TB_InspectedBy.Text;
                    worksheet.Cells[36, 65] = dataTable4.Rows[l]["ProdDateCode"].ToString();

                    DateTime prodSubmitDate = Convert.ToDateTime(dataTable4.Rows[l]["ProdSubmitDate"]);
                    string prodSubmitTime = prodSubmitDate.ToString("h:mm:ss tt");
                    worksheet.Cells[36, 87] = prodSubmitTime;

                    DateTime electricalInspectOut = Convert.ToDateTime(dataTable4.Rows[l]["ElectricalInspectOut"]);
                    string electricaltime = electricalInspectOut.ToString("h:mm:ss tt");
                    worksheet.Cells[37, 87] = electricaltime;

                    //worksheet.Cells[32, 45].Value = samplingSize;
                    worksheet.Cells[43, 60].Value = DateTime.Now;
                    worksheet.Cells[44, 60].Value = TB_InspectedBy.Text;

                    if (dataTable5.Rows.Count > 0)
                    {
                        mechanicalInspectStatus = dataTable5.Rows[l]["MechanicalInspectStatus"].ToString();
                        worksheet.Cells[43, 72] = dataTable5.Rows[l]["MechanicalInspectDate"].ToString();
                        worksheet.Cells[44, 72] = dataTable5.Rows[l]["MechanicalInspectName"].ToString();
                    }
                }
                earlySpec = false;
            }

            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Filter jenis file yang dapat dipilih (opsional)
            openFileDialog.Filter = "All files (*.*)|*.*|Text files (*.txt)|*.txt";

            // Set opsi lainnya (opsional)
            openFileDialog.Title = "Select a file";
            openFileDialog.InitialDirectory = @"C:\"; // Direktori awal yang ditampilkan
            openFileDialog.Multiselect = false; // Memungkinkan pemilihan beberapa file

            // Tampilkan dialog dan dapatkan hasil
            DialogResult result = openFileDialog.ShowDialog();

            // Jika pengguna memilih file dan menekan tombol "OK"
            if (result == DialogResult.OK)
            {
                // Dapatkan path file yang dipilih
                string selectedFilePath = openFileDialog.FileName;
                try
                {
                    // Baca isi file teks
                    string[] rows = File.ReadAllLines(selectedFilePath);

                    // Jika ada baris dalam file
                    if (rows.Length > 0)
                    {
                        // List untuk menyimpan data setiap kolom
                        List<double> responseData = new List<double>();
                        List<double> offsetData = new List<double>();
                        List<double> matchData = new List<double>();

                        // Membaca file
                        foreach (string row in rows)
                        {
                            string[] numColumns = row.Split('\t');
                            double dataRes;
                            double dataOff;
                            double dataMat;

                            if (numColumns.Length > 22) // Pastikan ada setidaknya 23 kolom
                            {
                                if (double.TryParse(numColumns[22], out dataRes))
                                {
                                    if (dataRes != 0 && dataRes < specMaxRespone)
                                    {
                                        double dataresfix = Math.Round(dataRes, 2, MidpointRounding.AwayFromZero);
                                        responseData.Add(dataresfix);
                                    }
                                }

                                if (double.TryParse(numColumns[18], out dataOff))
                                {
                                    if (dataOff != 0 && dataOff < specMaxOffset)
                                    {
                                        double dataofffix = Math.Round(dataOff, 2, MidpointRounding.AwayFromZero);
                                        offsetData.Add(dataofffix);
                                    }
                                }

                                if (double.TryParse(numColumns[24], out dataMat))
                                {
                                    if (dataMat != 0 && dataMat < specMaxMatch)
                                    {
                                        double datamatfix = Math.Round(dataMat, 2, MidpointRounding.AwayFromZero);
                                        matchData.Add(datamatfix);
                                    }
                                }
                            }
                        }

                        //RESPONSE
                        if (responseData.Count > 0)
                        {
                            Res = "Avail";
                            double averageRes = responseData.Average();
                            double averes = Math.Round(averageRes, 1, MidpointRounding.AwayFromZero);
                            earlyRespone = averes - 1.3;
                        }
                        else
                        {
                            Res = "NA";
                            PlaceImageInCell(worksheet, imagePath, 14, 4, 170, 200);
                        }

                        //OFFSET
                        if (offsetData.Count > 0)
                        {
                            Off = "Avail";
                            double averageOff = offsetData.Average();
                            double aveoff = Math.Round(averageOff, 1, MidpointRounding.AwayFromZero);
                            earlyOffset = aveoff - 1.1;
                        }
                        else
                        {
                            Off = "NA";
                            PlaceImageInCell(worksheet, imagePath, 14, 31, 170, 200);
                        }

                        //MATCH
                        if (matchData.Count > 0)
                        {
                            Mat = "Avail";
                            double averageMat = matchData.Average();
                            double avemat = Math.Round(averageMat, 2, MidpointRounding.AwayFromZero);
                            //earlyMatch = avemat - 0.11;
                        }
                        else
                        {
                            Mat = "NA";
                            PlaceImageInCell(worksheet, imagePath, 14, 54, 170, 200);
                        }


                        //Show Value to Form Plotting
                        //RESPONSE
                        if (Res == "Avail")
                        {
                            worksheet.Cells[5, 79] = specMinRespone + " < X < " + specMaxRespone + " (kV/W)";
                        }
                        else
                        {
                            worksheet.Cells[5, 79] = "NA";
                        }

                        //OFFSET
                        if (Off == "Avail")
                        {
                            worksheet.Cells[6, 79] = specMinOffset + " < X < " + specMaxOffset + " (V)";
                        }
                        else
                        {
                            worksheet.Cells[6, 79] = "NA";
                        }

                        //MATCH
                        if (Mat == "Avail")
                        {
                            worksheet.Cells[7, 79] = "X < " + specMaxMatch.ToString("N2") + " (kV/W)";
                            //worksheet.Cells[7, 79] = "X < 10 % (kV/W)";
                            specMinMatch = earlyMatch;
                        }
                        else
                        {
                            worksheet.Cells[7, 79] = "NA";
                        }

                        // Membaca file
                        foreach (string row in rows)
                        {
                            string[] numColumns = row.Split('\t');
                            double dataRes;
                            double dataOff;
                            double dataMat;
                            string errorCode = numColumns[15];
                            if (numColumns.Length > 22) // Pastikan ada setidaknya 23 kolom
                            {
                                if (reTest == 0)
                                {
                                    if (Res == "Avail")
                                    {
                                        if (double.TryParse(numColumns[22], out dataRes))
                                        {
                                            if (dataRes != 0 && dataRes < specMaxRespone)
                                            {
                                                double dataresfix = Math.Round(dataRes, 2, MidpointRounding.AwayFromZero);
                                                responseData.Add(dataresfix);
                                            }

                                            if (errorCode != "ENT")
                                            {
                                                double minValue = specMinRespone;
                                                double maxValue = specMaxRespone;
                                                valueRes = Math.Floor(dataRes * 10) / 10;
                                                if (valueRes >= minValue && valueRes <= maxValue && errorCode == "E00")
                                                {

                                                }
                                                else
                                                {
                                                    rejectedResponse++;
                                                    if (rejectedResponse != 0)
                                                    {
                                                        worksheet.Cells[37, 11] = "Response";
                                                    }
                                                    worksheet.Cells[37, 23] = rejectedResponse;
                                                }
                                            }
                                        }
                                    }

                                    if (Off == "Avail")
                                    {
                                        if (double.TryParse(numColumns[18], out dataOff))
                                        {
                                            if (dataOff != 0 && dataOff < specMaxOffset)
                                            {
                                                double dataofffix = Math.Round(dataOff, 2, MidpointRounding.AwayFromZero);
                                                offsetData.Add(dataofffix);
                                            }

                                            if (errorCode != "ENT")
                                            {
                                                double minValue = specMinOffset;
                                                double maxValue = specMaxOffset;
                                                valueOff = Math.Round(dataOff, 2, MidpointRounding.AwayFromZero);
                                                if (valueOff >= minValue && valueOff <= maxValue && errorCode == "E00")
                                                {

                                                }
                                                else
                                                {
                                                    rejectedOffset++;
                                                    if (rejectedOffset != 0)
                                                    {
                                                        worksheet.Cells[38, 11] = "Offset";
                                                    }
                                                    worksheet.Cells[38, 23] = rejectedOffset;
                                                }
                                            }
                                        }
                                    }

                                    if (Mat == "Avail")
                                    {
                                        if (double.TryParse(numColumns[24], out dataMat))
                                        {
                                            if (dataMat != 0 && dataMat < specMaxMatch)
                                            {
                                                double datamatfix = Math.Round(dataMat, 2, MidpointRounding.AwayFromZero);
                                                matchData.Add(datamatfix);
                                            }

                                            if (errorCode != "ENT")
                                            {
                                                double minValue = specMinMatch;
                                                //double maxValue = specMaxMatch;
                                                double maxValue = valueRes * 10 / 100;
                                                valueMat = Math.Round(dataMat, 2, MidpointRounding.AwayFromZero);
                                                if (valueMat >= minValue && valueMat <= maxValue && errorCode == "E00")
                                                {

                                                }
                                                else
                                                {
                                                    rejectedMatch++;
                                                    if (rejectedMatch != 0)
                                                    {
                                                        worksheet.Cells[39, 11] = "Match";
                                                    }
                                                    worksheet.Cells[39, 23] = rejectedMatch;
                                                }
                                            }
                                        }
                                    }
                                    if (errorCode == "E00" && samplingSize < samplingSizeMax)
                                    {
                                        samplingSize++;
                                    }
                                    worksheet.Cells[32, 45].Value = samplingSize;
                                }

                                if (reTest == 1)
                                {
                                    if (Res == "Avail")
                                    {
                                        if (double.TryParse(numColumns[22], out dataRes))
                                        {
                                            if (dataRes != 0 && dataRes < specMaxRespone)
                                            {
                                                double dataresfix = Math.Round(dataRes, 2, MidpointRounding.AwayFromZero);
                                                responseData.Add(dataresfix);
                                            }

                                            if (errorCode == "E00")
                                            {
                                                double minValue = specMinRespone;
                                                double maxValue = specMaxRespone;
                                                valueRes = Math.Floor(dataRes * 10) / 10;
                                                if (valueRes >= minValue && valueRes <= maxValue)
                                                {

                                                }
                                                else
                                                {
                                                    rejectedResponse++;
                                                    worksheet.Cells[37, 23] = rejectedResponse;
                                                }
                                            }
                                        }
                                    }

                                    if (Off == "Avail")
                                    {
                                        if (double.TryParse(numColumns[18], out dataOff))
                                        {
                                            if (dataOff != 0 && dataOff < specMaxOffset)
                                            {
                                                double dataofffix = Math.Round(dataOff, 2, MidpointRounding.AwayFromZero);
                                                offsetData.Add(dataofffix);
                                            }

                                            if (errorCode != "E00")
                                            {
                                                double minValue = specMinOffset;
                                                double maxValue = specMaxOffset;
                                                valueOff = Math.Round(dataOff, 2, MidpointRounding.AwayFromZero);
                                                if (valueOff >= minValue && valueOff <= maxValue)
                                                {

                                                }
                                                else
                                                {
                                                    rejectedOffset++;
                                                    worksheet.Cells[38, 23] = rejectedOffset;
                                                }
                                            }
                                        }
                                    }

                                    if (Mat == "Avail")
                                    {
                                        if (double.TryParse(numColumns[24], out dataMat))
                                        {
                                            if (dataMat != 0 && dataMat < specMaxMatch)
                                            {
                                                double datamatfix = Math.Round(dataMat, 2, MidpointRounding.AwayFromZero);
                                                matchData.Add(datamatfix);
                                            }

                                            if (errorCode == "E00")
                                            {
                                                double minValue = specMinMatch;
                                                //double maxValue = specMaxMatch;
                                                double maxValue = valueRes * 10 / 100;
                                                valueMat = Math.Round(dataMat, 2, MidpointRounding.AwayFromZero);
                                                if (valueMat >= minValue && valueMat <= maxValue)
                                                {

                                                }
                                                else
                                                {
                                                    rejectedMatch++;
                                                    worksheet.Cells[39, 23] = rejectedMatch;
                                                }
                                            }
                                        }
                                    }
                                    if (errorCode == "E00" && samplingSize < samplingSizeMax)
                                    {
                                        samplingSize++;
                                    }
                                    worksheet.Cells[32, 45].Value = samplingSize;
                                }
                            }
                        }
                        
                        if (reTest == 0)
                        {
                            if (rejectedResponse != 0 || rejectedOffset != 0 || rejectedMatch != 0)
                            {
                                reTest++;
                            }
                        }
                        else
                        {
                            reTest = 0;
                        }

                        // Response
                        int response1 = 0;
                        int response2 = 0;
                        int response3 = 0;
                        int response4 = 0;
                        int response5 = 0;
                        int response6 = 0;
                        int response7 = 0;
                        int response8 = 0;
                        int response9 = 0;
                        int response10 = 0;
                        int response11 = 0;
                        int response12 = 0;
                        int response13 = 0;
                        int response14 = 0;
                        int response15 = 0;
                        int response16 = 0;
                        int response17 = 0;
                        int response18 = 0;
                        int response19 = 0;
                        int response20 = 0;
                        int response21 = 0;
                        int response22 = 0;
                        int response23 = 0;
                        int response24 = 0;
                        int response25 = 0;
                        int response26 = 0;

                        // Offset
                        int offset1 = 0;
                        int offset2 = 0;
                        int offset3 = 0;
                        int offset4 = 0;
                        int offset5 = 0;
                        int offset6 = 0;
                        int offset7 = 0;
                        int offset8 = 0;
                        int offset9 = 0;
                        int offset10 = 0;
                        int offset11 = 0;
                        int offset12 = 0;
                        int offset13 = 0;
                        int offset14 = 0;
                        int offset15 = 0;
                        int offset16 = 0;
                        int offset17 = 0;
                        int offset18 = 0;
                        int offset19 = 0;
                        int offset20 = 0;
                        int offset21 = 0;
                        int offset22 = 0;

                        // Match
                        int match1 = 0;
                        int match2 = 0;
                        int match3 = 0;
                        int match4 = 0;
                        int match5 = 0;
                        int match6 = 0;
                        int match7 = 0;
                        int match8 = 0;
                        int match9 = 0;
                        int match10 = 0;
                        int match11 = 0;
                        int match12 = 0;
                        int match13 = 0;
                        int match14 = 0;
                        int match15 = 0;
                        int match16 = 0;
                        int match17 = 0;
                        int match18 = 0;
                        int match19 = 0;
                        int match20 = 0;
                        int match21 = 0;
                        int match22 = 0;

                        if (Res == "Avail")
                        {
                            //Range Data Response
                            double add1 = 0.0;
                            for (int n = 0; n <= 25; n++)
                            {
                                double specRes = earlyRespone;
                                double valueRangeRes = specRes + add1;
                                double valueResfix = Math.Round(valueRangeRes, 2, MidpointRounding.AwayFromZero);
                                worksheet.Cells[29, 3 + n] = valueResfix;
                                add1 += 0.1;
                            }

                            cellValueRes1 = worksheet.Cells[29, 3].Value2;
                            cellValueRes2 = worksheet.Cells[29, 4].Value2;
                            cellValueRes3 = worksheet.Cells[29, 5].Value2;
                            cellValueRes4 = worksheet.Cells[29, 6].Value2;
                            cellValueRes5 = worksheet.Cells[29, 7].Value2;
                            cellValueRes6 = worksheet.Cells[29, 8].Value2;
                            cellValueRes7 = worksheet.Cells[29, 9].Value2;
                            cellValueRes8 = worksheet.Cells[29, 10].Value2;
                            cellValueRes9 = worksheet.Cells[29, 11].Value2;
                            cellValueRes10 = worksheet.Cells[29, 12].Value2;
                            cellValueRes11 = worksheet.Cells[29, 13].Value2;
                            cellValueRes12 = worksheet.Cells[29, 14].Value2;
                            cellValueRes13 = worksheet.Cells[29, 15].Value2;
                            cellValueRes14 = worksheet.Cells[29, 16].Value2;
                            cellValueRes15 = worksheet.Cells[29, 17].Value2;
                            cellValueRes16 = worksheet.Cells[29, 18].Value2;
                            cellValueRes17 = worksheet.Cells[29, 19].Value2;
                            cellValueRes18 = worksheet.Cells[29, 20].Value2;
                            cellValueRes19 = worksheet.Cells[29, 21].Value2;
                            cellValueRes20 = worksheet.Cells[29, 22].Value2;
                            cellValueRes21 = worksheet.Cells[29, 23].Value2;
                            cellValueRes22 = worksheet.Cells[29, 24].Value2;
                            cellValueRes23 = worksheet.Cells[29, 25].Value2;
                            cellValueRes24 = worksheet.Cells[29, 26].Value2;
                            cellValueRes25 = worksheet.Cells[29, 27].Value2;
                            cellValueRes26 = worksheet.Cells[29, 28].Value2;
                        }

                        if (Off == "Avail")
                        {
                            //Range Data Offset
                            double add2 = 0.0;
                            for (int m = 0; m <= 21; m++)
                            {
                                double specOff = specMinOffset;
                                double valueRangeOff = specOff + add2;
                                double valueOfffix = Math.Round(valueRangeOff, 2, MidpointRounding.AwayFromZero);
                                worksheet.Cells[29, 30 + m] = valueOfffix;
                                add2 += 0.1;
                            }

                            //Offset
                            cellValueOff1 = worksheet.Cells[29, 30].Value2;
                            cellValueOff2 = worksheet.Cells[29, 31].Value2;
                            cellValueOff3 = worksheet.Cells[29, 32].Value2;
                            cellValueOff4 = worksheet.Cells[29, 33].Value2;
                            cellValueOff5 = worksheet.Cells[29, 34].Value2;
                            cellValueOff6 = worksheet.Cells[29, 35].Value2;
                            cellValueOff7 = worksheet.Cells[29, 36].Value2;
                            cellValueOff8 = worksheet.Cells[29, 37].Value2;
                            cellValueOff9 = worksheet.Cells[29, 38].Value2;
                            cellValueOff10 = worksheet.Cells[29, 39].Value2;
                            cellValueOff11 = worksheet.Cells[29, 40].Value2;
                            cellValueOff12 = worksheet.Cells[29, 41].Value2;
                            cellValueOff13 = worksheet.Cells[29, 42].Value2;
                            cellValueOff14 = worksheet.Cells[29, 43].Value2;
                            cellValueOff15 = worksheet.Cells[29, 44].Value2;
                            cellValueOff16 = worksheet.Cells[29, 45].Value2;
                            cellValueOff17 = worksheet.Cells[29, 46].Value2;
                            cellValueOff18 = worksheet.Cells[29, 47].Value2;
                            cellValueOff19 = worksheet.Cells[29, 48].Value2;
                            cellValueOff20 = worksheet.Cells[29, 49].Value2;
                            cellValueOff21 = worksheet.Cells[29, 50].Value2;
                            cellValueOff22 = worksheet.Cells[29, 51].Value2;
                        }

                        if (Mat == "Avail")
                        {
                            //Range Data Match
                            double add3 = 0.00;
                            for (int o = 0; o <= 21; o++)
                            {
                                double specMat = earlyMatch;
                                double valueRangeMat = specMat + add3;
                                double valueMatfix = Math.Round(valueRangeMat, 2, MidpointRounding.AwayFromZero);
                                worksheet.Cells[29, 53 + o] = valueMatfix;
                                add3 += 0.01;
                            }

                            //Match
                            cellValueMat1 = worksheet.Cells[29, 53].Value2;
                            cellValueMat2 = worksheet.Cells[29, 54].Value2;
                            cellValueMat3 = worksheet.Cells[29, 55].Value2;
                            cellValueMat4 = worksheet.Cells[29, 56].Value2;
                            cellValueMat5 = worksheet.Cells[29, 57].Value2;
                            cellValueMat6 = worksheet.Cells[29, 58].Value2;
                            cellValueMat7 = worksheet.Cells[29, 59].Value2;
                            cellValueMat8 = worksheet.Cells[29, 60].Value2;
                            cellValueMat9 = worksheet.Cells[29, 61].Value2;
                            cellValueMat10 = worksheet.Cells[29, 62].Value2;
                            cellValueMat11 = worksheet.Cells[29, 63].Value2;
                            cellValueMat12 = worksheet.Cells[29, 64].Value2;
                            cellValueMat13 = worksheet.Cells[29, 65].Value2;
                            cellValueMat14 = worksheet.Cells[29, 66].Value2;
                            cellValueMat15 = worksheet.Cells[29, 67].Value2;
                            cellValueMat16 = worksheet.Cells[29, 68].Value2;
                            cellValueMat17 = worksheet.Cells[29, 69].Value2;
                            cellValueMat18 = worksheet.Cells[29, 70].Value2;
                            cellValueMat19 = worksheet.Cells[29, 71].Value2;
                            cellValueMat20 = worksheet.Cells[29, 72].Value2;
                            cellValueMat21 = worksheet.Cells[29, 73].Value2;
                            cellValueMat22 = worksheet.Cells[29, 74].Value2;
                        }

                        int startRes = 1; // Indeks awal untuk batch ini
                        int endRes = 31;

                        //for (int i = 0; i < 31; i++)
                        if (Res == "Avail")
                        {
                            for (int i = startRes; i < endRes; i++)
                            {
                                string[] columns = rows[i].Split('\t');
                                string errorCode = columns[15];
                                // DATA RESPONSE
                                double data;
                                if (double.TryParse(columns[22], out data))
                                {
                                    valueRes = Math.Floor(data * 10) / 10;

                                    // Tentukan rentang nilai yang ingin difilter
                                    double minValue = specMinRespone; // Misalnya, rentang nilai minimum
                                    double maxValue = specMaxRespone; // Misalnya, rentang nilai maksimum

                                    if (valueRes < cellValueRes1 || valueRes > cellValueRes26 || errorCode != "E00")
                                    {
                                        endRes++;
                                    }

                                    // Filter data berdasarkan rentang nilai
                                    if (valueRes >= minValue && valueRes <= maxValue && errorCode == "E00")
                                    {
                                        // Tampilkan nilai yang memenuhi syarat di Excel
                                        if (valueRes == cellValueRes1)
                                        {
                                            if (response1 <= 13)
                                            {
                                                worksheet.Cells[28 - response1, 3] = "X";
                                                response1++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes2)
                                        {
                                            if (response2 <= 13)
                                            {
                                                worksheet.Cells[28 - response2, 4] = "X";
                                                response2++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes3)
                                        {
                                            if (response3 <= 13)
                                            {
                                                worksheet.Cells[28 - response3, 5] = "X";
                                                response3++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes4)
                                        {
                                            if (response4 <= 13)
                                            {
                                                worksheet.Cells[28 - response4, 6] = "X";
                                                response4++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes5)
                                        {
                                            if (response5 <= 13)
                                            {
                                                worksheet.Cells[28 - response5, 7] = "X";
                                                response5++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        if (valueRes == cellValueRes6)
                                        {
                                            if (response6 <= 13)
                                            {
                                                worksheet.Cells[28 - response6, 8] = "X";
                                                response6++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes7)
                                        {
                                            if (response7 <= 13)
                                            {
                                                worksheet.Cells[28 - response7, 9] = "X";
                                                response7++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes8)
                                        {
                                            if (response8 <= 13)
                                            {
                                                worksheet.Cells[28 - response8, 10] = "X";
                                                response8++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes9)
                                        {
                                            if (response9 <= 13)
                                            {
                                                worksheet.Cells[28 - response9, 11] = "X";
                                                response9++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes10)
                                        {
                                            if (response10 <= 13)
                                            {
                                                worksheet.Cells[28 - response10, 12] = "X";
                                                response10++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        if (valueRes == cellValueRes11)
                                        {
                                            if (response11 <= 13)
                                            {
                                                worksheet.Cells[28 - response11, 13] = "X";
                                                response11++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes12)
                                        {
                                            if (response12 <= 13)
                                            {
                                                worksheet.Cells[28 - response12, 14] = "X";
                                                response12++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes13)
                                        {
                                            if (response13 <= 13)
                                            {
                                                worksheet.Cells[28 - response13, 15] = "X";
                                                response13++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes14)
                                        {
                                            if (response14 <= 13)
                                            {
                                                worksheet.Cells[28 - response14, 16] = "X";
                                                response14++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes15)
                                        {
                                            if (response15 <= 13)
                                            {
                                                worksheet.Cells[28 - response15, 17] = "X";
                                                response15++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes16)
                                        {
                                            if (response16 <= 13)
                                            {
                                                worksheet.Cells[28 - response16, 18] = "X";
                                                response16++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes17)
                                        {
                                            if (response17 <= 13)
                                            {
                                                worksheet.Cells[28 - response17, 19] = "X";
                                                response17++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes18)
                                        {
                                            if (response18 <= 13)
                                            {
                                                worksheet.Cells[28 - response18, 20] = "X";
                                                response18++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes19)
                                        {
                                            if (response19 <= 13)
                                            {
                                                worksheet.Cells[28 - response19, 21] = "X";
                                                response19++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes20)
                                        {
                                            if (response20 <= 13)
                                            {
                                                worksheet.Cells[28 - response20, 22] = "X";
                                                response20++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes21)
                                        {
                                            if (response21 <= 13)
                                            {
                                                worksheet.Cells[28 - response21, 23] = "X";
                                                response21++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes22)
                                        {
                                            if (response22 <= 13)
                                            {
                                                worksheet.Cells[28 - response22, 24] = "X";
                                                response22++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes23)
                                        {
                                            if (response23 <= 13)
                                            {
                                                worksheet.Cells[28 - response23, 25] = "X";
                                                response23++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes24)
                                        {
                                            if (response24 <= 13)
                                            {
                                                worksheet.Cells[28 - response24, 26] = "X";
                                                response24++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes25)
                                        {
                                            if (response25 <= 13)
                                            {
                                                worksheet.Cells[28 - response25, 27] = "X";
                                                response25++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                        else if (valueRes == cellValueRes26)
                                        {
                                            if (response26 <= 13)
                                            {
                                                worksheet.Cells[28 - response26, 28] = "X";
                                                response26++;
                                            }
                                            else
                                            {
                                                endRes++;
                                            }
                                        }
                                    }
                                    else
                                    {

                                    }
                                }
                            }
                        }

                        int startOff = 1; // Indeks awal untuk batch ini
                        int endOff = 31;

                        if (Off == "Avail")
                        {
                            for (int i = startOff; i < endOff; i++)
                            {
                                string[] columns = rows[i].Split('\t');
                                string errorCode = columns[15];
                                double data;
                                //DATA OFFSET
                                if (double.TryParse(columns[18], out data))
                                {
                                    valueOff = Math.Floor(data * 10) / 10;

                                    // Tentukan rentang nilai yang ingin difilter
                                    double minValue = specMinOffset; // Misalnya, rentang nilai minimum
                                    double maxValue = specMaxOffset; // Misalnya, rentang nilai maksimum

                                    if (valueOff < cellValueOff1 || valueOff > cellValueOff22 || errorCode != "E00")
                                    {
                                        endOff++;
                                    }

                                    // Filter data berdasarkan rentang nilai
                                    if (valueOff >= minValue && valueOff <= maxValue && errorCode == "E00")
                                    {
                                        // Tampilkan nilai yang memenuhi syarat di Excel
                                        if (valueOff == cellValueOff1)
                                        {
                                            if (offset1 <= 13)
                                            {
                                                worksheet.Cells[28 - offset1, 30] = "X";
                                                offset1++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff2)
                                        {
                                            if (offset2 <= 13)
                                            {
                                                worksheet.Cells[28 - offset2, 31] = "X";
                                                offset2++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff3)
                                        {
                                            if (offset3 <= 13)
                                            {
                                                worksheet.Cells[28 - offset3, 32] = "X";
                                                offset3++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff4)
                                        {
                                            if (offset4 <= 13)
                                            {
                                                worksheet.Cells[28 - offset4, 33] = "X";
                                                offset4++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff5)
                                        {
                                            if (offset5 <= 13)
                                            {
                                                worksheet.Cells[28 - offset5, 34] = "X";
                                                offset5++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff6)
                                        {
                                            if (offset6 <= 13)
                                            {
                                                worksheet.Cells[28 - offset6, 35] = "X";
                                                offset6++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff7)
                                        {
                                            if (offset7 <= 13)
                                            {
                                                worksheet.Cells[28 - offset7, 36] = "X";
                                                offset7++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff8)
                                        {
                                            if (offset8 <= 13)
                                            {
                                                worksheet.Cells[28 - offset8, 37] = "X";
                                                offset8++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff9)
                                        {
                                            if (offset9 <= 13)
                                            {
                                                worksheet.Cells[28 - offset9, 38] = "X";
                                                offset9++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff10)
                                        {
                                            if (offset10 <= 13)
                                            {
                                                worksheet.Cells[28 - offset10, 39] = "X";
                                                offset10++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        if (valueOff == cellValueOff11)
                                        {
                                            if (offset11 <= 13)
                                            {
                                                worksheet.Cells[28 - offset11, 40] = "X";
                                                offset11++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff12)
                                        {
                                            if (offset12 <= 13)
                                            {
                                                worksheet.Cells[28 - offset12, 41] = "X";
                                                offset12++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff13)
                                        {
                                            if (offset13 <= 13)
                                            {
                                                worksheet.Cells[28 - offset13, 42] = "X";
                                                offset3++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff14)
                                        {
                                            if (offset14 <= 13)
                                            {
                                                worksheet.Cells[28 - offset14, 43] = "X";
                                                offset14++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff15)
                                        {
                                            if (offset15 <= 13)
                                            {
                                                worksheet.Cells[28 - offset15, 44] = "X";
                                                offset15++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff16)
                                        {
                                            if (offset16 <= 13)
                                            {
                                                worksheet.Cells[28 - offset16, 45] = "X";
                                                offset16++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff17)
                                        {
                                            if (offset17 <= 13)
                                            {
                                                worksheet.Cells[28 - offset17, 46] = "X";
                                                offset17++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff18)
                                        {
                                            if (offset18 <= 13)
                                            {
                                                worksheet.Cells[28 - offset18, 47] = "X";
                                                offset18++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff19)
                                        {
                                            if (offset19 <= 13)
                                            {
                                                worksheet.Cells[28 - offset19, 48] = "X";
                                                offset19++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff20)
                                        {
                                            if (offset20 <= 13)
                                            {
                                                worksheet.Cells[28 - offset20, 49] = "X";
                                                offset20++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff21)
                                        {
                                            if (offset21 <= 13)
                                            {
                                                worksheet.Cells[28 - offset21, 50] = "X";
                                                offset21++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                        else if (valueOff == cellValueOff22)
                                        {
                                            if (offset22 <= 13)
                                            {
                                                worksheet.Cells[28 - offset22, 51] = "X";
                                                offset22++;
                                            }
                                            else
                                            {
                                                endOff++;
                                            }
                                        }
                                    }
                                    else
                                    {

                                    }
                                }
                            }
                        }

                        int startMat = 1;
                        int endMat = 31;

                        if (Mat == "Avail")
                        {
                            for (int i = startMat; i < endMat; i++)
                            {
                                string[] columns = rows[i].Split('\t');
                                string errorCode = columns[15];
                                double data;

                                //DATA MATCH
                                if (double.TryParse(columns[24], out data))
                                {
                                    valueMat = data;
                                    double minValue = specMinMatch;
                                    double maxValue = specMaxMatch;

                                    if (valueMat < cellValueMat1 || valueMat > cellValueMat22 || errorCode != "E00")
                                    {
                                        endMat++;
                                    }

                                    // Filter data berdasarkan rentang nilai
                                    if (valueMat >= minValue && valueMat <= maxValue && errorCode == "E00")
                                    {
                                        // Tampilkan nilai yang memenuhi syarat di Excel
                                        if (valueMat == cellValueMat1)
                                        {
                                            if (match1 <= 13)
                                            {
                                                worksheet.Cells[28 - match1, 53] = "X";
                                                match1++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat2)
                                        {
                                            if (match2 <= 13)
                                            {
                                                worksheet.Cells[28 - match2, 54] = "X";
                                                match2++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat3)
                                        {
                                            if (match3 <= 13)
                                            {
                                                worksheet.Cells[28 - match3, 55] = "X";
                                                match3++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat4)
                                        {
                                            if (match4 <= 13)
                                            {
                                                worksheet.Cells[28 - match4, 56] = "X";
                                                match4++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat5)
                                        {
                                            if (match5 <= 13)
                                            {
                                                worksheet.Cells[28 - match5, 57] = "X";
                                                match5++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        if (valueMat == cellValueMat6)
                                        {
                                            if (match6 <= 13)
                                            {
                                                worksheet.Cells[28 - match6, 58] = "X";
                                                match6++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat7)
                                        {
                                            if (match7 <= 13)
                                            {
                                                worksheet.Cells[28 - match7, 59] = "X";
                                                match7++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat8)
                                        {
                                            if (match8 <= 13)
                                            {
                                                worksheet.Cells[28 - match8, 60] = "X";
                                                match8++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat9)
                                        {
                                            if (match9 <= 13)
                                            {
                                                worksheet.Cells[28 - match9, 61] = "X";
                                                match9++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat10)
                                        {
                                            if (match10 <= 13)
                                            {
                                                worksheet.Cells[28 - match10, 62] = "X";
                                                match10++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        if (valueMat == cellValueMat11)
                                        {
                                            if (match11 <= 13)
                                            {
                                                worksheet.Cells[28 - match11, 63] = "X";
                                                match11++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat12)
                                        {
                                            if (match12 <= 13)
                                            {
                                                worksheet.Cells[28 - match12, 64] = "X";
                                                match12++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat13)
                                        {
                                            if (match13 <= 13)
                                            {
                                                worksheet.Cells[28 - match13, 65] = "X";
                                                match13++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat14)
                                        {
                                            if (match14 <= 13)
                                            {
                                                worksheet.Cells[28 - match14, 66] = "X";
                                                match14++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat15)
                                        {
                                            if (match15 <= 13)
                                            {
                                                worksheet.Cells[28 - match15, 67] = "X";
                                                match15++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat16)
                                        {
                                            if (match16 <= 13)
                                            {
                                                worksheet.Cells[28 - match16, 68] = "X";
                                                match16++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat17)
                                        {
                                            if (match17 <= 13)
                                            {
                                                worksheet.Cells[28 - match17, 69] = "X";
                                                match17++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat18)
                                        {
                                            if (match18 <= 13)
                                            {
                                                worksheet.Cells[28 - match18, 70] = "X";
                                                match18++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat19)
                                        {
                                            if (match19 <= 13)
                                            {
                                                worksheet.Cells[28 - match19, 71] = "X";
                                                match19++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat20)
                                        {
                                            if (match20 <= 13)
                                            {
                                                worksheet.Cells[28 - match20, 72] = "X";
                                                match20++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat21)
                                        {
                                            if (match21 <= 13)
                                            {
                                                worksheet.Cells[28 - match21, 73] = "X";
                                                match21++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                        else if (valueMat == cellValueMat22)
                                        {
                                            if (match22 <= 13)
                                            {
                                                worksheet.Cells[28 - match22, 74] = "X";
                                                match22++;
                                            }
                                            else
                                            {
                                                endMat++;
                                            }
                                        }
                                    }
                                    else
                                    {

                                    }
                                }
                            }
                        }

                        //Judge
                        if (rejectedResponse == 0 && rejectedOffset == 0 && rejectedMatch == 0)
                        {
                            PlaceAcc2(worksheet, imageacc, 42, 24, 3); // Electrical Acc

                            if (mechanicalInspectStatus == "PASS")
                            {
                                PlaceinCell(worksheet, stampacc, 86, 42, 110, 110); //Mechanical Acc
                                PlaceAcc2(worksheet, imageacc, 42, 53, 3); // Mechanical Acc
                            }
                        }
                        else
                        {
                            worksheet.Cells[41, 23] = rejectedResponse;
                            PlaceAcc2(worksheet, imageacc, 42, 27, 3); //Electrical Rej

                            if (mechanicalInspectStatus == "PASS")
                            {
                                PlaceinCell(worksheet, stamprej, 86, 42, 110, 110); //Stamp Reject
                                PlaceAcc2(worksheet, imageacc, 42, 53, 3); // Mechanical Acc
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("INVALID");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading file: " + ex.Message);
                }
            }
            workbook.Save();

            // Tutup workbook
            workbook.Close();
            MessageBox.Show("Data Plotted Successfully");
            DisplayExcel("FORM");
        }

        void plottingNoise()
        {
            // Koneksi ke SQL Server
            string connectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123";
            string query1 = "SELECT * FROM dmQABLHeader WHERE QANo = @QA_Number";

            QANumber = cb_QANo.Text;
            DataTable dataTable1 = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query1, connection))
                {
                    // Parameter
                    command.Parameters.AddWithValue("@QA_Number", QANumber);
                    connection.Open();
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable1);
                    }
                }
            }

            clearDataNoise();

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            // Buka workbook yang sudah ada
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = workbook.Sheets["FORM"];

            worksheet.Cells[6, 10] = CB_TesterName.Text;

            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Filter jenis file yang dapat dipilih (opsional)
            openFileDialog.Filter = "All files (*.*)|*.*|Text files (*.txt)|*.txt";

            // Set opsi lainnya (opsional)
            openFileDialog.Title = "Select a file";
            openFileDialog.InitialDirectory = @"C:\"; // Direktori awal yang ditampilkan
            openFileDialog.Multiselect = false; // Memungkinkan pemilihan beberapa file

            // Tampilkan dialog dan dapatkan hasil
            DialogResult result = openFileDialog.ShowDialog();

            // Jika pengguna memilih file dan menekan tombol "OK"
            if (result == DialogResult.OK)
            {
                // Dapatkan path file yang dipilih
                string selectedFilePath = openFileDialog.FileName;
                try
                {
                    // Baca isi file teks
                    string[] rows = File.ReadAllLines(selectedFilePath);

                    // Jika ada baris dalam file
                    if (rows.Length > 0)
                    {

                        // List untuk menyimpan data setiap kolom
                        List<double> noiseData = new List<double>();

                         // Membaca file
                        foreach (string row in rows)
                        {
                            string[] numColumns = row.Split('\t');
                            double dataNoi;
                            if (numColumns.Length > 22) // Pastikan ada setidaknya 23 kolom
                            {
                                if (double.TryParse(numColumns[23], out dataNoi))
                                {
                                    if (dataNoi != 0 && dataNoi < specMaxNoise)
                                    {
                                        noiseData.Add(dataNoi);
                                    }
                                }
                            }
                        }

                        if (noiseData.Count > 0)
                        {
                            Noi = "Avail";
                            double averageNoi = noiseData.Average();
                            double avenoi = Math.Round(averageNoi, 0, MidpointRounding.AwayFromZero);
                            earlyNoise = avenoi - 9;
                        }
                        else
                        {
                            Noi = "NA";
                            PlaceImageInCell(worksheet, imagePath, 14, 76, 170, 200);
                        }


                        if (Noi == "Avail")
                        {
                            worksheet.Cells[8, 79] = "X < " + specMaxNoise + " (µVp)";
                            specMinNoise = earlyNoise;
                        }
                        else
                        {
                            worksheet.Cells[8, 79] = "NA";
                        }

                        if (Noi == "Avail")
                        {
                            foreach (string row in rows)
                            {
                                string[] numColumns = row.Split('\t');
                                double dataNoi;
                                string errorCodenoi = numColumns[15];
                                if (numColumns.Length > 22) // Pastikan ada setidaknya 23 kolom
                                {
                                    if (reTestnoi == 0)
                                    {
                                        if (double.TryParse(numColumns[23], out dataNoi))
                                        {
                                            if (dataNoi != 0 && dataNoi < specMaxNoise)
                                            {
                                                noiseData.Add(dataNoi);
                                            }

                                            if (errorCodenoi != "ENT")
                                            {
                                                double minValue = specMinNoise;
                                                double maxValue = specMaxNoise;
                                                valueNoi = dataNoi;
                                                if (valueNoi >= minValue && valueNoi <= maxValue && errorCodenoi == "E00")
                                                {

                                                }
                                                else
                                                {
                                                    rejectedNoise++;
                                                    if (rejectedNoise != 0)
                                                    {
                                                        worksheet.Cells[40, 11] = "Noise";
                                                    }
                                                    worksheet.Cells[40, 23] = rejectedNoise;
                                                }
                                            }
                                        }
                                    }

                                    if (reTestnoi == 1)
                                    {
                                        if (double.TryParse(numColumns[23], out dataNoi))
                                        {
                                            if (dataNoi != 0 && dataNoi < specMaxNoise)
                                            {
                                                noiseData.Add(dataNoi);
                                            }

                                            if (errorCodenoi == "E00")
                                            {
                                                double minValue = specMinNoise;
                                                double maxValue = specMaxNoise;
                                                valueNoi = dataNoi;
                                                if (valueNoi >= minValue && valueNoi <= maxValue)
                                                {

                                                }
                                                else
                                                {
                                                    rejectedNoise++;
                                                    worksheet.Cells[40, 23] = rejectedNoise;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            
                            if (reTestnoi == 0)
                            {
                                if (rejectedNoise != 0)
                                {
                                    reTestnoi++;
                                }
                            }
                            else
                            {
                                reTestnoi = 0;
                            }
                        }
                        
                        // Noise
                        int noise1 = 0;
                        int noise2 = 0;
                        int noise3 = 0;
                        int noise4 = 0;
                        int noise5 = 0;
                        int noise6 = 0;
                        int noise7 = 0;
                        int noise8 = 0;
                        int noise9 = 0;
                        int noise10 = 0;
                        int noise11 = 0;
                        int noise12 = 0;
                        int noise13 = 0;
                        int noise14 = 0;
                        int noise15 = 0;
                        int noise16 = 0;
                        int noise17 = 0;
                        int noise18 = 0;
                        int noise19 = 0;
                        int noise20 = 0;

                        if (Noi == "Avail")
                        {
                            int add4 = 0;
                            for (int k = 0; k <= 19; k++)
                            {
                                double specNoi = earlyNoise;
                                double valueRangeNoi = specNoi + add4;
                                worksheet.Cells[29, 76 + k] = valueRangeNoi;
                                add4 += 1;
                            }

                            //Noise
                            cellValueNoi1 = worksheet.Cells[29, 76].Value2;
                            cellValueNoi2 = worksheet.Cells[29, 77].Value2;
                            cellValueNoi3 = worksheet.Cells[29, 78].Value2;
                            cellValueNoi4 = worksheet.Cells[29, 79].Value2;
                            cellValueNoi5 = worksheet.Cells[29, 80].Value2;
                            cellValueNoi6 = worksheet.Cells[29, 81].Value2;
                            cellValueNoi7 = worksheet.Cells[29, 82].Value2;
                            cellValueNoi8 = worksheet.Cells[29, 83].Value2;
                            cellValueNoi9 = worksheet.Cells[29, 84].Value2;
                            cellValueNoi10 = worksheet.Cells[29, 85].Value2;
                            cellValueNoi11 = worksheet.Cells[29, 86].Value2;
                            cellValueNoi12 = worksheet.Cells[29, 87].Value2;
                            cellValueNoi13 = worksheet.Cells[29, 88].Value2;
                            cellValueNoi14 = worksheet.Cells[29, 89].Value2;
                            cellValueNoi15 = worksheet.Cells[29, 90].Value2;
                            cellValueNoi16 = worksheet.Cells[29, 91].Value2;
                            cellValueNoi17 = worksheet.Cells[29, 92].Value2;
                            cellValueNoi18 = worksheet.Cells[29, 93].Value2;
                            cellValueNoi19 = worksheet.Cells[29, 94].Value2;
                            cellValueNoi20 = worksheet.Cells[29, 95].Value2;
                        }

                        int startNoi = 1;
                        int endNoi = 31;

                        if (Noi == "Avail")
                        {
                            // Iterasi melalui setiap baris dan kolom
                            for (int i = startNoi; i < endNoi; i++)
                            {
                                string[] columns = rows[i].Split('\t');
                                string errorCode = columns[15];
                                double data;
                                //DATA NOISE
                                if (double.TryParse(columns[23], out data))
                                {
                                    valueNoi = data;

                                    // Tentukan rentang nilai yang ingin difilter
                                    double minValue = specMinNoise; // Misalnya, rentang nilai minimum
                                    double maxValue = specMaxNoise; // Misalnya, rentang nilai maksimum

                                    if (valueNoi < cellValueNoi1 || valueNoi > cellValueNoi20 || errorCode != "E00")
                                    {
                                        endNoi++;
                                    }

                                    // Filter data berdasarkan rentang nilai
                                    if (valueNoi >= minValue && valueNoi <= maxValue && errorCode == "E00")
                                    {
                                        // Tampilkan nilai yang memenuhi syarat di Excel
                                        if (valueNoi == cellValueNoi1)
                                        {
                                            if (noise1 <= 13)
                                            {
                                                worksheet.Cells[28 - noise1, 76] = "X";
                                                noise1++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi2)
                                        {
                                            if (noise2 <= 13)
                                            {
                                                worksheet.Cells[28 - noise2, 77] = "X";
                                                noise2++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi3)
                                        {
                                            if (noise3 <= 13)
                                            {
                                                worksheet.Cells[28 - noise3, 78] = "X";
                                                noise3++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi4)
                                        {
                                            if (noise4 <= 13)
                                            {
                                                worksheet.Cells[28 - noise4, 79] = "X";
                                                noise4++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi5)
                                        {
                                            if (noise5 <= 13)
                                            {
                                                worksheet.Cells[28 - noise5, 80] = "X";
                                                noise5++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi6)
                                        {
                                            if (noise6 <= 13)
                                            {
                                                worksheet.Cells[28 - noise6, 81] = "X";
                                                noise6++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi7)
                                        {
                                            if (noise7 <= 13)
                                            {
                                                worksheet.Cells[28 - noise7, 82] = "X";
                                                noise7++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi8)
                                        {
                                            if (noise8 <= 13)
                                            {
                                                worksheet.Cells[28 - noise8, 83] = "X";
                                                noise8++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi9)
                                        {
                                            if (noise9 <= 13)
                                            {
                                                worksheet.Cells[28 - noise9, 84] = "X";
                                                noise9++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi10)
                                        {
                                            if (noise10 <= 13)
                                            {
                                                worksheet.Cells[28 - noise10, 85] = "X";
                                                noise10++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        if (valueNoi == cellValueNoi11)
                                        {
                                            if (noise11 <= 13)
                                            {
                                                worksheet.Cells[28 - noise11, 86] = "X";
                                                noise11++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi12)
                                        {
                                            if (noise12 <= 13)
                                            {
                                                worksheet.Cells[28 - noise12, 87] = "X";
                                                noise12++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi13)
                                        {
                                            if (noise13 <= 13)
                                            {
                                                worksheet.Cells[28 - noise13, 88] = "X";
                                                noise13++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi14)
                                        {
                                            if (noise14 <= 13)
                                            {
                                                worksheet.Cells[28 - noise14, 89] = "X";
                                                noise14++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi15)
                                        {
                                            if (noise15 <= 13)
                                            {
                                                worksheet.Cells[28 - noise15, 90] = "X";
                                                noise15++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi16)
                                        {
                                            if (noise16 <= 13)
                                            {
                                                worksheet.Cells[28 - noise16, 91] = "X";
                                                noise16++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi17)
                                        {
                                            if (noise17 <= 13)
                                            {
                                                worksheet.Cells[28 - noise17, 92] = "X";
                                                noise17++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi18)
                                        {
                                            if (noise18 <= 13)
                                            {
                                                worksheet.Cells[28 - noise18, 93] = "X";
                                                noise18++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi19)
                                        {
                                            if (noise19 <= 13)
                                            {
                                                worksheet.Cells[28 - noise19, 94] = "X";
                                                noise19++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                        else if (valueNoi == cellValueNoi20)
                                        {
                                            if (noise20 <= 13)
                                            {
                                                worksheet.Cells[28 - noise20, 95] = "X";
                                                noise20++;
                                            }
                                            else
                                            {
                                                endNoi++;
                                            }
                                        }
                                    }
                                    else
                                    {

                                    }
                                }
                            }
                        }

                        //Judge
                        if (rejectedResponse == 0 && rejectedOffset == 0 && rejectedMatch == 0 && rejectedNoise == 0)
                        {
                            PlaceAcc2(worksheet, imageacc, 42, 24, 3); // Electrical Acc

                            if (mechanicalInspectStatus == "PASS")
                            {
                                PlaceinCell(worksheet, stampacc, 86, 42, 110, 110); //Mechanical Acc
                            }
                        }
                        else
                        {
                            worksheet.Cells[41, 23] = rejectedResponse;
                            PlaceAcc2(worksheet, imageacc, 42, 27, 3); //Electrical Rej

                            if (mechanicalInspectStatus == "PASS")
                            {
                                PlaceinCell(worksheet, stamprej, 86, 42, 110, 110); //Stamp Reject
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("INVALID");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading file: " + ex.Message);
                }
            }
            workbook.Save();

            // Tutup workbook
            workbook.Close();
            MessageBox.Show("Data Plotted Successfully");
            DisplayExcel("FORM");
        }

        bool Istester1 = false;
        bool Istester2 = false;
        private void BT_SelectFile_Click(object sender, EventArgs e)
        {
            if (CB_Tester.Text == "")
            {
                MessageBox.Show("Please Input The Tester!!!");
            }
            else if (CB_TesterName.Text == "")
            {
                MessageBox.Show("Please Input The Tester Name!!!");
            }
            else if (cb_QANo.Text == "")
            {
                MessageBox.Show("Please Input The QA Number!!!");
            }
            else if (TB_InspectedBy.Text == "")
            {
                MessageBox.Show("Please Input The Inspected!!!");
            }
            else
            {
                if (CB_Tester.SelectedItem.ToString() == "SIGNAL")
                {
                    if (Istester1 == true)
                    {
                        DialogResult result = MessageBox.Show("Do you want to continue to Signal Tester?", "Confirmation", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            //Reset variabel
                            samplingSize = 0;
                            rejectedResponse = 0;
                            rejectedOffset = 0;
                            rejectedMatch = 0;
                            rejectedNoise = 0;
                            earlySpec = true;
                            mechanicalInspectStatus = "";
                            plottingData();
                        }
                        else
                        {
                            this.Show();
                        }
                    }
                    else
                    {
                        Istester1 = true;
                        //Reset variabel
                        samplingSize = 0;
                        rejectedResponse = 0;
                        rejectedOffset = 0;
                        rejectedMatch = 0;
                        rejectedNoise = 0;
                        earlySpec = true;
                        mechanicalInspectStatus = "";
                        plottingData();
                    }
                }
                else if (CB_Tester.SelectedItem.ToString() == "NOISE")
                {
                    if (Istester2 == true)
                    {
                        DialogResult result = MessageBox.Show("Do you want to continue to Noise Tester?", "Confirmation", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            rejectedNoise = 0;
                            plottingNoise();
                        }
                        else
                        {
                            this.Show();
                        }
                    }
                    else
                    {
                        Istester2 = true;
                        plottingNoise();
                    }
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you want to print and save the data?", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                Form2 print = new Form2(this);
                print.Show();
            }
            else
            {
                this.Show();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you want to delete?", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                samplingSize = 0;
                rejectedResponse = 0;
                rejectedOffset = 0;
                rejectedMatch = 0;
                rejectedNoise = 0;
                mechanicalInspectStatus = "";
                earlySpec = true;
                clearDataAll();
            }
            else
            {
                this.Show();
            }
        }

        void Tester()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TesterName.txt"); // Path to your text file
            if (File.Exists(filePath))
            {
                try
                {
                    CB_TesterName.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    CB_TesterName.AutoCompleteSource = AutoCompleteSource.ListItems;
                    string[] lines = File.ReadAllLines(filePath);
                    CB_TesterName.Items.AddRange(lines);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading file: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("File not found: " + filePath);
            }
        }

        void QANo()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "QANumber.txt"); // Path to your text file
            if (File.Exists(filePath))
            {
                try
                {
                    cb_QANo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    cb_QANo.AutoCompleteSource = AutoCompleteSource.ListItems;
                    string[] lines = File.ReadAllLines(filePath);
                    cb_QANo.Items.AddRange(lines);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading file: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("File not found: " + filePath);
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            cb_QANo.Items.Clear();
            LoadQANumberData();
        }

        private void tableLayoutPanel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void cb_QANo_SelectedIndexChanged(object sender, EventArgs e)
        {
            earlySpec = true;
            Istester1 = false;
            Istester2 = false;
        }

        private void dateTimePicker1_ValueChanged_1(object sender, EventArgs e)
        {
            cb_QANo.Items.Clear();
            LoadQANumberData();
        }
    }
}
