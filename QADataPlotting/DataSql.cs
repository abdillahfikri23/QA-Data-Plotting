using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data.SqlClient;

namespace QADataPlotting
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //daftarTabelSQL();
            dataGridView1.DataSource = Select();
        }

        public DataGridView MyDataGridView
        {
            get { return dataGridView1; } // ganti dataGridView1 dengan nama sebenarnya
        }

        public DataTable Select()
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123";
            DataTable dt = new DataTable();
            try
            {
                //string sql = "SELECT * FROM dmQABLHeader"; //Week, Date, Type
                string sql = "SELECT * FROM dmQABLSerialPaper"; //SerialNo, Quantity
                //string sql = "SELECT * FROM dmQABLDefectElectrical"; //Defect Mode, Description 
                //string sql = "SELECT * FROM dmQABLDetail"; //Inspected, FA Number, Description, Sampling Size
                //string sql = "SELECT * FROM dmQABLVisualMechanical";
                //string sql = "SELECT * FROM TESTER_COINTRAY_DEVSPECQA";
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                conn.Open();
                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
            return dt;
        }

        void daftarTabelSQL()
        {
            // Ganti dengan informasi koneksi SQL Server Anda
            string connectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    // Query untuk mendapatkan daftar tabel
                    string query = "SELECT TABLE_NAME AS 'Table Name' FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'";

                    using (SqlDataAdapter adapter = new SqlDataAdapter(query, connection))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);

                        // Tampilkan data di DataGridView
                        dataGridView1.DataSource = dataTable;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Terjadi kesalahan: " + ex.Message);
                }
            }
        }
    }
}
