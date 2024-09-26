using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QADataPlotting
{
    public partial class Spec : Form
    {
        private QAPlotting _QAplot;
        public Spec(QAPlotting QAplot)
        {
            InitializeComponent();
            _QAplot = QAplot;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }

        public void AddValueRes(string text)
        {
            tb_Maxres.Text = text;
        }

        public void AddValueOff(string text)
        {
            tb_Maxoff.Text = text;
        }

        public void AddValueMat(string text)
        {
            tb_Maxmat.Text = text;
        }

        public void AddValueNoi(string text)
        {
            tb_Maxnoi.Text = text;
        }

        private void Spec_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cb_Res_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cb_Off_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cb_Mat_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cb_Noi_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
