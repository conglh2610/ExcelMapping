using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReaderTool
{
    public partial class Home : Form
    {
        public Home()
        {
            InitializeComponent();
        }

        private void excelReaderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var excelReaderFrm = new frmExcelReader();
            excelReaderFrm.Show();
        }
    }
}
