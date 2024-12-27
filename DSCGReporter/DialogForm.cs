using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DSCGReporter
{
    public partial class DialogForm : Form
    {
        public DataTable interbeds;
        public int InterbedId;
        public DialogForm()
        {
            InitializeComponent();
        }

        private void DialogForm_Load(object sender, EventArgs e)
        {
            comboBox1.DataSource = interbeds;
            comboBox1.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            InterbedId = (comboBox1.SelectedIndex + 1) * 100;
        }
    }
}
