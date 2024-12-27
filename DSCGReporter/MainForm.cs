using DevExpress.XtraBars;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DSCGReporter
{
    public partial class MainForm : DevExpress.XtraBars.FluentDesignSystem.FluentDesignForm
    {
        public MainForm()
        {
            InitializeComponent();
            authCombo.SelectedIndex = 0;
            netCombo.SelectedIndex = 0;
            reportTypeCombo.SelectedIndex = 0;
            gdbTypeCombo.SelectedIndex = 0;

            DSCGConnections.OpenCatalogConnection();
            DataTable projects = DataRepo.GetProjects();

            gridControl1.DataSource = projects;
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            EnableControls();
        }

        private void EnableControls()
        {
            loginTextBox.Enabled = authCombo.SelectedIndex == 1;
            passwordTextBox.Enabled = authCombo.SelectedIndex == 1;
            reportTypeCombo.Enabled = gdbTypeCombo.SelectedIndex == 0;
        }

        private void accordionControlElement5_Click(object sender, EventArgs e)
        {
            try
            {
                SetConnectionParameters();
                DSCGConnections.OpenGDBConnection();

                DialogForm InterbedForm = new DialogForm();
                InterbedForm.interbeds = DataRepo.GetInterbeds();

                InterbedForm.ShowDialog(this);

                if (InterbedForm.DialogResult == DialogResult.OK)
                {
                    labelControl1.Text = "Выполняется расчет";
                    Refresh();

                    DSCGReports_1.CreateBalanceReport(gdbTypeCombo.SelectedIndex + 1, reportTypeCombo.SelectedIndex + 1, InterbedForm.InterbedId, 1);

                    labelControl1.Text = "Отчет сформирован";
                }
            }
            finally
            {
                DSCGConnections.GDBConnection.Close();
            }

        }

        private void SetConnectionParameters()
        {
            DSCGConnections.IntegratedSecurity = authCombo.SelectedIndex == 0;
            DSCGConnections.UserName = loginTextBox.Text;
            DSCGConnections.Password = passwordTextBox.Text;
            DSCGConnections.Database = projectsGridView.GetFocusedRowCellValue("Database").ToString();

            if (netCombo.SelectedIndex == 0)
                DSCGConnections.ServerAddr = projectsGridView.GetFocusedRowCellValue("LocalServer").ToString().Trim();
            else
                DSCGConnections.ServerAddr = projectsGridView.GetFocusedRowCellValue("Server").ToString().Trim();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            EnableControls();
        }

        private void gdbTypeCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            reportTypeCombo.SelectedIndex = 0;
            EnableControls();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var x = Directory.GetCurrentDirectory();
            var y = Directory.GetParent(x);
        }
    }
}
