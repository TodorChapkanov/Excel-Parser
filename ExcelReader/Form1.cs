using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        private const string InvalidFilePathMessage = "Invalid File Path";
        public Form1()
        {
            InitializeComponent();
        }

        private void SubmitButton_Click(object sender, EventArgs e)
        {
            var dbContext = DatabaseOperator.GetDbContext();

            var filePath = this.FilePathTextBox.Text;

           DatabaseOperator.GetColumnNames();

            if (string.IsNullOrWhiteSpace(filePath))
            {
                MessageBox.Show(InvalidFilePathMessage);
                return;
            }

            DatabaseOperator.ParsFile(filePath,dbContext);


            try
            {
                DatabaseOperator.DisposeDbConnection(dbContext);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            
        }

       

        private void Button1_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var fileName = openFileDialog.FileName;
                this.FilePathTextBox.Text = fileName;
            }
        }
    }
}
