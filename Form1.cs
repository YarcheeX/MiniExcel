using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace MiniExcel
{
    public partial class MainForm : Form
    {
        
        
        public MainForm()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();

                ofd.Filter = "XML Files (*.xml; *.xlsx; *.xlsm; *.xlsb) |*.xml; *.xlsx; *.xlsm; *.xlsb";

                DialogResult res = ofd.ShowDialog();

                if (res == DialogResult.OK)
                {
                    string fn = ofd.FileName;
                    myTable1.LoadTable(fn);
                    
                }else
                {
                    throw new Exception("Файл не выбран!");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myTable1.saveTable();
        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    myTable1.saveTable(new System.IO.FileInfo(sfd.FileName));
                }
                else
                {
                    throw new Exception("Файл не выбран!");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void пустойЛистToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myTable1.TabPages.Add("Лист" + (myTable1.TabCount + 1).ToString());
            myTable1.makePage(myTable1.TabPages[myTable1.TabCount - 1]);
        }

        
    }
}
