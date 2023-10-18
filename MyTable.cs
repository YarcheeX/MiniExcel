using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using OfficeOpenXml;
using System.IO;

namespace MiniExcel
{
    class MyTable : TabControl
    {
        private List<DataGridView> dgvCollection = new List<DataGridView>();
        private string bufferCell;
       
        private Point lastClickPos;
        private ContextMenuStrip CMS;
        private bool saveFlag;
        private string fileName;
        
        private ContextMenuStrip dgvCellCms;
        private ContextMenuStrip dgvRowCms;
        private ContextMenuStrip dgvColCms;

        public MyTable()
        {
            saveFlag = true;
            CMS = GetCMS();
            this.SelectedIndexChanged += new System.EventHandler(this.myTabControl1_SelectedIndexChanged);
          
            bufferCell = string.Empty;

            dgvCellCms = new ContextMenuStrip();
            dgvCellCms.Items.Add("Копировать", null, new EventHandler(copyCell));
            dgvCellCms.Items.Add("Вставить", null, new EventHandler(pasteCell));
            dgvCellCms.Items.Add("Очистить содержимое", null, new EventHandler(clearCell));

            dgvRowCms = new ContextMenuStrip();
            dgvRowCms.Items.Add("Копировать", null, new EventHandler(copyRow));
            dgvRowCms.Items.Add("Вставить", null, new EventHandler(pasteRow));
            dgvRowCms.Items.Add("Очистить содержимое", null, new EventHandler(clearRow));

            dgvColCms = new ContextMenuStrip();
            dgvColCms.Items.Add("Копировать", null, new EventHandler(copyCol));
            dgvColCms.Items.Add("Вставить", null, new EventHandler(pasteCol));
            dgvColCms.Items.Add("Очистить содержимое", null, new EventHandler(clearCol));


        }

        public bool IsSaved()
        {
           return saveFlag;
        }
        private ContextMenuStrip GetCMS()
        {
            ContextMenuStrip cms = new ContextMenuStrip();

            ToolStripMenuItem deleteItem = new ToolStripMenuItem("Удалить");
            deleteItem.Click += DeleteItem_Clicked;
            cms.Items.Add(deleteItem);

            ToolStripMenuItem renameItem = new ToolStripMenuItem("Переименовтать");
            renameItem.Click += RenameItem_Clicked;
            cms.Items.Add(renameItem);
            return cms;

        }

        private void DeleteItem_Clicked(object sender, EventArgs e)
        {
            for (int i = 0; i < TabCount; i++)
            {
                Rectangle rect = this.GetTabRect(i);

                if (rect.Contains(this.PointToClient(lastClickPos)))
                {
                    this.deletePage(i);
                }
            }
        }

        private void RenameItem_Clicked(object sender, EventArgs e)
        {
            for (int i = 0; i < TabCount; i++)
            {
                Rectangle rect = this.GetTabRect(i);

                if (rect.Contains(this.PointToClient(lastClickPos)))
                {
                    renameListForm form = new renameListForm();
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        this.TabPages[i].Text = form.Controls.OfType<TextBox>().First().Text;
                        saveFlag = false;
                    }

                }

            }
        }
        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);

            if (e.Button == MouseButtons.Right)
            {
                lastClickPos = Cursor.Position;
                CMS.Show(Cursor.Position);
            }
        }

        public void LoadTable(string fn)
        {

            fileName = fn;
            if (!this.IsSaved())
            {
                DialogResult result = MessageBox.Show("У вас имеется несохраненный прогресс работы."
                    + "При открытии нового файла прогресс будет утерян. Открыть?",
                    "ВНИМАНИЕ!", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);

                if (result != DialogResult.OK) return;

                this.SelectedIndexChanged -= myTabControl1_SelectedIndexChanged;
                this.TabPages.Clear();
                dgvCollection.Clear();
                this.SelectedIndexChanged += myTabControl1_SelectedIndexChanged;
            }              
                
            var fi = new FileInfo(fileName);

            using (ExcelPackage exp = new ExcelPackage(fi))
            {
                ExcelWorkbook work_book = exp.Workbook;
                ExcelWorksheets ws_collection = work_book.Worksheets;

                foreach (ExcelWorksheet sh in ws_collection)
                {
                    this.TabPages.Add(sh.Name);
                }

                foreach (TabPage tp in this.TabPages)
                {
                    makePage(tp);
                }

                for (int i = 0; i < dgvCollection.Count; i++)
                {
                    SheetToTable(ws_collection[i], i);
                }
            }   

        }

        public void makePage(TabPage tp)
        {
            dgvCollection.Add(new DataGridView());
            DataGridView dgv = dgvCollection.Last();
            dgv.Dock = DockStyle.Fill;
            dgv.RowCount = 99;
            dgv.ColumnCount = 26;

            char header_ch = 'A';
            for (int j = 0; j <= 25; j++)
                dgv.Columns[j].HeaderText = header_ch++.ToString();
           
            dgv.RowHeadersVisible = true;

            for (int j = 0; j <= 98; j++)
                dgv.Rows[j].HeaderCell.Value = (j + 1).ToString();


            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.CellValueChanged += new DataGridViewCellEventHandler(cell_changed);
            dgv.CellMouseClick += new DataGridViewCellMouseEventHandler(dgv_cell_right_click);
            dgv.RowHeaderMouseClick += new DataGridViewCellMouseEventHandler(dgv_row_right_click);
            dgv.ColumnHeaderMouseClick += new DataGridViewCellMouseEventHandler(dgv_col_right_click);

            tp.Controls.Add(dgvCollection.Last());
            saveFlag = false;
        }

        public void deletePage(int i)
        {
            TabPages[i].Dispose();
            saveFlag = false;
        }

        public void saveTable()
        {
            if (IsSaved()) return;

            using (ExcelPackage exp = new ExcelPackage())
            {
                for (int i = 0; i < this.TabCount; i++)
                {
                    exp.Workbook.Worksheets.Add(this.TabPages[i].Text);
                }

                for (int i = 0; i < dgvCollection.Count; i++)
                {
                    dgvCollection[i].CurrentCell = null;
                    dgvCollection[i].SelectAll();
                    foreach (DataGridViewCell cell in dgvCollection[i].SelectedCells)
                    {
                        if (cell.Value != null)
                            exp.Workbook.Worksheets[i].Cells[cell.RowIndex + 1, cell.ColumnIndex + 1].Value = cell.Value;
                    }
                }

                exp.SaveAs(new FileInfo(fileName));
               
            }

            saveFlag = true;
        }

        public void saveTable(FileInfo fi)
        {
            if (!fi.Exists) return;
            using (ExcelPackage exp = new ExcelPackage())
            {
                for (int i = 0; i < this.TabCount; i++)
                {
                    exp.Workbook.Worksheets.Add(this.TabPages[i].Text);
                }

                for (int i = 0; i < dgvCollection.Count; i++)
                {
                    dgvCollection[i].CurrentCell = null;
                    dgvCollection[i].SelectAll();
                    foreach (DataGridViewCell cell in dgvCollection[i].SelectedCells)
                    {
                        if (cell.Value != null)
                            exp.Workbook.Worksheets[i].Cells[cell.RowIndex + 1, cell.ColumnIndex + 1].Value = cell.Value;
                    }
                }

                exp.SaveAs(fi);

            }

            saveFlag = true;
        }

        private void SheetToTable(ExcelWorksheet sheet, int index)
        {
            
            int rows;
            int cols;

            if (sheet.Dimension == null)
                return;
            rows = sheet.Dimension.End.Row;
            cols = sheet.Dimension.End.Column;


            for (int j = 1; j <= rows; j++)
            {
                for (int k = 1; k <= cols; k++)
                {
                    dgvCollection[index].Rows[j - 1].Cells[k - 1].Value = sheet.Cells[j, k].Value;
                }
            }      

        }

        private void myTabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.TabPages[this.SelectedIndex].Controls.OfType<DataGridView>().
             First().AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void dgv_cell_right_click(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex != -1 && e.RowIndex != -1 && e.Button == MouseButtons.Right)
            {
                DataGridViewCell c = (sender as DataGridView)[e.ColumnIndex, e.RowIndex];
                if (!c.Selected)
                {
                    c.DataGridView.ClearSelection();
                    c.DataGridView.CurrentCell = c;
                    c.Selected = true;

                }

                dgvCellCms.Show(Cursor.Position);
            }
        }

        private void copyCell(object sender, EventArgs e)
        {
            bufferCell = dgvCollection[this.SelectedIndex].SelectedCells[0].Value.ToString();
        }

        private void pasteCell(object sender, EventArgs e)
        {
            dgvCollection[this.SelectedIndex].SelectedCells[0].Value = bufferCell;
        }

        private void clearCell(object sender, EventArgs e)
        {
            dgvCollection[this.SelectedIndex].SelectedCells[0].Value = bufferCell;
        }


        private void dgv_row_right_click(object sender, DataGridViewCellMouseEventArgs e)
        {
            if(e.Button == MouseButtons.Right)
            {

                DataGridViewRowHeaderCell c = (DataGridViewRowHeaderCell)sender;
                
                DataGridViewRow r = c.DataGridView.Rows[c.RowIndex];
                
                if(!r.Selected)
                {
                    r.DataGridView.ClearSelection();
                    r.DataGridView.CurrentCell = r.Cells[0];
                    r.Selected = true;
                }

                dgvRowCms.Show(Cursor.Position);

            }
        }

        private void copyRow(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(dgvCollection[this.SelectedIndex].CurrentRow);
        }

        private void pasteRow(object sender, EventArgs e)
        {
            
            if (Clipboard.GetDataObject().GetDataPresent(typeof(DataGridViewRow)))
            {
                DataGridViewRow r = (DataGridViewRow)Clipboard.GetDataObject().GetData(typeof(DataGridViewRow));
                for(int i = 0; i < 26; i++)
                {
                    dgvCollection[this.SelectedIndex].CurrentRow.Cells[i].Value =
                        r.Cells[i].Value;
                }
            }
        }

        private void clearRow(object sender, EventArgs e)
        {
            foreach(DataGridViewCell cell in dgvCollection[this.SelectedIndex].CurrentRow.Cells)
            {
                cell.Value = null;
            }
        }

        private void dgv_col_right_click(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                DataGridViewColumn c = (sender as DataGridViewColumn);

                if (!c.Selected)
                {
                    c.DataGridView.ClearSelection();
                    c.DataGridView.CurrentCell = c.DataGridView.Rows[0].Cells[c.Index];
                    c.Selected = true;
                }

                dgvColCms.Show(Cursor.Position);

            }


        }

        private void copyCol(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(dgvCollection[this.SelectedIndex].SelectedColumns[0]);
        }

        private void pasteCol(object sender, EventArgs e)
        {
            if (Clipboard.GetDataObject().GetDataPresent(typeof(DataGridViewColumn)))
            {
                DataGridViewColumn c = (DataGridViewColumn)Clipboard.GetDataObject().GetData(typeof(DataGridViewColumn));
                for (int i = 0; i < 99; i++)
                {
                    dgvCollection[this.SelectedIndex].Rows[i].Cells[c.Index].Value =
                        c.DataGridView.Rows[i].Cells[c.Index].Value;
                }
            }

        }

        private void clearCol(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in dgvCollection[this.SelectedIndex].SelectedCells)
            {
                cell.Value = null;
            }

        }

        
        private void cell_changed(object sender, DataGridViewCellEventArgs e)
        {
            saveFlag = false;
        }


    }
}
