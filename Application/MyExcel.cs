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

namespace Excel
{
    public partial class MyExcel : Form
    {
        private readonly Table _table = new();
        public MyExcel()
        {
            InitializeComponent();
            InitializeDataGridView(10, 35);
        }

        private void InitializeDataGridView(int rows, int columns)
        {
            dataGridView1.ColumnHeadersVisible = true;
            dataGridView1.RowHeadersVisible = true;
            dataGridView1.ColumnCount = columns;
            for (var i = 0; i < columns; i++)
            {
                var columnName = _26BasedSystem.To26System(i);
                dataGridView1.Columns[i].Name = columnName;
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            for (var i = 0; i < rows; i++)
            {
                dataGridView1.Rows.Add("");
                dataGridView1.Rows[i].HeaderCell.Value = i.ToString();
            }

            dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            _table.SetTable(columns, rows);
        }


        private void calculateButton_Click(object sender, EventArgs e)
        {
            var col = dataGridView1.SelectedCells[0].ColumnIndex;
            var row = dataGridView1.SelectedCells[0].RowIndex;
            var expression = textBox1.Text;
            if (expression == "") return;
            _table.ChangeCellWithAllPointers(row, col, expression, dataGridView1);
            dataGridView1[col, row].Value = Table.Grid[row][col].value;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            var col = dataGridView1.SelectedCells[0].ColumnIndex;
            var row = dataGridView1.SelectedCells[0].RowIndex;
            var expression = "";
            try
            {
                expression = Table.Grid[row][col].expression;
            }
            catch 
            {
                MessageBox.Show("Selected incorrect cell");    
            }
            textBox1.Text = expression;
            textBox1.Focus();
        }

        private void addRowButton_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new System.Windows.Forms.DataGridViewRow();
            if (dataGridView1.Columns.Count == 0)
            {
                MessageBox.Show("There are no colums");  //???
                return;
            }
            dataGridView1.Rows.Add(row);
            dataGridView1.Rows[_table.RowCount].HeaderCell.Value = _table.RowCount.ToString();
            _table.AddRow(dataGridView1);

        }
        private void addColButton_Click(object sender, EventArgs e)
        {
            string name = _26BasedSystem.To26System(_table.ColCount);
            dataGridView1.Columns.Add(name, name);
            _table.AddCol(dataGridView1);
        }

        private void delRowButton_Click(object sender, EventArgs e)
        {
            if (!_table.DeleteRow(dataGridView1))
                return;
            dataGridView1.Rows.RemoveAt(_table.RowCount);
        }

        private void delColButton_Click(object sender, EventArgs e)
        {
            if (!_table.DeleteColumn(dataGridView1))
                return;
            dataGridView1.Columns.RemoveAt(_table.ColCount);
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "TableFile|*.txt";
            saveFileDialog1.Title = "Save table file";
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                FileStream fs = (FileStream)saveFileDialog1.OpenFile();
                StreamWriter sw = new StreamWriter(fs);
                _table.Save(sw);
                sw.Close();
                fs.Close();
            }
        }

        private void openButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "TableFile|*.txt";
            openFileDialog1.Title = "Open Table File";
            if (openFileDialog1.ShowDialog() != DialogResult.OK)
                return;
            StreamReader sr = new StreamReader(openFileDialog1.FileName);
            _table.Clear();
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            int.TryParse(sr.ReadLine(), out var row);
            int.TryParse(sr.ReadLine(), out var column);
            InitializeDataGridView(row, column);
            _table.Open(row, column, sr, dataGridView1);
            sr.Close();
        }
    }
}
