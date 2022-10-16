using System;
using System.IO;
using System.Windows.Forms;
using Excel.Models;
using Excel.Services;

namespace Excel
{
    public partial class MyExcel : Form
    {
        private Table _table = new();
        private readonly TableService tableService = new();
        public MyExcel()
        {
            InitializeComponent();
            InitializeDataGridView(10, 35);
            tableService = new TableService();
        }

        private void InitializeDataGridView(int rows, int columns)
        {
            dataGridView1.ColumnHeadersVisible = true;
            dataGridView1.RowHeadersVisible = true;
            dataGridView1.ColumnCount = columns;
            for (var i = 0; i < columns; i++)
            {
                var columnName = ColumnNameConverter.To26System(i);
                dataGridView1.Columns[i].Name = columnName;
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            for (var i = 0; i < rows; i++)
            {
                dataGridView1.Rows.Add("");
                dataGridView1.Rows[i].HeaderCell.Value = i.ToString();
            }

            dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            _table = tableService.CreateTable(columns, rows);
        }


        private void calculateButton_Click(object sender, EventArgs e)
        {
            var col = dataGridView1.SelectedCells[0].ColumnIndex;
            var row = dataGridView1.SelectedCells[0].RowIndex;
            var expression = textBox1.Text;
            if (expression == "") return;
            tableService.ChangeCellWithAllPointers(_table, row, col, expression, dataGridView1);
            dataGridView1[col, row].Value = _table.Sheet[row][col].Value;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            var col = dataGridView1.SelectedCells[0].ColumnIndex;
            var row = dataGridView1.SelectedCells[0].RowIndex;
            var expression = "";
            try
            {
                expression = _table.Sheet[row][col].Expression;
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
            dataGridView1.Rows[_table.RowsAmount].HeaderCell.Value = _table.RowsAmount.ToString();
            tableService.AddRow(_table, dataGridView1);

        }
        private void addColButton_Click(object sender, EventArgs e)
        {
            string name = ColumnNameConverter.To26System(_table.ColumnsAmount);
            dataGridView1.Columns.Add(name, name);
            tableService.AddCol(_table);
        }

        private void delRowButton_Click(object sender, EventArgs e)
        {
            if (!tableService.DeleteRow(_table, dataGridView1))
                return;
            dataGridView1.Rows.RemoveAt(_table.RowsAmount);
        }

        private void delColButton_Click(object sender, EventArgs e)
        {
            if (!tableService.DeleteColumn(_table, dataGridView1))
                return;
            dataGridView1.Columns.RemoveAt(_table.ColumnsAmount);
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
                tableService.Save(_table, sw);
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
            tableService.Clear(_table);
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            int.TryParse(sr.ReadLine(), out var row);
            int.TryParse(sr.ReadLine(), out var column);
            InitializeDataGridView(row, column);
            tableService.Open(_table, row, column, sr, dataGridView1);
            sr.Close();
        }
    }
}
