using System;
using System.IO;
using System.Windows.Forms;
using Excel.Interfaces;
using Excel.Models;
using Excel.Services;

namespace Excel
{
    public partial class MyExcel : Form
    {
        private Table _table = new();
        private readonly ITableService _tableService;
        
        public MyExcel(ITableService tableService)
        {
            _tableService = tableService;
            InitializeComponent();
            // TODO: Move to configs
            InitializeDataGridView(35, 35);
        }

        private void InitializeDataGridView(int rows, int columns)
        {
            dataGridView.ColumnHeadersVisible = true;
            dataGridView.RowHeadersVisible = true;
            dataGridView.ColumnCount = columns;
            for (var i = 0; i < columns; i++)
            {
                var columnName = CoordinateEncoder.Encode(i);
                dataGridView.Columns[i].Name = columnName;
                dataGridView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            for (var i = 0; i < rows; i++)
            {
                dataGridView.Rows.Add("");
                dataGridView.Rows[i].HeaderCell.Value = i.ToString();
            }

            dataGridView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            _table = _tableService.CreateTable(columns, rows);
        }


        private void calculateButton_Click(object sender, EventArgs e)
        {
            var col = dataGridView.SelectedCells[0].ColumnIndex;
            var row = dataGridView.SelectedCells[0].RowIndex;
            var expression = textBox.Text;
            if (expression == "") return;
            _tableService.ChangeCellWithAllPointers(_table, row, col, expression, dataGridView);
            dataGridView[col, row].Value = _table.Sheet[row][col].Value;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            var col = dataGridView.SelectedCells[0].ColumnIndex;
            var row = dataGridView.SelectedCells[0].RowIndex;
            var expression = "";
            try
            {
                expression = _table.Sheet[row][col].Expression;
            }
            catch 
            {
                MessageBox.Show("Selected incorrect cell");    
            }
            textBox.Text = expression;
            textBox.Focus();
        }

        private void addRowButton_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new System.Windows.Forms.DataGridViewRow();
            if (dataGridView.Columns.Count == 0)
            {
                MessageBox.Show("There are no colums");  //???
                return;
            }
            dataGridView.Rows.Add(row);
            dataGridView.Rows[_table.RowsAmount].HeaderCell.Value = _table.RowsAmount.ToString();
            _tableService.AddRow(_table, dataGridView);

        }
        private void addColButton_Click(object sender, EventArgs e)
        {
            string name = CoordinateEncoder.Encode(_table.ColumnsAmount);
            dataGridView.Columns.Add(name, name);
            _tableService.AddCol(_table);
        }

        private void delRowButton_Click(object sender, EventArgs e)
        {
            if (!_tableService.DeleteRow(_table, dataGridView))
                return;
            dataGridView.Rows.RemoveAt(_table.RowsAmount);
        }

        private void delColButton_Click(object sender, EventArgs e)
        {
            if (!_tableService.DeleteColumn(_table, dataGridView))
                return;
            dataGridView.Columns.RemoveAt(_table.ColumnsAmount);
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "TableFile|*.txt";
            saveFileDialog.Title = "Save table as a file";
            saveFileDialog.ShowDialog();

            if (string.IsNullOrEmpty(saveFileDialog.FileName)) return;
            
            var fs = (FileStream)saveFileDialog.OpenFile();
            using (StreamWriter sw = new StreamWriter(fs))
                _tableService.Save(_table, sw);
            fs.Close();
        }

        private void openButton_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "TableFile|*.txt";
            openFileDialog.Title = "Open Table";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;
            var sr = new StreamReader(openFileDialog.FileName);
            _tableService.Clear(_table);
            dataGridView.Rows.Clear();
            dataGridView.Columns.Clear();
            int.TryParse(sr.ReadLine(), out var row);
            int.TryParse(sr.ReadLine(), out var column);
            InitializeDataGridView(row, column);
            _tableService.Open(_table, row, column, sr, dataGridView);
            sr.Close();
        }
    }
}
