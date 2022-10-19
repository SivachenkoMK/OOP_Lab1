using System.IO;
using System.Windows.Forms;
using Application.Models;

namespace Application.Interfaces
{
    public interface ITableService
    {
        Table CreateTable(int columnsAmount, int rowsAmount);

        void Clear(Table table);

        void ChangeCellWithAllPointers(Table table, int row, int col, string expression,
            DataGridView dataGridView1);

        void AddRow(Table table, DataGridView dataGridView1);

        void AddCol(Table table);

        bool DeleteRow(Table table, DataGridView dataGridView1);

        bool DeleteColumn(Table table, DataGridView dataGridView1);

        void Save(Table table, FileStream stream);

        Table Open(string path);
    }
}