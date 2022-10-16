using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Excel.Models
{
    public class Table
    {
        public int ColumnsAmount { get; set; }
        public int RowsAmount { get; set; }

        public readonly List<List<Cell>> Sheet = new();

        public Table()
        {

        }

        public Table(int columnsAmount, int rowsAmount)
        {
            ColumnsAmount = columnsAmount;
            RowsAmount = rowsAmount;
        }
    }
}