using System;
using System.Collections.Generic;

namespace Application.Models
{
    [Serializable]
    public class Table
    {
        public int ColumnsAmount { get; set; }
        public int RowsAmount { get; set; }

        public readonly List<List<Cell>> Sheet = new();

        public readonly Dictionary<string, string> DisplayedValues = new();

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