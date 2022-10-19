using System;
using System.Collections.Generic;

namespace Application.Models
{
    [Serializable]
    public class Cell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Name { get; set; }
        public string Expression { get; set; } = "";
        public string Value { get; set; } = "0";

        public List<Cell> PointersToThis { get; set; } = new();
        public List<Cell> ReferencesFromThis { get; set; } = new();
        public List<Cell> NewReferencesFromThis { get; set; } = new();


        public Cell(int row, int column)
        {
            Row = row;
            Column = column;
            Name = CoordinateEncoder.Encode(column) + row;
        }
    }
}