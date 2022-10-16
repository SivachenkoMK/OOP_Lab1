using System.Collections.Generic;
using System.Linq;

namespace Excel
{
    public class Cell
    {
        public string Expression { get; set; }
        public string Value { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public string Name { get; set; }

        public readonly List<Cell> PointersToThis = new();
        public List<Cell> ReferencesFromThis = new();
        public readonly List<Cell> NewReferencesFromThis = new();


        public Cell(int row, int column)
        {
            Row = row;
            Column = column;
            Name = ColumnNameConverter.To26System(column) + row;
            Value = "0";
            Expression = "";
        }

        public void SetCell(string expression, string value, List<Cell> references, List<Cell> pointers)
        {
            Value = value;
            Expression = expression;
            ReferencesFromThis.Clear();
            ReferencesFromThis.AddRange(references);
            PointersToThis.Clear();
            PointersToThis.AddRange(pointers);
        }

        public bool CheckLoop(List<Cell> list)  //??
        {
            if (list.Any(cell => cell.Name == Name))
            {
                return false;
            }
            foreach (var point in PointersToThis)
            {
                if (list.Any(cell => cell.Name == point.Name))
                {
                    return false;
                }

                if (!point.CheckLoop(list))
                {
                    return false;
                }
            }
            return true;
        }

        public void AddPointersAndReferences()
        {
            foreach (var point in NewReferencesFromThis)
            {
                point.PointersToThis.Add(this);
            }
            ReferencesFromThis = NewReferencesFromThis;
        }

        public void DeletePointersAndReferences()
        {
            if (ReferencesFromThis == null) return;

            foreach (var point in ReferencesFromThis)
            {
                point.PointersToThis.Remove(this);
            }
            ReferencesFromThis = null;
        }
    }
}