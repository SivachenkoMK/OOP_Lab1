using System.Collections.Generic;
using System.Linq;
using Excel.Models;

namespace Excel.Services
{
    public class CellService : ICellService
    {
        public void UpdateCellData(Cell cell, string expression, string value, List<Cell> references, List<Cell> pointers)
        {
            cell.ReferencesFromThis.Clear();
            cell.ReferencesFromThis.AddRange(references);
            cell.PointersToThis.Clear();
            cell.PointersToThis.AddRange(pointers);
        }

        public bool IsLoop(Cell cell, List<Cell>? viewedCells)
        {
            viewedCells ??= new List<Cell>();
            if (viewedCells.Contains(cell))
                return true;
            var newReferences = cell.NewReferencesFromThis;
            if (newReferences.Any(newReferencedCell => newReferencedCell.Name == cell.Name))
            {
                return true;
            }
            
            viewedCells.Add(cell);

            return cell.NewReferencesFromThis.Any(point => IsLoop(point, viewedCells));
        }

        public void AddPointersAndReferences(Cell cell)
        {
            foreach (var point in cell.NewReferencesFromThis)
            {
                point.PointersToThis.Add(cell);
            }
            cell.ReferencesFromThis = cell.NewReferencesFromThis;
        }

        public void DeletePointersAndReferences(Cell cell)
        {
            if (cell.ReferencesFromThis == null) return;

            foreach (var point in cell.ReferencesFromThis)
            {
                point.PointersToThis.Remove(cell);
            }
            cell.ReferencesFromThis = null;
        }

    }
}