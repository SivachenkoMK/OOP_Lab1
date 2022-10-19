using System.Collections.Generic;
using System.Linq;
using Application.Interfaces;
using Application.Models;

namespace Application.Services
{
    public class CellService : ICellService
    {
        public void UpdateCellData(Cell cell, string expression, string value, List<Cell> references, List<Cell> pointers)
        {
            cell.ReferencesFromThis.Clear();
            cell.PointersToThis.Clear();
            cell.ReferencesFromThis.AddRange(references);
            cell.PointersToThis.AddRange(pointers);
        }

        public bool IsLoop(Cell cell, List<Cell>? viewedCells = null)
        {
            viewedCells ??= new List<Cell>();
            if (viewedCells.Contains(cell))
                return true;

            viewedCells.Add(cell);

            return cell.NewReferencesFromThis.Any(newReferencedCell => newReferencedCell.Name == cell.Name) || cell.NewReferencesFromThis.Any(point => IsLoop(point, viewedCells));
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
            foreach (var point in cell.ReferencesFromThis)
            {
                point.PointersToThis.Remove(cell);
            }
            cell.ReferencesFromThis = new List<Cell>();
        }
    }
}