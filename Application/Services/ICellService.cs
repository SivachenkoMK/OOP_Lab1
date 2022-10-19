using System.Collections.Generic;
using Excel.Models;

namespace Excel.Services
{
    public interface ICellService
    {
        void UpdateCellData(Cell cell, string expression, string value, List<Cell> references, List<Cell> pointers);

        bool IsLoop(Cell cell);

        void AddPointersAndReferences(Cell cell);

        void DeletePointersAndReferences(Cell cell);
    }
}