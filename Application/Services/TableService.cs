using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application.Configs;
using Application.Interfaces;
using Application.Models;
using Microsoft.Extensions.Options;

namespace Application.Services
{
    public class TableService : ITableService
    { 
        private readonly ICellService _cellService;
        private readonly ErrorMessages _errorMessages;

        public TableService(ICellService cellService, IOptions<ErrorMessages> errorMessagesOptions)
        {
            _cellService = cellService;
            _errorMessages = errorMessagesOptions.Value;
        }
        
        public Table CreateTable(int columnsAmount, int rowsAmount)
        {
            var table = new Table(columnsAmount, rowsAmount);

            for (var i = 0; i < table.RowsAmount; i++)
            {
                var newRow = new List<Cell>();
                for (var j = 0; j < table.ColumnsAmount; j++)
                {
                    newRow.Add(new Cell(i, j));
                    table.DisplayedValues.Add(newRow.Last().Name, "");
                }

                table.Sheet.Add(newRow);
            }
            return table;
        }

        public void Clear(Table table)
        {
            foreach (var column in table.Sheet)
                column.Clear();

            table.Sheet.Clear();
            table.DisplayedValues.Clear();

            table.RowsAmount = 0;
            table.ColumnsAmount = 0;
        }

        private void TryGetValue(string expression, out string value)
        {
            value = Evaluator.GetValue(expression).ToString();
            if (value == "∞")
                throw new ArgumentException(_errorMessages.DivisionByZero);
        }

        public void ChangeCellWithAllPointers(Table table, int row, int col, string expression,
            DataGridView dataGridView1) 
        {
            var cell = table.Sheet[row][col];
            _cellService.DeletePointersAndReferences(cell);
            cell.Expression = expression;
            cell.NewReferencesFromThis.Clear();

            if (expression != "")
            {
                if (expression[0] != '=') 
                {
                    cell.Value = expression;
                    table.DisplayedValues[GetFullNameForCell(row, col)] = expression;
                    foreach (var pointingToThis in cell.PointersToThis)
                    {
                        RefreshCellAndPointers(table, pointingToThis, dataGridView1);
                    }

                    return;
                }
            } 

            var newExpression = ConvertReferences(table, row, col, expression);
            if (newExpression != "")
                newExpression = newExpression.Remove(0, 1);

            if (_cellService.IsLoop(cell))
            {
                MessageBox.Show(_errorMessages.Loop);
                cell.Expression = "";
                cell.Value = "";
                dataGridView1[col, row].Value = "0";
                return;
            }
            
            _cellService.AddPointersAndReferences(cell);

            var isSuccessful = GetValue(dataGridView1, newExpression, cell, out var val);

            if (!isSuccessful)
                return;
            
            cell.Value = val;
            table.DisplayedValues[GetFullNameForCell(row, col)] = val;
            foreach (var pointingToThisCells in cell.PointersToThis)
                RefreshCellAndPointers(table, pointingToThisCells, dataGridView1);

        }

        private bool GetValue(DataGridView dataGridView1, string newExpression, Cell cell, out string value)
        {
            try
            {
                TryGetValue(newExpression, out value);
            }
            catch (ArgumentException argumentException)
            {
                MessageBox.Show(string.Format(_errorMessages.MessageAndName, argumentException.Message, cell.Name));
                WrongSetUpCellFormatting(dataGridView1, cell, out value);
                return false;
            }
            catch (Exception)
            {
                MessageBox.Show(string.Format(_errorMessages.Name, cell.Name));
                WrongSetUpCellFormatting(dataGridView1, cell, out value);
                return false;
            }
            
            return true;
        }

        private void WrongSetUpCellFormatting(DataGridView dataGridView1, Cell cell, out string value)
        {
            cell.Expression = "";
            cell.Value = "0";
            dataGridView1[cell.Column, cell.Row].Value = "0";
            value = "";
        }

        private string GetFullNameForCell(int row, int col)
        {
            var cell = new Cell(row, col);
            return cell.Name;
        }

        private bool RefreshCellAndPointers(Table table, Cell cell, DataGridView dataGridView1) 
        {
            cell.NewReferencesFromThis.Clear();
            var newExpression =
                ConvertReferences(table, cell.Row, cell.Column, cell.Expression); 
            newExpression = newExpression.Remove(0, 1); 
            var isSuccess = GetValue(dataGridView1, newExpression, cell, out var value); 

            if (!isSuccess)
                return false;

            table.Sheet[cell.Row][cell.Column].Value = value;
            table.DisplayedValues[GetFullNameForCell(cell.Row, cell.Column)] = value;
            dataGridView1[cell.Column, cell.Row].Value = value;

            return cell.PointersToThis.All(point => RefreshCellAndPointers(table, point, dataGridView1));
        }

        private void RefreshReferences(Table table)
        {
            foreach (var cell in table.Sheet.SelectMany(row => row))
            {
                cell.ReferencesFromThis.Clear();
                cell.NewReferencesFromThis.Clear();
                if (string.IsNullOrWhiteSpace(cell.Expression))
                    continue;
                if (cell.Expression[0] != '=') continue;
                    
                ConvertReferences(table, cell.Row, cell.Column, cell.Expression);
                cell.ReferencesFromThis.AddRange(cell.NewReferencesFromThis);
            }
        }

        private Table? _tempTable;

        private string ConvertReferences(Table table, int row, int col, string expr)
        {
            const string cellNamePattern = @"[A-Z]+[0-9]+";
            var regex = new Regex(cellNamePattern, RegexOptions.IgnoreCase);

            SetReferences(table, row, col, expr, regex);

            _tempTable = table;
            
            return regex.Replace(expr, ReferenceToValue);
        }

        private void SetReferences(Table table, int row, int col, string expr, Regex regex)
        {
            foreach (Match match in regex.Matches(expr))
            {
                if (!table.DisplayedValues.ContainsKey(match.Value)) continue;
                
                var nums = CoordinateEncoder.Decode(match.Value);
                table.Sheet[row][col].NewReferencesFromThis.Add(table.Sheet[nums.Item1][nums.Item2]);
            }
        }

        private string ReferenceToValue(Match m) 
        {
            if (!_tempTable!.DisplayedValues.ContainsKey(m.Value)) return m.Value;
            return _tempTable.DisplayedValues[m.Value] == "" ? "0" : _tempTable.DisplayedValues[m.Value];
        }

        public void AddRow(Table table, DataGridView dataGridView1)
        {
            var newRow = new List<Cell>();
            for (var i = 0; i < table.ColumnsAmount; i++)
            {
                newRow.Add(new Cell(table.RowsAmount, i));
                table.DisplayedValues.Add(newRow.Last().Name, "");
            }

            table.Sheet.Add(newRow);
            RefreshReferences(table);
            table.RowsAmount++;
        }

        public void AddCol(Table table)
        {
            for (var i = 0; i < table.RowsAmount; i++)
            {
                table.Sheet[i].Add(new Cell(i, table.ColumnsAmount));
                table.DisplayedValues.Add(table.Sheet[i].Last().Name, "");
            }

            RefreshReferences(table);
            table.ColumnsAmount++;
        }

        public bool DeleteRow(Table table, DataGridView dataGridView1)
        {
            var lastRowRef = new List<Cell>(); 
            var notEmptyCells = new List<Cell>();
            if (table.RowsAmount == 0)
                return false;
            var curCount = table.RowsAmount - 1;
            for (var i = 0; i < table.ColumnsAmount; i++)
            {
                var name = GetFullNameForCell(curCount, i);
                if (table.DisplayedValues[name] != "0" && table.DisplayedValues[name] != "" && table.DisplayedValues[name] != " ")
                    notEmptyCells.Add(table.Sheet[curCount][i]);
                if (table.Sheet[curCount][i].PointersToThis.Count != 0)
                    lastRowRef.AddRange(table.Sheet[curCount][i].PointersToThis);
            }

            if (lastRowRef.Count != 0 || notEmptyCells.Count != 0)
            {
                var errorMessage = "";
                if (notEmptyCells.Count != 0)
                {
                    errorMessage = notEmptyCells.Aggregate(_errorMessages.NonEmptyCellsPresent,
                        (current, cell) => current + string.Join(";", cell.Name));
                    errorMessage += "\n";
                }

                if (lastRowRef.Count != 0)
                {
                    errorMessage += _errorMessages.ReferencesToCurrentRowPresent + "\n";
                    errorMessage = lastRowRef.Aggregate(errorMessage,
                        (current, cell) => current + string.Join(";", cell.Name));
                }

                errorMessage += "\n" + _errorMessages.ConfirmDeletingThisRow;
                var res = MessageBox.Show(errorMessage, _errorMessages.Warning, MessageBoxButtons.YesNo);
                if (res == DialogResult.No)
                    return false;
            }

            for (var i = 0; i < table.ColumnsAmount; i++)
            {
                var name = GetFullNameForCell(curCount, i);
                table.DisplayedValues.Remove(name);
            }

            foreach (var cell in notEmptyCells)
            {
                foreach (var reference in cell.ReferencesFromThis)
                    reference.PointersToThis.Remove(cell);
            }

            foreach (var cell in lastRowRef.Where(cell => cell.Row != curCount))
            {
                RefreshCellAndPointers(table, cell, dataGridView1);
            }

            table.Sheet.RemoveAt(curCount);
            table.RowsAmount--;
            return true;
        }

        public bool DeleteColumn(Table table, DataGridView dataGridView1)
        {
            var lastColRef = new List<Cell>(); 
            var notEmptyCells = new List<Cell>();
            if (table.ColumnsAmount == 1)
                return false;
            var curCount = table.ColumnsAmount - 1;
            for (var i = 0; i < table.RowsAmount; i++)
            {
                var name = GetFullNameForCell(i, curCount);
                if (table.DisplayedValues[name] != "0" && table.DisplayedValues[name] != "" && table.DisplayedValues[name] != " ")
                    notEmptyCells.Add(table.Sheet[i][curCount]);
                if (table.Sheet[i][curCount].PointersToThis.Count != 0) 
                    lastColRef.AddRange(table.Sheet[i][curCount].PointersToThis);
            }

            if (lastColRef.Count != 0 || notEmptyCells.Count != 0)
            {
                var errorMessage = "";
                if (notEmptyCells.Count != 0)
                {
                    errorMessage = lastColRef.Aggregate(_errorMessages.NonEmptyCellsPresent,
                        (current, cell) => current + string.Join(";", cell.Name));
                    errorMessage += "\n";
                }

                if (lastColRef.Count != 0)
                {
                    errorMessage += _errorMessages.ReferencesToCurrentColumnPresent;
                    errorMessage = lastColRef.Aggregate(errorMessage,
                        (current, cell) => current + string.Join(";", cell.Name));
                }

                errorMessage += "\n" + _errorMessages.ConfirmDeletingThisColumn;
                var res = MessageBox.Show(errorMessage, _errorMessages.Warning, MessageBoxButtons.YesNo);
                if (res == DialogResult.No)
                    return false;
            }

            for (var i = 0; i < table.RowsAmount; i++)
            {
                var name = GetFullNameForCell(i, curCount);
                table.DisplayedValues.Remove(name);
            }

            foreach (var cell in notEmptyCells)
            {
                foreach (var reference in cell.ReferencesFromThis)
                    reference.PointersToThis.Remove(cell);
            }

            foreach (var cell in lastColRef)
                RefreshCellAndPointers(table, cell, dataGridView1);
            for (var i = 0; i < table.RowsAmount; i++)
                table.Sheet[i].RemoveAt(curCount);
            table.ColumnsAmount--;
            return true;

        }

        public void Save(Table table, FileStream stream)
        {
            IFormatter formatter = new BinaryFormatter();  
            formatter.Serialize(stream, table);
            stream.Close();
        }

        public Table Open(string path)
        {
            IFormatter formatter = new BinaryFormatter();  
            Stream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);  
            var table = (Table) formatter.Deserialize(stream);  
            stream.Close();
            
            return table;
        }
    }
}
