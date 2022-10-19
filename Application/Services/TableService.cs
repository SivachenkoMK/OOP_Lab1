﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel.Interfaces;
using Excel.Models;

namespace Excel.Services
{
    public class TableService : ITableService
    { 
        private readonly ICellService _cellService;

        public TableService(ICellService cellService)
        {
            _cellService = cellService;
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
        } //init table by 2 param

        public void Clear(Table table)
        {
            foreach (var column in table.Sheet)
                column.Clear();

            table.Sheet.Clear();
            table.DisplayedValues.Clear();

            table.RowsAmount = 0;
            table.ColumnsAmount = 0;
        }

        private string GetValue(string expression)
        {
            try
            {
                var res = Evaluator.GetValue(expression).ToString();
                if (res == "∞")
                    res = "Division by zero error";
                return res;
            }
            catch (Exception)
            {
                return "Error";
            }
        }

        public void ChangeCellWithAllPointers(Table table, int row, int col, string expression,
            DataGridView dataGridView1) //refresh cell value with check loops(Main func)
        {
            var currentCell = table.Sheet[row][col];
            _cellService.DeletePointersAndReferences(currentCell);
            currentCell.Expression = expression;
            currentCell.NewReferencesFromThis.Clear();

            if (expression != "")
            {
                if (expression[0] != '=') //expression not formula
                {
                    currentCell.Value = expression;
                    table.DisplayedValues[GetFullName(row, col)] = expression;
                    foreach (var cell in currentCell.PointersToThis)
                    {
                        RefreshCellAndPointers(table, cell, dataGridView1);
                    }

                    return;
                }
            } //expression formula

            string newExpression = ConvertReferences(table, row, col, expression);
            if (newExpression != "")
                newExpression = newExpression.Remove(0, 1);

            if (_cellService.IsLoop(currentCell)) //check new references for loop 
            {
                MessageBox.Show("There is a loop! Change the expression");
                currentCell.Expression = "";
                currentCell.Value = "";
                dataGridView1[col, row].Value = "0";
                return;
            }

            //new_references without loops
            _cellService.AddPointersAndReferences(currentCell);
            string val = GetValue(newExpression); //calculate ready expression

            if (val == "Error") //cannot calculate
            {
                MessageBox.Show("Error in cell " + currentCell.Name);
                currentCell.Expression = "";
                currentCell.Value = "0";
                dataGridView1[currentCell.Column, currentCell.Row].Value = "0";
                return;
            }

            currentCell.Value = val;
            table.DisplayedValues[GetFullName(row, col)] = val;
            foreach (var cell in currentCell.PointersToThis) //refresh all cells which has formula with currCell
                RefreshCellAndPointers(table, cell, dataGridView1);

        }

        private string GetFullName(int row, int col)
        {
            var cell = new Cell(row, col);
            return cell.Name;
        }

        private bool RefreshCellAndPointers(Table table, Cell cell, DataGridView dataGridView1) //refresh cell
        {
            cell.NewReferencesFromThis.Clear();
            var newExpression =
                ConvertReferences(table, cell.Row, cell.Column, cell.Expression); //expression without Cell Names
            newExpression = newExpression.Remove(0, 1); //remove '='
            var value = GetValue(newExpression); //calculate ready expression

            if (value == "Error")
            {
                MessageBox.Show("Error in cell " + cell.Name);
                cell.Expression = "";
                cell.Value = "0";
                dataGridView1[cell.Column, cell.Row].Value = "0";
                return false;
            }

            table.Sheet[cell.Row][cell.Column].Value = value;
            table.DisplayedValues[GetFullName(cell.Row, cell.Column)] = value;
            dataGridView1[cell.Column, cell.Row].Value = value;

            return cell.PointersToThis.All(point => RefreshCellAndPointers(table, point, dataGridView1));
        }

        private void RefreshReferences(Table table) //refresh only refs from each cell in all table
        {
            foreach (List<Cell> list in table.Sheet)
            {
                foreach (Cell cell in list)
                {
                    cell.ReferencesFromThis?.Clear();
                    cell.NewReferencesFromThis?.Clear();
                    if (cell.Expression == "")
                        continue;
                    if (cell.Expression[0] == '=') //has formula
                    {
                        ConvertReferences(table, cell.Row, cell.Column, cell.Expression);
                        cell.ReferencesFromThis?.AddRange(cell.NewReferencesFromThis!);
                    }
                }
            }
        }

        private Table tempTable;

        private string ConvertReferences(Table table, int row, int col, string expr) // 5+4*AA1-->5+4*('Value of AA1) and add references
        {
            const string cellNamePattern = @"[A-Z]+[0-9]+";
            var regex = new Regex(cellNamePattern, RegexOptions.IgnoreCase);

            SetReferences(table, row, col, expr, regex);

            tempTable = table;
            
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

        private string ReferenceToValue(Match m) //Evaluator for converting
        {
            if (!tempTable.DisplayedValues.ContainsKey(m.Value)) return m.Value;
            return tempTable.DisplayedValues[m.Value] == "" ? "0" : tempTable.DisplayedValues[m.Value];
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
            var lastRowRef = new List<Cell>(); //Cells that have references on the delete row
            var notEmptyCells = new List<Cell>();
            if (table.RowsAmount == 0)
                return false;
            var curCount = table.RowsAmount - 1;
            for (var i = 0; i < table.ColumnsAmount; i++)
            {
                var name = GetFullName(curCount, i);
                if (table.DisplayedValues[name] != "0" && table.DisplayedValues[name] != "" && table.DisplayedValues[name] != " ")
                    notEmptyCells.Add(table.Sheet[curCount][i]);
                if (table.Sheet[curCount][i].PointersToThis.Count != 0) //select cells that points to deleted cell
                    lastRowRef.AddRange(table.Sheet[curCount][i].PointersToThis);
            }

            if (lastRowRef.Count != 0 || notEmptyCells.Count != 0)
            {
                var errorMessage = "";
                if (notEmptyCells.Count != 0)
                {
                    errorMessage = notEmptyCells.Aggregate("There are not empty cells: ",
                        (current, cell) => current + string.Join(";", cell.Name));
                    errorMessage += "\n";
                }

                if (lastRowRef.Count != 0)
                {
                    errorMessage += "There are cells that point to cells from current Row:\n";
                    errorMessage = lastRowRef.Aggregate(errorMessage,
                        (current, cell) => current + string.Join(";", cell.Name));
                }

                errorMessage += "\nAre you sure want to delete this column?";
                var res = MessageBox.Show(errorMessage, "Warning", MessageBoxButtons.YesNo);
                if (res == DialogResult.No)
                    return false;
            }

            for (var i = 0; i < table.ColumnsAmount; i++)
            {
                var name = GetFullName(curCount, i);
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
            var lastColRef = new List<Cell>(); //Cells that have references on the delete column
            var notEmptyCells = new List<Cell>();
            if (table.ColumnsAmount == 1)
                return false;
            var curCount = table.ColumnsAmount - 1;
            for (var i = 0; i < table.RowsAmount; i++)
            {
                var name = GetFullName(i, curCount);
                if (table.DisplayedValues[name] != "0" && table.DisplayedValues[name] != "" && table.DisplayedValues[name] != " ")
                    notEmptyCells.Add(table.Sheet[i][curCount]);
                if (table.Sheet[i][curCount].PointersToThis.Count != 0) //select cells that points to deleted cell
                    lastColRef.AddRange(table.Sheet[i][curCount].PointersToThis);
            }

            if (lastColRef.Count != 0 || notEmptyCells.Count != 0)
            {
                var errorMessage = "";
                if (notEmptyCells.Count != 0)
                {
                    errorMessage = lastColRef.Aggregate("There are not empty cells: ",
                        (current, cell) => current + string.Join(";", cell.Name));
                    errorMessage += "\n";
                }

                if (lastColRef.Count != 0)
                {
                    errorMessage += "There are cells that point to cells from current column:\n";
                    errorMessage = lastColRef.Aggregate(errorMessage,
                        (current, cell) => current + string.Join(";", cell.Name));
                }

                errorMessage += "\nAre you sure want to delete this column?";
                var res = MessageBox.Show(errorMessage, "Warning", MessageBoxButtons.YesNo);
                if (res == DialogResult.No)
                    return false;
            }

            for (var i = 0; i < table.RowsAmount; i++)
            {
                var name = GetFullName(i, curCount);
                table.DisplayedValues.Remove(name);
            }

            foreach (var cell in notEmptyCells)
            {
                if (cell.ReferencesFromThis == null) continue;
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
