using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Excel
{
    public class Table
    {
        private const int DefaultCol = 35;
        private const int DefaultRow = 10;
        public int ColCount;
        public int RowCount;
        public static readonly List<List<Cell>> Grid = new();
        private readonly Dictionary<string, string> _dictionary = new();

        public Table()
        {
            SetTable(DefaultCol, DefaultRow);
        }

        public Table(int col, int row)
        {
            SetTable(col, row);
        }

        public void SetTable(int col, int row)
        {
            Clear();
            ColCount = col;
            RowCount = row;
            for (int i = 0; i < RowCount; i++)
            {
                var newRow = new List<Cell>();
                for (int j = 0; j < ColCount; j++)
                {
                    newRow.Add(new Cell(i, j));
                    _dictionary.Add(newRow.Last().GetName(), "");
                }

                Grid.Add(newRow);
            }
        } //init table by 2 param

        public void Clear()
        {
            foreach (List<Cell> list in Grid)
                list.Clear();
            Grid.Clear();
            _dictionary.Clear();
            RowCount = 0;
            ColCount = 0;
        }

        private string Calculate(string expression)
        {
            try
            {
                var res = (Calculator.Evaluate(expression)).ToString();
                if (res == "∞")
                    res = "Division by zero error";
                return res;
            }
            catch (Exception)
            {
                return "Error";
            }
        }

        public void ChangeCellWithAllPointers(int row, int col, string expression,
            DataGridView dataGridView1) //refresh cell value with check loops(Main func)
        {
            var currCell = Grid[row][col];
            currCell.DeletePointersAndReferences();
            currCell.Expression = expression;
            currCell.NewReferencesFromThis.Clear();

            if (expression != "")
            {
                if (expression[0] != '=') //expression not formula
                {
                    currCell.Value = expression;
                    _dictionary[FullName(row, col)] = expression;
                    foreach (Cell cell in currCell.PointersToThis)
                    {
                        RefreshCellAndPointers(cell, dataGridView1);
                    }

                    return;
                }
            } //expression formula

            string newExpression = ConvertReferences(row, col, expression);
            if (newExpression != "")
                newExpression = newExpression.Remove(0, 1);

            if (!currCell.CheckLoop(currCell.NewReferencesFromThis)) //check new references for loop 
            {
                MessageBox.Show("There is a loop! Change the expression");
                currCell.Expression = "";
                currCell.Value = "";
                dataGridView1[col, row].Value = "0";
                return;
            }

            //new_references without loops
            currCell.AddPointersAndReferences();
            string val = Calculate(newExpression); //calculate ready expression

            if (val == "Error") //cannot calculate
            {
                MessageBox.Show("Error in cell " + currCell.GetName());
                currCell.Expression = "";
                currCell.Value = "0";
                dataGridView1[currCell.Column, currCell.Row].Value = "0";
                return;
            }

            currCell.Value = val;
            _dictionary[FullName(row, col)] = val;
            foreach (Cell cell in currCell.PointersToThis) //refresh all cells which has formula with currCell
                RefreshCellAndPointers(cell, dataGridView1);

        }

        private static string FullName(int row, int col)
        {
            Cell cell = new Cell(row, col);
            return cell.GetName();
        }

        private bool RefreshCellAndPointers(Cell cell, DataGridView dataGridView1) //refresh cell
        {
            cell.NewReferencesFromThis.Clear();
            var newExpression =
                ConvertReferences(cell.Row, cell.Column, cell.Expression); //expression without Cell Names
            newExpression = newExpression.Remove(0, 1); //remove '='
            var value = Calculate(newExpression); //calculate ready expression

            if (value == "Error")
            {
                MessageBox.Show("Error in cell " + cell.GetName());
                cell.Expression = "";
                cell.Value = "0";
                dataGridView1[cell.Column, cell.Row].Value = "0";
                return false;
            }

            Grid[cell.Row][cell.Column].Value = value;
            _dictionary[FullName(cell.Row, cell.Column)] = value;
            dataGridView1[cell.Column, cell.Row].Value = value;

            return cell.PointersToThis.All(point => RefreshCellAndPointers(point, dataGridView1));
        }

        private void RefreshReferences() //refresh only refs from each cell in all table
        {
            foreach (List<Cell> list in Grid)
            {
                foreach (Cell cell in list)
                {
                    cell.ReferencesFromThis?.Clear();
                    cell.NewReferencesFromThis?.Clear();
                    if (cell.Expression == "")
                        continue;
                    if (cell.Expression[0] == '=') //has formula
                    {
                        ConvertReferences(cell.Row, cell.Column, cell.Expression);
                        cell.ReferencesFromThis?.AddRange(cell.NewReferencesFromThis!);
                    }
                }
            }
        }

        private string
            ConvertReferences(int row, int col, string expr) // 5+4*AA1-->5+4*('Value of AA1) and add references
        {
            const string cellNamePattern = @"[A-Z]+[0-9]+";
            var regex = new Regex(cellNamePattern, RegexOptions.IgnoreCase);

            foreach (Match match in regex.Matches(expr))
            {
                if (_dictionary.ContainsKey(match.Value)) //addReference
                {
                    var nums = _26BasedSystem.From26System(match.Value);
                    Grid[row][col].NewReferencesFromThis.Add(Grid[nums.Item1][nums.Item2]);
                }
            }

            MatchEvaluator evaluator = ReferenceToValue;
            var newExpression = regex.Replace(expr, evaluator);
            return newExpression;
        }

        private string ReferenceToValue(Match m) //Evaluator for converting
        {
            if (!_dictionary.ContainsKey(m.Value)) return m.Value;
            return _dictionary[m.Value] == "" ? "0" : _dictionary[m.Value];
        }

        public void AddRow(DataGridView dataGridView1)
        {
            var newRow = new List<Cell>();
            for (var i = 0; i < ColCount; i++)
            {
                newRow.Add(new Cell(RowCount, i));
                _dictionary.Add(newRow.Last().GetName(), "");
            }

            Grid.Add(newRow);
            RefreshReferences();
            RowCount++;
        }

        public void AddCol()
        {
            for (var i = 0; i < RowCount; i++)
            {
                Grid[i].Add(new Cell(i, ColCount));
                _dictionary.Add(Grid[i].Last().GetName(), "");
            }

            RefreshReferences();
            ColCount++;
        }

        public bool DeleteRow(DataGridView dataGridView1)
        {
            var lastRowRef = new List<Cell>(); //Cells that have references on the delete row
            var notEmptyCells = new List<Cell>();
            if (RowCount == 0)
                return false;
            var curCount = RowCount - 1;
            for (var i = 0; i < ColCount; i++)
            {
                var name = FullName(curCount, i);
                if (_dictionary[name] != "0" && _dictionary[name] != "" && _dictionary[name] != " ")
                    notEmptyCells.Add(Grid[curCount][i]);
                if (Grid[curCount][i].PointersToThis.Count != 0) //select cells that points to deleted cell
                    lastRowRef.AddRange(Grid[curCount][i].PointersToThis);
            }

            if (lastRowRef.Count != 0 || notEmptyCells.Count != 0)
            {
                var errorMessage = "";
                if (notEmptyCells.Count != 0)
                {
                    errorMessage = notEmptyCells.Aggregate("There are not empty cells: ",
                        (current, cell) => current + string.Join(";", cell.GetName()));
                    errorMessage += "\n";
                }

                if (lastRowRef.Count != 0)
                {
                    errorMessage += "There are cells that point to cells from current Row:\n";
                    errorMessage = lastRowRef.Aggregate(errorMessage,
                        (current, cell) => current + string.Join(";", cell.GetName()));
                }

                errorMessage += "\nAre you sure want to delete this column?";
                var res = MessageBox.Show(errorMessage, "Warning", MessageBoxButtons.YesNo);
                if (res == DialogResult.No)
                    return false;
            }

            for (var i = 0; i < ColCount; i++)
            {
                var name = FullName(curCount, i);
                _dictionary.Remove(name);
            }

            foreach (var cell in notEmptyCells)
            {
                if (cell.ReferencesFromThis == null) continue;

                foreach (var reference in cell.ReferencesFromThis)
                    reference.PointersToThis.Remove(cell);
            }

            foreach (var cell in lastRowRef.Where(cell => cell.Row != curCount))
            {
                RefreshCellAndPointers(cell, dataGridView1);
            }

            Grid.RemoveAt(curCount);
            RowCount--;
            return true;
        }

        public bool DeleteColumn(DataGridView dataGridView1)
        {
            var lastColRef = new List<Cell>(); //Cells that have references on the delete column
            var notEmptyCells = new List<Cell>();
            if (ColCount == 1)
                return false;
            var curCount = ColCount - 1;
            for (var i = 0; i < RowCount; i++)
            {
                var name = FullName(i, curCount);
                if (_dictionary[name] != "0" && _dictionary[name] != "" && _dictionary[name] != " ")
                    notEmptyCells.Add(Grid[i][curCount]);
                if (Grid[i][curCount].PointersToThis.Count != 0) //select cells that points to deleted cell
                    lastColRef.AddRange(Grid[i][curCount].PointersToThis);
            }

            if (lastColRef.Count != 0 || notEmptyCells.Count != 0)
            {
                var errorMessage = "";
                if (notEmptyCells.Count != 0)
                {
                    errorMessage = lastColRef.Aggregate("There are not empty cells: ",
                        (current, cell) => current + string.Join(";", cell.GetName()));
                    errorMessage += "\n";
                }

                if (lastColRef.Count != 0)
                {
                    errorMessage += "There are cells that point to cells from current column:\n";
                    errorMessage = lastColRef.Aggregate(errorMessage,
                        (current, cell) => current + string.Join(";", cell.GetName()));
                }

                errorMessage += "\nAre you sure want to delete this column?";
                var res = MessageBox.Show(errorMessage, "Warning", MessageBoxButtons.YesNo);
                if (res == DialogResult.No)
                    return false;
            }

            for (var i = 0; i < RowCount; i++)
            {
                var name = FullName(i, curCount);
                _dictionary.Remove(name);
            }

            foreach (var cell in notEmptyCells)
            {
                if (cell.ReferencesFromThis == null) continue;
                foreach (var reference in cell.ReferencesFromThis)
                    reference.PointersToThis.Remove(cell);
            }

            foreach (var cell in lastColRef)
                RefreshCellAndPointers(cell, dataGridView1);
            for (var i = 0; i < RowCount; i++)
                Grid[i].RemoveAt(curCount);
            ColCount--;
            return true;

        }

        public void Save(StreamWriter sw)
        {
            sw.WriteLine(RowCount);
            sw.WriteLine(ColCount);

            foreach (var cell in Grid.SelectMany(list => list).ToList())
            {
                sw.WriteLine(cell.GetName());
                sw.WriteLine(cell.Expression);
                sw.WriteLine(cell.Value);
                if (cell.ReferencesFromThis == null)
                    sw.WriteLine("0");
                else
                {
                    sw.WriteLine(cell.ReferencesFromThis.Count);
                    foreach (var refCell in cell.ReferencesFromThis)
                        sw.WriteLine(refCell.GetName());
                }

                if (cell.PointersToThis == null)
                    sw.WriteLine("0");
                else
                {
                    sw.WriteLine(cell.PointersToThis.Count);
                    foreach (var pointCell in cell.PointersToThis)
                        sw.WriteLine(pointCell.GetName());
                }
            }
        }

        public void Open(int row, int column, StreamReader sr, DataGridView dataGridView1)
        {
            for (var i = 0; i < row; i++)
            {
                for (var j = 0; j < column; j++)
                {
                    var index = sr.ReadLine();
                    if (index == null)
                        throw (new NullReferenceException("No fucking clue what is it, rework will be needed")); // TODO: Take a look at this place
                    var expression = sr.ReadLine();
                    var value = sr.ReadLine();

                    if (expression != "")
                        _dictionary[index] = value;
                    else
                        _dictionary[index] = "";

                    var refCount = Convert.ToInt32(sr.ReadLine());
                    var newRef = new List<Cell>();
                    for (var k = 0; k < refCount; k++)
                    {
                        var refer = sr.ReadLine();
                        var curRow = _26BasedSystem.From26System(refer).Item1;
                        var curCol = _26BasedSystem.From26System(refer).Item2;

                        if (curRow < RowCount && curCol < ColCount)
                            newRef.Add(Grid[curRow][curCol]);
                    }

                    var pointCount = Convert.ToInt32(sr.ReadLine());
                    var newPoint = new List<Cell>();
                    for (var k = 0; k < pointCount; k++)
                    {
                        var point = sr.ReadLine();
                        var curRow = _26BasedSystem.From26System(point).Item1;
                        var curCol = _26BasedSystem.From26System(point).Item2;
                        newPoint.Add(Grid[curRow][curCol]);
                    }

                    Grid[i][j].SetCell(expression, value, newRef, newPoint);
                    var columnIndex = Grid[i][j].Column;
                    var rowIndex = Grid[i][j].Row;
                    dataGridView1[columnIndex, rowIndex].Value = _dictionary[index];
                }
            }
        }
    }
}