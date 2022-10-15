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
        private Dictionary<string, string> dictionary = new();

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
                    dictionary.Add(newRow.Last().GetName(), "");
                }
                Grid.Add(newRow);
            }
        } //init table by 2 param
        public void Clear()
        {
            foreach (List<Cell> list in Grid)
                list.Clear();
            Grid.Clear();
            dictionary.Clear();
            RowCount = 0;
            ColCount = 0;
        }
        public string Calculate(string expression)
        {
            string res = null;
            try
            {
                res = (Calculator.Evaluate(expression)).ToString();
                if (res == "∞")
                    res = "Division by zero error";
                return res;
            }
            catch (Exception)
            {
                return "Error";
            }
        }
        public void ChangeCellWithAllPointers(int row, int col, string expression, DataGridView dataGridView1) //refresh cell value with check loops(Main func)
        {
            var currCell = Grid[row][col];
            currCell.DeletePointersAndReferences();
            currCell.expression = expression;
            currCell.new_referencesFromThis.Clear();

            if (expression != "")
            {
                if (expression[0] != '=')  //expression not formula
                {
                    currCell.value = expression;
                    dictionary[fullName(row, col)] = expression;
                    foreach (Cell cell in currCell.pointersToThis)
                    {
                        RefreshCellAndPointers(cell, dataGridView1);
                    }
                    return;
                }
            }  //expression formula
            string new_expression = ConvertReferences(row, col, expression);
            if (new_expression != "")
                new_expression = new_expression.Remove(0, 1);

            if (!currCell.CheckLoop(currCell.new_referencesFromThis))  //check new references for loop 
            {
                MessageBox.Show("There is a loop! Change the expression");
                currCell.expression = "";
                currCell.value = "";
                dataGridView1[col, row].Value = "0";
                return;
            }
            //new_references without loops
            currCell.AddPointersAndReferences();
            string val = Calculate(new_expression); //calculate ready expression

            if (val == "Error")  //cannot calculate
            {
                MessageBox.Show("Error in cell " + currCell.GetName());
                currCell.expression = "";
                currCell.value = "0";
                dataGridView1[currCell.column, currCell.row].Value = "0";
                return;
            }
            currCell.value = val;
            dictionary[fullName(row, col)] = val;
            foreach (Cell cell in currCell.pointersToThis)   //refresh all cells which has formula with currCell
                RefreshCellAndPointers(cell, dataGridView1);

        }
        private static string fullName(int row, int col)
        {
            Cell cell = new Cell(row, col);
            return cell.GetName();
        }
        public bool RefreshCellAndPointers(Cell cell, DataGridView dataGridView1) //refresh cell
        {
            cell.new_referencesFromThis.Clear();
            string new_expression = ConvertReferences(cell.row, cell.column, cell.expression); //expression without Cell Names
            new_expression = new_expression.Remove(0, 1); //remove '='
            string Value = Calculate(new_expression); //calculate ready expression

            if (Value == "Error")
            {
                MessageBox.Show("Error in cell " + cell.GetName());
                cell.expression = "";
                cell.value = "0";
                dataGridView1[cell.column, cell.row].Value = "0";
                return false;
            }
            Grid[cell.row][cell.column].value = Value;
            dictionary[fullName(cell.row, cell.column)] = Value;
            dataGridView1[cell.column, cell.row].Value = Value;

            foreach (Cell point in cell.pointersToThis)  //refresh all cells which points on current
            {
                if (!RefreshCellAndPointers(point, dataGridView1))
                    return false;
            }
            return true;
        }

        private void RefreshReferences()  //refresh only refs from each cell in all table
        {
            foreach (List<Cell> list in Grid)
            {
                foreach (Cell cell in list)
                {
                    if (cell.referencesFromThis != null)
                        cell.referencesFromThis.Clear();
                    if (cell.new_referencesFromThis != null)
                        cell.new_referencesFromThis.Clear();
                    if (cell.expression == "")
                        continue;
                    string new_expression = cell.expression;
                    if (cell.expression[0] == '=') //has formula
                    {
                        new_expression = ConvertReferences(cell.row, cell.column, cell.expression);
                        cell.referencesFromThis.AddRange(cell.new_referencesFromThis);
                    }
                }
            }
        }
        private string ConvertReferences(int row, int col, string expr)  // 5+4*AA1-->5+4*('Value of AA1) and add references
        {
            string cellNamePattern = @"[A-Z]+[0-9]+";
            Regex regex = new Regex(cellNamePattern, RegexOptions.IgnoreCase);
            var nums = new Tuple<int, int>(0, 0);

            foreach (Match match in regex.Matches(expr))
            {
                if (dictionary.ContainsKey(match.Value))  //addReference
                {
                    nums = _26BasedSystem.From26System(match.Value);
                    Grid[row][col].new_referencesFromThis.Add(Grid[nums.Item1][nums.Item2]);
                }
            }
            MatchEvaluator evaluator = ReferenceToValue;
            string new_expression = regex.Replace(expr, evaluator);
            return new_expression;
        }
        private string ReferenceToValue(Match m)   //Evaluator for converting
        {
            if (dictionary.ContainsKey(m.Value))
            {
                if (dictionary[m.Value] == "")
                    return "0";
                return dictionary[m.Value];
            }
            return m.Value;
        }

        public void AddRow(DataGridView dataGridView1)
        {
            List<Cell> newRow = new List<Cell>();
            for (int i = 0; i < ColCount; i++)
            {
                newRow.Add(new Cell(RowCount, i));
                dictionary.Add(newRow.Last().GetName(), "");
            }
            Grid.Add(newRow);
            RefreshReferences();
            RowCount++;
        }
        public void AddCol(DataGridView dataGridView1)
        {
            for (int i = 0; i < RowCount; i++)
            {
                Grid[i].Add(new Cell(i, ColCount));
                dictionary.Add(Grid[i].Last().GetName(), "");
            }
            RefreshReferences();
            ColCount++;
        }
        public bool DeleteRow(DataGridView dataGridView1)
        {
            List<Cell> lastRowRef = new List<Cell>(); //Cells that have references on the delete row
            List<Cell> notEmptyCells = new List<Cell>();
            if (RowCount == 0)
                return false;
            int curCount = RowCount - 1;
            for (int i = 0; i < ColCount; i++)
            {
                string name = fullName(curCount, i);
                if (dictionary[name] != "0" && dictionary[name] != "" && dictionary[name] != " ")
                    notEmptyCells.Add(Grid[curCount][i]);
                if (Grid[curCount][i].pointersToThis.Count != 0)  //select cells that points to deleted cell
                    lastRowRef.AddRange(Grid[curCount][i].pointersToThis);
            }

            if (lastRowRef.Count != 0 || notEmptyCells.Count != 0)
            {
                string errorMessage = "";
                if (notEmptyCells.Count != 0)
                {
                    errorMessage = "There are not empty cells: ";
                    foreach (Cell cell in notEmptyCells)
                        errorMessage += string.Join(";", cell.GetName());
                    errorMessage += "\n";
                }
                if (lastRowRef.Count != 0)
                {
                    errorMessage += "There are cells that point to cells from current Row:\n";
                    foreach (Cell cell in lastRowRef)
                        errorMessage += string.Join(";", cell.GetName());
                }
                errorMessage += "\nAre you sure want to delete this column?";
                DialogResult res = MessageBox.Show(errorMessage, "Warning", MessageBoxButtons.YesNo);
                if (res == DialogResult.No)
                    return false;
            }
            for (int i = 0; i < ColCount; i++)
            {
                string name = fullName(curCount,i);
                dictionary.Remove(name);
            }
            foreach (Cell cell in notEmptyCells)
            {
                if (cell.referencesFromThis != null)
                    foreach (Cell reference in cell.referencesFromThis)
                        reference.pointersToThis.Remove(cell);
            }
            
            foreach (Cell cell in lastRowRef)
            {
                if(cell.row != curCount)    
                    RefreshCellAndPointers(cell, dataGridView1);
            }
            Grid.RemoveAt(curCount);
            RowCount--;
            return true;
        }
        public bool DeleteColumn(DataGridView dataGridView1)
        {
            List<Cell> lastColRef = new List<Cell>(); //Cells that have references on the delete column
            List<Cell> notEmptyCells = new List<Cell>();
            if (ColCount == 1)
                return false;
            int curCount = ColCount - 1;
            for (int i = 0; i < RowCount; i++)
            {
                string name = fullName(i, curCount);
                if (dictionary[name] != "0" && dictionary[name] != "" && dictionary[name] != " ")
                    notEmptyCells.Add(Grid[i][curCount]);
                if (Grid[i][curCount].pointersToThis.Count != 0)  //select cells that points to deleted cell
                    lastColRef.AddRange(Grid[i][curCount].pointersToThis);
            }

            if (lastColRef.Count != 0 || notEmptyCells.Count != 0)
            {
                string errorMessage = "";
                if (notEmptyCells.Count != 0)
                {
                    errorMessage = "There are not empty cells: ";
                    foreach (Cell cell in lastColRef)
                        errorMessage += string.Join(";", cell.GetName());
                    errorMessage += "\n";
                }
                if (lastColRef.Count != 0)
                {
                    errorMessage += "There are cells that point to cells from current column:\n";
                    foreach (Cell cell in lastColRef)
                        errorMessage += string.Join(";", cell.GetName());
                }
                errorMessage += "\nAre you sure want to delete this column?";
                DialogResult res = MessageBox.Show(errorMessage, "Warning", MessageBoxButtons.YesNo);
                if (res == DialogResult.No)
                    return false;
            }
            for (int i = 0; i < RowCount; i++)
            {
                string name = fullName(i, curCount);
                dictionary.Remove(name);
            }
            foreach (Cell cell in notEmptyCells)
            {
                if (cell.referencesFromThis != null)
                {
                    foreach (Cell reference in cell.referencesFromThis)
                        reference.pointersToThis.Remove(cell);
                }
            }
            foreach (Cell cell in lastColRef)
                RefreshCellAndPointers(cell, dataGridView1);
            for (int i = 0; i < RowCount; i++)
                Grid[i].RemoveAt(curCount);
            ColCount--;
            return true;

        }

        public void Save(StreamWriter sw)
        {
            sw.WriteLine(RowCount);
            sw.WriteLine(ColCount);     

            foreach (List<Cell> list in Grid)
            {
                foreach (Cell cell in list)
                {
                    sw.WriteLine(cell.GetName());
                    sw.WriteLine(cell.expression);
                    sw.WriteLine(cell.value);
                    if (cell.referencesFromThis == null)
                        sw.WriteLine("0");
                    else
                    {
                        sw.WriteLine(cell.referencesFromThis.Count);
                        foreach (Cell refCell in cell.referencesFromThis)
                            sw.WriteLine(refCell.GetName());
                    }
                    if (cell.pointersToThis == null)
                        sw.WriteLine("0");
                    else
                    {
                        sw.WriteLine(cell.pointersToThis.Count);
                        foreach (Cell pointCell in cell.pointersToThis)
                            sw.WriteLine(pointCell.GetName());
                    }
                }
            }
        }
        public void Open(int row, int column, StreamReader sr, DataGridView dataGridView1)
        {
            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    var index = sr.ReadLine();
                    var expression = sr.ReadLine();
                    var value = sr.ReadLine();

                    if (expression != "")
                        dictionary[index] = value;
                    else
                        dictionary[index] = "";

                    var refCount = Convert.ToInt32(sr.ReadLine());
                    var newRef = new List<Cell>();
                    string refer;
                    for (var k = 0; k < refCount; k++)
                    {
                        refer = sr.ReadLine();
                        var curRow = _26BasedSystem.From26System(refer).Item1;
                        var curCol = _26BasedSystem.From26System(refer).Item2;

                        if (curRow < RowCount && curCol < ColCount)
                            newRef.Add(Grid[curRow][curCol]);
                    }

                    var pointCount = Convert.ToInt32(sr.ReadLine());
                    var newPoint = new List<Cell>();
                    string point;
                    for (var k = 0; k < pointCount; k++)
                    {
                        point = sr.ReadLine();
                        var curRow = _26BasedSystem.From26System(point).Item1;
                        var curCol = _26BasedSystem.From26System(point).Item2;
                        newPoint.Add(Grid[curRow][curCol]);
                    }
                    Grid[i][j].SetCell(expression, value, newRef, newPoint);
                    var columnIndex = Grid[i][j].column;
                    var rowIndex = Grid[i][j].row;
                    dataGridView1[columnIndex, rowIndex].Value = dictionary[index];
                }
            }
        }


    }

    public class Cell
    {
        public string expression { get; set; }
        public string value { get; set; }
        public int row { get; set; }
        public int column { get; set; }
        string name { get; set; }

        public List<Cell> pointersToThis = new List<Cell>();
        public List<Cell> referencesFromThis = new List<Cell>();
        public List<Cell> new_referencesFromThis = new List<Cell>();


        public Cell(int row, int column)
        {
            this.row = row;
            this.column = column;
            name = _26BasedSystem.To26System(column) + row;
            value = "0";
            expression = "";
        }

        public void SetCell(string expression, string value, List<Cell> references, List<Cell> pointers)
        {
            this.value = value;
            this.expression = expression;
            referencesFromThis.Clear();
            referencesFromThis.AddRange(references);
            pointersToThis.Clear();
            pointersToThis.AddRange(pointers);
        }

        public bool CheckLoop(List<Cell> list)  //??
        {
            if (list.Any(cell => cell.name == name))
            {
                return false;
            }
            foreach (var point in pointersToThis)
            {
                if (list.Any(cell => cell.name == point.name))
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
            foreach (var point in new_referencesFromThis)
            {
                point.pointersToThis.Add(this);
            }
            referencesFromThis = new_referencesFromThis;
        }

        public void DeletePointersAndReferences()
        {
            if (referencesFromThis == null) return;
            
            foreach (var point in referencesFromThis)
            {
                point.pointersToThis.Remove(this);
            }
            referencesFromThis = null;
        }
        public string GetName()
        {
            return name;
        }
    }
}