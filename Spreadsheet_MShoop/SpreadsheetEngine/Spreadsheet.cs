using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Xml;
using ExpressionEngine; // DLL Reference

namespace SpreadsheetEngine
{
    public class Spreadsheet
    {
        private class CellInstance : Cell // Instance of a cell to set protected member variables
        {
            public CellInstance(int row, int col) : base(row, col) { } // row and column
            public void SetValue(string value) // Set value of a cell
            {
                _value = value;
            }
        }

        public event PropertyChangedEventHandler CellPropertyChanged; // Cell triggers event, route it by calling this event 
        private Cell[,] _cells;
        private Dictionary<string, HashSet<string>> _references; // Store other cell references

        public Spreadsheet(int numberOfRows, int numberOfColumns) // Spreadsheet constructor
        {
            _references = new Dictionary<string, HashSet<string>>(); 
            _cells = new Cell[numberOfRows, numberOfColumns];
            for (int row = 0; row < numberOfRows; row++)
            {
                for (int col = 0; col < numberOfColumns; col++)
                {
                    CellInstance currentCell = new CellInstance(row, col);
                    currentCell.PropertyChanged += OnPropertyChanged; // subscribe to property changed
                    _cells[row, col] = currentCell;
                }
            }
        }

        public void EvalCell(Cell cell) // Evaluate a cell
        {
            CellInstance instance = cell as CellInstance;

            if (string.IsNullOrEmpty(cell.Text))
            {
                instance.SetValue(""); //set to nothing for default
                CellPropertyChanged(cell, new PropertyChangedEventArgs("Value")); // if value changed
            }

            else if (cell.Text[0] == '=' && cell.Text.Length > 1) // for an equation 
            {
                bool error = false;
                ExpTree expression = new ExpTree(); 
                expression.ExpressionString = cell.Text.Substring(1);
                List<string> expressionVar = expression.GetVar();

                foreach (string variable in expressionVar) // check for errors
                {
                    if (variable == cell.Name) // self reference
                    {
                        instance.SetValue("!(self reference)");
                        error = true; 
                        break;
                    }
                    if (GetCell(variable) == null) // cell does not exist
                    {
                        instance.SetValue("!(bad reference)");
                        error = true;
                        break;
                    }
                    if (IsCircularReference(variable, cell.Name)) // circular reference
                    {
                        instance.SetValue("!(circular reference)");
                        error = true;
                        break; 
                    }

                    Cell variableCell = GetCell(variable); 
                    double variableValue;

                    if (string.IsNullOrEmpty(variableCell.Value)) // check for empty value
                    {
                        expression.SetVar(variable, 0);
                    }

                    else if (!double.TryParse(variableCell.Value, out variableValue))
                    {
                        expression.SetVar(variable, 0);
                    }

                    else
                    {
                        expression.SetVar(variable, variableValue);
                    }
                }

                if (error) // for errors
                {
                    CellPropertyChanged(cell, new PropertyChangedEventArgs("Value"));
                    return;
                }
                //at this point, the variables are set
                instance.SetValue(expression.Eval().ToString()); // Evaluate expression and set value of Cell
                CellPropertyChanged(cell, new PropertyChangedEventArgs("Value")); 
            }

            else // if not an expression
            {
                instance.SetValue(cell.Text); // chage text of cell
                CellPropertyChanged(cell, new PropertyChangedEventArgs("Value")); 
            }

            if (_references.ContainsKey(instance.Name)) // Evaluate referenced cells
            {
                foreach (string cellName in _references[instance.Name])
                {
                    EvalCell(GetCell(cellName));
                }
            }
        }

        private bool IsCircularReference(string startCell, string currentCell) // Check for circular references
        {
            if (startCell == currentCell) 
            {
                return true;
            }

            if (!_references.ContainsKey(currentCell)) 
            {
                return false;
            }

            foreach(string reference in _references[currentCell]) 
            {
                if (reference == currentCell)
                {
                    return true;
                }

                if(IsCircularReference(startCell, reference))
                {
                    return true;
                }
            }
            return false; 
        }

        public Cell GetCell(int row, int col)
        {
            int get_row = row;
            int get_col = col;
            return _cells[row, col];
        }

        public Cell GetCell(string location) // Gets a cell given a location (example: "A1")
        {
            try
            {
                if (!Char.IsLetter(location[0])) // check starts with a character
                {
                    return null;
                }
                int col = (int)Char.ToUpper(location[0]) - 65;
                int row = Convert.ToInt16(location.Substring(1)) - 1;
                return GetCell(row, col);
            }
            catch
            {
                return null;
            }
        }

        public int RowCount
        {
            get { return _cells.GetLength(0); }
        }

        public int ColumnCount
        {
            get { return _cells.GetLength(1); }
        }

        private void RemoveReferences(string name) 
        {
            foreach (string key in _references.Keys) 
            {
                if (_references[key].Contains(name))
                {
                    _references[key].Remove(name);
                }
            }
        }

        private void AddReferences(string name, List<string> variablesReferenced)
        {
            foreach (string variable in variablesReferenced) // add reference if not already there
            {
                if (!_references.ContainsKey(variable))
                {
                    _references[variable] = new HashSet<string>(); 
                }
                _references[variable].Add(name); 
            }
        }

        public void Save(XmlWriter writer) // Save a spreadsheet
        {
            writer.WriteStartElement("Spreadsheet"); 
            foreach (Cell c in _cells)
            {
                if (!c.IsBlank) // ignore blank cells
                {
                    // Write each element
                    writer.WriteStartElement("Cell"); 
                    writer.WriteAttributeString("Name", c.Name);
                    writer.WriteElementString("Text", c.Text); 
                    writer.WriteElementString("BGColor", c.BGColor.ToString());
                    writer.WriteEndElement();
                }
            }
            writer.WriteEndElement();
        }

        public void Load(XmlElement element) // Load a spreadsheet
        {
            if (element.Name != "Spreadsheet")
            {
                return;
            }
            // Load each element
            foreach (XmlElement cellElement in element.GetElementsByTagName("Cell")) 
            {
                Cell cell = GetCell(cellElement.GetAttribute("Name").ToString());
                if (cell == null) { continue; } // avoid any null references

                XmlNodeList textList = cellElement.GetElementsByTagName("Text");
                if (textList != null)
                {
                    cell.Text = textList[0].InnerText;
                }

                XmlNodeList colorList = cellElement.GetElementsByTagName("BGColor");
                if (colorList != null)
                {
                    cell.BGColor = int.Parse(colorList[0].InnerText);
                }
            }
        }

        public void Clear() // Clear entire spreadsheet
        {
            foreach(Cell c in _cells)
            {
                c.Clear();
            }
        }

        public void OnPropertyChanged(object sender, PropertyChangedEventArgs e) // For each fired event property changed
        {
            if (e.PropertyName == "Text") // Text changed
            {
                CellInstance currentCell = sender as CellInstance;
                RemoveReferences(currentCell.Name);
                if (!string.IsNullOrEmpty(currentCell.Text) && currentCell.Text.StartsWith("=") && currentCell.Text.Length > 1) 
                {
                    ExpTree expression = new ExpTree();
                    expression.ExpressionString = currentCell.Text.Substring(1); 
                    AddReferences(currentCell.Name, expression.GetVar()); 
                }

                EvalCell(currentCell);
            }

            if (e.PropertyName == "BGColor") // Background Color changed
            {
                CellPropertyChanged(sender, new PropertyChangedEventArgs("BGColor"));
            }
        }
    }
}
