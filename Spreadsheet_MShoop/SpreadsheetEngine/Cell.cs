using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using ExpressionEngine; // DLL Reference

namespace SpreadsheetEngine
{
    public abstract class Cell : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private readonly int _row = 0;
        private readonly int _col = 0;
        private readonly string _name = "";
        private int _bgcolor = -1;
        protected string _text { get; set; }
        protected string _value { get; set; }

        public int RowIndex 
        {
            get { return _row; }
        }

        public int ColumnIndex
        {
            get { return _col; }
        }

        public string Name // Name of Cell (example: A1)
        {
            get { return _name; }
        }

        public string Text // Text typed in cell
        {
            get { return _text; }
            set
            {
                if (value == _text) { return; } // if text does not change
                else
                {
                    _text = value; // update text
                    PropertyChanged(this, new PropertyChangedEventArgs("Text")); // fire PropertyChanged event
                }
            }
        }

        public string Value
        {
            get { return _value; }
        }

        public int BGColor // Background color of a cell
        {
            get { return _bgcolor; }
            set
            {
                if (value == _bgcolor) { return; } // bg color doesn't change
                else
                {
                    _bgcolor = value; // update bg color
                    PropertyChanged(this, new PropertyChangedEventArgs("BGColor")); // fire off property changed
                }
            }
        }

        public bool IsBlank // Check if cell has text or background color
        {
            get
            {
                if (BGColor == -1 && string.IsNullOrEmpty(Text))
                {
                    return true;
                }
                return false;
            }
        }

        public Cell() { }
        public Cell(int row, int col) // Set row, column, name
        {
            _row = row;
            _col = col;
            _name += Convert.ToChar('A' + col);
            _name += (row + 1).ToString();
        }

        public void Clear() // Clear a cell
        {
            Text = "";
            BGColor = -1;
        }
    }
}
