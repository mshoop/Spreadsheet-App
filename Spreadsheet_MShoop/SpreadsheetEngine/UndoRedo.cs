using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    public interface IUndoRedoCmd // Undo / Redo Interface for command
    {
        IUndoRedoCmd Exec(Spreadsheet sheet);
    }

    public class BGColorCmd : IUndoRedoCmd // Background color commands
    {
        private string _cellName; 
        private int _cellColor;

        public BGColorCmd(string name, int color)
        {
            _cellName = name;   //set the name
            _cellColor = color; //set the color
        }

        public IUndoRedoCmd Exec(Spreadsheet sheet)
        {
            Cell theCell = sheet.GetCell(_cellName);    // get cell
            int oldColor = theCell.BGColor;             // store old color
            theCell.BGColor = _cellColor;               // set the color
            return new BGColorCmd(_cellName, oldColor);
        }
    }

    public class TextCmd : IUndoRedoCmd // Text edits undo/redo
    {
        private string _cellText, _cellName;
        public TextCmd(string name, string text)
        {
            _cellText = text;
            _cellName = name;
        }

        public IUndoRedoCmd Exec(Spreadsheet sheet)
        {
            Cell theCell = sheet.GetCell(_cellName);
            string oldText = theCell.Text;
            theCell.Text = _cellText;
            return new TextCmd(_cellName, oldText);
        }
    }

    public class UndoRedoCollection // Collection of Undos/Redos for commands
    {
        public string Description { get; set; } // Description of command
        private IUndoRedoCmd[] _cmds;           // store commands
        public UndoRedoCollection() { }

        public UndoRedoCollection(List<IUndoRedoCmd> cmds, string desc)
        {
            _cmds = cmds.ToArray();
            Description = desc;
        }

        public UndoRedoCollection Exec(Spreadsheet sheet)
        {
            List<IUndoRedoCmd> commands = new List<IUndoRedoCmd>();
            foreach (IUndoRedoCmd cmd in _cmds)
            {
                commands.Add(cmd.Exec(sheet));
            }
            return new UndoRedoCollection(commands, this.Description);
        }

    }

    public class UndoRedoSystem // Complete system for UndoRedo
    {
        private Stack<UndoRedoCollection> _undos = new Stack<UndoRedoCollection>(); // undo stack
        private Stack<UndoRedoCollection> _redos = new Stack<UndoRedoCollection>(); // redo stack

        public bool canUndo // check if can undo
        {
            get { return _undos.Count != 0; }
        }

        public bool canRedo // check if can redo
        {
            get { return _redos.Count != 0; }
        }

        public string UndoDescription // Description for Undo commands
        {
            get
            {
                if (canUndo)
                {
                    return _undos.Peek().Description;
                }
                return "";
            }
        }

        public string RedoDescription // Description for Redo commands
        {
            get
            {
                if (canRedo)
                {
                    return _redos.Peek().Description;
                }
                return "";
            }
        }
        
        public void AddUndo(UndoRedoCollection undo) // Add Undo comands (can move to workbook class later)
        {
            _undos.Push(undo);
            _redos.Clear();
        }

        public void PerformUndo(Spreadsheet sheet) // Perform Undo command
        {
            UndoRedoCollection undo = _undos.Pop(); 
            _redos.Push(undo.Exec(sheet)); 
        }

        public void PerformRedo(Spreadsheet sheet) // Perform Redo command
        {
            UndoRedoCollection redo = _redos.Pop();
            _undos.Push(redo.Exec(sheet));
        }

        public void Clear() // Clear undo/redo stacks
        {
            _undos.Clear();
            _redos.Clear();
        }

    }
}
