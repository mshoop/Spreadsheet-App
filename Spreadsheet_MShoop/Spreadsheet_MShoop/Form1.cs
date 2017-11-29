using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO; 
using SpreadsheetEngine; // DLL Reference

namespace Spreadsheet_MShoop
{
    public partial class Form1 : Form
    {
        public Workbook workbook = new Workbook();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            workbook.activeSheet.CellPropertyChanged += UpdateSpreadsheet; // Subscribe cell 
            dataGridView1.Columns.Clear(); // Clear columns
            for (char c = 'A'; c <= 'Z'; c++) // Set columns A-Z
            {
                string letter = c.ToString();
                dataGridView1.Columns.Add(letter, letter);
            }
            dataGridView1.RowCount = 50; // Number of rows 
            dataGridView1.RowHeadersWidth = 50; // Size of headers 
            for (int i = 0; i < 50; i++) // Set rows
            {
                dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }
            UpdateUndoRedoDescriptions(); // Update undo/redos
        }

        void UpdateSpreadsheet(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Value") // Cell value
            {
                Cell currentCell = sender as Cell;
                if (currentCell != null)
                {
                    int row = currentCell.RowIndex;
                    int col = currentCell.ColumnIndex;
                    string cellValue = currentCell.Value;
                    dataGridView1.Rows[row].Cells[col].Value = cellValue;
                }
            }

            else if (e.PropertyName == "BGColor") // Background color
            {
                Cell currentCell = sender as Cell;
                if (currentCell != null)
                {
                    int row = currentCell.RowIndex;
                    int col = currentCell.ColumnIndex;
                    int cellBG = currentCell.BGColor;
                    dataGridView1.Rows[row].Cells[col].Style.BackColor = System.Drawing.Color.FromArgb(cellBG);
                }
            }
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            Cell currentCell = workbook.activeSheet.GetCell(row, col);
            dataGridView1.Rows[row].Cells[col].Value = currentCell.Text;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            Cell currentCell = workbook.activeSheet.GetCell(row, col);
            string cellText;

            try
            {
                cellText = dataGridView1.Rows[row].Cells[col].Value.ToString();
            }
            catch (NullReferenceException)
            {
                cellText = ""; //set to empty if null
            }

            List<IUndoRedoCmd> undoCmd = new List<IUndoRedoCmd>();
            undoCmd.Add(new TextCmd(currentCell.Name, currentCell.Text));
            currentCell.Text = cellText;
            workbook.UndoRedoSys.AddUndo(new UndoRedoCollection(undoCmd, "cell text change"));
            UpdateUndoRedoDescriptions();
            dataGridView1.Rows[row].Cells[col].Value = currentCell.Value;
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            workbook.UndoRedoSys.PerformUndo(workbook.activeSheet);
            UpdateUndoRedoDescriptions();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            workbook.UndoRedoSys.PerformRedo(workbook.activeSheet);
            UpdateUndoRedoDescriptions();
        }

        private void chooseBackgroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<IUndoRedoCmd> undoCmd = new List<IUndoRedoCmd>(); 
            ColorDialog cd = new ColorDialog(); // color dialogue box
            if (cd.ShowDialog() == DialogResult.OK) // check if color is selected
            {
                int color = cd.Color.ToArgb(); // chosen color
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    Cell spreadsheetCell = workbook.activeSheet.GetCell(cell.RowIndex, cell.ColumnIndex); // get cell
                    undoCmd.Add(new BGColorCmd(spreadsheetCell.Name, spreadsheetCell.BGColor)); // add to undo list
                    spreadsheetCell.BGColor = color; //change selected BGColor
                }
                workbook.UndoRedoSys.AddUndo(new UndoRedoCollection(undoCmd, "cell background color change"));
                UpdateUndoRedoDescriptions(); 
            }
        }

        private void UpdateUndoRedoDescriptions()
        {
            ToolStripMenuItem edit = menuStrip1.Items[1] as ToolStripMenuItem;

            foreach (ToolStripItem item in edit.DropDownItems)
            {
                if (String.IsNullOrEmpty(item.Text)) { continue; }

                if (item.Text.Substring(0, 4) == "Undo") // Undo
                {
                    item.Enabled = workbook.UndoRedoSys.canUndo; // enable if possible
                    item.Text = "Undo " + workbook.UndoRedoSys.UndoDescription; // update drop down text
                }

                else if (item.Text.Substring(0, 4) == "Redo") // Redo
                {
                    item.Enabled = workbook.UndoRedoSys.canRedo;
                    item.Text = "Redo " + workbook.UndoRedoSys.RedoDescription;
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e) // file is clicked -> then save is clicked
        {
            if (dataGridView1.IsCurrentCellInEditMode) // commit change if currently editing cell
            {
                dataGridView1.EndEdit();
            }

            var saveDialog = new SaveFileDialog(); // Open save file dialog
            saveDialog.Filter = "XML files (*.xml)|*.xml"; // Only save as XML file

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                Stream stream = new FileStream(saveDialog.FileName, FileMode.Create, FileAccess.Write);
                workbook.Save(stream);
                stream.Dispose();
            }
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e) //when file is clicked -> then load is clicked
        {
            if (dataGridView1.IsCurrentCellInEditMode) // commit change if currently editing cell 
            {
                dataGridView1.EndEdit();
            }

            var loadDialog = new OpenFileDialog(); // load dialog box
            loadDialog.Filter = "XML files (*.xml)|*.xml"; // filter only for XML files

            if (loadDialog.ShowDialog() == DialogResult.OK)
            {
                Stream stream = new FileStream(loadDialog.FileName, FileMode.Open, FileAccess.Read);
                workbook.Load(stream);
                stream.Dispose();
            }
            UpdateUndoRedoDescriptions(); // after loading, clear undo/redo stacks
        }
    }
}
