using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;

namespace SpreadsheetEngine
{
    public class Workbook // to Save and Load workbooks
    {
        private List<Spreadsheet> _sheets = new List<Spreadsheet>();
        private int _activeSheetIndex; 
        public UndoRedoSystem UndoRedoSys = new UndoRedoSystem(); 

        public Workbook()
        {
            _sheets.Add(new Spreadsheet(50, 26));
            _activeSheetIndex = 0;
        }

        public Spreadsheet activeSheet
        {
            get { return _sheets[_activeSheetIndex]; }
        }

        public bool Save(Stream stream) //Saves to XML using Stream (successful or not)
        {
            // XML Writer Settings:
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.NewLineChars = "\r\n";
            settings.NewLineOnAttributes = false;
            settings.Indent = true;

            XmlWriter writer = XmlWriter.Create(stream, settings); // Using settings
            if (writer == null) { return false; } // check for error

            writer.WriteStartElement("Workbook"); // start writing
            foreach(Spreadsheet s in _sheets) // write every element
            {
                s.Save(writer);
            }

            writer.WriteEndElement(); // end 
            writer.Close(); // close 
            return true;
        }

        public bool Load(Stream stream) // Load stream from XML file
        {
            XmlDocument document = new XmlDocument();
            try
            {
                document.Load(stream);
            }
            catch
            {
                return false;
            }

            if (document == null)
            {
                return false;
            }

            foreach (Spreadsheet s in _sheets)
            {
                s.Clear(); // clear any existing data
            }

            int sheetIterator = 0;
            foreach (XmlElement sheetElement in document.GetElementsByTagName("Spreadsheet")) // load each element
            {
                _sheets[sheetIterator].Load(sheetElement);
            }

            UndoRedoSys.Clear(); // clear undo redo stack
            return true;
        }
    }
}
