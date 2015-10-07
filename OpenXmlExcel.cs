using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocumentFormat.OpenXml;
using System.Collections;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace OpenXmlExcel
{
    public class ExcelDocument : IDisposable
    {
        public SpreadsheetDocument document;
        public ExcelSheetCollection Sheets;

        private bool disposed = false;

        public ExcelDocument(string filename)
        {
            if (File.Exists(filename))
                document = SpreadsheetDocument.Open(filename, true);
            else
            {
                document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);
                WorkbookPart workbookpart = document.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
            }
            Sheets = new ExcelSheetCollection(this);
        }

        public int InsertSharedStringItem(string text)
        {
            SharedStringTablePart shareStringPart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            else
                shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
     
            if (shareStringPart.SharedStringTable == null)
                shareStringPart.SharedStringTable = new SharedStringTable();

            int i = 0;
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                    return i;
                i++;
            }

            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));

            return i;
        }

        public void Close()
        {
            this.Dispose();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                document.Close();
            }
            disposed = true;
        }
    }

    public class ExcelSheetCollection
    {
        private SortedDictionary<uint, ExcelSheet> List;
        public ExcelDocument document;

        public int Length { get { return this.List.Count; } }

        public ExcelSheet this[string name]
        {
            get { return this.List.Values.Where(p => String.Compare(p.Name, name, true) == 0).FirstOrDefault(); }
        }

        public ExcelSheet this[uint id]
        {
            get { if (this.List.ContainsKey(id)) return this.List[id]; else return null; }
        }

        public ExcelSheetCollection(ExcelDocument doc)
        {
            this.document = doc;
            this.List = new SortedDictionary<uint, ExcelSheet>();
            foreach (Sheet sheet in doc.document.WorkbookPart.Workbook.Descendants<Sheet>())
                this.List.Add(sheet.SheetId, new ExcelSheet(sheet, doc));
        }

        public ExcelSheet AddNew(string name = "")
        {
            ExcelSheet sheet = new ExcelSheet(document, name);
            this.List.Add(sheet.Id, sheet);
            return sheet;
        }
    }

    public class ExcelSheet
    {
        public string Name { get { return sheet.Name; } set { sheet.Name.Value = value; } }
        public uint Id { get { return sheet.SheetId; } }
        public ExcelDocument document;
        public Worksheet worksheet;
        public Sheet sheet;
        private ExcelRowCollection _rows;
        public ExcelRowCollection Rows { get { return _rows; } }

        public ExcelCell this[string address]
        {
            get
            {
                return _rows[ExcelCell.GetRowIndex(address)].Cells[ExcelCell.GetColumnName(address)];
            }
        }

        public ExcelCell this[uint row, uint col]
        {
            get
            {
                return _rows[row].Cells[col];
            }
        }

        public ExcelSheet(ExcelDocument doc, string name)
        {
            this.document = doc;
            WorksheetPart newWorksheetPart = doc.document.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet = worksheet;

            Sheets sheets = doc.document.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets == null)
                sheets = doc.document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            string relationshipId = doc.document.WorkbookPart.GetIdOfPart(newWorksheetPart);
            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;

            // Give the new worksheet a name.
            string sheetName = "Sheet";
            if (String.IsNullOrEmpty(name))
                sheetName = getSheetName(sheetName, sheets.Elements<Sheet>(), 1);
            else
                sheetName = getSheetName(name, sheets.Elements<Sheet>());
            // Append the new worksheet and associate it with the workbook.
            sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);

            _rows = new ExcelRowCollection(this);
        }

        public ExcelSheet(Sheet sheet, ExcelDocument doc)
        {
            this.sheet = sheet;
            this.document = doc;
            this.worksheet = ((WorksheetPart)(doc.document.WorkbookPart.GetPartById(sheet.Id))).Worksheet;
            _rows = new ExcelRowCollection(this);
        }

        private string getSheetName(string name, IEnumerable<Sheet> sheets, int number = 0)
        {
            string sheetName = name + (number > 0 ? number.ToString() : String.Empty);
            if(sheets.Count(s => String.Compare(s.Name, sheetName, true) == 0) > 0)
                sheetName = getSheetName(name, sheets, number + 1);
            return sheetName;
        }
    }

    public class ExcelRowCollection
    {
        private SortedDictionary<uint, ExcelRow> List;
        public ExcelSheet sheet;

        public ExcelRow this[UInt32 rowIndex] 
        { 
            get 
            {
                if (rowIndex < 1) throw new ArgumentException("Index row can not be less than 1!");
                return this.GetRow(rowIndex); 
            } 
        }

        public int Length { get { return this.List.Count; } }

        public ExcelRowCollection(ExcelSheet sheet)
        {
            this.sheet = sheet;
            this.List = new SortedDictionary<uint, ExcelRow>();
            foreach(Row row in sheet.worksheet.GetFirstChild<SheetData>().Elements<Row>())
                this.List.Add(row.RowIndex, new ExcelRow(row, sheet));
        }

        public ExcelRow GetRow(UInt32 rowIndex)
        {
            if (rowIndex < 1) throw new ArgumentException("Index row can not be less than 1!");
            if (this.List.ContainsKey(rowIndex))
                return this.List[rowIndex];

            Row newRow = new Row() { RowIndex = rowIndex };
            ExcelRow nextRow = this.List.FirstOrDefault(p => p.Key > rowIndex).Value;
            if(nextRow != null)
                sheet.worksheet.GetFirstChild<SheetData>().InsertBefore(newRow, nextRow.row);
            else
                sheet.worksheet.GetFirstChild<SheetData>().Append(newRow);
            
            ExcelRow retVal = new ExcelRow(newRow, sheet);
            this.List.Add(rowIndex, retVal);

            return retVal;
        }
    }

    public class ExcelRow
    {
        public UInt32 index;
        public Row row;
        public ExcelCellCollection Cells;
        public ExcelSheet sheet;

        public ExcelRow(Row row, ExcelSheet sheet)
        {
            this.row = row;
            this.index = row.RowIndex.Value;
            this.sheet = sheet;
            this.Cells = new ExcelCellCollection(this);
        }
    }

    public class ExcelCellCollection
    {
        private SortedDictionary<uint, ExcelCell> List;
        public ExcelRow row;

        public int Length { get { return this.List.Count; } }

        public ExcelCell this[string name]
        {
            get 
            {
                uint index = ExcelCell.GetColumnIndex(name);
                if (index < 1) throw new ArgumentException("Index column cannot be less than 1!");
                if(this.List.ContainsKey(index))
                    return this.List[index];

                return GetCell(index);
            }
        }

        public ExcelCell this[uint index]
        {
            get
            {
                if (index < 1) throw new ArgumentException("Index column cannot be less than 1!");
                if (this.List.ContainsKey(index))
                    return this.List[index];

                return GetCell(index);
            }
        }

        public ExcelCellCollection(ExcelRow row)
        {
            this.row = row;
            this.List = new SortedDictionary<uint, ExcelCell>();
            foreach (Cell cell in row.row.Elements<Cell>())
            {
                ExcelCell newCell = new ExcelCell(cell, row);
                this.List.Add(newCell.Index, newCell);
            }
        }

        public ExcelCell GetCell(uint index)
        {
            if (index < 1) throw new ArgumentException("Index column cannot be less than 1!");
            if (this.List.ContainsKey(index))
                return this.List[index];

            string address = ExcelCell.ConvertToLetter(index) + row.index.ToString();
            Cell newCell = new Cell() { CellReference = address };
            ExcelCell nextCell = this.List.FirstOrDefault(p => p.Key > index).Value;
            if (nextCell != null)
                row.row.InsertBefore(newCell, nextCell.cell);
            else
                row.row.Append(newCell);
            ExcelCell retVal = new ExcelCell(newCell, row);
            this.List.Add(retVal.Index, retVal);

            return retVal;
        }
    }

    public class ExcelCell
    {
        public string Address { get { return cell.CellReference.Value; } }
        public uint Index { get { return GetColumnIndex(this.Name); } }
        public string Name { get { return GetColumnName(cell.CellReference.Value); } }
        public ExcelRow row;
        public Cell cell;

        public string Value
        {
            get { return getCellValue(); }
            set { setCellValue(value); }
        }

        public ExcelCell(Cell cell, ExcelRow row)
        {
            this.row = row;
            this.cell = cell;
        }

        private string getCellValue()
        {
            string value = cell.InnerText;
            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        var stringTable = row.sheet.document.document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (stringTable != null)
                            value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                        break;
                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                }
            }
            return value;
        }

        private void setCellValue(string value)
        {
            int index = row.sheet.document.InsertSharedStringItem(value);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        public static string ConvertToLetter(uint iCol)
        {
           uint iAlpha, iRemainder;
           iAlpha = iCol / 27;
           iRemainder = iCol - (iAlpha * 26);
           string result = String.Empty;
           if(iAlpha > 0)
              result = ((char)(iAlpha + 64)).ToString();
           if(iRemainder > 0)
              result = result + ((char)(iRemainder + 64));
           return result;
        }

        public static string GetColumnName(string cellName)
        {
            Match match = Regex.Match(cellName, @"[A-Za-z]+");
            return match.Value;
        }

        public static uint GetRowIndex(string cellName)
        {
            Match match = Regex.Match(cellName, @"\d+");
            return uint.Parse(match.Value);
        }

        public static uint GetColumnIndex(string leters)
        {
            uint result = 0;
            foreach (char leter in leters)
                result = (((uint)leter) - 64) + (result * 26);
            return result;
        }
    }
}
