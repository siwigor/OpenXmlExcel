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
using System.Globalization;

namespace OpenXmlExcel
{
    public class ExcelDocument : IDisposable
    {
        public SpreadsheetDocument document;
        public Stylesheet stylesheet;
        public ExcelSheetCollection Sheets;
        public static uint DateTypeIndex;

        private bool disposed = false;

        public ExcelDocument(string filename)
        {
            if (File.Exists(filename))
            {
                document = SpreadsheetDocument.Open(filename, true);
                if (document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().Count() == 0)
                {
                    WorkbookStylesPart stylespart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                    stylespart.Stylesheet = GetStylesheet();
                    DateTypeIndex = 1;
                    stylesheet = stylespart.Stylesheet;
                }
                else
                {
                    WorkbookStylesPart stylespart = document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();
                    IEnumerable<CellFormat> cellformats = stylespart.Stylesheet.CellFormats.Elements<CellFormat>();
                    if (cellformats.Where(f => f.NumberFormatId == 14).Count() > 0)
                    {
                        uint dateformatindex = 0;
                        foreach (CellFormat format in cellformats)
                            if (format.NumberFormatId == 14) break; else dateformatindex++;
                        DateTypeIndex = dateformatindex;
                    }
                    else
                    {
                        stylespart.Stylesheet.CellFormats.Append(new CellFormat()
                        {
                            BorderId = 0,
                            FillId = 0,
                            FontId = 0,
                            NumberFormatId = 14,
                            FormatId = 0,
                            ApplyNumberFormat = true
                        });
                        stylespart.Stylesheet.CellFormats.Count = (uint)stylespart.Stylesheet.CellFormats.ChildElements.Count;
                        DateTypeIndex = stylespart.Stylesheet.CellFormats.Count - 1;
                    }
                    stylesheet = stylespart.Stylesheet;
                }
            }
            else
            {
                document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);
                WorkbookPart workbookpart = document.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                WorkbookStylesPart stylespart = workbookpart.AddNewPart<WorkbookStylesPart>();
                stylespart.Stylesheet = GetStylesheet();
                DateTypeIndex = 1;
                stylesheet = stylespart.Stylesheet;
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

        private static Stylesheet GetStylesheet()
        {
            Stylesheet stylesheet = new Stylesheet();
            Fonts fonts = new Fonts();
            fonts.Append(new Font()
            {
                FontName = new FontName() { Val = "Calibri" },
                FontSize = new FontSize() { Val = 11 },
                FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
            });
            fonts.Count = (uint)fonts.ChildElements.Count;
            
            Fills fills = new Fills();
            fills.Append(new Fill() { PatternFill = new PatternFill() { PatternType = PatternValues.None } });
            fills.Append(new Fill() { PatternFill = new PatternFill() { PatternType = PatternValues.Gray125 } });
            fills.Count = (uint)fills.ChildElements.Count;

            Borders borders = new Borders();
            borders.Append(new Border()
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Count = (uint)borders.ChildElements.Count;

            CellStyleFormats cellstyleformats = new CellStyleFormats();
            cellstyleformats.Append(new CellFormat()
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            });
            cellstyleformats.Count = (uint)cellstyleformats.ChildElements.Count;

            CellFormats cellformats = new CellFormats();
            cellformats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 0,
                FormatId = 0
            });
            cellformats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 14,
                FormatId = 0,
                ApplyNumberFormat = true
            });
            cellformats.Count = (uint)cellformats.ChildElements.Count;

            CellStyles cellstyles = new CellStyles();
            cellstyles.Append(new CellStyle()
            {
                Name = "Normal",
                FormatId = 0,
                BuiltinId = 0
            });
            cellstyles.Count = (uint)cellstyles.ChildElements.Count;

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellstyleformats);
            stylesheet.Append(cellformats);
            stylesheet.Append(cellstyles);

            return stylesheet;
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

        public void RemoveCell(uint rowIndex, uint colIndex)
        {
            if (_rows.Contains(rowIndex))
                _rows[rowIndex].Cells.Remove(colIndex);
        }

        public void RemoveCell(string address)
        {
            RemoveCell(ExcelCell.GetRowIndex(address), ExcelCell.GetColumnIndex(ExcelCell.GetColumnName(address)));
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

        public bool Contains(uint rowIndex)
        {
            return this.List.ContainsKey(rowIndex);
        }

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

        public bool Contains(uint colIndex)
        {
            return this.List.ContainsKey(colIndex);
        }

        public bool Contains(string colName)
        {
            return this.List.ContainsKey(ExcelCell.GetColumnIndex(colName));
        }

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

        public void Remove(uint index)
        {
            if (this.List.ContainsKey(index))
            {
                this.List[index].cell.Remove();
                this.List.Remove(index);
            }
        }

        public void Remove(string colName)
        {
            Remove(ExcelCell.GetColumnIndex(colName));
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

        public object Value
        {
            get { return getCellValue(); }
            set { setCellValue(value); }
        }

        public ExcelCell(Cell cell, ExcelRow row)
        {
            this.row = row;
            this.cell = cell;
        }

        private object getCellValue()
        {
            string cellText = cell.InnerText;
            object value = cellText;
            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        var stringTable = row.sheet.document.document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (stringTable != null)
                            value = stringTable.SharedStringTable.ElementAt(int.Parse(cellText)).InnerText;
                        break;
                    case CellValues.Boolean:
                        if(cellText == "0") value = false; else value = true;
                        break;
                    case CellValues.Number:
                        value = Convert.ToDecimal(cellText, CultureInfo.InvariantCulture);
                        break;
                    case CellValues.Date:
                        value = DateTime.FromOADate(Convert.ToDouble(cellText, CultureInfo.InvariantCulture));
                        break;
                    default:
                        value = cellText;
                        break;
                }
            }
            return value;
        }

        private void setCellValue(object value)
        {
            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (String.IsNullOrEmpty(text)) return;
            if(value is string)
            {
                int index = row.sheet.document.InsertSharedStringItem(text);
                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            }
            else if(ExcelCell.IsNumericType(value.GetType()))
            {
                cell.CellValue = new CellValue(text);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
            else
            {
                cell.CellValue = new CellValue(text);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
            
        }

        public string GetString()
        {
            string value = cell.InnerText;
            if (cell.DataType != null)
            {
                if (cell.DataType.Value == CellValues.SharedString)
                {
                    var stringTable = row.sheet.document.document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (stringTable != null)
                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            return value;
        }

        public bool GetBoolean()
        {
            return (cell.InnerText != "0");
        }

        public decimal GetNumber()
        {
            return Convert.ToDecimal(cell.InnerText, CultureInfo.InvariantCulture);
        }

        public DateTime GetDate()
        {
            return DateTime.FromOADate(Convert.ToDouble(cell.InnerText, CultureInfo.InvariantCulture));
        }

        public void SetValue(object value, ExcelCellType type)
        {
            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (String.IsNullOrEmpty(text)) return;
            switch (type)
            {
                case ExcelCellType.String:
                    cell.CellValue = new CellValue(text);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    break;
                case ExcelCellType.SharedString:
                    int index = row.sheet.document.InsertSharedStringItem(text);
                    cell.CellValue = new CellValue(index.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    break;
                case ExcelCellType.Number:
                    cell.CellValue = new CellValue(text.Replace(',', '.'));
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;
                case ExcelCellType.Date:
                    text = Convert.ToDateTime(value).ToOADate().ToString(CultureInfo.InvariantCulture);
                    cell.CellValue = new CellValue(text);
                    cell.DataType = CellValues.Number;
                    cell.StyleIndex = ExcelDocument.DateTypeIndex;
                    break;
                case ExcelCellType.Boolean:
                    text = Convert.ToInt32(value, CultureInfo.InvariantCulture).ToString();
                    cell.CellValue = new CellValue(text);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
                    break;
            }
        }

        public enum ExcelCellType { String, SharedString, Number, Date, Boolean };

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

        private static bool IsNumericType(Type type)
        {
            if (type == null)
                return false;

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.SByte:
                case TypeCode.Single:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    return true;
                case TypeCode.Object:
                    if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        return IsNumericType(Nullable.GetUnderlyingType(type));
                    }
                    return false;
            }
            return false;
        }
    }
}
