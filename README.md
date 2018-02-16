# OpenXmlExcel
### Class for creating and opening files Excel. Based on the Open XML SDK.

### Example:
```
---New Document----
ExcelDocument doc = new ExcelDocument(saveDlg.FileName);
ExcelSheet sheet = doc.Sheets.AddNew("MySheetTest");
ExcelCell cell = sheet["A5"];
cell.Value = "test";
cell = sheet["B6"];
cell.Value = "test";
cell = sheet["C5"];
cell.Value = "01234";
cell = sheet.Rows[7].Cells[2];
cell.Value = "testi";
cell = sheet.Rows[7].Cells[3];
cell.Value = "testi";
cell = sheet.Rows[8].Cells[2];
cell.Value = "testi";
sheet["D8"].Value = "test test";
doc.Close();
-----Open Document------
ExcelDocument doc = new ExcelDocument(openDlg.FileName);
doc.Sheets[1]["A5"].Value = "test";
ExcelSheet sheet = doc.Sheets.AddNew("Test");
sheet["D5"].Value = "test";
doc.Close();
```
