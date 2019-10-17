using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TestReadExcel
{
	public static class SpreadSheet
	{
		/// <summary>
		/// Reader of configfile
		/// </summary>
		/// <param name="path"></param>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public static DataTable GetSheetData(string path, string sheetName)
		{
			string localFile = path;
			SpreadsheetDocument excelFile = SpreadsheetDocument.Open(localFile, false);
			//Get Sheet Reference
			Sheet sheet = excelFile.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
			//Get Sheet Rows
			string relationshipId = sheet.Id.Value;
			WorksheetPart worksheetPart = (WorksheetPart)excelFile.WorkbookPart.GetPartById(relationshipId);
			Worksheet workSheet = worksheetPart.Worksheet;
			SheetData sheetData = workSheet.GetFirstChild<SheetData>();
			IEnumerable<Row> rows = sheetData.Descendants<Row>();
			//Create Datatable
			DataTable dt = new DataTable(sheet.Name);
			//Datatable Header
			foreach (Cell cell in rows.ElementAt(0))
			{
				dt.Columns.Add(GetCellValue(excelFile, cell));
			}
			//Datatable Data
			foreach (Row row in rows)
			{
				AddNewRow(excelFile, dt, row);
			}
			if (dt.Rows.Count > 0) dt.Rows.RemoveAt(0);
			excelFile.Close();
			return dt;
		}
		/// <summary>
		/// This method reads the content of each excel cell
		/// </summary>
		/// <param name="document">SpreadsheetDocument</param>
		/// <param name="cell">Cell</param>
		/// <returns>string</returns>
		private static string GetCellValue(SpreadsheetDocument document, Cell cell)
		{
			SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
			string value;
			if (cell.CellValue == null)
			{
				value = "";
			}
			else
			{
				value = cell.CellValue.InnerXml;
			}
			if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
			{
				return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
			}
			else
			{
				return value;
			}
		}
		/// <summary>
		/// This method creates a new row in DataTable
		/// </summary>
		/// <param name="excelFile"></param>
		/// <param name="dt"></param>
		/// <param name="row"></param>
		private static void AddNewRow(SpreadsheetDocument excelFile, DataTable dt, Row row)
		{
			DataRow tempRow = dt.NewRow();
			for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
			{
				tempRow[i] = GetCellValue(excelFile, row.Descendants<Cell>().ElementAt(i));
			}
			dt.Rows.Add(tempRow);
		}
	}
}
