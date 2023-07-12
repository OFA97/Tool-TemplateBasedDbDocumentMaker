﻿using DbDocumentMaker.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace DbDocumentMaker.Utility
{
    static class NpoiHelper
    {
        /// <summary>
        /// Sets the content of the cell.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="location">The location.</param>
        /// <param name="cellContent">Content of the cell.</param>
        public static void SetCellContent(this ISheet sheet, Point location, string cellContent)
        {
            sheet.GetRow(location.Y).GetCell(location.X).SetCellValue(cellContent);
        }


        /// <summary>
        /// Sets the cell hyperlink.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="location">The location.</param>
        /// <param name="cellContent">Content of the cell.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        public static void SetCellHyperlink(this ISheet sheet, Point location, string cellContent, string sheetName)
        {

            var link = new XSSFHyperlink(HyperlinkType.Document);
            link.Address = ("'" + sheetName + "'!A1");
            sheet.GetRow(location.Y).GetCell(location.X).Hyperlink = link;
            sheet.GetRow(location.Y).GetCell(location.X).SetCellValue(cellContent);
        }

        /// <summary>
        /// Finds the cell location.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="cellContent">Content of the cell.</param>
        /// <returns></returns>
        public static Point? FindCellLocation(this ISheet sheet, string cellContent)
        {

            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);

                if (row == null)
                {
                    continue;
                }

                for (int j = 0; j <= row.LastCellNum; j++)
                {
                    var cell = row.GetCell(j);

                    if (cell == null)
                    {
                        continue;
                    }

                    if (cell.CellType == CellType.String)
                    {
                        if (cell.RichStringCellValue.String.Trim() == cellContent)
                        {
                            return new Point(j, i);
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Finds the index of the cell.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="cellContent">Content of the cell.</param>
        /// <returns></returns>
        public static int? FindCellIndex(this IRow row, string cellContent)
        {
            for (int j = 0; j <= row.LastCellNum; j++)
            {
                var cell = row.GetCell(j);

                if (cell == null)
                {
                    continue;
                }

                if (cell.CellType == CellType.String)
                {
                    if (cell.RichStringCellValue.String.Trim() == cellContent)
                    {
                        return j;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Removes the first match row.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="matchedContent">Content of the matched.</param>
        public static void RemoveFirstMatchRow(this ISheet sheet, string matchedContent)
        {
            var location = sheet.FindCellLocation(matchedContent);
            if (location.HasValue)
                sheet.RemoveRow(sheet.GetRow(location.Value.Y));
        }

        /// <summary>
        /// Sets the first content of the match cell.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="matchedContent">Content of the matched.</param>
        /// <param name="newContent">The new content.</param>
        public static void SetFirstMatchCellContent(this ISheet sheet, string matchedContent, string newContent)
        {
            var location = sheet.FindCellLocation(matchedContent);
            if (location.HasValue)
                sheet.SetCellContent(location.Value, newContent);
        }

        /// <summary>
        /// Sets the first match cell content in row.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="matchedContent">Content of the matched.</param>
        /// <param name="newContent">The new content.</param>
        public static void SetFirstMatchCellContentInRow(this ISheet sheet, int rowIndex, string matchedContent, string newContent)
        {
            if (sheet.LastRowNum > rowIndex)
            {
                var columnIndex = sheet.GetRow(rowIndex).FindCellIndex(matchedContent);
                if (columnIndex.HasValue)
                    sheet.SetCellContent(new Point(columnIndex.Value, rowIndex), newContent);
            }
        }

        /// <summary>
        /// Sets the first match cell hyperlink in row.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="matchedContent">Content of the matched.</param>
        /// <param name="newContent">The new content.</param>
        /// <param name="linkedSheetName">Name of the linked sheet.</param>
        public static void SetFirstMatchCellHyperlinkInRow(this ISheet sheet, int rowIndex, string matchedContent, string newContent, string linkedSheetName)
        {
            if (sheet.LastRowNum > rowIndex)
            {
                var columnIndex = sheet.GetRow(rowIndex).FindCellIndex(matchedContent);
                if (columnIndex.HasValue)
                    sheet.SetCellHyperlink(new Point(columnIndex.Value, rowIndex), newContent, linkedSheetName);
            }
        }

        public static void ReadExcelFile(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(fs);

                for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
                {
                    ISheet sheet = workbook.GetSheetAt(sheetIndex);

                    DescMap(sheet,
                        (sheet.SheetName == "Table List") ? "Table" : "Column");
                }
            }
        }

        private static void DescMap(ISheet sheet, string Type)
        {
            var Location_Name = sheet.FindCellLocation($"{Type} Name");
            var Location_Desc = sheet.FindCellLocation($"{Type} Description");

            int beginCell_Name = Location_Name.Value.X;
            int beginCell_Desc = Location_Desc.Value.X;
            int beginRow = Location_Name.Value.Y + 1;

            List<object> list = new List<object>();

            for (int row = beginRow; row <= sheet.LastRowNum; row++)
            {
                string name = sheet.GetRow(row).GetCell(beginCell_Name).ToString();
                string desc = sheet.GetRow(row).GetCell(beginCell_Desc).ToString();

                if (string.IsNullOrWhiteSpace(desc)) continue;

                if (Type == "Table")
                {
                    list.Add(new Table() { TableName = name, Description = desc });
                }

                if (Type == "Column")
                {
                    string TableName = sheet.GetRow(0).GetCell(2).ToString();
                    list.Add(new Column() { TableName = TableName, ColumnName = name, Description = desc });
                }
            }

            Export.Procedure(list);
        }
    }
}
