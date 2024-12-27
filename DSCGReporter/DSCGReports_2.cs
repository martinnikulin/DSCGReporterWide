using DevExpress.Spreadsheet;
using System;
using System.Collections;
using System.Data;
using System.Linq;
using System.Diagnostics;
using System.Collections.Generic;


namespace DSCGReporter
{
    public static partial class DSCGReports_1
    {
        private static void CreateWorkbook(string excelTemplate)
        {
            workbook = new Workbook();
            workbook.LoadDocument(excelTemplate, DocumentFormat.Xltx);
        }

        private static void SaveWorkbook(string fileName)
        {
            workbook.SaveDocument(fileName, DocumentFormat.Xlsx);
        }

        private static void AddHeaders()
        {
            AddHeader("Подсчет", resColIndexCalc);
            AddHeader("Сводная", resColIndexSvod);
        }

        private static void AddHeader(string worksheetName, int colIndex)
        {
            Worksheet worksheet = workbook.Worksheets[worksheetName];
            int firstCol = AddReservesHeader(worksheet, colIndex);
            AddVariantsHeader(worksheet, firstCol);
        }

        private static void AddVariantsHeader(Worksheet worksheet, int firstCol)
        {
            foreach (string text in resHeaders)
            {
                int lastCol = AddThHeader(worksheet, firstCol);
                AddCaption(worksheet, text, 1, 1, firstCol, lastCol - 1, "");
                firstCol = lastCol;
            }
        }

        private static int AddThHeader(Worksheet worksheet, int firstCol)
        {
            DataTable variants = DataRepo.Variants;

            var thList = variants.AsEnumerable().Select(row => new { Thnn = row.Field<int>("Thnn"), ThCond = row.Field<string>("ThCond") }).Distinct();
            int lastCol = 0;
            foreach (var thRow in thList)
            {
                lastCol = AddAshHeader(worksheet, firstCol, thRow.Thnn);
                AddCaption(worksheet, thRow.ThCond, 2, 2, firstCol, lastCol - 1, " м.");
                firstCol = lastCol;
            }
            return lastCol;
        }

        private static int AddAshHeader(Worksheet worksheet, int firstCol, int Thnn)
        {
            DataTable variants = DataRepo.Variants;

            var ashList = variants.AsEnumerable().Select(row => new { Thnn = row.Field<int>("Thnn"), Ashnn = row.Field<int>("Ashnn"), AshCond = row.Field<string>("AshCond") }).Where(p => p.Thnn == Thnn);
            int lastCol = 0;
            foreach (var ashRow in ashList)
            {
                lastCol = AddCategoriesHeader(worksheet, firstCol);
                AddCaption(worksheet, ashRow.AshCond, 3, 3, firstCol, lastCol - 1, "%");
                firstCol = lastCol;
            }
            return lastCol;
        }

        private static int AddCategoriesHeader(Worksheet worksheet, int firstCol)
        {
            int i = firstCol;
            foreach (string cat in DataRepo.Categories)
            {
                worksheet[catRowIndex, i].Value = cat;
                i++;
            }
            return i;
        }

        private static void AddCaption(Worksheet worksheet, string text, int row1, int row2, int col1, int col2, string suffix)
        {
            worksheet.MergeCells(worksheet.Range.FromLTRB(col1, row1, col2, row2));
            worksheet[row1, col1].Value = text + suffix;
        }

        private static int AddReservesHeader(Worksheet worksheet, int firstCol)
        {
            int i = firstCol;
            foreach (string text in resHeaders)
            {
                AddCategoriesHeader(worksheet, i);
                AddCoalCMHeaders(worksheet, i, i + DataRepo.Categories.Count - 1, text);
                i += DataRepo.Categories.Count;
            }
            return i;
        }

        private static void AddCoalCMHeaders(Worksheet worksheet, int firstCol, int lastCol, string text)
        {
            worksheet.MergeCells(worksheet.Range.FromLTRB(firstCol, 1, lastCol, 3));
            worksheet[1, firstCol].Value = text;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private static void WriteValueToCell(Worksheet worksheet, int excelRow, int excelColumnIndex, object value)
        {
            if (!(value == DBNull.Value || value.ToString() == string.Empty))
            {
                if (value.GetType() == (typeof(string)))
                {
                    worksheet[excelRow, excelColumnIndex].Value = value.ToString();
                }
                else if (value.GetType() == (typeof(decimal)))
                {
                    worksheet[excelRow, excelColumnIndex].Value = Convert.ToDouble(value);
                }
                if (value.ToString().Substring(0, 1) == "=")
                {
                    worksheet[excelRow, excelColumnIndex].Formula = value.ToString();
                }
            }
        }
        private static void FormatRow(Worksheet worksheet, int level, int excelRow)
        {
            int resColIndex, col1, col2;
            if (worksheet.Name == "Подсчет")
            {
                resColIndex = resColIndexCalc;
                col1 = 14;
                col2 = 13;
            }
            else
            {
                resColIndex = resColIndexSvod;
                col1 = 1;
                col2 = 0;
            }
            int tableSize = resColIndex + (2 + 2 * variantsCount) * categoriesCount - 1;

            worksheet.MergeCells(worksheet.Range.FromLTRB(0, 0, tableSize, 0));

            if (level <= 7 || level >= 11)
            {
                worksheet.Range.FromLTRB(0, excelRow, 0, excelRow).Font.Bold = true;
                worksheet.Range.FromLTRB(0, excelRow, 0, excelRow).Font.Size = 11;
            }
            if (level <= 7 || level == 11 || level == 101)
            {
                worksheet.MergeCells(worksheet.Range.FromLTRB(0, excelRow, tableSize, excelRow));
            }
            if ((level >= 11 && level <= 16) || (level >= 101 && level <= 106))
                worksheet.MergeCells(worksheet.Range.FromLTRB(0, excelRow, col1, excelRow));
            if (level == 17 || level == 107)
                worksheet.MergeCells(worksheet.Range.FromLTRB(0, excelRow, col2, excelRow));
        }
    }
}
