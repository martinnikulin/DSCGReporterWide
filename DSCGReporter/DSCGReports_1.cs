using DevExpress.Spreadsheet;
using System;
using System.Collections;
using System.Data;
using System.Linq;
using System.Diagnostics;
using System.Collections.Generic;
using DevExpress.Export.Xl;

namespace DSCGReporter
{
    public static partial class DSCGReports_1
    {
        const int resColIndexCalc = 15;
        const int resColIndexSvod = 2;
        const int sumColIndex = 2;
        const int catRowIndex = 4; // строка Excel для категорий запасов

        public static int GdbType;
        public static int ReportType;
        public static int InterbedId;
        public static int VersionId;

        private static Workbook workbook;
        private static int variantsCount;
        private static int categoriesCount;

        private static int firstFieldIndex = 5;
        private static int lastFieldIndex = 19;
        private static int resCoalFieldIndex = 20;
        private static int resTotalFieldIndex = 21;
        private static string[] resHeaders = { "Запасы по угольным пачкам, тыс.т.", "Запасы по горной массе, тыс.т." };
        private static int firstExcelRow = 5;

        public static void CreateBalanceReport(int gdbType, int reportType, int interbedId, int versionId)
        {
            DataRepo.GdbType = gdbType;
            DataRepo.ReportType = reportType;
            DataRepo.InterbedId = interbedId;
            DataRepo.VersionId = versionId;
            DataRepo.OpenTables();

            variantsCount = DataRepo.Variants.Rows.Count;
            categoriesCount = DataRepo.Categories.Count;

            DataTable reserves = DataRepo.GetReserves();
            if (reserves.Rows.Count > 0)
            {
                CreateWorkbook("BalanceReserves.xltx");
                AddHeaders();
                WriteTables();
                SaveWorkbook("..\\Балансовые запасы.xlsx");
            }
            else
            { }
        }

        private static void WriteTables()
        {
            Worksheet detailWorksheet = workbook.Worksheets["Подсчет"];
            Worksheet totalWorksheet = workbook.Worksheets["Сводная"];

            DataTable reserves = DataRepo.GetReserves();
            try
            {
                workbook.BeginUpdate();
                int excelRowSvod = firstExcelRow;
                for (int i = 0; i <= reserves.Rows.Count - 1; i++)
                {
                    DataRow dataRow = reserves.Rows[i];
                    int level = Convert.ToInt32(dataRow["Level"]);
                    int excelRow = firstExcelRow + i;
                    if (level < 8)
                    {
                        WriteCaptionRow(detailWorksheet, dataRow, excelRow);
                        FormatRow(detailWorksheet, level, excelRow);
                    }
                    else if (level < 11)
                    {
                        WriteBlockRow(detailWorksheet, dataRow, excelRow);
                        FormatRow(detailWorksheet, level, excelRow);
                    }
                    else
                    {
                        WriteSummaryRow(dataRow, level, excelRow, excelRowSvod);
                        FormatRow(detailWorksheet, level, excelRow);
                        FormatRow(totalWorksheet, level, excelRowSvod);
                        excelRowSvod = excelRowSvod + 1;
                    }                   
                }
                detailWorksheet.GetUsedRange().Borders.SetAllBorders(XlColor.DefaultForeground, BorderLineStyle.Thin);
                totalWorksheet.GetUsedRange().Borders.SetAllBorders(XlColor.DefaultForeground, BorderLineStyle.Thin);
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        private static void WriteCaptionRow(Worksheet worksheet, DataRow dataRow, int excelRow)
        {
            WriteValueToCell(worksheet, excelRow, 0, dataRow[firstFieldIndex]);
        }

        private static void WriteBlockRow(Worksheet worksheet, DataRow dataRow, int excelRow)
        {
            // аттрибуты блока и предварительные расчеты
            int excelColumn = 0;
            for (int j = firstFieldIndex; j <= lastFieldIndex; j++)
            {
                WriteValueToCell(worksheet, excelRow, excelColumn, dataRow[j]);
                excelColumn++;
            }

            int variant = Convert.ToInt32(dataRow["variant"]);

            // запасы по углю
            excelColumn = GetResExcelColumn(dataRow, "ResCoal", false);
            WriteValueToCell(worksheet, excelRow, excelColumn, dataRow[resCoalFieldIndex]);
            if (DataRepo.SumCatIndex > 1 & dataRow[4].ToString().CompareTo("C2") < 0)
            {
                excelColumn = resColIndexCalc + DataRepo.SumCatIndex;
                WriteValueToCell(worksheet, excelRow, excelColumn, dataRow[resCoalFieldIndex]);
            }
            // запасы по углю по вариантам
            excelColumn = GetResExcelColumn(dataRow, "ResCoal", true);
            WriteVariants(worksheet, excelRow, excelColumn, dataRow[resCoalFieldIndex], variant);
            if (DataRepo.SumCatIndex > 1 & dataRow[4].ToString().CompareTo("C2") < 0)
            {
                excelColumn = resColIndexCalc + 2 * DataRepo.Categories.Count + DataRepo.SumCatIndex;
                WriteVariants(worksheet, excelRow, excelColumn, dataRow[resCoalFieldIndex], variant);
            }


            // запасы по горной массе
            excelColumn = GetResExcelColumn(dataRow, "ResTotal", false);
            WriteValueToCell(worksheet, excelRow, excelColumn, dataRow[resTotalFieldIndex]);
            if (DataRepo.SumCatIndex > 1 & dataRow[4].ToString().CompareTo("C2") < 0)
            {
                excelColumn = resColIndexCalc + DataRepo.SumCatIndex + DataRepo.Categories.Count;
                WriteValueToCell(worksheet, excelRow, excelColumn, dataRow[resTotalFieldIndex]);
            }
            // запасы по горной массе по вариантам
            excelColumn = GetResExcelColumn(dataRow, "ResTotal", true);
            WriteVariants(worksheet, excelRow, excelColumn, dataRow[resTotalFieldIndex], variant);
            if (DataRepo.SumCatIndex > 1 & dataRow[4].ToString().CompareTo("C2") < 0)
            {
                excelColumn = resColIndexCalc + DataRepo.SumCatIndex + categoriesCount * (2 + DataRepo.Variants.Rows.Count);
                WriteVariants(worksheet, excelRow, excelColumn, dataRow[resTotalFieldIndex], variant);
            }
        }

        private static void WriteSummaryRow(DataRow dataRow, int level, int excelRow, int excelRowSvod)
        {
            Worksheet detailWorksheet = workbook.Worksheets["Подсчет"];
            Worksheet totalWorksheet = workbook.Worksheets["Сводная"];

            // первое поле и марка угля
            string descr = dataRow[firstFieldIndex].ToString();

            string descr1 = (level == 11) ? descr.Substring(0, 5) : descr;
            WriteValueToCell(detailWorksheet, excelRow, 0, descr1);
            WriteValueToCell(detailWorksheet, excelRow, resColIndexCalc - 1, dataRow[lastFieldIndex]);

            descr1 = (level == 11) ? descr.Substring(6) : descr;
            WriteValueToCell(totalWorksheet, excelRowSvod, 0, descr1);
            WriteValueToCell(totalWorksheet, excelRowSvod, resColIndexSvod - 1, dataRow[lastFieldIndex]);

            string value = dataRow["ResCoal"].ToString();
            // запасы по углю, по горной массе и по вариантам
            int excelColumn, excelColumnSvod;
            if (value != "")
            {
                for (int i = 0; i < (2 + 2 * variantsCount) * categoriesCount; i++)
                {
                    excelColumn = resColIndexCalc + i;
                    string columnName = GetExcelColumnName(excelColumn + 1);
                    string formula = "=СУММ(" + value.Replace(";", ";" + columnName).Substring(2) + ")";
                    WriteValueToCell(detailWorksheet, excelRow, excelColumn, formula);
                    excelColumnSvod = resColIndexSvod + i;
                    WriteValueToCell(totalWorksheet, excelRowSvod, excelColumnSvod, "='Подсчет'!" + columnName + (excelRow + 1).ToString());
                }
            }
        }
        private static int GetResExcelColumn(DataRow dataRow, string fieldName, bool isVariantColumn)
        {
            string category = Convert.ToString(dataRow["categoryName"]);
            int icat = DataRepo.Categories.IndexOf(category);

            if (!(isVariantColumn))
            {
                return (fieldName == "ResCoal") ? resColIndexCalc + icat : resColIndexCalc + categoriesCount + icat;
            }
            else 
            {
                return (fieldName == "ResCoal") ? resColIndexCalc + 2 * categoriesCount + icat : resColIndexCalc + categoriesCount * (2 + DataRepo.Variants.Rows.Count) + icat;
            }
       
        }

        private static void WriteVariants(Worksheet worksheet, int excelRow, int firstExcelColumn, object value, int variant)
        {
            DataTable subVariants = DataRepo.SubVariants[GetVariantIndex(variant)];
            foreach (DataRow dataRow in subVariants.AsEnumerable())
            {
                int subVariant = dataRow.Field<int>("SubVariant");
                int columnShift = (GetVariantIndex(subVariant)) * DataRepo.Categories.Count;
                WriteValueToCell(worksheet, excelRow, firstExcelColumn + columnShift, value);
            }
        }

        private static int GetVariantIndex(int variant)
        {
            return variant - (DataRepo.Variants.Rows.Count) * (DataRepo.InterbedId / 100 - 1) - 1;
        }
    }
}
