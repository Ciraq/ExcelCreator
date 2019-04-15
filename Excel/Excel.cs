using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.IO;

namespace Excel
{
    public static class Excel
    {
        public static void ExcelCreator()
        {
            string filepath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            string docName = $"{filepath}\\{DateTime.Now.ToString("yyyyMMddHHmmssff")}.xlsx";
            FileInfo sheetinfo = new FileInfo(docName);

            ExcelPackage pck = new ExcelPackage(sheetinfo);
            Color headercolor = ColorTranslator.FromHtml("#D3D3D3");

            var DefterMain = pck.Workbook.Worksheets.Add("defterMain");
            DefterMain.Cells["A1"].Value = "vkn";
            DefterMain.Cells["B1"].Value = "period_start";
            DefterMain.Cells["C1"].Value = "period_end";
            DefterMain.Cells["D1"].Value = "sube_kodu";

            DefterMain.Cells["A1:D1"].Style.Font.Bold = true;
            DefterMain.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            DefterMain.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(headercolor);
            DefterMain.Cells.AutoFitColumns();

            var EntryHeader = pck.Workbook.Worksheets.Add("EntryHeader");
            EntryHeader.Cells["A1"].Value = "enteredBy";
            EntryHeader.Cells["B1"].Value = "enteredDate";
            EntryHeader.Cells["C1"].Value = "enteredNumber";
            EntryHeader.Cells["D1"].Value = "entryComment";
            EntryHeader.Cells["E1"].Value = "totalDebit";
            EntryHeader.Cells["F1"].Value = "totalCredit";
            EntryHeader.Cells["G1"].Value = "entryNumberCounter";

            EntryHeader.Cells["A1:G1"].Style.Font.Bold = true;
            EntryHeader.Cells["A1:G1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            EntryHeader.Cells["A1:G1"].Style.Fill.BackgroundColor.SetColor(headercolor);
            EntryHeader.Cells.AutoFitColumns();


            var EntryDetail = pck.Workbook.Worksheets.Add("entryDetail");
            EntryDetail.Cells["A1"].Value = "lineNumber";
            EntryDetail.Cells["B1"].Value = "lineNumberCounter";
            EntryDetail.Cells["C1"].Value = "accountMainID";
            EntryDetail.Cells["D1"].Value = "accountMainDescription";
            EntryDetail.Cells["E1"].Value = "accountSubDescription";
            EntryDetail.Cells["F1"].Value = "accountSubID";
            EntryDetail.Cells["G1"].Value = "amount";
            EntryDetail.Cells["H1"].Value = "debitCreditCode";
            EntryDetail.Cells["I1"].Value = "postingDate";
            EntryDetail.Cells["J1"].Value = "documentType";
            EntryDetail.Cells["K1"].Value = "documentTypeDescription";

            EntryDetail.Cells["A1:K1"].Style.Font.Bold = true;
            EntryDetail.Cells["A1:K1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            EntryDetail.Cells["A1:K1"].Style.Fill.BackgroundColor.SetColor(headercolor);
            EntryDetail.Cells.AutoFitColumns();

            pck.Save();
        }
    }
}
