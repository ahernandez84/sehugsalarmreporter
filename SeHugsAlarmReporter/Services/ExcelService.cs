using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

using NLog;

namespace SeHugsAlarmReporter.Services
{
    class ExcelService
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public static bool GenerateReport(string fileNamePath, string reportName, DataTable[] dt, Color? reportColor = null)
        {
            try
            {
                using (ExcelPackage p = new ExcelPackage())
                {
                    p.Workbook.Properties.Author = "2019 Schneider Electric Homewood";
                    p.Workbook.Properties.Title = reportName;

                    ExcelWorksheet ws = CreateSheet(p, reportName);

                    //Merging cells and create a center heading for out table
                    ws.Cells[1, 1].Value = $"{reportName}";
                    ws.Cells[1, 1, 1, 4].Merge = true;
                    ws.Cells[1, 1, 1, 4].Style.Font.Bold = true;
                    ws.Cells[1, 1, 1, 4].Style.Font.Size = 12;
                    ws.Cells[1, 1, 1, 4].Style.Font.Name = "Calibri";
                    ws.Cells[1, 1, 1, 4].Style.Font.Color.SetColor(Color.White);
                    ws.Cells[1, 1, 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[1, 1, 1, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[1, 1, 1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[1, 1, 1, 4].Style.Fill.BackgroundColor.SetColor(reportColor ?? Color.FromArgb(99, 89, 158));
                    ws.Row(1).Height = 30;
                    ws.Row(2).Height = 21;

                    int rowIndex = 2;

                    CreateHeader(ws, ref rowIndex, dt[0], reportColor);
                    CreateData(ws, ref rowIndex, dt[0]);

                    rowIndex++; //start on next empty row

                    CreateHeader(ws, ref rowIndex, dt[1], reportColor);
                    CreateData(ws, ref rowIndex, dt[1]);

                    for (int i = 1; i <= dt[1].Columns.Count; i++)
                    {
                        if (i < 4)
                            ws.Column(i).AutoFit();
                        else
                            ws.Column(i).Width = 110;
                    }

                    var chart = (ExcelBarChart)ws.Drawings.AddChart("crtHugsAlarms", OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered);

                    chart.SetPosition(rowIndex + 2, 0, 2, 0);
                    chart.SetSize(400, 400);
                    chart.Series.Add("B8:B13", "A8:A13");
                    chart.Title.Text = "Hugs Alarms";
                    chart.Style = eChartStyle.Style31;
                    chart.Legend.Remove();

                    //string file = fileNamePath + $" from {DateTime.Now.ToString("yyyyMMdd")}.xlsx";
                    string file = fileNamePath + $".xlsx";

                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(file, bin);

                    return true;
                }
            }
            catch (Exception ex) { logger.Error(ex, "ExcelService <GenerateReport> method."); Console.WriteLine($"ExcelService: {ex.Message}"); return false; }
        }

        #region local methods
        private static ExcelWorksheet CreateSheet(ExcelPackage p, string sheetName)
        {
            p.Workbook.Worksheets.Add(sheetName);
            ExcelWorksheet ws = p.Workbook.Worksheets[1];
            ws.Name = sheetName; 
            ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

            return ws;
        }

        private static void CreateHeader(ExcelWorksheet ws, ref int rowIndex, DataTable dt, Color? reportColor = null)
        {
            int colIndex = 1;
            foreach (DataColumn dc in dt.Columns) //Creating Headings
            {
                var cell = ws.Cells[rowIndex, colIndex];

                //Setting the background color of header cells
                var fill = cell.Style.Fill;
                fill.PatternType = ExcelFillStyle.Solid;
                fill.BackgroundColor.SetColor(reportColor ?? Color.FromArgb(99, 89, 158));

                //Setting Top/left,right/bottom borders.
                var border = cell.Style.Border;
                if (colIndex == 1)
                    border.Bottom.Style = border.Top.Style = border.Left.Style = ExcelBorderStyle.Medium;
                else if (colIndex == dt.Columns.Count)
                    border.Bottom.Style = border.Top.Style = border.Right.Style = ExcelBorderStyle.Medium;
                else
                    border.Bottom.Style = border.Top.Style = ExcelBorderStyle.Medium;

                //Setting Value in cell
                cell.Value = dc.ColumnName.Contains("Column") ? "" : dc.ColumnName;

                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(Color.White);
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                colIndex++;
            }
        }

        private static void CreateData(ExcelWorksheet ws, ref int rowIndex, DataTable dt)
        {
            int colIndex = 0;
            foreach (DataRow dr in dt.Rows) // Adding Data into rows
            {
                colIndex = 1;
                rowIndex++;

                foreach (DataColumn dc in dt.Columns)
                {
                    var cell = ws.Cells[rowIndex, colIndex];

                    cell.Value = dr[dc.ColumnName];

                    //Setting borders of cell
                    var border = cell.Style.Border;
                    border.Left.Style = border.Right.Style = border.Bottom.Style = ExcelBorderStyle.Thin;
                    colIndex++;
                }
            }
        }
        #endregion


    }
}
