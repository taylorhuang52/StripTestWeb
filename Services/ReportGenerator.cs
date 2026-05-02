using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using StripTestWeb.Models;

namespace StripTestWeb.Services
{
    public static class ReportGenerator
    {
        static readonly Color CLR_HEADER_BG = ColorTranslator.FromHtml("#4472C4");
        static readonly Color CLR_HEADER_FG = Color.White;
        static readonly Color CLR_ROW_ODD   = ColorTranslator.FromHtml("#DCE6F1");
        static readonly Color CLR_ROW_EVEN  = Color.White;
        static readonly Color CLR_BORDER    = ColorTranslator.FromHtml("#95B3D7");
        static readonly Color CLR_WARN      = ColorTranslator.FromHtml("#FFFF00"); // >0.5%
        static readonly Color CLR_ALERT     = ColorTranslator.FromHtml("#FF0000"); // >1%

        // Returns xlsx bytes for HTTP download
        public static byte[] Generate(IEnumerable<LogResult> results, DateTime reportDate)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var pkg = new ExcelPackage();
            var ws = pkg.Workbook.Worksheets.Add("工作表1");

            int col = 1;
            void AddSingle(string label) { ws.Cells[1, col].Value = label; StyleHeader(ws.Cells[1, col]); col++; }
            void AddMerged(string label)
            {
                ws.Cells[1, col, 1, col + 1].Merge = true;
                ws.Cells[1, col].Value = label;
                StyleHeader(ws.Cells[1, col]);
                col += 2;
            }

            AddSingle("Test Date/ Time");
            AddSingle("Report Date");
            AddSingle("Lot No.");
            AddSingle("Input");
            AddSingle("Output");
            AddSingle("Yield");
            AddMerged("VTH");
            AddMerged("VTH < 1.5 V");
            AddMerged("VTH > 6 V");
            AddMerged("VFSD");
            AddMerged("BVDSS 2");
            AddMerged("BVDSS 1");
            AddMerged("Delta3");
            AddMerged("IPD-10V");
            AddMerged("IDSS1");
            AddMerged("IDSS2");
            AddMerged("IDSS3");

            int row = 2;
            foreach (var r in results)
            {
                bool odd = (row % 2 == 0);
                int c = 1;

                void WriteStr(string s)   { ws.Cells[row, c].Value = s; StyleData(ws.Cells[row, c], odd); c++; }
                void WriteDate(DateTime dt)
                {
                    ws.Cells[row, c].Value = dt.Date;
                    ws.Cells[row, c].Style.Numberformat.Format = "yyyy/MM/dd";
                    StyleData(ws.Cells[row, c], odd); c++;
                }
                void WriteInt(int v)      { ws.Cells[row, c].Value = v; StyleData(ws.Cells[row, c], odd); c++; }
                void WriteYield(double v)
                {
                    ws.Cells[row, c].Value = v;
                    ws.Cells[row, c].Style.Numberformat.Format = "0.00%";
                    StyleData(ws.Cells[row, c], odd); c++;
                }
                void WritePair(int fail, int total)
                {
                    double rate = total > 0 ? (double)fail / total : 0.0;
                    Color? hi = rate > 0.01 ? CLR_ALERT : rate > 0.005 ? CLR_WARN : (Color?)null;
                    ws.Cells[row, c].Value = fail;
                    StyleData(ws.Cells[row, c], odd, hi); c++;
                    ws.Cells[row, c].Value = rate;
                    ws.Cells[row, c].Style.Numberformat.Format = "0.00%";
                    StyleData(ws.Cells[row, c], odd, hi); c++;
                }

                WriteStr(r.TestDateTime);
                WriteDate(reportDate);
                WriteStr(r.LotNo);
                WriteInt(r.Input);
                WriteInt(r.Output);
                WriteYield(r.Yield);
                WritePair(r.FailCounts.GetValueOrDefault("VTH"),     r.Input);
                WritePair(r.VthLt15,                                  r.Input);
                WritePair(r.VthGt6,                                   r.Input);
                WritePair(r.FailCounts.GetValueOrDefault("VFSD"),    r.Input);
                WritePair(r.FailCounts.GetValueOrDefault("BVDSS 2"), r.Input);
                WritePair(r.FailCounts.GetValueOrDefault("BVDSS 1"), r.Input);
                WritePair(r.FailCounts.GetValueOrDefault("Delta3"),  r.Input);
                WritePair(r.FailCounts.GetValueOrDefault("IPD-10V"), r.Input);
                WritePair(r.FailCounts.GetValueOrDefault("IDSS1"),   r.Input);
                WritePair(r.FailCounts.GetValueOrDefault("IDSS2"),   r.Input);
                WritePair(r.FailCounts.GetValueOrDefault("IDSS3"),   r.Input);
                row++;
            }

            ws.Column(1).Width = 18; ws.Column(2).Width = 13; ws.Column(3).Width = 22;
            ws.Column(4).Width = 9;  ws.Column(5).Width = 9;  ws.Column(6).Width = 9;
            for (int i = 7; i <= 28; i++) ws.Column(i).Width = 9;
            ws.View.FreezePanes(2, 1);
            ws.Cells[1, 1, 1, col - 1].AutoFilter = true;

            return pkg.GetAsByteArray();
        }

        static void StyleHeader(ExcelRange cell)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Font.Color.SetColor(CLR_HEADER_FG);
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(CLR_HEADER_BG);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment   = ExcelVerticalAlignment.Center;
            cell.Style.WrapText = true;
            SetBorder(cell);
        }

        static void StyleData(ExcelRange cell, bool odd, Color? highlight = null)
        {
            Color bg = highlight ?? (odd ? CLR_ROW_ODD : CLR_ROW_EVEN);
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(bg);
            cell.Style.Font.Color.SetColor(highlight == CLR_ALERT ? Color.White : Color.Black);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment   = ExcelVerticalAlignment.Center;
            SetBorder(cell);
        }

        static void SetBorder(ExcelRange cell)
        {
            foreach (var b in new[] { cell.Style.Border.Top, cell.Style.Border.Bottom,
                                       cell.Style.Border.Left, cell.Style.Border.Right })
            { b.Style = ExcelBorderStyle.Thin; b.Color.SetColor(CLR_BORDER); }
        }
    }
}
