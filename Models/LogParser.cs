using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDataReader;

namespace StripTestWeb.Models
{
    public class LogResult
    {
        public string FilePath     { get; set; } = "";
        public string FileName     { get; set; } = "";
        public string LotNo        { get; set; } = "";
        public string TestDateTime { get; set; } = "";
        public int    Input        { get; set; }
        public int    Output       { get; set; }
        public double Yield        => Input > 0 ? (double)Output / Input : 0;
        public Dictionary<string, int> FailCounts { get; set; } = new();
        public int    VthLt15      { get; set; }
        public int    VthGt6       { get; set; }
        public string? Error       { get; set; }
    }

    public static class TestItems
    {
        public static readonly string[] Judged =
        {
            "VTH", "VFSD", "BVDSS 2", "BVDSS 1", "Delta3",
            "IPD-10V", "IDSS1", "IDSS2", "IDSS3"
        };
    }

    public static class LogParser
    {
        static double? ParseLimit(object? raw)
        {
            if (raw == null) return null;
            var m = Regex.Match(raw.ToString()!.Trim(), @"[-+]?\d*\.?\d+");
            return m.Success && double.TryParse(m.Value,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double v) ? v : null;
        }

        static double? ToDouble(object? raw)
        {
            if (raw == null) return null;
            if (raw is double d) return d;
            if (raw is float  f) return f;
            if (raw is int    i) return i;
            if (raw is long   l) return l;
            return double.TryParse(raw.ToString(),
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double v) ? v : null;
        }

        // Parse from a stream (uploaded file — no filesystem path needed)
        public static LogResult ParseStream(Stream stream, string fileName, string ext)
        {
            var result = new LogResult
            {
                FilePath     = fileName,
                FileName     = fileName,
                TestDateTime = DateTime.Now.ToString("yyyyMMdd-HH:mmss"),
            };
            foreach (var tc in TestItems.Judged) result.FailCounts[tc] = 0;

            try
            {
                IExcelDataReader reader = ext == ".xls"
                    ? ExcelReaderFactory.CreateBinaryReader(stream)
                    : ExcelReaderFactory.CreateOpenXmlReader(stream);

                using (reader)
                {
                    var ds = reader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
                    });

                    // Read TestDateTime from Measure!B1
                    foreach (DataTable t in ds.Tables)
                    {
                        if (!t.TableName.Equals("Measure", StringComparison.OrdinalIgnoreCase)) continue;
                        if (t.Rows.Count > 0 && t.Columns.Count > 1)
                        {
                            var raw = t.Rows[0][1]?.ToString()?.Trim();
                            if (!string.IsNullOrEmpty(raw)) result.TestDateTime = raw;
                        }
                        break;
                    }

                    // Find LoggerData sheet
                    DataTable? sheet = null;
                    foreach (DataTable t in ds.Tables)
                        if (t.TableName.Equals("LoggerData", StringComparison.OrdinalIgnoreCase))
                        { sheet = t; break; }

                    if (sheet == null) { result.Error = "Sheet 'LoggerData' not found."; return result; }

                    ParseSheet(sheet, result, fileName);
                }
            }
            catch (Exception ex) { result.Error = ex.Message; }

            return result;
        }

        static void ParseSheet(DataTable sheet, LogResult result, string fileName)
        {
            const int ROW_MIN = 3, ROW_MAX = 4, ROW_HEADER = 9, ROW_DATA = 10;

            if (sheet.Rows.Count <= ROW_HEADER)
            { result.Error = "LoggerData sheet has insufficient rows."; return; }

            var headerRow = sheet.Rows[ROW_HEADER];
            var colIdx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < sheet.Columns.Count; c++)
            {
                var v = headerRow[c]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(v) && !colIdx.ContainsKey(v)) colIdx[v] = c;
            }

            result.LotNo = Path.GetFileNameWithoutExtension(fileName);

            var minRow = sheet.Rows[ROW_MIN];
            var maxRow = sheet.Rows[ROW_MAX];
            var limits = new Dictionary<string, (double? min, double? max)>();
            foreach (var tc in TestItems.Judged)
            {
                if (!colIdx.TryGetValue(tc, out int ci)) continue;
                limits[tc] = (ParseLimit(minRow[ci]), ParseLimit(maxRow[ci]));
            }

            colIdx.TryGetValue("VTH", out int vthCol);
            int totalDies = 0, passDies = 0;

            for (int r = ROW_DATA; r < sheet.Rows.Count; r++)
            {
                var row = sheet.Rows[r];
                if (row[0] == null || row[0] == DBNull.Value ||
                    string.IsNullOrWhiteSpace(row[0].ToString())) break;

                totalDies++;

                // VTH derived (all dies, independent of cascading)
                double? vth = ToDouble(row[vthCol]);
                if (vth.HasValue) { if (vth.Value < 1.5) result.VthLt15++; if (vth.Value > 6.0) result.VthGt6++; }

                // Cascading fail
                bool eligible = true;
                foreach (var tc in TestItems.Judged)
                {
                    if (!eligible) break;
                    if (!colIdx.TryGetValue(tc, out int ci)) continue;
                    if (!limits.TryGetValue(tc, out var lim)) continue;
                    double? val = ToDouble(row[ci]);
                    bool failed = val == null
                        || (lim.min.HasValue && val.Value < lim.min.Value)
                        || (lim.max.HasValue && val.Value > lim.max.Value);
                    if (failed) { result.FailCounts[tc]++; eligible = false; }
                }
                if (eligible) passDies++;
            }

            result.Input  = totalDies;
            result.Output = passDies;
        }
    }
}
