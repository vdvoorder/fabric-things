// Tabular Editor C# script to detect outliers using the IQR method.
// For numeric columns: values outside Q1 − 1.5×IQR or Q3 + 1.5×IQR.
// For string columns:  rows whose string length falls outside the length IQR fence.
// Date, boolean, and other column types are silently skipped.
// Vibe-coded by Ruben Van de Voorde with Claude Code.
//
// Instructions
// ------------
// 1. Save this script as a macro with a context of 'Column'
// 2. Configure a keyboard shortcut for the macro if using Tabular Editor 3
// 3. Select one or more columns from the SAME table and run the script
//
// Output columns
// --------------
// # Non-Blank      — total non-blank rows (denominator for %)
// [numeric columns only]
// Lower Fence      — Q1 − 1.5 × IQR
// Upper Fence      — Q3 + 1.5 × IQR
// # Low Outliers   — values below lower fence
// # High Outliers  — values above upper fence
// [string columns only]
// Lower Len        — Q1(length) − 1.5 × IQR(length)
// Upper Len        — Q3(length) + 1.5 × IQR(length)
// # Len Outliers   — rows whose string length falls outside the fence
// # Outliers       — total outliers (low + high for numeric; len for string)
// % Outliers       — outlier % with bar chart (conditional: only if any outliers found)

// ----------------------------------------------------------------------------------------------------------//
// Timing infrastructure

var _stopwatch = System.Diagnostics.Stopwatch.StartNew();

string FormatTiming()
{
    return $"({_stopwatch.ElapsedMilliseconds / 1000.0:F2}s⏱️)";
}

// ----------------------------------------------------------------------------------------------------------//
// Helper functions

int GetBarLength(double percentage)
{
    if (percentage <= 0) return 0;
    if (percentage >= 100) return 12;
    return Math.Min(12, Math.Max(1, (int)Math.Round(percentage / 100.0 * 12)));
}

string GeneratePercentageBar(int barLength)
{
    return barLength > 0 ? new string('\u2588', barLength) : "";
}

// ----------------------------------------------------------------------------------------------------------//
// Main script execution

if (Selected.Columns.Count() == 0)
{
    Info("Please select one or more columns to analyze.");
}
else
{
    var firstTable = Selected.Columns.First().Table;

    if (!Selected.Columns.All(c => c.Table == firstTable))
    {
        Info("Please select columns from the same table only.");
    }
    else
    {
        var numericCols = Selected.Columns
            .Where(c => c.DataType == DataType.Double || c.DataType == DataType.Int64 || c.DataType == DataType.Decimal)
            .ToList();
        var stringCols = Selected.Columns
            .Where(c => c.DataType == DataType.String)
            .ToList();

        if (numericCols.Count == 0 && stringCols.Count == 0)
        {
            Info("No numeric or text columns in selection. VizOutliers only analyzes numeric and text columns.");
        }
        else
        {
            string tableDaxName = firstTable.DaxObjectFullName;

            // Tuple: name, nonBlank, lower (fence or len fence), upper, lowOut, highOut, lenOut, isNumeric
            var results = new List<(string name, long nonBlank, double lower, double upper, long lowOut, long highOut, long lenOut, bool isNumeric)>();

            // Numeric columns — IQR on values
            foreach (var col in numericCols)
            {
                string colDax = col.DaxObjectFullName;
                string dax = $@"
VAR _q1      = PERCENTILE.INC({colDax}, 0.25)
VAR _q3      = PERCENTILE.INC({colDax}, 0.75)
VAR _iqr     = _q3 - _q1
VAR _lower   = _q1 - 1.5 * _iqr
VAR _upper   = _q3 + 1.5 * _iqr
VAR _nonBlank = COUNTROWS(FILTER({tableDaxName}, NOT ISBLANK({colDax})))
VAR _lowOut  = COUNTROWS(FILTER({tableDaxName}, NOT ISBLANK({colDax}) && {colDax} < _lower))
VAR _highOut = COUNTROWS(FILTER({tableDaxName}, NOT ISBLANK({colDax}) && {colDax} > _upper))
RETURN ROW(
    ""NonBlank"", _nonBlank,
    ""Lower"",    _lower,
    ""Upper"",    _upper,
    ""LowOut"",   _lowOut,
    ""HighOut"",  _highOut
)";
                try
                {
                    var result = EvaluateDax(dax) as System.Data.DataTable;
                    if (result != null && result.Rows.Count > 0)
                    {
                        var r        = result.Rows[0];
                        long nb      = r["[NonBlank]"] == DBNull.Value ? 0 : Convert.ToInt64(r["[NonBlank]"]);
                        double lower = r["[Lower]"]    == DBNull.Value ? 0 : Convert.ToDouble(r["[Lower]"]);
                        double upper = r["[Upper]"]    == DBNull.Value ? 0 : Convert.ToDouble(r["[Upper]"]);
                        long lowOut  = r["[LowOut]"]   == DBNull.Value ? 0 : Convert.ToInt64(r["[LowOut]"]);
                        long highOut = r["[HighOut]"]  == DBNull.Value ? 0 : Convert.ToInt64(r["[HighOut]"]);
                        results.Add((col.Name, nb, lower, upper, lowOut, highOut, 0, true));
                    }
                }
                catch
                {
                    results.Add((col.Name, 0, 0, 0, 0, 0, 0, true));
                }
            }

            // String columns — IQR on string length
            foreach (var col in stringCols)
            {
                string colDax = col.DaxObjectFullName;
                string dax = $@"
VAR _nonBlank  = FILTER({tableDaxName}, NOT ISBLANK({colDax}))
VAR _nonBlankCt = COUNTROWS(_nonBlank)
VAR _q1Len     = PERCENTILEX.INC(_nonBlank, LEN({colDax}), 0.25)
VAR _q3Len     = PERCENTILEX.INC(_nonBlank, LEN({colDax}), 0.75)
VAR _iqrLen    = _q3Len - _q1Len
VAR _lowerLen  = _q1Len - 1.5 * _iqrLen
VAR _upperLen  = _q3Len + 1.5 * _iqrLen
VAR _lenOut    = COUNTROWS(FILTER(_nonBlank, LEN({colDax}) < _lowerLen || LEN({colDax}) > _upperLen))
RETURN ROW(
    ""NonBlank"",  _nonBlankCt,
    ""LowerLen"",  _lowerLen,
    ""UpperLen"",  _upperLen,
    ""LenOut"",    _lenOut
)";
                try
                {
                    var result = EvaluateDax(dax) as System.Data.DataTable;
                    if (result != null && result.Rows.Count > 0)
                    {
                        var r         = result.Rows[0];
                        long nb       = r["[NonBlank]"]  == DBNull.Value ? 0 : Convert.ToInt64(r["[NonBlank]"]);
                        double lowerL = r["[LowerLen]"]  == DBNull.Value ? 0 : Convert.ToDouble(r["[LowerLen]"]);
                        double upperL = r["[UpperLen]"]  == DBNull.Value ? 0 : Convert.ToDouble(r["[UpperLen]"]);
                        long lenOut   = r["[LenOut]"]    == DBNull.Value ? 0 : Convert.ToInt64(r["[LenOut]"]);
                        results.Add((col.Name, nb, lowerL, upperL, 0, 0, lenOut, false));
                    }
                }
                catch
                {
                    results.Add((col.Name, 0, 0, 0, 0, 0, 0, false));
                }
            }

            // Pre-scan for conditional columns
            bool anyNumeric  = results.Any(r => r.isNumeric);
            bool anyString   = results.Any(r => !r.isNumeric);
            bool anyOutliers = results.Any(r => (r.isNumeric ? r.lowOut + r.highOut : r.lenOut) > 0);

            // Build output DataTable
            var outputTable = new System.Data.DataTable();

            string colHeader = $"Column {FormatTiming()}";
            outputTable.Columns.Add(colHeader,       typeof(string));
            outputTable.Columns.Add("# Non-Blank",   typeof(long));

            if (anyNumeric)
            {
                outputTable.Columns.Add("Lower Fence",    typeof(double));
                outputTable.Columns.Add("Upper Fence",    typeof(double));
                outputTable.Columns.Add("# Low Outliers", typeof(long));
                outputTable.Columns.Add("# High Outliers",typeof(long));
            }
            if (anyString)
            {
                outputTable.Columns.Add("Lower Len",      typeof(double));
                outputTable.Columns.Add("Upper Len",      typeof(double));
                outputTable.Columns.Add("# Len Outliers", typeof(long));
            }

            outputTable.Columns.Add("# Outliers", typeof(long));

            string outBarHeader = "% Outliers (Bars)" + new string('\u00A0', 5);
            if (anyOutliers)
            {
                outputTable.Columns.Add("% Outliers", typeof(double));
                outputTable.Columns.Add(outBarHeader, typeof(string));
            }

            foreach (var (name, nonBlank, lower, upper, lowOut, highOut, lenOut, isNumeric) in results)
            {
                long   totalOut = isNumeric ? lowOut + highOut : lenOut;
                double pctOut   = nonBlank > 0 ? Math.Round((double)totalOut / nonBlank * 100, 1) : 0.0;

                var row = outputTable.NewRow();
                row[colHeader]      = name;
                row["# Non-Blank"]  = nonBlank;

                if (anyNumeric)
                {
                    row["Lower Fence"]     = isNumeric ? (object)lower   : DBNull.Value;
                    row["Upper Fence"]     = isNumeric ? (object)upper   : DBNull.Value;
                    row["# Low Outliers"]  = isNumeric ? (object)lowOut  : DBNull.Value;
                    row["# High Outliers"] = isNumeric ? (object)highOut : DBNull.Value;
                }
                if (anyString)
                {
                    row["Lower Len"]       = !isNumeric ? (object)lower  : DBNull.Value;
                    row["Upper Len"]       = !isNumeric ? (object)upper  : DBNull.Value;
                    row["# Len Outliers"]  = !isNumeric ? (object)lenOut : DBNull.Value;
                }

                row["# Outliers"] = totalOut;

                if (anyOutliers)
                {
                    row["% Outliers"] = pctOut;
                    row[outBarHeader] = GeneratePercentageBar(GetBarLength(pctOut));
                }

                outputTable.Rows.Add(row);
            }

            if (outputTable.Rows.Count > 0)
                outputTable.Output();
            else
                Info("No results to display.");
        }
    }
}
