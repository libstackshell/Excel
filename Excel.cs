using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ScriptStack.Runtime;
using System.Collections.ObjectModel;
using System.Text;

namespace Excel
{

    public class Excel : Model
    {
        private static ReadOnlyCollection<Routine> exportedRoutines;

        public Excel()
        {
            if (exportedRoutines != null) return;

            var routines = new List<Routine>();

            routines.Add(new Routine((Type)null, "xlsx.new"));
            routines.Add(new Routine((Type)null, "xlsx.load", (Type)null, "Loads an xlsx file from the specified path. params: path"));
            routines.Add(new Routine((Type)null, "xlsx.close", (Type)null, "close an open workbook. params: workbook handle"));

            routines.Add(new Routine((Type)null, "xlsx.add_ws", (Type)null, (Type)null, "add a worksheet to a workbook. params: workbook handle, sheet name"));
            routines.Add(new Routine((Type)null, "xlsx.get_ws", (Type)null, (Type)null, "get a worksheet from a workbook. params workbook handle, sheet name"));
            routines.Add(new Routine((Type)null, "xlsx.remove_ws", (Type)null, (Type)null, "remove a worksheet from a workbook. params: workbook handle, sheet name"));
            routines.Add(new Routine((Type)null, "xlsx.list_ws", (Type)null, "list all the worksheets from a workbook as array. params: workbook handle"));

            routines.Add(new Routine((Type)null, "xlsx.rows", (Type)null, "Total number of rows in this sheet. params: sheet"));
            routines.Add(new Routine((Type)null, "xlsx.usedrows", (Type)null, "Number of used rows in this sheet. params: sheet"));
            routines.Add(new Routine((Type)null, "xlsx.columns", (Type)null, "Total number of columns in this sheet. params: sheet"));
            routines.Add(new Routine((Type)null, "xlsx.usedcolumns", (Type)null, "Number of used columns in this sheet. params: sheet"));

            List<Type> cellParams = new List<Type>();
            cellParams.Add((Type)null);
            cellParams.Add(typeof(int));
            cellParams.Add(typeof(int));
            cellParams.Add((Type)null);
            routines.Add(new Routine((Type)null, "xlsx.set", cellParams, "Set the value of a cell. params: sheet, row, col, value"));
            routines.Add(new Routine((Type)null, "xlsx.set_formula", cellParams, "Set the value of a cell to be a formula. params: sheet, row, col, formula"));
            routines.Add(new Routine((Type)null, "xlsx.get", (Type)null, typeof(int), typeof(int), "Get the value of a cell. params: sheet, row, col"));
            routines.Add(new Routine((Type)null, "xlsx.get_formula", (Type)null, typeof(int), typeof(int), "Get the formula of a cell. params: sheet, row, col"));

            List<Type> csvParams = new List<Type>();
            csvParams.Add((Type)null);        // sheet (IXLWorksheet)
            csvParams.Add(typeof(string));    // separator, z.B. ";", ",", "\t"
            csvParams.Add(typeof(bool));      // quoteFields (always quote)
            routines.Add(new Routine((Type)null, "xlsx.tocsv", csvParams, "Get a sheet as csv string. params: sheet, separator, quoteFields"));

            exportedRoutines = routines.AsReadOnly();

        }

        public ReadOnlyCollection<Routine> Routines => exportedRoutines;

        public object Invoke(string routine, List<object> parameters) {

            if(routine == "xlsx.new")
            {
                return new XLWorkbook();
            }
            if(routine == "xlsx.load")
            {
                string path = (string)parameters[0];
                XLWorkbook wb = new XLWorkbook(path);
                return wb;
            }
            if(routine == "xlsx.close")
            {
                XLWorkbook wb = parameters[0] as XLWorkbook;
                wb.Dispose();
                return null;
            }

            if (routine == "xlsx.add_ws")
            {
                XLWorkbook wb = parameters[0] as XLWorkbook;            
                var sheet = wb.Worksheets.Add((string)parameters[1]);
                return sheet;
            }
            if (routine == "xlsx.get_ws")
            {
                XLWorkbook wb = (XLWorkbook)parameters[0];
                try
                {
                    var sheet = wb.Worksheet((string)parameters[1]);
                    return sheet;
                }
                catch
                {
                    return null;
                }
            }
            if(routine == "xlsx.remove_ws")
            {
                XLWorkbook wb = parameters[0] as XLWorkbook;
                wb.Worksheets.Delete((string)parameters[1]);
                return null;
            }
            if(routine == "xlsx.list_ws")
            {
                XLWorkbook wb = parameters[0] as XLWorkbook;
                //List<string> sheetNames = new List<string>();
                ArrayList ret = new ArrayList();
                foreach(var sheet in wb.Worksheets)
                {
                    ret.Add(sheet.Name);
                }
                return ret;
            }


            if (routine == "xlsx.get")
            {
                var sheet = parameters[0] as IXLWorksheet;
                int row = (int)parameters[1];
                int col = (int)parameters[2];
                return sheet.Cell(row, col).Value;
            }
            if (routine == "xlsx.set")
            {
                var sheet = parameters[0] as IXLWorksheet;
                int row = (int)parameters[1];
                int col = (int)parameters[2];

                switch(parameters[3].GetType().ToString())
                {

                    case "System.Char":
                        {
                            sheet.Cell(row, col).Value = (char)parameters[3];
                            break;
                        }
                    case "System.Int32":
                        {
                            sheet.Cell(row, col).Value = (int)parameters[3];
                            break;
                        }
                    case "System.Single":
                        {
                            sheet.Cell(row, col).Value = (float)parameters[3];
                            break;
                        }
                    case "System.Double":
                        {
                            sheet.Cell(row, col).Value = (double)parameters[3];
                            break;
                        }
                    case "System.Decimal":
                        {
                            sheet.Cell(row, col).Value = (decimal)parameters[3];
                            break;
                        }
                    case "System.String":
                    default:
                        {
                            sheet.Cell(row, col).Value = (string)parameters[3];
                            break;
                        }

                }

                return null;
            }
            if(routine == "xlsx.set_formula")
            {
                var sheet = parameters[0] as IXLWorksheet;
                int row = (int)parameters[1];
                int col = (int)parameters[2];
                string formula = (string)parameters[3];
                sheet.Cell(row, col).FormulaA1 = formula;
                return null;
            }
            if(routine == "xlsx.get_formula")
            {
                var sheet = parameters[0] as IXLWorksheet;
                int row = (int)parameters[1];
                int col = (int)parameters[2];
                return sheet.Cell(row, col).FormulaA1;
            }

            if (routine == "xlsx.rows")
            {
                var sheet = parameters[0] as IXLWorksheet;
                if (sheet == null) return 0;

                // Letzte Zeile mit echtem Inhalt (Wert/Formel), nicht nur Format/Validation
                var lastRow = sheet.LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 0;
                return lastRow;
            }

            if (routine == "xlsx.columns")
            {
                var sheet = parameters[0] as IXLWorksheet;
                if (sheet == null) return 0;

                var lastCol = sheet.LastColumnUsed(XLCellsUsedOptions.Contents)?.ColumnNumber() ?? 0;
                return lastCol;
            }

            if (routine == "xlsx.usedrows")
            {
                var sheet = parameters[0] as IXLWorksheet;
                return sheet.RowsUsed().Count();
            }
            if (routine == "xlsx.usedcolumns")
            {
                var sheet = parameters[0] as IXLWorksheet;
                return sheet.ColumnsUsed().Count();
            }

            if (routine == "xlsx.tocsv")
            {
                var sheet = parameters[0] as IXLWorksheet;
                if (sheet == null) return "";

                string sepStr = (parameters.Count > 1 && parameters[1] != null) ? (string)parameters[1] : ";";
                char sep = string.IsNullOrEmpty(sepStr) ? ';' : sepStr[0];

                bool alwaysQuote = (parameters.Count > 2 && parameters[2] != null) && (bool)parameters[2];

                var range = sheet.RangeUsed();
                if (range == null) return ""; // leeres Sheet

                int firstRow = range.RangeAddress.FirstAddress.RowNumber;
                int lastRow = range.RangeAddress.LastAddress.RowNumber;
                int firstCol = range.RangeAddress.FirstAddress.ColumnNumber;
                int lastCol = range.RangeAddress.LastAddress.ColumnNumber;

                var sb = new StringBuilder();

                for (int r = firstRow; r <= lastRow; r++)
                {
                    for (int c = firstCol; c <= lastCol; c++)
                    {
                        if (c > firstCol) sb.Append(sep);

                        var cell = sheet.Cell(r, c);

                        // GetFormattedString() entspricht eher dem, was Excel anzeigt (Zahlen/Datum-Format)
                        string text = cell.IsEmpty() ? "" : cell.GetFormattedString();

                        sb.Append(EscapeCsv(text, sep, alwaysQuote));
                    }

                    if (r < lastRow) sb.AppendLine();
                }

                return sb.ToString();
            }

            return null;

        }

        private static string EscapeCsv(string value, char sep, bool alwaysQuote)
        {
            value ??= "";

            bool mustQuote =
                alwaysQuote ||
                value.IndexOf(sep) >= 0 ||
                value.IndexOf('"') >= 0 ||
                value.IndexOf('\r') >= 0 ||
                value.IndexOf('\n') >= 0;

            if (!mustQuote) return value;

            // Quotes verdoppeln
            value = value.Replace("\"", "\"\"");
            return $"\"{value}\"";
        }

    }

}
