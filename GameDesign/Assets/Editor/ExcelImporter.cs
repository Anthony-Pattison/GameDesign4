/*
This "Excel Importer" script provides a wrapper around EPPlus with convenience methods 
to make common operations on tables and named ranges easier.
This code is provided under an MIT licence:

Copyright (c) 2026 Douglas Gregory

Permission is hereby granted, free of charge, to any person obtaining a copy of this 
software and associated documentation files (the "Software"), to deal in the Software 
without restriction, including without limitation the rights to use, copy, modify, 
merge, publish, distribute, sublicense, and/or sell copies of the Software, and to 
permit persons to whom the Software is furnished to do so, subject to the following 
conditions:

The above copyright notice and this permission notice shall be included in all copies
or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR 
PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, 
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.
*/


using System.Collections.Generic;
using UnityEngine;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Table;

public class ExcelImporter
{
    /// <summary>
    /// Represents a table with named columns.
    /// </summary>
    public struct Table {
        readonly ExcelTable _source;
        
        /// <summary>
        /// Number of data rows in the table, not counting headers or totals.
        /// </summary>
        public int RowCount { get { return _source.Address.Rows - (_source.ShowTotal ? 1 : 0) - (_source.ShowHeader ? 1 : 0); } }

        /// <summary>
        /// Number of columns in the table.
        /// </summary>
        public int ColumnCount { get { return _source.Columns.Count; } }

        // TODO: Investigate why built-in ExcelTableColumnCollection[name] fails.
        readonly Dictionary<string, ExcelTableColumn> _columnLookup;        

        public Table(ExcelTable source) {
            _source = source;
            _columnLookup = new Dictionary<string, ExcelTableColumn>(source.Columns.Count);
            foreach (var column in source.Columns) {
                _columnLookup.Add(column.Name, column);
            }
        }

        /// <summary>
        /// Checks whether the given column name is present in this table.
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public bool HasColumn(string columnName) {
            return _columnLookup.ContainsKey(columnName);
        }

        /// <summary>
        /// Gets a label from the table and tries to map it to an Enumeration.
        /// </summary>
        /// <typeparam name="T">The type of the Enumeration to match.</typeparam>
        /// <param name="row">The row of the table to read, starting at 1.</param>
        /// <param name="columnName">The name of the column to read - case sensitive.</param>
        /// <param name="value">If successful, the matched Enumeration value will be stored here. Otherwise, the default value of the Enumeration is used.</param>
        /// <returns>true if a matching Enumeration value was found, false otherwise.</returns>
        public bool TryGetEnum<T>(int row, string columnName, out T value) where T:struct, System.Enum {
            string name = GetValue<string>(row, columnName);

            value = default;
            if (string.IsNullOrWhiteSpace(name)) return false;

            if (System.Enum.TryParse<T>(name, out value)) return true;
            
            Debug.LogError($"Unknown {typeof(T).Name} value '{name}' in table {_source.Name}, row {row}, column {columnName}.");
            return false;
        }

        /// <summary>
        /// Reads a value of a given type from a specific row in the named column of this table.
        /// </summary>
        /// <typeparam name="T">The type of data to expect: string, float, int, bool all work. For enums, use TryGetEnum instead.</typeparam>
        /// <param name="row">The row of the table to read, starting at 1.</param>
        /// <param name="columnName">The name of the column to read - case sensitive.</param>
        /// <returns></returns>
        public T GetValue<T>(int row, string columnName) {
            if (row < 1 || row > RowCount) {                
                Debug.LogError($"Tried to access row {row} of table '{_source.Name}'. Valid rows are 1 - {RowCount}.");
                return default;
            }

            if (!TryGetColumn(columnName, out ExcelTableColumn column)) {
                return default;
            }

            var start = _source.Address.Start;
            return _source.WorkSheet.GetValue<T>(start.Row + row, start.Column + column.Position);
        }

        /// <summary>
        /// Reads an entire column of values from the table and returns them as a new array.
        /// </summary>
        /// <typeparam name="T">The type of data to expect: string, float, int, bool all work.</typeparam>
        /// <param name="columnName">The name of the column to read - case sensitive.</param>
        /// <returns></returns>
        public T[] GetValues<T>(string columnName) {
            if (!TryGetColumn(columnName, out ExcelTableColumn firstColumn))
                return null;

            int rows = RowCount;
            var values = new T[rows];

            var start = _source.Address.Start;
            int startRow = start.Row + 1;
            int startColumn = start.Column + firstColumn.Position;

            for (int row = 0; row < rows; row++) {
                values[row] = _source.WorkSheet.GetValue<T>(startRow + row, startColumn);
            }

            return values;
        }

        /// <summary>
        /// Reads a range of consecutive columns into a new 2D array.
        /// </summary>
        /// <typeparam name="T">The type of data to expect: string, float, int, bool all work.</typeparam>
        /// <param name="startColumnName">The name of the column to start reading from - case sensitive.</param>
        /// <param name="columns">How many columns to read (width of returned array)</param>
        /// <returns></returns>
        public T[,] GetValues<T>(string startColumnName, int columns) {
            if (!TryGetColumn(startColumnName, out ExcelTableColumn firstColumn))
                return null;

            int rows = RowCount;
            var values = new T[rows, columns];

            var start = _source.Address.Start;
            int startRow = start.Row + 1;
            int startColumn = start.Column + firstColumn.Position;

            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++) {
                    values[row, column] = _source.WorkSheet.GetValue<T>(startRow + row, startColumn + column);
                }
            }
            return values;
        }

        /// <summary>
        /// Attempts to overwrite a range of consecutive columns in the table with new values. This will expand the table's row count, if necessary.
        /// </summary>
        /// <typeparam name="T">The data type being written.</typeparam>
        /// <param name="startColumnName">The name of the column to start writing into - case sensitive.</param>
        /// <param name="values">A 2D array of values to be written, organized into [rows, columns].</param>
        /// <returns>True if successful, false if there was an error.</returns>
        public bool TrySetValues<T>(string startColumnName, T[,] values) {
            if (!TryGetColumn(startColumnName, out ExcelTableColumn firstColumn))
                return false;                        

            int rows = values.GetLength(0);

            int columns = values.GetLength(1);
            var start = _source.Address.Start;
            int startRow = start.Row + 1;
            int startColumn = start.Column + firstColumn.Position;

            if (rows > RowCount) {
                _source.WorkSheet.InsertRow(startRow, rows - RowCount);
            }

            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++) {
                    _source.WorkSheet.SetValue(startRow + row, startColumn + column, values[row, column]);
                }
            }
            return true;
        }

        /// <summary>
        /// Attempts to overwrite a the values in the column with new ones, starting at row 1. This will expand the table's row count, if necessary.
        /// </summary>
        /// <typeparam name="T">The data type being written.</typeparam>
        /// <param name="columnName">The name of the column to write into - case sensitive.</param>
        /// <param name="values">An array of values to be written.</param>
        /// <param name="appendIfAbsent">If true, this will add the column to the table if no column of that name is found.</param>
        /// <returns></returns>
        public bool TrySetValues<T>(string columnName, T[] values, bool appendIfAbsent = false) {

            var start = _source.Address.Start;
            int startColumn;
            if (!TryGetColumn(columnName, out ExcelTableColumn firstColumn, !appendIfAbsent)) {
                if (!appendIfAbsent) return false;

                startColumn = start.Column + ColumnCount;
                _source.WorkSheet.InsertColumn(startColumn, 1);
                _source.WorkSheet.SetValue(_source.Address.Start.Row, startColumn, columnName);
                _columnLookup.Add(columnName, _source.Columns[_columnLookup.Count]);
                
            } else {
                startColumn = start.Column + firstColumn.Position;
            }

            int rows = values.Length;            
            int startRow = start.Row + 1;             

            if (rows > RowCount) {
                _source.WorkSheet.InsertRow(startRow, rows - RowCount);
            }

            for (int row = 0; row < rows; row++) {
                    _source.WorkSheet.SetValue(startRow + row, startColumn, values[row]);
            }

            return true;
        }

        /// <summary>
        /// Convenience method to fill a column with row numbers 1-n.
        /// </summary>
        /// <param name="columnName">Name of the column to fill.</param>
        public void NumberRows(string columnName) {
            if (!TryGetColumn(columnName, out ExcelTableColumn firstColumn))
                return;

            int rows = RowCount;
            var start = _source.Address.Start;
            int startRow = start.Row;
            int startColumn = start.Column + firstColumn.Position;
            for (int row = 1; row <= rows; row++) {
                _source.WorkSheet.SetValue(startRow + row, startColumn, row);
            }
        }

        bool TryGetColumn(string columnName, out ExcelTableColumn result, bool reportError = true) {
            if (_columnLookup.TryGetValue(columnName, out result))
                return true;
            
            if (!reportError) return false;
                
            string info = "Valid columns are...";
            string comparison = columnName.Trim().ToLowerInvariant();
            bool nearMatch = false;
            foreach (var label in _source.Columns){
                info += $" '{label.Name}'";
                nearMatch |= label.Name.Trim().ToLowerInvariant() == comparison;
            }
            if (nearMatch) info += "\n(Check capitalization and whitespace)";
            Debug.LogError($"Cannot find column named '{columnName}' in table {_source.Name}.\n{info}");
            
            return false;
        }
    }

    /// <summary>
    /// Represents a named range in the Excel workbook.
    /// </summary>
    public struct Range {
        readonly ExcelNamedRange _source;
        readonly ExcelWorksheet _sheet;

        public int RowCount { get { return _source.Rows; } }
        public int ColumnCount { get { return _source.Columns; } }

        // Creates a new Range object by wrapping the EPPlus ExcelNamedRange.
        public Range(ExcelWorksheet sheet, ExcelNamedRange source) {
            _sheet = sheet;
            _source = source;
        }

        /// <summary>
        /// Gets a label from the range and tries to map it to an Enumeration.
        /// </summary>
        /// <typeparam name="T">The type of the Enumeration to match.</typeparam>
        /// <param name="row">The row to read from: the top row of the range is 1.</param>
        /// <param name="column">The column to read from: the leftmost column in the range is 1.</param>
        /// <param name="value">The enumeration value that was read, or the default value, if mapping failed.</param>
        /// <returns>true if the mapping was successful, false otherwise.</returns>
        public bool TryGetEnum<T>(out T value, int row = 1, int column = 1) where T : struct, System.Enum
        {
            string name = GetValue<string>(row, column);
            if (System.Enum.TryParse<T>(name, out value))
                return true;

            Debug.LogError($"Unknown {typeof(T).Name} value '{name}' in range '{_source.Name}', row {row}, column {column}.");
            return false;
        }

        /// <summary>
        /// Reads a value of a given type from a specific cell in this named range.
        /// </summary>
        /// <typeparam name="T">The type of data to expect: string, float, int, bool all work. For enums, use TryGetEnum instead.</typeparam>
        /// <param name="row">The row to read from: the top row of the range is 1.</param>
        /// <param name="column">The column to read from: the leftmost column in the range is 1.</param>
        /// <returns>The value read from the range.</returns>
        public T GetValue<T>(int row = 1, int column = 1) {
            var start = _source.Start;            
            return _sheet.GetValue<T>(start.Row + row-1, start.Column + column-1);
        }

        /// <summary>
        /// Gets all values from this named range as a 2D array.
        /// </summary>
        /// <typeparam name="T">The type of data to expect: string, float, int, bool all work.</typeparam>
        /// <returns></returns>
        public T[,] GetValues<T>() {
            int rows = RowCount;
            int columns = ColumnCount;
            var values = new T[rows, columns];
            var start = _source.Start;
            int startRow = start.Row;
            int startColumn = start.Column;

            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++)
                    values[row, column] = _sheet.GetValue<T>(startRow + row, startColumn + column);
            }
            return values;
        }

        /// <summary>
        /// Writes a value into the specified cell of the named range (top-left by default).
        /// </summary>
        /// <typeparam name="T">The type of data being written.</typeparam>
        /// <param name="value">The data to write.</param>
        /// <param name="row">The row to write into: the top row of the range is 1 (default).</param>
        /// <param name="column">The column to write into: the leftmost column in the range is 1 (default).</param>
        public void SetValue<T>(T value, int row = 1, int column = 1) {
            var start = _source.Start;
            _sheet.SetValue(start.Row + row-1, start.Column + column-1, value);
        }

        /// <summary>
        /// Overwrites a named range with a 2D array of values. This will add rows/columns to the sheet if needed to accommodate the provided data.
        /// /// </summary>
        /// <typeparam name="T">The type of data to write.</typeparam>
        /// <param name="values">A 2D array of values to be written, organized into [rows, columns].</param>
        public void SetValues<T>(T[,] values) {
            int rows = values.GetLength(0);
            int columns = values.GetLength(1);

            var start = _source.Start;
            int startRow = start.Row;
            int startColumn = start.Column;

            if (rows > RowCount) {
                _sheet.InsertRow(startRow+1, rows - RowCount);
            }

            if (columns > ColumnCount) {
                _sheet.InsertColumn(startColumn+1, columns - ColumnCount);
            }

            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++)
                    _sheet.SetValue(startRow + row, startColumn + column, values[row, column]);
            }
        }

        /// <summary>
        /// Convenience method to fill the first column of the named range with the numbes 1 to [row count].
        /// </summary>
        public void NumberRows() {
            int rows = RowCount;
            var start = _source.Start;
            int startRow = start.Row-1;
            int startColumn = start.Column;
            for (int row = 1; row <= rows; row++) {
                _sheet.SetValue(startRow + row, startColumn, row);
            }
        }

        /// <summary>
        /// Convenience method to fill the first row of the named range with the numbers 1 to [column count].
        /// </summary>
        public void NumberColumns() {
            int columns = ColumnCount;
            var start = _source.Start;
            int startRow = start.Row;
            int startColumn = start.Column - 1;
            for (int column = 1; column <= columns; column++) {
                _sheet.SetValue(startRow, startColumn + column, column);
            }
        }

        /// <summary>
        /// Add rows/columns to the worksheet to expand the named range to fit the provided dimensions.
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="columns"></param>
        public void ExpandToFit(int rows, int columns) {
            var start = _source.Start;
            int startRow = start.Row;
            int startColumn = start.Column;

            if (rows > RowCount) {
                _sheet.InsertRow(startRow+1, rows - RowCount);
            }

            if (columns > ColumnCount) {
                _sheet.InsertColumn(startColumn+1, columns - ColumnCount);
            }
        }
    }

    ExcelPackage _package;

    Dictionary<string, Table> _tables = new Dictionary<string, Table>();
    Dictionary<string, Range> _namedRanges = new Dictionary<string, Range>();

    
    /// <summary>
    /// Imports an Excel file from the given path and wraps it in an editable form.
    /// </summary>
    /// <param name="filePath">Path to the .xlsx workbook. Relative paths are assumed to start at the Assets folder. Use forward slashes.</param>
    public ExcelImporter(string filePath) {        
        var path = Path.IsPathRooted(filePath) ? filePath : Path.Combine(Application.dataPath, filePath);
      
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
            _package = new ExcelPackage();
            _package.Load(stream);
        }

        var workbook = _package.Workbook;

        foreach (var sheet in workbook.Worksheets) {
            
            foreach (var table in sheet.Tables) {
                _tables.Add(table.Name, new Table(table));
            }               
        }

        foreach (var range in workbook.Names) {
            var sheetName = range.FullAddress.Substring(0, range.FullAddress.IndexOf('!'));
            if (sheetName.StartsWith("'")) {
                sheetName = sheetName.Substring(1, sheetName.Length - 2);
            }
            var sheet = workbook.Worksheets[sheetName];
            
            _namedRanges.Add(range.Name, new Range(sheet, range));
        }
    }

    /// <summary>
    /// Gets a table from the workbook by name.
    /// </summary>
    /// <param name="name">The name of the table. Case sensitive.</param>
    /// <param name="table">Populated with the table found, if any.</param>
    /// <returns>True if the table is found, false otherwise.</returns>
    public bool TryGetTable(string name, out Table table) {
        if (_tables.TryGetValue(name, out table)) {
            return true;
        }
        return false;
    }

    /// <summary>
    /// Gets a named range from the workbook by name.
    /// </summary>
    /// <param name="name">The name of the range. Case sensitive.</param>
    /// <param name="range">Populated with the range found, if any.</param>
    /// <returns>True if the range is found, false otherwise.</returns>
    public bool TryGetNamedRange(string name, out Range range) {
        if (_namedRanges.TryGetValue(name, out range)) {
            return true;
        }
        return false;
    }

    /// <summary>
    /// Saves the excel file under a new name.
    /// </summary>
    /// <param name="filePath">Path to the .xlsx workbook. Relative paths are assumed to start at the Assets folder.</param>
    public void SaveAs(string filePath) {
        var path = Path.IsPathRooted(filePath) ? filePath : Path.Combine(Application.dataPath, filePath);
        var data = _package.GetAsByteArray();
        File.WriteAllBytes(path, data);
    }    
}
