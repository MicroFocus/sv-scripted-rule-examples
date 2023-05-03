using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace OpenText.ExcelHelper {
    /// <summary>
    /// Serializable simplified representation of the Excel sheet, which supports fast lookup for a cell data based on the indexed column.
    /// </summary>
    [Serializable()]
    public class ExcelSheet {
        private string[,] data;

        private Dictionary<string, int> headerToColumnMap;
        private Dictionary<string, int> lookupValueToRowMap;

        private bool supportsLookup = false;
        private bool supportsHeader = false;

        /// <summary>
        /// Creates ExcelSheet without support for a lookup column or a header.
        /// </summary>
        public ExcelSheet(string filePath, string sheetName) : this(filePath, sheetName, false, null) {

        }

        /// <summary>
        /// Creates ExcelSheet with support for a lookup column and allows to specify if the sheet has a header.
        /// </summary>
        public ExcelSheet(string filePath, string sheetName, bool hasHeader, string lookupColumn = null) {
            // read  
            IWorkbook workbook = ReadWorkbook(filePath);

            // get value  
            ISheet sheet = workbook.GetSheet(sheetName);

            LoadSheet(sheet);

            if (hasHeader) {
                CreateHeaderMap();
                supportsHeader = true;
            }

            if (lookupColumn != null) {
                CreateLookupIndex(lookupColumn);
                supportsLookup = true;
            }
        }

        /// <summary>
        /// Creates ExcelSheet with support for a lookup column and allows to specify if the sheet has a header.
        /// </summary>
        public ExcelSheet(string filePath, string sheetName, bool hasHeader, int lookupColumn) {
            // read  
            IWorkbook workbook = ReadWorkbook(filePath);

            // get value  
            ISheet sheet = workbook.GetSheet(sheetName);

            LoadSheet(sheet);

            if (hasHeader) {
                CreateHeaderMap();
                supportsHeader = true;
            }

            CreateLookupIndex(lookupColumn);
            supportsLookup = true;
        }

        private void LoadSheet(ISheet sheet) {
            // determin number of rows
            int rowsCount = sheet.LastRowNum;

            if (rowsCount == 0) {
                data = new string[0, 0];
            }

            // determine number of columns 
            short columnsCount = sheet.GetRow(0).LastCellNum;

            data = new string[rowsCount, columnsCount];

            for (int row = 0; row < rowsCount; row++) {
                for (int column = 0; column < columnsCount; column++) {
                    ICell cell = sheet.GetRow(row)?.GetCell(column);

                    data[row, column] = cell == null ? string.Empty : GetCellValue(cell);
                }
            }
        }
        private string GetCellValue(ICell cell) {
            string retval = string.Empty;

            CellType cellType = cell.CellType;
            if (cellType == CellType.Formula) {
                cellType = cell.CachedFormulaResultType;
            }

            switch (cellType) {
                case CellType.Numeric:
                    retval = cell.NumericCellValue.ToString();
                    break;
                case CellType.String:
                    retval = cell.StringCellValue.ToString();
                    break;
                case CellType.Boolean:
                    retval = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Error:
                    retval = cell.ErrorCellValue.ToString();
                    break;
                default:
                    retval = cell.ToString();
                    break;
            }
            return retval;
        }

        private void CreateHeaderMap() {
            headerToColumnMap = new Dictionary<string, int>();

            int columnCount = data.GetLength(1);
            for (int i = 0; i < columnCount; i++) {
                headerToColumnMap.Add(data[0, i], i);
            }
        }

        private void CreateLookupIndex(string lookupHeaderName) {
            if (!supportsHeader) {
                throw new Exception("Column header navigation not supported.");
            }

            if (!headerToColumnMap.ContainsKey(lookupHeaderName)) {
                throw new Exception(GetIncorrectHeaderExceptionMessage(lookupHeaderName));
            }

            int column = headerToColumnMap[lookupHeaderName];

            CreateLookupIndex(column);
        }

        private void CreateLookupIndex(int column) {
            lookupValueToRowMap = new Dictionary<string, int>();

            int rowsCount = data.GetLength(0);
            for (int row = 0; row < rowsCount; row++) {
                string key = data[row, column];
                if (!lookupValueToRowMap.ContainsKey(key)) {
                    lookupValueToRowMap.Add(key, row);
                }
            }
        }

        /// <summary>
        /// Returns a cell value from a row identified by a lookupValue and a column identified by columnName.
        /// </summary>
        /// 
        /// <param name="lookupValue">lookup value to identify a row</param>
        /// <param name="columnName">column name from which to return a cell value</param>
        /// <returns>Cell value</returns>
        /// <exception cref="Exception">If lookup is not supported, header is not supported or header is incorrect</exception>
        public string LookupCellValue(string lookupValue, string columnName) {
            if (!supportsLookup) {
                throw new Exception("Lookup not supported.");
            }
            if (!supportsHeader) {
                throw new Exception("Column header navigation not supported.");
            }

            string retval = string.Empty;

            int column;
            if (headerToColumnMap.TryGetValue(columnName, out column)) {
                return LookupCellValue(lookupValue, column);
            }

            throw new Exception(GetIncorrectHeaderExceptionMessage(columnName));
        }

        /// <summary>
        /// Returns a cell value from the row identified by the lookupValue and column identified by column index.
        /// </summary>
        /// 
        /// <param name="lookupValue">lookup value</param>
        /// <param name="column">zero-based column index</param>
        /// <returns>Cell value or null if no row was found based on the lookup value</returns>
        /// <exception cref="Exception">If lookup is not supported</exception>
        public string LookupCellValue(string lookupValue, int column) {
            if (!supportsLookup) {
                throw new Exception("Lookup not supported.");
            }

            string retval = null;

            int row;
            if (lookupValueToRowMap.TryGetValue(lookupValue, out row)) {
                retval = data[row, column];
            }

            return retval;
        }

        /// <summary>
        /// Returns a cell value from the row identified by the row index and a column identified by a column name.
        /// </summary>
        /// 
        /// <param name="row">zero-based row index</param>
        /// <param name="columnName">column name</param>
        /// <returns>Cell value</returns>
        /// <exception cref="Exception">If header is not supported or header is incorrect</exception>
        public string GetCellValue(int row, string columnName) {
            if (!supportsHeader) {
                throw new Exception("Column header navigation not supported.");
            }

            string retval = null;

            int column;
            if (headerToColumnMap.TryGetValue(columnName, out column)) {
                retval = GetCellValue(row, column);
            } else {
                throw new Exception(GetIncorrectHeaderExceptionMessage(columnName));
            }

            return retval;
        }

        /// <summary>
        /// Gets a cell value from given row and column.
        /// </summary>
        /// <param name="row">zero-based row index</param>
        /// <param name="column">zero-based column index</param>
        /// <returns></returns>
        public string GetCellValue(int row, int column) {
            return data[row, column];
        }

        private static IWorkbook ReadWorkbook(string file) {
            IWorkbook workBook;
            using (FileStream fs = File.OpenRead(file)) {
                workBook = new XSSFWorkbook(fs);
            }
            return workBook;
        }

        private string GetIncorrectHeaderExceptionMessage(string header) {
            return $"Header [{header}] does not exist. Please use one of the following: [" + String.Join(",", headerToColumnMap.Keys.ToArray<string>()) + "]";
        }
    }
}
