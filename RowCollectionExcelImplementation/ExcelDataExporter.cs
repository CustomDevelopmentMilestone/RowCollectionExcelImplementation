using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;


using Microsoft.Office.Interop.Excel;
using DataTable = Microsoft.Office.Interop.Excel.DataTable;

namespace VideoOS.CustomDevelopment.VodafonePlugin
{
    public class ExcelDataExporter
    {
        private Microsoft.Office.Interop.Excel.Application excelApp;

        public void LoadDataToWorkBook(string fileName, string[] data)
        {
            Workbook workbook = OpenExcelFile(fileName);

            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            worksheet.Columns.ClearFormats();
            worksheet.Rows.ClearFormats();

            int usedRowCount = GetFirstEmptyRow(worksheet);

            Range startFill = worksheet.Cells[usedRowCount, 1];
            Range endFill = worksheet.Cells[usedRowCount, data.Length];
            Range fillRange = worksheet.get_Range(startFill, endFill);
            //filling data and saving
            fillRange.Value = data;
            workbook.Save();
            CloseExcelFile(workbook);
        }

        private void CloseExcelFile(Workbook workbook)
        {
            //clean up stuffs
            workbook.Close(false, Type.Missing, Type.Missing);
            Marshal.ReleaseComObject(workbook);
            excelApp.Quit();
            Marshal.FinalReleaseComObject(excelApp);
        }

        private Workbook OpenExcelFile(string fileName)
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null)
            {
                throw new System.ApplicationException("EXCEL could not be started. Check that your office installation and project references are correct.");
            }

            excelApp.Visible = false;//to make excel app invisible
            excelApp.DisplayAlerts = false;//not to have popup dialogs

            Workbook workbook = null;
            //open the workbook if present, else create new one
            if (System.IO.File.Exists(fileName))
                workbook = excelApp.Workbooks.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            else
            {
                workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                workbook.SaveAs(fileName);
            }

            //select the first sheet        
            return workbook;
        }


        private int GetFirstEmptyRow(Worksheet sheet)
        {
            int startRow = sheet.UsedRange.Row;
            int rowCount = sheet.UsedRange.Rows.Count;
            int freshRow = startRow + rowCount;

            if (startRow == 1 && rowCount == 1)
            {
                string cellValue = (string)(sheet.Cells[1, 1] as Range).Value;

                if (IsSheetEmpty(cellValue))
                {
                    freshRow = 1;
                }
            }

            return freshRow;
        }

        private bool IsSheetEmpty(string cellValue)
        {
            return cellValue == null;
        }

        public void ConvertExcelToCSV(string excelFileName, string CSVFileName)
        {
            Workbook workbook = OpenExcelFile(excelFileName);

            workbook.SaveAs(CSVFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges,
                Type.Missing, Type.Missing);

            CloseExcelFile(workbook);
        }

        static void convertExcelToCSV(string sourceFile, string targetFile)
        {
            OleDbConnection conn = null;
            StreamWriter wrtr = null;
            OleDbCommand cmd = null;
            OleDbDataAdapter da = null;

            string fileType = Path.GetExtension(sourceFile);
            string strConn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceFile + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1\";";//for xlsx, using as default
            if (fileType.Equals(".xls", StringComparison.CurrentCultureIgnoreCase))
                strConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sourceFile + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\";";//for xls

            try
            {
                conn = new OleDbConnection(strConn);
                conn.Open();

                cmd = new OleDbCommand("SELECT * FROM [" + "sheet1" + "$]", conn);//considering sheet1 as worksheetname
                cmd.CommandType = CommandType.Text;
                wrtr = new StreamWriter(targetFile);

                da = new OleDbDataAdapter(cmd);
                var dt = new System.Data.DataTable();
                da.Fill(dt);

                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    string rowString = "";
                    for (int y = 0; y < dt.Columns.Count; y++)
                    {
                        rowString += dt.Rows[x][y].ToString() + ",";
                    }
                    wrtr.WriteLine(rowString);
                }

            }

            catch (Exception ex)
            {

                throw;
            }

            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                conn.Dispose();
                cmd.Dispose();
                da.Dispose();
                wrtr.Close();
                wrtr.Dispose();

            }

        }

        private bool IsRowPresent(Worksheet worksheet, int rowIndex, List<string> rowStringsList)
        {
            for (int compareRow = rowIndex - 1; compareRow >= 1; --compareRow)
            {
                Range compareRowFirstCell = worksheet.Cells[compareRow, 1];
                Range compareRowLastCell = worksheet.Cells[compareRow, worksheet.UsedRange.Columns.Count];
                Range compareCompleteRow = worksheet.Range[compareRowFirstCell, compareRowLastCell];
                object[,] compareRowStrings = compareCompleteRow.Value;
                List<string> compareRowStringsList = compareRowStrings.Cast<string>().ToList();
                if (rowStringsList.SequenceEqual(compareRowStringsList))
                {
                    return true;
                }
            }
            return false;
        }

        private IEnumerator<Range> Rows(Worksheet worksheet, int lastRow)
        {
            for (int index = 1; index < lastRow; index++)
            {
                Range compareRowFirstCell = worksheet.Cells[index, 1];
                Range compareRowLastCell = worksheet.Cells[index, worksheet.UsedRange.Columns.Count];
                Range compareCompleteRow = worksheet.Range[compareRowFirstCell, compareRowLastCell];
                yield return compareCompleteRow;
            }
        }

    }

    /// <summary>
    /// Something that supports maintaining a collection of rows that should not contain duplicates.
    /// Where a row is a string[]
    /// </summary>
    public abstract class RowCollection<T> : IComparable<T>
    {
        public abstract void Add(string[] row);
        public abstract void Remove(T key);
        public abstract void Contains(T key);


        public abstract int CompareTo(T other);
    }

}
