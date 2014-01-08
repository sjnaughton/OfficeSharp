using System;
using System.Collections;
using System.Collections.Generic;
//using System.Data;
using System.Diagnostics;
using System.Net;
using System.Runtime.InteropServices.Automation;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
//using MS.Internal.ComAutomation
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
//using Microsoft.VisualBasic;


namespace OfficeSharp
{
    public class ExcelHelper : IDisposable
    {
        private dynamic _excel;
        private dynamic _workbook;

        #region "Constants"
        const int XlListObjectSourceType_xlSrcRange = 1;
        const int XlYesNoGuess_xlYes = 1;
        const int XlDVType_xlValidateList = 3;
        const int _XlDVAlertStyle_xlValidAlertStop = 1;
        #endregion

        public dynamic Excel
        {
            get
            {
                try
                {
                    if (_excel == null)
                        _excel = AutomationFactory.CreateObject("Excel.Application");
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                return _excel;
            }
        }

        public void Quit()
        {
            try
            {

                if ((_excel != null))
                {
                    _excel.DisplayAlerts = false;

                    if ((_workbook != null))
                        _workbook.Close(false);

                    _excel.Quit();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public dynamic Workbook
        {
            get
            {
                try
                {
                    if (_workbook == null)
                        _workbook = Excel.ActiveWorkbook;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                return _workbook;
            }
            set { _workbook = value; }
        }

        // Use this method to retrieve active Word session
        public bool GetExcel()
        {
            try
            {
                // If GetObject throws an exception, then Word is 
                // either not running or is not available.
                _excel = AutomationFactory.GetObject("Excel.Application");
                return true;
            }
            catch
            {
                return false;
            }
        }

        // Use this method to explicity create a new instance of Word
        public bool CreateExcel()
        {
            try
            {
                // If CreateObject throws an exception, then Word is not available.
                _excel = AutomationFactory.CreateObject("Excel.Application");
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void OpenWorkbook(string filePath)
        {
            dynamic o = null;
            o = Excel;
            if (o != null)
            {
                o.WorkBooks.Open(filePath);
            }
            else
            {
                throw new InvalidProgramException("Error: Could not open Excel spreadsheet. Please check it is installed.");
            }
        }

        public void CreateWorkbook()
        {
            dynamic o = null;
            o = Excel;
            if (o != null)
            {
                Workbook = o.Workbooks.Add();
            }
        }

        public void Cells(dynamic RowIndex, dynamic ColumnIndex, dynamic value)
        {
            if (Workbook != null)
            {
                Workbook.ActiveSheet.Cells(RowIndex, ColumnIndex).Value = value;
            }
        }

        public void CloseWorkbook()
        {
            if (Workbook != null)
            {
                Workbook.Save();
                Workbook.Close();
                Workbook = null;
            }
        }

        public void ShowWorkbook()
        {
            dynamic o = Excel;
            if (o != null)
            {
                o.Visible = true;
            }
        }

        public IEnumerable<ExcelWorkSheet> GetWorkSheets(int workBookIndex)
        {
            dynamic workBook = Excel.WorkBooks(workBookIndex);

            List<ExcelWorkSheet> workSheets = new List<ExcelWorkSheet>();
            int index = 1;
            foreach (dynamic workSheet in workBook.WorkSheets)
            {
                workSheets.Add(new ExcelWorkSheet(index, workSheet.Name));
                index = index + 1;
            }
            return workSheets;
        }

        public string[,] UsedRange(int workBookIndex, int workSheetIndex)
        {
            dynamic workBook = Excel.WorkBooks(workBookIndex);
            dynamic workSheet = workBook.WorkSheets(workSheetIndex);
            dynamic excelRange = workSheet.UsedRange;
            int columnCount = excelRange.Columns.Count;
            int rowCount = excelRange.Rows.Count;


            string[,] valueArray = new string[rowCount, columnCount];
            for (var i = 1; i <= rowCount; i++)
            {
                for (var j = 1; j <= columnCount; j++)
                {
                    valueArray[i - 1, j - 1] = excelRange(i, j).Value;
                }
            }
            return valueArray;
        }

        #region "IDisposable Support"
        // To detect redundant calls
        private bool disposedValue;

        // IDisposable
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed dynamics).
                }

                Excel.DisplayAlerts = false;
                Excel.Quit();


                // TODO: free unmanaged resources (unmanaged dynamics) and override Finalize() below.
                // TODO: set large fields to null.

                //_beforeCloseEvent.RemoveEventHandler(New ComAutomationEventHandler(AddressOf BeforeCloseEventHandler))
                //_beforeCloseEvent = Nothing
                _excel = null;
            }
            this.disposedValue = true;
        }

        // TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        //Protected Overrides Sub Finalize()
        //    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        //    Dispose(False)
        //    MyBase.Finalize()
        //End Sub

        // This code added by Visual Basic to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }

    public class ExcelWorkSheet
    {
        public ExcelWorkSheet(int index, string name)
        {
            this.Index = index;
            this.Name = name;
        }
        public int Index { get; set; }
        public string Name { get; set; }
    }
}
