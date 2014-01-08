using System;
using System.Collections;
using System.Collections.Generic;
//using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Automation;
using System.Security;
using System.Windows;
using System.Windows.Controls;
using Microsoft.LightSwitch;
using Microsoft.LightSwitch.Client;
using Microsoft.LightSwitch.Details;
using Microsoft.LightSwitch.Framework;
using Microsoft.LightSwitch.Model;
using Microsoft.LightSwitch.Presentation;
using Microsoft.LightSwitch.Presentation.Extensions;
using Microsoft.LightSwitch.Runtime.Shell.Framework;
using Microsoft.LightSwitch.Sdk.Proxy;
using Microsoft.LightSwitch.Threading;
using Microsoft.VisualStudio.ExtensibilityHosting;
//using System.Threading.Tasks;

namespace OfficeSharp
{

    public static class Excel
    {
        private static string[,] _excelDocRange;
        private static IVisualCollection _collection;
        private static List<ColumnMapping> _columnMappings;
        private static Dictionary<string, IEntityObject> navProperties = new Dictionary<string, IEntityObject>();
        private static FileInfo _fileInfo;

        #region "Export Overloads"
        // exports a collection to a workbook starting at the specified location
        public static dynamic Export(IVisualCollection collection, string Workbook, string Worksheet, string Range)
        {
            ExcelHelper xlProxy = new ExcelHelper();
            dynamic wb = null;
            dynamic ws = null;
            dynamic rg = null;
            dynamic result = null;

            try
            {
                wb = xlProxy.Excel.Workbooks.Open(Workbook);
                if ((wb != null))
                {
                    ws = wb.Worksheets(Worksheet);

                    if ((ws != null))
                    {
                        rg = ws.Range(Range);

                        List<string> columnNames = new List<string>();

                        if (collection.Count > 0)
                        {
                            // get column properties
                            IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();

                            int columnCounter = 0;
                            int rowCounter = 0;

                            // add columns names to the list
                            foreach (IEntityProperty entityProperty in columnProperties)
                            {
                                columnNames.Add(entityProperty.Name);

                                rg.Offset(rowCounter, columnCounter).Value = entityProperty.DisplayName;
                                columnCounter += 1;
                            }

                            // add values on the row following the headers
                            rowCounter = 1;

                            // iterate the collection and extract values by column name
                            foreach (IEntityObject entityObj in collection)
                            {
                                for (int i = 0; i <= columnNames.Count - 1; i++)
                                {
                                    rg.Offset(rowCounter, i).Value = LightSwitchHelper.GetValue(entityObj, columnNames[i]);
                                }
                                rowCounter += 1;
                            }

                        }

                        result = wb;
                        xlProxy.ShowWorkbook();
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException comException)
            {
                result = null;
                xlProxy.Quit();

                switch (comException.ErrorCode)
                {
                    case -2147352565: // Bad worksheet name
                        throw new System.ArgumentException("Unknown worksheet", "Worksheet");
                    case -2146827284: // Bad path parameter or invalid range reference
                        if (comException.InnerException == null)
                            throw new System.ArgumentException("Invalid range reference", "Range");
                        else
                            throw new System.ArgumentException("Can't open Excel workbook", comException.InnerException);
                    default:
                        throw comException;
                }
            }

            return result;
        }

        // exports a collection to a workbook starting at the specified location
        public static dynamic Export(IVisualCollection collection, string Workbook, string Worksheet, string Range, List<string> ColumnNames)
        {

            List<ColumnMapping> mappings = new List<ColumnMapping>();
            ColumnMapping map = default(ColumnMapping);
            foreach (string name in ColumnNames)
            {
                map = new ColumnMapping("", name);
                map.TableField.DisplayName = name;
                mappings.Add(map);
            }

            return Export(collection, Workbook, Worksheet, Range, mappings);
        }

        // exports a collection to a workbook starting at the specified location
        public static dynamic Export(IVisualCollection collection, string Workbook, string Worksheet, string Range, List<ColumnMapping> ColumnNames)
        {
            ExcelHelper xlProxy = new ExcelHelper();
            dynamic wb = null;
            dynamic ws = null;
            dynamic rg = null;
            dynamic result = null;

            try
            {
                wb = xlProxy.Excel.Workbooks.Open(Workbook);
                if ((wb != null))
                {
                    ws = wb.Worksheets(Worksheet);
                    if ((ws != null))
                    {
                        rg = ws.Range(Range);

                        if (collection.Count > 0)
                        {
                            // get column properties
                            int columnCounter = 0;
                            int rowCounter = 0;

                            // add columns names to the list
                            foreach (var map in ColumnNames)
                            {
                                if (map.TableField.DisplayName.Length > 0)
                                    rg.Offset(rowCounter, columnCounter).Value = map.TableField.DisplayName;
                                else
                                    rg.Offset(rowCounter, columnCounter).Value = map.TableField.Name;
                                columnCounter += 1;
                            }

                            // add values on the row following the headers
                            rowCounter = 1;

                            // iterate the collection and extract values by column name
                            foreach (IEntityObject entityObj in collection)
                            {
                                for (int i = 0; i <= ColumnNames.Count - 1; i++)
                                {
                                    try
                                    {
                                        rg.Offset(rowCounter, i).Value = LightSwitchHelper.GetValue(entityObj, ColumnNames[i].TableField.Name);
                                    }
                                    catch (Exception ex) { }
                                }
                                rowCounter += 1;
                            }

                        }

                        result = wb;
                        xlProxy.ShowWorkbook();
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException comException)
            {
                result = null;
                xlProxy.Quit();
                switch (comException.ErrorCode)
                {
                    case -2147352565:
                        // Bad worksheet name
                        throw new System.ArgumentException("Unknown worksheet", "Worksheet");
                    case -2146827284:
                        // Bad path parameter or invalid range reference
                        if (comException.InnerException == null)
                            throw new System.ArgumentException("Invalid range reference", "Range");
                        else
                            throw new System.ArgumentException("Can't open Excel workbook", comException.InnerException);
                    default:
                        throw comException;
                }
            }

            return result;
        }

        // exports collection to a new workbook starting at cell A1 on the first worksheet
        public static dynamic Export(IVisualCollection collection)
        {
            dynamic result = null;
            try
            {
                ExcelHelper excel = new ExcelHelper();
                excel.CreateWorkbook();

                List<string> columnNames = new List<string>();

                if (collection.Count > 0)
                {
                    // get column properties
                    IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();

                    int columnCounter = 1;
                    int rowCounter = 1;

                    // add columns names to the list
                    foreach (IEntityProperty entityProperty in columnProperties)
                    {
                        columnNames.Add(entityProperty.Name);

                        // add column headers to Excel workbook
                        excel.Cells(rowCounter, columnCounter, entityProperty.DisplayName);
                        columnCounter += 1;
                    }

                    // add values on the row following the headers
                    rowCounter = 2;

                    // iterate the collection and extract values by column name
                    foreach (IEntityObject entityObj in collection)
                    {
                        for (int i = 0; i <= columnNames.Count - 1; i++)
                        {
                            excel.Cells(rowCounter, i + 1, LightSwitchHelper.GetValue(entityObj, columnNames[i]));
                        }
                        rowCounter += 1;
                    }
                }

                excel.ShowWorkbook();
                result = excel.Workbook;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        #endregion

        #region "ExportEntityCollection Overloads"
        // exports an IEntityCollection to a workbook starting at the specified location
        public static dynamic ExportEntityCollection(IEntityCollection collection, string Workbook, string Worksheet, string Range)
        {
            ExcelHelper xlProxy = new ExcelHelper();
            dynamic wb = null;
            dynamic ws = null;
            dynamic rg = null;
            dynamic result = null;

            try
            {
                wb = xlProxy.Excel.Workbooks.Open(Workbook);
                if ((wb != null))
                {
                    ws = wb.Worksheets(Worksheet);

                    if ((ws != null))
                    {
                        rg = ws.Range(Range);

                        List<string> columnNames = new List<string>();

                        // get column properties
                        IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();

                        int columnCounter = 0;
                        int rowCounter = 0;

                        // add columns names to the list
                        foreach (IEntityProperty entityProperty in columnProperties)
                        {
                            columnNames.Add(entityProperty.Name);

                            rg.Offset(rowCounter, columnCounter).Value = entityProperty.DisplayName;
                            columnCounter += 1;
                        }

                        // add values on the row following the headers
                        rowCounter = 1;

                        // iterate the collection and extract values by column name
                        foreach (IEntityObject entityObj in collection)
                        {
                            for (int i = 0; i <= columnNames.Count - 1; i++)
                            {
                                rg.Offset(rowCounter, i).Value = LightSwitchHelper.GetValue(entityObj, columnNames[i]);
                            }
                            rowCounter += 1;
                        }

                        result = wb;
                        xlProxy.ShowWorkbook();
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException comException)
            {
                result = null;
                xlProxy.Quit();

                switch (comException.ErrorCode)
                {
                    case 2147352565:
                        // Bad worksheet name
                        throw new System.ArgumentException("Unknown worksheet", "Worksheet");
                    case -2146827284:
                        // Bad path parameter or invalid range reference
                        if (comException.InnerException == null)
                            throw new System.ArgumentException("Invalid range reference", "Range");
                        else
                            throw new System.ArgumentException("Can't open Excel workbook", comException.InnerException);
                    default:
                        throw comException;
                }
            }

            return result;
        }

        // exports a collection to a workbook starting at the specified location
        public static dynamic ExportEntityCollection(IEntityCollection collection, string Workbook, string Worksheet, string Range, List<string> ColumnNames)
        {

            List<ColumnMapping> mappings = new List<ColumnMapping>();
            ColumnMapping map = default(ColumnMapping);
            foreach (string name in ColumnNames)
            {
                map = new ColumnMapping("", name);
                map.TableField.DisplayName = name;
                mappings.Add(map);
            }

            return ExportEntityCollection(collection, Workbook, Worksheet, Range, mappings);
        }

        // exports a collection to a workbook starting at the specified location
        public static dynamic ExportEntityCollection(IEntityCollection collection, string Workbook, string Worksheet, string Range, List<ColumnMapping> ColumnNames)
        {
            ExcelHelper xlProxy = new ExcelHelper();
            dynamic wb = null;
            dynamic ws = null;
            dynamic rg = null;
            dynamic result = null;

            try
            {
                wb = xlProxy.Excel.Workbooks.Open(Workbook);
                if ((wb != null))
                {
                    ws = wb.Worksheets(Worksheet);
                    if ((ws != null))
                    {
                        rg = ws.Range(Range);

                        // get column properties
                        int columnCounter = 0;
                        int rowCounter = 0;

                        // add columns names to the list
                        foreach (ColumnMapping map in ColumnNames)
                        {
                            if (map.TableField.DisplayName.Length > 0)
                                rg.Offset(rowCounter, columnCounter).Value = map.TableField.DisplayName;
                            else
                                rg.Offset(rowCounter, columnCounter).Value = map.TableField.Name;
                            columnCounter += 1;
                        }

                        // add values on the row following the headers
                        rowCounter = 1;

                        // iterate the collection and extract values by column name
                        foreach (IEntityObject entityObj in collection)
                        {
                            for (int i = 0; i <= ColumnNames.Count - 1; i++)
                            {
                                try
                                {
                                    rg.Offset(rowCounter, i).Value = LightSwitchHelper.GetValue(entityObj, ColumnNames[i].TableField.Name);
                                }
                                catch (Exception ex) { }
                            }
                            rowCounter += 1;
                        }

                        result = wb;
                        xlProxy.ShowWorkbook();

                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException comException)
            {
                result = null;
                xlProxy.Quit();
                switch (comException.ErrorCode)
                {
                    case -2147352565:
                        // Bad worksheet name
                        throw new System.ArgumentException("Unknown worksheet", "Worksheet");
                    case -2146827284:
                        // Bad path parameter or invalid range reference
                        if (comException.InnerException == null)
                            throw new System.ArgumentException("Invalid range reference", "Range");
                        else
                            throw new System.ArgumentException("Can't open Excel workbook", comException.InnerException);
                    default:
                        throw comException;
                }
            }

            return result;
        }

        // exports collection to a new workbook starting at cell A1 on the first worksheet
        public static dynamic ExportEntityCollection(IEntityCollection collection)
        {
            dynamic result = null;

            try
            {
                ExcelHelper excel = new ExcelHelper();
                excel.CreateWorkbook();

                List<string> columnNames = new List<string>();

                // get column properties
                IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();

                int columnCounter = 1;
                int rowCounter = 1;

                // add columns names to the list
                foreach (IEntityProperty entityProperty in columnProperties)
                {
                    columnNames.Add(entityProperty.Name);

                    // add column headers to Excel workbook
                    excel.Cells(rowCounter, columnCounter, entityProperty.DisplayName);
                    columnCounter += 1;
                }

                // add values on the row following the headers
                rowCounter = 2;

                // iterate the collection and extract values by column name
                foreach (IEntityObject entityObj in collection)
                {
                    for (int i = 0; i <= columnNames.Count - 1; i++)
                    {
                        excel.Cells(rowCounter, i + 1, LightSwitchHelper.GetValue(entityObj, columnNames[i]));
                    }
                    rowCounter += 1;
                }

                excel.ShowWorkbook();
                result = excel.Workbook;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        #endregion

        #region "Export IEnumerable overloads"
        // exports collection to a new workbook starting at cell A1 on the first worksheet
        public static dynamic Export<T>(IEnumerable<T> collection, IEnumerable<IEntityProperty> columnProperties)
        {
            dynamic result = null;
            try
            {
                ExcelHelper excel = new ExcelHelper();
                excel.CreateWorkbook();

                List<string> columnNames = new List<string>();

                if (collection.Count() > 0)
                {
                    // get column properties
                    //IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();

                    int columnCounter = 1;
                    int rowCounter = 1;

                    // add columns names to the list
                    foreach (IEntityProperty entityProperty in columnProperties)
                    {
                        columnNames.Add(entityProperty.Name);

                        // add column headers to Excel workbook
                        excel.Cells(rowCounter, columnCounter, entityProperty.DisplayName);
                        columnCounter += 1;
                    }

                    // add values on the row following the headers
                    rowCounter = 2;

                    // iterate the collection and extract values by column name IEntityObject entityObj in collection)
                    var entities = collection.OfType<IEntityObject>();
                    foreach(var entityObj in entities)
                    {
                        for (int i = 0; i <= columnNames.Count - 1; i++)
                        {
                            excel.Cells(rowCounter, i + 1, LightSwitchHelper.GetValue(entityObj, columnNames[i]));
                        }
                        rowCounter += 1;
                    }

                    //// iterate the collection and extract values by column name IEntityObject entityObj in collection)
                    //var entities = collection.OfType<IEntityObject>();
                    //Parallel.ForEach(entities, entityObj =>
                    //{
                    //    //for (int i = 0; i <= columnNames.Count - 1; i++)
                    //    Parallel.For(0, columnNames.Count, i =>
                    //    {
                    //        excel.Cells(rowCounter, i + 1, LightSwitchHelper.GetValue(entityObj, columnNames[i]));
                    //    });
                    //    rowCounter += 1;
                    //});
                }

                excel.ShowWorkbook();
                result = excel.Workbook;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        #endregion

        #region "Import Overloads"
        // Imports a range starting at the location specified by workbook, worksheet, and range.
        // Workbook should be the full path to the workbook.
        public static void Import(IVisualCollection collection, string Workbook, string Worksheet, string Range)
        {
            _collection = collection;

            Dispatchers.Main.BeginInvoke(() =>
            {
                ExcelHelper xlProxy = new ExcelHelper();
                dynamic wb = null;
                dynamic ws = null;
                dynamic rg = null;

                wb = xlProxy.Excel.Workbooks.Open(Workbook);
                if ((wb != null))
                {
                    ws = wb.Worksheets(Worksheet);
                    if ((ws != null))
                    {
                        rg = ws.Range(Range);
                        _excelDocRange = ConvertToArray(rg);

                        List<FieldDefinition> tablePropertyChoices = GetTablePropertyChoices();

                        _columnMappings = new List<ColumnMapping>();
                        Int32 numColumns = _excelDocRange.GetLength(1);
                        for (var i = 0; i <= numColumns - 1; i++)
                        {
                            _columnMappings.Add(new ColumnMapping(_excelDocRange[0, i], tablePropertyChoices));
                        }

                        ColumnMapper columnMapperContent = new ColumnMapper();
                        columnMapperContent.OfficeColumn.Text = "Excel Column";
                        ScreenChildWindow columnMapperWindow = new ScreenChildWindow();
                        columnMapperContent.DataContext = _columnMappings;
                        columnMapperWindow.Closed += OnMappingDialogClosed;

                        //set parent to current screen
                        IServiceProxy sdkProxy = VsExportProviderService.GetExportedValue<IServiceProxy>();
                        columnMapperWindow.Owner = (Control)sdkProxy.ScreenViewService.GetScreenView(_collection.Screen).RootUI;
                        columnMapperWindow.Show(_collection.Screen, columnMapperContent);

                    }
                }

                rg = null;
                ws = null;
                wb.Close(false);
                wb = null;
                xlProxy.Dispose();
            });
        }

        // Imports a range starting at the location specified by workbook, worksheet, and range.
        // Workbook should be the full path to the workbook.
        public static void Import(IVisualCollection collection, string Workbook, string Worksheet, string Range, List<ColumnMapping> ColumnMapping)
        {
            _collection = collection;
            _columnMappings = ColumnMapping;

            // Ensure that we have the correct and complete field definition.
            foreach (OfficeSharp.ColumnMapping map in _columnMappings)
            {
                map.TableField = collection.GetFieldDefinition(map.TableField.Name);
            }

            Dispatchers.Main.BeginInvoke(() =>
            {
                ExcelHelper xlProxy = new ExcelHelper();
                dynamic wb = null;
                dynamic ws = null;
                dynamic rg = null;

                wb = xlProxy.Excel.Workbooks.Open(Workbook);
                if ((wb != null))
                {
                    ws = wb.Worksheets(Worksheet);
                    if ((ws != null))
                    {
                        rg = ws.Range(Range);
                        _excelDocRange = ConvertToArray(rg);
                        ValidateData();
                    }
                }

                rg = null;
                ws = null;
                wb.Close(false);
                wb = null;
                xlProxy.Dispose();
            });
        }

        // Imports a range from workbook chosen by end-user. Assumes table starts at cell A1 on the first worksheet.
        public static void Import(IVisualCollection collection)
        {
            _collection = collection;

            Dispatchers.Main.BeginInvoke(() =>
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Multiselect = false;
                dialog.Filter = "Excel Documents(*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv|All files (*.*)|*.*";
                if (dialog.ShowDialog() == true)
                {
                    FileInfo f = dialog.File;
                    try
                    {
                        ExcelHelper excel = new ExcelHelper();
                        excel.OpenWorkbook(f.FullName);
                        _excelDocRange = excel.UsedRange(1, 1);
                        excel.Dispose();

                        List<FieldDefinition> tablePropertyChoices = GetTablePropertyChoices();

                        _columnMappings = new List<ColumnMapping>();
                        Int32 numColumns = _excelDocRange.GetLength(1);
                        for (var i = 0; i <= numColumns - 1; i++)
                        {
                            _columnMappings.Add(new ColumnMapping(_excelDocRange[0, i], tablePropertyChoices));
                        }

                        ColumnMapper columnMapperContent = new ColumnMapper();
                        columnMapperContent.OfficeColumn.Text = "Excel Column";
                        ScreenChildWindow columnMapperWindow = new ScreenChildWindow();
                        columnMapperContent.DataContext = _columnMappings;
                        columnMapperWindow.Closed += OnMappingDialogClosed;

                        //set parent to current screen
                        IServiceProxy sdkProxy = VsExportProviderService.GetExportedValue<IServiceProxy>();
                        columnMapperWindow.Owner = (Control)sdkProxy.ScreenViewService.GetScreenView(_collection.Screen).RootUI;
                        columnMapperWindow.Show(_collection.Screen, columnMapperContent);

                    }
                    catch (SecurityException ex)
                    {
                        collection.Screen.Details.Dispatcher.BeginInvoke(() => { _collection.Screen.ShowMessageBox("Error: Silverlight Security error. Could not load Excel document. Make sure the document is in your 'Documents' directory."); });
                    }
                    catch (COMException comEx)
                    {
                        _collection.Screen.Details.Dispatcher.BeginInvoke(() => { _collection.Screen.ShowMessageBox("Error: Could not open this file.  It may not be a valid Excel document."); });
                    }
                }
            });
        }
        #endregion

        // Returns an Excel.Application dynamic
        public static dynamic GetExcel()
        {
            ExcelHelper xlProxy = new ExcelHelper();
            dynamic xl = null;

            // first try and get a reference to a running instance
            if (xlProxy.GetExcel())
            {
                xl = xlProxy.Excel;
            }
            else
            {
                // next try and create a new instance
                if (xlProxy.CreateExcel())
                {
                    xl = xlProxy.Excel;
                }
                else
                {
                    // can't get Excel - return Nothing by default
                }
            }
            return xl;
        }

        // Returns an Excel.Workbook dynamic
        public static dynamic GetWorkbook(dynamic Excel, string WorkbookPath)
        {
            dynamic wb = null;

            // Validate Word argument is actually Word.Application
            if (!IsExcelApplicationObject(Excel))
            {
                throw new System.ArgumentException("'Excel' is not the expected type of dynamic. Expected dynamic should be an Excel.Application dynamic.", "Excel");
            }

            if (!File.Exists(WorkbookPath))
            {
                throw new System.ArgumentException("File '" + WorkbookPath + "' does not exist.", "WorkbookPath");
            }

            try
            {
                wb = Excel.Workbooks.Open(WorkbookPath);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return wb;
        }

        private static bool IsExcelApplicationObject(dynamic app)
        {
            bool isValid = false;
            string s = null;
            try
            {
                s = app.Name;
                if (s == "Microsoft Excel")
                {
                    isValid = true;
                }
            }
            catch (Exception ex)
            {
            }
            return isValid;
        }

        private static string[,] ConvertToArray(dynamic rg)
        {
            Int32 rowCount = rg.Rows.Count;
            Int32 columnCount = rg.Columns.Count;

            string[,] valueArray = new string[rowCount, columnCount];
            for (var i = 1; i <= rowCount; i++)
            {
                for (var j = 1; j <= columnCount; j++)
                {
                    valueArray[i - 1, j - 1] = rg.Cells(i, j).Value;
                }
            }
            return valueArray;
        }

        private static void AddItemsToCollection()
        {
            //This should always be called on the logical thread
            Debug.Assert(_collection.Screen.Details.Dispatcher.CheckAccess(), "Expected to run on the logical thread");

            //Add Items to Collection
            dynamic numRows = _excelDocRange.GetLength(0) - 1;
            for (var i = 1; i <= numRows; i++)
            {
                var newRow = _collection.AddNew() as IEntityObject;
                Int32 currentRow = i;
                Int32 nOfficeColumnIndex = default(Int32);
                foreach (ColumnMapping mapping in _columnMappings)
                {
                    if (mapping.TableField != null && mapping.TableField.Name != "<Ignore>")
                    {
                        nOfficeColumnIndex = GetColumnIndex(mapping.OfficeColumn);
                        if (nOfficeColumnIndex >= 0)
                        {
                            string value = _excelDocRange[currentRow, nOfficeColumnIndex];
                            if (string.IsNullOrEmpty(value) && mapping.TableField.IsNullable)
                            {
                                newRow.Details.Properties[mapping.TableField.Name].Value = null;
                            }
                            else if (mapping.TableField.EntityType != null)
                            {
                                //Get cached results
                                if (navProperties.ContainsKey(String.Format("{0}_{1}", mapping.TableField.Name, value)))
                                {
                                    newRow.Details.Properties[mapping.TableField.Name].Value = navProperties[String.Format("{0}_{1}", mapping.TableField.Name, value)];
                                }
                            }
                            else
                            {
                                var propValue = newRow.Details.Properties[mapping.TableField.Name].Value;
                                LightSwitchHelper.TryConvertValue(mapping.TableField.TypeName, value, ref propValue);
                            }
                        }
                    }
                }
            }
        }

        private static void ValidateData()
        {
            List<string> errorList = new List<string>();
            dynamic numRows = _excelDocRange.GetLength(0) - 1;

            //Dispatch to the logical thread
            _collection.Screen.Details.Dispatcher.BeginInvoke(() =>
            {
                try
                {
                    //Work through each row of data

                    for (var i = 1; i <= numRows; i++)
                    {
                        Int32 currentRow = i;
                        Int32 count = 0;
                        Int32 nOfficeColumnIndex = default(Int32);

                        // Work through each mapped column

                        foreach (ColumnMapping mapping in _columnMappings)
                        {
                            // Make sure the current column should be processed

                            if (mapping.TableField != null && mapping.TableField.Name != "<Ignore>")
                            {
                                nOfficeColumnIndex = GetColumnIndex(mapping.OfficeColumn);


                                if (nOfficeColumnIndex >= 0)
                                {
                                    // Read in the value for the current row and column
                                    string value = _excelDocRange[currentRow, nOfficeColumnIndex];

                                    // Process values without null or empty string values

                                    if (!(string.IsNullOrEmpty(value)))
                                    {
                                        // Try and validate based on the Entity type
                                        bool isValid = false;
                                        if (mapping.TableField.EntityType != null)
                                        {
                                            isValid = LightSwitchHelper.ValidData(value, currentRow, mapping, _collection, navProperties, errorList);
                                        }
                                        else
                                        {
                                            isValid = LightSwitchHelper.ValidData(value, currentRow, mapping, errorList);
                                        }


                                    }
                                    else if (string.IsNullOrEmpty(value) && mapping.TableField.IsNullable == false && mapping.TableField.EntityType == null && mapping.TableField.TypeName != "String")
                                    {
                                        errorList.Add("Column: " + mapping.OfficeColumn + " Row:" + currentRow.ToString() + " A value must be specified for " + mapping.TableField.DisplayName + ". A default value will be used.");

                                    }

                                }
                                else
                                {
                                    // Couldn't find the column in the Excel range
                                    errorList.Add("Column: " + mapping.OfficeColumn + " Row:" + currentRow.ToString() + " Could not locate column in Office document.");
                                }
                            }
                            count = count + 1;
                        }
                    }
                    if (errorList.Count > 0)
                    {
                        DisplayErrors(errorList);
                    }
                    else
                    {
                        //Add Items to Collection
                        AddItemsToCollection();
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            });
        }

        private static Int32 GetColumnIndex(string sOfficeColumnName)
        {
            Int32 nResult = -1;
            Int32 nColumn = default(Int32);

            for (nColumn = 0; nColumn <= _excelDocRange.GetLength(1) - 1; nColumn++)
            {
                if (_excelDocRange[0, nColumn] == sOfficeColumnName)
                {
                    nResult = nColumn;
                    break; // TODO: might not be correct. Was : Exit For
                }
            }
            return nResult;
        }

        private static void DisplayErrors(List<string> errorList)
        {
            Dispatchers.Main.BeginInvoke(() =>
            {
                //Display some sort of dialog indicating that errors occurred
                IServiceProxy sdkProxy = VsExportProviderService.GetExportedValue<IServiceProxy>();
                ErrorList errorDialog = new ErrorList();
                ScreenChildWindow errorWindow = new ScreenChildWindow();
                errorDialog.DataContext = errorList;
                errorWindow.DataContext = errorList;
                errorWindow.Owner = (Control)sdkProxy.ScreenViewService.GetScreenView(_collection.Screen).RootUI;

                errorWindow.Closed += OnErroDialogClosed;
                errorWindow.Show(_collection.Screen, errorDialog);
            });
        }

        #region "Dialog Closing Methods"
        private static void OnErroDialogClosed(dynamic sender, EventArgs e)
        {
            ScreenChildWindow errorWindow = sender;
            bool? result = errorWindow.DialogResult;
            if (result.HasValue && result.Value == true)
            {
                _collection.Screen.Details.Dispatcher.BeginInvoke(() => { AddItemsToCollection(); });
            }
        }

        private static void OnMappingDialogClosed(dynamic sender, EventArgs e)
        {
            ScreenChildWindow mappingWindow = sender;
            bool? result = mappingWindow.DialogResult;
            if (result.HasValue && result.Value == true)
            {
                ValidateData();
            }
        }
        #endregion

        private static List<FieldDefinition> GetTablePropertyChoices()
        {
            var entityType = _collection.Details.GetModel().ElementType as IEntityType;
            List<FieldDefinition> tablePropertyChoices = new List<FieldDefinition>();

            foreach (IEntityPropertyDefinition p in entityType.Properties)
            {
                //Ignore hidden fields and computed field
                if (p.Attributes.Where(a => a.Class.Name == "Computed").FirstOrDefault() == null)
                {
                    if (!(p.PropertyType is ISequenceType))
                    {
                        //ignore collections and entities
                        FieldDefinition fd = new FieldDefinition();
                        fd.Name = p.Name;
                        fd.DisplayName = p.Name;

                        bool isNullable = false;
                        fd.TypeName = p.PropertyType.GetPropertyType(ref isNullable);
                        fd.IsNullable = isNullable;
                        if (fd.TypeName == "Entity")
                        {
                            fd.EntityType = (IEntityType)p.PropertyType;
                        }

                        tablePropertyChoices.Add(fd);
                    }
                }
            }
            tablePropertyChoices.Add(new FieldDefinition("<Ignore>", "<Ignore>", "", false));
            return tablePropertyChoices;
        }
    }
}
