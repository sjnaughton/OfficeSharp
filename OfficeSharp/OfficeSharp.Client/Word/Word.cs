using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Automation;
using System.Security;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
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

namespace OfficeSharp.Word
{
    public static class Word
    {
        public static dynamic GenerateDocument(string Template, IEntityObject Item, List<ColumnMapping> ColumnMappings)
        {
            dynamic doc = null;
            WordHelper wordProxy = new WordHelper();

            wordProxy.CreateWord();
            wordProxy.OpenDocument(Template);
            PopulateContentControls(ColumnMappings, Item, wordProxy);
            doc = wordProxy.Document;
            wordProxy.ShowDocument();

            return doc;
        }

        #region "Export Overloads"
        // Exports an IVisualCollection to a table in either the active (UseActiveDocument = True) or a new document (UseActiveDocument = False)
        public static dynamic Export(IVisualCollection collection, bool UseActiveDocument)
        {
            dynamic doc = null;
            WordHelper wordProxy = new WordHelper();
            bool bUseActiveDocument = false;
            dynamic rg = null;

            // if Word is active then use it
            if (wordProxy.GetWord())
            {
                // obtain a reference to the selection range
                if (UseActiveDocument)
                {
                    rg = wordProxy.Word.Selection.Range;
                    bUseActiveDocument = true;
                }
                else
                {
                    wordProxy.CreateDocument();
                }
            }
            else
            {
                wordProxy.CreateWord();
                wordProxy.CreateDocument();
            }

            List<string> columnNames = new List<string>();

            if (collection.Count > 0)
            {
                // get column properties
                IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();
                int columnCounter = 1;
                int rowCounter = 1;

                // add table
                dynamic oTable = null;
                if (bUseActiveDocument)
                {
                    oTable = wordProxy.AddTable(collection.Count + 1, columnProperties.Count(), rg);
                }
                else
                {
                    oTable = wordProxy.AddTable(collection.Count + 1, columnProperties.Count());
                }


                // add columns names to the list
                foreach (IEntityProperty entityProperty in columnProperties)
                {
                    columnNames.Add(entityProperty.Name);

                    // add column headers to table
                    wordProxy.SetTableCell(oTable, 1, columnCounter, entityProperty.DisplayName);
                    columnCounter += 1;
                }

                // add values on the row following the headers
                rowCounter = 2;

                // iterate the collection and extract values by column name
                foreach (IEntityObject entityObj in collection)
                {
                    for (int i = 0; i <= columnNames.Count - 1; i++)
                    {
                        wordProxy.SetTableCell(oTable, rowCounter, i + 1, LightSwitchHelper.GetValue(entityObj, columnNames[i]));
                    }
                    rowCounter += 1;
                }

            }

            doc = wordProxy.Document;
            wordProxy.ShowDocument();

            return doc;
        }

        // Exports a collection to a table in a Word document located at DocumentPath. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic Export(string DocumentPath, string BookmarkName, int StartRow, bool BuildColumnHeadings, IVisualCollection collection)
        {
            dynamic doc = null;
            dynamic word = null;
            dynamic result = null;

            try
            {
                word = GetWord();
                if ((word != null))
                {
                    doc = GetDocument(word, DocumentPath);
                    if ((doc != null))
                    {
                        result = Export(doc, BookmarkName, StartRow, BuildColumnHeadings, collection);
                        word.Visible = true;
                    }
                    else
                    {
                        throw new System.Exception("Could not open the document '" + DocumentPath + "'.");
                    }
                }
                else
                {
                    throw new System.Exception("Could not obtain a reference to Microsoft Word.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        // Exports a collection to a table in given Document. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic Export(dynamic Document, string BookmarkName, int StartRow, bool BuildColumnHeadings, IVisualCollection collection)
        {
            List<string> columnNames = new List<string>();


            if (collection.Count > 0)
            {
                // get column properties
                IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();
                int columnCounter = 1;
                int rowCounter = StartRow;

                // validate that the Documnet argument is expected type of Word.Document
                if (!IsWordDocumentObject(Document))
                {
                    throw new System.ArgumentException("'Document' is not the expected type of dynamic. Expected dynamic should be a Word.Document dynamic.", "Document");
                }

                // validate the BookmarkName argument
                if (!IsValidBookmark(Document, BookmarkName))
                {
                    throw new System.ArgumentException("'BookmarkName' was not found in 'Document'", "BookmarkName");
                }

                // validate that the bookmark is part of a table
                if (Document.Bookmarks(BookmarkName).Range.Tables.Count == 0)
                {
                    throw new System.ArgumentException("No table was found at the bookmark", "BookmarkName");
                }

                // add table
                dynamic oTable = null;
                oTable = Document.Bookmarks(BookmarkName).Range.Tables(1);

                // validate the StartRow argument
                if (StartRow > oTable.Rows.Count)
                {
                    throw new System.ArgumentException("'StartRow' is greater then the number of rows in the table", "StartRow");
                }

                // add columns names to the list
                foreach (IEntityProperty entityProperty in columnProperties)
                {
                    columnNames.Add(entityProperty.Name);
                    if (columnCounter > oTable.Columns.Count)
                    {
                        oTable.Columns.Add();
                    }
                    if (BuildColumnHeadings)
                    {
                        oTable.Cell(rowCounter, columnCounter).Range.Text = entityProperty.DisplayName;
                    }
                    columnCounter += 1;
                }

                // add values on the row following the headers
                if (BuildColumnHeadings)
                {
                    rowCounter += 1;
                }

                // iterate the collection and extract values by column name
                foreach (IEntityObject entityObj in collection)
                {
                    for (int i = 0; i <= columnNames.Count - 1; i++)
                    {
                        if (rowCounter > oTable.Rows.Count)
                        {
                            oTable.Rows.Add();
                        }
                        oTable.Cell(rowCounter, i + 1).Range.Text = LightSwitchHelper.GetValue(entityObj, columnNames[i]);
                    }
                    rowCounter += 1;
                }

            }
            else
            {
                // No items in the collection
            }

            return Document;
        }

        // Exports a collection to a table in given Document. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic Export(string DocumentPath, string BookmarkName, int StartRow, bool BuildColumnHeadings, IVisualCollection collection, List<string> ColumnNames)
        {
            dynamic functionReturnValue = null;
            List<ColumnMapping> mappings = new List<ColumnMapping>();
            ColumnMapping map = default(ColumnMapping);
            dynamic doc = null;
            WordHelper wordProxy = new WordHelper();
            FieldDefinition fd = default(FieldDefinition);

            // if Word is active then use it
            if (!wordProxy.GetWord())
            {
                if (!wordProxy.CreateWord())
                {
                    throw new System.Exception("Could not start Microsoft Word.");
                }
            }

            wordProxy.OpenDocument(DocumentPath);
            doc = wordProxy.Document;

            foreach (string name in ColumnNames)
            {
                fd = collection.GetFieldDefinition(name);
                map = new ColumnMapping("", name);
                if (fd == null)
                {
                    map.TableField.DisplayName = name;
                }
                else
                {
                    map.TableField = fd;
                }
                mappings.Add(map);
            }

            functionReturnValue = Export(doc, BookmarkName, StartRow, BuildColumnHeadings, collection, mappings);
            wordProxy.ShowDocument();
            return functionReturnValue;
        }

        // Exports a collection to a table in given Document. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic Export(string DocumentPath, string BookmarkName, int StartRow, bool BuildColumnHeadings, IVisualCollection collection, List<ColumnMapping> ColumnNames)
        {
            dynamic functionReturnValue = null;
            dynamic doc = null;
            WordHelper wordProxy = new WordHelper();

            // if Word is active then use it
            if (!wordProxy.GetWord())
            {
                if (!wordProxy.CreateWord())
                {
                    throw new System.Exception("Could not start Microsoft Word.");
                }
            }

            wordProxy.OpenDocument(DocumentPath);
            doc = wordProxy.Document;

            functionReturnValue = Export(doc, BookmarkName, StartRow, BuildColumnHeadings, collection, ColumnNames);
            wordProxy.ShowDocument();
            return functionReturnValue;
        }

        // Exports a collection to a table in given Document. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic Export(dynamic Document, string BookmarkName, int StartRow, bool BuildColumnHeadings, IVisualCollection collection, List<string> ColumnNames)
        {
            List<ColumnMapping> mappings = new List<ColumnMapping>();
            ColumnMapping map = default(ColumnMapping);
            FieldDefinition fd = default(FieldDefinition);

            foreach (string name in ColumnNames)
            {
                fd = collection.GetFieldDefinition(name);
                map = new ColumnMapping("", name);
                if (fd == null)
                {
                    map.TableField.DisplayName = name;
                }
                else
                {
                    map.TableField = fd;
                }
                mappings.Add(map);
            }

            return Export(Document, BookmarkName, StartRow, BuildColumnHeadings, collection, mappings);
        }

        // Exports an IVisualCollection to a table in given Document. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic Export(dynamic Document, string BookmarkName, int StartRow, bool BuildColumnHeadings, IVisualCollection collection, List<ColumnMapping> ColumnNames)
        {


            if (collection.Count > 0)
            {
                // get column properties
                int columnCounter = 1;
                int rowCounter = StartRow;

                // validate that the Documnet argument is expected type of Word.Document
                if (!IsWordDocumentObject(Document))
                {
                    throw new System.ArgumentException("'Document' is not the expected type of dynamic. Expected dynamic should be a Word.Document dynamic.", "Document");
                }

                // validate the BookmarkName argument
                if (!IsValidBookmark(Document, BookmarkName))
                {
                    throw new System.ArgumentException("'BookmarkName' was not found in 'Document'", "BookmarkName");
                }

                // validate that the bookmark is part of a table
                if (Document.Bookmarks(BookmarkName).Range.Tables.Count == 0)
                {
                    throw new System.ArgumentException("No table was found at the bookmark", "BookmarkName");
                }

                // add table
                dynamic oTable = null;
                oTable = Document.Bookmarks(BookmarkName).Range.Tables(1);

                // validate the StartRow argument
                if (StartRow > oTable.Rows.Count)
                {
                    throw new System.ArgumentException("'StartRow' is greater then the number of rows in the table", "StartRow");
                }

                // add columns names to the list
                foreach (ColumnMapping map in ColumnNames)
                {
                    if (columnCounter > oTable.Columns.Count)
                    {
                        oTable.Columns.Add();
                    }
                    if (BuildColumnHeadings)
                    {
                        if (map.TableField.DisplayName.Length > 0)
                        {
                            oTable.Cell(rowCounter, columnCounter).Range.Text = map.TableField.DisplayName;
                        }
                        else
                        {
                            oTable.Cell(rowCounter, columnCounter).Range.Text = map.TableField.Name;
                        }
                    }
                    columnCounter += 1;
                }

                // add values on the row following the headers
                if (BuildColumnHeadings)
                {
                    rowCounter += 1;
                }

                // iterate the collection and extract values by column name
                foreach (IEntityObject entityObj in collection)
                {
                    for (int i = 0; i <= ColumnNames.Count - 1; i++)
                    {
                        if (rowCounter > oTable.Rows.Count)
                        {
                            oTable.Rows.Add();
                        }
                        try
                        {
                            oTable.Cell(rowCounter, i + 1).Range.Text = LightSwitchHelper.GetValue(entityObj, ColumnNames[i].TableField.Name);
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                    rowCounter += 1;
                }

            }
            else
            {
                // No items in the collection
            }

            return Document;
        }

        #endregion

        #region "ExportEntityCollection Overloads"

        // Exports an IEntityCollection to a table in either the active (UseActiveDocument = True) or a new document (UseActiveDocument = False)
        public static dynamic ExportEntityCollection(IEntityCollection collection, bool UseActiveDocument)
        {
            dynamic doc = null;
            WordHelper wordProxy = new WordHelper();
            bool bUseActiveDocument = false;
            dynamic rg = null;

            // if Word is active then use it
            if (wordProxy.GetWord())
            {
                // obtain a reference to the selection range
                if (UseActiveDocument)
                {
                    rg = wordProxy.Word.Selection.Range;
                    bUseActiveDocument = true;
                }
                else
                {
                    wordProxy.CreateDocument();
                }
            }
            else
            {
                wordProxy.CreateWord();
                wordProxy.CreateDocument();
            }

            List<string> columnNames = new List<string>();

            // get column properties
            IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();
            int columnCounter = 1;
            int rowCounter = 1;

            // add table
            dynamic oTable = null;
            if (bUseActiveDocument)
            {
                oTable = wordProxy.AddTable(1, columnProperties.Count(), rg);
            }
            else
            {
                oTable = wordProxy.AddTable(1, columnProperties.Count());
            }


            // add columns names to the list
            foreach (IEntityProperty entityProperty in columnProperties)
            {
                columnNames.Add(entityProperty.Name);

                // add column headers to table
                wordProxy.SetTableCell(oTable, 1, columnCounter, entityProperty.DisplayName);
                columnCounter += 1;
            }

            // add values on the row following the headers
            rowCounter = 2;

            // iterate the collection and extract values by column name
            foreach (IEntityObject entityObj in collection)
            {
                oTable.Rows.Add();
                for (int i = 0; i <= columnNames.Count - 1; i++)
                {
                    wordProxy.SetTableCell(oTable, rowCounter, i + 1, LightSwitchHelper.GetValue(entityObj, columnNames[i]));
                }
                rowCounter += 1;
            }


            doc = wordProxy.Document;
            wordProxy.ShowDocument();

            return doc;
        }

        // Exports an IEntityCollection to a table in a Word document located at DocumentPath. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic ExportEntityCollection(string DocumentPath, string BookmarkName, int StartRow, bool BuildColumnHeadings, IEntityCollection collection)
        {
            dynamic doc = null;
            dynamic word = null;
            dynamic result = null;

            try
            {
                word = GetWord();
                if ((word != null))
                {
                    doc = GetDocument(word, DocumentPath);
                    if ((doc != null))
                    {
                        result = ExportEntityCollection(doc, BookmarkName, StartRow, BuildColumnHeadings, collection);
                        word.Visible = true;
                    }
                    else
                    {
                        throw new System.Exception("Could not open the document '" + DocumentPath + "'.");
                    }
                }
                else
                {
                    throw new System.Exception("Could not obtain a reference to Microsoft Word.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        // Exports an IEntityCollection to a table in given Document. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic ExportEntityCollection(dynamic Document, string BookmarkName, int StartRow, bool BuildColumnHeadings, IEntityCollection collection)
        {
            List<string> columnNames = new List<string>();

            // get column properties
            IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();
            int columnCounter = 1;
            int rowCounter = StartRow;

            // validate that the Document argument is expected type of Word.Document
            if (!IsWordDocumentObject(Document))
            {
                throw new System.ArgumentException("'Document' is not the expected type of dynamic. Expected dynamic should be a Word.Document dynamic.", "Document");
            }

            // validate the BookmarkName argument
            if (!IsValidBookmark(Document, BookmarkName))
            {
                throw new System.ArgumentException("'BookmarkName' was not found in 'Document'", "BookmarkName");
            }

            // validate that the bookmark is part of a table
            if (Document.Bookmarks(BookmarkName).Range.Tables.Count == 0)
            {
                throw new System.ArgumentException("No table was found at the bookmark", "BookmarkName");
            }

            // add table
            dynamic oTable = null;
            oTable = Document.Bookmarks(BookmarkName).Range.Tables(1);

            // validate the StartRow argument
            if (StartRow > oTable.Rows.Count)
            {
                throw new System.ArgumentException("'StartRow' is greater then the number of rows in the table", "StartRow");
            }

            // add columns names to the list
            foreach (IEntityProperty entityProperty in columnProperties)
            {
                columnNames.Add(entityProperty.Name);
                if (columnCounter > oTable.Columns.Count)
                {
                    oTable.Columns.Add();
                }
                if (BuildColumnHeadings)
                {
                    oTable.Cell(rowCounter, columnCounter).Range.Text = entityProperty.DisplayName;
                }
                columnCounter += 1;
            }

            // add values on the row following the headers
            if (BuildColumnHeadings)
            {
                rowCounter += 1;
            }

            // iterate the collection and extract values by column name
            foreach (IEntityObject entityObj in collection)
            {
                for (int i = 0; i <= columnNames.Count - 1; i++)
                {
                    if (rowCounter > oTable.Rows.Count)
                    {
                        oTable.Rows.Add();
                    }
                    oTable.Cell(rowCounter, i + 1).Range.Text = LightSwitchHelper.GetValue(entityObj, columnNames[i]);
                }
                rowCounter += 1;
            }

            return Document;
        }

        // Exports an IEntityCollection to a table in given Document. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic ExportEntityCollection(string DocumentPath, string BookmarkName, int StartRow, bool BuildColumnHeadings, IEntityCollection collection, List<string> ColumnNames)
        {
            dynamic functionReturnValue = null;
            List<ColumnMapping> mappings = new List<ColumnMapping>();
            ColumnMapping map = default(ColumnMapping);

            dynamic doc = null;
            WordHelper wordProxy = new WordHelper();

            // if Word is active then use it
            if (!wordProxy.GetWord())
            {
                if (!wordProxy.CreateWord())
                {
                    throw new System.Exception("Could not start Microsoft Word.");
                }
            }

            wordProxy.OpenDocument(DocumentPath);
            doc = wordProxy.Document;

            foreach (string name in ColumnNames)
            {
                map = new ColumnMapping("", name);
                map.TableField.DisplayName = name;
                mappings.Add(map);
            }

            functionReturnValue = ExportEntityCollection(doc, BookmarkName, StartRow, BuildColumnHeadings, collection, mappings);
            wordProxy.ShowDocument();
            return functionReturnValue;
        }

        // Exports an IEntityCollection to a table in given Document. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic ExportEntityCollection(string DocumentPath, string BookmarkName, int StartRow, bool BuildColumnHeadings, IEntityCollection collection, List<ColumnMapping> ColumnNames)
        {
            dynamic functionReturnValue = null;
            dynamic doc = null;
            WordHelper wordProxy = new WordHelper();

            // if Word is active then use it
            if (!wordProxy.GetWord())
            {
                if (!wordProxy.CreateWord())
                {
                    throw new System.Exception("Could not start Microsoft Word.");
                }
            }

            wordProxy.OpenDocument(DocumentPath);
            doc = wordProxy.Document;

            functionReturnValue = ExportEntityCollection(doc, BookmarkName, StartRow, BuildColumnHeadings, collection, ColumnNames);
            wordProxy.ShowDocument();
            return functionReturnValue;
        }

        // Exports an IEntityCollection to a table in given Document. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic ExportEntityCollection(dynamic Document, string BookmarkName, int StartRow, bool BuildColumnHeadings, IEntityCollection collection, List<string> ColumnNames)
        {
            List<ColumnMapping> mappings = new List<ColumnMapping>();
            ColumnMapping map = default(ColumnMapping);

            foreach (string name in ColumnNames)
            {
                map = new ColumnMapping("", name);
                map.TableField.DisplayName = name;
                mappings.Add(map);
            }

            return ExportEntityCollection(Document, BookmarkName, StartRow, BuildColumnHeadings, collection, mappings);
        }

        // Exports an IEntityCollection to a table in given Document. BookmarkName is the name of the bookmark associated
        // with the table.
        public static dynamic ExportEntityCollection(dynamic Document, string BookmarkName, int StartRow, bool BuildColumnHeadings, IEntityCollection collection, List<ColumnMapping> ColumnNames)
        {

            // get column properties
            int columnCounter = 1;
            int rowCounter = StartRow;

            // validate that the Document argument is expected type of Word.Document
            if (!IsWordDocumentObject(Document))
            {
                throw new System.ArgumentException("'Document' is not the expected type of dynamic. Expected dynamic should be a Word.Document dynamic.", "Document");
            }

            // validate the BookmarkName argument
            if (!IsValidBookmark(Document, BookmarkName))
            {
                throw new System.ArgumentException("'BookmarkName' was not found in 'Document'", "BookmarkName");
            }

            // validate that the bookmark is part of a table
            if (Document.Bookmarks(BookmarkName).Range.Tables.Count == 0)
            {
                throw new System.ArgumentException("No table was found at the bookmark", "BookmarkName");
            }

            // add table
            dynamic oTable = null;
            oTable = Document.Bookmarks(BookmarkName).Range.Tables(1);

            // validate the StartRow argument
            if (StartRow > oTable.Rows.Count)
            {
                throw new System.ArgumentException("'StartRow' is greater then the number of rows in the table", "StartRow");
            }

            // add columns names to the list
            foreach (ColumnMapping map in ColumnNames)
            {
                if (columnCounter > oTable.Columns.Count)
                {
                    oTable.Columns.Add();
                }
                if (BuildColumnHeadings)
                {
                    if (map.TableField.DisplayName.Length > 0)
                    {
                        oTable.Cell(rowCounter, columnCounter).Range.Text = map.TableField.DisplayName;
                    }
                    else
                    {
                        oTable.Cell(rowCounter, columnCounter).Range.Text = map.TableField.Name;
                    }
                }
                columnCounter += 1;
            }

            // add values on the row following the headers
            if (BuildColumnHeadings)
            {
                rowCounter += 1;
            }

            // iterate the collection and extract values by column name
            foreach (IEntityObject entityObj in collection)
            {
                for (int i = 0; i <= ColumnNames.Count - 1; i++)
                {
                    if (rowCounter > oTable.Rows.Count)
                    {
                        oTable.Rows.Add();
                    }
                    try
                    {
                        oTable.Cell(rowCounter, i + 1).Range.Text = LightSwitchHelper.GetValue(entityObj, ColumnNames[i].TableField.Name);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
                rowCounter += 1;
            }

            return Document;
        }

        #endregion

        // Returns a Word.Application dynamic
        public static dynamic GetWord()
        {
            WordHelper wordProxy = new WordHelper();
            dynamic w = null;

            // first try and get a reference to a running instance
            if (wordProxy.GetWord())
                w = wordProxy.Word;
            else
            {
                // next try and create a new instance
                if (wordProxy.CreateWord())
                    w = wordProxy.Word;
                else
                {
                    // can't get word - return Nothing by default
                }
            }
            return w;
        }

        // Returns a Word.Document dynamic
        public static dynamic GetDocument(dynamic Word, string DocumentPath)
        {
            dynamic doc = null;

            // Validate Word argument is actually Word.Application
            if (!IsWordApplicationObject(Word))
                throw new System.ArgumentException("'Word' is not the expected type of dynamic. Expected dynamic should be a Word.Application dynamic.", "Word");

            if (!File.Exists(DocumentPath))
                throw new System.ArgumentException(String.Format("File '{0}' does not exist.", DocumentPath), "DocumentPath");

            try
            {
                doc = Word.Documents.Open(DocumentPath);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return doc;
        }

        // Saves the Document as PDF to the path and file name supplied by FullName
        // Optionally shows the PDF document after it is created
        public static void SaveAsPDF(dynamic Document, string FullName, bool ShowPDF)
        {
            bool bSaved = false;
            try
            {
                Document.SaveAs(FullName, 17);
                // WdSaveFormat.wdFormatPDF
                bSaved = true;
                if (ShowPDF)
                    Document.FollowHyperlink(FullName);
            }
            catch (System.Runtime.InteropServices.COMException comException)
            {
                switch (comException.ErrorCode)
                {
                    case  // ERROR: Case labels with binary operators are unsupported : Equality
    -2146824090:
                        // Adobe Reader or equivalent not found
                        // Two known conditions will generate this exception - when
                        // a bad path is provided in the FullName parameter or 
                        // when the document is saved but Adobe Reader isn't installed
                        if (bSaved)
                        {
                            throw new System.ArgumentException("Adobe Reader or equivalent not found", "ShowPDF");
                        }
                        else
                        {
                            throw new System.ArgumentException("An exception occured while attempting to save as PDF. Please ensure that the file path is correct.", "FullName");
                        }
                        break;
                    default:
                        throw comException;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static bool IsWordDocumentObject(dynamic doc)
        {
            bool isValid = false;
            int i = 0;
            try
            {
                i = doc.ContentControls.Count;
                isValid = true;
            }
            catch (Exception ex)
            {
            }
            return isValid;
        }

        private static bool IsWordApplicationObject(dynamic app)
        {
            bool isValid = false;
            string s = null;
            try
            {
                s = app.Name;
                if (s == "Microsoft Word")
                {
                    isValid = true;
                }
            }
            catch (Exception ex)
            {
            }
            return isValid;
        }

        private static bool IsValidBookmark(dynamic doc, string sBookmark)
        {
            bool isValid = false;
            foreach (dynamic b in doc.Bookmarks)
            {
                if (b.Name == sBookmark)
                {
                    isValid = true;
                    break; // TODO: might not be correct. Was : Exit For
                }
            }
            return isValid;
        }

        private static void PopulateContentControls(List<ColumnMapping> columnMappings, IEntityObject item, WordHelper wordProxy)
        {
            foreach (ColumnMapping mapping in columnMappings)
            {
                if (mapping.TableField != null && mapping.TableField.Name != "<Ignore>")
                {
                    string value = item.Details.Properties[mapping.TableField.Name].Value.ToString();
                    wordProxy.SetContentControl(mapping.OfficeColumn, value);
                }
            }
        }
    }
}