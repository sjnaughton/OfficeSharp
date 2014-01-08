using System;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices.Automation;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace OfficeSharp.Word
{
    class WordHelper : IDisposable
    {
        private dynamic _word;

        private dynamic _document;
        public dynamic Word
        {
            get
            {
                try
                {
                    return _word;
                }
                catch (Exception ex)
                {
                    return null;
                }
            }
        }

        public dynamic Document
        {
            get
            {
                try
                {
                    if (_document == null)
                    {
                        if (_word.Documents.Count > 0)
                            _document = _word.ActiveDocument;
                        else
                            _document = _word.Documents.Add();
                    }

                    return _document;
                }
                catch (Exception ex)
                {
                    return null;
                }
            }
            set { _document = value; }
        }

        // Use this method to retrieve active Word session
        public bool GetWord()
        {
            try
            {
                // If GetObject throws an exception, then Word is 
                // either not running or is not available.
                _word = AutomationFactory.GetObject("Word.Application");
                return true;
            }
            catch
            {
                return false;
            }
        }

        // Use this method to explicity create a new instance of Word
        public bool CreateWord()
        {
            try
            {
                // If CreateObject throws an exception, then Word is not available.
                _word = AutomationFactory.CreateObject("Word.Application");
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void OpenDocument(string filePath)
        {
            dynamic o = null;
            o = Word;
            if (o != null)
            {
                if (File.Exists(filePath))
                    o.Documents.Open(filePath);
                else
                    throw new ArgumentException("The file '" + filePath + "' does not exist.", "FilePath");
            }
            else
                throw new InvalidProgramException("Error: Could not open Word document. Please make sure Word is installed.");
        }

        public void CreateDocument()
        {
            dynamic o = null;
            o = Word;
            if (o != null)
                Document = o.Documents.Add();
        }

        public void CloseDocument()
        {
            if (Document != null)
            {
                Document.Save();
                Document.Close();
                Document = null;
            }
        }

        public void ShowDocument()
        {
            dynamic o = null;
            o = Word;
            if (o != null)
                o.Visible = true;
        }

        // Adds a table to the end of the Document
        public dynamic AddTable(Int32 RowCount, Int32 ColumnCount)
        {
            dynamic o = null;
            o = null;

            if (Document != null)
            {
                try
                {
                    Document.Tables.Add(Document.Bookmarks.Item("\\endofdoc").Range, RowCount, ColumnCount);
                }
                catch (Exception Ex)
                {
                    // Method call works, but throws a COM exception on returning back to Silverlight
                }

                if (_document.Tables.Count > 0)
                {
                    // use last table
                    o = Document.Tables(1);
                    o.ApplyStyleHeadingRows = true;
                    o.Style = "Medium Shading 1 - Accent 1";
                    o.ApplyStyleFirstColumn = false;
                }
            }
            return o;
        }

        // Adds a table to the specified range of the Document
        public dynamic AddTable(Int32 RowCount, Int32 ColumnCount, dynamic LocationRange)
        {
            dynamic o = null;
            o = null;

            if (Document != null)
            {
                try
                {
                    Document.Tables.Add(LocationRange, RowCount, ColumnCount);
                }
                catch (Exception Ex)
                {
                    // Method call works, but throws a COM exception on returning back to Silverlight
                }

                if (LocationRange.Tables.Count > 0)
                {
                    // use last table
                    o = LocationRange.Tables(1);
                    o.ApplyStyleHeadingRows = true;
                    o.Style = "Medium Shading 1 - Accent 1";
                    o.ApplyStyleFirstColumn = false;
                }
            }
            return o;
        }

        public void SetTableCell(dynamic Table, Int32 Row, Int32 Column, string Value)
        {
            Table.Cell(Row, Column).Range.Text = Value;
        }

        public Int32 ContentControlCount()
        {
            return Document.ContentControls.Count;
        }

        public void SetContentControl(string ContentControlTitle, string Value)
        {
            foreach (dynamic cc in Document.ContentControls)
            {
                if (cc.Title == ContentControlTitle)
                {
                    cc.Range.Text = Value;
                    // Populate every content control having the title = ContentControlTitle
                    //Exit For
                }
            }
        }

        public IEnumerable<ContentControl> GetContentControls
        {
            get
            {
                List<ContentControl> controls = new List<ContentControl>();

                foreach (dynamic cc in Document.ContentControls)
                {
                    controls.Add(new ContentControl(cc.Title));
                }
                return controls;
            }
        }

        #region "IDisposable Support"
        // To detect redundant calls
        private bool disposedValue;

        // IDisposable
        protected virtual void Dispose(bool disposing)
        {

            if (!this.disposedValue)
            {
                Word.DisplayAlerts = false;
                Word.Quit();

                _word = null;
            }
            this.disposedValue = true;
        }

        // This code added by Visual Basic to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }


    internal class ContentControl
    {
        public ContentControl(string Title)
        {
            this.Title = Title;
        }
        public string Title { get; set; }
    }
}
