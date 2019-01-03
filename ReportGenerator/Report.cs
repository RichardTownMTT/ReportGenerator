using Microsoft.Office.Interop.Word;
using System;

namespace ReportGenerator
{
    public class Report
    {
        private string _filename;
        private Document _reportDocument;
        private Microsoft.Office.Interop.Word.Application _app;

        public Report(string filename)
        {
            _filename = filename;
            LoadReportDocument();
        }

        private void LoadReportDocument()
        {
            _app = new Microsoft.Office.Interop.Word.Application();
            try
            {
                _reportDocument = _app.Documents.Open(_filename);
                _app.Visible = true;
            }
            catch (Exception e)
            {
                //Make sure that the word app has been displosed of.
                if (_app != null)
                {
                    _app.Quit();
                    _app = null;
                }
                throw;
            }
        }

        public void SetTestNumbers()
        {
            if (_reportDocument == null || _app == null)
            {
                throw new NullReferenceException("Report document not loaded");
            }
            //The test number needs setting on all tests which contain the word test in the first cell of the table.  Test number format is "Test #1"

            int testCounter = 1;
            int totalTables = _reportDocument.Tables.Count;

            for (int tableCounter = 1; tableCounter <= totalTables; tableCounter++)
            {
                Table currentTable = _reportDocument.Tables[tableCounter];

                //Check the first cell contains test
                string cellText = currentTable.Cell(1, 1).Range.Text;
                if (cellText.Contains("Test"))
                {
                    //Set the test number 
                    currentTable.Cell(1, 1).Range.Text = "Test #" + testCounter;
                    testCounter++;
                }
            }
        }

        //Saves the file with the current date / time on the file
        public bool Save()
        {
            if (_reportDocument == null || _app == null)
            {
                throw new NullReferenceException("Report document not loaded");
            }

            int fileExtensionInt = _filename.IndexOf(".docx");
            string dateStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string saveFileName = _filename.Insert(fileExtensionInt, " - " + dateStamp);

            _reportDocument.SaveAs2(saveFileName);

            _app.Quit(WdSaveOptions.wdDoNotSaveChanges);
            _app = null;

            return true;
        }
    }
}