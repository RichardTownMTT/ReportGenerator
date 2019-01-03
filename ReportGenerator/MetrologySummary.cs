using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;

namespace ReportGenerator
{
    public class MetrologySummary
    {
        public static bool CreateImperialMetSum(string filename)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Document doc;
            try
            {
                doc = app.Documents.Open(filename);
                app.Visible = true;

                bool isMetSum = ValidateMetSumDocument(doc);
                if (!isMetSum)
                {
                    MessageBox.Show("Document is not a Metrology Summary Document");
                    return false;
                }

                CreateMetSumDoc(doc);

                int fileExtensionInt = filename.IndexOf(".docx");
                string saveFileName = filename.Insert(fileExtensionInt, " - Imperial");

                doc.SaveAs2(saveFileName);

                app.Quit(WdSaveOptions.wdDoNotSaveChanges);
                app = null;

                return true;
            }
            catch (Exception e)
            {
                throw;
            }
            finally
            {
                //Make sure that the word app has been displosed of.
                if (app != null)
                {
                    app.Quit();
                    app = null;
                }
            }
        }

        private static void CreateMetSumDoc(Document doc)
        {
            string forwardslashLengthPattern = "/[0-9]{3,}mm";
            string singleValuePattern = "[0-9].?[0-9]{2,}mm";

            Table metTable = doc.Tables[1];
            int noRows = metTable.Rows.Count;
            int noColumns = metTable.Columns.Count;

            //Loop through all cells in table and look for mm
            int rowCounter = 1;
            int columnCounter = 1;

            //Set the title to the imperial version
            metTable.Cell(1, 1).Range.Text = "METROLOGY SUMMARY - Imperial";

            while (rowCounter <= noRows)
            {
                while (columnCounter <= noColumns)
                {
                    try
                    {
                        string cellText = metTable.Cell(rowCounter, columnCounter).Range.Text;
                        Match lengthFound = Regex.Match(cellText, forwardslashLengthPattern);

                        if (lengthFound.Success)
                        {
                            string textToReplace = lengthFound.ToString();
                            string replacement = ConvertMmToInchLength(textToReplace);
                            cellText = cellText.Replace(textToReplace, replacement);
                            metTable.Cell(rowCounter, columnCounter).Range.Text = cellText;
                        }

                        MatchCollection singleFoundCollection = Regex.Matches(cellText, singleValuePattern);

                        foreach (var item in singleFoundCollection)
                        {
                            string replace = item.ToString();
                            string replacementText = ConvertMmToInch(replace);
                            cellText = cellText.Replace(replace, replacementText);
                            metTable.Cell(rowCounter, columnCounter).Range.Text = cellText;
                        }
                    }
                    catch (Exception ex)
                    {
                        //Catch any cells which are outside of the count - e.g. merged cells will change the number in the row / column
                    }

                    columnCounter++;
                }

                columnCounter = 1;
                rowCounter++;
            }
        }

        private static string ConvertMmToInch(string input)
        {
            const Decimal inchConversion = 25.4M;
            if (string.IsNullOrEmpty(input))
            {
                return "";
            }

            string retval;

            //Remove mm and /
            input = input.Replace("mm", "");


            if (Decimal.TryParse(input, out decimal converted))
            {
                converted = converted / inchConversion;
            }
            else
            {
                MessageBox.Show("Unable to convert mm for " + input);
                throw new Exception("Unable to convert mm for " + input);
            }

            retval = Math.Round(converted, 6).ToString() + "in";
            return retval;
        }

        private static string ConvertMmToInchLength(string input)
        {
            const Decimal inchConversion = 25.4M;
            if (string.IsNullOrEmpty(input))
            {
                return "";
            }

            string retval;

            //Remove mm and /
            input = input.Replace("mm", "");
            input = input.Replace("/", "");


            if (Decimal.TryParse(input, out decimal converted))
            {
                converted = converted / inchConversion;
            }
            else
            {
                MessageBox.Show("Unable to convert mm for " + input);
                throw new Exception("Unable to convert mm for " + input);
            }

            retval = "/" + Math.Round(converted, 4).ToString() + "in";
            return retval;
        }

        private static bool ValidateMetSumDocument(Document doc)
        {
            int tblCount = doc.Tables.Count;

            if (tblCount == 0)
            {
                return false;
            }
            if (tblCount > 1)
            {
                return false;
            }

            Table wordTable = doc.Tables[1];

            Cell firstCell = wordTable.Cell(1, 1);
            string titleText = firstCell.Range.Text;

            if (titleText.Contains("Metrology Summary") || titleText.Contains("METROLOGY SUMMARY"))
            {
                return true;
            }

            return false;
        }
    }
}