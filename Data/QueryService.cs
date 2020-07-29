using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using System.IO;
using System.Reflection;
using System.Globalization;

namespace BlazorOnGitHubPages.Data
{
    public class QueryService
    {
        public string CreateQuery(MemoryStream mStream, string fName)
        {
            //Create an instance of ExcelEngine
            string queryStr = "SELECT Name, Case_Status__c, Case_Sub_Status__c, Confirmed_Hearing_Date__c, Lead_Advocate1__r.Name, Advocate_Status__c FROM Account WHERE ";

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Load xls to file stream


                FileStream inputStream = new FileStream(fName, FileMode.Create);

                //mStream.Position = 0;
                mStream.WriteTo(inputStream);
                mStream.Close();

                inputStream.Position = 0;
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Create a workbook
                //IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                int lastRow = worksheet.UsedRange.LastRow;
                int lastCol = worksheet.UsedRange.LastColumn;

                //Disable gridlines in the worksheet
                //worksheet.IsGridLinesVisible = false;

                //Enter values to the cells from A5 to C5
                //worksheet.Range["A5"].Number = 4;
                //worksheet.Range["B5"].Text = "Tony";
                //worksheet.Range["C5"].Text = "HR";

                char fNameColLetter = 'A', lNameColLetter = 'A', acctColLetter = (char)('A' + lastCol);
                string fNameCol, lNameCol;
                for (int j = 0; j < lastCol; j++)
                {
                    fNameCol = fNameColLetter.ToString() + "1";
                    if (worksheet.Range[fNameCol].Text != "First Name")
                    {
                        fNameColLetter++;
                    }
                    else
                        break;
                }

                for (int j = 0; j < lastCol; j++)
                {
                    lNameCol = lNameColLetter.ToString() + "1";
                    if (worksheet.Range[lNameCol].Text != "Last Name")
                    {
                        lNameColLetter++;
                    }
                    else
                        break;
                }

                //worksheet.Range[fNameCol].Text = "Test";
                //worksheet.Range[lNameCol].Text = "Test";
               // string acctCol = acctColLetter.ToString() + "1";
             //   worksheet.Range[acctCol].Text = "Name";
                int i = 2;


                while (i <= lastRow)
                {
                    string fCol = fNameColLetter.ToString() + i.ToString();
                //    string accountCol = acctColLetter.ToString() + i.ToString();
                    string lCol = lNameColLetter.ToString() + i.ToString();
                  //  worksheet.Range[accountCol].Text = worksheet.Range[fCol].Text + " " + worksheet.Range[lCol].Text;
                    //WHERE Name = 'Gina Jensen'  OR Name = 'Robin Oka'
                    //OR Name = 'Gordon Bolt'
                    //OR Name = 'Victor Lopez'
                    //OR Name = 'Joanne Abernathy'
                    queryStr += (" Name = '" +(worksheet.Range[fCol].Text + " " + worksheet.Range[lCol].Text) + "'");
                    if (i < lastRow)
                        queryStr += " OR ";
                    i++;
                }


                //IWorksheet newWorksheet;
                //newWorksheet = workbook.Worksheets.Create("Generated Query");
                //newWorksheet.Range["A1"].Text = queryStr;
                inputStream.Close();
                return queryStr;
                //using (MemoryStream stream = new MemoryStream())
                //{
                //    //Save the created Excel document to MemoryStream
                //    workbook.SaveAs(stream);
                //    return stream;
                //}
            }
        }
    }
}
