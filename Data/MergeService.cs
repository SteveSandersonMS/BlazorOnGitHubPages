using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using System.IO;
using System.Reflection;
using BlazorOnGitHubPages.Models;
using System.Globalization;


namespace BlazorOnGitHubPages.Data
{
    public class MergeService
    {
        public MemoryStream MergeExcel(MemoryStream csvStream, MemoryStream xlsStream)
        {
            //Create an instance of ExcelEngine

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Load xls to file stream

               // DateTime now = DateTime.Now;
               // string csvFileName = now.ToString() + ".csv";
               // string xlsFileName = now.ToString() + ".xlsx";
                FileStream csvInputStream = new FileStream("test.csv", FileMode.Create);

                //mStream.Position = 0;
                csvStream.WriteTo(csvInputStream);
                csvStream.Close();

                //reset write position from tail to head
                csvInputStream.Position = 0;
                IWorkbook csvworkbook = application.Workbooks.Open(csvInputStream);
                IWorksheet csvworksheet = csvworkbook.Worksheets[0];

                int csvLastRow = csvworksheet.UsedRange.LastRow;
                // int csvLastCol = csvworksheet.UsedRange.LastColumn;

                Dictionary<string, List<csvData>> csvPairs = new Dictionary<string, List<csvData>>();

               // csvData[] csvDatas = new csvData[csvLastRow - 1];
                for(int i = 2; i <= csvLastRow; i++)
                {
                    csvData newCsvData = new csvData();

                    newCsvData.Name = csvworksheet.Range["A" + i.ToString()].Text;
                    newCsvData.Case_Status__c = csvworksheet.Range[("B" + i.ToString())].Text;
                    newCsvData.Case_Sub_Status__c = csvworksheet.Range[("C" + i.ToString())].Text;
                    newCsvData.Confirmed_Hearing_Date__c = csvworksheet.Range[("D" + i.ToString())].Text;
                    newCsvData.Lead_Advocate1__r_Name = csvworksheet.Range[("E" + i.ToString())].Text;
                    newCsvData.Advocate_Status__c = csvworksheet.Range[("F" + i.ToString())].Text;

                    //csvPairs.Add(newCsvData.Name + newCsvData.Confirmed_Hearing_Date__c, newCsvData);
                    if (!csvPairs.ContainsKey(newCsvData.Name))
                    {
                        csvPairs.Add(newCsvData.Name, new List<csvData> { newCsvData });
                    }
                    else
                    {
                        csvPairs[newCsvData.Name].Add(newCsvData);
                    }
                }
                csvInputStream.Close();


                FileStream xlsInputStream = new FileStream("test.xlsx", FileMode.Create);

                xlsStream.WriteTo(xlsInputStream);
                xlsStream.Close();

                xlsInputStream.Position = 0;
                IWorkbook xlsworkbook = application.Workbooks.Open(xlsInputStream);
                IWorksheet xlsworksheet = xlsworkbook.Worksheets[0];

                

                List<xlsData> xlsList = new List<xlsData>();


                int xlsLastRow = xlsworksheet.UsedRange.LastRow;

                //xlsData[] xlsDatas = new xlsData[xlsLastRow - 1];
                for (int i = 2; i <= xlsLastRow; i++)
                {
                    //string joinKey = xlsworksheet.Range[("D" + i.ToString())].Text + xlsworksheet.Range[("G" + i.ToString())].Text;
                   // string joinKey = xlsworksheet.Range[("D" + i.ToString())].Text;

                    //if (csvPairs.ContainsKey(joinKey))
                    //{
                    //    csvData match = csvPairs[joinKey];
                    //    if (null != match)
                    //    {
                            
                    //        xlsworksheet.Range["L" + i.ToString()].Text = match.Case_Status__c;
                    //        xlsworksheet.Range["M" + i.ToString()].Text = match.Case_Sub_Status__c;
                    //        xlsworksheet.Range["N" + i.ToString()].Text = match.Lead_Advocate1__r_Name;
                            
                    //        /*
                    //        xlsworksheet.Range["L" + i.ToString()].Text = "1";
                    //        xlsworksheet.Range["M" + i.ToString()].Text = "2";
                    //        xlsworksheet.Range["N" + i.ToString()].Text = "3";
                    //        */

                    //    }
                    //}
                    xlsData newXlsData = new xlsData();
                    //newXlsData.Id = (i - 1).ToString();
                    //newXlsData.LastName = xlsworksheet.Range["A" + i.ToString()].Text;
                    //newXlsData.FirstName = xlsworksheet.Range["B" + i.ToString()].Text;
                    //newXlsData.MiddleName = xlsworksheet.Range[("C" + i.ToString())].Text;

                    ////newXlsData.Account = xlsworksheet.Range[("D" + i.ToString())].Text;
                    //newXlsData.Account = xlsworksheet.Range["B" + i.ToString()].Text + " " + xlsworksheet.Range["A" + i.ToString()].Text;

                    //newXlsData.Last4SSN = xlsworksheet.Range[("E" + i.ToString())].Text;
                    //newXlsData.HearingOfficeWithJurisdiction = xlsworksheet.Range[("F" + i.ToString())].Text;
                    //newXlsData.HearingScheduledDate = xlsworksheet.Range[("G" + i.ToString())].Text;
                    //newXlsData.HearingTime = xlsworksheet.Range[("H" + i.ToString())].Text;
                    //newXlsData.ALJLastName = xlsworksheet.Range[("I" + i.ToString())].Text;
                    //newXlsData.MedicalExpert = xlsworksheet.Range[("J" + i.ToString())].Text;
                    //newXlsData.VocationalExpert = xlsworksheet.Range[("K" + i.ToString())].Text;


                    newXlsData.Id = (i - 1).ToString();
                    newXlsData.LastName = xlsworksheet.Range["A" + i.ToString()].Text;
                    newXlsData.FirstName = xlsworksheet.Range["B" + i.ToString()].Text;
                   // newXlsData.MiddleName = xlsworksheet.Range[("C" + i.ToString())].Text;

                    //newXlsData.Account = xlsworksheet.Range[("D" + i.ToString())].Text;
                    newXlsData.Account = xlsworksheet.Range["B" + i.ToString()].Text + " " + xlsworksheet.Range["A" + i.ToString()].Text;
                    newXlsData.Last4SSN = "";
                    newXlsData.MiddleName = "";
                   // newXlsData.Last4SSN = xlsworksheet.Range[("E" + i.ToString())].Text;
                    newXlsData.HearingOfficeWithJurisdiction = xlsworksheet.Range[("C" + i.ToString())].Text;
                    //newXlsData.HearingScheduledDate = xlsworksheet.Range[("F" + i.ToString())].Text;
                    newXlsData.HearingScheduledDate = xlsworksheet.Range[("D" + i.ToString())].DateTime;
                    newXlsData.HearingTime = xlsworksheet.Range[("E" + i.ToString())].Text;
                    newXlsData.ALJLastName = xlsworksheet.Range[("F" + i.ToString())].Text;
                    newXlsData.MedicalExpert = xlsworksheet.Range[("G" + i.ToString())].Text;
                    newXlsData.VocationalExpert = xlsworksheet.Range[("H" + i.ToString())].Text;

                    xlsList.Add(newXlsData);
                }

                //xlsworksheet.Range["A1"].Text = "ID";
                //xlsworksheet.Range["B1"].Text = "Last Name";
                //xlsworksheet.Range["C1"].Text = "First Name";
                //xlsworksheet.Range["D1"].Text = "Middle Name";
                //xlsworksheet.Range["E1"].Text = "Account";
                //xlsworksheet.Range["F1"].Text = "Last 4 SSN";
                //xlsworksheet.Range["G1"].Text = "Hearing Office with Jurisdiction";
                //xlsworksheet.Range["H1"].Text = "Hearing Scheduled Date";
                //xlsworksheet.Range["I1"].Text = "Hearing Time";
                //xlsworksheet.Range["J1"].Text = "ALJ Last Name";
                //xlsworksheet.Range["K1"].Text = "Medical Expert";
                //xlsworksheet.Range["L1"].Text = "Vocational Expert";
                //xlsworksheet.Range["M1"].Text = "Case_Status__c";
                //xlsworksheet.Range["N1"].Text = "Case_Sub_Status__c";
                //xlsworksheet.Range["O1"].Text = "Confirmed_Hearing_Date__c";
                //xlsworksheet.Range["P1"].Text = "Lead_Advocate1__r.Name";


                //xlsworksheet.Range["A1"].Text = "ID";
                xlsworksheet.Range["A1"].Text = "Last Name";
                xlsworksheet.Range["B1"].Text = "First Name";
                //xlsworksheet.Range["D1"].Text = "Middle Name";
                xlsworksheet.Range["C1"].Text = "Account";
               // xlsworksheet.Range["F1"].Text = "Last 4 SSN";
                xlsworksheet.Range["D1"].Text = "Hearing Office with Jurisdiction";
                xlsworksheet.Range["E1"].Text = "Hearing Scheduled Date";
                xlsworksheet.Range["F1"].Text = "Hearing Time";
                xlsworksheet.Range["G1"].Text = "ALJ Last Name";
                xlsworksheet.Range["H1"].Text = "Medical Expert";
                xlsworksheet.Range["I1"].Text = "Vocational Expert";
                xlsworksheet.Range["J1"].Text = "Case_Status__c";
                xlsworksheet.Range["K1"].Text = "Case_Sub_Status__c";
                xlsworksheet.Range["L1"].Text = "Confirmed_Hearing_Date__c";
                xlsworksheet.Range["M1"].Text = "Lead_Advocate1__r.Name";
                xlsworksheet.Range["N1"].Text = "Advocate_Status__c";

                int curRow = 2;

              




                for (int i = 2; i <= xlsList.Count + 1; i++)
                {

                    //xlsworksheet.Range["A" + curRow.ToString()].Text = xlsList[i - 2].Id;
                    //xlsworksheet.Range["B" + curRow.ToString()].Text = xlsList[i - 2].LastName;
                    //xlsworksheet.Range["C" + curRow.ToString()].Text = xlsList[i - 2].FirstName;
                    //xlsworksheet.Range["D" + curRow.ToString()].Text = xlsList[i - 2].MiddleName;
                    //xlsworksheet.Range["E" + curRow.ToString()].Text = xlsList[i - 2].Account;
                    //xlsworksheet.Range["F" + curRow.ToString()].Text = xlsList[i - 2].Last4SSN;
                    //xlsworksheet.Range["G" + curRow.ToString()].Text = xlsList[i - 2].HearingOfficeWithJurisdiction;
                    //xlsworksheet.Range["H" + curRow.ToString()].Text = xlsList[i - 2].HearingScheduledDate;
                    //xlsworksheet.Range["I" + curRow.ToString()].Text = xlsList[i - 2].HearingTime;
                    //xlsworksheet.Range["J" + curRow.ToString()].Text = xlsList[i - 2].ALJLastName;
                    //xlsworksheet.Range["K" + curRow.ToString()].Text = xlsList[i - 2].MedicalExpert;
                    //xlsworksheet.Range["L" + curRow.ToString()].Text = xlsList[i - 2].VocationalExpert;


                    //xlsworksheet.Range["A" + curRow.ToString()].Text = xlsList[i - 2].Id;
                    xlsworksheet.Range["A" + curRow.ToString()].Text = xlsList[i - 2].LastName;
                    xlsworksheet.Range["B" + curRow.ToString()].Text = xlsList[i - 2].FirstName;
                    //xlsworksheet.Range["D" + curRow.ToString()].Text = xlsList[i - 2].MiddleName;
                    xlsworksheet.Range["C" + curRow.ToString()].Text = xlsList[i - 2].Account;
                   // xlsworksheet.Range["F" + curRow.ToString()].Text = xlsList[i - 2].Last4SSN;
                    xlsworksheet.Range["D" + curRow.ToString()].Text = xlsList[i - 2].HearingOfficeWithJurisdiction;
                    //xlsworksheet.Range["E" + curRow.ToString()].Text = xlsList[i - 2].HearingScheduledDate;
                    xlsworksheet.Range["E" + curRow.ToString()].DateTime = xlsList[i - 2].HearingScheduledDate;
                    xlsworksheet.Range["F" + curRow.ToString()].Text = xlsList[i - 2].HearingTime;
                    xlsworksheet.Range["G" + curRow.ToString()].Text = xlsList[i - 2].ALJLastName;
                    xlsworksheet.Range["H" + curRow.ToString()].Text = xlsList[i - 2].MedicalExpert;
                    xlsworksheet.Range["I" + curRow.ToString()].Text = xlsList[i - 2].VocationalExpert;

                    // var temp = xlsList[i - 2];

                    if (csvPairs.ContainsKey(xlsList[i - 2].Account))
                    {
                        //List<csvData> csvList = csvPairs[xlsList[i - 2].Account];
                        List<csvData> csvList = csvPairs[xlsList[i - 2].Account];
                        //List<csvData> csvList = csvPairs[xlsList[i - 2].Account];
                        for (int j = 0; j < csvList.Count; j++)
                        {                         
                            if(0 == j)
                            {
                                xlsworksheet.Range["J" + curRow.ToString()].Text = csvList[j].Case_Status__c;
                                xlsworksheet.Range["K" + curRow.ToString()].Text = csvList[j].Case_Sub_Status__c;
                                xlsworksheet.Range["L" + curRow.ToString()].Text = csvList[j].Confirmed_Hearing_Date__c;
                                //xlsworksheet.Range["M" + curRow.ToString()].Text = csvList.Count.ToString();
                                xlsworksheet.Range["M" + curRow.ToString()].Text = csvList[j].Lead_Advocate1__r_Name;
                                xlsworksheet.Range["N" + curRow.ToString()].Text = csvList[j].Advocate_Status__c;
                            }
                            
                            else
                            {
                                curRow++;
                                xlsworksheet.Range["A" + curRow.ToString()].Text = " ";
                                xlsworksheet.Range["B" + curRow.ToString()].Text = " ";
                                xlsworksheet.Range["C" + curRow.ToString()].Text = " ";
                                xlsworksheet.Range["D" + curRow.ToString()].Text = " ";
                                xlsworksheet.Range["E" + curRow.ToString()].Text = " ";
                                xlsworksheet.Range["F" + curRow.ToString()].Text = " ";
                                xlsworksheet.Range["G" + curRow.ToString()].Text = " ";
                                xlsworksheet.Range["H" + curRow.ToString()].Text = " ";
                                xlsworksheet.Range["I" + curRow.ToString()].Text = " ";
                                xlsworksheet.Range["J" + curRow.ToString()].Text = csvList[j].Case_Status__c;
                                xlsworksheet.Range["K" + curRow.ToString()].Text = csvList[j].Case_Sub_Status__c;
                                xlsworksheet.Range["L" + curRow.ToString()].Text = csvList[j].Confirmed_Hearing_Date__c;
                                //xlsworksheet.Range["M" + curRow.ToString()].Text = csvList.Count.ToString();
                                xlsworksheet.Range["M" + curRow.ToString()].Text = csvList[j].Lead_Advocate1__r_Name;
                                xlsworksheet.Range["N" + curRow.ToString()].Text = csvList[j].Advocate_Status__c;

                            }

                        }
                        
                    }
                    else
                    {
                        xlsworksheet.Range["J" + curRow.ToString()].Text = "NULL";
                        xlsworksheet.Range["K" + curRow.ToString()].Text = "NULL";
                        xlsworksheet.Range["L" + curRow.ToString()].Text = "NULL";
                        xlsworksheet.Range["M" + curRow.ToString()].Text = "NULL";
                        xlsworksheet.Range["N" + curRow.ToString()].Text = "NULL";
                    }
                    curRow++;



                }

                //Create a workbook
                //IWorkbook workbook = application.Workbooks.Create(1);
                //IWorksheet xlsworksheet = xlsworkbook.Worksheets[0];

                //int lastRow = worksheet.UsedRange.LastRow;
                //int lastCol = worksheet.UsedRange.LastColumn;

                //Disable gridlines in the worksheet
                //worksheet.IsGridLinesVisible = false;

                //Enter values to the cells from A5 to C5
                //xlsworksheet.Range["A1"].Text = csvDatas[1].Lead_Advocate1__r_Name;
                // worksheet.Range["B5"].Text = "Tony";
                // worksheet.Range["C5"].Text = "HR";

                //File.Delete("test.csv");
                //File.Delete("test.xlsx");


                using (MemoryStream stream = new MemoryStream())
                {
                    //Save the created Excel document to MemoryStream
                    xlsworkbook.SaveAs(stream);
                    xlsInputStream.Close();
                    //csvworkbook.SaveAs(stream, ",");
                    File.Delete("test.csv");
                    File.Delete("test.xlsx");
                    return stream;
                }
            }
        }
    }
}
