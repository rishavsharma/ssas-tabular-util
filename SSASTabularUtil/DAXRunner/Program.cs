using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ADOTabular;
using ADOTabular.AdomdClientWrappers;
using ClosedXML.Excel;
using CommandLine;

namespace DAXRunner
{
    class Program
    {
        public class Options
        {
            [Option('i', "input", Required = true, HelpText = "Input Excel")]
            public String inExcel { get; set; }
            [Option('o', "out", Required = true, HelpText = "Out Directory")]
            public String outDir { get; set; }
            [Option('t', "threads", Required = true, HelpText = "Parallel Threads")]
            public int threads { get; set; }
        }
        public static Utils.StatusRow runComparison(DataRow row, DirectoryInfo od, Utils.StatusRow statusrow)
        {
            string testName = row["NAME"].ToString();
            Console.WriteLine("Test [" + testName + "] Started executing");
            string srcSSAS = "Data Source = " + row["SRC_SSAS"].ToString();
            string tgtSSAS = "Data Source = " + row["TGT_SSAS"].ToString();
            string srcSSASModel = row["SRC_SSAS_MODEL"].ToString();
            string tgtSSASModel = row["TGT_SSAS_MODEL"].ToString();
            string srcQuery = row["SRC_DAX"].ToString();
            string tgtQuery = row["TGT_DAX"].ToString();
            DateTime srcStartTime = DateTime.Now, srcEndTime = DateTime.Now, tgtStartTime = DateTime.Now, tgtEndTime = DateTime.Now;

            String srcErrorMsg = "", tgtErrorMsg = "", status = "PASS";
            int srcNoOfRows = 0, tgtNoOfRows = 0, srcNoOfColumn = 0, tgtNoOfColumn = 0;
            bool srcError = false, tgtError = false;
            ADOTabularConnection srcConx = null;
            ADOTabularConnection tgtConx = null;
            try
            {



                //Console.WriteLine("--------------------Src DAX---------------------------");
                //Console.WriteLine(srcQuery);
                //Console.WriteLine("--------------------tgt DAX---------------------------");
                //Console.WriteLine(tgtQuery);
                DataTable srcDT = new DataTable();
                DataTable tgtDT = new DataTable();
                DataTable ret = new DataTable();

                try
                {
                    srcConx = new ADOTabularConnection(srcSSAS, ADOTabular.Enums.AdomdType.AnalysisServices);
                    srcConx.ChangeDatabase(srcSSASModel);
                    srcStartTime = DateTime.Now;
                    srcDT = srcConx.ExecuteDaxQueryDataTable(srcQuery);

                }
                catch (Exception ee)
                {
                    Console.WriteLine(ee.Message);
                    Console.WriteLine(ee.StackTrace);
                    srcErrorMsg = ee.Message;
                    srcError = true;
                }
                finally
                {
                    srcEndTime = DateTime.Now;
                    srcConx.Close();
                }

                try
                {
                    tgtConx = new ADOTabularConnection(tgtSSAS, ADOTabular.Enums.AdomdType.AnalysisServices);
                    tgtConx.ChangeDatabase(tgtSSASModel);
                    tgtStartTime = DateTime.Now;
                    tgtDT = tgtConx.ExecuteDaxQueryDataTable(tgtQuery);
                }
                catch (Exception ee)
                {
                    Console.WriteLine(ee.Message);
                    Console.WriteLine(ee.StackTrace);
                    tgtErrorMsg = ee.Message;
                    tgtError = true;
                }
                finally
                {
                    tgtEndTime = DateTime.Now;
                    tgtConx.Close();
                }

                if (srcError || tgtError)
                {
                    status = "FAILED";
                }
                else
                {
                    ret = Utils.getDifferentRecords(srcDT, tgtDT);
                    srcNoOfRows = srcDT.Rows.Count;
                    tgtNoOfRows = tgtDT.Rows.Count;
                    srcNoOfColumn = srcDT.Columns.Count;
                    tgtNoOfColumn = tgtDT.Columns.Count;
                    if (srcNoOfRows == 0 || tgtNoOfRows == 0)
                    {
                        status = "FAILED";
                        srcErrorMsg = "No Data:" + srcNoOfRows;
                        tgtErrorMsg = "No Data:" + tgtNoOfRows;
                    }
                    else if (srcNoOfRows != tgtNoOfRows)
                    {
                        status = "FAILED";
                        srcErrorMsg = "Number of rows:" + srcNoOfRows;
                        tgtErrorMsg = "Number of rows:" + tgtNoOfRows;
                    }
                    else if (ret.Rows.Count > 0)
                    {
                        status = "FAILED";
                        srcErrorMsg = "Data Mismatch";
                    }
                    else if (srcNoOfColumn != tgtNoOfColumn)
                    {
                        status = "FAILED";
                        srcErrorMsg = "Number of Columns:" + srcNoOfColumn;
                        tgtErrorMsg = "Number of Columns:" + tgtNoOfColumn;
                    }
                }

                DataTable statusDT = Utils.getStatusDataTable();
                statusrow.NAME = testName;
                statusrow.STATUS = status;
                statusrow.SRC_SSASQueryStartTime = srcStartTime;
                statusrow.SRC_SSASQueryEndTime = srcEndTime;
                statusrow.SRC_SSASQueryExecutionTime = (double)srcEndTime.Subtract(srcStartTime).Seconds;
                statusrow.TGT_SSASQueryStartTime = tgtStartTime;
                statusrow.TGT_SSASQueryEndTime = tgtEndTime;
                statusrow.TGT_SSASQueryExecutionTime = (double)tgtEndTime.Subtract(tgtStartTime).Seconds;
                statusrow.SRC_Exception = srcErrorMsg;
                statusrow.TGT_Exception = tgtErrorMsg;
                statusrow.SRC_DAX = srcQuery;
                statusrow.TGT_DAX = tgtQuery;

                Utils.getStatusRow(statusDT, statusrow);
                using (var workbook = new XLWorkbook())
                {
                    try
                    {
                        workbook.AddWorksheet(statusDT, "Status");
                        workbook.AddWorksheet(srcDT, "Source");
                        workbook.AddWorksheet(tgtDT, "Target");
                        workbook.AddWorksheet(ret, "Delta");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        Console.WriteLine(e.StackTrace);
                    }
                    finally
                    {
                        workbook.SaveAs(od.FullName + "\\" + status + "_" + testName + ".xlsx");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                statusrow.STATUS = "FAILED";
            }
            finally
            {
                srcConx.Close();
                tgtConx.Close();
                //Utils.getStatusRow(overallStatusDT, statusrow);
                //Console.WriteLine("Tast [" + testName + "] Finished with Status:" + statusrow.STATUS);
            }

            return statusrow;

        }
        static void Main(string[] args)
        {
            //Console.WriteLine("Set threads done: " + System.Threading.ThreadPool.SetMaxThreads(10,10));
            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       DataTable excel = Utils.ImportExceltoDatatable(o.inExcel, "Queries");

                       DirectoryInfo od = new DirectoryInfo(o.outDir);
                       od.Create();
                       DataTable overallStatusDT = Utils.getStatusDataTable();
                       using (var statusWorkbook = new XLWorkbook())
                       {
                           try
                           {
                               var listOfTasks = new List<Task<Utils.StatusRow>>();
                               using (SemaphoreSlim concurrency = new SemaphoreSlim(o.threads))
                               {
                                   
                                   foreach (DataRow row in excel.Rows)
                                   {
                                       Utils.StatusRow statusrow = new Utils.StatusRow();
                                       string testName = row["NAME"].ToString();
                                       if (String.IsNullOrEmpty(testName))
                                       {
                                           continue;
                                       }
                                       concurrency.Wait();
                                       //Task<Utils.StatusRow> n = new Task<Utils.StatusRow>(() => Program.runComparison(row, od, statusrow));
                                       Task<Utils.StatusRow> n = Task<Utils.StatusRow>.Factory.StartNew(() =>
                                       {
                                           try
                                           {
                                               statusrow = Program.runComparison(row, od, statusrow);
                                           }
                                           finally
                                           {
                                               concurrency.Release();
                                           }
                                           return statusrow;
                                       });
                                       
                                       listOfTasks.Add(n);
                                       
                                       //statusrow =Program.runComparison(row, od, statusrow);
                                       //Utils.getStatusRow(overallStatusDT, statusrow);
                                   }
                                   
                                   Task.WaitAll(listOfTasks.ToArray());
                                   //foreach (Task<Utils.StatusRow> task in listOfTasks){
                                   //    task.Start();
                                   //}
                               }


                               foreach (Task<Utils.StatusRow> task in listOfTasks)
                               {
                                   Utils.StatusRow statusrow = task.Result;
                                   Utils.getStatusRow(overallStatusDT, statusrow);
                               }
                           }
                           catch (Exception e)
                           {
                               Console.WriteLine(e.Message);
                               Console.WriteLine(e.StackTrace);
                           }
                           finally
                           {
                               statusWorkbook.AddWorksheet(overallStatusDT, "Status");
                               statusWorkbook.SaveAs(od.FullName + "\\" + "_OverallStatus.xlsx");
                           }
                       }

                   });
        }
    }
}
