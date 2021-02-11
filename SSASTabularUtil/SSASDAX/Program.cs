using System;
using System.Data;
using System.IO;
using ADOTabular;
using ADOTabular.AdomdClientWrappers;
using ClosedXML.Excel;
using CommandLine;

namespace SSASDAX
{
    class Program
    {
        private static bool keepRunning = true;
        public class Options
        {
            public String SSAS { get; set; }
            [Option('i', "input", Required = true, HelpText = "Input Excel")]
            public String inExcel { get; set; }
            [Option('o', "out", Required = true, HelpText = "Out Directory")]
            public String outDir { get; set; }
        }

        private static DataTable RunQuery(string server,string model,string query)
        {
            ADOTabularConnection srcConx = null;
            String srcErrorMsg = "", status = "PASS";
            DataTable srcDT = new DataTable();
            DataTable ret = new DataTable();
            try
            {
               
                try
                {
                    status = "PASS";
                    srcConx = new ADOTabularConnection(server, ADOTabular.Enums.AdomdType.AnalysisServices);
                    srcConx.ChangeDatabase(model);
                    srcDT = srcConx.ExecuteDaxQueryDataTable(query);
                    //recordCount = srcDT.Rows.Count;
                    //SPID = srcConx.SPID;
                    //Console.WriteLine("Session:" + srcConx.SPID);
                }
                catch (Exception ee)
                {
                    Console.WriteLine(ee.Message);
                    Console.WriteLine(ee.StackTrace);
                    srcErrorMsg = ee.Message;
                }
                finally
                {
                    //srcEndTime = DateTime.Now;
                    srcConx.Close();
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                status = "FAILED";
            }
            finally
            {
                srcConx.Close();
                //utils.getStatusRow(overallStatusDT, testName, status, srcStartTime, srcEndTime, (double)srcEndTime.Subtract(srcStartTime).TotalMilliseconds,
                //    recordCount, srcErrorMsg, srcQuery, SPID);                
            }
            return srcDT;

        }
        public static void runAllDAX(Options o)
        {
            Utils utils = new Utils();
            DataTable excel = utils.ImportExceltoDatatable(o.inExcel, "Queries");
            int recordCount = 0;
            DirectoryInfo od = new DirectoryInfo(o.outDir);
            od.Create();
            DataTable overallStatusDT = utils.getStatusDataTable();
            using (var statusWorkbook = new XLWorkbook())
            {
                try
                {
                    foreach (DataRow row in excel.Rows)
                    {
                        if (!keepRunning)
                        {
                            break;
                        }
                        string testName = row["NAME"].ToString();
                        if (testName.Trim().Equals(""))
                        {
                            continue;
                        }
                        string srcSSAS = "Data Source = " + row["SRC_SSAS"].ToString();
                        string srcSSASModel = row["SRC_SSAS_MODEL"].ToString();
                        string srcQuery = row["SRC_DAX"].ToString();
                        DateTime srcStartTime = DateTime.Now, srcEndTime = DateTime.Now;
                        String srcErrorMsg = "", status = "PASS";
                        int SPID = 0;
                        ADOTabularConnection srcConx = null;
                        try
                        {



                            Console.WriteLine("--------------------Src DAX (" + srcSSASModel + ")---------------------------");
                            Console.WriteLine(srcQuery);
                            DataTable srcDT = new DataTable();
                            DataTable ret = new DataTable();
                            try
                            {
                                status = "PASS";
                                srcConx = new ADOTabularConnection(srcSSAS, ADOTabular.Enums.AdomdType.AnalysisServices);
                                srcConx.ChangeDatabase(srcSSASModel);
                                srcStartTime = DateTime.Now;
                                srcDT = srcConx.ExecuteDaxQueryDataTable(srcQuery);
                                recordCount = srcDT.Rows.Count;
                                SPID = srcConx.SPID;
                                Console.WriteLine("Session:" + srcConx.SPID);
                            }
                            catch (Exception ee)
                            {
                                Console.WriteLine(ee.Message);
                                Console.WriteLine(ee.StackTrace);
                                srcErrorMsg = ee.Message;
                            }
                            finally
                            {
                                srcEndTime = DateTime.Now;
                                srcConx.Close();
                            }

                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                            Console.WriteLine(e.StackTrace);
                            status = "FAILED";
                        }
                        finally
                        {
                            srcConx.Close();
                            utils.getStatusRow(overallStatusDT, testName, status, srcStartTime, srcEndTime, (double)srcEndTime.Subtract(srcStartTime).TotalMilliseconds,
                                recordCount, srcErrorMsg, srcQuery, SPID);
                        }
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
        }
        static void Main(string[] args)
        {
            Console.CancelKeyPress += delegate (object sender, ConsoleCancelEventArgs e) {
                e.Cancel = true;
                Program.keepRunning = false;
            };
            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       Program.runAllDAX(o);

                   });
        }
    }
}
