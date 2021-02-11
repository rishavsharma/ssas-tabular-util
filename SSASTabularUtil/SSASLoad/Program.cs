using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using ADOTabular;
using ADOTabular.AdomdClientWrappers;
using System.Data;
using CommandLine;
namespace SSASLoadTest
{
    class Program
    {
        private static bool keepRunning = true;
        private static void RunQuery(string server, string model, string query)
        {
            DateTime srcStartTime = DateTime.Now, srcEndTime = DateTime.Now;
            ADOTabularConnection srcConx = null;
            String srcErrorMsg = "", status = null;
            DataTable srcDT = new DataTable();
            int recordCount = 0;
            int SPID = 0;
            try
            {

                try
                {
                    status = "PASS";
                    srcConx = new ADOTabularConnection(server,ADOTabular.Enums.AdomdType.AnalysisServices);
                    srcConx.ChangeDatabase(model); 
                    srcDT = srcConx.ExecuteDaxQueryDataTable(query); 
                    recordCount = srcDT.Rows.Count;
                    SPID = srcConx.SPID;
                    srcEndTime = DateTime.Now;
                    Console.WriteLine("Session:" + srcConx.SPID + "["+ (srcEndTime - srcStartTime).TotalSeconds.ToString()+"]");
                }
                catch (Exception ee)
                {
                    Console.WriteLine(ee.Message);
                    Console.WriteLine(ee.StackTrace);
                    srcErrorMsg = ee.Message;
                }
                finally
                {
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
            }
        }
        public class Options
        {
            [Option('i', "input", Required = true, HelpText = "Input Excel")]
            public String inExcel { get; set; }
            [Option('t', "threads", Required = true, HelpText = "Number of Threads")]
            public int numThreads { get; set; }
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
                       Console.WriteLine("Set threads done: "+ThreadPool.SetMaxThreads(o.numThreads, o.numThreads));
                       Utils utils = new Utils();
                       DataTable excel = utils.ImportExceltoDatatable(o.inExcel, "Queries");
                       DateTime srcStartTime = DateTime.Now, srcEndTime = DateTime.Now;
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
                           string srcSSAS = "Data Source = " +  row["SERVER"].ToString();
                           string srcSSASModel = row["MODEL"].ToString();
                           string srcQuery = row["DAX"].ToString();
                           ThreadPool.QueueUserWorkItem(state => Program.RunQuery(srcSSAS, srcSSASModel, srcQuery));
                       }
                       srcEndTime = DateTime.Now;
                       
                       Console.ReadLine();
                       Console.WriteLine("Time Taken:" + (srcEndTime - srcStartTime).TotalSeconds.ToString());
                   });
            
        }
    }
}
