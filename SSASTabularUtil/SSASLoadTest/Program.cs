using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ADOTabular;
using ADOTabular.AdomdClientWrappers;
using System.Data;
namespace SSASLoadTest
{
    class Program
    {
        private static void RunQuery(string server, string model, string query)
        {
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
                    srcConx = new ADOTabularConnection(server, AdomdType.AnalysisServices);
                    srcConx.ChangeDatabase(model);
                    srcDT = srcConx.ExecuteDaxQueryDataTable(query);
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
        static void Main(string[] args)
        {
            
            Console.WriteLine("Hello World!");
            Console.ReadKey();
        }
    }
}
