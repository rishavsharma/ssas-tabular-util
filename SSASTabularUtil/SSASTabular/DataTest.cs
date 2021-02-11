using ADOTabular;
using ADOTabular.AdomdClientWrappers;
using ClosedXML.Excel;
using CommandLine;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using Teradata.Client.Provider;

namespace SSASTabular
{
    class DataTest
    {
        public class Options
        {
            [Option('t', "teradata", Required = true, HelpText = "Teradata Server")]
            public String Server { get; set; }
            [Option('s', "ssas", Required = true, HelpText = "SSAS Server")]
            public String SSAS { get; set; }
            [Option('i', "input", Required = true, HelpText = "Input Excel")]
            public String inExcel { get; set; }
            [Option('o', "out", Required = true, HelpText = "Out Directory")]
            public String outDir { get; set; }
        }
        private const int MAX_ROWS = 10000;
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       Utils utils = new Utils();
                       DataTable excel = utils.ImportExceltoDatatable(o.inExcel, "Queries");

                       DirectoryInfo od = new DirectoryInfo(o.outDir);
                       od.Create();

                       DataTable dtSummary = new DataTable();
                       dtSummary.Columns.Add(new DataColumn("SSASModel", typeof(String)));
                       dtSummary.Columns.Add(new DataColumn("TableName", typeof(String)));                       
                       dtSummary.Columns.Add(new DataColumn("TestResults", typeof(String)));
                       dtSummary.Columns.Add(new DataColumn("TDQueryStartTime", typeof(DateTime)));
                       dtSummary.Columns.Add(new DataColumn("TDQueryEndTime", typeof(DateTime)));
                       dtSummary.Columns.Add(new DataColumn("TDQueryExecutionTime(minutes)", typeof(System.Double)));

                       dtSummary.Columns.Add(new DataColumn("SSASQueryStartTime", typeof(DateTime)));
                       dtSummary.Columns.Add(new DataColumn("SSASQueryEndTime", typeof(DateTime)));
                       dtSummary.Columns.Add(new DataColumn("SSASQueryExecutionTime(minutes)", typeof(System.Double)));
                       dtSummary.Columns.Add(new DataColumn("Exception", typeof(String)));
                       dtSummary.Columns.Add(new DataColumn("TDQuery", typeof(String)));
                       dtSummary.Columns.Add(new DataColumn("DAX", typeof(String)));

                       foreach (DataRow row in excel.Rows)
                       {
                           
                           string ssasModel = row["SSAS_MODEL"].ToString();
                           string tableName = row["TABLE_NAME"].ToString();
                           string tdQuery = row["TD_QUERY"].ToString();
                           string daxQuery = row["DAX"].ToString();
                           DateTime tdStartTime = DateTime.Now, tdEndTime= DateTime.Now, SSASStartTime= DateTime.Now, SSASEndTime = DateTime.Now;

                           if (tableName.ToString().Equals(""))
                           {
                               continue;
                           }
                           DataRow drSummary = dtSummary.NewRow();
                           drSummary["SSASModel"] = ssasModel;
                           drSummary["TableName"] = tableName;
                           drSummary["TDQuery"] = tdQuery;
                           drSummary["DAX"] = daxQuery;
                           drSummary["TestResults"] = "Failed";
                           Console.WriteLine("---------------------------------------------------------------");
                           Console.WriteLine(ssasModel + " -> " + tableName);


                           using (var workbook = new XLWorkbook())
                           {
                               try
                               {
                                   Console.WriteLine(tdQuery);
                                   TdConnection cn = new TdConnection(o.Server);
                                   cn.Open();
                                   TdCommand cmd = null;                                   
                                   TdDataAdapter adapter = null;
                                   DataTable dtt = new DataTable();
                                   tdStartTime = DateTime.Now;
                                   try
                                   {
                                       cmd = new TdCommand(tdQuery, cn);
                                       cmd.CommandTimeout = 1200;
                                       adapter = new TdDataAdapter(cmd);
                                       tdStartTime = DateTime.Now;
                                       adapter.Fill(dtt);
                                       
                                       tdEndTime = DateTime.Now;
                                       workbook.Worksheets.Add(dtt, "TD");
                                       cmd.Connection.Close();
                                       cmd.Dispose();
                                       adapter.Dispose();
                                   }
                                   catch(Exception e)
                                   {
                                       tdEndTime = DateTime.Now;
                                       cmd.Connection.Close();
                                       cmd.Dispose();
                                       adapter.Dispose();
                                       throw e;
                                   }

                                   drSummary["TDQueryStartTime"] = tdStartTime;
                                   drSummary["TDQueryEndTime"] = tdEndTime;
                                   drSummary["TDQueryExecutionTime(minutes)"] = (double)tdEndTime.Subtract(tdStartTime).Seconds / 60;


                                   //ADOTabularConnection conx = new ADOTabularConnection(@"Data Source = .; Catalog = AW;", AdomdType.AnalysisServices);
                                   System.Data.DataTable dt = new DataTable();
                                   ADOTabularConnection conx=null;
                                   try
                                   {
                                       conx = new ADOTabularConnection(o.SSAS, ADOTabular.Enums.AdomdType.AnalysisServices);
                                       if (!ssasModel.Trim().Equals(""))
                                       {
                                           conx.ChangeDatabase(ssasModel);
                                       }

                                       Console.WriteLine("--------------------DAX---------------------------");
                                       Console.WriteLine(daxQuery);



                                       SSASStartTime = DateTime.Now;
                                       drSummary["SSASQueryStartTime"] = SSASStartTime;
                                       dt = conx.ExecuteDaxQueryDataTable(daxQuery);
                                       SSASEndTime = DateTime.Now;
                                       conx.Close();
                                       conx.Dispose();
                                   }
                                   catch(Exception e)
                                   {
                                       SSASEndTime = DateTime.Now;
                                       conx.Close();
                                       conx.Dispose();
                                       throw e;
                                   }
                                   
                                   drSummary["SSASQueryEndTime"] = SSASEndTime;
                                   drSummary["SSASQueryExecutionTime(minutes)"] = (double)SSASEndTime.Subtract(SSASStartTime).Seconds / 60;


                                   int noOfRows = dt.Rows.Count;

                                   if (noOfRows > MAX_ROWS)
                                   {
                                       dt = dt.Rows.Cast<DataRow>().Take(MAX_ROWS).CopyToDataTable();

                                   }
                                   workbook.Worksheets.Add(dt, "SSAS");

                                   DataTable ret = utils.getDifferentRecords(dtt, dt);

                                   DataTable dtOutput = new DataTable();
                                   dtOutput.Columns.Add(new DataColumn("SSASModel", typeof(String)));
                                   dtOutput.Columns.Add(new DataColumn("TableName", typeof(String)));

                                   dtOutput.Columns.Add(new DataColumn("TDQuery", typeof(String)));
                                   dtOutput.Columns.Add(new DataColumn("DAX", typeof(String)));
                                   dtOutput.Columns.Add(new DataColumn("TestResults", typeof(String)));

                                   dtOutput.Columns.Add(new DataColumn("TDQueryStartTime", typeof(DateTime)));
                                   dtOutput.Columns.Add(new DataColumn("TDQueryEndTime", typeof(DateTime)));
                                   dtOutput.Columns.Add(new DataColumn("TDQueryExecutionTime(minutes)", typeof(System.Double)));

                                   dtOutput.Columns.Add(new DataColumn("SSASQueryStartTime", typeof(DateTime)));
                                   dtOutput.Columns.Add(new DataColumn("SSASQueryEndTime", typeof(DateTime)));
                                   dtOutput.Columns.Add(new DataColumn("SSASQueryExecutionTime(minutes)", typeof(System.Double)));

                                   DataRow dr = dtOutput.NewRow();

                                   dr["SSASModel"] = ssasModel;
                                   dr["TableName"] = tableName;
                                   dr["TDQuery"] = tdQuery;
                                   dr["DAX"] = daxQuery;

                                   dr["TDQueryStartTime"] = tdStartTime;
                                   dr["TDQueryEndTime"] = tdEndTime;
                                   dr["TDQueryExecutionTime(minutes)"] = (double) tdEndTime.Subtract(tdStartTime).Seconds / 60;
                                   
                                   dr["SSASQueryStartTime"] = SSASStartTime;
                                   dr["SSASQueryEndTime"] = SSASEndTime;
                                   dr["SSASQueryExecutionTime(minutes)"] = (double) SSASEndTime.Subtract(SSASStartTime).Seconds / 60;

                                   
                                   if (noOfRows == 0)
                                   {
                                       //tableName = tableName + "_NODATA";
                                       dr["TestResults"] = "No data";
                                       drSummary["TestResults"] = "No data";
                                   }
                                   else if (ret.Rows.Count > 0)
                                   {
                                       workbook.Worksheets.Add(ret, "ERRORS");
                                       //tableName = tableName + "_FAILED";
                                       dr["TestResults"] = "Failed";
                                       drSummary["TestResults"] = "Failed";
                                   }
                                   else
                                   {
                                       //tableName = tableName + "_PASSED";
                                       dr["TestResults"] = "Passed";
                                       drSummary["TestResults"] = "Passed";
                                   }
                                   dtOutput.Rows.Add(dr);

                                   workbook.Worksheets.Add(dtOutput, "TestOutput");
                                   
                               }
                               catch (Exception e)
                               {
                                   Console.WriteLine(e.Message);
                                   Console.WriteLine(e.StackTrace);
                                   tableName = tableName + "_Exception";
                                   DataTable ex = new DataTable("Exception");
                                   ex.Columns.AddRange(new DataColumn[2] { new DataColumn("ERROR1"), new DataColumn("ERROR2") });
                                   drSummary["Exception"] = e.Message;
                                   ex.Rows.Add("Failed", e.Message);
                                   ex.Rows.Add("", e.StackTrace);
                                   workbook.Worksheets.Add(ex, "Exception");
                               }
                               finally
                               {
                                   dtSummary.Rows.Add(drSummary);
                                   workbook.SaveAs(od.FullName + "\\" + tableName + ".xlsx");
                                   XLWorkbook summaryworkbook = new XLWorkbook();
                                   summaryworkbook.Worksheets.Add(dtSummary, "TestSummary");
                                   summaryworkbook.SaveAs(od.FullName + "\\TestSummary.xlsx");
                               }

                           }

                       }

                       
                   });



        }
    }
}
