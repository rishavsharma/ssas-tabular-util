using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAXRunner
{
    class Utils
    {
        public static int ERROR_ROWS = 20000;
        public class StatusRow
        {
            public String NAME { get; set; }
            public String STATUS { get; set; }
            public DateTime SRC_SSASQueryStartTime { get; set; }
            public DateTime SRC_SSASQueryEndTime { get; set; }
            public System.Double SRC_SSASQueryExecutionTime { get; set; }
            public DateTime TGT_SSASQueryStartTime { get; set; }
            public DateTime TGT_SSASQueryEndTime { get; set; }
            public System.Double TGT_SSASQueryExecutionTime { get; set; }
            public String SRC_Exception { get; set; }
            public String TGT_Exception { get; set; }
            public String SRC_DAX { get; set; }
            public String TGT_DAX { get; set; }
        }

        public static DataTable ImportExceltoDatatable(string filePath, string sheetName)
        {
            DataTable dt = new DataTable();
            try
            {
                // Open the Excel file using ClosedXML.
                // Keep in mind the Excel file cannot be open when trying to read it
                using (XLWorkbook workBook = new XLWorkbook(filePath))
                {
                    //Read the first Sheet from Excel file.
                    IXLWorksheet workSheet = workBook.Worksheet(sheetName);

                    //Loop through the Worksheet rows.
                    bool firstRow = true;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        
                        //Use the first row to add columns to DataTable.
                        if (firstRow)
                        {
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }
                            firstRow = false;
                        }
                        else
                        {
                            //Add rows to DataTable.
                            dt.Rows.Add();
                            int i = 0;

                            foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                            {
                                if (cell == null)
                                {
                                    continue;
                                }
                                if (cell.IsEmpty())
                                {
                                    continue;

                                }
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                i++;
                            }
                        }
                    }
                }
            }catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return dt;
            
        }

        public static string getNullReplace(string val)
        {
            if (string.IsNullOrEmpty(val))
            {
                return "0";
            }
            return val;
        }

        public static DataTable getDifferentRecords(DataTable tbl1, DataTable tbl2)
        {
            //Create Empty Table 
            
            DataTable ResultDataTable = new DataTable("ResultDataTable");
            try
            {
                if (tbl1.Rows.Count != tbl2.Rows.Count || tbl1.Columns.Count != tbl2.Columns.Count)
                {
                    string msg = tbl1.Rows.Count + " SRC-[Rows]-TGT " + tbl2.Rows.Count + " - " + tbl1.Columns.Count + " SRC-[Count]-TGT " + tbl2.Columns.Count;
                    //Console.WriteLine(msg);
                    ResultDataTable.Columns.AddRange(new DataColumn[2] { new DataColumn("ERROR1"), new DataColumn("ERROR2") });
                    ResultDataTable.Rows.Add("Count Failed", "" );
                    ResultDataTable.Rows.Add( msg, "" );
                    return ResultDataTable;
                }
                //Console.WriteLine("First Table");
                for (int i = 0; i < tbl1.Columns.Count; i++)
                {
                    //Console.WriteLine(tbl1.Columns[i].ColumnName + "-" + tbl1.Columns[i].DataType);
                    ResultDataTable.Columns.Add(tbl1.Columns[i].ColumnName);
                }
                //Console.WriteLine("Second Table");
                //for (int i = 0; i < tbl2.Columns.Count; i++)
                //{
                //    Console.WriteLine(tbl2.Columns[i].ColumnName + "-" + tbl2.Columns[i].DataType);
                //}



                int ERROR_LIMIT = ERROR_ROWS;
                int errorCount = 0;
                for (int i = 0; i < tbl1.Rows.Count; i++)
                {
                    //errorCount++;
                    ResultDataTable.Rows.Add();
                    if (errorCount > ERROR_LIMIT)
                    {
                        break;
                    }
                    for (int c = 0; c < tbl1.Columns.Count; c++)
                    {
                        //Console.WriteLine(tbl1.Rows[i][c] + " [compare] " + tbl2.Rows[i][c]);
                        if (tbl1.Columns[c].DataType == typeof(System.Double) || tbl1.Columns[c].DataType == typeof(System.Decimal))
                        {
                            double x, y;
                            x = Math.Round(Convert.ToDouble(getNullReplace(tbl1.Rows[i][c].ToString()))/100,0)*100;
                            y = Math.Round(Convert.ToDouble(getNullReplace(tbl2.Rows[i][c].ToString()))/100,0)*100;

                            if (x != y)
                            {
                                //++errorCount;
                                String msg = tbl1.Rows[i][c].ToString() + " [SRC-[not Equal]-TGT] " + tbl2.Rows[i][c].ToString();
                                //Console.WriteLine(msg);
                                //ResultDataTable.Rows.Add();
                                ResultDataTable.Rows[i][c] = msg;
                                errorCount++;
                            }
                        }
                        else
                        {
                            if (!Equals(getNullReplace(tbl1.Rows[i][c].ToString()), getNullReplace(tbl2.Rows[i][c].ToString())))
                            {
                                //++errorCount;
                                String msg = tbl1.Rows[i][c].ToString() + " [SRC-[not Equal]-TGT] " + tbl2.Rows[i][c].ToString();
                                //Console.WriteLine(msg);
                                //ResultDataTable.Rows.Add();
                                ResultDataTable.Rows[i][c] = msg;
                                errorCount++;
                            }
                        }
                    }
                }
                if (errorCount == 0)
                {
                    ResultDataTable.Rows.Clear();
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
            //if (ResultDataTable.Rows.Count > 0)
            //{
            //    bool isEmpty = true;
            //    for (int r = 0; r < ResultDataTable.Rows.Count; r++)
            //    {
            //        for (int i = 0; i < ResultDataTable.Columns.Count; i++)
            //        {
            //            string cv = ResultDataTable.Rows[r][i].ToString().Trim();
            //            if (!String.IsNullOrEmpty(cv))
            //            {
            //                isEmpty = false;
            //                break;
            //            }
            //        }
            //        if (!isEmpty)
            //        {
            //            break;
            //        }
            //    }

            //    if (isEmpty)
            //    {
            //        ResultDataTable.Rows.Clear();
            //    }
            //}
            
            return ResultDataTable;
        }

        public static DataTable getStatusDataTable()
        {
            DataTable dtSummary = new DataTable();
            dtSummary.Columns.Add(new DataColumn("NAME", typeof(String)));
            dtSummary.Columns.Add(new DataColumn("STATUS", typeof(String)));
            dtSummary.Columns.Add(new DataColumn("SRC_SSASQueryStartTime", typeof(DateTime)));
            dtSummary.Columns.Add(new DataColumn("SRC_SSASQueryEndTime", typeof(DateTime)));
            dtSummary.Columns.Add(new DataColumn("SRC_SSASQueryExecutionTime", typeof(System.Double)));
            dtSummary.Columns.Add(new DataColumn("TGT_SSASQueryStartTime", typeof(DateTime)));
            dtSummary.Columns.Add(new DataColumn("TGT_SSASQueryEndTime", typeof(DateTime)));
            dtSummary.Columns.Add(new DataColumn("TGT_SSASQueryExecutionTime", typeof(System.Double)));
            dtSummary.Columns.Add(new DataColumn("SRC_Exception", typeof(String)));
            dtSummary.Columns.Add(new DataColumn("TGT_Exception", typeof(String)));
            dtSummary.Columns.Add(new DataColumn("SRC_DAX", typeof(String)));
            dtSummary.Columns.Add(new DataColumn("TGT_DAX", typeof(String)));
            return dtSummary;
        }

        public static DataRow getStatusRow(DataTable dt, String NAME, String STATUS, DateTime SRC_SSASQueryStartTime, 
            DateTime SRC_SSASQueryEndTime, System.Double SRC_SSASQueryExecutionTime, DateTime TGT_SSASQueryStartTime, 
            DateTime TGT_SSASQueryEndTime, System.Double TGT_SSASQueryExecutionTime, String SRC_Exception, String TGT_Exception, String SRC_DAX, String TGT_DAX)
        {
            DataRow dr = dt.NewRow();
            dr["NAME"] = NAME;
            dr["STATUS"] = STATUS;
            dr["SRC_SSASQueryStartTime"] = SRC_SSASQueryStartTime;
            dr["SRC_SSASQueryEndTime"] = SRC_SSASQueryEndTime;
            dr["SRC_SSASQueryExecutionTime"] = SRC_SSASQueryExecutionTime;
            dr["TGT_SSASQueryStartTime"] = TGT_SSASQueryStartTime;
            dr["TGT_SSASQueryEndTime"] = TGT_SSASQueryEndTime;
            dr["TGT_SSASQueryExecutionTime"] = TGT_SSASQueryExecutionTime;
            dr["SRC_Exception"] = SRC_Exception;
            dr["TGT_Exception"] = TGT_Exception;
            dr["SRC_DAX"] = SRC_DAX;
            dr["TGT_DAX"] = TGT_DAX;
            dt.Rows.Add(dr);
            return dr;
        }

        public static DataRow getStatusRow(DataTable dt, Utils.StatusRow row)
        {
            DataRow dr = dt.NewRow();
            dr["NAME"] = row.NAME;
            dr["STATUS"] = row.STATUS;
            dr["SRC_SSASQueryStartTime"] = row.SRC_SSASQueryStartTime;
            dr["SRC_SSASQueryEndTime"] = row.SRC_SSASQueryEndTime;
            dr["SRC_SSASQueryExecutionTime"] = row.SRC_SSASQueryExecutionTime;
            dr["TGT_SSASQueryStartTime"] = row.TGT_SSASQueryStartTime;
            dr["TGT_SSASQueryEndTime"] = row.TGT_SSASQueryEndTime;
            dr["TGT_SSASQueryExecutionTime"] = row.TGT_SSASQueryExecutionTime;
            dr["SRC_Exception"] = row.SRC_Exception;
            dr["TGT_Exception"] = row.TGT_Exception;
            dr["SRC_DAX"] = row.SRC_DAX;
            dr["TGT_DAX"] = row.TGT_DAX;
            dt.Rows.Add(dr);
            return dr;
        }
    }
}
