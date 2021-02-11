using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSASTabular
{
    class Utils
    {
        public DataTable ImportExceltoDatatable(string filePath, string sheetName)
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

        public string getNullReplace(string val)
        {
            if (string.IsNullOrEmpty(val))
            {
                return "0";
            }
            return val;
        }
        public DataTable getDifferentRecords(DataTable tbl1, DataTable tbl2)
        {
            //Create Empty Table 
            
            DataTable ResultDataTable = new DataTable("ResultDataTable");
            try
            {
                if (tbl1.Rows.Count != tbl2.Rows.Count || tbl1.Columns.Count != tbl2.Columns.Count)
                {
                    string msg = tbl1.Rows.Count + " TD-[Rows]-SSAS " + tbl2.Rows.Count + " - " + tbl1.Columns.Count + " TD-[Count]-SSAS " + tbl2.Columns.Count;
                    Console.WriteLine(msg);
                    ResultDataTable.Columns.AddRange(new DataColumn[2] { new DataColumn("ERROR1"), new DataColumn("ERROR2") });
                    ResultDataTable.Rows.Add("Count Failed", "" );
                    ResultDataTable.Rows.Add( msg, "" );
                    return ResultDataTable;
                }
                Console.WriteLine("First Table");
                for (int i = 0; i < tbl1.Columns.Count; i++)
                {
                    Console.WriteLine(tbl1.Columns[i].ColumnName + "-" + tbl1.Columns[i].DataType);
                    ResultDataTable.Columns.Add(tbl1.Columns[i].ColumnName);
                }
                Console.WriteLine("Second Table");
                for (int i = 0; i < tbl2.Columns.Count; i++)
                {
                    Console.WriteLine(tbl2.Columns[i].ColumnName + "-" + tbl2.Columns[i].DataType);
                }



                int ERROR_LIMIT = 100;
                int errorCount = 0;
                for (int i = 0; i < tbl1.Rows.Count; i++)
                {
                    if(errorCount > ERROR_LIMIT)
                    {
                        break;
                    }
                    for (int c = 0; c < tbl1.Columns.Count; c++)
                    {
                        //Console.WriteLine(tbl1.Rows[i][c] + " [compare] " + tbl2.Rows[i][c]);
                        if (tbl1.Columns[c].DataType == typeof(System.Double) || tbl1.Columns[c].DataType == typeof(System.Decimal))
                        {
                            double x, y;
                            x = Convert.ToDouble(getNullReplace(tbl1.Rows[i][c].ToString()));
                            y = Convert.ToDouble(getNullReplace(tbl2.Rows[i][c].ToString()));
                            if (!Equals(Math.Round(x,3), Math.Round(y,3)))
                            {
                                ++errorCount;
                                String msg = tbl1.Rows[i][c].ToString() + " [TD-[not Equal]-SSAS] " + tbl2.Rows[i][c].ToString();
                                //Console.WriteLine(msg);
                                ResultDataTable.Rows.Add();
                                ResultDataTable.Rows[errorCount][c] = msg;

                            }
                        }
                        else
                        {
                            if (!Equals(getNullReplace(tbl1.Rows[i][c].ToString()), getNullReplace(tbl2.Rows[i][c].ToString())))
                            {
                                ++errorCount;
                                String msg = tbl1.Rows[i][c].ToString() + " [TD-[not Equal]-SSAS] " + tbl2.Rows[i][c].ToString();
                                //Console.WriteLine(msg);
                                ResultDataTable.Rows.Add();
                                ResultDataTable.Rows[errorCount][c] = msg;

                            }
                        }
                    }
                }
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return ResultDataTable;
        }
    }
}
