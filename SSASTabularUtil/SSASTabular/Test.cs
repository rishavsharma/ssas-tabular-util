using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSASTabular
{
    class Test
    {
        static void Main(string[] args)
        {
            using (var workbook = new XLWorkbook())
            {
                DataTable dt = new DataTable();
                dt.Columns.AddRange(new DataColumn[3] { new DataColumn("Item"), new DataColumn("Price"), new DataColumn("Total") });
                dt.Rows.Add("Shirt", 200, null);
                dt.Rows.Add("Football", 30, null);
                dt.Rows.Add("Bat", 22.50, 0);
                dt.Rows.Add("Ring", 25, null);
                dt.Rows.Add("Band", 77, 0);
                dt.Rows.Add("Glass", 57, 0);
                workbook.Worksheets.Add(dt, "TD");
                DataTable dt1 = new DataTable();
                dt1.Columns.AddRange(new DataColumn[3] { new DataColumn("Item"), new DataColumn("Price"), new DataColumn("Total") });
                dt1.Rows.Add("Shirt", 200, null);
                dt1.Rows.Add("Football", 30, 0);
                dt1.Rows.Add("Bat", 22.5, 0);
                dt1.Rows.Add("Ring", 25, null);
                dt1.Rows.Add("Band", 77, 0);
                dt1.Rows.Add("Glasss", 57, 0);
                dt1.Rows.Add("Random", 57, 0);
                

                if (dt1.Rows.Count > 6)
                {
                    dt1 = dt1.Rows.Cast<DataRow>().Take(6).CopyToDataTable();

                }
                workbook.Worksheets.Add(dt1, "SSAS");
                Utils u = new Utils();
                DataTable dt2 = u.getDifferentRecords(dt, dt1);
                if (dt.Rows.Count > 0) { workbook.Worksheets.Add(dt2, "Errors"); }
                
                workbook.SaveAs("Test.xlsx");
            }
        }
    }
}
