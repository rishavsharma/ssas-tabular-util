using ADOTabular;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSASLoad
{
    class QueryRunner
    {
        static void Main(string[] args)
        {
            ADOTabularConnection srcConx = new ADOTabularConnection("Data Source=ntedatsgv77", ADOTabular.Enums.AdomdType.AnalysisServices);
            srcConx.ChangeDatabase("R12 FRR Summary In-Memory");
            System.Data.DataTable dataTable = srcConx.ExecuteDaxQueryDataTable("EVALUATE 'D Time'");
            Console.WriteLine(dataTable.Rows.Count);
            Console.ReadLine();

        }
    }
}
