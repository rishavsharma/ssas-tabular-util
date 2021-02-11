using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.AnalysisServices.Tabular;
using TOM = Microsoft.AnalysisServices.Tabular;
using TD = TabularEditor.TOMWrapper.Utils;
using System.IO;
using System.Reflection;

namespace SSASTabular
{
    class Program
    {
        static void Main(string[] args)
        {
            
            if (args.Length == 0)
            {
                Console.WriteLine(@"
List the databases on server
SSASTabular -S server -L [ALL | dbName] [-F exportFolder]
-L ALL | dbName (list the database)
-F = FolderName (where databases exported to folder)

Deploy Models including roles
SSASTabular -S server -B modelFile -D dbName
            ");
                Console.ReadLine();
                System.Environment.Exit(0);
            }
            List<string> upperArgList = args.Select(arg => arg.ToUpper()).ToList();
            int sIndex = upperArgList.IndexOf("-S");
            Server ssasServer = new Server();
            String server = ".";
            if (sIndex != -1 && args.Length > 1)
            {
                server = args[sIndex + 1];                
            }
            ssasServer.Connect(server);
            int lIndex = upperArgList.IndexOf("-L");
            if (lIndex != -1)
            {
                System.Collections.IList dbList = ssasServer.Databases.Cast<Database>().OrderBy(db => db.Name).ToList();
                foreach (Database item in dbList)
                {
                    if (args[lIndex+1].Equals("ALL") || args[lIndex + 1].Equals(item.Name)) {
                        Console.WriteLine(item.Name);
                        int fIndex = upperArgList.IndexOf("-F");
                        if (-1 != fIndex)
                        {
                            String folder = args[fIndex + 1];
                            String json = TOM.JsonSerializer.SerializeDatabase(item);
                            System.IO.FileInfo file = new System.IO.FileInfo(folder + "\\" + item.Name + ".bim");
                            file.Directory.Create();
                            System.IO.File.WriteAllText(file.FullName, json);
                        }
                    }
                }
                Console.ReadLine();
            }

            int dIndex = upperArgList.IndexOf("-D");
            int bIndex = upperArgList.IndexOf("-B");
            if (dIndex != -1 && bIndex !=-1)
            {
                Console.WriteLine("Deploying:"+ args[bIndex + 1]);
                string contents = File.ReadAllText(args[bIndex+1]);
                Database db = TOM.JsonSerializer.DeserializeDatabase(contents);
                TD.TabularDeployer.Deploy(db, server, args[dIndex+1]);
                Console.WriteLine("Deploying Completed..");
                Console.ReadLine();
            }
            
            System.Environment.Exit(0);            
            
        }
    }
}
