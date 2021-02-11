using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using Microsoft.AnalysisServices.Tabular;
using TOM = Microsoft.AnalysisServices.Tabular;
using TD = TabularEditor.TOMWrapper.Utils;
namespace SSASDeploy
{
    class Deploy
    {
        public class Options
        {
            [Option('s', "ssas", Required = true, HelpText = "SSAS Server")]
            public String ssasServer { get; set; }
            [Option('a', "action", Required = true, HelpText = "Action: deploy|export|list|tmsl|xmla")]
            public String action { get; set; }
            [Option('f', "input", Required = false, HelpText = "Input list file for database")]
            public String infile { get; set; }
            [Option('o', "out", Required = false, HelpText = "Out Directory")]
            public String outDir { get; set; }
            [Option('c', "conn", Required = false, HelpText = "Deploy Connection")]
            public bool deployConnection { get; set; }
            [Option('r', "roles", Required = false, HelpText = "Deploy Roles")]
            public bool deployRoles { get; set; }
            [Option('m', "member", Required = false, HelpText = "Deploy Roles Members")]
            public bool deployMembers { get; set; }
            [Option('t', "partition", Required = false, HelpText = "Deploy Partition")]
            public bool deployPartition { get; set; }
            [Option('x', "script", Required = false, HelpText = "Use with export to generate xmla script into out dir")]
            public bool script { get; set; }
            [Option('p', "recal", Required = false, HelpText = "Recal the models")]
            public bool recal { get; set; }
            [Option('n', "name", Required = false, HelpText = "name of the Model")]
            public String nameSSASModel { get; set; }
        }
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       try
                       {
                           Server ssasServer = new Server();
                           ssasServer.Connect(o.ssasServer);
                           if (o.action.Equals("deploy"))
                           {
                               TD.DeploymentOptions dopt = TD.DeploymentOptions.Default;
                               if (o.deployConnection)
                               {
                                   dopt.DeployConnections = true;
                               }
                               if (o.deployRoles)
                               {
                                   dopt.DeployRoles = true;
                               }
                               if (o.deployMembers)
                               {
                                   dopt.DeployRoleMembers = true;
                               }

                               if (o.deployPartition)
                               {
                                   dopt.DeployPartitions = true;
                               }

                               if (o.infile.EndsWith("bim", StringComparison.CurrentCultureIgnoreCase))
                               {
                                   string contents = File.ReadAllText(o.infile);
                                   Console.WriteLine("Deploying:" + o.infile);
                                   Database db = TOM.JsonSerializer.DeserializeDatabase(contents);
                                   TD.TabularDeployer.Deploy(db, ssasServer, o.nameSSASModel, dopt);
                                   Console.WriteLine("Deploying Completed..");
                                   if (o.recal)
                                   {
                                       Console.WriteLine("Recal:" + db.Name);
                                       var result = ssasServer.Execute("{\"refresh\": {\"type\": \"calculate\",\"objects\": [{\"database\": \"" + o.nameSSASModel + "\"}]}}");
                                       if (result.ContainsErrors)
                                       {
                                           throw new Exception(string.Join("\n", result.Cast<Microsoft.AnalysisServices.XmlaResult>().SelectMany(r => r.Messages.Cast<Microsoft.AnalysisServices.XmlaMessage>().Select(m => m.Description)).ToArray()));
                                       }
                                       ssasServer.Refresh();
                                       Console.WriteLine("Recal Done:" + db.Name);
                                   }
                               }
                               else
                               {
                                   foreach (string line in File.ReadLines(o.infile, Encoding.UTF8))
                                   {
                                       try
                                       {
                                           string[] arg = line.Split(',');
                                           Console.WriteLine("Deploying:" + arg[0]);
                                           string contents = File.ReadAllText(arg[1]);
                                           Database db = TOM.JsonSerializer.DeserializeDatabase(contents);
                                           TD.TabularDeployer.Deploy(db, ssasServer, arg[0], dopt);
                                           Console.WriteLine("Deploying Completed..");
                                           if (o.recal)
                                           {
                                               Console.WriteLine("Recal:" + db.Name);
                                               var result = ssasServer.Execute("{\"refresh\": {\"type\": \"calculate\",\"objects\": [{\"database\": \""+ arg[0] + "\"}]}}");
                                               if (result.ContainsErrors)
                                               {
                                                   throw new Exception(string.Join("\n", result.Cast<Microsoft.AnalysisServices.XmlaResult>().SelectMany(r => r.Messages.Cast<Microsoft.AnalysisServices.XmlaMessage>().Select(m => m.Description)).ToArray()));
                                               }
                                               ssasServer.Refresh();
                                               Console.WriteLine("Recal Done:" + db.Name);
                                           }

                                       }
                                       catch (Exception e)
                                       {
                                           Console.WriteLine(e.Message);
                                       }
                                   }
                               }

                           }else if (o.action.Equals("export"))
                           {
                               string outdir = o.outDir;
                               if(outdir == null)
                               {
                                   outdir = "out";
                               }
                               foreach (string line in File.ReadLines(o.infile, Encoding.UTF8))
                               {
                                   try
                                   {
                                       Database db = ssasServer.Databases[line];
                                       if (o.script)
                                       {
                                           Console.WriteLine("Exporting script:" + db.Name);
                                           System.IO.FileInfo file = new System.IO.FileInfo(outdir + "\\" + db.Name + ".xmla");
                                           file.Directory.Create();
                                           var rawTmsl = TOM.JsonScripter.ScriptCreateOrReplace(db, false);
                                           System.IO.File.WriteAllText(file.FullName, rawTmsl);
                                           Console.WriteLine("Exporting finished");

                                       }
                                       else
                                       {
                                           Console.WriteLine("Exporting:" + line);                                           
                                           String json = TOM.JsonSerializer.SerializeDatabase(db);
                                           System.IO.FileInfo file = new System.IO.FileInfo(outdir + "\\" + db.Name + ".bim");
                                           file.Directory.Create();
                                           System.IO.File.WriteAllText(file.FullName, json);
                                           Console.WriteLine("Exporting finished");
                                       }
                                   }
                                   catch(Exception e)
                                   {
                                       Console.WriteLine(e.Message);
                                   }
                               }
                           }
                           else if (o.action.Equals("list"))
                           {
                               System.Collections.IList dbList = ssasServer.Databases.Cast<Database>().OrderBy(db => db.Name).ToList();
                               foreach (Database item in dbList)
                               {
                                   Console.WriteLine(item.Name);
                               }
                           }
                           else if (o.action.Equals("tmsl"))
                           {
                               foreach (string line in File.ReadLines(o.infile, Encoding.UTF8))
                               {
                                   try
                                   {
                                       Console.WriteLine("Executing for:" + line);
                                       string tmsl = File.ReadAllText(line, Encoding.UTF8);
                                       var result = ssasServer.Execute(tmsl);
                                       if (result.ContainsErrors)
                                       {
                                           throw new Exception(string.Join("\n", result.Cast<Microsoft.AnalysisServices.XmlaResult>().SelectMany(r => r.Messages.Cast<Microsoft.AnalysisServices.XmlaMessage>().Select(m => m.Description)).ToArray()));
                                       }
                                       ssasServer.Refresh(); 
                                       Console.WriteLine("Executed Successfully");
                                   }catch(Exception e)
                                   {
                                       Console.WriteLine(e.Message);
                                   }
                               }
                           }
                           else if (o.action.Equals("xmla"))
                           {
                               foreach (string line in File.ReadLines(o.infile, Encoding.UTF8))
                               {
                                   Console.WriteLine("Executing for:" + line);
                                   var result = ssasServer.Execute(line);
                                   if (result.ContainsErrors)
                                   {
                                       throw new Exception(string.Join("\n", result.Cast<Microsoft.AnalysisServices.XmlaResult>().SelectMany(r => r.Messages.Cast<Microsoft.AnalysisServices.XmlaMessage>().Select(m => m.Description)).ToArray()));
                                   }
                                   ssasServer.Refresh();
                                   Console.WriteLine("Executed Successfully");
                               }
                           }
                       }
                       catch(Exception e)
                       {
                           Console.WriteLine(e.Message);
                       }
                   });
        }
    }
}
