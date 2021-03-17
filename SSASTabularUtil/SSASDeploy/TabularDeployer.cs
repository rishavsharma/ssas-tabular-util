extern alias json;
using Microsoft.AnalysisServices;
using Microsoft.AnalysisServices.Tabular;
using json.Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using TOM = Microsoft.AnalysisServices.Tabular;

namespace TabularEditor.TOMWrapper.Utils
{
    public class TabularDeployer
    {
        public static string GetTMSL(TOM.Database db, TOM.Server server, string targetDatabaseID, DeploymentOptions options, bool includeRestricted = false)
        {
            if (db == null) throw new ArgumentNullException("db");
            if (string.IsNullOrWhiteSpace(targetDatabaseID)) throw new ArgumentNullException("targetDatabaseID");
            if (options.DeployRoleMembers && !options.DeployRoles) throw new ArgumentException("Cannot deploy Role Members when Role deployment is disabled.");

            if (server.Databases.Contains(targetDatabaseID) && options.DeployMode == DeploymentMode.CreateDatabase) throw new ArgumentException("The specified database already exists.");

            if (!server.Databases.Contains(targetDatabaseID)) return DeployNewTMSL(db, targetDatabaseID, options, includeRestricted);
            else return DeployExistingTMSL(db, server, targetDatabaseID, options, includeRestricted);

            // TODO: Check if invalid CalculatedTableColumn perspectives/translations can give us any issues here
            // Should likely be handled similar to what we do in TabularModelHandler.SaveDB()
        }

       

        public static void Deploy(TOM.Database db, string targetConnectionString, string targetDatabaseName)
        {
            Deploy(db, targetConnectionString, targetDatabaseName, DeploymentOptions.Default);
        }

        public static void SaveModelMetadataBackup(string connectionString, string targetDatabaseID, string backupFilePath)
        {
            using (var s = new TOM.Server())
            {
                s.Connect(connectionString);
                if (s.Databases.Contains(targetDatabaseID))
                {
                    var db = s.Databases[targetDatabaseID];

                    var dbcontent = TOM.JsonSerializer.SerializeDatabase(db);
                }
                s.Disconnect();
            }
        }

        

        /// <summary>
        /// Deploys the specified database to the specified target server and database ID, using the specified options.
        /// Returns a list of DAX errors (if any) on objects inside the database, in case the deployment was succesful.
        /// </summary>
        /// <param name="db"></param>
        /// <param name="targetConnectionString"></param>
        /// <param name="targetDatabaseID"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        internal static void Deploy(TOM.Database db, string targetConnectionString, string targetDatabaseID, DeploymentOptions options)
        {
            if (string.IsNullOrWhiteSpace(targetConnectionString)) throw new ArgumentNullException("targetConnectionString");
            var s = new TOM.Server();
            s.Connect(targetConnectionString);

            var tmsl = GetTMSL(db, s, targetDatabaseID, options, true);
            var result = s.Execute(tmsl);

            if(result.ContainsErrors)
            {
                throw new Exception(string.Join("\n", result.Cast<XmlaResult>().SelectMany(r => r.Messages.Cast<XmlaMessage>().Select(m => m.Description)).ToArray()));
            }

            s.Refresh();
            var deployedDB = s.Databases[targetDatabaseID];
            
        }

        public static void Deploy(TOM.Database db, TOM.Server s, string targetDatabaseID, DeploymentOptions options)
        {
            var tmsl = GetTMSL(db, s, targetDatabaseID, options, true);
            var result = s.Execute(tmsl);

            if (result.ContainsErrors)
            {
                throw new Exception(string.Join("\n", result.Cast<XmlaResult>().SelectMany(r => r.Messages.Cast<XmlaMessage>().Select(m => m.Description)).ToArray()));
            }

            s.Refresh();
            var deployedDB = s.Databases[targetDatabaseID];

        }

        private static string GetName(TOM.NamedMetadataObject obj)
        {
            if (obj is TOM.Hierarchy) return string.Format("hierarchy {0}[{1}]", GetName((obj as TOM.Hierarchy).Table), obj.Name);
            if (obj is TOM.Measure) return string.Format("measure {0}[{1}]", GetName((obj as TOM.Measure).Table), obj.Name);
            if (obj is TOM.Column) return string.Format("column {0}[{1}]", GetName((obj as TOM.Column).Table), obj.Name);
            if (obj is TOM.Partition) return string.Format("partition '{0}' on table {1}", obj.Name, GetName((obj as TOM.Partition).Table));
            if (obj is TOM.Table) return string.Format("'{0}'", obj.Name);
            else return string.Format("{0} '{1}'", ((TOM.ObjectType)obj.ObjectType).ToString(), obj.Name);
        }

        /// <summary>
        /// This method transforms a JObject representing a Create TMSL script, so that the database is deployed
        /// using the proper ID and Name values. In addition, of the DeploymentOptions specify that roles should
        /// not be deployed, they are stripped from the TMSL script.
        /// </summary>
        private static JObject TransformCreateTmsl(JObject tmslJObj, string newDbId, DeploymentOptions options)
        {
            tmslJObj["create"]["database"]["id"] = newDbId;
            tmslJObj["create"]["database"]["name"] = newDbId;

            if (!options.DeployRoles)
            {
                // Remove roles if present
                var roles = tmslJObj.SelectToken("create.database.model.roles") as JArray;
                if (roles != null) roles.Clear();
            }

            return tmslJObj;
        }

        private static string DeployNewTMSL(TOM.Database db, string newDbId, DeploymentOptions options, bool includeRestricted)
        {
            var rawTmsl = TOM.JsonScripter.ScriptCreate(db, includeRestricted);

            var jTmsl = JObject.Parse(rawTmsl);

            return TransformCreateTmsl(jTmsl, newDbId, options).ToString();
        }

        /// <summary>
        /// This method transforms a JObject representing a CreateOrReplace TMSL script, so that the script points
        /// to the correct database to be overwritten, and that the correct ID and Name properties are set. In
        /// addition, the method will replace any Roles, RoleMembers, Data Sources and Partitions in the TMSL with
        /// the corresponding TMSL from the specified orgDb, depending on the provided DeploymentOptions.
        /// </summary>
        private static JObject TransformCreateOrReplaceTmsl(JObject tmslJObj, TOM.Database orgDb, DeploymentOptions options)
        {
            // Deployment target / existing database (note that TMSL uses the NAME of an existing database, not the ID, to identify the object)
            tmslJObj["createOrReplace"]["object"]["database"] = orgDb.Name;
            tmslJObj["createOrReplace"]["database"]["id"] = orgDb.ID;
            tmslJObj["createOrReplace"]["database"]["name"] = orgDb.Name;

            var model = tmslJObj.SelectToken("createOrReplace.database.model");

            var roles = model["roles"] as JArray;
            if (!options.DeployRoles)
            {
                // Remove roles if present and add original:
                roles = new JArray();
                model["roles"] = roles;
                foreach (var role in orgDb.Model.Roles) roles.Add(JObject.Parse(TOM.JsonSerializer.SerializeObject(role)));
            }
            else if (roles != null && !options.DeployRoleMembers)
            {
                foreach (var role in roles)
                {
                    var members = new JArray();
                    role["members"] = members;

                    // Remove members if present and add original:
                    var roleName = role["name"].Value<string>();
                    
                    if (orgDb.Model.Roles.Contains(roleName))
                    {
                        foreach (var member in orgDb.Model.Roles[roleName].Members)
                            members.Add(JObject.Parse(TOM.JsonSerializer.SerializeObject(member)));
                    }
                }
            }

            if (!options.DeployConnections)
            {
                // Remove dataSources if present
                var dataSources = new JArray();
                model["dataSources"] = dataSources;
                foreach (var ds in orgDb.Model.DataSources) dataSources.Add(JObject.Parse(TOM.JsonSerializer.SerializeObject(ds)));
            }

            if (!options.DeployPartitions)
            {
                var tables = tmslJObj.SelectToken("createOrReplace.database.model.tables") as JArray;
                foreach (var table in tables)
                {
                    var tableName = table["name"].Value<string>();
                    if (orgDb.Model.Tables.Contains(tableName))
                    {
                        var t = orgDb.Model.Tables[tableName];

                        var partitions = new JArray();
                        table["partitions"] = partitions;
                        foreach (var pt in t.Partitions) partitions.Add(JObject.Parse(TOM.JsonSerializer.SerializeObject(pt)));
                    }
                }
            }

            return tmslJObj;
        }

        private static string DeployExistingTMSL(TOM.Database db, TOM.Server server, string dbId, DeploymentOptions options, bool includeRestricted)
        {
            var rawTmsl = TOM.JsonScripter.ScriptCreateOrReplace(db, includeRestricted);

            var orgDb = server.Databases[dbId];

            var jTmsl = JObject.Parse(rawTmsl);

            return TransformCreateOrReplaceTmsl(jTmsl, orgDb, options).ToString();
        }
    }

    public class DeploymentResult
    {
        public readonly IReadOnlyList<string> Issues;
        public readonly IReadOnlyList<string> Warnings;
        public readonly IReadOnlyList<string> Unprocessed;
        public DeploymentResult(IEnumerable<string> issues, IEnumerable<string> warnings, IEnumerable<string> unprocessed)
        {
            Issues = issues.ToList();
            Warnings = warnings.ToList();
            Unprocessed = unprocessed.ToList();
        }
    }
    public class DeploymentOptions
    {
        public DeploymentMode DeployMode = DeploymentMode.CreateOrAlter;
        public bool DeployConnections = false;
        public bool DeployPartitions = false;
        public bool DeployRoles = false;
        public bool DeployRoleMembers = false;

        /// <summary>
        /// Default deployment. Does not overwrite connections, partitions or role members.
        /// </summary>
        public static DeploymentOptions Default = new DeploymentOptions();

        /// <summary>
        /// Full deployment.
        /// </summary>
        public static DeploymentOptions Full = new DeploymentOptions() { DeployConnections = true, DeployPartitions = true, DeployRoles = true, DeployRoleMembers = true };

        /// <summary>
        /// StructureOnly deployment. Does not overwrite roles or role members.
        /// </summary>
        public static DeploymentOptions StructureOnly = new DeploymentOptions() { DeployRoles = false };
    }

    public enum DeploymentMode
    {
        CreateDatabase,
        CreateOrAlter
    }
}
