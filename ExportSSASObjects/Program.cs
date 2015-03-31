using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using AMO = Microsoft.AnalysisServices;
using System.Data;
using System.Data.SqlClient;

namespace ExportSSASObjects
{

    static class ExtensionMethods
    {

        // http://stackoverflow.com/questions/13451085/exception-when-addwithvalue-parameter-is-null
        public static SqlParameter AddWithNullableValue(this SqlParameterCollection collection, string parameterName, object value)
        {
            if (value == null)
                return collection.AddWithValue(parameterName, DBNull.Value);
            else
                return collection.AddWithValue(parameterName, value);
        }

    }

    class Program
    {

        private class ResultRow
        {
            public string ServerName;
            public string DatabaseName;
            public string AnalyticalObject;
            public string ElementName;
            public string TransPath;
            public string TransKey;
            public string TransObject;
            public int? TransLanguage;
            public string TransValue;

            public ResultRow(string ServerName_, string DatabaseName_, string AnalyticalObject_, string ElementName_, string TransPath_, string TransKey_, string TransObject_ = null, int? TransLanguage_ = null, string TransValue_ = null)
            {
                ServerName = ServerName_;
                DatabaseName = DatabaseName_;
                AnalyticalObject = AnalyticalObject_;
                ElementName = ElementName_;
                TransPath = TransPath_;
                TransKey = TransKey_;
                TransObject = TransObject_;
                TransLanguage = TransLanguage_;
                TransValue = TransValue_;
            }

            public int Write (SqlConnection SqlConn, string TableName) 
            {
                int AffectedRows;

                using (SqlCommand Command = new SqlCommand())
                {
                    Command.Connection = SqlConn;
                    Command.CommandType = CommandType.Text;
                    Command.CommandText = "INSERT INTO " + @TableName + " (ServerName, DatabaseName, AnalyticalObject, ElementName, TransPath, TransKey, TransObject, TransLanguage, TransValue) VALUES (@ServerName, @DatabaseName, @AnalyticalObject, @ElementName, @TransPath, @TransKey, @TransObject, @TransLanguage, @TransValue) ";
                    Command.Parameters.AddWithNullableValue("@ServerName", ServerName);
                    Command.Parameters.AddWithNullableValue("@DatabaseName", DatabaseName);
                    Command.Parameters.AddWithNullableValue("@AnalyticalObject", AnalyticalObject);
                    Command.Parameters.AddWithNullableValue("@ElementName", ElementName);
                    Command.Parameters.AddWithNullableValue("@TransPath", TransPath);
                    Command.Parameters.AddWithNullableValue("@TransKey", TransKey);
                    Command.Parameters.AddWithNullableValue("@TransObject", TransObject);
                    Command.Parameters.AddWithNullableValue("@TransLanguage", TransLanguage);
                    Command.Parameters.AddWithNullableValue("@TransValue", TransValue);

                    try
                    {
                        AffectedRows = Command.ExecuteNonQuery();
                    }
                    catch (SqlException)
                    {
                        throw;
                    }
                }
                
                return AffectedRows;

            }
        }

        static void UnhandledExceptionTrapper(object sender, UnhandledExceptionEventArgs e)
        {

            Console.WriteLine(e.ExceptionObject.ToString());
            Console.WriteLine("");
            Console.WriteLine("Press Enter to continue");
            Console.ReadLine();
            Environment.Exit(1);
        }

        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;

            AMO.Database AnalysisDb;

            ArrayList results = new ArrayList();

            string AnalysisServerName;
            string AnalysisDatabaseName;
            string SqlServerName;
            string SqlDatabaseName;

            if (args.Length<3)
            {
                throw new Exception("Incorrect number of parameters! \r\n\r\nSyntax: ExportSSASObjects [AnalysisServerName] [AnalysisDatabaseName] [SqlServerName] [SqlDatabaseName]\r\n");
            }
            else
            {
                AnalysisServerName = args[0];
                AnalysisDatabaseName = args[1];
                SqlServerName = args[2];
                SqlDatabaseName = args[3];
            }

            string SqlConnectionString = "Integrated Security=true;Data Source=" + SqlServerName + ";Database=" + SqlDatabaseName;

            // try connection to sql server first
            using (SqlConnection SqlConn = new SqlConnection())
            {
                try
                {
                    SqlConn.ConnectionString = SqlConnectionString;
                    SqlConn.Open();
                    SqlConn.Close();
                }
                catch
                {
                    throw;
                }

            }        

            using (AMO.Server AnalysisServer = new AMO.Server())
            {

                // connect to analysis server
                try
                {
                    AnalysisServer.Connect(@"Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" + AnalysisDatabaseName + ";Data Source=" + AnalysisServerName);
                }
                catch
                {
                    throw;
                }

                // select the analysis database
                AnalysisDb = AnalysisServer.Databases.FindByName(AnalysisDatabaseName);
                if (AnalysisDb == null)
                    throw new Exception("Cannot connect to the database: " + AnalysisDatabaseName);

                // MAIN

                Console.Write("Reading analysis database " + AnalysisDatabaseName + " on server " + AnalysisServerName + " ... ");

                // dimensions
                foreach (AMO.Dimension dim in AnalysisDb.Dimensions)
                {

                    results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, dim.ID, dim.Name, "Dimension", "[" + dim.ID + "]"));
                    foreach (AMO.Translation t in dim.Translations)
                        results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, dim.ID, dim.Name, "Dimension", "[" + dim.ID + "]", "Caption", t.Language, t.Caption));

                    // dimension attributes
                    foreach (AMO.DimensionAttribute attr in dim.Attributes)
                    {
                        results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, dim.ID, attr.Name, "Dimension.Attribute", "[" + dim.ID + "].[" + attr.ID + "]"));
                        foreach (AMO.AttributeTranslation t in attr.Translations)
                            results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, dim.ID, attr.Name, "Dimension.Attribute", "[" + dim.ID + "].[" + attr.ID + "]", "Caption", t.Language, t.Caption));
                    }

                    // dimension hierarchies
                    foreach (AMO.Hierarchy h in dim.Hierarchies)
                    {
                        results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, dim.ID, h.Name, "Dimension.Hierarchy", "[" + dim.ID + "].[" + h.ID + "]"));
                        foreach (AMO.Translation t in h.Translations)
                            results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, dim.ID, h.Name, "Dimension.Hierarchy", "[" + dim.ID + "].[" + h.ID + "]", "Caption", t.Language, t.Caption));

                        // dimension hierarchy levels
                        foreach (AMO.Level l in h.Levels)
                        {
                            results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, dim.ID, l.Name, "Dimension.Hierarchy.Level", "[" + dim.ID + "].[" + h.ID + "].[" + l.ID + "]"));
                            foreach (AMO.Translation t in l.Translations)
                                results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, dim.ID, l.Name, "Dimension.Hierarchy.Level", "[" + dim.ID + "].[" + h.ID + "].[" + l.ID + "]", "Caption", t.Language, t.Caption));
                        }
                    }

                }

                // cubes
                foreach (AMO.Cube cube in AnalysisDb.Cubes)
                {
                    results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, cube.Name, "Cube", "[" + cube.ID + "]"));
                    foreach (AMO.Translation t in cube.Translations)
                        results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, cube.Name, "Cube", "[" + cube.ID + "]", "Caption", t.Language, t.Caption));

                    // cube dimensions
                    foreach (AMO.CubeDimension d in cube.Dimensions)
                    {
                        results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, d.Name, "Cube.Dimension", "[" + cube.ID + "].[" + d.ID + "]"));
                        foreach (AMO.Translation t in d.Translations)
                            results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, d.Name, "Cube.Dimension", "[" + cube.ID + "].[" + d.ID + "]", "Caption", t.Language, t.Caption));
                    }

                    // measure groups
                    foreach (AMO.MeasureGroup mg in cube.MeasureGroups)
                    {
                        results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, mg.Name, "Cube.MeasureGroup", "[" + cube.ID + "].[" + mg.ID + "]"));
                        foreach (AMO.Translation t in mg.Translations)
                            results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, mg.Name, "Cube.MeasureGroup", "[" + cube.ID + "].[" + mg.ID + "]", "Caption", t.Language, t.Caption));

                        // measures
                        foreach (AMO.Measure m in mg.Measures)
                        {
                            results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, m.Name, "Cube.MeasureGroup.Measure", "[" + cube.ID + "].[" + mg.ID + "].[" + m.ID + "]"));
                            foreach (AMO.Translation t in m.Translations)
                                results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, m.Name, "Cube.MeasureGroup.Measure", "[" + cube.ID + "].[" + mg.ID + "].[" + m.ID + "]", "Caption", t.Language, t.Caption));

                            // display folders for measures
                            if (!string.IsNullOrEmpty(m.DisplayFolder))
                                results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, m.Name, "Cube.MeasureGroup.Measure", "[" + cube.ID + "].[" + mg.ID + "].[" + m.ID + "]", "DisplayFolder", null, m.DisplayFolder));

                            foreach (AMO.Translation t in m.Translations)
                                if (!string.IsNullOrEmpty(t.DisplayFolder))
                                    results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, m.Name, "Cube.MeasureGroup.Measure", "[" + cube.ID + "].[" + mg.ID + "].[" + m.ID + "]", "DisplayFolder", t.Language, t.DisplayFolder));

                        }

                    }

                    // calculated measures (mdx)
                    foreach (AMO.MdxScript mdx in cube.MdxScripts)
                    {
                        foreach (AMO.CalculationProperty cp in mdx.CalculationProperties)
                        {
                            results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, cp.CalculationReference, "Cube.MdxScript.CalculationProperty", "[" + cube.ID + "].[MdxScript].[" + cp.CalculationReference + "]"));
                            foreach (AMO.Translation t in cp.Translations)
                                results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, cp.CalculationReference, "Cube.MdxScript.CalculationProperty", "[" + cube.ID + "].[MdxScript].[" + cp.CalculationReference + "]", "Caption", t.Language, t.Caption));

                            // display folders for calculated measures
                            if (!string.IsNullOrEmpty(cp.DisplayFolder))
                                results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, cp.CalculationReference, "Cube.MdxScript.CalculationProperty", "[" + cube.ID + "].[MdxScript].[" + cp.CalculationReference + "]", "DisplayFolder", null, cp.DisplayFolder));

                            foreach (AMO.Translation t in cp.Translations)
                                if (!string.IsNullOrEmpty(t.DisplayFolder))
                                    results.Add(new ResultRow(AnalysisServerName, AnalysisDatabaseName, cube.ID, cp.CalculationReference, "Cube.MdxScript.CalculationProperty", "[" + cube.ID + "].[MdxScript].[" + cp.CalculationReference + "]", "DisplayFolder", t.Language, t.DisplayFolder));

                        }
                    }

                }

                Console.Write("DONE!\n\r");

            }

            // output data
            using (SqlConnection SqlConn = new SqlConnection())
            {
                try
                {

                    Console.Write("Writing to SQL database " + SqlDatabaseName + " on server " + SqlServerName + " ... ");

                    SqlConn.ConnectionString = SqlConnectionString;
                    SqlConn.Open();

                    using (SqlCommand Command = new SqlCommand())
                    {
                        Command.Connection = SqlConn;
                        Command.CommandText =
                            "IF OBJECT_ID('dbo.SSASDatabaseImport', 'U') IS NULL " +
                            "CREATE TABLE [dbo].[SSASDatabaseImport] (           " +
	                        "    [ServerName] [nvarchar](255) NULL,              " +
	                        "    [DatabaseName] [nvarchar](255) NULL,            " +
	                        "    [AnalyticalObject] [nvarchar](255) NULL,        " +
	                        "    [ElementName] [nvarchar](255) NULL,             " +
	                        "    [TransPath] [nvarchar](255) NULL,               " +
	                        "    [TransKey] [nvarchar](255) NULL,                " +
	                        "    [TransObject] [nvarchar](255) NULL,             " +
	                        "    [TransLanguage] [int] NULL,                     " +
	                        "    [TransValue] [nvarchar](255) NULL               " +
                            ")                                                   ";

                        try
                        {
                            Command.ExecuteNonQuery();
                        }
                        catch
                        {
                            throw;
                        }

                    }

                    using (SqlCommand Command = new SqlCommand())
                    {
                        Command.Connection = SqlConn;
                        Command.CommandText = "DELETE FROM [dbo].[SSASDatabaseImport] WHERE ServerName=@ServerName AND DatabaseName=@DatabaseName ";
                        Command.Parameters.AddWithNullableValue("@ServerName", AnalysisServerName);
                        Command.Parameters.AddWithNullableValue("@DatabaseName", AnalysisDatabaseName);

                        try
                        {
                            Command.ExecuteNonQuery();
                        }
                        catch
                        {
                            throw;
                        }

                    }

                    foreach (ResultRow result in results)
                    {
                        result.Write(SqlConn, "dbo.SSASDatabaseImport");
                    }

                    SqlConn.Close();

                    Console.Write("DONE!\n\r");
 
                }
                catch
                {
                    throw;
                }

            }    

        }
    }

}