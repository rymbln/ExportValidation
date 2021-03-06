﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using ConsoleApplication1;
using ExportValidation.Common;

namespace ExportValidationConsole
{
    class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var strPath = args[0];
                var strServer = args[1];
                var strDBName = args[2];
                var strUser = args[3];
                var strPassword = args[4];
                var strMethod = args[5];
                var strProject = args[6];

                var conn = Tools.GetConnectionString(strServer, strDBName, strUser, strPassword);
                if (strMethod.Equals("RUN_VALIDATION"))
                {
                    using (conn)
                    {
                        var res = new ReturnProc(null, null);
                        try
                        {
                            res = Tools.RunProcedure(conn, "RUN_VALIDATION", strProject);
                      //      index = Tools.GetIndex(conn, "RUN_VALIDATION");
                            if (res.Data.Count > 0)
                            {
                                ExcelGeneration.GenerateDocument(strPath, res);
                                
                            }
                            else
                            {
                                Log.Write("NoData");
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Write(ex);
                        }
                        finally
                        {
                            Log.Write("Finish");
                        }
                    }
                }
                else if (strMethod.Equals("RUN_EXPORT"))
                {
//                    var data = new List<QueryData>();
                    var res = new ReturnProc(null, null);
                    try
                    {
                        res = Tools.RunProcedure(conn, "RUN_EXPORT", strProject);
                     
                        if (res.Data.Count > 0)
                        {
                            ExcelGeneration.GenerateDocument2(strPath, res);
                            Log.Write("Finish");
                        }
                        else
                        {
                            Log.Write("NoData");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Write(ex);
                    }
                    finally
                    {
                        Log.Write("Finish");
                    }
                }

                else if (strMethod.Equals("RUN_EXPORT_CSV"))
                {
                    string fileName = "";
                    string sql = "";

                    var index = new List<IndexData>();

                    try
                    {
                        index = Tools.GetIndex(conn, "RUN_EXPORT");
                        Log.Write("Finish");
                    }
                    catch (Exception ex)
                    {
                        Log.Write(ex);
                    }
                    finally
                    {
                        if (index.Count > 0)
                        {
                            foreach (var queryData in index)
                            {
                                fileName = queryData.NameList;
                                sql = queryData.SelectCommand;
                                Tools.ExportToCSVFile(strPath, strProject, fileName, sql, conn,
                                    Encoding.GetEncoding(1251), ";", true);
                            }
                      
                        }
                        else
                        {
                            Log.Write("NoData");
                        }
                            Log.Write("Finish");
                        
                    }
                }
                else if (strMethod.Equals("RUN_QUERY"))
                {
                    try
                    {
                        Tools.GetQueriesInFormat(conn, strProject, strPath);
                      
                    }
                    catch (Exception ex)
                    {
                        Log.Write(ex);
                    }
                    finally
                    {
                        Log.Write("Finish");
                    }

                }
                else if (strMethod.Equals("RUN_ACTIVITY"))
                {
                    var res = new ReturnProc(null, null);
                    try
                    {
                        res = Tools.RunProcedure(conn, "RUN_ACTIVITY", strProject);
                 //       var index = Tools.GetIndex(conn, "RUN_ACTIVITY");
                        if (res.Data.Count > 0)
                        {
                            ExcelGeneration.GenerateDocument(strPath, res);
                     
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Write(ex);

                    }
                    finally
                    {
                        Log.Write("Finish");
                    }
                }
                else if (strMethod.Equals("RUN_SYNC"))
                {
                    try
                    {
                        Tools.RunProcedureNonQuery(conn, "RUN_SYNC");
                    }
                    catch (Exception ex)
                    {
                        Log.Write(ex);
                    }
                    finally
                    {
                        Log.Write("Finish");
                    }
                }
                else
                {

                }
            }
            catch (Exception)
            {
                Log.Write("Error input");
            }

            finally
            {
                Console.ReadKey();
            }


        }


    }
}
