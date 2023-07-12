using DbDocumentMaker.Models;
using NPOI.HSSF.Record.Aggregates.Chart;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Column = DbDocumentMaker.Models.Column;

namespace DbDocumentMaker.Utility
{
    static class Export
    {
        public static string TableFile => @"TableSP.sql";
        public static string ColumnFile => @"ColumnSP.sql";

        public static void Procedure(List<object> list)
        {
            List<string> table_ms_sp = new List<string>();
            List<string> column_ms_sp = new List<string>();

            foreach (object obj in list)
            {
                if (obj is Table table)
                {
                    table_ms_sp.Add(CrtTableDescSP(table));
                }

                if (obj is Column column)
                {
                    column_ms_sp.Add(CrtColumnDescSP(column));
                }
            }

            ConfigContent cc = Config.GetInstance().Content;

            OutGrammar(table_ms_sp, Path.Combine(cc.OutputDocLocation, TableFile));
            OutGrammar(column_ms_sp, Path.Combine(cc.OutputDocLocation, ColumnFile));

        }

        private static void OutGrammar(List<string> _ms, string _file)
        {
            string _m = string.Join(";", _ms);
            if (!File.Exists(_file))
            {
                using (StreamWriter sw = File.CreateText(_file))
                {
                    sw.Write(_m);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(_file))
                {
                    sw.Write(_m);
                }
            }
        }

        private static string CrtTableDescSP(Table obj)
        {
            string param = $@"
@name = N'MS_Description',
@value = '{obj.Description}',
@level0type = N'SCHEMA',
@level0name = N'dbo',
@level1type = N'TABLE',
@level1name = N'{obj.TableName}';";

            string SQL = $@"
IF EXISTS(
    SELECT * FROM sys.extended_properties
    WHERE major_id = OBJECT_ID('{obj.TableName}')
    AND minor_id = 0
    AND class = 1
    AND name = 'MS_Description'
)
BEGIN
    EXEC sys.sp_updateextendedproperty {param}
END
ELSE
BEGIN
    EXEC sys.sp_addextendedproperty {param}
END";

            return SQL;
        }

        private static string CrtColumnDescSP(Column obj)
        {
            string param = $@"'MS_Description', '{obj.Description}', 'user', 'dbo', 'table', '{obj.TableName}', 'column', '{obj.ColumnName}'";

            string SQL = $@"
IF not exists(SELECT * FROM ::fn_listextendedproperty (NULL, 'user', 'dbo', 'table', '{obj.TableName}', 'column', '{obj.ColumnName}'))
BEGIN  
    exec sp_addextendedproperty {param}
END  
ELSE
BEGIN  
    exec sp_updateextendedproperty {param}
END";

            return SQL;
        }
    }
}
