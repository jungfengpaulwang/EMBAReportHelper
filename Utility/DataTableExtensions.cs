using System;
using System.Collections.Generic;
using System.Data;
using System.Collections;
using System.Reflection;

namespace ReportHelper
{
    public static class DataTableExtensions
    {
        public static DataTable ToDataTable(this object obj, string dataTableName, string columnName)
        {
            return ToDataTable(new object[] { obj }, dataTableName, columnName, null);
        }

        public static DataTable ToDataTable(this object obj, string dataTableName, string columnName, System.Type columnType)
        {
            return ToDataTable(new object[] { obj }, dataTableName, columnName, columnType);
        }

        public static DataTable ToDataTable(this IEnumerable list, string dataTableName, string columnName)
        {
            return ToDataTable(list, dataTableName, columnName, null);
        }

        public static DataTable ToDataTable(this IEnumerable list, string dataTableName, string columnName, System.Type columnType)
        {
            if (list == null)
                return ToDataTable(new object[] { string.Empty }, dataTableName, columnName, Type.GetType("System.String"));

            if (list.GetType() == Type.GetType("System.String"))
                return ToDataTable(new object[] { list }, dataTableName, columnName, Type.GetType("System.String"));

            if (list.GetType() == Type.GetType("System.Byte[]"))
                return ToDataTable(new object[] { list }, dataTableName, columnName, Type.GetType("System.Byte[]"));

            DataTable dt = new DataTable(dataTableName);
            try
            {
                if (columnType == null)
                    dt.Columns.Add(columnName, Type.GetType("System.String"));
                else
                    dt.Columns.Add(columnName, columnType);

                foreach (object item in list)
                {
                    DataRow row = dt.NewRow();

                    row[columnName] = item;

                    dt.Rows.Add(row);
                }

                dt.AcceptChanges();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return dt;
        }
    }
}

//  object o = si.GetType().GetProperty(propertyName).GetValue(si, null);