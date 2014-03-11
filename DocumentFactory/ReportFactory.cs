using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace ReportHelper
{
    class ReportFactory
    {
        /// <summary>
        /// 工廠依據訂單生產報表
        /// </summary>
        /// <param name="order">訂單</param>
        /// <returns>實作 Document 之物件</returns>
        public static T CreateExcelDocument<T>(string order) where T : ExcelDocument 
        {
            var obj = (T)Assembly.Load(typeof(T).Assembly.FullName).CreateInstance(typeof(T).Assembly.GetName().Name + "." + order); 
            return obj; 
        }
    }
}
