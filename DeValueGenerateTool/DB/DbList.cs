using System;
using System.Data;

namespace DeValueGenerateTool.DB
{
    public class DbList
    {
        /// <summary>
        /// 导入模板(包括:‘标准色’数据 及 ‘样品色’数据)
        /// </summary>
        /// <returns></returns>
        public DataTable ImportDt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 5; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 1:
                        dc.ColumnName = "角度";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "L";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 3:
                        dc.ColumnName = "A";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 4:
                        dc.ColumnName = "B";
                        dc.DataType = Type.GetType("System.Double"); 
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 导出语句
        /// </summary>
        /// <returns></returns>
        public DataTable ExportDt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 2; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 1:
                        dc.ColumnName = "DE值";
                        dc.DataType = Type.GetType("System.Double"); 
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

    }
}
