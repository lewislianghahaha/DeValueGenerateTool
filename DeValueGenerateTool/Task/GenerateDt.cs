using System.Data;
using DeValueGenerateTool.DB;

namespace DeValueGenerateTool.Task
{
    public class GenerateDt
    {
        DbList dbList=new DbList();

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="standardColorDt">'标准色'DT</param>
        /// <param name="sampleColorDt">‘样品色’DT</param>
        public DataTable GenerateExcelSourceDt(DataTable standardColorDt, DataTable sampleColorDt)
        {
            //获取导出临时表
            var resultdt = dbList.ExportDt();



            return resultdt;
        }
    }
}
