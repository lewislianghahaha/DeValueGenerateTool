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
        /// <param name="standardColorDt"></param>
        /// <param name="sampleColorDt"></param>
        public DataTable GenerateExcelSourceDt(DataTable standardColorDt, DataTable sampleColorDt)
        {
            //获取导出临时表
            var resultdt = dbList.ExportDt();



            return resultdt;
        }
    }
}
