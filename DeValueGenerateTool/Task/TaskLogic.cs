using System.Data;

namespace DeValueGenerateTool.Task
{
    public class TaskLogic
    {
        ImportDt importDt=new ImportDt();
        GenerateDt generateDt=new GenerateDt();
        ExportDt exportDt=new ExportDt();

        #region 变量参数
        private int _taskid;
        private string _fileAddress;       //文件地址
        private DataTable _dt;             //获取dt(从EXCEL获取的DT)
        private DataTable _tempdt;         //保存运算成功的DT(导出时使用)

        private DataTable _resultTable;   //返回DT
        private bool _resultMark;        //返回是否成功标记
        #endregion

        #region Set
        /// <summary>
        /// 中转ID
        /// </summary>
        public int TaskId { set { _taskid = value; } }

        /// <summary>
        /// //接收文件地址信息
        /// </summary>
        public string FileAddress { set { _fileAddress = value; } }

        /// <summary>
        /// 获取dt(从EXCEL获取的DT)
        /// </summary>
        public DataTable Data { set { _dt = value; } }
        #endregion

        #region Get
        /// <summary>
        ///返回DataTable至主窗体
        /// </summary>
        public DataTable RestulTable => _resultTable;

        /// <summary>
        ///  返回是否成功标记
        /// </summary>
        public bool ResultMark => _resultMark;

        /// <summary>
        /// 返回运算成功的表头DT(导出时使用)
        /// </summary>
        public DataTable Tempdt => _tempdt;
        #endregion

        public void StartTask()
        {
            switch (_taskid)
            {
                //导入
                case 0:
                    OpenExcelImporttoDt(_fileAddress);
                    break;
                //运算
                case 1:

                    break;
                //导出
                case 2:

                    break;
            }
        }

        /// <summary>
        /// 导入EXCEL数据(会出现两次,先导入‘标准色’Excel记录,再导入‘样品色’Excel记录)
        /// 注:需检查_resultdt有没有值,若有,就先清空内容,再赋值
        /// </summary>
        /// <param name="fileAddress"></param>
        private void OpenExcelImporttoDt(string fileAddress)
        {
            //
            if (_resultTable?.Rows.Count >= 0)
            {
                _resultTable.Rows.Clear();
                _resultTable.Columns.Clear();
            }
            //
            _resultTable = importDt.OpenExcelImporttoDt(fileAddress);
        }

        /// <summary>
        /// 运算
        /// </summary>
        private void GenerateRecord()
        {
            
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="fileAddress"></param>
        private void ExportDtToExcel(string fileAddress)
        {
            
        }
    }
}
