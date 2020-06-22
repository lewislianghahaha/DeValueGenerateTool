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
        private string _fileAddress;        //文件地址

        private DataTable _standardColorDt; //获取‘标准色’EXCEL导入的DT
        private DataTable _sampleColorDt;   //获取‘样品色’EXCEL导入的DT

        private DataTable _tempdt;         //保存运算成功的DT(导出时使用)

        private DataTable _resultTable;    //返回DT
        private bool _resultMark;          //返回是否成功标记
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
        /// 获取‘标准色’DT
        /// </summary>
        public DataTable StandardColorDt { set { _standardColorDt = value; } }

        /// <summary>
        /// 获取‘样品色’DT
        /// </summary>
        public DataTable SampleColorDt { set { _sampleColorDt = value; } }
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
                    GenerateRecord(_standardColorDt,_sampleColorDt);
                    break;
                //导出
                case 2:
                    ExportDtToExcel(_fileAddress, _tempdt);
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
            //检测若发现_resultTable有值,即先将_resultTable清空,再进行赋值
            if (_resultTable?.Rows.Count >= 0)
            {
                _resultTable.Rows.Clear();
                _resultTable.Columns.Clear();
            }
            //导入
            _resultTable = importDt.OpenExcelImporttoDt(fileAddress);
        }

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="standardColorDt">标准色DT</param>
        /// <param name="sampleColorDt">样品色DT</param>
        private void GenerateRecord(DataTable standardColorDt, DataTable sampleColorDt)
        {
            //检测若发现_tempdt有值,即先将_tempdt清空,再进行赋值
            if (_tempdt?.Rows.Count >= 0)
            {
                _tempdt.Rows.Clear();
                _tempdt.Columns.Clear();
            }
            _tempdt = generateDt.GenerateExcelSourceDt(standardColorDt, sampleColorDt);
            _resultMark = _tempdt.Rows.Count > 0;
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="fileAddress"></param>
        /// <param name="tempdt">整理后的DT</param>
        private void ExportDtToExcel(string fileAddress, DataTable tempdt)
        {
            _resultMark = exportDt.ExportDtToExcel(fileAddress, tempdt);
        }
    }
}
