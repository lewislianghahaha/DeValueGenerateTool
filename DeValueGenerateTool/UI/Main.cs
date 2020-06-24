using System;
using System.Data;
using System.Threading;
using System.Windows.Forms;
using BomOfferOrder.UI;
using DeValueGenerateTool.Task;

namespace DeValueGenerateTool.UI
{
    public partial class Main : Form
    {
        TaskLogic task=new TaskLogic();
        Load load=new Load();

        #region 变量参数
        //获取‘标准色’DT
        private DataTable _standardColorDt;
        //获取‘样品色’DT
        private DataTable _sampleColorDt;
        #endregion

        public Main()
        {
            InitializeComponent();
            OnRegisterEvents();
            OnInitial();
        }

        private void OnRegisterEvents()
        {
            btnimport.Click += Btnimport_Click;
            this.FormClosing += Main_FormClosing;
        }

        private void OnInitial()
        {
            //检测若_standardColorDt及_sampleColorDt有值,即在初始化时清空这两个临时表
            if (_standardColorDt?.Rows.Count >= 0)
            {
                _standardColorDt.Rows.Clear();
                _standardColorDt.Columns.Clear();
            }
            if (_sampleColorDt?.Rows.Count >= 0)
            {
                _sampleColorDt.Rows.Clear();
                _sampleColorDt.Columns.Clear();
            }
        }

        /// <summary>
        /// 导入记录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btnimport_Click(object sender, EventArgs e)
        {
            try
            {
                //提示信息
                var mess_Sample = "‘标准色’模板数据导入完成,请继续执行导入‘样品色’模板数据";
                var mess_Gen = "已导入成功,进入运算.";
                var mess_export = "运算成功,导出至Excel.";

                //检测若_standardColorDt及_sampleColorDt有值,即在初始化时清空这两个临时表
                OnInitial();

                if(!GetAddAndImportDt(0)) throw new Exception("没有导入'标准色'模板数据,请重新导入'标准色'数据");

                if (_standardColorDt.Rows.Count == 0) throw new Exception("不能成功导入'标准色'EXCEL内容,请检查模板内数据是否正确.");
                else
                {
                    MessageBox.Show(mess_Sample, $"提示",MessageBoxButtons.OK,MessageBoxIcon.Information);

                    if(!GetAddAndImportDt(1)) throw new Exception("没有导入'样品色'模板数据,请重新导入'标准色'数据,再导入'样品色'数据");

                    if(_sampleColorDt.Rows.Count == 0) throw new Exception("不能成功导入'样品色'EXCEL内容,请检查模板内数据是否正确.");
                    else
                    {
                        MessageBox.Show(mess_Gen,$"提示",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        if(!Generatedt()) throw new Exception("运算结果没有记录,请检查是否与实际情况一致");
                        else
                        {
                            MessageBox.Show(mess_export, $"提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ExportDt();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, $"错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取导入地址及导入
        /// </summary>
        /// <param name="id">0:导入‘标准色’模板 1:导入‘样品色’模板</param>
        private bool GetAddAndImportDt(int id)
        {
            //定义是否进行导入的标记
            var result = false;

            var openFileDialog = new OpenFileDialog { Filter = $"Xlsx文件|*.xlsx" };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var fileAdd = openFileDialog.FileName;

                //将所需的值赋到Task类内
                task.TaskId = 0;
                task.FileAddress = fileAdd;

                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();

                if (id == 0)
                {
                    _standardColorDt = task.RestulTable.Copy();
                    result = true;
                }
                else
                {
                    _sampleColorDt = task.RestulTable.Copy();
                    result = true;
                }
            }
            return result;
        }

        /// <summary>
        /// 运算
        /// </summary>
        /// <returns></returns>
        bool Generatedt()
        {
            var result = true;
            try
            {
                task.TaskId = 1;
                task.StandardColorDt = _standardColorDt;  //标准色DT
                task.SampleColorDt = _sampleColorDt;      //样品色DT

                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();

                result = task.ResultMark;
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }

        /// <summary>
        /// 导出功能
        /// </summary>
        void ExportDt()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog { Filter = $"Xlsx文件|*.xlsx" };
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var fileAdd = saveFileDialog.FileName;

                    task.TaskId = 2;
                    task.FileAddress = fileAdd;

                    //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                    new Thread(Start).Start();
                    load.StartPosition = FormStartPosition.CenterScreen;
                    load.ShowDialog();

                    if (!task.ResultMark) throw new Exception("导出异常");
                    else
                    {
                        MessageBox.Show($"导出成功!可从EXCEL中查阅导出效果", $"提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, $"错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 关闭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            var clickMessage = $"是否退出?";
            if (e.CloseReason != CloseReason.ApplicationExitCall)
            {
                var result = MessageBox.Show(clickMessage, $"提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                //当点击"OK"按钮时执行以下操作
                if (result == DialogResult.Yes)
                {
                    //退出前将_standardColorDt,_sampleColorDt表清空
                    if (_standardColorDt?.Rows.Count >= 0)
                    {
                        _standardColorDt.Rows.Clear();
                        _standardColorDt.Columns.Clear();
                    }
                    if (_sampleColorDt?.Rows.Count >= 0)
                    {
                        _sampleColorDt.Rows.Clear();
                        _sampleColorDt.Columns.Clear();
                    }
                    //允许窗体关闭
                    e.Cancel = false;
                }
                else
                {
                    //将Cancel属性设置为 true 可以"阻止"窗体关闭
                    e.Cancel = true;
                }
            }
        }

        /// <summary>
        ///子线程使用(重:用于监视功能调用情况,当完成时进行关闭LoadForm)
        /// </summary>
        private void Start()
        {
            task.StartTask();

            //当完成后将Form2子窗体关闭
            this.Invoke((ThreadStart)(() => {
                load.Close();
            }));
        }
    }
}
