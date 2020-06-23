using System;
using System.Data;
using DeValueGenerateTool.DB;

namespace DeValueGenerateTool.Task
{
    public class GenerateDt
    {
        DbList dbList=new DbList();

        #region 变量参数
        //角度临时表
        private DataTable _temprangedt;
        //计算同一个配方内各角度DE相关值
        private DataTable _tempdt;
        #endregion

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="standardColorDt">'标准色'DT</param>
        /// <param name="sampleColorDt">‘样品色’DT</param>
        public DataTable GenerateExcelSourceDt(DataTable standardColorDt, DataTable sampleColorDt)
        {
            var resultdt = new DataTable();
            //初始化角度临时表
            OnInitializeTempRange();

            try
            {
                //获取‘样品色’的内部色号记录(注:唯一值)
                var tempheaddt = GetTempHeaddt(sampleColorDt);
                //循环tempheaddt,并利用它提供的‘内部色号’分别放到standardColorDt 及 sampleColorDt进行获取数据
                foreach (DataRow rows in tempheaddt.Rows)
                {
                    var standardrows = standardColorDt.Select("内部色号='" + Convert.ToString(rows[0]) + "'");
                    var samplerows = sampleColorDt.Select("内部色号='" + Convert.ToString(rows[0]) + "'");

                    //使用‘角度’临时表循环获取'同一个内部色号'下的角度对应的L A B值
                    for (var i = 0; i < _temprangedt.Rows.Count; i++)
                    {
                        _tempdt?.Merge(GetTempDt(standardrows,samplerows,Convert.ToString(_temprangedt.Rows[i])));
                    }
                    //最后将结果插入至resultdt内(注:需使用SUM(5个角度的DE值)/5)
                    resultdt.Merge(GenerateDeValueToDt(_tempdt));
                    //最后将_tempdt临时表数据清空
                    if (_tempdt?.Rows.Count >= 0)
                    {
                        _tempdt.Rows.Clear();
                        _tempdt.Columns.Clear();
                    }
                }
            }
            catch (Exception)
            {
                resultdt.Rows.Clear();
                resultdt.Columns.Clear();
            }

            //退出前将各TEMP临时表清空
            if (_temprangedt?.Rows.Count >= 0)
            {
                _temprangedt.Rows.Clear();
                _temprangedt.Columns.Clear();
            }
            if (_tempdt?.Rows.Count >= 0)
            {
                _tempdt.Rows.Clear();
                _tempdt.Columns.Clear();
            }
            return resultdt;
        }

        /// <summary>
        /// 整理并插入至导出临时表内
        /// </summary>
        /// <param name="tempdt"></param>
        /// <returns></returns>
        private DataTable GenerateDeValueToDt(DataTable tempdt)
        {
            //获取导出临时表
            var resultdt = dbList.ExportDt();
            //定义DE值和变量
            var sumde = 0.0;

            //循环将tempdt的DE值求和/5得出结果将插入至resultdt内
            foreach (DataRow rows in tempdt.Rows)
            {
                sumde += Convert.ToDouble(rows[5]);
            }
            //插入
            var newrow = resultdt.NewRow();
            newrow[0] = tempdt.Rows[0][0];  //内部色号
            newrow[1] = sumde / 5;          //DE值
            resultdt.Rows.Add(newrow);
            return resultdt;
        }

        /// <summary>
        /// 获取‘样品色’的内部色号记录(注:唯一值)
        /// </summary>
        /// <returns></returns>
        private DataTable GetTempHeaddt(DataTable samplesourcedt)
        {
            var resultdt = dbList.GetTempHeadDt();
            //定义内部色号变量
            var code = "";

            foreach (DataRow rows in samplesourcedt.Rows)
            {
                if (string.IsNullOrEmpty(code))
                {
                    var newrow = resultdt.NewRow();
                    newrow[0] = rows[0];
                    resultdt.Rows.Add(newrow);
                }
                else if (code != "" && code != Convert.ToString(rows[0]))
                {
                    var newrow = resultdt.NewRow();
                    newrow[0] = rows[0];
                    resultdt.Rows.Add(newrow);
                }
                code = Convert.ToString(rows[0]);
            }
            return resultdt;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="standardrows">'标准色'行记录</param>
        /// <param name="samplerows">‘样品色’行记录</param>
        /// <param name="range">角度ID</param>
        /// <returns></returns>
        private DataTable GetTempDt(DataRow[] standardrows,DataRow[] samplerows,string range)
        {
            //获取‘内部色号对应各角度结果临时表’
            var resultdt = dbList.GetTempDt();

            //判断若standardrows行数大于0,即执行
            if (standardrows.Length > 0)
            {
                //检测若standardrows 并且 samplerows没有rangeid对应的角度数据,即插入空白值,并马上跳出
                if (CheckIncludeRange(standardrows,samplerows,range))
                {
                    //循环samplerows行记录
                    for (var i = 0; i < samplerows.Length; i++)
                    {
                        //若判断到samplerows内包含rangeid,即继续
                        if (Convert.ToString(samplerows[i][1]).Contains(range))
                        {
                            for (var j = 0; j < standardrows.Length; j++)
                            {
                                //计算 DEL DEA DEB值,若standardrows没有所包含的‘角度’
                                //注:分别将‘样品色’的L A B 值与‘标准色’的L A B 值相减得出DEL DEA DEB值
                                //而DE值公式为:MATH.SQRT(DEL*DEL+DEA*DEA+DEB*DEB)
                                if (Convert.ToString(standardrows[j][1]).Contains(range))
                                {
                                    //分别获取DEL DEA DEB值
                                    var del = Convert.ToDouble(samplerows[i][2]) - Convert.ToDouble(standardrows[j][2]);
                                    var dea = Convert.ToDouble(samplerows[i][3]) - Convert.ToDouble(standardrows[j][3]);
                                    var deb = Convert.ToDouble(samplerows[i][4]) - Convert.ToDouble(standardrows[j][4]);

                                    var newrow = resultdt.NewRow();
                                    newrow[0] = Convert.ToString(samplerows[0][0]);            //内部色号
                                    newrow[1] = range;                                         //角度
                                    newrow[2] = del;                                           //DEL
                                    newrow[3] = dea;                                           //DEA
                                    newrow[4] = deb;                                           //DEB
                                    newrow[5] = Math.Sqrt(del * del + dea * dea + deb * deb);  //DE
                                    resultdt.Rows.Add(newrow);
                                }
                            }
                        }
                    }
                }
                else
                {
                    resultdt = InsertEmptyTempdt(Convert.ToString(samplerows[0][0]),range,resultdt);
                }
            }
            //若为0,即将空值插入
            else
            {
                resultdt = InsertEmptyTempdt(Convert.ToString(samplerows[0][0]), range, resultdt);
            }
            return resultdt;
        }

        /// <summary>
        /// 检测standardrows 或 samplerows有没有rangeid对应的角度数据
        /// </summary>
        /// <param name="standardrows">'标准色'行记录</param>
        /// <param name="samplerows">‘样品色’行记录</param>
        /// <param name="range"></param>
        /// <returns></returns>
        private bool CheckIncludeRange(DataRow[] standardrows,DataRow[] samplerows,string range)
        {
            var standardvalue = false;
            var samplevalue = false;
            var result = false;

            //检测standardrows是否有range记录
            for (var i = 0; i < standardrows.Length; i++)
            {
                if (Convert.ToString(standardrows[i][1]).Contains(range))
                {
                    standardvalue = true;
                    break;
                }
            }

            //检测samplerows是否有range记录
            for (var i = 0; i < samplerows.Length; i++)
            {
                if (Convert.ToString(samplerows[i][1]).Contains(range))
                {
                    samplevalue = true;
                    break;
                }
            }

            //必须要两者都有记录才会返回true
            if (standardvalue && samplevalue)
            {
                result = true;
            }
            return result;
        }

        /// <summary>
        /// 插入空白记录TEMPDT
        /// </summary>
        /// <param name="code"></param>
        /// <param name="range"></param>
        /// <param name="resultdt"></param>
        /// <returns></returns>
        private DataTable InsertEmptyTempdt(string code,string range,DataTable resultdt)
        {
            var newrow = resultdt.NewRow();
            newrow[0] = code;    //内部色号
            newrow[1] = range;   //角度
            newrow[2] = 0;       //DEL
            newrow[3] = 0;       //DEA
            newrow[4] = 0;       //DEB
            newrow[5] = 0;       //DE
            resultdt.Rows.Add(newrow);
            return resultdt;
        }

        /// <summary>
        /// 初始化角度临时表
        /// </summary>
        private void OnInitializeTempRange()
        {
            var dt = new DataTable();

            //创建表头
            for (var i = 0; i < 5; i++)
            {
                var dc = new DataColumn();
                switch (i)
                {
                    case 0:
                        dc.ColumnName = "Range";
                        break;
                }
                dt.Columns.Add(dc);
            }

            //创建行内容
            for (var j = 0; j < 5; j++)
            {
                var dr = dt.NewRow();

                switch (j)
                {
                    case 0:
                        dr[0] = "15";
                        break;
                    case 1:
                        dr[0] = "25";
                        break;
                    case 2:
                        dr[0] = "45";
                        break;
                    case 3:
                        dr[0] = "75";
                        break;
                    case 4:
                        dr[0] = "110";
                        break;
                }
                dt.Rows.Add(dr);
            }
            _temprangedt = dt.Copy();
        }

    }
}
