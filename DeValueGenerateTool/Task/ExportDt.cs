﻿using System;
using System.Data;
using NPOI.XSSF.UserModel;

namespace DeValueGenerateTool.Task
{
    public class ExportDt
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="fileAdd"></param>
        /// <param name="sourcedt"></param>
        /// <returns></returns>
        public bool ExportDtToExcel(string fileAdd, DataTable sourcedt)
        {
            var result = true;
            var sheetcount = 0;  //记录所需的sheet页总数
            var rownum = 1;

            try
            {
                //声明一个WorkBook
                var xssfWorkbook = new XSSFWorkbook();

                //执行sheet页(注:1)先列表temp行数判断需拆分多少个sheet表进行填充; 以一个sheet表有100W行记录填充为基准)
                sheetcount = sourcedt.Rows.Count % 1000000 == 0 ? sourcedt.Rows.Count / 1000000 : sourcedt.Rows.Count / 1000000 + 1;

                //i为EXCEL的Sheet页数ID
                for (var i = 1; i <= sheetcount; i++)
                {
                    //创建sheet页
                    var sheet = xssfWorkbook.CreateSheet("Sheet" + i);
                    //创建"标题行"
                    var row = sheet.CreateRow(0);

                    //创建sheet页各列标题
                    for (var j = 0; j < sourcedt.Columns.Count; j++)
                    {
                        //设置列宽度
                        sheet.SetColumnWidth(j, (int)((20 + 0.72) * 256));
                        //创建标题
                        row.CreateCell(j).SetCellValue("制造商");
                    }

                    //计算进行循环的起始行
                    var startrow = (i - 1) * 1000000;
                    //计算进行循环的结束行
                    var endrow = i == sheetcount ? sourcedt.Rows.Count : i * 1000000;


                }
            }
            catch (Exception)
            {
                result = false;
            }

            return result;
        }
    }
}
