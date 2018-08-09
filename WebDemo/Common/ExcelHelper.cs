using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace Demo
{
    /// <summary>
    /// Excel导入导出助手
    /// NuGet：EPPlus.Core
    /// </summary>
    public class ExcelHelper
    {

        /// private ExcelHelper(){}

        /// <summary>
        /// Excel文件 Content-Type
        /// </summary>
        private const string Excel = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        #region Excel导出

        /// <summary>
        /// Excel导出
        /// </summary>
        /// <param name="keyValuePairs">字典表【名称，数据】</param>
        /// <param name="sWebRootFolder">网站根文件夹</param>
        /// <param name="tuple">item1:The virtual path of the file to be returned.|item2:The Content-Type of the file</param>
        public static void Export(Dictionary<string, DataTable> keyValuePairs, string sWebRootFolder, out Tuple<string, string> tuple, string sFileName)
        {
            if (string.IsNullOrWhiteSpace(sWebRootFolder))
                tuple = Tuple.Create("", Excel);

            sFileName = $"{sFileName}.xlsx";

            //if (File.Exists(sFileName))
            //{

                //string sFileName = $"{DateTime.Now.ToString("yyyyMMddHHmmssfff")}-{FormatGuid.GetGuid(FormatGuid.GuidType.N)}.xlsx";
                FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    foreach (var item in keyValuePairs)
                    {
                        string worksheetTitle = item.Key; //表名称
                        var dt = item.Value; //数据表

                        // 添加worksheet
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetTitle);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                if (i == 0)
                                {
                                    //添加表头
                                    worksheet.Cells[1, j + 1].Value = dt.Columns[j].ColumnName;
                                    worksheet.Cells[1, j + 1].Style.Font.Bold = true;
                                }
                                else
                                {
                                    //添加值
                                    worksheet.Cells[i + 1, j + 1].Value = dt.Rows[i][j].ToString();
                                }
                            }
                        }
                    }
                    package.Save();
                //}
                tuple = Tuple.Create(sFileName, Excel);
            }
            //tuple = null;
        }
        #endregion     
    }
}
