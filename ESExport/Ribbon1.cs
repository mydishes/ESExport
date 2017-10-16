using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ESExport
{
    public partial class Ribbon1
    {
        #region RegistStoreData
        public static readonly string registryPath = @"Software\YourCompanyName\YourAddInName";

        public static void StoreInRegistry(string keyName, string value)
        {
            RegistryKey rootKey = Registry.CurrentUser;
            using (RegistryKey rk = rootKey.CreateSubKey(registryPath))
            {
                rk.SetValue(keyName, value, RegistryValueKind.String);
            }
        }

        public static string ReadFromRegistry(string keyName, string defaultValue = "")
        {
            RegistryKey rootKey = Registry.CurrentUser;
            using (RegistryKey rk = rootKey.OpenSubKey(registryPath, false))
            {
                if (rk == null)
                {
                    return defaultValue;
                }

                var res = rk.GetValue(keyName, defaultValue);
                if (res == null)
                {
                    return defaultValue;
                }

                return res.ToString();
            }
        }
        #endregion
        public static readonly string EndSign = @"#End";
        public static string exportPath = ReadFromRegistry("ExportPath");

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        void Save2File(string name, string data)
        {
            var path = Path.Combine(exportPath, name);
            using (StreamWriter sw = File.CreateText(path))
            {
                sw.Write(data);
            }
            MessageBox.Show("导出文件: " + path);
        }

        /// <summary>
        /// 检查有效数据区
        /// </summary>
        void CheckDataCount(Excel.Worksheet worksheet, out int rowCount, out int columnCount)
        {
            rowCount = worksheet.UsedRange.Rows.Count;
            columnCount = worksheet.UsedRange.Columns.Count;
            int k;
            for (k = 1; k <= columnCount; k++)
            {
                var value = (worksheet.Cells[1, k] as Excel.Range).Value;
                if (value == null)
                {
                    break;
                }
                string param_name = value.ToString();
                param_name = param_name.Trim();
                if (param_name == EndSign)
                {
                    break;
                }
                if (string.IsNullOrEmpty(param_name))
                {
                    break;
                }
            }
            columnCount = k - 1;

            for (k = 1; k <= rowCount; k++)
            {
                bool null_row = true;
                for (int j = 1; j <= columnCount; j++)
                {
                    var value = (worksheet.Cells[k, j] as Excel.Range).Value;
                    if (value == null)
                    {
                        continue;
                    }
                    string cell_var = value.ToString();
                    if (j == 1 && cell_var == EndSign)
                    {
                        break;
                    }
                    if (!string.IsNullOrEmpty(cell_var.Trim()))
                    {
                        null_row = false;
                        break;
                    }
                }
                if (null_row)
                {
                    break;
                }
            }
            rowCount = k - 1;
        }


        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // export c#
            Excel.Worksheet worksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            int rowCount, columnCount;
            CheckDataCount(worksheet, out rowCount, out columnCount);
            if (rowCount < 3)
            {
                MessageBox.Show("至少需要3行数据!（变量名,类型,数据)");
                return;
            }
            string[] param_names = new string[columnCount];
            string[] param_types = new string[columnCount];
            for (int k = 1; k <= columnCount; k++)
            {
                var value1 = (worksheet.Cells[1, k] as Excel.Range).Value;
                var value2 = (worksheet.Cells[2, k] as Excel.Range).Value;
                if (value1 == null || value2 == null)
                {
                    MessageBox.Show("前两行定义不能有空!");
                    return;
                }
                string param_name = value1.ToString();
                param_name = param_name.Trim();
                string param_type = value2.ToString();
                param_type = param_type.Trim();
                if (string.IsNullOrEmpty(param_name) || string.IsNullOrEmpty(param_type))
                {
                    MessageBox.Show("前两行定义不能有空!");
                    return;
                }
                param_names[k - 1] = param_name;
                param_types[k - 1] = param_type;
            }

            string s = System.Configuration.ConfigurationManager.AppSettings.Get("template");
            //去注释
            Regex regex = new Regex("//[^\\r\\n]*[\\r\\n]+");
            var match = regex.Match(s);
            while (match.Success)
            {
                s = s.Replace(match.Value, "");
                match = regex.Match(s);
            }
            // 替换表格名字
            string clsName = worksheet.Name;
            regex = new Regex(@"#SHEET_NAME#");
            match = regex.Match(s);
            while (match.Success)
            {
                s = s.Replace(match.Value, clsName);
                match = regex.Match(s);
            }

            // 参数定义
            regex = new Regex("#PARAM_DEF#.*#PARAM_END#", RegexOptions.Singleline);
            match = regex.Match(s);
            if (match.Success)
            {
                string formatStr = match.Value.Replace("#PARAM_DEF#", "");
                formatStr = formatStr.Replace("#PARAM_END#", "").Trim();

                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < columnCount; i++)
                {
                    sb.AppendFormat(formatStr, param_types[i], param_names[i]);
                    if (i != columnCount - 1)
                    {
                        sb.AppendLine();
                    }
                }
                s = s.Replace(match.Value, sb.ToString());
            }
            else
            {
                MessageBox.Show("#PARAM_DEF# 匹配失败！");
                return;
            }

            // 数据定义
            regex = new Regex("#DATA_DEF#.*#DATA_END#", RegexOptions.Singleline);
            match = regex.Match(s);
            if (match.Success)
            {
                string matchAllStr = match.Value;
                string formatStr = matchAllStr.Replace("#DATA_DEF#", "");
                formatStr = formatStr.Replace("#DATA_END#", "").Trim();

                regex = new Regex("#PAIR_DEF#.*#PAIR_END#", RegexOptions.Singleline);
                match = regex.Match(formatStr);
                if (match.Success)
                {
                    string pairStr = match.Value;
                    string formatStr2 = pairStr.Replace("#PAIR_DEF#", "");
                    formatStr2 = formatStr2.Replace("#PAIR_END#", "").Trim();


                    regex = new Regex(@"string");
                    StringBuilder sb = new StringBuilder();
                    StringBuilder sb2 = new StringBuilder();
                    for (int i = 3; i <= rowCount; i++)
                    {
                        sb2.Clear();
                        for (int j = 1; j <= columnCount; j++)
                        {

                            string stype = param_types[j - 1];
                            string sname = param_names[j - 1];
                            bool strflg = regex.Match(stype).Success;

                            var value = (worksheet.Cells[i, j] as Excel.Range).Value;
                            bool nullflg = value == null;
                            if (!nullflg)
                            {
                                string vstr = value.ToString();
                                vstr = vstr.Trim();
                                if (string.IsNullOrEmpty(vstr))
                                {
                                    nullflg = true;
                                }
                                else
                                {
                                    if (strflg)
                                    {
                                        vstr = string.Format("\"{0}\"", vstr);
                                    }
                                    sb2.AppendFormat(formatStr2, sname, vstr);
                                }
                            }
                            if (nullflg)
                            {
                                // null cell
                                if (strflg)
                                {
                                    // string type?
                                    sb2.AppendFormat(formatStr2, sname, "\"\"");
                                }
                                else
                                {
                                    sb2.AppendFormat(formatStr2, sname, 0);
                                }
                            }
                            if (j != columnCount)
                            {
                                sb2.Append(',');
                            }
                        }

                        sb.Append(formatStr.Replace(pairStr, sb2.ToString()));
                        if (i != rowCount)
                        {
                            sb.AppendLine();
                        }
                    }
                    s = s.Replace(matchAllStr, sb.ToString());
                }
                else
                {
                    MessageBox.Show("#PAIR_DEF# 匹配失败！");
                    return;
                }
            }
            else
            {
                MessageBox.Show("#DATA_DEF# 匹配失败！");
                return;
            }
            Save2File(worksheet.Name + ".cs", s);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            // export csv
            Excel.Worksheet worksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            int rowCount, columnCount;
            CheckDataCount(worksheet, out rowCount, out columnCount);
            if (rowCount < 3)
            {
                MessageBox.Show("至少需要3行数据!（变量名,类型,数据)");
                return;
            }
            StringBuilder sb = new StringBuilder();
            for (int i = 1; i <= rowCount; i++)
            {
                Excel.Range range = worksheet.UsedRange.Rows[i];
                for (int j = 1; j <= columnCount; j++)
                {
                    Excel.Range cell = range.Columns[j];
                    if (cell.Value != null)
                    {
                        sb.Append(cell.Value.ToString().Trim());
                    }
                    if (j != columnCount)
                    {
                        sb.Append(',');
                    }
                }
                if (i != rowCount)
                {
                    sb.Append("\t\n");
                }
            }
            Save2File(worksheet.Name + ".csv", sb.ToString());
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = exportPath;
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                exportPath = dialog.FileName;
                StoreInRegistry("ExportPath", exportPath);

                MessageBox.Show("设置工作目录: " + exportPath);
            }
        }
    }
}
