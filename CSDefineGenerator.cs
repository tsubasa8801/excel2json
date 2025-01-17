﻿using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;

namespace excel2json
{
    /// <summary>
    /// 根据表头，生成C#类定义数据结构
    /// 表头使用三行定义：字段名称、字段类型、注释
    /// </summary>
    class CSDefineGenerator
    {
        struct FieldDef
        {
            public string name;
            public string type;
            public string comment;
        }

        string mCode;

        public string code {
            get {
                return this.mCode;
            }
        }

        const string CST_Suffix = "Config";

        private string UpperFirstChar(string s)
        {
            char[] a = s.ToCharArray();
            a[0] = char.ToUpper(a[0]);
            return new string(a);
        }

        public CSDefineGenerator(string excelName, ExcelLoader excel, string excludePrefix, string namespaceStr)
        {
            //-- 创建代码字符串
            StringBuilder sb = new StringBuilder();
            //文件头注释
            sb.AppendLine("//");
            sb.AppendLine("// Auto Generated Code By excel2json");
            sb.AppendLine("// https://neil3d.gitee.io/coding/excel2json.html");
            sb.AppendLine("// 1. 每个 Sheet 形成一个 Struct 定义, Sheet 的名称作为 Struct 的名称");
            sb.AppendLine("// 2. 表格约定：第一行是变量名称，第二行是变量类型");
            sb.AppendLine();
            sb.AppendFormat("// Generate From {0}.xlsx", excelName);
            sb.AppendLine();
            sb.AppendLine();

            //添加using
            sb.AppendLine("using System.Collections.Generic;");
            sb.AppendLine();
            sb.AppendLine();

            string tabPrex = string.Empty;

            //添加namespace
            if (string.IsNullOrEmpty(namespaceStr) == false)
            {
                sb.AppendLine("namespace " + namespaceStr + "\r\n{");
                tabPrex = "\t";
            }
     

            for (int i = 0; i < excel.Sheets.Count; i++)
            {
                DataTable sheet = excel.Sheets[i];
                sb.Append(_exportSheet(sheet, excludePrefix, tabPrex));
            }

            string className = UpperFirstChar(excelName);

            sb.Append(_exportTableClass(excel.Sheets, className, excludePrefix, tabPrex));

            //end of namespace     
            if (string.IsNullOrEmpty(namespaceStr) == false)
            {
                sb.Append('}');
                sb.AppendLine();
            }           

            sb.AppendLine();
            sb.AppendLine("// End of Auto Generated Code");

            mCode = sb.ToString();
        }

        private string _exportTableClass(DataTableCollection dtc, string tableName, string excludePrefix, string tabPrex)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("{1}public class {0}\r\n{1}{{", tableName, tabPrex);
            sb.AppendLine();

            foreach (DataTable sheet in dtc)
            {
                string sheetName = sheet.TableName;
                if (excludePrefix.Length > 0 && sheetName.StartsWith(excludePrefix))
                    continue;

                string className = sheet.TableName + CST_Suffix;

                sb.AppendFormat("{2}\tpublic List<{0}> {1} = new List<{0}>();", className, sheet.TableName, tabPrex);
                sb.AppendLine();
            }

            sb.AppendFormat("{0}}}", tabPrex);
            sb.AppendLine();
            return sb.ToString();

        }

        private string _exportSheet(DataTable sheet, string excludePrefix, string tabPrex)
        {
            if (sheet.Columns.Count < 0 || sheet.Rows.Count < 2)
                return "";

            string sheetName = sheet.TableName;
            if (excludePrefix.Length > 0 && sheetName.StartsWith(excludePrefix))
                return "";

            // get field list
            List<FieldDef> fieldList = new List<FieldDef>();
            DataRow typeRow = sheet.Rows[0];
            DataRow commentRow = sheet.Rows[1];

            foreach (DataColumn column in sheet.Columns)
            {
                // 过滤掉包含指定前缀的列
                string columnName = column.ToString();
                if (excludePrefix.Length > 0 && columnName.StartsWith(excludePrefix))
                    continue;

                string fieldType = typeRow[column].ToString();

                if (string.IsNullOrEmpty(fieldType))
                {
                    continue;
                }

                FieldDef field;
                field.name = column.ToString();
                field.type = fieldType;
                field.comment = commentRow[column].ToString();

                fieldList.Add(field);
            }

            string className = sheet.TableName + CST_Suffix;

            // export as string
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("{1}public class {0}\r\n{1}{{", className, tabPrex);           
            sb.AppendLine();

            foreach (FieldDef field in fieldList)
            {
                sb.AppendFormat("{3}\tpublic {0} {1}; // {2}", field.type, field.name, field.comment, tabPrex);
                sb.AppendLine();
            }

            sb.AppendFormat("{0}}}", tabPrex);
            sb.AppendLine();
            sb.AppendLine();
            return sb.ToString();
        }

        public void SaveToFile(string filePath, Encoding encoding)
        {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mCode);
            }
        }
    }
}
