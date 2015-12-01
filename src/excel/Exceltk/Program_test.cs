using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;

namespace ExcelToolKit
{
	internal class Program
	{
        //const string xls = @"C:\Users\Jennal\Desktop\Work\RDTower.CMD.xlsx2lua\linq\测试.xlsx";
        const string xls = @"C:\Users\Jennal\Desktop\Work\Tmp\xlsx2lua\机器人设置.xlsx";
        const string sheet = "配置";
        //const string sheet = "出怪5-3";
        const string ID = "id";
        //const string ID = "battleNumber";
        const int LINE_HEADER = 1;
        const int LINE_TYPE = 2;
        const int LINE_START = 3;

        const string luaFile = @"C:\Users\Jennal\Desktop\Work\Tmp\xlsx2lua\output.lua";

		private static void Main(string[] args)
		{
            var tmpFileName = xls + ".4linq";
            if (File.Exists(tmpFileName))
            {
                File.Delete(tmpFileName);
            }
            File.Copy(xls, tmpFileName);

            try
            {
                using (FileStream stream = File.Open(tmpFileName, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader excelReader;

                    if (Path.GetExtension(xls) == ".xls")
                    {
                        excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else
                    {
                        excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }

                    DataSet dataSet = excelReader.AsDataSet();
                    DataTable dataTable = dataSet.Tables[sheet];

                    Utils.clearCache();
                    var dict = BuildDict(dataTable);
                    var content = "return " + dict.convertToLua();
                    File.WriteAllText(luaFile, content);
                    excelReader.Close();
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                File.Delete(tmpFileName);
            }
		}

        static Dictionary<string, object> BuildDict(DataTable table)
        {
            Dictionary<string, object> dict = new Dictionary<string, object>();

            var headers = table.Rows[LINE_HEADER].ItemArray.toStringArray();
            var headerTypes = table.Rows[LINE_TYPE].ItemArray.toStringArray();

            for (int i = LINE_START; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i].ItemArray.toStringArray();
                string id = i.ToString();

                for (int j = 0; j < row.Length; j++)
                {
                    var column = row[j];
                    if (headers[j] == ID)
                    {
                        id = column;
                    }
                    if (string.IsNullOrEmpty(id))
                    {
                        break;
                    }

                    string key = string.Format("[{0}].{1}", id, headers[j]);
                    key = key.parseRefKeys(row, headers, headerTypes);
                    key = key.resolveKeys(i);

                    Utils.setKeyType(key, headerTypes[j]);
                    dict.setValue(key, column);
                }

                if (string.IsNullOrEmpty(id))
                {
                    break;
                }
            }

            return dict;
        }
	}

    static class Utils
    {
        static private Dictionary<string, bool> KeyLineInced = new Dictionary<string, bool>();
        static private Dictionary<string, int> KeyCounts = new Dictionary<string, int>();
        static private Dictionary<string, string> KeyTypeDict = new Dictionary<string, string>();
        static public void clearCache()
        {
            KeyLineInced.Clear();
            KeyCounts.Clear();
            KeyTypeDict.Clear();
        }

        static public bool isNumber(this string str)
        {
            Regex reg = new Regex(@"^-?\d+(\.\d+)?$");
            return reg.IsMatch(str);
        }

        static public string[] toStringArray(this Object[] arr)
        {
            return arr.Select(o => o.ToString()).ToArray();
        }

        static public string wrapQuoteWithType(this string value, string type)
        {
            if (type == "string")
            {
                return string.Format(@"""{0}""", value);
            }

            return value;
        }

        static public string toLua(string name, string type, string value)
        {
            //solve type
            var types = type.Split("|".ToCharArray());

            //如果type存在 |Y，表示是Key，返回空
            if (types.Length > 1 && types[1] == "K")
            {
                return "";
            }

            //如果type存在 |Y，表示可选，如果没有数据则返回空
            if (types.Length > 1 && types[1] == "Y")
            {
                if (types[0] == "string" && string.IsNullOrEmpty(value))
                {
                    return "";
                }
                else if (types[0] == "number" && Double.Parse(value) == 0)
                {
                    return "";
                }
            }

            //solve value
            value = value.wrapQuoteWithType(types[0]);

            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            //solve key
            Regex reg = new Regex(@"^[^a-zA-Z_\[]");
            if (reg.IsMatch(name))
            {
                name = string.Format(@"[""{0}""]", name);
            }

            return string.Format("{0}={1}", name, value);
        }

        static public void setKeyType(string key, string type)
        {
            KeyTypeDict[key] = type;
        }

        static public string getKeyType(string key)
        {
            if (!KeyTypeDict.ContainsKey(key))
            {
                return "";
            }
            return KeyTypeDict[key];
        }

        static public string parseRefKeys(this string key, string[] row, string[] headers, string[] headerTypes)
        {
            //解析引用的key
            Regex reg = new Regex(@"\[\$(.*?)\]");
            key = reg.Replace(key, new MatchEvaluator(match =>
            {
                if (match.Groups.Count < 2) return match.ToString();

                var idName = match.Groups[1].Value;
                for (int i = 0; i != headers.Length; ++i)
                {
                    if (idName == headers[i])
                    {
                        var types = headerTypes[i].Split("|".ToCharArray());
                        var k = row[i].wrapQuoteWithType(types[0]);
                        return string.Format("[{0}]", k);
                    }
                }

                return match.ToString();
            }));

            return key;
        }

        static public string resolveKeys(this string key, int line)
        {
            var result = key;
            var start = result.IndexOf("[]", 0);
            while (start >= 0)
            {
                var k = result.Substring(0, start);
                if (!KeyCounts.ContainsKey(k))
                {
                    KeyCounts.Add(k, 0);
                }

                var lineIncKey = string.Format("{0}.{1}", line, k);
                if (!KeyLineInced.ContainsKey(lineIncKey))
                {
                    KeyCounts[k] = KeyCounts[k] + 1;
                    KeyLineInced[lineIncKey] = true;
                }
                result = result.Insert(start + 1, KeyCounts[k].ToString());

                start = result.IndexOf("[]", start);
            }

            return result;
        }

        static public void setValue(this Dictionary<string, object> dict, string key, string val)
        {
            //过滤空的数据，用于第N行设置子节点信息，某些外层字段可以忽略不填
            if (string.IsNullOrEmpty(val))
            {
                return;
            }

            var keys = key.Split(".".ToCharArray());
            var d = dict;
            for (int i = 0; i < keys.Length - 1; i++)
            {
                var k = keys[i];
                if (!d.ContainsKey(k)) d.Add(k, new Dictionary<string, object>());

                d = d[k] as Dictionary<string, object>;
            }

            d[keys[keys.Length - 1]] = val;
        }

        static public string convertToLua(this Dictionary<string, object> dict, string parentKey = null)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("{");

            foreach (var item in dict)
            {
                string key = item.Key;
                string value;

                if (!string.IsNullOrEmpty(parentKey))
                {
                    key = parentKey + "." + key;
                }

                if (item.Value is Dictionary<string, object>)
                {
                    value = (item.Value as Dictionary<string, object>).convertToLua(key);
                }
                else
                {
                    value = item.Value as string;
                }

                value = Utils.toLua(
                    item.Key,
                    Utils.getKeyType(key),
                    value
                );
                if (!string.IsNullOrEmpty(value))
                {
                    sb.Append(value);
                    sb.Append(",");
                }
            }
            //remove last ","
            if (sb.Length > 1) // ==> {
            {
                sb.Remove(sb.Length - 1, 1);
            }
            sb.Append("}");

            if (sb.Length == 2) // ==> {} Empty
            {
                return "";
            }

            return sb.ToString();
        }
    }
}