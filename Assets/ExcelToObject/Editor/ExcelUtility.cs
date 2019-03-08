using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using Excel;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using System.Text;
using System.Reflection;
using System.Reflection.Emit;
using System;

public class ExcelUtility
{
    /*
     * 自定义结构
     * 第一行：Contains(c) = 客户端 Contains(s) = 服务器 null = 备注
     * 第二行：注释
     * 第三行：字段名称
     * 第四行：字段类型
     * 第五行：数据起始行   
    */
    private const int CLIENT_SERVER_ROW = 0;
    private const int NAME_ROW = 2;
    private const int TYPE_ROW = 3;
    private const int DATA_START_ROW = 4;

    // 程序集
    private const string ASSEMBLY = "Assembly-CSharp";

    // 表格数据集合
    private DataSet mResultSet;
    // 文件名称
    private string fileName;

	public ExcelUtility (string excelFile)
	{
        fileName = Path.GetFileNameWithoutExtension(excelFile);
        FileStream stream = File.Open (excelFile, FileMode.Open, FileAccess.Read);
        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader (stream);
        mResultSet = excelReader.AsDataSet ();
	}

	/// <summary>
	/// 转换为Json
	/// </summary>
	public void ConvertToJson (string writePath, Encoding encoding)
	{
		//判断Excel文件中是否存在数据表
		if (mResultSet.Tables.Count < 1)
			return;

		//默认读取第一个数据表
		DataTable sheet = mResultSet.Tables [0];

		//判断数据表内是否存在数据
		if (sheet.Rows.Count <= DATA_START_ROW)
			return;

		//读取数据表行数和列数
		int rowCount = sheet.Rows.Count;
		int colCount = sheet.Columns.Count;

		//准备一个列表存储整个表的数据
		List<Dictionary<string, object>> table = new List<Dictionary<string, object>> ();

		//读取数据
		for (int i = DATA_START_ROW; i < rowCount; i++) {
			//准备一个字典存储每一行的数据
			Dictionary<string, object> row = new Dictionary<string, object> ();
			for (int j = 0; j < colCount; j++) {
				//读取第1行数据作为表头字段
				string field = sheet.Rows [NAME_ROW] [j].ToString ();
				//Key-Value对应
				row [field] = sheet.Rows [i] [j];
			}

			//添加到表数据中
			table.Add (row);
		}

		//生成Json字符串
		string json = JsonConvert.SerializeObject (table, Newtonsoft.Json.Formatting.Indented);


	    /*絮大王添加代码：去掉所有的反斜杠+支持数组*/
	    json = JsonSupportArray(json);


		//写入文件
		using (FileStream fileStream=new FileStream(writePath, FileMode.Create,FileAccess.Write)) {
			using (TextWriter textWriter = new StreamWriter(fileStream, encoding)) {
				textWriter.Write (json);
			}
		}
	}

	/// <summary>
	/// 转换为CSV
	/// </summary>
	public void ConvertToCSV (string writePath, Encoding encoding)
	{
		//判断Excel文件中是否存在数据表
		if (mResultSet.Tables.Count < 1)
			return;

		//默认读取第一个数据表
		DataTable sheet = mResultSet.Tables [0];

		//判断数据表内是否存在数据
		if (sheet.Rows.Count <= DATA_START_ROW)
			return;

		//读取数据表行数和列数
		int rowCount = sheet.Rows.Count;
		int colCount = sheet.Columns.Count;

		//创建一个StringBuilder存储数据
		StringBuilder stringBuilder = new StringBuilder ();

		//读取数据
		for (int i = 0; i < rowCount; i++) {
			for (int j = 0; j < colCount; j++) {
				//使用","分割每一个数值
				stringBuilder.Append (sheet.Rows [i] [j] + ",");
			}
			//使用换行符分割每一行
			stringBuilder.Append ("\r\n");
		}

		//写入文件
		using (FileStream fileStream = new FileStream(writePath, FileMode.Create, FileAccess.Write)) {
			using (TextWriter textWriter = new StreamWriter(fileStream, encoding)) {
				textWriter.Write (stringBuilder.ToString ());
			}
		}

	}

    /// <summary>
    /// 转换为Xml
    /// </summary>
    public void ConvertToXml (string writePath)
	{
		//判断Excel文件中是否存在数据表
		if (mResultSet.Tables.Count < 1)
			return;

		//默认读取第一个数据表
		DataTable sheet = mResultSet.Tables [0];

		//判断数据表内是否存在数据
		if (sheet.Rows.Count <= DATA_START_ROW)
			return;

		//读取数据表行数和列数
		int rowCount = sheet.Rows.Count;
		int colCount = sheet.Columns.Count;

		//创建一个StringBuilder存储数据
		StringBuilder stringBuilder = new StringBuilder ();
		//创建Xml文件头
		stringBuilder.Append ("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
		stringBuilder.Append ("\r\n");
		//创建根节点
		stringBuilder.Append ("<Table>");
		stringBuilder.Append ("\r\n");
		//读取数据
		for (int i = DATA_START_ROW; i < rowCount; i++) {
			//创建子节点
			stringBuilder.Append ("  <Row>");
			stringBuilder.Append ("\r\n");
			for (int j = 0; j < colCount; j++) {
				stringBuilder.Append ("   <" + sheet.Rows [NAME_ROW] [j].ToString () + ">");
				stringBuilder.Append (sheet.Rows [i] [j].ToString ());
				stringBuilder.Append ("</" + sheet.Rows [NAME_ROW] [j].ToString () + ">");
				stringBuilder.Append ("\r\n");
			}
			//使用换行符分割每一行
			stringBuilder.Append ("  </Row>");
			stringBuilder.Append ("\r\n");
		}
		//闭合标签
		stringBuilder.Append ("</Table>");
		//写入文件
		using (FileStream fileStream = new FileStream(writePath, FileMode.Create, FileAccess.Write)) {
			using (TextWriter textWriter = new StreamWriter(fileStream,Encoding.GetEncoding("utf-8"))) {
				textWriter.Write (stringBuilder.ToString ());
			}
		}
	}

    /// <summary>
    /// 转换为ScriptableObject
    /// </summary>
    public void ConvertToScriptableObject(string writePath)
    {
        //判断Excel文件中是否存在数据表
        if (mResultSet.Tables.Count < 1)
            return;

        //默认读取第一个数据表
        DataTable sheet = mResultSet.Tables[0];

        //判断数据表内是否存在数据
        if (sheet.Rows.Count <= DATA_START_ROW)
            return;

        var tabName = string.Format("Tab_{0}", fileName);
        var recordName = string.Format("Record_{0}", fileName);

        //数据赋值
        var assembly = Assembly.Load(ASSEMBLY);
        List<object> recordList = new List<object>();
        for (int i = DATA_START_ROW; i < sheet.Rows.Count; i++)
        {
            var recordType = assembly.GetType(recordName);
            var obj = recordType.GetConstructor(Type.EmptyTypes).Invoke(null);
            for (int j = 0; j < sheet.Columns.Count; j++)
            {
                var cs = sheet.Rows[CLIENT_SERVER_ROW][j];
                if (cs == DBNull.Value)
                {
                    continue;
                }
                var type = sheet.Rows[TYPE_ROW][j].ToString();
                var name = sheet.Rows[NAME_ROW][j].ToString();
                var value = sheet.Rows[i][j];
                var field = recordType.GetField(name);
                switch (type)
                {
                    case "int":
                        field.SetValue(obj, value == DBNull.Value ? 0 : Convert.ChangeType(value, field.FieldType));
                        break;
                    case "string":
                        field.SetValue(obj, value == DBNull.Value ? null : (string)value);
                        break;
                    case "float":
                        field.SetValue(obj, value == DBNull.Value ? 0f : (float)value);
                        break;
                    default:
                        Debug.LogError(string.Format("{0}.xlsx文件配置了不支持的数据类型:{1}", fileName, type));
                        break;
                }
            }
            recordList.Add(obj);
        }
        var tabType = assembly.GetType(tabName);
        var scriptableObject = ScriptableObject.CreateInstance(tabType);
        var soType = scriptableObject.GetType();
        var constructor = soType.GetConstructor(new Type[] { typeof(List<object>) });
        var tabObj = constructor.Invoke(new object[] { recordList }) as ScriptableObject;
        UnityEditor.AssetDatabase.CreateAsset(tabObj, string.Format("{0}/{1}.asset", "Assets",tabName));
    }

    /// <summary>
    /// 生成cs文件 继承ScriptableObject
    /// </summary>
    public void GenerateCSFile(string writePath)
    {
        //判断Excel文件中是否存在数据表
        if (mResultSet.Tables.Count < 1)
            return;

        //默认读取第一个数据表
        DataTable sheet = mResultSet.Tables[0];

        //判断数据表内是否存在数据
        if (sheet.Rows.Count <= DATA_START_ROW)
            return;
           
        var template =

        #region 类文件格式
@"using UnityEngine;
using System;
using System.Collections.Generic;

[Serializable]
public class @ClassName_Record : RecordBasse
{
    @body
}

public class @ClassName_Tab : TabBase
{
    private static string fileName = ""@ClassName_Tab"";
    private Dictionary<int, @ClassName_Record> mDic;
    [SerializeField]
    private List<@ClassName_Record> mList;

    public @ClassName_Tab(List<object> objList)
    {
        mList = new List<@ClassName_Record>();
        foreach (var item in objList)
        {
            mList.Add(item as @ClassName_Record);
        }
    }

    public static List<@ClassName_Record> GetRecordList()
    {
        return TabMgr.GetTab<@ClassName_Tab>(fileName).mList;
    }

    public static @ClassName_Record GetRecord(int tag)
    {
        @ClassName_Tab tab = TabMgr.GetTab<@ClassName_Tab>(fileName);
        if(null == tab.mDic)
        {
            tab.mDic = new Dictionary<int, @ClassName_Record>();
            foreach (var item in tab.mList)
            {
                tab.mDic.Add(item.tag, item);
            }
        }

        @ClassName_Record record;
        if(tab.mDic.TryGetValue(tag,out record))
        {
            return record; 
        }
        return null;
    }
}";
        #endregion

        //读取数据
        var body = new StringBuilder();
        for (int j = 0; j < sheet.Columns.Count; j++)
        {
            var cs = sheet.Rows[CLIENT_SERVER_ROW][j];
            if (cs == DBNull.Value)
            {
                continue;
            }

            var type = sheet.Rows[TYPE_ROW][j];
            if (type == DBNull.Value)
            {
                Debug.LogError(string.Format("{0}.xlsx {1}行{2}列配置错误...", fileName, TYPE_ROW + 1, j + 1));
                return;
            }

            var name = sheet.Rows[NAME_ROW][j];
            if (name == DBNull.Value)
            {
                Debug.LogError(string.Format("{0}.xlsx {1}行{2}列配置错误...", fileName, NAME_ROW + 1,j + 1));
                return;
            }

            var file = string.Format("public {0} {1};", type, name);
            if (j > 0) { body.Append("\r\n"); }
            if(body.ToString().IndexOf(file, StringComparison.Ordinal) != -1)
            {
                Debug.LogError(string.Format("{0}.xlsx {1}重命名...", fileName, name));
                return;
            }
            body.Append("\t");
            body.Append(file);
        }

        template = template.Replace("@ClassName_Record", string.Format("Record_{0}", fileName))
                           .Replace("@ClassName_Tab", string.Format("Tab_{0}", fileName))
                           .Replace("@body", body.ToString().Trim());

        writePath = string.Format("{0}/Tab_{1}.cs", writePath, fileName);
        var writer = new StreamWriter(writePath, false, Encoding.UTF8);
        writer.Write(template);
        writer.Close();
    }




    #region 絮大王添加的代码
    /// <summary>
    /// 让Json支持数组
    /// ————————————————————————————————————————————————————
    /// (这个方法是絮大王自己添加的)
    /// 此方法是为了让Excel在转化成Json的时候，支持数组  (格式为["元素1","元素2"])
    /// 
    /// * 为什么要这样做？
    ///   在源代码中，转化后的Json字符串是不支持数组的  (转化后，会变成这样：   "[\"元素1\",\"元素2\"]")
    /// 
    /// * 此方法的做了什么？
    ///   把Json字符串传递过来，就会返回一个支持数组的Json字符串  (格式为["元素1","元素2"])
    /// ——————————————————————————————————————————————————————
    /// </summary>
    /// <param name="jsonContent">json的文本内容</param>
    /// <returns></returns>
    private string JsonSupportArray(string jsonContent)
    {
        //去掉所有的反斜杠  （把"\"替换成""）
        jsonContent = jsonContent.Replace("\\", string.Empty);

        //为了能够支持数组，所以：
        //把所有的"[替换成[   
        //并且把所有的]"替换成]
        jsonContent = jsonContent.Replace("\"[", "[");
        jsonContent = jsonContent.Replace("]\"", "]");

        //把所有的".0" 替换成 ""
        //不然的话，"1"会显示为"1.0"
        jsonContent = jsonContent.Replace(".0,", "");

        return jsonContent;
    }

    #endregion
}

