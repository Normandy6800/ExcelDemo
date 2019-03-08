using UnityEngine;
using System.Collections.Generic;

public class RecordBasse
{
}

public class TabBase : ScriptableObject
{
}

public class TabMgr
{
    private static Dictionary<string, TabBase> tabDic = new Dictionary<string, TabBase>();

    public static T GetTab<T>(string key) where T : TabBase
    {
        TabBase tab;
        if (!tabDic.TryGetValue(key,out tab))
        {
            tab = UnityEditor.AssetDatabase.LoadAssetAtPath<T>(GetRecordPath(key));
            tabDic.Add(key, tab);
        }
        return tab as T;
    }

    public static void Clear()
    {
        tabDic.Clear();
    }

    public static string GetRecordPath(string name)
    {
        return string.Format("Assets/{0}.asset", name);
    }
}

