using UnityEngine;
using System;
using System.Collections.Generic;

[Serializable]
public class Record_CardLevel : RecordBasse
{
    public int tag;
	public int exp;
}

public class Tab_CardLevel : TabBase
{
    private static string fileName = "Tab_CardLevel";
    private Dictionary<int, Record_CardLevel> mDic;
    [SerializeField]
    private List<Record_CardLevel> mList;

    public Tab_CardLevel(List<object> objList)
    {
        mList = new List<Record_CardLevel>();
        foreach (var item in objList)
        {
            mList.Add(item as Record_CardLevel);
        }
    }

    public static List<Record_CardLevel> GetRecordList()
    {
        return TabMgr.GetTab<Tab_CardLevel>(fileName).mList;
    }

    public static Record_CardLevel GetRecord(int tag)
    {
        Tab_CardLevel tab = TabMgr.GetTab<Tab_CardLevel>(fileName);
        if(null == tab.mDic)
        {
            tab.mDic = new Dictionary<int, Record_CardLevel>();
            foreach (var item in tab.mList)
            {
                tab.mDic.Add(item.tag, item);
            }
        }

        Record_CardLevel record;
        if(tab.mDic.TryGetValue(tag,out record))
        {
            return record; 
        }
        return null;
    }
}