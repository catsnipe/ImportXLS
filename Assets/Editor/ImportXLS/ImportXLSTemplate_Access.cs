using System.Collections;
using System.Collections.Generic;
using UnityEngine;

/// <summary>
/// ImportXLSTemplate_Access: ScriptableObject singleton accessor
/// </summary>
public partial class ImportXLSTemplate_Access
{
    /// <summary>
    /// データカウント
    /// </summary>
    public static int Count
    {
        get
        {
            return table.Count;
        }
    }
    /// <summary>
    /// 全行を返す
    /// </summary>
    public static List<ImportXLSTemplate_Class.Row> Rows
    {
        get
        {
            return table.Rows;
        }
    }

    static ImportXLSTemplate_Class table;

    /// <summary>
    /// テーブルを設定
    /// </summary>
    public static void SetTable(Object obj)
    {
        table = obj as ImportXLSTemplate_Class;
    }

    /// <summary>
    /// 行を配列順で取得する
    /// </summary>
    /// <param name="index">配列番号</param>
    public static ImportXLSTemplate_Class.Row GetRow(int index)
    {
        return table.GetRow(index);
    }

//$$REGION INDEX_FIND$$
//$$REGION END_INDEX_FIND$$
}
