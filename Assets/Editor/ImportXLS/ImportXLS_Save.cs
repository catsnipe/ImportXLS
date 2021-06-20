using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using UnityEditor;

public partial class ImportXLS : EditorWindow
{
    /// <summary>
    /// シートを保存
    /// </summary>
    void saveSheet(List<SheetEntity> list)
    {
        int cnt = 0;
        
        try
        {
            int max = list.Count*3;
            foreach (SheetEntity ent in list)
            {
                setNames_Table_Accessor_Data(ent);

                string sheetName = ent.Sheet.SheetName;
                if (cancelableProgressBar(cnt++, max, ent.TableName) == true)
                {
                    break;
                }
                saveTableClass(ent, classDir);

                if (createAccess == true)
                {
                    if (cancelableProgressBar(cnt++, max, ent.AccessorName) == true)
                    {
                        break;
                    }
                    saveTableAccess(ent, classDir);
                }

                if (cancelableProgressBar(cnt++, max, ent.DataName + IMPORT_FILENAME_SUFFIX) == true)
                {
                    break;
                }
                saveImporter(ent, xlsPath, dataDir);
            }
        }
        finally
        {
            EditorUtility.ClearProgressBar();
        }
    }

    /// <summary>
    /// ScriptableObject の基幹となるクラスを生成し、保存
    /// </summary>
    static bool saveTableClass(SheetEntity report, string classdir)
    {
        // Class は統一クラスとして既に出力されている
        if (report.ClassName != report.Sheet.SheetName)
        {
            return true;
        }

        string text = createTableClass(report);

        if (text != null)
        {
            string classpath = pathCombine(classdir, report.TableName) + ".cs";

            // unix 形式に合わせる
            text = text.Replace("\r\n", "\n");
            File.WriteAllText(classpath, text, Encoding.UTF8);
            AssetDatabase.ImportAsset(classpath);
        }

        return text != null;
    }

    /// <summary>
    /// ScriptableObject をアクセスするシングルトンクラスを生成し、保存
    /// </summary>
    static bool saveTableAccess(SheetEntity report, string classdir)
    {
        // テーブルがない
        if (report.Classes.Count == 0)
        {
            return true;
        }
        // Class は統一クラスとして既に出力されている
        if (report.ClassName != report.Sheet.SheetName)
        {
            return true;
        }

        string text = createTableAccess(report);

        if (text != null)
        {
            string classpath = pathCombine(classdir, report.AccessorName) + ".cs";
            // unix 形式に合わせる
            text = text.Replace("\r\n", "\n");
            File.WriteAllText(classpath, text, Encoding.UTF8);
            AssetDatabase.ImportAsset(classpath);
        }

        return text != null;
    }

    /// <summary>
    /// ScriptableObject を自動生成する Editor クラスを生成し、保存
    /// </summary>
    static bool saveImporter(SheetEntity report, string xlspath, string datadir)
    {
        // テーブルがない
        if (report.Classes.Count == 0)
        {
            return true;
        }

        string text = createImporter(report, xlspath, datadir);

        if (text == null)
        {
            return false;
        }

        string dataname = report.DataName;
        string workdir  = pathCombine(searchImportXLSDirectory(), IMPORT_DIRECTORY);
        string outpath  = pathCombine(workdir, dataname + IMPORT_FILENAME_SUFFIX + ".cs");

        if (workdir != null)
        {
            CompleteDirectory(workdir);
            // unix 形式に合わせる
            text = text.Replace("\r\n", "\n");
            File.WriteAllText(outpath, text, Encoding.UTF8);
            AssetDatabase.ImportAsset(outpath);
        }
        return true;
    }

}
