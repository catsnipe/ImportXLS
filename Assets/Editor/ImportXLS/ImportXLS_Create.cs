using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;
using UnityEditor;
using UnityEngine;

public partial class ImportXLS : EditorWindow
{
    /// <summary>
    /// ScriptableObject の基幹となるクラスを生成
    /// </summary>
    static string createTableClass(SheetEntity report)
    {
        ISheet                        sheet         = report.Sheet;
        Dictionary<string, ClassInfo> classes       = report.Classes;
        Dictionary<string, EnumInfo>  enums         = report.Enums;
        Dictionary<string, ConstInfo> consts        = report.Consts;
        string                        tablename     = report.TableName;

        string                        import_dir    = searchImportXLSDirectory();
        string                        template_file = $"{nameof(ImportXLSTemplate_Class)}.cs";
        string                        text          = null;

        if (import_dir != null)
        {
            text = File.ReadAllText(pathCombine(import_dir, template_file), Encoding.UTF8);
        }
        if (text == null)
        {
            // テンプレートが見つからない
            dialog_error(MSG_CLASSTMPL_NOTFOUND, template_file);
            return null;
        }

        // enum
        var sb_enum        = new StringBuilder();
        var sb_global_enum = new StringBuilder();
        foreach (var pair in enums)
        {
            StringBuilder sb;
            string        indent = "";

            if (pair.Key.IndexOf(TRIGGER_GLOBAL_ENUM) == 0)
            {
                sb = sb_global_enum;
            }
            else
            {
                sb = sb_enum;
                indent = "\t";
            }
            if (string.IsNullOrEmpty(pair.Value.Comment) == false)
            {
                addCommentText(sb, pair.Value.Comment, indent.Length == 0 ? 0 : 1);
            }
            sb.AppendLine($"{indent}public enum {pair.Value.GroupName}");
            sb.AppendLine($"{indent}{{");
            foreach (var member in pair.Value.Members)
            {
                Member row = member.Value;

                addCommentText(sb, row.Comment, indent.Length == 0 ? 1 : 2);
                if (string.IsNullOrEmpty(row.Suffix) == true)
                {
                    sb.AppendLine($"{indent}\t{member.Key},");
                }
                else
                {
                    sb.AppendLine($"{indent}\t{member.Key} = {row.Suffix},");
                }
            }
            sb.AppendLine($"{indent}}};");
            sb.AppendLine( "");
        }

        // const
        var sb_const = new StringBuilder();
        foreach (var pair in consts)
        {
            StringBuilder sb     = sb_const;
            string        indent = "\t";

            foreach (var member in pair.Value.Members)
            {
                if (string.IsNullOrEmpty(member.Value.Comment) == false)
                {
                    addCommentText(sb, member.Value.Comment, indent.Length == 0 ? 0 : 1);
                }

                string value = "";
                switch (member.Value.Type)
                {
                    case "string":
                        value = $"\"{member.Value.Suffix}\"";
                        break;
                    case "float":
                        value = $"{member.Value.Suffix}f";
                        break;
                    case "int":
                        value = $"{member.Value.Suffix}";
                        break;
                }
                sb.AppendLine($"{indent}public const {member.Value.Type} {member.Key} = {value};");
            }
        }

        var sb_class      = new StringBuilder();
        var sb_index      = new StringBuilder();
        var sb_index_find = new StringBuilder();
        var tableNames    = new Dictionary<string, string>();
        int padding       = 0;

        foreach (var entity in sheetList)
        {
            if (tableNames.ContainsKey(entity.ClassName) == false)
            {
                tableNames.Add(entity.ClassName, entity.TableName);
            }
        }

        foreach (var pair in classes)
        {
            foreach (var member in pair.Value.Members)
            {
                if (padding < member.Value.Type.Length)
                {
                    padding = member.Value.Type.Length;
                }
            }
        }
        foreach (var pair in classes)
        {
            if (string.IsNullOrEmpty(pair.Value.Comment) == false)
            {
                addCommentText(sb_class, pair.Value.Comment, 1);
            }
            sb_class.AppendLine( "\t[System.Serializable]");
            sb_class.AppendLine($"\tpublic class {pair.Key}");
            sb_class.AppendLine( "\t{");
            foreach (var member in pair.Value.Members)
            {
                Member row  = member.Value;
                string type = row.Type;

                if (row.IsEnum == true && type.IndexOf('.') >= 0)
                {
                    string classname = type.Substring(0, type.IndexOf('.'));
                    if (tableNames.ContainsKey(classname) == false)
                    {
                        log_error(MSG_NOT_FOUND_ENUMTBL, report.Sheet.SheetName);
                        continue;
                    }
                    type = type.Replace(classname, tableNames[classname]);
                }

                addCommentText(sb_class, row.Comment, 2);
                if (row.Suffix == null)
                {
                    sb_class.AppendLine($"\t\tpublic {type.PadRight(padding)} {member.Key};");
                }
                else
                {
                    sb_class.AppendLine($"\t\tpublic {type.PadRight(padding)} {member.Key} {row.Suffix};");
                }

                if (row.Indexer == true)
                {
                    string name = $"{member.Key}Rows";

                    sb_index.AppendLine($"\tDictionary<{type}, Row> {name};");
                    sb_index_find.AppendLine( "\t/// <summary>");
                    sb_index_find.AppendLine($"\t/// {member.Key}");
                    sb_index_find.AppendLine( "\t/// </summary>");
                    sb_index_find.AppendLine($"\tpublic Row FindRowBy{member.Key}({type} val)");
                    sb_index_find.AppendLine( "\t{");
                    sb_index_find.AppendLine($"\t\tif ({name} == null)");
                    sb_index_find.AppendLine( "\t\t{");
                    sb_index_find.AppendLine($"\t\t\t{name} = new Dictionary<{type}, Row>();");
                    sb_index_find.AppendLine($"\t\t\tRows.ForEach( (row) => {name}.Add(row.{member.Key}, row) );");
                    sb_index_find.AppendLine( "\t\t}");
                    sb_index_find.AppendLine($"\t\tif ({name}.ContainsKey(val) == false)");
                    sb_index_find.AppendLine( "\t\t{");
                    sb_index_find.AppendLine( "\t\t\tDebug.LogError(\"cannot find: {val}\");");
                    sb_index_find.AppendLine( "\t\t\treturn null;");
                    sb_index_find.AppendLine( "\t\t}");
                    sb_index_find.AppendLine($"\t\treturn {name}[val];");
                    sb_index_find.AppendLine( "\t}");
                    sb_index_find.AppendLine( "");
                }
            }
            sb_class.AppendLine( "\t};");
            sb_class.AppendLine( "");
        }

        int start;
        int end;

        // 最初からある Row を消す
        start = text.IndexOf(CLASSTMPL_CLASS_SIGN) + CLASSTMPL_CLASS_SIGN.Length + "\r\n".Length;
        end   = text.IndexOf(CLASSTMPL_CLASS_ENDSIGN);
        text = text.Remove(start, end - start);

        text = text.Replace(CLASSTMPL_GENUM_ENDSIGN + "\r\n", sb_global_enum.ToString());
        text = text.Replace(CLASSTMPL_ENUM_ENDSIGN + "\r\n", sb_enum.ToString());
        text = text.Replace(CLASSTMPL_CONST_ENDSIGN + "\r\n", sb_const.ToString());
        text = text.Replace(CLASSTMPL_CLASS_ENDSIGN + "\r\n", sb_class.ToString());
        text = text.Replace(CTMPL_INDEX_ENDSIGN + "\r\n", sb_index.ToString());
        text = text.Replace(CTMPL_INDEX_FIND_ENDSIGN + "\r\n", sb_index_find.ToString());
        text = text.Replace(CLASSTMPL_GENUM_SIGN + "\r\n", "");
        text = text.Replace(CLASSTMPL_ENUM_SIGN + "\r\n", "");
        text = text.Replace(CLASSTMPL_CONST_SIGN + "\r\n", "");
        text = text.Replace(CLASSTMPL_CLASS_SIGN + "\r\n", "");
        text = text.Replace(CTMPL_INDEX_SIGN + "\r\n", "");
        text = text.Replace(CTMPL_INDEX_FIND_SIGN + "\r\n", "");
        if (report.Classes.Count == 0 && sb_enum.Length == 0 && sb_const.Length == 0)
        {
            // テーブルがない場合、余計なコードを一切消す
            start = text.IndexOf(CLASSTMPL_TABLE_SIGN) + CLASSTMPL_TABLE_SIGN.Length + "\r\n".Length;
            end   = text.IndexOf(CLASSTMPL_TABLE_ENDSIGN);
            text = text.Remove(start, end - start);
        }
        else
        if (report.Classes.Count == 0)
        {
            // Row がない場合、余計なコードを一切消す
            start = text.IndexOf(CLASSTMPL_CODE_SIGN) + CLASSTMPL_CODE_SIGN.Length + "\r\n".Length;
            end   = text.IndexOf(CLASSTMPL_CODE_ENDSIGN);
            text = text.Remove(start, end - start);
        }
        text = text.Replace(CLASSTMPL_TABLE_SIGN + "\r\n", "");
        text = text.Replace(CLASSTMPL_TABLE_ENDSIGN + "\r\n", "");
        text = text.Replace(CLASSTMPL_CODE_SIGN + "\r\n", "");
        text = text.Replace(CLASSTMPL_CODE_ENDSIGN + "\r\n", "");
        text = text.Replace("\t", "    ");
        text = text.Replace(nameof(ImportXLSTemplate_Class), tablename);

        return text;
    }

    /// <summary>
    /// ScriptableObject をアクセスするシングルトンクラスを生成
    /// </summary>
    static string createTableAccess(SheetEntity report)
    {
        ISheet                        sheet         = report.Sheet;
        Dictionary<string, ClassInfo> classes       = report.Classes;
        string                        accessorname  = report.AccessorName;
        string                        tablename     = report.TableName;

        string import_dir    = searchImportXLSDirectory();
        string template_file = $"{nameof(ImportXLSTemplate_Access)}.cs";
        string text          = null;

        if (import_dir != null)
        {
            text = File.ReadAllText(pathCombine(import_dir, template_file), Encoding.UTF8);
        }
        if (text == null)
        {
            // テンプレートが見つからない
            dialog_error(MSG_CLASSTMPL_NOTFOUND, template_file);
            return null;
        }

        var sb_index_find = new StringBuilder();
        foreach (var pair in classes)
        {
            foreach (var member in pair.Value.Members)
            {
                Member row = member.Value;

                if (row.Indexer == true)
                {
                    string name = $"{member.Key}Rows";

                    sb_index_find.AppendLine( "\t/// <summary>");
                    sb_index_find.AppendLine($"\t/// {member.Key}");
                    sb_index_find.AppendLine( "\t/// </summary>");
                    sb_index_find.AppendLine($"\tpublic static {tablename}.Row FindRowBy{member.Key}({row.Type} val)");
                    sb_index_find.AppendLine( "\t{");
                    sb_index_find.AppendLine($"\t\treturn table.FindRowBy{member.Key}(val);");
                    sb_index_find.AppendLine( "\t}");
                    sb_index_find.AppendLine( "");
                }
            }
        }

        text = text.Replace(nameof(ImportXLSTemplate_Class), tablename);
        text = text.Replace(nameof(ImportXLSTemplate_Access), accessorname);
        text = text.Replace(CTMPL_INDEX_FIND_ENDSIGN + "\r\n", sb_index_find.ToString());
        text = text.Replace(CTMPL_INDEX_FIND_SIGN + "\r\n", "");
        text = text.Replace("\t", "    ");

        return text;
    }

    /// <summary>
    /// ScriptableObject を自動生成する Editor クラスを生成
    /// </summary>
    static string createImporter(SheetEntity report, string xlspath, string datadir)
    {
        ISheet                       sheet         = report.Sheet;
        string[,]                    grid          = report.Grid;
        Dictionary<string, PosIndex> posList       = report.PosList;
        string                       sheetname     = sheet.SheetName;
        string                       classname     = report.ClassName;
        string                       tablename     = report.TableName;
        string                       datapath      = pathCombine(datadir, report.DataName) + ".asset";

        string                       import_dir    = searchImportXLSDirectory();
        string                       template_file = IMPORT_TEMPLATE_FILE;
        string                       text          = null;

        if (import_dir != null)
        {
            text = File.ReadAllText(pathCombine(import_dir, template_file), Encoding.UTF8);
        }
        if (text == null)
        {
            // テンプレートが見つからない
            dialog_error(MSG_CLASSTMPL_NOTFOUND, template_file);
            return null;
        }

        var tableNames = new Dictionary<string, string>();

        foreach (var entity in sheetList)
        {
            if (tableNames.ContainsKey(entity.ClassName) == false)
            {
                tableNames.Add(entity.ClassName, entity.TableName);
            }
        }

        StringBuilder sb_import = new StringBuilder();

        PosIndex id = posList[TRIGGER_ID];

        string preMember = null;

        for (int c = id.C; c < grid.GetUpperBound(1)+1; c++)
        {
            string member  = grid[id.R, c];
            string typestr = grid[id.R+1, c];

            if (string.IsNullOrEmpty(member) == true)
            {
                break;
            }

            if (c > id.C)
            {
                sb_import.AppendLine($"\t\t\t\t\tcell = grid[r, ++c];");
            }

            if (typestr == null)
            {
                typestr = "";
            }

            if (id.R >= 1)
            {
                // クラス宣言の存在するメンバか調査
                string classMember = "";
                int    count_blank = 0;
                for (int r = id.R-1; r > -1; r--)
                {
                    string cell = grid[r, c];
                    if (string.IsNullOrEmpty(cell) == true)
                    {
                        // blank
                        if (++count_blank >= MARGINROW_BETWEEN_CTG)
                        {
                            break;
                        }
                        continue;
                    }
                    else
                    // コメントは無視
                    if (CheckSignComment(cell) == true)
                    {
                        continue;
                    }
                    else
                    {
                        classMember += cell + ".";
                    }
                }
                member = classMember + member;
                // delete '*'
                member = member.Replace(SIGN_INDEXER, "");
            }

            bool   isList  = false;
            bool   isNull  = false;
            string getFunc = null;

            if (typestr.IndexOf(SIGN_LIST) >= 0)
            {
                isList = true;
                typestr = typestr.Replace(SIGN_LIST, "");
            }
            else
            {
                isList = false;
            }

            if (typestr.IndexOf(SIGN_ENUM) >= 0)
            {
                getFunc = "GetEnum  ";
                typestr = typestr.Replace(SIGN_ENUM, "");

                if (typestr.IndexOf(".") >= 0)
                {
                    string clsname = typestr.Substring(0, typestr.IndexOf('.'));
                    if (tableNames.ContainsKey(clsname) == false)
                    {
                        log_error(MSG_NOT_FOUND_ENUMTBL, report.Sheet.SheetName);
                        continue;
                    }
                    typestr = typestr.Replace(clsname, tableNames[clsname]);
                }

                //if (typestr.IndexOf(".") >= 0)
                //{
                //    string[] types = typestr.Split('.');
                //    typestr = $"{types[0]}{CLASS_TABLE_SUFFIX}.{types[1]}";
                //}
            }
            else
            {
                switch (typestr)
                {
                    case "bool":
                        getFunc = "GetBool  ";
                        break;
                    case "int":
                        getFunc = "GetInt   ";
                        break;
                    case "float":
                        getFunc = "GetFloat ";
                        break;
                    case "string":
                        getFunc = "GetString";
                        break;
                    default:
                        if (string.IsNullOrEmpty(typestr) == false)
                        {
                            log($"no importing member was found. type:{typestr} member:{member}");
                        }
                        getFunc = "GetNull  ";
                        isNull  = true;
                        break;
                }
            }

            if (isList == false)
            {
                if (isNull == true)
                {
                    sb_import.AppendLine($"\t\t\t\t\tif (import(r, c, cell, ImportXLS.{getFunc}()) == false) break;");
                }
                else
                {
                    sb_import.AppendLine($"\t\t\t\t\tif (import(r, c, cell, ImportXLS.{getFunc}(cell, out row.{member})) == false) break;");
                }
            }
            else
            {
                if (preMember != member)
                {
                    sb_import.AppendLine($"\t\t\t\t\tif (import(r, c, cell, ImportXLS.{getFunc}(cell, out {typestr} {member})) == false) break;");
                }
                else
                {
                    sb_import.AppendLine($"\t\t\t\t\tif (import(r, c, cell, ImportXLS.{getFunc}(cell, out {member})) == false) break;");
                }
                sb_import.AppendLine($"\t\t\t\t\trow.{member}.Add({member});");
            }

            preMember = member;
        }

        text = text.Replace(IMPORTTMPL_SHEET_NAME, sheetname);
        text = text.Replace(IMPORTTMPL_TABLE_NAME, tablename);
        text = text.Replace(IMPORTTMPL_EXCELL_PATH, xlspath);
        text = text.Replace(IMPORTTMPL_EXPORT_PATH, datapath);
        text = text.Replace(IMPORTTMPL_IMPORT_ROW, sb_import.ToString());
        text = text.Replace("\t", "    ");
        text = text.Replace(nameof(ImportXLSTemplate_Class), classname);

        return text;
    }
}
