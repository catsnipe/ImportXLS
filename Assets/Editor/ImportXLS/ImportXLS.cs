//#define JAPANESE

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Security.Cryptography;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor;
using UnityEngine;
using System.Security.Policy;

public partial class ImportXLS : EditorWindow
{
    /// <summary>
    /// コメントを示すキーワード
    /// </summary>
    static readonly string[] SIGN_COMMENTS
                                          = new string[] { "///", "[[[", "]]]", "[[", "]]", "#" };
    /// <summary>
    /// ID のポジションを示すキーワード
    /// </summary>
    public const string TRIGGER_ID        = "ID";
    /// <summary>
    /// [CLASS] のポジションを示すキーワード
    /// </summary>
    const string TRIGGER_CLASS            = "[CLASS]";
    /// <summary>
    /// [ENUM] のポジションを示すキーワード
    /// </summary>
    const string TRIGGER_ENUM             = "[ENUM]";
    /// <summary>
    /// [CONST] のポジションを示すキーワード
    /// </summary>
    const string TRIGGER_CONST            = "[CONST]";
    /// <summary>
    /// [GLOBAL_ENUM] のポジションを示すキーワード
    /// </summary>
    const string TRIGGER_GLOBAL_ENUM      = "[GLOBAL_ENUM]";
    /// <summary>
    /// 最大行数
    /// </summary>
    const int    ROWS_MAX                 = 2000;
    /// <summary>
    /// 最大列数
    /// </summary>
    const int    COLS_MAX                 = 200;

    const int    MARGINROW_BETWEEN_CTG    = 2;
    const string SIGN_ENUM                = "enum:";
    const string SIGN_LIST                =  "[]";
    const string SIGN_INDEXER             =  "*";
    const string SIGN_COMMA               =  ".";

    const string CLASSTMPL_GENUM_SIGN     = "//$$REGION GLOBAL_ENUM$$";
    const string CLASSTMPL_GENUM_ENDSIGN  = "//$$REGION_END GLOBAL_ENUM$$";
    const string CLASSTMPL_ENUM_SIGN      = "//$$REGION ENUM$$";
    const string CLASSTMPL_ENUM_ENDSIGN   = "//$$REGION_END ENUM$$";
    const string CLASSTMPL_CONST_SIGN     = "//$$REGION CONST$$";
    const string CLASSTMPL_CONST_ENDSIGN  = "//$$REGION_END CONST$$";
    const string CLASSTMPL_CLASS_SIGN     = "//$$REGION CLASS$$";
    const string CLASSTMPL_CLASS_ENDSIGN  = "//$$REGION_END CLASS$$";
    const string CLASSTMPL_CODE_SIGN      = "//$$REGION CODE$$";
    const string CLASSTMPL_CODE_ENDSIGN   = "//$$REGION_END CODE$$";
    const string CLASSTMPL_TABLE_SIGN     = "//$$REGION TABLE$$";
    const string CLASSTMPL_TABLE_ENDSIGN  = "//$$REGION_END TABLE$$";
    const string CTMPL_INDEX_SIGN         = "//$$REGION INDEX$$";
    const string CTMPL_INDEX_ENDSIGN      = "//$$REGION END_INDEX$$";
    const string CTMPL_INDEX_FIND_SIGN    = "//$$REGION INDEX_FIND$$";
    const string CTMPL_INDEX_FIND_ENDSIGN = "//$$REGION END_INDEX_FIND$$";

    const string CLASSTMPL_ROW            = "Row";
    const string IMPORTTMPL_SHEET_NAME    = "$$SHEET_NAME$$";
    const string IMPORTTMPL_TABLE_NAME    = "$$TABLE_NAME$$";
    const string IMPORTTMPL_EXCELL_PATH   = "$$EXCELL_PATH$$";
    const string IMPORTTMPL_EXPORT_PATH   = "$$EXPORT_PATH$$";
    const string IMPORTTMPL_IMPORT_ROW    = "$$IMPORT_ROW$$";

    const string IMPORT_TEMPLATE_FILE     = "ImportXLSTemplate_Import.txt";
    const string IMPORT_DIRECTORY         = "importer/";
    const string IMPORT_FILENAME_SUFFIX   = "_Import";
    const string CLASS_TABLE_SUFFIX       = "_Table";
    const string ASSETS_CLASS             = "Assets/Scripts/";
    const string ASSETS_RESOURCE          = "Assets/Resources/";
    const string PREFIX_TABLE             = "";
    const string PREFIX_ACCESS            = "X_";
    const string PREFIX_DATA              = "Data_";

    const string PREFS_CLASS_DIRECTORY      = ".classdir";
    const string PREFS_DATA_DIRECTORY       = ".datadir";
    const string PREFS_CREATE_ACCESSS       = ".access";
    const string PREFS_TOGETHER_CLASS       = ".together";
    const string PREFS_SHEET_NO             = ".sheetno";
    const string PREFS_PRESUFFIX_COMBO      = ".pscombo";
    const string PREFS_PRESUFFIX_TABLE      = ".pstable";
    const string PREFS_PRESUFFIX_ACCESSOR   = ".psaccessor";
    const string PREFS_PRESUFFIX_COMBO_DATA = ".pscdata";
    const string PREFS_PRESUFFIX_DATA       = ".psdata";

    static readonly string CLASS_NAME             = $"{nameof(ImportXLS)}";

#if JAPANESE
    static readonly string MSG_ROWMAXOVER         = "[{0}] 行数が作成可能最大数を超えています[{1}].\r\n- セルに見えない空白文字が含まれている可能性があります.\r\n- どうしても最大を増やしたい場合は、" + CLASS_NAME + ".ROWS_MAX の値を増やします.";
    static readonly string MSG_COLMAXOVER         = "[{0}] 列数が作成可能最大数を超えています[{1}].\r\n- セルに見えない空白文字が含まれている可能性があります.\r\n- どうしても最大を増やしたい場合は、" + CLASS_NAME + ".COLS_MAX の値を増やします.";
    static readonly string MSG_ID_ONLYONE         = "[{0}: {1}] ID は 1 テーブルに 1 つのみです.";
    static readonly string MSG_ENUMNAME_ONLYONE   = "[{0}: {1}] 同名の ENUM が既に存在します.";
    static readonly string MSG_ENUMNAME_NOTFOUND  = "[{0}: {1}] ENUM のグループ名がありません.";
    static readonly string MSG_CONSTNAME_ONLYONE  = "[{0}: {1}] 同名の CONST クラスが既に存在します.";
    static readonly string MSG_CONSTNAME_NOTFOUND = "[{0}: {1}] CONST のクラス名がありません.";
    static readonly string MSG_SAMEMEMBER         = "[{0}: {1}] 既に同じフィールドがあります. '{2}'";
    static readonly string MSG_TYPE_NOTFOUND      = "[{0}: {1}] フィールドの型がありません. '{2}'";
    static readonly string MSG_NEED_BLANKROW      = "[{0}: {1}] 各カテゴリ間は最低 {2} 行マージンを取る必要があります.";
    static readonly string MSG_CLASSTMPL_NOTFOUND = "{0} が見つかりません.";
    static readonly string MSG_DIRECTORY_INVALID  = "ディレクトリは '{0}' から始まる相対パスを指定してください.";
    static readonly string MSG_CREATE_ENVIRONMENT = "管理クラスを作成しました.\r\nReimport で Scriptable Object が作成可能です.";
    static readonly string MSG_USE_SAME_CLASS     = "同一クラスが存在するため、最初のクラス定義を使用します. '{0}'";
    static readonly string MSG_SHEETNAME_CANT_JP  = "{0}: 日本語名のシートは無視します.";
    static readonly string MSG_NOT_FOUND_ENUMTBL  = "{0}: enum で存在しないシート名が指定されています.";
    static readonly string MSG_CREATE_ACCESS      = "*シングルトンなアクセスクラス (X_[SHEET_NAME]) を作成します";
    static readonly string MSG_TOGETHER_CLASS     = "*同じフィールドルールのテーブルがある場合、最初のシートにまとめます";

    static readonly string MSG_CANCEL             = "ユーザーキャンセルされました.";
#else
    static readonly string MSG_ROWMAXOVER         = "[{0}] The maximum number of lines that can created has been exceeded[{1}].\r\n- The cell may contain invisible whitespace.\r\n- どうしても最大を増やしたい場合は、" + CLASS_NAME + ".ROWS_MAX の値を増やします.";
    static readonly string MSG_COLMAXOVER         = "[{0}] The maximum number of columns that can be created has been exceeded[{1}].\r\n- The cell may contain invisible whitespace.";
    static readonly string MSG_ID_ONLYONE         = "[{0}: {1}] Only one ID per table.";
    static readonly string MSG_ENUMNAME_ONLYONE   = "[{0}: {1}] 'Enum' with the same name already exists.";
    static readonly string MSG_ENUMNAME_NOTFOUND  = "[{0}: {1}] There is no definition for 'Enum'.";
    static readonly string MSG_CONSTNAME_ONLYONE  = "[{0}: {1}] There is a Const Class with the same name.";
    static readonly string MSG_CONSTNAME_NOTFOUND = "[{0}: {1}] Missing Const Class Name.";
    static readonly string MSG_SAMEMEMBER         = "[{0}: {1}] The table has the same field. '{2}'";
    static readonly string MSG_TYPE_NOTFOUND      = "[{0}: {1}] There is no field type. '{2}'";
    static readonly string MSG_NEED_BLANKROW      = "[{0}: {1}] There must be a minimum of {2} line margins between each category.";
    static readonly string MSG_CLASSTMPL_NOTFOUND = "{0} not found.";
    static readonly string MSG_DIRECTORY_INVALID  = "The directory specifies a relative path. It starts with '{0}'.";
    static readonly string MSG_CREATE_ENVIRONMENT = "I created a management class.\r\nCreate a Scriptable Object with 'Reimport'.";
    static readonly string MSG_USE_SAME_CLASS     = "Use the first class definition because the same class exists. '{0}'";
    static readonly string MSG_SHEETNAME_CANT_JP  = "{0}: Ignore Japanese Sheet.";
    static readonly string MSG_NOT_FOUND_ENUMTBL  = "{0}: A non-existent sheet name is specified in the enum.";
    static readonly string MSG_CREATE_ACCESS      = "*Create a singleton access class (X_ [SHEET_NAME])";
    static readonly string MSG_TOGETHER_CLASS     = "*If you have tables with the same field rules, put them together in the first sheet.";

    static readonly string MSG_CANCEL             = "User canceled.";
#endif

    /// <summary>
    /// ポジションインデックス
    /// </summary>
    public class PosIndex
    {
        /// <summary>行</summary>
        public int R;
        /// <summary>列</summary>
        public int C;
        /// <summary>(Enumなどの)名前</summary>
        public string Name;
    }
    
    /// <summary>
    /// クラスメンバ、enum メンバ、const メンバを示す
    /// </summary>
    class Member
    {
        /// <summary>型</summary>
        public string Type;
        /// <summary>追加後尾文字列</summary>
        public string Suffix;
        /// <summary>コメント</summary>
        public string Comment;
        /// <summary>FindRow 可能なインデクサは true</summary>
        public bool   Indexer;
        /// <summary>enum は true</summary>
        public bool   IsEnum;
    }
    
    /// <summary>
    /// クラス情報
    /// </summary>
    class ClassInfo
    {
        /// <summary></summary>
        public string Comment;
        /// <summary></summary>
        public Dictionary<string, Member> Members = new Dictionary<string, Member>();
    }

    /// <summary>
    /// enum 情報
    /// </summary>
    class EnumInfo
    {
        /// <summary></summary>
        public string GroupName;
        /// <summary></summary>
        public string Comment;
        /// <summary></summary>
        public Dictionary<string, Member> Members = new Dictionary<string, Member>();
    }
    
    /// <summary>
    /// const 情報
    /// </summary>
    class ConstInfo
    {
        /// <summary></summary>
        public Dictionary<string, Member> Members = new Dictionary<string, Member>();
    }
    
    /// <summary>
    /// シート情報
    /// </summary>
    class SheetEntity
    {
        public Dictionary<string, ClassInfo> Classes = new Dictionary<string, ClassInfo>();
        public Dictionary<string, EnumInfo>  Enums   = new Dictionary<string, EnumInfo>();
        public Dictionary<string, ConstInfo> Consts  = new Dictionary<string, ConstInfo>();
        public Dictionary<string, PosIndex>  PosList = new Dictionary<string, PosIndex>();
        public string[,]                     Grid;
        public ISheet                        Sheet;
        public string                        Text    = null;
        public string                        Hash;
        public string                        ClassName;
        public string                        TableName;
        public string                        AccessorName;
        public string                        DataName;

        /// <summary>
        /// クラスや enum の情報からハッシュを取得する
        /// </summary>
        public void CreateHash()
        {
            StringBuilder sb = new StringBuilder();
            if (Classes.Count == 0)
            {
                Hash = null;
            }
            else
            {
                foreach (var cls in Classes)
                {
                    sb.Append(cls.Key + " ");
                    foreach (var member in cls.Value.Members)
                    {
                        sb.Append(member.Key + " ");
                        sb.Append(member.Value.Type + " ");
                        sb.Append(member.Value.Suffix + " ");
                    }
                }

                SHA256CryptoServiceProvider hashProvider = new SHA256CryptoServiceProvider();
                Hash =
                    string.Join(
                        "",
                        hashProvider.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()))
                        .Select(x => $"{x:x2}")
                    );
            }
        }
    }

    enum eComboPreSuffix
    {
        Prefix,
        Suffix,
    }

    static string            colTexts = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    static List<SheetEntity> sheetList;
    static string            xlsPath;
    static string            prefsKey;
    static string            classDir;
    static string            dataDir;
    static eComboPreSuffix   presuffixCombo;
    static string            presuffixTable;
    static string            presuffixAccessor;
    static eComboPreSuffix   presuffixComboData;
    static string            presuffixData;
    static bool              createAccess;
    static bool              togetherClass;
    static int               sheetNo;

    Vector2                  scroll;

    /// <summary>
    /// GUI window
    /// </summary>
    void OnGUI()
    {
        if (sheetList == null || sheetList.Count == 0)
        {
            Close();
            return;
        }
        // 前回選択していたシート番号のページがない
        if (sheetNo >= sheetList.Count)
        {
            sheetNo = sheetList.Count - 1;
        }

        GUILayout.Label("DIRECTORY", EditorStyles.boldLabel);
        GUILayout.Space(10);

        classDir = EditorGUILayout.TextField("class dir:", classDir);
        dataDir = EditorGUILayout.TextField("scriptable object dir:", dataDir);
        GUILayout.Space(10);
        GUILayout.BeginHorizontal();
        createAccess  = EditorGUILayout.Toggle("create accessor:", createAccess, GUILayout.MaxWidth(180));
        EditorGUILayout.LabelField(MSG_CREATE_ACCESS);
        GUILayout.EndHorizontal();
        GUILayout.BeginHorizontal();
        togetherClass = EditorGUILayout.Toggle("together class:", togetherClass, GUILayout.MaxWidth(180));
        EditorGUILayout.LabelField(MSG_TOGETHER_CLASS);
        GUILayout.EndHorizontal();
        GUILayout.Space(20);

        GUILayout.Label("CLASS NAME / SCRIPTABLE OBJECT NAME", EditorStyles.boldLabel);
        GUILayout.Space(10);

        // table class name
        GUILayout.BeginHorizontal();
        GUILayout.Label("table class name:", GUILayout.MaxWidth(150));
        presuffixCombo = (eComboPreSuffix)EditorGUILayout.EnumPopup(presuffixCombo, GUILayout.MaxWidth(60));
        GUILayout.Space(50);
        if (presuffixCombo == eComboPreSuffix.Prefix)
        {
            GUI.skin.textField.alignment = TextAnchor.MiddleRight;
            presuffixTable = EditorGUILayout.TextField(presuffixTable, new GUILayoutOption[] { GUILayout.MaxWidth(100) });
            GUI.skin.textField.alignment = TextAnchor.LowerLeft;
            GUILayout.Label("[SHEET_NAME]", new GUILayoutOption[] { GUILayout.MaxWidth(100) });
        }
        else
        {
            GUI.skin.label.alignment = TextAnchor.MiddleRight;
            GUILayout.Label("[SHEET_NAME]", new GUILayoutOption[] { GUILayout.MaxWidth(100) });
            GUI.skin.label.alignment = TextAnchor.LowerLeft;
            presuffixTable = EditorGUILayout.TextField(presuffixTable, new GUILayoutOption[] { GUILayout.MaxWidth(100) });
        }
        GUILayout.EndHorizontal();

        // accessor class name
        GUILayout.BeginHorizontal();
        GUILayout.Label("accessor class name:", GUILayout.MaxWidth(150));
        GUILayout.Space(60);
        GUILayout.Space(50);
        if (presuffixCombo == eComboPreSuffix.Prefix)
        {
            GUI.skin.textField.alignment = TextAnchor.MiddleRight;
            presuffixAccessor = EditorGUILayout.TextField(presuffixAccessor, new GUILayoutOption[] { GUILayout.MaxWidth(100) });
            GUI.skin.textField.alignment = TextAnchor.LowerLeft;
            GUILayout.Label("[SHEET_NAME]", new GUILayoutOption[] { GUILayout.MaxWidth(100) });
        }
        else
        {
            GUI.skin.label.alignment = TextAnchor.MiddleRight;
            GUILayout.Label("[SHEET_NAME]", new GUILayoutOption[] { GUILayout.MaxWidth(100) });
            GUI.skin.label.alignment = TextAnchor.LowerLeft;
            presuffixAccessor = EditorGUILayout.TextField(presuffixAccessor, new GUILayoutOption[] { GUILayout.MaxWidth(100) });
        }
        GUILayout.EndHorizontal();
        GUILayout.Space(10);

        // data class name
        GUILayout.BeginHorizontal();
        GUILayout.Label("scriptable obj name:", GUILayout.MaxWidth(150));
        presuffixComboData = (eComboPreSuffix)EditorGUILayout.EnumPopup(presuffixComboData, GUILayout.MaxWidth(60));
        GUILayout.Space(50);
        if (presuffixComboData == eComboPreSuffix.Prefix)
        {
            GUI.skin.textField.alignment = TextAnchor.MiddleRight;
            presuffixData = EditorGUILayout.TextField(presuffixData, new GUILayoutOption[] { GUILayout.MaxWidth(100) });
            GUI.skin.textField.alignment = TextAnchor.LowerLeft;
            GUILayout.Label("[SHEET_NAME]", new GUILayoutOption[] { GUILayout.MaxWidth(100) });
        }
        else
        {
            GUI.skin.label.alignment = TextAnchor.MiddleRight;
            GUILayout.Label("[SHEET_NAME]", new GUILayoutOption[] { GUILayout.MaxWidth(100) });
            GUI.skin.label.alignment = TextAnchor.LowerLeft;
            presuffixData = EditorGUILayout.TextField(presuffixData, new GUILayoutOption[] { GUILayout.MaxWidth(100) });
        }
        GUILayout.EndHorizontal();
        GUILayout.Space(20);


        GUILayout.Label("SHEET", EditorStyles.boldLabel);
        List<string> names = new List<string>();
        foreach (SheetEntity ent in sheetList)
        {
            names.Add(ent.Sheet.SheetName);
        }
        sheetNo = EditorGUILayout.Popup(sheetNo, names.ToArray());
        GUILayout.Space(20);

        SheetEntity entity = sheetList[sheetNo];

        GUILayout.BeginHorizontal();
        GUILayout.Label("CLASS SAMPLE", EditorStyles.boldLabel);
        if (GUILayout.Button("update", GUILayout.MaxWidth(150)) || entity.Text == null)
        {
            setNames_Table_Accessor_Data(entity);

            if (entity.ClassName == entity.Sheet.SheetName)
            {
                entity.Text = createTableClass(entity);
            }
            else
            {
                entity.Text = string.Format(MSG_USE_SAME_CLASS, entity.ClassName);
            }
        }
        GUILayout.EndHorizontal();

        scroll = EditorGUILayout.BeginScrollView(scroll);
        EditorGUILayout.TextArea(entity.Text, GUILayout.ExpandHeight(true));
        EditorGUILayout.EndScrollView();
        GUILayout.Space(20);

        using (new EditorGUILayout.HorizontalScope())
        {
            if (GUILayout.Button("CREATE ALL"))
            {
                setClassName();
                if (checkSaveDirectory() == true)
                {
                    saveSheet(sheetList);

                    dialog(MSG_CREATE_ENVIRONMENT);

                    Close();
                }
            }
            if (GUILayout.Button($"create '{entity.Sheet.SheetName}'"))
            {
                setClassName();
                if (checkSaveDirectory() == true)
                {
                    saveSheet(new List<SheetEntity>() { entity });

                    dialog(MSG_CREATE_ENVIRONMENT);

                    Close();
                }
            }
            if (GUILayout.Button("close"))
            {
                checkSaveDirectory();
                Close();
            }
        }
    }
    
    /// <summary>
    /// XLS を右クリック - ImportXLS
    /// </summary>
    [MenuItem ("Assets/ImportXLS")]
    static void Import()
    {
        if (Selection.objects.Length == 0)
        {
            log_error("no selecting excell");
            return;
        }
        var obj = Selection.objects[0];

        xlsPath  = AssetDatabase.GetAssetPath(obj);
        prefsKey = Path.GetFileNameWithoutExtension(xlsPath);

        // prefs からディレクトリを復帰（エクセル名単位）
        loadPrefs();

        using (FileStream stream = File.Open (xlsPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            var window = ScriptableObject.CreateInstance<ImportXLS>();
            sheetList = new List<SheetEntity>();

            IWorkbook book = null;
            if (Path.GetExtension(xlsPath) == ".xls")
            {
                book = new HSSFWorkbook(stream);
            }
            else
            {
                book = new XSSFWorkbook(stream);
            }

            bool error = false;

            for (int sheetno = 0; sheetno < book.NumberOfSheets; ++sheetno)
            {
                ISheet sheet = book.GetSheetAt(sheetno);

                // 日本語のシートは無視
                if (Encoding.GetEncoding("Shift_JIS").GetByteCount(sheet.SheetName) != sheet.SheetName.Length)
                {
                    log_warning(MSG_SHEETNAME_CANT_JP, sheet.SheetName);
                    continue;
                }

                SheetEntity entity = new SheetEntity();

                entity.Sheet = sheet;
                entity.Grid = GetGrid(sheet, entity.PosList);
                // 解析失敗
                if (entity.Grid == null)
                {
                    continue;
                }

                // クラス解析
                if (entity.PosList.ContainsKey(TRIGGER_ID) == true)
                {
                    if (analyzeClasses(entity) == false)
                    {
                        continue;
                    }
                }

                // enum解析
                if (analyzeEnumsAndConsts(entity) == false)
                {
                    continue;
                }

                // クラスコメント解析（あれば）
                analyzeClassComments(entity);

                entity.CreateHash();

                sheetList.Add(entity);
            }

            setClassName();

            if (error == false)
            {
                window.Show();
            }
        }
    }


    /// <summary>
    /// シートのインポートに使うクラス名を取得する（１クラス複数シート対応）
    /// ハッシュを確認し、全く同じクラス内容であれば、一番左にあるシート名＝クラスとし、マルチシート対応する
    /// </summary>
    static void setClassName()
    {
        var classNameByHash = new Dictionary<string, string>();
        foreach (SheetEntity ent in sheetList)
        {
            if (ent.Hash == null)
            {
                ent.ClassName = ent.Sheet.SheetName;
            }
            else
            {
                if (togetherClass == true)
                {
                    if (classNameByHash.ContainsKey(ent.Hash) == false)
                    {
                        ent.ClassName = ent.Sheet.SheetName;
                        classNameByHash.Add(ent.Hash, ent.ClassName);
                    }
                    else
                    {
                        ent.ClassName = classNameByHash[ent.Hash];
                    }
                }
                else
                {
                    ent.ClassName = ent.Sheet.SheetName;
                }
            }

            setNames_Table_Accessor_Data(ent);
        }
    }
    
    /// <summary>
    /// テーブルクラス名、アクセッサクラス名を作成
    /// </summary>
    static void setNames_Table_Accessor_Data(SheetEntity entity)
    {
        if (presuffixCombo == eComboPreSuffix.Prefix)
        {
            entity.TableName    = presuffixTable  + entity.ClassName;
            entity.AccessorName = presuffixAccessor + entity.ClassName;
        }
        else
        {
            entity.TableName    = entity.ClassName + presuffixTable;
            entity.AccessorName = entity.ClassName + presuffixAccessor;
        }

        if (presuffixComboData == eComboPreSuffix.Prefix)
        {
            entity.DataName     = presuffixData   + entity.ClassName;
        }
        else
        {
            entity.DataName     = entity.ClassName + presuffixData;
        }
    }
    
    /// <summary>
    /// コメント文字列を除去して返す
    /// </summary>
    static string removeSignComment(string str)
    {
        if (str == null)
        {
            return null;
        }
        foreach (string sign in SIGN_COMMENTS)
        {
            str = str.Replace(sign, "");
        }
        return str;
    }
    
    /// <summary>
    /// シートの最大行数、最大列数を調べる
    /// </summary>
    static bool getRowAndColumnMax(ISheet sheet, out int rowMax, out int colMax)
    {
        string name  = sheet.SheetName;

        rowMax = sheet.LastRowNum+1;
        colMax = 0;

        for (int r = 0; r < rowMax; r++)
        {
            IRow row = sheet.GetRow(r);
            if (checkRowIsNullOrEmpty(row) == true)
            {
                continue;
            }

            ICell cell = row.Cells[row.Cells.Count-1];
            if (colMax < cell.ColumnIndex+1)
            {
                colMax = cell.ColumnIndex+1;
            }
        }
        
        // 予め決めておいたバッファ最大量を超える場合、エラー
        if (rowMax >= ROWS_MAX)
        {
            log_error(MSG_ROWMAXOVER, name, $"{rowMax} >= {ROWS_MAX}");
            return false;
        }
        if (colMax >= COLS_MAX)
        {
            log_error(MSG_COLMAXOVER, name, $"{colMax} >= {COLS_MAX}");
            return false;
        }

        return true;
    }
    
    /// <summary>
    /// １行分のセルデータを取得
    /// </summary>
    /// <param name="sheet">シート</param>
    /// <param name="grid">グリッド情報</param>
    /// <param name="posList">ID, [CLASS], enum のデータポジションを示すリスト</param>
    /// <param name="r">行数</param>
    /// <param name="marginrow_between_category">カテゴリ先頭までに、この値が 0 である必要がある</param>
    /// <returns>false..取得失敗</returns>
    static bool getCells(ISheet sheet, string[,] grid, Dictionary<string, PosIndex> posList, int r, ref int marginrow_between_category)
    {
        string name  = sheet.SheetName;

        IRow   row   = sheet.GetRow(r);
        if (checkRowIsNullOrEmpty(row) == true)
        {
            if (--marginrow_between_category < 0)
            {
                marginrow_between_category = 0;
            }
            return true;
        }

        for (int c = 0; c < row.Cells.Count; c++)
        {
            ICell    cell     = row.Cells[c];
            string   cellstr  = cell.ToString();
            bool     category = false;
            CellType celltype = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;

            switch (celltype)
            {
                case CellType.Numeric:
                    cellstr = cell.NumericCellValue.ToString();
                    break;
                case CellType.Boolean:
                    cellstr = cell.BooleanCellValue.ToString();
                    break;
                case CellType.String:
                    cellstr = cell.StringCellValue.ToString();
                    break;
            }

            int col = cell.ColumnIndex;
            grid[r, col] = cellstr;

            // テーブルのトリガー ID
            if (cellstr == TRIGGER_ID)
            {
                if (posList.ContainsKey(TRIGGER_ID) == true)
                {
                    // 既に ID がある
                    log_error(MSG_ID_ONLYONE, name, GetXLS_RC(r, col));
                    return false;
                }
                posList.Add(TRIGGER_ID, new PosIndex(){ R=r, C=col});

                category = true;
            }
            else
            // テーブルのクラスコメント
            if (cellstr == TRIGGER_CLASS)
            {
                if (posList.ContainsKey(TRIGGER_CLASS) == true)
                {
                    // 既に ID がある
                    log_error(MSG_ID_ONLYONE, name, GetXLS_RC(r, col));
                    return false;
                }
                posList.Add(TRIGGER_CLASS, new PosIndex(){ R=r, C=col});

                category = true;
            }
            else
            // enum グループ
            if (cellstr == TRIGGER_ENUM)
            {
                if (c+1 >= row.Cells.Count)
                {
                    // enum 名がない
                    log_error(MSG_ENUMNAME_NOTFOUND, name, GetXLS_RC(r, col));
                    return false;
                }
                string ename = row.Cells[c+1].ToString();
                string key   = name + "." + ename;
                if (posList.ContainsKey(key) == true)
                {
                    // 既に同じ enum がある
                    log_error(MSG_ENUMNAME_ONLYONE, name, GetXLS_RC(r, col));
                    return false;
                }
                posList.Add(key, new PosIndex(){ R=r, C=col+1, Name=ename });

                category = true;
            }
            else
            // enum グループ
            if (cellstr == TRIGGER_CONST)
            {
                if (c+1 >= row.Cells.Count)
                {
                    // enum 名がない
                    log_error(MSG_CONSTNAME_NOTFOUND, name, GetXLS_RC(r, col));
                    return false;
                }
                string ename = row.Cells[c+1].ToString();
                string key   = TRIGGER_CONST + "." + ename;
                if (posList.ContainsKey(key) == true)
                {
                    // 既に同じ const がある
                    log_error(MSG_CONSTNAME_ONLYONE, name, GetXLS_RC(r, col));
                    return false;
                }
                posList.Add(key, new PosIndex(){ R=r, C=col+1, Name=ename });

                category = true;
            }
            else
            // enum グループ
            if (cellstr.ToLower() == TRIGGER_GLOBAL_ENUM.ToLower())
            {
                if (c+1 >= row.Cells.Count)
                {
                    // enum 名がない
                    log_error(MSG_ENUMNAME_NOTFOUND, name, GetXLS_RC(r, col));
                    return false;
                }
                string ename = row.Cells[c+1].ToString();
                string key   = TRIGGER_GLOBAL_ENUM + "." + ename;
                if (posList.ContainsKey(key) == true)
                {
                    // 既に同じ enum がある
                    log_error(MSG_ENUMNAME_ONLYONE, name, GetXLS_RC(r, col));
                    return false;
                }
                posList.Add(key, new PosIndex(){ R=r, C=col+1, Name=ename });

                category = true;
            }

            if (category == true)
            {
                if (marginrow_between_category > 0)
                {
                    // 各カテゴリ間は最低マージン 2 行必要
                    log_error(MSG_NEED_BLANKROW, name, GetXLS_RC(r, col), MARGINROW_BETWEEN_CTG);
                    return false;
                }
                marginrow_between_category = MARGINROW_BETWEEN_CTG;
            }
        }

        return true;
    }
    
    /// <summary>
    /// Row が null または空行か確認する
    /// </summary>
    /// <param name="row">確認する行</param>
    /// <returns>true..null または空行</returns>
    static bool checkRowIsNullOrEmpty(IRow row)
    {
        if (row == null)
        {
            return true;
        }

        // エクセル上、見た目に何もないがデータとして "" だけ検出される行を null とみなし無視する
        // null ではないのに、Cells.Count = 0 の row を返すこともある…
        for (int c = 0; c < row.Cells.Count; c++)
        {
            // なにか入っていた
            if (string.IsNullOrEmpty(row.Cells[c].ToString()) == false)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// １行上にコメントがある場合、その文字列を返す
    /// </summary>
    static string getCommentUp(string[,] grid, int r, int c)
    {
        string cell = null;
        if (r >= 1)
        {
            cell = grid[r-1, c];
        }
        if (CheckSignComment(cell) == true)
        {
            return removeSignComment(cell);
        }
        return "";
    }

    /// <summary>
    /// １行右にコメントがある場合、その文字列を返す
    /// </summary>
    static string getCommentRight(string[,] grid, int r, int c)
    {
        string cell = null;
        if ((c+1) < (grid.GetUpperBound(1)+1))
        {
            cell = grid[r, c+1];
        }
        if (CheckSignComment(cell) == true)
        {
            return removeSignComment(cell);
        }
        return "";
    }
    
    /// <summary>
    /// ImportXLS.cs のあるディレクトリを取得する
    /// </summary>
    static string searchImportXLSDirectory()
    {
        string[] files = Directory.GetFiles(Application.dataPath, $"{nameof(ImportXLS)}.cs", SearchOption.AllDirectories);
        if (files != null && files.Length == 1)
        {
            // フルパスから相対パスに
            string path = Path.GetDirectoryName(files[0]).Replace("\\", "/");
            path = Path.Combine(Path.GetFileNameWithoutExtension(Application.dataPath), path.Replace(Application.dataPath + "/", ""));
            return path;
        }
        return null;
    }
    
    /// <summary>
    /// セーブディレクトリが適切かチェック＆修正. 問題なければ prefs に保存
    /// </summary>
    /// <returns>true..成功</returns>
    static bool checkSaveDirectory()
    {
        // windows と mac の違いを吸収
        classDir = classDir.Replace("\\", "/").Trim();
        dataDir  = dataDir.Replace("\\", "/").Trim();

        // パスは Assets/ から始まる相対パス
        string assetPath = Path.GetFileName(Application.dataPath);

        if (classDir.IndexOf(assetPath) != 0)
        {
            dialog_error(MSG_DIRECTORY_INVALID, assetPath);
            return false;
        }
        if (dataDir.IndexOf(assetPath) != 0)
        {
            dialog_error(MSG_DIRECTORY_INVALID, assetPath);
            return false;
        }
        
        // フォルダを予め作成しておく
        CompleteDirectory(classDir);
        CompleteDirectory(dataDir);

        // prefs にディレクトリを保存（エクセル名単位）
        savePrefs();

        return true;
    }

    /// <summary>
    /// prefs にディレクトリを保存（エクセル名単位）
    /// </summary>
    static void savePrefs()
    {
        EditorPrefs.SetString(prefsKey + PREFS_CLASS_DIRECTORY, classDir);
        EditorPrefs.SetString(prefsKey + PREFS_DATA_DIRECTORY, dataDir);
        EditorPrefs.SetBool(prefsKey + PREFS_CREATE_ACCESSS, createAccess);
        EditorPrefs.SetBool(prefsKey + PREFS_TOGETHER_CLASS, togetherClass);
        EditorPrefs.SetInt(prefsKey + PREFS_SHEET_NO, sheetNo);
        EditorPrefs.SetInt(prefsKey + PREFS_PRESUFFIX_COMBO, (int)presuffixCombo);
        EditorPrefs.SetString(prefsKey + PREFS_PRESUFFIX_TABLE, presuffixTable);
        EditorPrefs.SetString(prefsKey + PREFS_PRESUFFIX_ACCESSOR, presuffixAccessor);
        EditorPrefs.SetInt(prefsKey + PREFS_PRESUFFIX_COMBO_DATA, (int)presuffixComboData);
        EditorPrefs.SetString(prefsKey + PREFS_PRESUFFIX_DATA, presuffixData);
    }

    /// <summary>
    /// prefs からディレクトリを復帰（エクセル名単位）
    /// </summary>
    static void loadPrefs()
    {
        classDir      = EditorPrefs.GetString(prefsKey + PREFS_CLASS_DIRECTORY);
        dataDir       = EditorPrefs.GetString(prefsKey + PREFS_DATA_DIRECTORY);
        createAccess  = EditorPrefs.GetBool(prefsKey + PREFS_CREATE_ACCESSS, true);
        togetherClass = EditorPrefs.GetBool(prefsKey + PREFS_TOGETHER_CLASS, false);
        sheetNo       = EditorPrefs.GetInt(prefsKey + PREFS_SHEET_NO);

        // 初期値
        if (string.IsNullOrEmpty(classDir) == true || classDir.IndexOf(ASSETS_CLASS) != 0)
        {
            classDir = ASSETS_CLASS;
        }
        if (string.IsNullOrEmpty(dataDir) == true || dataDir.IndexOf(ASSETS_RESOURCE) != 0)
        {
            dataDir = ASSETS_RESOURCE;
        }

        int no;

        no = EditorPrefs.GetInt(prefsKey + PREFS_PRESUFFIX_COMBO);
        Enum.TryParse(no.ToString(), out eComboPreSuffix presuffixCombo);
        presuffixTable     = EditorPrefs.GetString(prefsKey + PREFS_PRESUFFIX_TABLE, PREFIX_TABLE);
        presuffixAccessor  = EditorPrefs.GetString(prefsKey + PREFS_PRESUFFIX_ACCESSOR, PREFIX_ACCESS);

        no = EditorPrefs.GetInt(prefsKey + PREFS_PRESUFFIX_COMBO_DATA);
        Enum.TryParse(no.ToString(), out eComboPreSuffix presuffixComboData);
        presuffixData      = EditorPrefs.GetString(prefsKey + PREFS_PRESUFFIX_DATA, PREFIX_DATA);
    }

    /// <summary>
    /// コメントテキスト生成
    /// </summary>
    /// <param name="sb">Append する StringBuilder</param>
    /// <param name="description">説明</param>
    /// <param name="indent">タブインデントの数</param>
    static void addCommentText(StringBuilder sb, string description, int indent)
    {
        string tab = "".PadLeft(indent, '\t');

        if (string.IsNullOrEmpty(description) == true)
        {
//            sb.AppendLine($"{tab}///<summary>\r\n{tab}/// \r\n{tab}///</summary>");
        }
        else
        {
            string[] descs = description.Replace("\r", "").Split('\n');
            sb.AppendLine($"{tab}///<summary>");
            foreach (string desc in descs)
            {
                sb.AppendLine($"{tab}/// {desc}");
            }
            sb.AppendLine($"{tab}///</summary>");
        }
    }
    
    /// <summary>
    /// Path.Combine の後、フォルダ区切りを / にして返す
    /// </summary>
    static string pathCombine(string path0, string path1)
    {
        return Path.Combine(path0, path1).Replace("\\", "/");
    }

    /// <summary>
    /// ログ表示
    /// </summary>
    static void log(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        Debug.Log($"{CLASS_NAME}:" + msg);
    }

    /// <summary>
    /// 警告表示
    /// </summary>
    static void log_warning(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        Debug.LogWarning($"{CLASS_NAME}:" + msg);
    }

    /// <summary>
    /// エラー表示
    /// </summary>
    static void log_error(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        Debug.LogError($"{CLASS_NAME}:" + msg);
    }


    /// <summary>
    /// ダイアログ表示
    /// </summary>
    static void dialog(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        EditorUtility.DisplayDialog($"{CLASS_NAME}", msg, "ok");
    }

    /// <summary>
    /// エラー表示
    /// </summary>
    static void dialog_error(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        EditorUtility.DisplayDialog($"{CLASS_NAME}", $"[ERROR]\r\n{msg}", "ok");
    }

    /// <summary>
    /// ok/cancel ダイアログ表示
    /// </summary>
    static bool dialog_select(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        return EditorUtility.DisplayDialog($"{CLASS_NAME}", msg, "ok", "cancel");
    }

    /// <summary>
    /// キャンセルつき進捗バー (index+1)/max %
    /// </summary>
    static bool cancelableProgressBar(int index, int max, string msg)
    {
        float	perc = (float)(index+1) / (float)max;
        
        bool result =
            EditorUtility.DisplayCancelableProgressBar(
                nameof(ImportXLS),
                perc.ToString("00.0%") + "　" + msg,
                perc
            );
        if (result == true)
        {
            EditorUtility.ClearProgressBar();
            dialog(MSG_CANCEL);
            return true;
        }
        return false;
    }
}
