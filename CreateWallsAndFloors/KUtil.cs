using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CreateWallsAndFloors
{
    public static class KUtil
    {
        //壁の線分（Curve）を抽出するための関数。連続した2マス以上の線分を見つける
        public static IList<Curve> FindContinuousOnes(int[,] array, int a)
        {
            //結果として返す変数
            IList<Curve> curves = new List<Curve>();

            //行列のサイズを取得
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);

            // 横方向の連続を確認
            for (int i = 0; i < rows; i++)
            {
                double startOffsetX = 0;
                double endOffsetX = 0;
                int start = -1;
                for (int j = 0; j < cols; j++)
                {
                    if (array[i, j] == a && start == -1)
                    {
                        // 連続の開始点を記録
                        start = j;
                        if (j >0)
                        {
                            if (a == 2 && array[i, j - 1] == 1)//******内壁を追跡中に左が外壁1だったら1つ減じる。（i-1>=0とする）
                            {
                                startOffsetX = -1;
                            }
                        }
                    }
                    else if (array[i, j] != a && start != -1)
                    {
                        // 連続が終了した場合
                        if (j - start > 1) // 2つ以上の連続のみ出力
                        {
                            //"横方向の連続はデータ配列上で: ({i}, {start}) から ({i}, {j - 1})"
                            //Revitモデル上では行はY軸、列はX軸にあたるので入れ替え。さらにY軸は大小の方向が逆なので
                            //最大値からiを引いている
                            //XYZ startPoint = new XYZ(i, start, 0);
                            //XYZ endPoint = new XYZ(i, j - 1, 0);

                            if (a == 2 && array[i, j] == 1)//******内壁を追跡中に右が外壁だったら1つ増やす。
                            {
                                endOffsetX = 1;
                            }

                            XYZ startPoint = new XYZ(start+ startOffsetX, -i ,  0);
                            XYZ endPoint = new XYZ((j - 1)+ endOffsetX, -i, 0);
                            Line line = Line.CreateBound(startPoint, endPoint);
                            curves.Add(line);
                            startOffsetX = 0;//オフセットのリセット
                            endOffsetX = 0;
                        }
                        start = -1; // 開始点をリセット
                    }
                }
                // 行末まで連続が続いた場合の処理
                if (start != -1 && cols - start > 1) // 行末までの連続が2つ以上の場合
                {
                    //"横方向の連続はデータ配列上で:  ({i}, {start}) から ({i}, {cols - 1})");
                    //Revitモデル上では行はY軸、列はX軸にあたるので入れ替え。さらにY軸は大小の方向が逆なので
                    //最大値からiを引いている
                    //XYZ startPoint = new XYZ(i, start, 0);
                    //XYZ endPoint = new XYZ(i, cols - 1, 0);
                    XYZ startPoint = new XYZ(start , -i,  0);
                    XYZ endPoint = new XYZ((cols - 1) , -i, 0);
                    Line line = Line.CreateBound(startPoint, endPoint);
                    curves.Add(line);
                }
            }

            // 縦方向の連続を確認
            for (int j = 0; j < cols; j++)
            {
                double startOffsetY = 0;
                double endOffsetY = 0;
                int start = -1;
                for (int i = 0; i < rows; i++)
                {
                    if (array[i, j] == a && start == -1)
                    {
                        // 連続の開始点を記録
                        start = i;
                        if (i > 0)
                        {
                            if (a == 2 && array[i-1, j] == 1)//******内壁を追跡中に上が外壁1だったら1つ減じる。（i-1>=0とする）
                            {
                                startOffsetY = -1;
                            }
                        }
                    }
                    else if (array[i, j] != a && start != -1)
                    {
                        // 連続が終了した場合
                        if (i - start > 1) // 2つ以上の連続のみ出力
                        {
                            //{start}, {j}) から ({i - 1}, {j})
                            //Revitモデル上では行はY軸、列はX軸にあたるので入れ替え。さらにY軸は大小の方向が逆なので
                            //最大値からiを引いている
                            //XYZ startPoint = new XYZ(start, j, 0);
                            //XYZ endPoint = new XYZ(i - 1, j, 0);

                            if (a == 2 && array[i, j] == 1)//******内壁を追跡中に下が外壁だったら1つ増やす。
                            {
                                endOffsetY = 1;
                            }

                            XYZ startPoint = new XYZ(j , -start- startOffsetY, 0);//方向が逆であることに注意
                            XYZ endPoint = new XYZ(j , -(i - 1) - endOffsetY, 0);//同上
                            Line line = Line.CreateBound(startPoint, endPoint);
                            curves.Add(line);
                        }
                        start = -1; // 開始点をリセット
                    }
                }
                // 列末まで連続が続いた場合の処理
                if (start != -1 && rows - start > 1) // 列末までの連続が2つ以上の場合
                {
                    //"横方向の連続はデータ配列上で:  ({start}, {j}) から ({rows - 1}, {j})");
                    //Revitモデル上では行はY軸、列はX軸にあたるので入れ替え。さらにY軸は大小の方向が逆なので
                    //最大値からiを引いている
                    //XYZ startPoint = new XYZ(start, j, 0);
                    //XYZ endPoint = new XYZ(rows - 1, j, 0);
                    XYZ startPoint = new XYZ(j , -start ,  0);
                    XYZ endPoint = new XYZ(j , -(rows - 1),  0);
                    Line line = Line.CreateBound(startPoint, endPoint);
                    curves.Add(line);
                }
            }
            return curves;
        }
        //Excelファイルを読むためのダイアログを表示する
        public static string OpenExcel()
        {
            string fileName = string.Empty;

            //OpenFileDialog
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "ファイル選択";
                openFileDialog.Filter = "ExcelFiles | *.xls;*.xlsx;*.xlsm";
                //初期表示フォルダはデスクトップ
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                //ファイル選択ダイアログを開く
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog.FileName;
                }
            }
            return fileName;
        }
        //ミリメーターを内部単位に変換する
        public static double CVmmToInt(double x)
        {
            return UnitUtils.ConvertToInternalUnits(x, UnitTypeId.Millimeters);
        }
        //内部単位をミリメーターに変換する
        public static double CVintTomm(double x)
        {
            return UnitUtils.ConvertFromInternalUnits(x, UnitTypeId.Millimeters);
        }
        //窓やドアを取り付ける
        public static void CreateWindowAndDoor(UIDocument uidoc, Document doc, string fsFamilyName,
            string fsName, string levelName, double xCoord, double yCoord, double offset)
        {
            // LINQ 指定のファミリシンボルを探して取得する
            FamilySymbol familySymbol = (from fs in new FilteredElementCollector(doc).
                 OfClass(typeof(FamilySymbol)).
                 Cast<FamilySymbol>()
                                         where (fs.Family.Name == fsFamilyName && fs.Name == fsName)
                                         select fs).First();

            // LINQ 指定のレベルを取得する
            Level level = (from lvl in new FilteredElementCollector(doc).
                           OfClass(typeof(Level)).
                           Cast<Level>()
                           where (lvl.Name == levelName)
                           select lvl).First();

            // 座標mmを内部単位に変換する
            double x = xCoord;
            double y = yCoord;

            XYZ xyz = new XYZ(x, y, level.Elevation+ offset);

            //部材を挿入する壁で、最も近いホスト壁を取得する
            //壁を全部取得する
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Wall));
            //当該レベルにある壁だけ抽出する。リストにする
            List<Wall> walls = collector.Cast<Wall>().Where(wl => wl.LevelId == level.Id).ToList();

            Wall wall = null;
            //距離の初期値を最大値にしておく
            double distance = double.MaxValue;
            //壁を全部調べて、最も近いホスト壁を抽出する
            foreach (Wall w in walls)
            {
                //壁のカーブからの最短直線距離を測る
                double proximity = (w.Location as LocationCurve).Curve.Distance(xyz);
                //これまでの値よりも小さい場合には、より接近した壁だと思われるので、その壁をホスト候補とする
                if (proximity < distance)
                {
                    distance = proximity;
                    wall = w;
                }
            }

            // Create window.
            //using (Transaction t = new Transaction(doc, "Create window and door"))
            //{
                //t.Start();

                if (!familySymbol.IsActive)
                {
                    //ファミリシンボルがアクティブではないので、アクティブにする
                    familySymbol.Activate();
                    doc.Regenerate();
                }

                // 部材の配置
                // ホストであるwallを指定しない場合には部材はホストなし
                FamilyInstance window = doc.Create.NewFamilyInstance(xyz, familySymbol, wall, level,StructuralType.NonStructural);
                //t.Commit();
            //}
            //string prompt = "部材は配置されました";
            //TaskDialog.Show("Revit", prompt);
        }
        //目的のレベル（階）を探して返す
        public static Level GetLevel(Document doc, string levelName)
        {
            Level result = null;
            //エレメントコレクターの作成
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            //レベルを全て検出する
            ICollection<Element> collection = collector.OfClass(typeof(Level)).ToElements();
            //目的のレベルを探す
            foreach (Element element in collection)
            {
                Level level = element as Level;
                if (null != level)
                {
                    if (level.Name == levelName)
                    {
                        result = level;
                    }
                }
            }
            return result;
        }
        //床作成のための外壁座標を取得する（数字が入っているセルの座標を取得する）
        public static SortedDictionary<int, XYZ> FindFloorLine(string[,] array)
        {
            //自動ソートされるディクショナリ
            var dict = new SortedDictionary<int, XYZ>();
            //行列のサイズを取得
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    int numeric;
                    if (int.TryParse(array[i, j], out numeric))
                    {
                        //dict.Add(numeric, new XYZ(i, j, 0));
                        dict.Add(numeric, new XYZ(j , -i, 0));
                    }
                }
            }
            return dict;
        }
        //床のためのカーブを作る
        public static CurveLoop FloorProfile(ModelSetting modelsetting)
        {
            //床のためのカーブを作る
            CurveLoop floorLoop = new CurveLoop();
            XYZ first = new XYZ();//始点を記憶
            XYZ previous = new XYZ();//直前の点を記憶
            bool isFirstPoint = true;//始点かどうか
            foreach (var kvp in modelsetting.floorCurveLoop)
            {
                if (isFirstPoint)
                {
                    first = kvp.Value;//開始点を記憶
                    previous = kvp.Value;//直前の点を記憶
                    isFirstPoint = false;
                }
                else
                {
                    floorLoop.Append(Line.CreateBound(previous, kvp.Value));//線分をループに追加していく
                    previous = kvp.Value;//直前の点を記憶
                }
            }
            floorLoop.Append(Line.CreateBound(previous, first));//最後は始点に戻る
            //ループの完成
            return floorLoop;
        }
        public static bool ReadInputInfo(out ModelSetting modelsetting)
        {
            //Excelから入力条件データを読んでcolorArrayとvalueArrayを作成する            
            int[,] colorArray;//Excelのセルの色を数値で保持する
            string[,] valueArray;//Excelのセルの文字を保持する

            //設定とモデリング情報が入っているExcelファイルを読む
            string excelFile = KUtil.OpenExcel();

            using (var excel = new ExcelPackage(new FileInfo(excelFile)))
            {
                //独自のModelSettingクラス
                modelsetting = new ModelSetting();

                //設定をExcelシートから読む---------------------------------------------
                int sheetPage = 1;//SPECシート, 企業のExcelは0から始まる模様だがmicrosoft365では１から開始。確認要
                //シートを取得
                var sheet = excel.Workbook.Worksheets[sheetPage];
                //エラー項目
                string errorItem = "";
                //階高設定を取得。内部単位に変換
                var temp = sheet.Cells[1, 2].Value;
                if (temp != null)
                {
                    modelsetting.kaidaka = KUtil.CVmmToInt(Convert.ToDouble(temp.ToString()));
                }
                else
                {
                    errorItem = "Excel入力情報のエラー：\n階高設定がない" + "\n";
                }
                //ひと目盛り設定を取得。内部単位に変換
                temp = sheet.Cells[2, 2].Value;
                if (temp != null)
                {
                    modelsetting.memori = KUtil.CVmmToInt(Convert.ToDouble(temp.ToString()));
                }
                else
                {
                    errorItem = "目盛り設定がない" + "\n";
                }
                //外壁シンボルを取得
                temp = sheet.Cells[4, 3].Value;
                if (temp != null)
                {
                    modelsetting.outWallTypeName = temp.ToString();
                }
                else
                {
                    errorItem = "外壁シンボル設定がない" + "\n";
                }
                //内壁シンボルを取得
                temp = sheet.Cells[5, 3].Value;
                if (temp != null)
                {
                    modelsetting.innerWallTypeName = temp.ToString();
                }
                else
                {
                    errorItem = "内壁タイプ設定がない" + "\n";
                }
                //床シンボルを取得
                temp = sheet.Cells[6, 3].Value;
                if (temp != null)
                {
                    modelsetting.floorTypeName = temp.ToString();
                }
                else
                {
                    errorItem = "床タイプ設定がない" + "\n";
                }
                //天井シンボルを取得
                temp = sheet.Cells[7, 3].Value;
                if (temp != null)
                {
                    modelsetting.ceilingTypeName = temp.ToString();
                }
                else
                {
                    errorItem = "天井タイプ設定がない" + "\n";
                }
                //窓のファミリ名とシンボル名を取得
                temp = sheet.Cells[8, 2].Value;
                if (temp != null)
                {
                    modelsetting.windowFamilyName = temp.ToString();
                }
                else
                {
                    errorItem = "窓のファミリ名設定がない" + "\n";
                }
                temp = sheet.Cells[8, 3].Value;
                if (temp != null)
                {
                    modelsetting.windowTypeName = temp.ToString();
                }
                else
                {
                    errorItem = "窓のタイプ名設定がない" + "\n";
                }
                //ドアのファミリ名とシンボル名を取得
                temp = sheet.Cells[9, 2].Value;
                if (temp != null)
                {
                    modelsetting.doorFamilyName = temp.ToString();
                }
                else
                {
                    errorItem = "ドアのファミリ名設定がない" + "\n";
                }
                temp = sheet.Cells[9, 3].Value;
                if (temp != null)
                {
                    modelsetting.doorTypeName = temp.ToString();
                }
                else
                {
                    errorItem = "ドアのタイプ名設定がない" + "\n";
                }
                //柱のファミリ名とシンボル名を取得
                temp = sheet.Cells[10, 2].Value;
                if (temp != null)
                {
                    modelsetting.columnFamilyName = temp.ToString();
                }
                else
                {
                    errorItem = "柱のファミリ名設定がない" + "\n";
                }
                temp = sheet.Cells[10, 3].Value;
                if (temp != null)
                {
                    modelsetting.columnTypeName = temp.ToString();
                }
                else
                {
                    errorItem = "柱のタイプ名設定がない" + "\n";
                }


                if (!errorItem.Equals(""))
                {
                    TaskDialog.Show("Error", errorItem);

                    return false;
                }

                //モデリングに必要な情報をExcelシートから読む--------------------------
                //Excelシート番号。
                sheetPage = 3;//level１の情報
                //シートを取得
                sheet = excel.Workbook.Worksheets[sheetPage];
                var cellsRange = sheet.Dimension;

                //外壁、内壁の位置を解析するためのデータ配列------------
                //データ配列（壁の色用。外壁の赤は1を代入、内壁の青は2を代入）
                colorArray = new int[cellsRange.End.Row, cellsRange.End.Column];
                //床を作成するための座標を検出するための配列。外壁で数値が入っているものだけを拾った配列
                //窓のW、ドアのDも入っている
                //データ配列（文字用）
                valueArray = new string[cellsRange.End.Row, cellsRange.End.Column];

                //データが存在する最大の行番号
                int maxI = colorArray.GetLength(0);
                //Excelを読んで解析用のモデルに数値を入れる
                //行をスキャンする
                for (int i = cellsRange.Start.Row; i <= cellsRange.End.Row; i++)//セルの行、列番号は１から始まる
                {
                    //列をスキャンする
                    for (int j = cellsRange.Start.Column; j <= cellsRange.End.Column; j++)
                    {
                        //セルの色を取得
                        string color = sheet.Cells[i, j].Style.Fill.BackgroundColor.Rgb ?? "xxxxxxxx";
                        //セルの値を取得
                        string val = (sheet.Cells[i, j].Value ?? "").ToString();
                        //セルの文字をvalueArray配列に代入（数値、W、D、C等）
                        if (val != null)
                        {
                            valueArray[i - 1, j - 1] = val;
                        }

                        //色を抽出する
                        color = color.Substring(2, 6);
                        switch (color)
                        {
                            case @"FF0000"://赤　外壁
                                colorArray[i - 1, j - 1] = 1;
                                break;
                            case @"00FF00"://緑　内壁
                                colorArray[i - 1, j - 1] = 2;
                                break;
                        }
                    }
                }
            }
            //外壁のカーブを取得する
            IList<Curve> GaihekiCurves = KUtil.FindContinuousOnes(colorArray, 1);
            modelsetting.gaihekiCurves = GaihekiCurves;
            //内壁のカーブを取得する
            IList<Curve> NaihekiCurves = KUtil.FindContinuousOnes(colorArray, 2);
            modelsetting.naihekiCurves = NaihekiCurves;
            //床の外形線のラインの情報を取得する
            SortedDictionary<int, XYZ> floorCurveLoop = FindFloorLine(valueArray);
            modelsetting.floorCurveLoop = floorCurveLoop;
            //窓の位置
            IList<XYZ> windowsPosition = FindSymbolPosition(valueArray, "W");
            modelsetting.windowsPoints = windowsPosition;
            //ドアの位置
            IList<XYZ> doorsPosition = FindSymbolPosition(valueArray, "D");
            modelsetting.doorsPoints = doorsPosition;
            //柱の位置
            IList<XYZ> columnsPosition = FindSymbolPosition(valueArray, "C");
            modelsetting.columnsPoints = columnsPosition;

            return true;
        }
        public static IList<XYZ> FindSymbolPosition(string[,] array, string symbol)
        {
            //自動ソートされるディクショナリ
            var result = new List<XYZ>();
            //行列のサイズを取得
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if(array[i, j] != null)
                    {
                        //シンボルと一致したらリストに入れる
                        if (array[i, j].Equals(symbol))
                        {
                            result.Add(new XYZ(j, -i, 0));
                        }
                    }
                }
            }
            return result;
        }
        public static void CreateColumn(Document doc, String levelName1,string levelName2,string columnFamilyName, string columnSymbolName,double x,double y)
        {
            string message = null;

                try
                {
                // 柱のファミリとタイプを取得
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                    FamilySymbol columnSymbol = collector
                        .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                        .FirstOrDefault(q => q.Name == columnSymbolName && q.Family.Name == columnFamilyName);

                    if (columnSymbol == null)
                    {
                        message = "指定された柱のファミリまたはタイプが見つかりません。";
                    }

                    // レベルを取得
                    Level level1 = GetLevel(doc, levelName1);
                    Level level2 = GetLevel(doc, levelName2);

                    if (level1 == null || level2 == null)
                    {
                        message = "指定されたレベルが見つかりません。";
                    }

                    // トランザクションを開始
                    //using (Transaction trans = new Transaction(doc, "Place Column"))
                    //{
                        //trans.Start();

                        // ファミリシンボルをアクティブにする
                        if (!columnSymbol.IsActive)
                        {
                            columnSymbol.Activate();
                        }

                        // 柱を作成
                        XYZ location = new XYZ(x, y, 0);
                        FamilyInstance column = doc.Create.NewFamilyInstance(location, columnSymbol, level1, Autodesk.Revit.DB.Structure.StructuralType.Column);

                        // 上部の拘束を設定
                        
                        Parameter topLevelParam = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                        if (topLevelParam != null && !topLevelParam.IsReadOnly)
                        {
                            topLevelParam.Set(level2.Id);
                        }
                        
                        //trans.Commit();
                    //}

                }
                catch (Exception ex)
                {
                    message = ex.Message;
                }
            if (message != null)
            {
                TaskDialog.Show("Error", message);
            }
        }
        public static void CreateGrids(Document doc)
        {
            double y1 = -100;
            double y2 = 30;
            double[] x = {2,34,74,114,147 };
            string[] symbX = {"1","2","3","4","5" };

            double x1 = -30;
            double x2 = 180;
            double[] y = { -2, -33, -66 };
            string[] symbY = { "A", "B", "C"};

            for (int i = 0; i < x.Length; i++) 
            {
                XYZ start = new XYZ(x[i], y1, 0);
                XYZ end = new XYZ(x[i], y2, 0);
                Line line = Line.CreateBound(start, end);
                Grid grid = Grid.Create(doc, line);
                grid.Name = symbX[i];
            }

            for (int i = 0; i < y.Length; i++)
            {
                XYZ start = new XYZ(x1, y[i], 0);
                XYZ end = new XYZ(x2, y[i], 0);
                Line line = Line.CreateBound(start, end);
                Grid grid = Grid.Create(doc, line);
                grid.Name = symbY[i];
            }
        }
    }
}
