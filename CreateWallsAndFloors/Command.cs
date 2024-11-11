#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using Autodesk.Revit.DB.Mechanical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System.Windows.Forms.VisualStyles;


#endregion
//当ソフトにはEPPlus4.5.3.3を使用しています。これはLGPLライセンスです。著作権はEPPlus Software社です。
namespace CreateWallsAndFloors
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;


            //独自のModelSettingクラスをインスタンス化。設定情報を保持する
            ModelSetting modelsetting = new ModelSetting();

            //Excelから設定情報を読む。壁の座標なども読んで設定する
            bool result = KUtil.ReadInputInfo(out modelsetting);

            //Excel入力にエラーがあったときは終了
            if(!result) return Result.Failed;

            //ファミリ等の設定の読み出し
            double wallHeight = modelsetting.kaidaka;
            double offset = 0;//床からのオフセット（壁）
            double windowOffset = KUtil.CVmmToInt(1000);//床からのオフセット（窓）mm
            string outWallTypeName = modelsetting.outWallTypeName;
            string innerWallTypeName = modelsetting.innerWallTypeName;
            string floorTypeName = modelsetting.floorTypeName;
            string levelName = @"レベル 1";
            string levelName2 = @"レベル 2";
            double level2Height = 4500;//mm
            double level2HeightFeet = UnitUtils.ConvertToInternalUnits(level2Height, UnitTypeId.Millimeters);

            string windowFamilyName = modelsetting.windowFamilyName;
            string windowSymbolName = modelsetting.windowTypeName;
            string doorFamilyName = modelsetting.doorFamilyName;
            string doorSymbolName = modelsetting.doorTypeName;
            string columnFamilyName = modelsetting.columnFamilyName;
            string columnSymbolName = modelsetting.columnTypeName;

            //目的のレベル（階）を取得する
            Level level = KUtil.GetLevel(doc, levelName);//1階
            Level level2 = KUtil.GetLevel(doc, levelName2);//2階

            //外壁、内壁、床タイプの検索
            WallType outWallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault(q => q.Name == outWallTypeName);
            WallType innerWallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault(q => q.Name == innerWallTypeName);
            FloorType floorType = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().FirstOrDefault(q => q.Name == floorTypeName);

            //床のためのループプロファイル（線分をつないだもの）を作る
            CurveLoop floorLoop = KUtil.FloorProfile(modelsetting);

            //データチェック
            //ModelOut(modelsetting);

            //モデルの生成
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("TransactionWallFloor");
                try
                {
                    // レベル2の高さを設定
                    level2.Elevation = level2HeightFeet;

                    //通芯
                    KUtil.CreateGrids(doc);

                    //外壁を作成
                    foreach (Curve curve in modelsetting.gaihekiCurves)
                    {
                        Wall wall = Wall.Create(doc, curve, outWallType.Id, level.Id, wallHeight, offset, false, false);
                    }
                    //内壁を作成
                    foreach (Curve curve in modelsetting.naihekiCurves)
                    {
                        Wall wall = Wall.Create(doc, curve, innerWallType.Id, level.Id, wallHeight, offset, false, false);
                    }
                    //床を作成
                    Floor floor = Floor.Create(doc, new List<CurveLoop> { floorLoop }, floorType.Id, level.Id);
                    //柱の作成
                    foreach (XYZ point in modelsetting.columnsPoints)
                    {
                        KUtil.CreateColumn(doc, levelName, levelName2, columnFamilyName, columnSymbolName, point.X, point.Y);//シンボルを探すこともメソッド内で実行
                    }
                    //窓の作成
                    foreach (XYZ point in modelsetting.windowsPoints)
                    {
                        KUtil.CreateWindowAndDoor(uidoc, doc, windowFamilyName, windowSymbolName, levelName, point.X, point.Y, windowOffset);//シンボルを探すこともメソッド内で実行
                    }
                    //ドアの作成
                    
                    foreach (XYZ point in modelsetting.doorsPoints)
                    {
                        KUtil.CreateWindowAndDoor(uidoc, doc, doorFamilyName, doorSymbolName, levelName, point.X, point.Y, 0);//シンボルを探すこともメソッド内で実行
                    }
                    
                }
                catch (Exception e)
                {
                    message = e.Message;
                    tx.RollBack();
                    return Result.Failed;
                }
                tx.Commit();
            }
            
            return Result.Succeeded;
        }    
        private void ModelOut(ModelSetting ms)
        {
            /*
            //Excelファイルに保存するためにファイル名をユーザーにきくためのダイアログを表示する
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "ExcelFiles | *.xls;*.xlsx;*.xlsm";
            saveFileDialog.Title = "EXCELファイル保存";
            saveFileDialog.ShowDialog();
            //ユーザーがファイル拡張子入れない場合にはxlsになってしまう。あとで開くときに警告が出るのであらかじめxlsxに変更しておく
            //filenameはユーザーが何も入れずにキャンセルボタンを押すと空になってしまう。次の次のif文で空のチェックしている
            String filename = saveFileDialog.FileName;
            //半角空白が含まれているとExcelアプリ起動時にファイルが見つからないエラーが出る場合あるので半角空白をアンダースコアに置き換える■
            filename = filename.Replace(" ", "_");
            if (filename.IndexOf("xlsx") < 0)
            {
                //SaveFileDialogオブジェクトのFileName属性のはうも変更しておく
                saveFileDialog.FileName = filename.Split('.')[0] + ".xlsx";
            }
            */
            //保存ファイル名が正常に入力されていればExcel保存を実行する
            String filename = @"C:\Users\katsu\Desktop\ModelCheck.xlsx";
            if (filename != "")
            {
                //ユーザーが指定したExcelファイル名でファイルストリームを取得する
                FileInfo newFile = new FileInfo(filename);
                //ExcelPackageを作成する
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    // ワークシートを追加
                    ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
                    int i = 1;
                    sheet.Cells[i ,1].Value = "階高"; sheet.Cells[i ,2].Value = KUtil.CVintTomm(ms.kaidaka);
                    i++;
                    sheet.Cells[i ,1].Value = "目盛り"; sheet.Cells[i ,2].Value = KUtil.CVintTomm(ms.memori);
                    i++;
                    //外壁
                    sheet.Cells[i ,1].Value = "外壁のCurve";
                    foreach (Curve x in ms.gaihekiCurves)
                    {
                        sheet.Cells[i ,2].Value = "(" + (x as Line).GetEndPoint(0).X + "," + (x as Line).GetEndPoint(0).Y + "," + (x as Line).GetEndPoint(0).Z + ")("
                            + (x as Line).GetEndPoint(1).X + "," + (x as Line).GetEndPoint(1).Y + "," + (x as Line).GetEndPoint(1).Z + ")";
                        i++;
                    }
                    //内壁
                    sheet.Cells[i ,1].Value = "内壁のCurve";
                    foreach (Curve x in ms.naihekiCurves)
                    {
                        sheet.Cells[i ,2].Value = "(" + (x as Line).GetEndPoint(0).X + "," + (x as Line).GetEndPoint(0).Y + "," + (x as Line).GetEndPoint(0).Z + ")("
                            + (x as Line).GetEndPoint(1).X + "," + (x as Line).GetEndPoint(1).Y + "," + (x as Line).GetEndPoint(1).Z + ")";
                        i++;
                    }
                    //床
                    sheet.Cells[i ,1].Value = "床のCurve";
                    foreach (var kvp in ms.floorCurveLoop)
                    {
                        sheet.Cells[i ,2].Value = "(" + kvp.Value.X + "," + kvp.Value.Y + "," + kvp.Value.Z + ")";
                        i++;
                    }
                    //ドア
                    sheet.Cells[i ,1].Value = "ドア";
                    foreach (XYZ x in ms.doorsPoints)
                    {
                        sheet.Cells[i ,2].Value = "(" + x.X + "," + x.Y + "," + x.Z + ")";
                        i++;
                    }
                    //窓
                    sheet.Cells[i ,1].Value = "窓";
                    foreach (XYZ x in ms.windowsPoints)
                    {
                        sheet.Cells[i ,2].Value = "(" + x.X + "," + x.Y + "," + x.Z + ")";
                        i++;
                    }
                    sheet.Cells[i ,1].Value = "外壁タイプ"; sheet.Cells[i,2].Value = ms.outWallTypeName;
                    i++;
                    sheet.Cells[i, 1].Value = "内壁タイプ"; sheet.Cells[i, 2].Value = ms.innerWallTypeName;
                    i++;
                    sheet.Cells[i, 1].Value = "床タイプ"; sheet.Cells[i, 2].Value = ms.floorTypeName;
                    i++;
                    sheet.Cells[i, 1].Value = "窓ファミリ名"; sheet.Cells[i, 2].Value = ms.windowFamilyName;
                    i++;
                    sheet.Cells[i, 1].Value = "窓タイプ"; sheet.Cells[i, 2].Value = ms.windowTypeName;
                    i++;
                    sheet.Cells[i, 1].Value = "ドアファミリ名"; sheet.Cells[i, 2].Value = ms.doorFamilyName;
                    i++;
                    sheet.Cells[i, 1].Value = "ドアタイプ"; sheet.Cells[i, 2].Value = ms.doorTypeName;
                    i++;

                    // 保存
                    package.Save();
                }
            }
        }
    }
    //ユーザーが入力した情報を保持するためのクラス
    public class ModelSetting
    {
        public double kaidaka { set; get; }
        public double memori { set; get; }
        public IList<Curve> gaihekiCurves { set; get; }//外壁のCurve
        public IList<Curve> naihekiCurves { set; get; }//内壁のCurve
        public SortedDictionary<int, XYZ> floorCurveLoop { set; get; }//床のcurveLoop
        public IList<XYZ> doorsPoints { set; get; }//１つのX,Y,Zの座標で表現
        public IList<XYZ> windowsPoints { set; get; }//１つのX,Y,Zの座標で表現
        public IList<XYZ> columnsPoints { set; get; }//１つのX,Y,Zの座標で表現
        public string outWallTypeName { set; get; }//外壁タイプ＝シンボル
        public string innerWallTypeName { set; get; }//内壁タイプ＝シンボル
        public string floorTypeName { set; get; }//床タイプ＝シンボル
        public string ceilingFamilyName { set; get; }//天井ファミリ名
        public string ceilingTypeName { set; get; }//天井タイプ＝シンボル
        public string windowFamilyName { set; get; }//窓ファミリ名
        public string windowTypeName { set; get; }//窓タイプ＝シンボル
        public string doorFamilyName { set; get; }//ドアファミリ名
        public string doorTypeName { set; get; }//ドアタイプ＝シンボル
        public string columnFamilyName { set; get; }//柱ファミリ名
        public string columnTypeName { set; get; }//柱タイプ＝シンボル

        public ModelSetting()
        {
        }
    }
}
