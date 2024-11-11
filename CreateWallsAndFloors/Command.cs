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
//���\�t�g�ɂ�EPPlus4.5.3.3���g�p���Ă��܂��B�����LGPL���C�Z���X�ł��B���쌠��EPPlus Software�Ђł��B
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


            //�Ǝ���ModelSetting�N���X���C���X�^���X���B�ݒ����ێ�����
            ModelSetting modelsetting = new ModelSetting();

            //Excel����ݒ����ǂށB�ǂ̍��W�Ȃǂ��ǂ�Őݒ肷��
            bool result = KUtil.ReadInputInfo(out modelsetting);

            //Excel���͂ɃG���[���������Ƃ��͏I��
            if(!result) return Result.Failed;

            //�t�@�~�����̐ݒ�̓ǂݏo��
            double wallHeight = modelsetting.kaidaka;
            double offset = 0;//������̃I�t�Z�b�g�i�ǁj
            double windowOffset = KUtil.CVmmToInt(1000);//������̃I�t�Z�b�g�i���jmm
            string outWallTypeName = modelsetting.outWallTypeName;
            string innerWallTypeName = modelsetting.innerWallTypeName;
            string floorTypeName = modelsetting.floorTypeName;
            string levelName = @"���x�� 1";
            string levelName2 = @"���x�� 2";
            double level2Height = 4500;//mm
            double level2HeightFeet = UnitUtils.ConvertToInternalUnits(level2Height, UnitTypeId.Millimeters);

            string windowFamilyName = modelsetting.windowFamilyName;
            string windowSymbolName = modelsetting.windowTypeName;
            string doorFamilyName = modelsetting.doorFamilyName;
            string doorSymbolName = modelsetting.doorTypeName;
            string columnFamilyName = modelsetting.columnFamilyName;
            string columnSymbolName = modelsetting.columnTypeName;

            //�ړI�̃��x���i�K�j���擾����
            Level level = KUtil.GetLevel(doc, levelName);//1�K
            Level level2 = KUtil.GetLevel(doc, levelName2);//2�K

            //�O�ǁA���ǁA���^�C�v�̌���
            WallType outWallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault(q => q.Name == outWallTypeName);
            WallType innerWallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault(q => q.Name == innerWallTypeName);
            FloorType floorType = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().FirstOrDefault(q => q.Name == floorTypeName);

            //���̂��߂̃��[�v�v���t�@�C���i�������Ȃ������́j�����
            CurveLoop floorLoop = KUtil.FloorProfile(modelsetting);

            //�f�[�^�`�F�b�N
            //ModelOut(modelsetting);

            //���f���̐���
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("TransactionWallFloor");
                try
                {
                    // ���x��2�̍�����ݒ�
                    level2.Elevation = level2HeightFeet;

                    //�ʐc
                    KUtil.CreateGrids(doc);

                    //�O�ǂ��쐬
                    foreach (Curve curve in modelsetting.gaihekiCurves)
                    {
                        Wall wall = Wall.Create(doc, curve, outWallType.Id, level.Id, wallHeight, offset, false, false);
                    }
                    //���ǂ��쐬
                    foreach (Curve curve in modelsetting.naihekiCurves)
                    {
                        Wall wall = Wall.Create(doc, curve, innerWallType.Id, level.Id, wallHeight, offset, false, false);
                    }
                    //�����쐬
                    Floor floor = Floor.Create(doc, new List<CurveLoop> { floorLoop }, floorType.Id, level.Id);
                    //���̍쐬
                    foreach (XYZ point in modelsetting.columnsPoints)
                    {
                        KUtil.CreateColumn(doc, levelName, levelName2, columnFamilyName, columnSymbolName, point.X, point.Y);//�V���{����T�����Ƃ����\�b�h���Ŏ��s
                    }
                    //���̍쐬
                    foreach (XYZ point in modelsetting.windowsPoints)
                    {
                        KUtil.CreateWindowAndDoor(uidoc, doc, windowFamilyName, windowSymbolName, levelName, point.X, point.Y, windowOffset);//�V���{����T�����Ƃ����\�b�h���Ŏ��s
                    }
                    //�h�A�̍쐬
                    
                    foreach (XYZ point in modelsetting.doorsPoints)
                    {
                        KUtil.CreateWindowAndDoor(uidoc, doc, doorFamilyName, doorSymbolName, levelName, point.X, point.Y, 0);//�V���{����T�����Ƃ����\�b�h���Ŏ��s
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
            //Excel�t�@�C���ɕۑ����邽�߂Ƀt�@�C���������[�U�[�ɂ������߂̃_�C�A���O��\������
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "ExcelFiles | *.xls;*.xlsx;*.xlsm";
            saveFileDialog.Title = "EXCEL�t�@�C���ۑ�";
            saveFileDialog.ShowDialog();
            //���[�U�[���t�@�C���g���q����Ȃ��ꍇ�ɂ�xls�ɂȂ��Ă��܂��B���ƂŊJ���Ƃ��Ɍx�����o��̂ł��炩����xlsx�ɕύX���Ă���
            //filename�̓��[�U�[���������ꂸ�ɃL�����Z���{�^���������Ƌ�ɂȂ��Ă��܂��B���̎���if���ŋ�̃`�F�b�N���Ă���
            String filename = saveFileDialog.FileName;
            //���p�󔒂��܂܂�Ă����Excel�A�v���N�����Ƀt�@�C����������Ȃ��G���[���o��ꍇ����̂Ŕ��p�󔒂��A���_�[�X�R�A�ɒu�������遡
            filename = filename.Replace(" ", "_");
            if (filename.IndexOf("xlsx") < 0)
            {
                //SaveFileDialog�I�u�W�F�N�g��FileName�����̂͂����ύX���Ă���
                saveFileDialog.FileName = filename.Split('.')[0] + ".xlsx";
            }
            */
            //�ۑ��t�@�C����������ɓ��͂���Ă����Excel�ۑ������s����
            String filename = @"C:\Users\katsu\Desktop\ModelCheck.xlsx";
            if (filename != "")
            {
                //���[�U�[���w�肵��Excel�t�@�C�����Ńt�@�C���X�g���[�����擾����
                FileInfo newFile = new FileInfo(filename);
                //ExcelPackage���쐬����
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    // ���[�N�V�[�g��ǉ�
                    ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
                    int i = 1;
                    sheet.Cells[i ,1].Value = "�K��"; sheet.Cells[i ,2].Value = KUtil.CVintTomm(ms.kaidaka);
                    i++;
                    sheet.Cells[i ,1].Value = "�ڐ���"; sheet.Cells[i ,2].Value = KUtil.CVintTomm(ms.memori);
                    i++;
                    //�O��
                    sheet.Cells[i ,1].Value = "�O�ǂ�Curve";
                    foreach (Curve x in ms.gaihekiCurves)
                    {
                        sheet.Cells[i ,2].Value = "(" + (x as Line).GetEndPoint(0).X + "," + (x as Line).GetEndPoint(0).Y + "," + (x as Line).GetEndPoint(0).Z + ")("
                            + (x as Line).GetEndPoint(1).X + "," + (x as Line).GetEndPoint(1).Y + "," + (x as Line).GetEndPoint(1).Z + ")";
                        i++;
                    }
                    //����
                    sheet.Cells[i ,1].Value = "���ǂ�Curve";
                    foreach (Curve x in ms.naihekiCurves)
                    {
                        sheet.Cells[i ,2].Value = "(" + (x as Line).GetEndPoint(0).X + "," + (x as Line).GetEndPoint(0).Y + "," + (x as Line).GetEndPoint(0).Z + ")("
                            + (x as Line).GetEndPoint(1).X + "," + (x as Line).GetEndPoint(1).Y + "," + (x as Line).GetEndPoint(1).Z + ")";
                        i++;
                    }
                    //��
                    sheet.Cells[i ,1].Value = "����Curve";
                    foreach (var kvp in ms.floorCurveLoop)
                    {
                        sheet.Cells[i ,2].Value = "(" + kvp.Value.X + "," + kvp.Value.Y + "," + kvp.Value.Z + ")";
                        i++;
                    }
                    //�h�A
                    sheet.Cells[i ,1].Value = "�h�A";
                    foreach (XYZ x in ms.doorsPoints)
                    {
                        sheet.Cells[i ,2].Value = "(" + x.X + "," + x.Y + "," + x.Z + ")";
                        i++;
                    }
                    //��
                    sheet.Cells[i ,1].Value = "��";
                    foreach (XYZ x in ms.windowsPoints)
                    {
                        sheet.Cells[i ,2].Value = "(" + x.X + "," + x.Y + "," + x.Z + ")";
                        i++;
                    }
                    sheet.Cells[i ,1].Value = "�O�ǃ^�C�v"; sheet.Cells[i,2].Value = ms.outWallTypeName;
                    i++;
                    sheet.Cells[i, 1].Value = "���ǃ^�C�v"; sheet.Cells[i, 2].Value = ms.innerWallTypeName;
                    i++;
                    sheet.Cells[i, 1].Value = "���^�C�v"; sheet.Cells[i, 2].Value = ms.floorTypeName;
                    i++;
                    sheet.Cells[i, 1].Value = "���t�@�~����"; sheet.Cells[i, 2].Value = ms.windowFamilyName;
                    i++;
                    sheet.Cells[i, 1].Value = "���^�C�v"; sheet.Cells[i, 2].Value = ms.windowTypeName;
                    i++;
                    sheet.Cells[i, 1].Value = "�h�A�t�@�~����"; sheet.Cells[i, 2].Value = ms.doorFamilyName;
                    i++;
                    sheet.Cells[i, 1].Value = "�h�A�^�C�v"; sheet.Cells[i, 2].Value = ms.doorTypeName;
                    i++;

                    // �ۑ�
                    package.Save();
                }
            }
        }
    }
    //���[�U�[�����͂�������ێ����邽�߂̃N���X
    public class ModelSetting
    {
        public double kaidaka { set; get; }
        public double memori { set; get; }
        public IList<Curve> gaihekiCurves { set; get; }//�O�ǂ�Curve
        public IList<Curve> naihekiCurves { set; get; }//���ǂ�Curve
        public SortedDictionary<int, XYZ> floorCurveLoop { set; get; }//����curveLoop
        public IList<XYZ> doorsPoints { set; get; }//�P��X,Y,Z�̍��W�ŕ\��
        public IList<XYZ> windowsPoints { set; get; }//�P��X,Y,Z�̍��W�ŕ\��
        public IList<XYZ> columnsPoints { set; get; }//�P��X,Y,Z�̍��W�ŕ\��
        public string outWallTypeName { set; get; }//�O�ǃ^�C�v���V���{��
        public string innerWallTypeName { set; get; }//���ǃ^�C�v���V���{��
        public string floorTypeName { set; get; }//���^�C�v���V���{��
        public string ceilingFamilyName { set; get; }//�V��t�@�~����
        public string ceilingTypeName { set; get; }//�V��^�C�v���V���{��
        public string windowFamilyName { set; get; }//���t�@�~����
        public string windowTypeName { set; get; }//���^�C�v���V���{��
        public string doorFamilyName { set; get; }//�h�A�t�@�~����
        public string doorTypeName { set; get; }//�h�A�^�C�v���V���{��
        public string columnFamilyName { set; get; }//���t�@�~����
        public string columnTypeName { set; get; }//���^�C�v���V���{��

        public ModelSetting()
        {
        }
    }
}
