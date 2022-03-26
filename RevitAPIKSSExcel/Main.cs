using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPIKSSExcel
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;


            //Выбрали все помещения
            string roomInfo = string.Empty;
            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            ////Создали файл экселя и записали в него данные
            //string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "rooms.xlsx");
            //using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            //{
            //    IWorkbook workbook = new XSSFWorkbook();
            //    ISheet sheet = workbook.CreateSheet("Data");

            //    int rowIndex = 0;
            //    foreach (var room in rooms)
            //    {
            //        sheet.SetCellValue(rowIndex, columnIndex: 0, room.Name);
            //        sheet.SetCellValue(rowIndex, columnIndex: 1, room.Number);
            //        sheet.SetCellValue(rowIndex, columnIndex: 2, room.Area);
            //        rowIndex++;
            //    }
            //    workbook.Write(stream);
            //    workbook.Close();
            //}

            ////Запускаем файл экселя автоматически
            //System.Diagnostics.Process.Start(excelPath);






            ////Открываем диалоговое окно для открытия файла Excel
            //System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog
            //{
            //    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            //    Filter = "Excel files(*.xlsx) | *.xlsx"
            //};
            //string filePath = string.Empty;
            //if(openFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    filePath = openFileDialog1.FileName;
            //}
            //if(string.IsNullOrEmpty(filePath))
            //{
            //    return Result.Cancelled;
            //}

            ////Читаем данные из файла Excel
            //using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            //{
            //    IWorkbook workbook = new XSSFWorkbook(filePath);
            //    ISheet sheet = workbook.GetSheetAt(index: 0);

            //    int rowIndex = 0;
            //    while (sheet.GetRow(rowIndex) != null)
            //    {
            //        if (sheet.GetRow(rowIndex).GetCell(0) == null ||
            //            sheet.GetRow(rowIndex).GetCell(1) == null)
            //        {
            //            rowIndex++;
            //            continue;
            //        }
            //        string name = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
            //        string number = sheet.GetRow(rowIndex).GetCell(1).StringCellValue;

            //        var room = rooms.FirstOrDefault(r => r.Number.Equals(number));

            //        if (room == null)
            //        {
            //            rowIndex++;
            //            continue;
            //        }    

            //        using (var ts = new Transaction(doc, "Set parameter"))
            //        {
            //            ts.Start();
            //            room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(name);
            //            ts.Commit();
            //        }

            //        rowIndex++;
            //    }
            //}




            ////Запись данных в файл JSON
            //var roomDataList = new List<RoomData>();
            //foreach (var room in rooms)
            //{
            //    roomDataList.Add(new RoomData
            //    {
            //        Name = room.Name,
            //        Number = room.Number
            //    });
            //}

            //string json = JsonConvert.SerializeObject(roomDataList, Formatting.Indented);
            //File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "data.json"), json);


            //Чтение данных из файла JSON
            System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "Json files(*.json) | *.json"
            };
            string filePath = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
            }
            if (string.IsNullOrEmpty(filePath))
            {
                return Result.Cancelled;
            }

            string json = File.ReadAllText(filePath);
            List<RoomData> roomDataList = JsonConvert.DeserializeObject<List<RoomData>>(json);

            while (roomDataList != null)
            {
                //JObject jsonObject = JObject.Parse(json);
                var result = JObject.Parse(json);
                var name = result.SelectToken("Name").ToString();
                var number = result.SelectToken("Number").ToString();
                

                var room = rooms.FirstOrDefault(r => r.Number.Equals(number));

                if (room == null)
                {
                    continue;
                }

                using (var ts = new Transaction(doc, "Set parameter"))
                {
                    ts.Start();
                    room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(name);
                    ts.Commit();
                }
            }

            return Result.Succeeded;
        }
    }
}
