using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Windows.Forms;

namespace RevitAPITraining_DataOutputTXT
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //Сбор всех стен
            var walls = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            //Запись данных в строку
            string wallInfo = string.Empty;
            foreach (var wall in walls)
            {
                string wallVolume = wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsValueString();

                wallInfo += $"{wall.WallType.Name}\t{wallVolume}{Environment.NewLine}";
            }

            //Созранение файла
            var saveDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "wallInfo.csv",
                DefaultExt = "*.csv",
            };

            string selectedFilePath = string.Empty;
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveDialog.FileName;
            }
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                return Result.Cancelled;
            }
            File.WriteAllText(selectedFilePath, wallInfo);


            return Result.Succeeded;
        }
    }
}
