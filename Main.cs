using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIFirstApp
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            IList<Reference> selectedElementRefList = uidoc.Selection.PickObjects(ObjectType.Face, "Выберите несколько стен по грани");

            var WallList = new List<Wall>();
            double Value = 0;
            double Sum = 0;

            foreach (var selectedElement in selectedElementRefList)
            {
                Element element = doc.GetElement(selectedElement);

                if (element is Wall)
                {
                    Wall owall = (Wall)element;
                    Parameter volumeParameter = element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                    if (volumeParameter.StorageType == StorageType.Double)
                    {
                        Value = UnitUtils.ConvertFromInternalUnits(volumeParameter.AsDouble(), UnitTypeId.CubicMeters);
                        TaskDialog.Show("Объём i-ой стены ", Value.ToString());
                    }
                    Sum += Value;
                }
            }
            TaskDialog.Show("Сумма объёмов стен ", Sum.ToString());

            return Result.Succeeded;
        }
    }
}
