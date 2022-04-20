using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
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
        private double lengthMarginMeters;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            lengthMarginMeters = 0;
            var categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves));

            using (Transaction ts = new Transaction(doc, "Add Parameter"))
            {
                ts.Start();
                CreateSharedParameter(uiapp.Application, doc, "Длина с запасом", categorySet, BuiltInParameterGroup.PG_LENGTH, true);
                ts.Commit();
            }

            IList<Reference> selectedElementRefList = uidoc.Selection.PickObjects(ObjectType.Element, "Выберите трубы");

            foreach (var selectedElement in selectedElementRefList)
            {
                Element elem = doc.GetElement(selectedElement);
                if (elem.Category.Name == "Трубы")
                {
                    Parameter length = elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    if (length.StorageType == StorageType.Double)
                    {
                        double lengthMeters = UnitUtils.ConvertFromInternalUnits(length.AsDouble(), UnitTypeId.Meters);
                        lengthMarginMeters = lengthMeters * 1.1;

                        using (Transaction ts = new Transaction(doc, "Set Parameters"))
                        {
                            ts.Start();
                            Parameter lengthMarginParam = elem.LookupParameter("Длина с запасом");
                            lengthMarginParam.Set(lengthMarginMeters);
                            ts.Commit();
                        }
                    }
                }
                else
                {
                    TaskDialog.Show("Ошибка", "Выбранный элемент не относится к трубам");
                    return Result.Failed;
                }
            }
            return Result.Succeeded;
        }


        private void CreateSharedParameter(Application application, Document doc, string parameterName, CategorySet categorySet,
            BuiltInParameterGroup builtInParameterGroup, bool isInstance)
        {
            DefinitionFile definitionFile = application.OpenSharedParameterFile();
            if (definitionFile == null)
            {
                TaskDialog.Show("Ошибка", "не найден файл общих параметров");
                return;
            }
            Definition definition = definitionFile.Groups.
                SelectMany(group => group.Definitions).FirstOrDefault(def=>def.Name.Equals(parameterName));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }
            Binding binding = application.Create.NewTypeBinding(categorySet);
            if (isInstance)
                binding = application.Create.NewInstanceBinding(categorySet);

            BindingMap map = doc.ParameterBindings;
            map.Insert(definition, binding, builtInParameterGroup);
        }
    }
}
