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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves));

            using (Transaction ts = new Transaction(doc, "Add parametr"))
            {
                ts.Start();
                CreateSharedParameter(uiapp.Application, doc, "Наименование", categorySet, BuiltInParameterGroup.PG_DATA, true);
                ts.Commit();
            }
            List<Pipe> pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();
            foreach (var selectedElement in pipes)
            {
                Element element = doc.GetElement(selectedElement.Id);
                if (element is Pipe)
                {
                    Parameter outerDiamParameter = element.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    Parameter innerDiamParameter = element.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                    if (outerDiamParameter.StorageType == StorageType.Double && innerDiamParameter.StorageType == StorageType.Double)
                    {
                        double outerDiamValue = UnitUtils.ConvertFromInternalUnits(outerDiamParameter.AsDouble(), UnitTypeId.Millimeters);
                        double innerDiamValue = UnitUtils.ConvertFromInternalUnits(innerDiamParameter.AsDouble(), UnitTypeId.Millimeters);

                        using (Transaction ts1 = new Transaction(doc, "Set parameter"))
                        {
                            ts1.Start();
                            var pipe = element as Pipe;
                            Parameter name = pipe.LookupParameter("Наименование");
                            string output = $"{outerDiamValue}/{innerDiamParameter}";
                            name.Set(output);
                            ts1.Commit();
                        }
                    }

                }
            }

            TaskDialog.Show("Успешно", message);
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
