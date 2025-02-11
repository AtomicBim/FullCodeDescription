using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace FullCodeDescription
{
    [Transaction(TransactionMode.Manual)]
    public class CopyCode : IExternalCommand
    {
        public class TypeInfo
        {
            public string Category { get; set; }
            public string TypeName { get; set; }
            public string Code { get; set; }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            var typeInfos = new List<TypeInfo>();

            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .WhereElementIsElementType();

                foreach (ElementType elemType in collector)
                {
                    Parameter codeParam = elemType.get_Parameter(BuiltInParameter.UNIFORMAT_CODE);
                    if (codeParam != null && !string.IsNullOrEmpty(codeParam.AsString()))
                    {
                        TypeInfo info = new TypeInfo
                        {
                            Category = elemType.Category?.Name ?? "Без категории",
                            TypeName = elemType.Name,
                            Code = codeParam.AsString()
                        };
                        typeInfos.Add(info);
                        Debug.WriteLine($"Export: Category={info.Category}, Type={info.TypeName}, Code={info.Code}");
                    }
                }

                string directory = doc.PathName != null
                    ? Path.GetDirectoryName(doc.PathName)
                    : Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                string jsonPath = Path.Combine(directory, $"{doc.Title}_TypeCodes.json");
                string jsonContent = JsonConvert.SerializeObject(typeInfos, Formatting.Indented);
                File.WriteAllText(jsonPath, jsonContent);

                Debug.WriteLine($"Exported {typeInfos.Count} items to {jsonPath}");
                TaskDialog.Show("Успешно", $"Экспортировано {typeInfos.Count} элементов\nПуть: {jsonPath}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.Message);
                return Result.Failed;
            }
        }
    }
}