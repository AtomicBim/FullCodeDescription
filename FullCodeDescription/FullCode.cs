using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace FullCodeDescription
{
    [Transaction(TransactionMode.Manual)]
    public class FullCode : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            // Получаем все элементы в проекте
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType(); // Получаем экземпляры, а не типы

            try
            {
                using (Transaction trans = new Transaction(doc, "Обновление Т_Полное наименование"))
                {
                    trans.Start();

                    foreach (Element elem in collector)
                    {
                        // Получаем тип элемента
                        ElementId typeId = elem.GetTypeId();
                        ElementType elemType = doc.GetElement(typeId) as ElementType;

                        if (elemType != null)
                        {
                            // Получаем значение параметра "Код по классификатору"
                            Parameter codeParam = elemType.get_Parameter(BuiltInParameter.UNIFORMAT_CODE);

                            // Получаем значение параметра "Описание по классификатору"
                            Parameter descParam = elemType.get_Parameter(BuiltInParameter.UNIFORMAT_DESCRIPTION);

                            if (codeParam != null && descParam != null && !string.IsNullOrEmpty(codeParam.AsString()))
                            {
                                // Получаем параметр Т_Полное наименование
                                Parameter fullNameParam = elemType.LookupParameter("Т_Полное наименование");

                                if (fullNameParam != null && fullNameParam.StorageType == StorageType.String)
                                {
                                    string code = codeParam.AsString();
                                    string description = descParam.AsString();

                                    // Формируем полное наименование
                                    string fullName = string.IsNullOrEmpty(description)
                                        ? code
                                        : $"{code}_{description}";

                                    // Записываем значение
                                    fullNameParam.Set(fullName);
                                }
                            }
                        }
                    }

                    trans.Commit();
                }

                TaskDialog.Show("Успех", "Параметры успешно обновлены");
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Ошибка", $"Произошла ошибка при выполнении скрипта:\n{ex.Message}");
                return Result.Failed;
            }
        }

    }
}
