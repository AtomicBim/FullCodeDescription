using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using Microsoft.Win32;

namespace FullCodeDescription
{
    [Transaction(TransactionMode.Manual)]
    public class PasteCode : IExternalCommand
    {
        public class TypeInfo
        {
            public string Category { get; set; }
            public string TypeName { get; set; }
            public string Code { get; set; }
        }

        private string CreateKey(string category, string typeName)
        {
            return $"{category}|{typeName}".ToLower();
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            int successCount = 0;
            int notFoundCount = 0;
            var problemElements = new List<(string Element, string Reason)>();

            try
            {
                // Выбор файла
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    FilterIndex = 1
                };

                if (openFileDialog.ShowDialog() != true)
                {
                    TaskDialog.Show("Отмена", "Операция отменена пользователем");
                    return Result.Cancelled;
                }

                // Чтение JSON
                string jsonContent = File.ReadAllText(openFileDialog.FileName);
                var typeInfos = JsonConvert.DeserializeObject<List<TypeInfo>>(jsonContent);

                if (typeInfos == null || !typeInfos.Any())
                {
                    TaskDialog.Show("Ошибка", "Файл пуст или имеет неверный формат");
                    return Result.Failed;
                }

                Debug.WriteLine($"Loaded {typeInfos.Count} items from JSON");

                // Создаем словарь для поиска, обрабатывая возможные дубликаты
                var codeMap = new Dictionary<string, TypeInfo>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in typeInfos)
                {
                    string key = CreateKey(item.Category, item.TypeName);
                    if (!codeMap.ContainsKey(key))
                    {
                        codeMap.Add(key, item);
                    }
                    else
                    {
                        Debug.WriteLine($"Duplicate key found: {key}");
                        // Можно также добавить в отчет информацию о дубликатах
                    }
                }

                // Вывод первых 10 ключей из JSON для отладки
                Debug.WriteLine("First 10 keys in JSON:");
                foreach (var key in codeMap.Keys.Take(10))
                {
                    Debug.WriteLine($"JSON key: {key}");
                }

                // Получаем все типы
                var collector = new FilteredElementCollector(doc).WhereElementIsElementType();

                using (Transaction trans = new Transaction(doc, "Обновление кодов"))
                {
                    trans.Start();

                    foreach (ElementType elemType in collector)
                    {
                        try
                        {
                            string category = elemType.Category?.Name ?? "Без категории";
                            string typeName = elemType.Name;
                            string key = CreateKey(category, typeName);
                            Debug.WriteLine($"Processing: {key}");

                            if (codeMap.TryGetValue(key, out TypeInfo matchingType))
                            {
                                Parameter codeParam = elemType.get_Parameter(BuiltInParameter.UNIFORMAT_CODE);

                                if (codeParam == null)
                                {
                                    problemElements.Add((key, "Параметр не найден"));
                                    continue;
                                }

                                if (codeParam.IsReadOnly)
                                {
                                    problemElements.Add((key, "Параметр не доступен для записи"));
                                    continue;
                                }

                                string currentValue = codeParam.AsString();
                                if (!string.Equals(currentValue, matchingType.Code, StringComparison.OrdinalIgnoreCase))
                                {
                                    try
                                    {
                                        codeParam.Set(matchingType.Code);
                                        string newValue = codeParam.AsString();

                                        if (string.Equals(newValue, matchingType.Code, StringComparison.OrdinalIgnoreCase))
                                        {
                                            successCount++;
                                            Debug.WriteLine($"Successfully updated: {key} -> {matchingType.Code}");
                                        }
                                        else
                                        {
                                            problemElements.Add((key, $"Не удалось установить значение. Текущее: {newValue}, Ожидаемое: {matchingType.Code}"));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        problemElements.Add((key, $"Ошибка при установке значения: {ex.Message}"));
                                    }
                                }
                                else
                                {
                                    Debug.WriteLine($"Value already correct for: {key}");
                                }
                            }
                            else
                            {
                                notFoundCount++;
                                Debug.WriteLine($"No match found for: {key}");
                            }
                        }
                        catch (Exception ex)
                        {
                            string elementInfo = $"{elemType.Category?.Name ?? "Без категории"}|{elemType.Name}";
                            problemElements.Add((elementInfo, ex.Message));
                            Debug.WriteLine($"Error processing element {elementInfo}: {ex.Message}");
                        }
                    }

                    trans.Commit();
                }

                // Формируем сообщение с результатами
                string resultMessage = $"Обработка завершена:\n\n" +
                                     $"Успешно обновлено: {successCount}\n" +
                                     $"Не найдено соответствий: {notFoundCount}";

                if (problemElements.Any())
                {
                    resultMessage += $"\n\nПроблемные элементы ({problemElements.Count}):\n";
                    foreach (var (elem, reason) in problemElements.Take(10))
                    {
                        resultMessage += $"- {elem}: {reason}\n";
                    }
                    if (problemElements.Count > 10)
                    {
                        resultMessage += "...";
                    }
                }

                TaskDialog.Show("Результаты", resultMessage);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Ошибка", $"Произошла ошибка:\n{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
