using System;
using System.Collections.Generic;
using System.Configuration;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using Serilog;
using Serilog.Core;

namespace UpgradedExcelVlookup
{
    internal class Program
    {
        static Logger logFile;

        static void Main(string[] args)
        {
            string logFileName = CreateLogFile();
            if (logFileName == null) { Environment.Exit(1); }

            BuildConfig();

            logFile.Information("Получение путей к файлам...");
            string pathToFirst = ConfigurationManager.AppSettings.Get("First");
            string pathToSecond = ConfigurationManager.AppSettings.Get("Second");

            logFile.Information("Получение диапазонов значений...");
            string firstValueRange = ConfigurationManager.AppSettings.Get("FirstRange");
            string secondValueRange = ConfigurationManager.AppSettings.Get("SecondRange");

            WorkingWithData(pathToFirst, pathToSecond, firstValueRange, secondValueRange);

            Console.WriteLine("Программа завершила свою работу.\nЖурнал событий находится по пути: " + logFileName);
        }

        static string CreateLogFile()
        {
            try
            {
                string pathToExe = Environment.CurrentDirectory.ToString();
                string pathToLogFile = pathToExe + "\\Logs";
                string logFileName = pathToLogFile + "\\Log_" + GeActualDateTime() + ".txt";

                logFile = new LoggerConfiguration().MinimumLevel.Debug().WriteTo.File(logFileName).CreateLogger();

                logFile.Information("Журнал событий создан и находится по пути: " + logFileName);

                return logFileName;
            }
            catch(Exception e)
            {
                Console.WriteLine("Возникла ошибка при создании журнала событий: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Возможно, отстуствует папка 'Logs' ");

                return null;
            }
        }

        static string GeActualDateTime()
        {
            string dateTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss").Replace(":", "_");
            return dateTime;
        }

        static void BuildConfig()
        {
            Console.WriteLine("Сгенерировать пути к excel-файлам и диапазоны по умолчанию? Y/n ");

            while (true)
            {
                string generateAppConfigChecker = Console.ReadLine();

                if (generateAppConfigChecker == "Y" || generateAppConfigChecker == "y")
                {
                    logFile.Information("Выбрано создание путей к excel-файлам и диапазонов значений по умолчанию...");

                    string pathToExe = Environment.CurrentDirectory;

                    string pathToFirst = pathToExe + "\\Exceles\\First.xlsx";
                    string pathToSecond = pathToExe + "\\Exceles\\Second.xlsx";

                    ConfigurationManager.AppSettings.Set("First", pathToFirst);
                    ConfigurationManager.AppSettings.Set("Second", pathToSecond);
                    ConfigurationManager.AppSettings.Set("FirstRange", "Лист1!B11:N30");
                    ConfigurationManager.AppSettings.Set("SecondRange", "Лист1!C11:O30");

                    break;
                }
                else if (generateAppConfigChecker == "N" || generateAppConfigChecker == "n") 
                {
                    logFile.Information("Выбрано создание путей к excel-файлам и диапазонов значений из файла конфигураций...");
                    break; 
                }
                else
                {
                    logFile.Error("Введены некоректные данные при выборе путей к Excel-файлам и диапозонов значений");
                    Console.WriteLine("Введите коректные данные!");
                    continue;
                }
            }
        }

        static void WorkingWithData(string pathToFirst, string pathToSecond, string firstValueRange, string secondValueRange)
        {
            logFile.Information("Загрузка данных из Excel-файлов...");
            var firstData = ExcelDataLoad(pathToFirst, firstValueRange);
            var secondtData = ExcelDataLoad(pathToSecond, secondValueRange);

            logFile.Information("Сравнение данных из Excel-файлов...");
            ExcelDataComparison(firstData, secondtData);
        }

        static List<List<string>> ExcelDataLoad(string filePath, string valueRange)
        {
            var fileData = new List<List<string>>();

            using (var excelPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                var mainList = excelPackage.Workbook.Worksheets.FirstOrDefault();

                if (mainList != null)
                {
                    if (valueRange != null)
                    {
                        var range = mainList.Cells[valueRange];
                        var rows = range.Value as object[,];

                        if (rows != null)
                        {
                            for (int row = 0; row < rows.GetLength(0); row++)
                            {
                                var rowData = new List<string>();

                                for (int col = 0; col < rows.GetLength(1); col++)
                                {
                                    // На данном этапе я решил заменить все пустые значения на null, чтобы отслеживать состояние на этапе разработки, да и для удобного вывода
                                    var cellValue = rows[row, col]?.ToString() ?? "null";
                                    rowData.Add(cellValue);
                                }

                                fileData.Add(rowData);
                            }
                        }
                        else
                        {
                            logFile.Error("Не удалось загрузить данные из диапазона");
                            Console.WriteLine("Ошибка: Не удалось загрузить данные из диапазона.");
                        }
                    }
                    else
                    {
                        logFile.Error("Не указан диапазон значений");
                        Console.WriteLine("Ошибка: Не указан диапазон значений.");
                    }
                }
                else
                {
                    logFile.Error("Не найден лист в файле Excel");
                    Console.WriteLine("[] Ошибка: Не найден лист в файле Excel.");
                }
                logFile.Information("Данные из Excel-файла " + filePath + " успешно загружены");
            }

            return fileData;
        }

        static void ExcelDataComparison(List<List<string>> firstData, List<List<string>> secondData)
        {
            if (firstData.Count > 0)
            {
                if (secondData.Count > 0)
                {
                    int reportCount = 0;

                    for (int i = 0; i < firstData.Count; i++)
                    {
                        var firstRow = firstData[i];
                        var valueToCompare = firstRow[0];

                        var secondRow = secondData.FirstOrDefault(row => row[0] == valueToCompare);

                        if (secondRow != null)
                        {

                            for (int j = 1; j < firstRow.Count; j++)
                            {
                                if (firstRow[j] != secondRow[j])
                                {
                                    reportCount++;

                                    // тут я не стал уже заморачиваться и искать как указать фактическое значение столбца в Excel, это лишь займет больше времени
                                    string report = $"В строке первого столбца '{valueToCompare}': ";
                                    report += $"столбец {j}(от начала столбца): {firstRow[j]} -> {secondRow[j]}, ";

                                    logFile.Information(report);
                                }
                            }
                        }
                        else
                        {
                            logFile.Error("Соответствующая строка из второго файла отстуствует. Номер строки первого файла: " + (i + 1));
                        }
                    }
                    if (reportCount < 1)
                    {
                        logFile.Information("Различия в файле не обнаружены");
                    }
                }
                else
                {
                    logFile.Error("Отсутствуют данные из второго файла");
                }
            }
            else
            {
                logFile.Error("Отсутствуют данные из первого файла");
            }
        }
    }
}
