using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using OfficeOpenXml;

namespace XML_Data_Reader
{
    class Program
    {
        static void Main(string[] args)
        {
            //Ustawienie kontekstu licencji dla pakietu Excel
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //Pobranie ścieżki do XML referencyjnego
            Console.WriteLine("Proszę podać ścieżkę do pliku xml urządzenia referencyjnego");
            string referencePath = Console.ReadLine();

            //Pobranie ścieżki do folderu z XML do analizy
            Console.WriteLine("Proszę podać ścieżkę do folderu z plikami xml do sprawdzenia");
            string folderPath = Console.ReadLine();
            string[] xmlFiles = Directory.GetFiles(folderPath, "*.xml");

            //Utworzenie ścieżki wyjściowej dla pliku Excel
            string outputPath = Path.Combine(folderPath, "output.xlsx");
            
            //Wybór typu wyświetlania danych
            Console.WriteLine($"Aby otrzymać pełne dane wpisz 'full', aby otrzymać tylko parametry z różniącymi się wartościami wpisz 'diff'");
            string displayType = Console.ReadLine();

            //Przetworzenie pliku XML referencyjnego i przechowanie danych w słowniku
            Dictionary<int, ParameterInfo> referenceParameters = ProcessXmlFile(referencePath);

            //Inicjalizacja pakietu Excel
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Dane");

                //Ustawienie indeksu wiersza referencyjnego
                int rowReferenceIndex = 2;
                worksheet.Cells[1, 3].Value = referenceParameters[31318].Value;
                foreach (var parameter in referenceParameters)
                {
                    //Wypełnienie arkusza danymi z pliku referencyjnego
                    worksheet.Cells[rowReferenceIndex, 1].Value = parameter.Key;
                    worksheet.Cells[rowReferenceIndex, 2].Value = parameter.Value.Name;
                    worksheet.Cells[rowReferenceIndex, 3].Value = parameter.Value.Value;
                    rowReferenceIndex++;
                }

                //Analiza danych zgodnie z wybranym typem wyświetlania
                if (displayType == "full")
                {
                    ProcessXmlToWorksheet(worksheet, xmlFiles, referencePath, referenceParameters);
                    
                }
                else if (displayType == "diff")
                {
                    ProcessXmlToWorksheet(worksheet, xmlFiles, referencePath, referenceParameters);

                    //Usuwanie wierszy z takimi samymi ustawieniami we wszystkich plikach
                    for(int row = worksheet.Dimension.End.Row; row >= 2; row--)
                    {
                        if(AllCellsInRowEqual(worksheet, row))
                        {
                            worksheet.DeleteRow(row);
                            //Console.WriteLine(AllCellsInRowEqual(worksheet, row));
                        }
                    }

                }
                else
                {
                    Console.WriteLine("Wrong input");
                    return;
                }

                var fileInfo = new FileInfo(outputPath);
                package.SaveAs(fileInfo);
                
            }
            //Informacje o zakończeniu eksportu danych do pliku Excel
            Console.WriteLine("Export to .xls finished");
            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
            System.Diagnostics.Process.Start(outputPath);

        }
        //Klasa zawierająca informacje o parametrze
        class ParameterInfo
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }

        //Metoda do przetwarzania pliku XMl i przechowania danych w słowniku
        static Dictionary<int, ParameterInfo> ProcessXmlFile(string xmlFilePath)
        {
            Dictionary<int, ParameterInfo> parameters = new Dictionary<int, ParameterInfo>();

            try
            {
                //Wczytanie pliku XML
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlFilePath);

                //Wyodrębnienie węzłow parametrów z pliku XML
                XmlNodeList parameterNodes = xmlDoc.SelectNodes("//chanelParameter/ArrayOfParameter/Parameter");

                //Iteracja przez węzły parametrów i zapisanie informacji do słownika
                foreach (XmlNode parameterNode in parameterNodes)
                {
                    int index;
                    if (int.TryParse(parameterNode.SelectSingleNode("index").InnerText, out index))
                    {
                        string name = parameterNode.SelectSingleNode("name").InnerText;
                        string value = parameterNode.SelectSingleNode("value")?.InnerText;

                        parameters[index] = new ParameterInfo { Name = name, Value = value };
                    
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return parameters;
        }
        //Metoda do przetwarzania plików XML i zapisywania danych do arkusza Excel
        static void ProcessXmlToWorksheet(ExcelWorksheet worksheet, string[] xmlFiles, string referencePath, Dictionary<int,ParameterInfo> referenceParameters)
        {
            //Wskazanie początkowej komórki do której wprowadzane mają być dane
            int columnIndex = 4;
            int rowIndex = 2;
            
            //Iteracja przez wszystkie pliki XML znajdujące się we wskazanym folderze 
            foreach (string xmlFilePath in xmlFiles.Where(filePath => filePath != referencePath))
            {
                Dictionary<int, ParameterInfo> parameters = ProcessXmlFile(xmlFilePath);

                //Zapis wartości o konkretnym kluczu w wierszu nagłówkowym
                worksheet.Cells[1, columnIndex].Value = parameters[31318].Value;
                
                //Iteracja przez klucze parametrów pliku referencyjnego
                foreach (var refKey in referenceParameters)
                {
                    //Sprawdzenie czy plik zawiera dany klucz
                    if (parameters.ContainsKey(refKey.Key))
                    {
                        //Zapis wartości parametru do odpowiedniej komórki w arkuszu
                        ParameterInfo parameter = parameters[refKey.Key];
                        worksheet.Cells[rowIndex, columnIndex].Value = parameter.Value;
                        rowIndex++;
                    }
                    else
                    {
                        //Zapis informacji o brakujących danych, jeśli w pliku nie było parametru o podanym kluczu
                        worksheet.Cells[rowIndex, columnIndex].Value = "Missing data";
                        rowIndex++;
                    }
                }

                columnIndex++;
                rowIndex = 2;
            }
        }

        static bool AllCellsInRowEqual(ExcelWorksheet worksheet, int rowIndex)
        {
            if (worksheet.Dimension == null)
            {
                return false;
            }
            int columnCount = worksheet.Dimension.End.Column;
            object firstCellValue = worksheet.Cells[rowIndex, 3].Value;

            for (int col=4; col <= columnCount; col++)
            {
                object currentCellValue = worksheet.Cells[rowIndex, col].Value;
                if (!object.Equals(currentCellValue, firstCellValue))
                {
                    return false;
                }
            }
            return true;
        }
    }
}
