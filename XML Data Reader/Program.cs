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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Console.WriteLine("Proszę podać ścieżkę do pliku xml urządzenia referencyjnego");
            string referencePath = Console.ReadLine();

            Console.WriteLine("Proszę podać ścieżkę do folderu z plikami xml do sprawdzenia");
            string folderPath = Console.ReadLine();
            string[] xmlFiles = Directory.GetFiles(folderPath, "*.xml");

            string outputPath = Path.Combine(folderPath, "output.xlsx");
            
            Console.WriteLine($"Aby otrzymać pełne dane wpisz 'full', aby otrzymać tylko parametry z różniącymi się wartościami wpisz 'diff'");
            string displayType = Console.ReadLine();

            Dictionary<int, ParameterInfo> referenceParameters = ProcessXmlFile(referencePath);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Dane");

                int rowReferenceIndex = 2;
                worksheet.Cells[1, 3].Value = referenceParameters[31318].Value;
                foreach (var parameter in referenceParameters)
                {
                    worksheet.Cells[rowReferenceIndex, 1].Value = parameter.Key;
                    worksheet.Cells[rowReferenceIndex, 2].Value = parameter.Value.Name;
                    worksheet.Cells[rowReferenceIndex, 3].Value = parameter.Value.Value;
                    rowReferenceIndex++;
                }


                if (displayType == "full")
                {
                    ProcessXmlToWorksheet(worksheet, xmlFiles, referencePath, referenceParameters);
                    
                }
                else if (displayType == "diff")
                {
                    ProcessXmlToWorksheet(worksheet, xmlFiles, referencePath, referenceParameters);

                    //Usuwanie wierszy z takimi samymi ustawieniami we wszystkich plikach
                    for (int row=2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        if (AllCellsInRowEqual(worksheet, row))
                        {
                            worksheet.Cells[row, 1].Value = "All Equal";
                            //worksheet.DeleteRow(row);
                            Console.WriteLine("Test");
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
            Console.WriteLine("Export to .xls finished");
            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
            System.Diagnostics.Process.Start(outputPath);


            
        }
        class ParameterInfo
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }


        static Dictionary<int, ParameterInfo> ProcessXmlFile(string xmlFilePath)
        {
            Dictionary<int, ParameterInfo> parameters = new Dictionary<int, ParameterInfo>();

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlFilePath);

                XmlNodeList parameterNodes = xmlDoc.SelectNodes("//chanelParameter/ArrayOfParameter/Parameter");

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
        static void ProcessXmlToWorksheet(ExcelWorksheet worksheet, string[] xmlFiles, string referencePath, Dictionary<int,ParameterInfo> referenceParameters)
        {
            int columnIndex = 4;
            int rowIndex = 2;
            
            foreach (string xmlFilePath in xmlFiles.Where(filePath => filePath != referencePath))
            {
                Dictionary<int, ParameterInfo> parameters = ProcessXmlFile(xmlFilePath);

                worksheet.Cells[1, columnIndex].Value = parameters[31318].Value;
                foreach (var refKey in referenceParameters)
                {
                    if (parameters.ContainsKey(refKey.Key))
                    {
                        ParameterInfo parameter = parameters[refKey.Key];
                        worksheet.Cells[rowIndex, columnIndex].Value = parameter.Value;
                        rowIndex++;
                    }
                    else
                    {
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

        static void CompareParameterValues(Dictionary<int, ParameterInfo> parameters1, Dictionary<int, ParameterInfo> parameters2)
        {
            Console.WriteLine("Comparing parameter values between the two XML files:");

            foreach (var kvp1 in parameters1)
            {
                if (parameters2.ContainsKey(kvp1.Key))
                {
                    ParameterInfo parameter2 = parameters2[kvp1.Key];
                    
                    if (kvp1.Value.Value == parameter2.Value)
                    {
                        Console.WriteLine($"Parameter with index {kvp1.Key} ({kvp1.Value.Name}) has the same value in both files: {kvp1.Value}");
                    }
                    else
                    {
                        Console.WriteLine($"Parameter with index {kvp1.Key}({kvp1.Value.Name}) has different values:");
                        Console.WriteLine($"   Value in the first file: {kvp1.Value.Value}");
                        Console.WriteLine($"   Value in the second file: {parameter2.Value}");
                    }
                }
                else
                {
                    Console.WriteLine($"Parameter with index {kvp1.Key} exists in the first file, but not in the second file.");
                }
            }

            foreach (var kvp2 in parameters2)
            {
                if (!parameters1.ContainsKey(kvp2.Key))
                {
                    Console.WriteLine($"Parameter with index {kvp2.Key} exists in the second file, but not in the first file.");
                }
            }
        }

       
    }
}
