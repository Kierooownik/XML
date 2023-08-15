using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace XML_Data_Reader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please enter the path to the XML file:");
            string xmlFilePath = Console.ReadLine();

            ProcessXmlFile(xmlFilePath);

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey(); // Czekaj na wciśnięcie dowolnego klawisza
        }

        static void ProcessXmlFile(string xmlFilePath)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlFilePath);

                XmlNodeList arrayOfParameterNodes = xmlDoc.SelectNodes("//chanelParameter/ArrayOfParameter");

                foreach (XmlNode arrayOfParameterNode in arrayOfParameterNodes)
                {
                    XmlAttribute nameAttribute = arrayOfParameterNode.Attributes["name"];
                    if(nameAttribute !=null)
                    {
                        string arrayOfParameterName = nameAttribute.Value;
                        Console.WriteLine();
                        Console.WriteLine(arrayOfParameterName);
                    }
                    

                    XmlNodeList parameterNodes = arrayOfParameterNode.SelectNodes("Parameter");

                    foreach (XmlNode parameterNode in parameterNodes)
                    {
                        string name = parameterNode.SelectSingleNode("name").InnerText;
                        string value = parameterNode.SelectSingleNode("value").InnerText;
                        string index = parameterNode.SelectSingleNode("index").InnerText;

                        Console.WriteLine($"{index} {name} = {value}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
