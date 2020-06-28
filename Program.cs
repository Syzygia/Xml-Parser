using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace Xml_Parser
{
    class Program
    {
        static dynamic config;
        static ExcelPackage excel = new ExcelPackage();
        static int nodesParsed = 0;
        [STAThread]
        static void Main(string[] args)
        {
            LoadConfig(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "config.json");
           
            string reply = "";
            while (reply != "n")
            {
                ShowDesciption();
                reply = Console.ReadLine();
                switch (reply)
                {
                    case "!config":
                        {
                            using (OpenFileDialog openFileDialog = new OpenFileDialog())
                            {
                                openFileDialog.Title = "Choose config file";
                                openFileDialog.InitialDirectory = "c:\\";
                                openFileDialog.Filter = "txt files (*.txt)|*.txt| json files (*.json)|*.json";
                                openFileDialog.FilterIndex = 1;
                                openFileDialog.RestoreDirectory = true;
                                openFileDialog.CheckFileExists = true;
                                openFileDialog.CheckPathExists = true;

                                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                {
                                    LoadConfig(openFileDialog.FileName);
                                }

                            }
                            break;
                        }
                    case "n":
                        {
                            return;
                        }
                    case "y":
                        {
                            try
                            {
                                LoadHeaders();
                                //to write in correct row
                                int filesParsed = 2;                                
                                foreach (var filename in LoadXml())
                                {
                                    excel.Workbook.Worksheets.First().Cells[filesParsed, 1].Value =
                                        Path.GetFileNameWithoutExtension(filename);
                                    nodesParsed = 2;
                                    XElement data = XElement.Load(filename);
                                    foreach (var node in config.Nodes)
                                    {                     
                                        WasBond = false;
                                        ParseNode(node, data, filesParsed);
                                    }
                                    ++filesParsed;
                                }
                                SaveExcel();
                                Console.WriteLine("Your data have been succesfully parsed!" 
                                    + System.Environment.NewLine +
                                    "Do you want to parse anything else? (y/n)");
                                if (Console.ReadLine() == "y")
                                {

                                }
                                else
                                {
                                    return;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                            break;
                        }
                }
            }
        }

        static bool WasBond = false;
        static void ParseNode(dynamic node, XElement data, int filesParsed)
        {
            //if title has a name
            bool isNumber = isNumeric((string)node.title);            
            
            if (node.Childrens is null)
            {                
                /*Console.WriteLine(data.Elements((string)node.NodeName)
                    .Where(x => (string)x.Attribute("title") ==
                    (string)node.title).FirstOrDefault().Attribute("value").Value);*/
                var parsedData = isNumber ?
                    data.Elements((string)node.NodeName).ToList()
                    [int.Parse((string)node.title)].Attribute("value").Value
                    : data.Elements((string)node.NodeName)
                    .Where(x => (string)x.Attribute("title")
                    == (string)node.title).FirstOrDefault().Attribute("value").Value;
                if (node.bond is null)
                {
                    nodesParsed += WasBond ? 1 : 0;
                    var excelData = excel.Workbook.Worksheets["Worksheet1"]
                        .Cells[filesParsed, nodesParsed++];                
                    excelData.Value = parsedData;
                }
               else
                {
                    var excelData = excel.Workbook.Worksheets["Worksheet1"]
                        .Cells[filesParsed, nodesParsed];
                    excelData.Value += " " + parsedData;
                    WasBond = true;
                }
                    
            }
            else
            {
                foreach (var node_ in node.Childrens)
                {
                    ParseNode(node_, isNumber?
                        data.Elements((string)node.NodeName).ToList()[int.Parse((string)node.title)]
                        : data.Elements((string)node.NodeName).Where(x => (string)x.Attribute("title")
                        == (string)node.title).FirstOrDefault(),
                        filesParsed
                        );
                }
            }
            
        }

        static void ShowDesciption ()
        {
            Console.WriteLine("This programm extract data from custom fileds of xml Specca files and save it in Excel file" + System.Environment.NewLine +
                "Fields can be customized by txt/json file, according to the rules of json format" + System.Environment.NewLine +
                "By default programm tries to use config.txt file in the same directory as this exe file" + System.Environment.NewLine +
                "If you want to choose config file - type: !config" + System.Environment.NewLine +
                "If config file is set and you wish to  continue type: y" + System.Environment.NewLine +
                "If you wish to exit programm type: n");
        }
        static IEnumerable<string> LoadXml ()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Choose xml files to parse";
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "xml files (*.xml)|*.xml";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = true;
                openFileDialog.CheckFileExists = true;
                openFileDialog.CheckPathExists = true;

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return openFileDialog.FileNames;
                }
                else
                {
                    Console.WriteLine("Something goes wrong when loading xml files, try again");
                    return null;
                }
            }
        }

        static void LoadConfig (string path)
        {
            try
            {
                config = JObject.Parse(File.ReadAllText(path));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
        }

        static void LoadHeaders ()
        {            
            ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("Worksheet1");
            int i = 0;
            foreach (var header in config.Headers)
            {
                sheet.Cells[1, ++i].Value = (string)header;
            }            
        }

        static void SaveExcel ()
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Title = "Choose folder to save Excel file";
                saveFileDialog.InitialDirectory = "c:\\";
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.Filter = "xls files (*.xls)| *.xls | xlsx files (*.xlsx)| *.xlsx";
                saveFileDialog.FilterIndex = 2;
                saveFileDialog.RestoreDirectory = false;                
                saveFileDialog.CheckPathExists = true;

                if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    excel.SaveAs(new FileInfo(saveFileDialog.FileName));
                }
                else
                {
                    Console.WriteLine("Something goes wrong when trying to save excel file");                   
                }

            }
        }

        static bool isNumeric(string data)
        {
            return data.All(c => c >= '0' && c <= '9');
        }
    }
}
