using Newtonsoft.Json.Linq;
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
                                foreach (var filename in LoadXml())
                                {
                                    XElement data = XElement.Load(filename);
                                    foreach (var node in config.Nodes)
                                    {
                                        ParseNode(node, data);
                                    }
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

        static void ParseNode(dynamic node, XElement data)
        {
            //if title has a name
            bool isNumeric = (((string)node.title).All(c => c >= '0' && c <= '9'));            
            
            if (node.Childrens is null)
            {                
                Console.WriteLine(data.Elements((string)node.NodeName).Where(x => (string)x.Attribute("title") == (string)node.title).FirstOrDefault().Attribute("value").Value);
                Excel.Application xlApp = new
                Microsoft.Office.Interop.Excel.Application();
            }
            else
            {
                foreach (var node_ in node.Childrens)
                {
                    ParseNode(node_, isNumeric?
                        data.Elements((string)node.NodeName).ToList()[int.Parse((string)node.title)]
                        : data.Elements((string)node.NodeName).Where(x => (string)x.Attribute("title") == (string)node.title).FirstOrDefault()
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
    }
}
