using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel;

namespace XLSX2RESW.Classes
{
    public class Elaborator
    {
        private static List<string> errors = new List<string>();

        public static void Convert(string[] files)
        {
            //iterate all files
            foreach(var file in files)
            {
                try
                {
                    //check if item is a file
                    if(!File.GetAttributes(file).HasFlag(FileAttributes.Directory))
                    {
                        //check if file is supported
                        if(Path.GetExtension(file) == Constants.InputExtension)
                        {
                            //get folder path
                            string FolderPath = Path.GetDirectoryName(file);

                            //read file into JLanguage
                            List<JLanguage> myoutput = ReadFile(file);

                            //elaborate file
                            string outputFolder = FolderPath + "\\" + Path.GetFileNameWithoutExtension(file) + Constants.OutputFileNameSuffix;

                            if(!Directory.Exists(outputFolder))
                            {
                                Directory.CreateDirectory(outputFolder);

                                foreach(var items in myoutput)
                                {
                                    Directory.CreateDirectory(outputFolder + "\\" + items.code);

                                    File.AppendAllText(outputFolder + "\\" + items.code + "\\" + Constants.OutputFileName, 
                                        Properties.Resources.resw_start);

                                    foreach(var item in items.values)
                                    {
                                        File.AppendAllText(outputFolder + "\\" + items.code + "\\" + Constants.OutputFileName, 
                                            "\n" + String.Format(Properties.Resources.resw_value, item.id, item.value));
                                    }

                                    File.AppendAllText(outputFolder + "\\" + items.code + "\\" + Constants.OutputFileName, 
                                        "\n" + Properties.Resources.resw_end);
                                }

                                //DONE!
                            }
                            else
                            {
                                errors.Add(Path.GetFileName(file) + " not converted. " + Path.GetFileName(outputFolder) + " folder already exist. Please delete it and retry.");
                            }
                        }
                            else
                        {
                            errors.Add(Path.GetFileName(file) + " not converted. This program supports only " + Constants.InputExtension + " files.");
                        }
                    }
                    else
                    {
                        errors.Add(Path.GetFileName(file) + " not converted. This program supports only files.");
                    }
                }
                catch(Exception ex)
                {
                    switch(ex.HResult)
                    {
                        case -2147024864:
                            errors.Add(Path.GetFileName(file) + " not converted. Please close your " + Constants.InputExtension + " file from excel first!");
                            break;

                        default:
                            errors.Add(Path.GetFileName(file) + " not converted. "+ ex.Message);
                            Debug.WriteLine(ex.Message + "\n\n" + ex.StackTrace, "Exception " + ex.HResult);
                            break;
                    }
                }
            }

            //check any errors
            if(errors.Count > 0)
            {
                StringBuilder message = new StringBuilder();
                message.AppendLine("The following items were not converted:");
                message.AppendLine();
                foreach(var error in errors)
                {
                    message.AppendLine("- " + error);
                }

                //print all errors
                MessageBox.Show(message.ToString(), "Warning");
            }
        }

        private static List<JLanguage> ReadFile(string file)
        {
            //open file
            FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            //get total languages
            int totalLanguages = GetTotalLanguages(result);

            //get total rows
            int totalRows = GetTotalRows(result);

            //output to return
            List<JLanguage> projects = new List<JLanguage>();

            //iterate languages
            for(int i = 1; i <= totalLanguages; i++)
            {
                string project_code = "";
                List<JValues> project_values = new List<JValues>();


                project_code = result.Tables[0].Rows[1][i].ToString();
                for(int j = 2; j < totalRows; j++)
                {
                    project_values.Add(new JValues()
                    {
                        id = FixID(result.Tables[0].Rows[j][0].ToString()),
                        value = FixValue(result.Tables[0].Rows[j][i].ToString())
                    });
                }

                projects.Add(new JLanguage() { code = project_code, values = project_values });
            }

            excelReader.Close();

            return projects;
        }

        private static int GetTotalRows(DataSet result)
        {
            int rows = 0;

            foreach(DataRow item in result.Tables[0].Rows)
            {
                var cell = item[0].ToString();

                if(cell != null && cell != "")
                {
                    rows++;
                }
            }

            return rows;
        }

        private static int GetTotalLanguages(DataSet result)
        {
            int languages = 0;

            foreach(DataColumn item in result.Tables[0].Columns)
            {
                var cell = result.Tables[0].Rows[1][item].ToString();

                if(cell != null && cell != "")
                {
                    languages++;
                }
            }

            return languages - 1;
        }

        private static string FixID(string id)
        {
            return id.Replace("\n", "").Replace("\"", "").Replace("&", "&amp;").Replace(" ", "_");
        }

        private static string FixValue(string value)
        {
            return value.Replace("&", "&amp;");
        }
    }
}
