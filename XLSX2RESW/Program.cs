using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel;

namespace XLSX2RESW
{
    static class Program
    {
        private static string filePath = null;
        private static string supportedExtension = "xlsx";
        private static string outputFileName = "Resources.resw";
        private static string suffixOutputFileName = "_XLSX2RESW";

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {

                //check args
                if(args.Length == 1)
                {
                    filePath = args[0];

                    //check if item dropped is a file
                    if(!File.GetAttributes(filePath).HasFlag(FileAttributes.Directory))
                    {
                        //check if file is a .csv
                        if(Path.GetExtension(filePath) == "." + supportedExtension)
                        {
                            string folderPath = Path.GetDirectoryName(filePath);

                            List<JProject> myoutput = ReadFile();

                            //elaborate file
                            string outputFolder = folderPath + "\\" + Path.GetFileNameWithoutExtension(filePath) + suffixOutputFileName;

                            if(!Directory.Exists(outputFolder))
                            {
                                Directory.CreateDirectory(outputFolder);

                                foreach(var items in myoutput)
                                {
                                    Directory.CreateDirectory(outputFolder + "\\" + items.code);

                                    File.AppendAllText(outputFolder + "\\" + items.code + "\\" + outputFileName, Properties.Resources.resw_start);

                                    foreach(var item in items.values)
                                    {
                                        File.AppendAllText(outputFolder + "\\" + items.code + "\\" + outputFileName, "\n" + String.Format(Properties.Resources.resw_value, item.id, item.value));
                                    }

                                    File.AppendAllText(outputFolder + "\\" + items.code + "\\" + outputFileName, "\n" + Properties.Resources.resw_end);
                                }

                                //DONE!
                            }
                            else
                            {
                                MessageBox.Show("The " + Path.GetFileNameWithoutExtension(filePath) + suffixOutputFileName + " folder already exist.\nPlease delete it and retry.", "Error");
                            }
                        }
                        else
                        {
                            MessageBox.Show("This application only supports drag & drop of 1 " + supportedExtension + " file!", "Error");
                        }
                    }
                    else
                    {
                        MessageBox.Show("This application only supports drag & drop of 1 " + supportedExtension + " file!", "Error");
                    }
                }
                else
                {
                    MessageBox.Show("This application only supports drag & drop of 1 " + supportedExtension + " file!", "Error");
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace, "Exception");
            }

            //exit
            Application.Exit();
        }

        private static List<JProject> ReadFile()
        {
            //open xlsx file
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            //get total languages
            int totalLanguages = GetTotalLanguages(result);

            //get total rows
            int totalRows = GetTotalRows(result);

            //output to return
            List<JProject> projects = new List<JProject>();

            //iterate languages
            for(int i = 1; i <= totalLanguages; i++)
            {
                string project_code = "";
                List<JValues> project_values = new List<JValues>();


                project_code = result.Tables[0].Rows[1][i].ToString();
                for(int j = 2; j <= totalRows; j++)
                {
                    project_values.Add(new JValues()
                    {
                        id = FixID(result.Tables[0].Rows[j][0].ToString()),
                        value = FixValue(result.Tables[0].Rows[j][i].ToString())
                    });
                }

                projects.Add(new JProject() { code = project_code, values = project_values });
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
