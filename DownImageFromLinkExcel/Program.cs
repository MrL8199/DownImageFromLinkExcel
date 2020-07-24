using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace DownImageFromLinkExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            List<DataModel> dataModels = LoadExcel(@"C:\Users\dinhv\source\repos\DownImageFromLinkExcel\DownImageFromLinkExcel\bin\Debug\file.xlsx");
            foreach(DataModel dataModel in dataModels)
            {
                string[] vs = dataModel.Link.Split('/');
                string pathSaveFile = vs[vs.Length-1].ToString();
                //dataModel.Link = pathSaveFile;
                Uri uri = new Uri(dataModel.Link);
                using (var client = new WebClient())
                {
                    client.DownloadFile(dataModel.Link, pathSaveFile);
                }
            }
            Console.WriteLine("Done!");
            Console.ReadKey();
        }

        public static List<DataModel> LoadExcel(string filePath)
        {
            List<DataModel> result;
            try
            {
                List<DataModel> list = new List<DataModel>();
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.FirstOrDefault<ExcelWorksheet>();
                    for (int i = 2; i <= excelWorksheet.Dimension.End.Row; i++)
                    {
                        bool flag = !string.IsNullOrEmpty(excelWorksheet.Cells[i, 1].Text);
                        if (flag)
                        {
                            DataModel item = new DataModel
                            {
                                Link = excelWorksheet.Cells[i, 1].Text,
                            };
                            list.Add(item);
                        }
                    }
                    result = list;
                }
            }
            catch
            {
                result = null;
            }
            return result;
        }
    }

    class DataModel
    {
        private string link;

        public DataModel()
        {
            
        }

        public string Link { get => link; set => link = value; }
    }
}
