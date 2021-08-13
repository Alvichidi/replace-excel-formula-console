using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace ExcelExternalLinkConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length != 3)
            {
                return;
            }

            string filePaths = args[0];
            string oldStrings = args[1];
            string replaceStrings = args[2];

            string[] filePathssArray = filePaths.Split(';');
            string[] oldStringsArray = oldStrings.Split(';');
            string[] replaceStringsArray = replaceStrings.Split(';');

            if (filePathssArray.Length == 0 || oldStringsArray.Length == 0 || replaceStringsArray.Length == 0)
            {
                return;
            }

            Application oExcel = null;
            Workbook oBook = null;

            try
            {
                oExcel = new Application();

                foreach (var file in filePathssArray)
                {
                    oBook = OpenBook(oExcel, file);
                    UpdateLinks(oBook, oldStringsArray, replaceStringsArray);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oBook != null)
                {
                    oBook.Close(false, Type.Missing, Type.Missing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);

                }
                oBook = null;

                if (oExcel != null)
                {
                    oExcel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
                }
                oExcel = null;
                Console.WriteLine("Done!");
            }
            return;
        }

        private static void UpdateLinks(Workbook oBook, string[] oldStringsArray, string[] replaceStringsArray)
        {
            int numberChange = oldStringsArray.Length;

            int numberSheets = oBook.Worksheets.Count;
            for (int i = 1; i <= numberSheets; i++)
            {
                Worksheet worksheet = (Worksheet)oBook.Worksheets[i];

                int numberRows = worksheet.UsedRange.Rows.Count;
                int numberCells = worksheet.UsedRange.Cells.Count;

                for (int y = 1; y <= numberRows; y++)
                {
                    for (int x = 1; x <= numberCells; x++)
                    {
                        Range tagetCell = (Range)worksheet.Cells[y, x];
                        if (tagetCell != null)
                        {
                            string formule = tagetCell.Formula;
                            for (int c = 0; c < numberChange; c++)
                            {
                                formule = formule.Replace(oldStringsArray[c], replaceStringsArray[c]);
                                int tempIndex = formule.IndexOf("]");
                                if (tempIndex > 0)
                                {
                                    string targetFilePath = formule.Substring(2, tempIndex - 2);
                                    targetFilePath = targetFilePath.Replace("[", "");
                                    bool isExist = File.Exists(targetFilePath);
                                    if (isExist)
                                    {
                                        tagetCell.Formula = formule;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            oBook.Save();
        }

        public static Workbook OpenBook(Application excelInstance, string filepath)
        {
            Workbook book = excelInstance.Workbooks.Open(
                filepath, 0, false, 5, "", "", false, XlPlatform.xlWindows, "",
                true, false, 0, true, false, false
            );
            return book;
        }
    }
}
