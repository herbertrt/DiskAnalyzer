using System;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace DiskAanalyzer
{
    class DiskAanalyzer
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Hello Worlds");

            System.Diagnostics.Process.Start("CMD.exe", "/C dir c:\\ /a /s > output.txt").WaitForExit();
            ShowInExcel();

            Console.Write("Press any key to continue...");
            Console.ReadKey(true);
        }


        static void ShowInExcel()
        {

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            StreamReader reader = File.OpenText("output.txt");
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet1, xlWorkSheet2;
            object misValue = System.Reflection.Missing.Value;


            string pat1 = @" Directory of ([^\n]*)$";
            string pat2 = @"[^\n]* File\(s\)[ ]*([0-9,]*) bytes";
            string pat3 = @"([0-9/ :]*(A|P)M)[ ]*([0-9,]+)[ ]*([^\n]*)$";


            Regex r1 = new Regex(pat1);
            Regex r2 = new Regex(pat2);
            Regex r3 = new Regex(pat3);

            if (xlApp == null)
            {
                Console.Write("Excel is not properly installed!!");
                return;
            }


            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkBook.Worksheets.Add();

            xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            xlWorkSheet1.Cells[1, 1] = "Folder";
            xlWorkSheet1.Cells[1, 2] = "Size";

            xlWorkSheet2.Cells[1, 1] = "Folder";
            xlWorkSheet2.Cells[1, 2] = "FileName";
            xlWorkSheet2.Cells[1, 3] = "Size";
            xlWorkSheet2.Cells[1, 4] = "Date";

            xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            string line, cwd = "";
            int i = 2;
            int j = 2;
            while ((line = reader.ReadLine()) != null)
            {

                Match m3 = r3.Match(line);

                if (m3.Success)
                {

                    xlWorkSheet2.Cells[j, 1] = cwd;
                    xlWorkSheet2.Cells[j, 2] = m3.Groups[4].ToString();
                    xlWorkSheet2.Cells[j, 3] = m3.Groups[3].ToString();
                    xlWorkSheet2.Cells[j++, 4] = m3.Groups[1].ToString();


                }
                else
                {
                    Match m1 = r1.Match(line);

                    if (m1.Success)
                    {

                        cwd = m1.Groups[1].ToString();

                        xlWorkSheet1.Cells[i, 1] = cwd;


                    }

                    else
                    {
                        Match m2 = r2.Match(line);
                        if (m2.Success)
                        {

                            xlWorkSheet1.Cells[i++, 2] = m2.Groups[1].ToString();
                        }

                    }

                }



            }


            xlWorkBook.SaveAs("diskusage.xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet1);
            Marshal.ReleaseComObject(xlWorkSheet2);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }


    }




}
