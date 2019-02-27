using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WADOU.COMMON;
using System.Speech.Synthesis;
using Microsoft.Office.Core;
using Excel=Microsoft.Office.Interop.Excel;
using System.Reflection;
using ExcelReport;

namespace ConsoleApplication1
{
    public class Program
    {
        static void Main(string[] args)
        {
            /*
            string path = @"c:\excel\MAINMODEL.xlsx";
            string saveas = @"c:\excel\newexcel.xlsx";
            if (File.Exists(path))
            {
                InsertPicToExcel ipt = new InsertPicToExcel();
                ipt.Open(path);
                //ipt.insertdata(11,1,"88");
                //ipt.InsertPicture("C11", @"c:\excel\1.jpg", 85, 80);
                for(int i =200;i>10;i--)
                {
                    ipt.deleterow(i);
                }
                ipt.SaveFile(saveas);
                ipt.Dispose();
            }
            */
            DateTime a = Convert.ToDateTime("2019-01-06 07:35:16.000");
            DateTime b = Convert.ToDateTime("2019-01-05 11:38:40.000");
            if(b<=a)
            {
                int xx = 0;
            }
            else
            {
                string sql = string.Format("/*dialect*/UPDATE atdrecord SET Mark='JB' from atdrecord where BADGENUMBER = '180077' and DATEDIFF(DAY, '{0}', CHECKTIME) >= 0 AND DATEDIFF(DAY, '{0}', CHECKTIME) <= DATEDIFF(DAY, '{0}', '{1}')",a.ToString("yyyy-MM-dd"),b.ToString("yyyy-MM-dd"));

            }
        }
    }
}

