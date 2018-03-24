using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace DownloadFromNeT
{
   class Program
   {
      static void Main(string[] args)
      {
         // программа скачивает с сайта НБ РБ excel файл с курсами валют
         // выбирает последнюю актуальную дату 
         // выводит на экран курс доллара США на эту дату
         
         Console.WriteLine("Saving... a file with exchange rates");
         WebClient wc = new WebClient();
         wc.DownloadFile("http://www.nbrb.by/statistics/rates/ratesDaily.asp?yr=2018", @"exchange_of_currency.xls");
         Console.WriteLine("Download complete!");
         Console.ReadLine();
         

         // поиск курсов
         string fileName = @"D:\itproject\visual studio\DownloadFromNeT\DownloadFromNeT\bin\Debug\exchange_of_currency.xls";

         Excel.Workbook excelBook;
         Excel.Worksheet excelsheet;

         Excel.Application excelApp = new Excel.Application();

         excelBook = excelApp.Workbooks.Open(fileName); // открываем книгу
         excelsheet = excelBook.Worksheets["year"];

         Console.WriteLine("U.S. dollar (USD)");
         Console.WriteLine("last date ");
         
         int lastRow = excelsheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
         string lastIndx = null;
         for (int i = lastRow; i >= 1; i--)
         {
            if (excelsheet.Cells[i, 1].Value != null)
            {
               lastIndx = i.ToString();
               break;
            }
         }


         Console.WriteLine(excelsheet.Range["A" + lastIndx].Value.ToShortDateString().ToString());
         Console.WriteLine(excelsheet.Range["AY" + lastIndx].Value.ToString() + " BYN");

      }
   }
}
;