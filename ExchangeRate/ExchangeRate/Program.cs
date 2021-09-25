using System;
using System.Collections.Generic;

namespace ExchangeRate
{
    class Program
    {
        static void Main(string[] args)
        {

            var startDate = new DateTime(2021, 1, 1);
            var endDate = new DateTime(2021, 1, 1);

            List<List<ExcelColumnList>>excelColumnAll = new List<List<ExcelColumnList>>();
            while (startDate <= endDate)
            {
                var dateString  = startDate.ToString("dd.MM.yyyy");
                WriteEсxelClass writeExelClass = new WriteEсxelClass();
                HtmlBaseClass htmlBaseClass = new HtmlBaseClass();

                
                string htmlString = htmlBaseClass.GetHtml("https://cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To=" + dateString);
                htmlString = htmlBaseClass.ClearSting(htmlString);

                excelColumnAll.Add(writeExelClass.ExcelColumnsWrite(htmlString, dateString));
                startDate = startDate.AddDays(1);
                Console.WriteLine(dateString);
            }
            //writeExelClass.ExportToExcel();
            //html_base_class.WriteFile(html, dateString);
        }
    }
}
