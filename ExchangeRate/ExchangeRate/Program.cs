using System;
using System.Collections.Generic;

namespace ExchangeRate
{
    class Program
    {
        static void Main(string[] args)
        {

            var startDate = new DateTime(2021, 1, 1);
            var endDate = new DateTime(2021, 1, 3);

            List<List<ExcelColumnList>>excelColumnAll = new List<List<ExcelColumnList>>();
            WriteExcelClass writeExcelClass = new WriteExcelClass();
            HtmlBaseClass htmlBaseClass = new HtmlBaseClass();
            while (startDate <= endDate)
            {
                var dateString  = startDate.ToString("dd.MM.yyyy");
     

                
                string htmlString = htmlBaseClass.GetHtml("https://cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To=" + dateString);
                htmlString = htmlBaseClass.ClearSting(htmlString);

                excelColumnAll.Add(writeExcelClass.ExcelColumnsWrite(htmlString, dateString));
                startDate = startDate.AddDays(1);
                Console.WriteLine(dateString);
            }
            writeExcelClass.ExportToExcel(excelColumnAll);
            
        }
    }
}
