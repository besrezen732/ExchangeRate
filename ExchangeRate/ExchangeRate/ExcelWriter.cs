using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExchangeRate;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExchangeRate
{
    class ExcelColumnList 
    {
        public string Date {get; set; }
        public string NumCode { get; set; }
        public string StrCode { get; set; }
        public string Count { get; set; }
        public string Name { get; set; }
        public string Course { get; set; }
    }

    class WriteEсxelClass
    {
        public List<ExcelColumnList> ExcelColumnsWrite(string html, string dateString)
        {
            var htmlList = html.Replace("<tr>","").Trim().Split("</tr>").ToList();
            List<ExcelColumnList> excelColumnList = new List<ExcelColumnList>();

            foreach (string hl in htmlList)
            {
                string[] separatingStrings = { "</th>", "</td>" };

                var hlMass= hl.Replace("<th>", "").Replace("<td>", "").Trim().Split(separatingStrings, System.StringSplitOptions.RemoveEmptyEntries);
                excelColumnList.Add(new ExcelColumnList()
                {
                    Date = dateString,
                    NumCode = hlMass[0],
                    StrCode = hlMass[1],
                    Count = hlMass[2],
                    Name = hlMass[3],
                    Course = hlMass[4]
                });
            }
           
            return excelColumnList;
        }

        public void ExportToExcel()
        {
            string fileName = string.Format(@"{0}\LogFile\ExchangeRate.xlsx", Environment.CurrentDirectory);
            //Объявляем приложение
            Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            // Сделать приложение Excel видимым
            excelApp.Visible = true;
            var xlWB = excelApp.Workbooks.Add();
            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet workSheet = (Excel.Worksheet) excelApp.Worksheets.get_Item(1);


            // Установить заголовки столбцов в ячейках
            workSheet.Cells[1, "A"] = "Date";
            workSheet.Cells[1, "B"] = "NumCode";
            workSheet.Cells[1, "C"] = "StrCode";
            workSheet.Cells[1, "D"] = "Count";
            workSheet.Cells[1, "E"] = "Name";
            workSheet.Cells[1, "F"] = "Course";


            if (System.IO.File.Exists(fileName)) System.IO.File.Delete(fileName);
            xlWB.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookDefault); //формат Excel 2007
            xlWB.Close(false); //false - закрыть рабочую книгу не сохраняя изменения
            excelApp.Quit();
        }
    }

    #region График в excel

    /*
    class BasePrintExcelClass
    {
    public void WriteExcel(string excelString1)
    {
        Console.WriteLine(excelString1);
       
        // Экспорт данных в Excel
            int n = 50;
            double[] Y = new double[n];
            Console.WriteLine("Для экспорта данных в Excel нажмите любую клавишу!");
            //Console.ReadKey();
            for (int i = 0; i < n; i++)
                Y[i] = 1 - Math.Exp(-i / 10.0);
            OutInExcel(Y, n); // Вызов метода
            Console.WriteLine("Вы сохранили данные в *.xlsx файле?");            
        }

      
        // Сохранение данных с построением графика в MS Excel
        static void OutInExcel(double[] Y, int n)
        { // Создаём объект - экземпляр нашего приложения
            Excel.Application excelApp = new Excel.Application();
            // Создаём экземпляр рабочей книги Excel
            Excel.Workbook workBook;
            // Создаём экземпляр листа Excel
            Excel.Worksheet workSheet;
            workBook = excelApp.Workbooks.Add();
            workSheet = (Excel.Worksheet) workBook.Worksheets.get_Item(1);
            // Заполняем первый столбец листа из массива Y[0..n-1]
            for (int j = 1; j <= n; j++)
                workSheet.Cells[j, 1] = Y[j - 1];
            // Вывод текста
            Excel.Range rng = workSheet.Range["c1"];
            rng.Formula = "Результат";
            // Настраиваем линейный график через интерфейсы ChartObjects, ChartObject, Chart, Range
            Excel.ChartObjects chartObjs = (Excel.ChartObjects) workSheet.ChartObjects();
            Excel.ChartObject chartObj = chartObjs.Add(70, 30, 300, 200);
            Excel.Chart xlChart = chartObj.Chart;
            string sRange = "A1:A" + n.ToString(); // Вывод в 1 столбец - n чисел
            Excel.Range rng2 = workSheet.Range[sRange]; //
            // Устанавливаем тип диаграммы
            xlChart.ChartType = Excel.XlChartType.xlLineStacked; // тип графика - линейный
            // Устанавливаем источник данных (значения от 1 до n)
            xlChart.SetSourceData(rng2);
            // Открываем созданный excel-файл
            excelApp.Visible = true;
            excelApp.UserControl = true;
            workSheet.SaveAs(string.Format(@"{0}\Price.xlsx", Environment.CurrentDirectory));
            excelApp.Quit();
        }
      

    }
          */

    #endregion

}




