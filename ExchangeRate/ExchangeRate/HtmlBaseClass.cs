using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace ExchangeRate
{
    class HtmlBaseClass
    {
        public string GetHtml(string urlAddress)
        {
            string getHtmlString = String.Empty;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlAddress);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                Stream receiveStream = response.GetResponseStream();
                StreamReader readStream = null;

                if (response.CharacterSet == null)
                {
                    readStream = new StreamReader(receiveStream);
                }
                else
                {
                    readStream = new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));
                }
                getHtmlString = readStream.ReadToEnd();
                response.Close();
                readStream.Close();              
            }
            return getHtmlString;
        }
        public void WriteFile( string html, string fileName)
        {
            // создаем каталог для файла
            string path = $"{Directory.GetCurrentDirectory()}\\LogFile\\";
            DirectoryInfo dirInfo = new DirectoryInfo(path);
            if (!dirInfo.Exists)
            {
                dirInfo.Create();
            }
            var filePath = $"{path}\\{fileName}.txt";
            FileInfo fileInf = new FileInfo(filePath);
            if (fileInf.Exists)
            {
                fileInf.Delete();
            }
            using (FileStream fstream = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                // преобразуем строку в байты
                byte[] array = System.Text.Encoding.Default.GetBytes(html);
                // запись массива байтов в файл
                fstream.Write(array, 0, array.Length);
                Console.WriteLine("Текст записан в файл");
            }
        }

        public string ClearSting(string html)
        {
            var startPosition = html.IndexOf("<table class=\"data\">", StringComparison.Ordinal);
            var endPosition = html.IndexOf("</table>", StringComparison.Ordinal);
            html = html.Substring(startPosition, endPosition - startPosition);

            html = (html.Substring(html.IndexOf("<tr>", StringComparison.Ordinal),
                html.LastIndexOf("</tr>", StringComparison.Ordinal) - html.IndexOf("<tr>", StringComparison.Ordinal)).Replace("  ","") + "</tr>").Trim();

            return html;
        }

        

    }

    
}

