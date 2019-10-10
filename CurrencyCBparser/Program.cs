
using System.Xml;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Linq;

namespace CurrencyCBparser
{
    class Program
    {
        const string url = "http://www.cbr.ru/scripts/XML_daily.asp?date_req=21.08.2019";
        static void Main(string[] args)
        {
            //Создаем Excel-файл
            Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook wb = oXL.Workbooks.Add(Excel.XlSheetType.xlWorksheet);
            Excel.Worksheet ws = (Excel.Worksheet)oXL.ActiveSheet;
            //Добавляем заголовки 
            ws.Cells[1, 1] = "Валюта";
            ws.Cells[1, 2] = "Курс";

            //пишем данные со 2ой строки
            int row = 2;

            //загружаем xml-данные 
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(url);
            XmlElement xRoot = xDoc.DocumentElement;
          
            // обход всех узлов в корневом элементе Valute
            foreach (XmlNode xnode in xRoot)
            {

                // обходим все дочерние узлы элемента Valute 
                foreach (XmlNode childnode in xnode.ChildNodes)
                {
                    // если узел - название валюты,записываем в файл
                    if (childnode.Name == "Name")
                    {
                        ws.Cells[row, 1] = childnode.InnerText;
                    }
                    // если узел - курс валюты,записываем в файл
                    if (childnode.Name == "Value")
                    {
                        ws.Cells[row, 2].NumberFormat = "@";
                        ws.Cells[row, 2] = childnode.InnerText.ToString();

                    }
                    
                }
                row++;
            }

            oXL.Visible = true;
        }
    }
}