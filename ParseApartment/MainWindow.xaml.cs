using System.Windows;
using System.IO;
using System.Net;
using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;

namespace ParseApartment
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        Excel.Application excelapp; // excel
        Excel.Workbooks excelappworkbooks; // рабочие книги
        Excel.Workbook excelappworkbook; // рабочая книга
        Excel.Sheets excelsheets;
        Excel.Worksheet excelworksheet;

        Excel.Range excelcells; // ЯЧЕЙКА ВЫДЕЛЕНАЯ СЕЙЧАС 

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            string uri1 = @"http://www.moyareklama.by/%D0%93%D0%BE%D0%BC%D0%B5%D0%BB%D1%8C/%D0%BA%D0%B2%D0%B0%D1%80%D1%82%D0%B8%D1%80%D1%8B_%D0%BF%D1%80%D0%BE%D0%B4%D0%B0%D0%B6%D0%B0/%D0%B2%D1%81%D0%B5/8/1/";
            string uri2 = @"http://www.moyareklama.by/%D0%93%D0%BE%D0%BC%D0%B5%D0%BB%D1%8C/%D0%BA%D0%B2%D0%B0%D1%80%D1%82%D0%B8%D1%80%D1%8B_%D0%BF%D1%80%D0%BE%D0%B4%D0%B0%D0%B6%D0%B0/%D0%B2%D1%81%D0%B5/8/2/";

            string[] uri = new string[] { uri1,uri2 };

            try
            {
                InsertInExcel(uri);
                if (MessageBox.Show("Открыть файл Excel?", "Внимание", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                {
                    Process.Start((Directory.GetCurrentDirectory() + "\\file.xlsb"));
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка");
            }
            
        }

        private string GetStreaam(string uri)
        {
            //получаем html код
            WebRequest request = WebRequest.Create(uri);
            WebResponse response = request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            StreamReader readeer = new StreamReader(dataStream);
            string stream = readeer.ReadToEnd();
            return stream;
        }

        private string[,] LoadData(string uri)
        {
            // достаем элементы
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(GetStreaam(uri));
            HtmlNodeCollection price = doc.DocumentNode.SelectNodes("//div[@class='price_realty']");
            HtmlNodeCollection adress = doc.DocumentNode.SelectNodes("//div[@class='title_realty']");
            // запихиваем в массив [стоимость,адрес]
            if (price != null)
            {
                string[,] data = new string[2, price.Count];
                for (int i = 0; i < price.Count; i++)
                {
                    data[0, i] = price[i].InnerText.Trim();
                    data[1, i] = adress[i].InnerText.Trim();
                }
                return data;
            }
            return null;
        }

        private void InsertInExcel(string[] uriMass)
        {
            excelapp = new Excel.Application();
            excelapp.Visible = false;
            excelapp.SheetsInNewWorkbook = 1; // задаем колисество листов в книге
            excelapp.Workbooks.Add(Type.Missing); // добовляем книгу
            //Запрашивать сохранение
            excelapp.DisplayAlerts = true;
            //Получаем набор ссылок на объекты Workbook (на созданные книги)
            excelappworkbooks = excelapp.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            excelappworkbook = excelappworkbooks[1];
            //Устанавливаем формат
            excelappworkbook.Saved = false; // узаываем что не сохранено ( для того что бы запросило сохранить при закрытии)

            //Получаем массив ссылок на листы выбранной книги
            excelsheets = excelappworkbook.Worksheets;
            //Получаем ссылку на лист 1
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            //название столбцев
            excelcells = (Excel.Range)excelworksheet.Cells[1, 1];
            excelcells.Value2 = "Стоимость ";
            excelcells = (Excel.Range)excelworksheet.Cells[1, 2];
            excelcells.Value2 = "Адресс ";

            //щетчик последней занятой ячейки
            int k = 2;

            //пробегаемся по всем ссылкам в массиве и добовляем из них данные в Excel
            for (int i = 0; i < uriMass.Length; i++)
            {
                string[,] data = LoadData(uriMass[i]);

                //Вывод в ячейки используя номер строки и столбца Cells[строка, столбец]
                for (int m = 0; m < data.Length / 2; m++)
                {
                    excelcells = (Excel.Range)excelworksheet.Cells[k, 1];
                    excelcells.Value2 = data[0, m];
                    excelcells = (Excel.Range)excelworksheet.Cells[k, 2];
                    excelcells.Value2 = data[1, m];
                    k++;
                }
            }
            // сохраняем данные
            excelappworkbook.SaveAs((Directory.GetCurrentDirectory() + "\\file.xlsb"), Excel.XlFileFormat.xlExcel12, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Excel.XlSaveAsAccessMode.xlShared, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing);
            excelapp.Quit();
        }
    }
}
