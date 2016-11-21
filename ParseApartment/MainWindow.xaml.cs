using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Web;
using System.Net;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace ParseApartment
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
       
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            string uri1 = @"http://www.moyareklama.by/%D0%93%D0%BE%D0%BC%D0%B5%D0%BB%D1%8C/%D0%BA%D0%B2%D0%B0%D1%80%D1%82%D0%B8%D1%80%D1%8B_%D0%BF%D1%80%D0%BE%D0%B4%D0%B0%D0%B6%D0%B0/%D0%B2%D1%81%D0%B5/8/1/";
            string uri2 = @"http://www.moyareklama.by/%D0%93%D0%BE%D0%BC%D0%B5%D0%BB%D1%8C/%D0%BA%D0%B2%D0%B0%D1%80%D1%82%D0%B8%D1%80%D1%8B_%D0%BF%D1%80%D0%BE%D0%B4%D0%B0%D0%B6%D0%B0/%D0%B2%D1%81%D0%B5/8/2/";

            WebRequest request = WebRequest.Create(uri1);
            WebResponse response = request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            StreamReader readeer = new StreamReader(dataStream);
            string html = readeer.ReadToEnd();

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
            HtmlNodeCollection c = doc.DocumentNode.SelectNodes("//div[@class='sa_header']");

            if (c != null)
            {
                
            }

                int startIndex= html.IndexOf("class=\"sa_header");
            html = html.Remove(0,startIndex);


            File.WriteAllText("file.txt",html);

        }
    }
}
