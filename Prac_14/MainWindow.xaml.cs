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
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using DocumentFormat.OpenXml.Packaging;

namespace Prac_14
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public float accepted;

        public MainWindow()
        {
            InitializeComponent();
        }

        //Загрузка формы
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            sum_buy.IsEnabled = false;
            change.IsEnabled = false;

            List<string> list = Sub.Full_Combobox();

            foreach (string asd in list)
            {
                stationery.Items.Add(asd);
            }
        }

        //Изменение текста в текстбокс "количество"
        private void count_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(stationery.Text.Length > 0)
            {
                int price = Sub.KnowPrice(stationery.SelectedItem.ToString());
                int thing;
                int.TryParse(count.Text, out thing);
                if (count.Text.Length > 0 && (count.Text.Length > 0) && ((thing != 0) && (thing > 0)))
                    sum_buy.Text = (price * thing).ToString();
            }
        }

        //Событие выбора в комбобокс
        private void stationery_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(count.Text.Length > 0)
            {
                int price = Sub.KnowPrice(stationery.SelectedItem.ToString());
                int thing;
                int.TryParse(count.Text, out thing);
                if (count.Text.Length > 0 && (count.Text.Length > 0) && ((thing != 0) && (thing > 0)))
                    sum_buy.Text = (price * thing).ToString();
            }
        }

        //Кнопка "Рассчитать"
        private void calculate_Click(object sender, RoutedEventArgs e)
        {
            if (Sub.CheckNumber(sum_buy.Text) && Sub.CheckNumber(sum_accept.Text) && Sub.CheckNumber(percent.Text))
            {
                double sum_buy_T = Convert.ToInt32(sum_buy.Text);
                double sum_accept_T = Convert.ToInt32(sum_accept.Text);
                double percent_T = Convert.ToInt32(percent.Text);

                double change_T = sum_accept_T - (sum_buy_T - ((sum_buy_T / 100 * percent_T)));
                if (change_T > 0)
                {
                    change_T = Math.Round(change_T);
                    change.Text = change_T.ToString();
                    Sub.Add_History(stationery.Text, count.Text, percent.Text, change.Text);
                    MessageBox.Show("Успешно)))");
                }
                else
                    MessageBox.Show("Нет денег(");
            }
            else MessageBox.Show("Не все поля заполнены или некорректно!");
        }

        private void save_Click(object sender, RoutedEventArgs e)
        {
            var fileName = $"{"Чек"}_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.docx";//Имя + указание даты создания файла воизбежании замены файла с таким же названием
            //var savePath =  System.IO.Path.GetFullPath(@"..\..\..\Фотки");
            var savePath = @"Z:\ИС-20-1 Головкин\C#\14pr\Prac_14\Prac_14\Cheks\" + fileName;
            //savePath = savePath + "\\" + fileName; // Тут указываем путь сохранения файла

            var wordApp = new Application();
            var document = wordApp.Documents.Add();
            document.Content.SetRange(0, 0);

            var companyName = "ООО \"Dead Inside (хи-хи ха-ха)\"";
            var welcomeText = "Добро пожаловать";
            var kkmNumber = "ККМ 00075411 #3969";
            var inn = "ИНН 1087746942040";
            var ekls = "ЭКЛЗ 3851495566";
            Random random = new Random();
            int num = random.Next(000000001, 999999999);
            var checkNumber = $"Чек №{num}";
            var dateTime = $"{DateTime.Now.ToString("yyyyMMdd_HHmmss")} СИС.";
            var stuffname = $"{stationery.Text}";
            var line = "----------------------";


            document.Content.Text = $"{companyName}\n{welcomeText}\n{kkmNumber}\n{inn}\n{ekls}\n{checkNumber}\n{dateTime}\n{line}" +
                $"\nНазвание товара: {stuffname}\nКоличество: {count.Text}\nСумма товара: {sum_buy.Text}\nСумма принято: {sum_accept.Text}\nСдача: {change.Text}\n{line}";

            document.SaveAs2(savePath);
            document.Close();
            wordApp.Quit();
            MessageBox.Show("Чек сохранён!");
        }
    }
}
