using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data.Sql;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

namespace Prac_14
{
    public static class Sub
    {
        //Подключение к бд
        public static string connection = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\maxim\Desktop\шарашка\C#\Prac_14\DB.mdf;Integrated Security=True;Connect Timeout=30";
        public static SqlConnection sql = new SqlConnection(connection);

        //Запрос для списка товаров
        public static List<string> Full_Combobox()
        {
            List<string> list = new List<string>();
            sql.Open();

            SqlCommand com = new SqlCommand($"select Наименование from Stationery", sql);
            SqlDataReader dr = com.ExecuteReader();

            while (dr.Read())
            {
                list.Add(dr.GetString(0));
            }

            sql.Close();
            return list;
        }

        //Получаем цену товара
        public static int KnowPrice(string stationery)
        {
            sql.Open();

            SqlCommand com = new SqlCommand($"select [Цена] from [Stationery] where [Наименование] = '{stationery}'", sql);
            int price = Convert.ToInt32(com.ExecuteScalar());

            sql.Close();
            return price;
        }

        public static bool CheckNumber(string number)
        {
            Regex r = new Regex(@"^[0-9]+$");
            if (r.IsMatch(number)) return true;
            else return false;

            //@"^(?=.{6,}$)(?=.*\d)"
            //@"^\d$"
            // @"^[0-9]+$"
        }

        public static void Add_History(string name, string count, string percent, string price)
        {
            sql.Open();
            SqlCommand com1 = new SqlCommand($"select Код from [Stationery] where [Наименование] = '{name}'", sql);
            int id = Convert.ToInt32(com1.ExecuteScalar());
            SqlCommand com2 = new SqlCommand($"insert into Sell(Дата, Код_товара, Количество, Скидка, Стоимость) values(GETDATE(), {id}, {Convert.ToInt32(count)}, {Convert.ToInt32(percent)}, {Convert.ToInt32(price)})", sql);
            com2.ExecuteNonQuery();
            sql.Close();
        }
    }
}
