using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Collections.Generic;

namespace marché
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<Product> products = initList();

            int peachCount = products.Where(o => o.Name == "Pêches").Count();
            Console.WriteLine($"Il y a {peachCount} vendeurs de pêches");

            Product maxWatermelon = products.Where(o => o.Name == "Pastèques").OrderBy(o => o.Quantity).Last();
            Console.WriteLine($"C'est {maxWatermelon.Producer} qui a le plus de pastèques (stand {maxWatermelon.Location}, {maxWatermelon.Quantity} pièces)");
        }

        static List<Product> initList()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\pa78gum\Documents\GitHub\323-Programmation_fonctionnelle\exos\marché\Place du marché.xlsx");
            Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[2];

            Excel.Range xlRange = xlWorkSheet.UsedRange;
            Excel.Range firstRow = xlRange.Rows[1];
            firstRow.Delete();
            xlRange = xlWorkSheet.UsedRange;

            List<Product> products = new List<Product>();

            foreach (Excel.Range row in xlRange.Rows)
            {
                int location = Convert.ToInt32((row.Cells[1] as Excel.Range).Value2);
                string producer = Convert.ToString((row.Cells[2] as Excel.Range).Value2);
                string name = Convert.ToString((row.Cells[3] as Excel.Range).Value2);
                int quantity = Convert.ToInt32((row.Cells[4] as Excel.Range).Value2);
                string unity = Convert.ToString((row.Cells[5] as Excel.Range).Value2);
                float price = Convert.ToSingle((row.Cells[6] as Excel.Range).Value2);

                Product p = new Product(location, producer, name, quantity, unity, price);
                products.Add(p);
            }
            return products;
        }
    }
    public class Product
    {
        public int Location;
        public string Producer;
        public string Name;
        public int Quantity;
        public string Unity;
        public float Price;

        public Product(int location, string producer, string name, int quantity, string unity, float price)
        {
            Location = location;
            Producer = producer;
            Name = name;
            Quantity = quantity;
            Unity = unity;
            Price = price;
        }
    }
}
