using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelData
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath;
            Console.WriteLine("Enter The XLSX File Path:");
            filePath = Console.ReadLine();
            var file = new FileInfo(filePath);

            List<ItemModel> dataFromExcel = await LoadExcelFile(file);
            double totPrice = 0;
            Console.WriteLine("Id   First Name    Last Name   Email     Salary");
            foreach (var item in dataFromExcel)
            {               
                Console.WriteLine($"{item.Id} {item.FirstName} {item.LastName} {item.Email} ${item.Salary}");
                totPrice += item.Salary;
            }
            Console.WriteLine("--------------------------------");
            Console.WriteLine($"Total Salary Costs: ${totPrice.ToString("N0")}/Month");
            Console.ReadLine();
        }

        private static async Task <List<ItemModel>> LoadExcelFile(FileInfo file)
        {
            List<ItemModel> output = new();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[0];

            int row = 2;
            int col = 1;
                while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
                {
                    ItemModel item = new();
                    item.FirstName = ws.Cells[row, col].Value.ToString();
                    item.LastName = ws.Cells[row, col + 1].Value.ToString();
                    item.Email = ws.Cells[row, col + 2].Value.ToString();
                    item.Salary = double.Parse(ws.Cells[row, col + 3].Value.ToString());
                    item.Id = int.Parse(ws.Cells[row, col + 4].Value.ToString());
                    output.Add(item);
                    row += 1;
                }
            return output;
        }
    }
}
