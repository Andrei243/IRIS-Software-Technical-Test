using IRIS_Software_Technical_Test.Enums;
using System.Text;

namespace IRIS_Software_Technical_Test
{
    internal class Program
    {
        private static bool ReadAndCheckExcelColumn(out string? input)
        {
            input = Console.ReadLine();
            return input!= null && input.All(e => char.IsAsciiLetterUpper(e));
        }

        private static void ConvertExcelToNumber()
        {
            var excelConvertOpningText = "Please type the Excel column you want to convert using only letters from 'A' to 'Z'";
            Console.WriteLine(excelConvertOpningText);
            string? column;
            while(!ReadAndCheckExcelColumn(out column))
            {
                Console.WriteLine("You need to type a word only with uppercase letters");
                Console.WriteLine(excelConvertOpningText);
            }
            var power = 1;
            var columnNumber = 0;
            while(column?.Length > 0)
            {
                columnNumber += power * (column.Last() - 'A' + 1);
                power *= 26;
                column = column.Substring(0, column.Length - 1);
            }
            Console.WriteLine(columnNumber);
        }

        private static void ConvertNumberToExcel()
        {
            var numberConvertOpeningText = "Please type the number you want to convert to an Excel column";
            Console.WriteLine(numberConvertOpeningText);
            long columnNumber = 0;

            while(!long.TryParse(Console.ReadLine(), out columnNumber) && columnNumber <= 0)
            {
                Console.WriteLine("You need to type a number bigger than 0");
                Console.WriteLine(numberConvertOpeningText);
            }

            var columnbuilder = new StringBuilder();
            while(columnNumber > 0)
            {
                var modulo = (columnNumber - 1) % 26;
                columnbuilder.Append((char)('A' + modulo));
                columnNumber = (columnNumber-modulo) / 26;
            }
            Console.WriteLine(columnbuilder.ToString());
        }

        static void Main()
        {
            var openingText = "Type 1 to convert from  Excel column to number or 2 for number to Excel column";
            Console.WriteLine(openingText);
            int option;
            while(!int.TryParse(Console.ReadLine(), out option) || !Enum.IsDefined(typeof(OperationType), option))
            {
                Console.WriteLine("You need to either type 1 or 2");
                Console.WriteLine(openingText);
            }

            switch ((OperationType)option)
            {
                case OperationType.ExcelToNumber:
                    ConvertExcelToNumber();
                    break;
                case OperationType.NumberToExcel:
                    ConvertNumberToExcel();
                    break;
            }
        }
    }
}