using ClosedXML.Excel; //https://www.nuget.org/packages/ClosedXML/
using System.ComponentModel;

namespace ConsoleTesteData
{
    class Program
    {
        static void Main(string[] args)
        {
            var monthList = GenerateMonths();
            GenerateFile(monthList, "AnnualCalendar.xlsx", "Provisorio" );
        }

        static void GenerateFile(ICollection<String> monthList, string fileName, string tets)
        {
            string filePathName = System.IO.Directory.GetCurrentDirectory() + "\\"+ fileName;
            if (File.Exists(filePathName))
                File.Delete(filePathName);

            using (var workbook = new XLWorkbook())
            {
                foreach (var month in monthList)
                {
                    int count = 1;
                    var worksheet = workbook.Worksheets.Add(month.ElementAt(count));
                     
                    count ++;
                }
                workbook.SaveAs(filePathName);
            }
        }

        /// <summary>
        /// Generate months of the year to create tables in the file. xlsx
        /// </summary>
        /// <returns>Every month of the year.</returns>
        static List<String> GenerateMonths()
        {
            List<string> retur = new List<string>();
            
            for (int i = 1; i <= 12; i++)
            {
                DateTime firstDayOfTheMonth = new DateTime(DateTime.Now.Year, i, 1);
                string nameOfTheMonth = firstDayOfTheMonth.ToString("MMMM");
                retur.Add(nameOfTheMonth);
            }

            return retur;
        }

        static string GenerateDaysOfTheWeek(IXLWorksheet worksheet, Dictionary<int, int> yearAndMonthCurrent)
        {
            // Preenche os dias da semana
            for (int dia = 1; dia <= diasNoMes; dia++)
            {
                DateTime dataAtualizada = new DateTime(yearAndMonthCurrent.Keys, mes, dia);

                // Determina o dia da semana (0 = domingo, 1 = segunda, ..., 6 = sábado)
                int diaSemana = (int)dataAtualizada.DayOfWeek;

                // Preenche a célula correspondente ao dia
                worksheet.Cells(mes + 1, diaSemana + 2).Value = dia;
            }
        }

    }
}