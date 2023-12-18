using ClosedXML.Excel;//https://www.nuget.org/packages/ClosedXML/

class Program
{
    static async Task Main()
    {
        await GenerateFileAsync("AnnualCalendar.xlsx");
        Console.WriteLine("Arquivo disponível!");
    }

    static async Task GenerateFileAsync(string fileName)
    {
        string filePathName = Directory.GetCurrentDirectory() + "\\"+ fileName;
        if (File.Exists(filePathName))
            File.Delete(filePathName);
        
        // Create the spreadsheet.
        using (var workbook = new XLWorkbook())
        {
            var year = 2024;
            int countMonth = 1;
            //Generate months of the year.
            List<string> monthsOfTheYear = GenerateMonths(year);
            foreach (var month in monthsOfTheYear)
            {
                //Each month will correspond to a table.
                var worksheet = workbook.Worksheets.Add(month);
                // Generate the days of the month.
                List<DateTime> daysOfMonth = GenerateDaysOfMonth(year, countMonth);
                int countDay = 1;
                int lines = 3;
                worksheet.Range("A1:G1").Merge().Value = $"Calendário anual de leitura familiar!";
                worksheet.Cell("C2").Value = "Leitura Bíblica";
                worksheet.Cell("C2").Style.Font.Bold = true;
                LettersIntoBold(worksheet, "A1:G1");
                worksheet.Range("A1:G1").Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                foreach (var day in daysOfMonth)
                {
                    worksheet.Cells("A" + lines).Value = $"{day.ToString($"dd")}";
                    worksheet.Cells("A" + lines).Style.Font.Bold = true;
                    if(day.DayOfWeek == DayOfWeek.Sunday)
                        worksheet.Cells("B" + lines).Style.Font.Bold = true;
                    worksheet.Cells("B" + lines).Value = $"{day.DayOfWeek}" ;
                    countDay ++;
                    lines++;
                }
                
                countMonth ++;
            }
            
            await Task.Run(() => workbook.SaveAs(filePathName));
        }
    }

    /// <summary>
    /// Generate months of the year to create tables in the file. xlsx
    /// </summary>
    /// <param name="year">Current year.</param>
    /// <returns>Every month of the year.</returns>
    static List<string> GenerateMonths(int year)
    {
        List<string> retur = new List<string>();
        
        for (int i = 1; i <= 12; i++)
        {
            DateTime firstDayOfTheMonth = new DateTime(year, i, 1);
            string nameOfTheMonth = firstDayOfTheMonth.ToString("MMMM");
            // First capital letter of month name.
            nameOfTheMonth = char.ToUpper(nameOfTheMonth[0]) + nameOfTheMonth[1..];

            retur.Add(nameOfTheMonth);
        }

        return retur;
    }

    /// <summary>
    /// Generate every day of the month.
    /// </summary>
    /// <param name="year">Current year.</param>
    /// <param name="month">Current month.</param>
    /// <returns>From the beginning to the month.</returns>
    static List<DateTime> GenerateDaysOfMonth(int year, int month)
    {
        var daysOfMonth = new List<DateTime>();

        // Get the first and last day of the month
        var firstDay = new DateTime(year, month, 1);
        var lastDay = firstDay.AddMonths(1).AddDays(-1);

        // Add all days of the month to the list
        for (var currentDate = firstDay; currentDate <= lastDay; currentDate = currentDate.AddDays(1))
            daysOfMonth.Add(currentDate);
        
        return daysOfMonth;
    }

    /// <summary>
    /// Highlight bold letters for use in the application. 
    /// </summary>
    /// <param name="value">Context of use.</param>
    /// <param name="address">Address of columns and/or rows.</param>
    static bool LettersIntoBold(IXLWorksheet value, string address)
    {
        return value.Range(address).Merge().Style.Font.Bold = true;
    }
}
