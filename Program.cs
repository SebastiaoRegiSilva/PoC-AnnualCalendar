using System.Globalization;
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
                // Create functions to distribute the code.
                worksheet.Range("A1:G1").Merge().Value = $"Calendário anual de leitura familiar!";
                LettersIntoBold(worksheet, "A1:G1");
                worksheet.Range("A1:G1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Range("D2:E2").Value = "Leitura Bíblica";
                LettersIntoBold(worksheet, "D2:E2");    

                foreach (var day in daysOfMonth)
                {
                    worksheet.Cell("A" + lines).Value = GetWeekOfYear(day);
                    worksheet.Cell("B" + lines).Value = $"{day.ToString($"dd")}";
                    worksheet.Range("A" + lines, $"B" + lines).Style.Font.Bold = true;
                    // Lista de semanas para mesclar verticalmente e alinhar centrado. 
                    worksheet.Range("A" + lines).Merge().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    // Column width in pixel.
                    worksheet.Columns("A","B").Width = 2.5;
                    worksheet.Column("C").Width = 10;
                    if(day.DayOfWeek == DayOfWeek.Sunday)
                        worksheet.Cell("C" + lines).Style.Font.Bold = true;
                    worksheet.Cell("C" + lines).Value = $"{day.DayOfWeek}" ;
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

    /// <summary>
    /// Discover the week of the year.
    /// </summary>
    /// <param name="date">Date.</param>
    static int GetWeekOfYear(DateTime date)
    {
        CultureInfo culture = new CultureInfo("pt-BR");
        Calendar calendar = culture.Calendar;
        CalendarWeekRule rule = culture.DateTimeFormat.CalendarWeekRule;
        DayOfWeek firstDayOfWeek = culture.DateTimeFormat.FirstDayOfWeek;

        return calendar.GetWeekOfYear(date, rule, firstDayOfWeek);
    }
}
