using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography; //https://www.nuget.org/packages/ClosedXML/


class Program
{
    static async Task Main()
    {
        await GenerateFileAsync("AnnualCalendar.xlsx");
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
                foreach (var day in daysOfMonth)
                {
                    worksheet.Cells("A" + countDay).Value = $"{day.DayOfWeek}" ;
                    worksheet.Cells("B" + countDay).Value = $"{day.ToString($"dd/{month}")}";
                    countDay ++;
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
}
