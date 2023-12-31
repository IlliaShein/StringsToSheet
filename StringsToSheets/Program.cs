using OfficeOpenXml;
using System.Text.Json;


namespace StringsToSheets
{
    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var builder = WebApplication.CreateBuilder(args);

            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();
            builder.Services.AddAuthorization();

            var app = builder.Build();

            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseHttpsRedirection();
            app.UseAuthorization();


            app.MapPost("/create-sheet", async (HttpContext context) =>
            {
                string requestBody;
                using (StreamReader reader = new StreamReader(context.Request.Body))
                {
                    requestBody = await reader.ReadToEndAsync();
                }

                string[] strings;
                try
                {
                    strings = JsonSerializer.Deserialize<string[]>(requestBody);
                }
                catch (JsonException)
                {
                    context.Response.StatusCode = 400;
                    return;
                }

                CreateSheet(strings);

                context.Response.StatusCode = 200;
            });

            app.Run();
        }

        private static void CreateSheet(string[] strings)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "StringsSheet.xlsx");
            var package = new ExcelPackage(new FileInfo(path));
            ExcelWorksheet resultSheet = package.Workbook.Worksheets.Add("StringsSheet");

            for (int i = 0; i < strings.Length; i++)
            {
                resultSheet.Cells[i + 1, 1].Value = strings[i];
            }

            package.Save();
        }
    }
}
