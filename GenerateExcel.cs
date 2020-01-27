using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System.IO;
using System.Net.Mail;
using System.Net;

namespace ERNI.GenerateExcel
{
    public static class GenerateExcel
    {
        [FunctionName("GenerateExcel")]
        public static void Run([TimerTrigger("* * * * *")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            using(var excel = new ExcelPackage()) {
                ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("HelloSheet");

                sheet.Cells["B3"].Value ="Hello World";

                excel.Workbook.Properties.Title = "Sample Excel";
                excel.Workbook.Properties.Author = "Niel";

                //Send email
                var mail = new MailMessage();
                mail.To.Add("arni@erni.ph");
                mail.Subject = "Azure Function Test";
                mail.From = new MailAddress("cool@gmail.com");
                mail.Body = "Hi All, ";

                var mem = new MemoryStream(excel.GetAsByteArray());
                mem.Position = 0;
                mail.Attachments.Add(new Attachment(mem, "HelloSheet.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

                var smpt = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    Credentials = new NetworkCredential("cool@gmail.com", "mycoolpassword")
                };

                smpt.Send(mail);


            }
        }
    }
}
