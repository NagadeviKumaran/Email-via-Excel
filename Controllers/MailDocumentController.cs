using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Net.Mail;
using System.Net;
using MimeKit;
using System.IO;
using System.Text;

namespace ExcelEmail.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MailDocumentController : ControllerBase
    {
        public MailDocumentController()
        {
            // Set the license context for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }


        [HttpPost("UploadAndSend")]
        public async Task<IActionResult> UploadAndSend(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
                return BadRequest("Please select a valid file to upload.");

            var employeeDetails = new List<(string EmployeeName, string CompanyName, string Month, string Year, string Email)>();

            try
            {
                using (var stream = new MemoryStream())
                {
                    await excelFile.CopyToAsync(stream);

                    // Process the uploaded Excel file using EPPlus
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0]; // Use the first worksheet
                        var rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++) // Skip header row
                        {
                            var employeeName = worksheet.Cells[row, 1].Text;
                            var companyName = worksheet.Cells[row, 2].Text;
                            var month = worksheet.Cells[row, 3].Text;
                            var year = worksheet.Cells[row, 4].Text;
                            var email = worksheet.Cells[row, 5].Text;

                            if (!string.IsNullOrWhiteSpace(employeeName) && !string.IsNullOrWhiteSpace(email))
                            {
                                employeeDetails.Add((employeeName, companyName, month, year, email));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error reading Excel file: {ex.Message}");
            }

            // Process each employee's details and send emails
            foreach (var (employeeName, companyName, month, year, email) in employeeDetails)
            {
                try
                {
                    string pdfFileName = $"{employeeName.Trim()} - {companyName} - {month} {year}.pdf";
                    byte[] pdfContent = GetPdfContent(employeeName, companyName, month, year);

                    var emailSent = await SendEmailAsync(email, pdfContent, pdfFileName);
                    if (!emailSent)
                    {
                        return StatusCode(500, $"Failed to send email to {email}");
                    }
                }
                catch (FileNotFoundException fnfe)
                {
                    return NotFound(fnfe.Message);
                }
                catch (Exception ex)
                {
                    return StatusCode(500, $"Error processing {email}: {ex.Message}");
                }
            }

            return Ok("File uploaded and emails sent successfully.");
        }

        private byte[] GetPdfContent(string employeeName, string companyName, string month, string year)
        {
            string documentsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Documents");
            string pdfFileName = $"{employeeName.Trim()} - {companyName} - {month} {year}.pdf";
            string pdfFilePath = Path.Combine(documentsFolder, pdfFileName);

            if (!System.IO.File.Exists(pdfFilePath))
            {
                throw new FileNotFoundException($"File not found for {employeeName}: {pdfFilePath}");
            }

            return System.IO.File.ReadAllBytes(pdfFilePath);
        }

        private async Task<bool> SendEmailAsync(string emailAddress, byte[] pdfContent, string pdfFileName)
        {
            try
            {
                using (var smtpClient = new SmtpClient("smtp.gmail.com")
                {
                    Port = 587,
                    Credentials = new NetworkCredential("nagadevikumaran01@gmail.com", "lnde fzwr vgbz ybny"), // Replace with your email and app password
                    EnableSsl = true,
                })
                {
                    var mailMessage = new MailMessage
                    {
                        From = new MailAddress("nagadevikumaran01@gmail.com"), // Replace with your email
                        Subject = "Salary Slip",
                        Body = $"Hi, please find attached the salary slip for {pdfFileName}.",
                        IsBodyHtml = true,
                    };

                    mailMessage.To.Add(emailAddress);

                    if (pdfContent != null && pdfContent.Length > 0)
                    {
                        var attachment = new Attachment(new MemoryStream(pdfContent), pdfFileName, "application/pdf");
                        mailMessage.Attachments.Add(attachment);
                    }

                    await smtpClient.SendMailAsync(mailMessage);
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to send email to {emailAddress}: {ex.Message}");
                return false;
            }
        }
    }
}