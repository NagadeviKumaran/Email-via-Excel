using ClosedXML.Excel;
using ExcelEmail.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MimeKit;
using OfficeOpenXml;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace ExcelEmail.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EmailController : ControllerBase
    {
        public EmailController()
        {
            // Set the license context for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        //[HttpPost("upload")]
        //public async Task<IActionResult> Upload(IFormFile file)
        //{
        //    if (file == null || file.Length == 0)
        //        return BadRequest("No file uploaded");

        //    // Create a MemoryStream to hold the uploaded file
        //    using (var memoryStream = new MemoryStream())
        //    {
        //        // Copy the uploaded file to the MemoryStream
        //        await file.CopyToAsync(memoryStream);

        //        // Reset the position of the stream to the beginning
        //        memoryStream.Position = 0;

        //        // Optionally, modify the Excel file here using EPPlus
        //        // e.g., add or remove sheets, update cell values, etc.

        //        // Return the file as a downloadable response
        //        var fileName = file.FileName; // Original file name
        //        var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; // Excel MIME type

        //        // Return the file directly from the memory stream
        //        return File(memoryStream.ToArray(), contentType, fileName);
        //    }
        //}

        
        [HttpPost("upload-and-send")]
        public async Task<IActionResult> UploadAndSend(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded");

            var emailAddresses = new List<string>();

            // Read the email addresses from the uploaded Excel file
            using (var memoryStream = new MemoryStream())
            {
                await file.CopyToAsync(memoryStream);
                memoryStream.Position = 0;

                using (var package = new ExcelPackage(memoryStream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++) // Skip header row
                    {
                        var email = worksheet.Cells[row, 2].Text;
                        if (!string.IsNullOrEmpty(email))
                        {
                            emailAddresses.Add(email);
                        }
                    }
                }
            }

            if (emailAddresses.Count == 0)
                return BadRequest("No email addresses found in the file.");

            // Prepare PDF content
            byte[] pdfContent = GetPdfContent(); // Get PDF content as byte array
            string pdfFileName = "SalarySlip.pdf"; // The PDF file name

            // Loop through email addresses and send emails with PDF attachment
            foreach (var email in emailAddresses)
            {
                bool isSuccess = await SendEmailAsync(email, pdfContent, pdfFileName);
                if (!isSuccess)
                {
                    return StatusCode(500, $"Failed to send email to {email}");
                }
            }

            return Ok("All emails sent successfully.");
        }

        // Method to send email with PDF attachment
        private async Task<bool> SendEmailAsync(string emailAddress, byte[] pdfContent, string pdfFileName)
        {
            try
            {
                using (var smtpClient = new SmtpClient("smtp.gmail.com") // Replace with your SMTP server
                {
                    Port = 587, // Change to the correct SMTP port
                    Credentials = new NetworkCredential("nagadevikumaran01@gmail.com", "eagn tbuz gsor ohec"), // Replace with your credentials
                    EnableSsl = true,
                })
                {
                    var mailMessage = new MailMessage
                    {
                        From = new MailAddress("nagadevikumaran01@gmail.com"), // Replace with your email address
                        Subject = "Hello",
                        Body = "Hi, please find the attached PDF.",
                        IsBodyHtml = true,
                    };

                    // Add recipient email
                    mailMessage.To.Add(emailAddress);

                    // Attach PDF
                    if (pdfContent != null && pdfContent.Length > 0)
                    {
                        var attachment = new Attachment(new MemoryStream(pdfContent), pdfFileName, "application/pdf");
                        mailMessage.Attachments.Add(attachment);
                    }

                    // Send email
                    await smtpClient.SendMailAsync(mailMessage);
                }

                return true;
            }
            catch (SmtpException ex)
            {
                // Log error (implement proper logging as needed)
                Console.WriteLine($"Failed to send email to {emailAddress}: {ex.Message}");
                return false;
            }
        }

        // Method to get the PDF content as a byte array
        private byte[] GetPdfContent()
        {
            // Here you can generate a PDF using FastReport, iTextSharp, or any other library
            // For demonstration, we will read a file from the disk
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "pdf", "SalarySlip.pdf"); // Replace with the path to your PDF file
            return System.IO.File.ReadAllBytes(filePath);
        }
    }
}


