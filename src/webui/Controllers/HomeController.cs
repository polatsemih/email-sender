using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using webui.EmailSender;
using webui.Models;

namespace webui.Controllers
{
    public class HomeController : Controller
    {
        private readonly ISmtpEmailSender _smtpemailSender;

        public HomeController(ISmtpEmailSender smtpemailSender)
        {
            this._smtpemailSender = smtpemailSender;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public IActionResult ImportExcel()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ImportExcel(IFormFile file) {
            if (file == null)
            {
                return NotFound();
            }

            var fileExtension = Path.GetExtension(file.FileName);
            if (fileExtension != ".xlsx" || fileExtension != ".xlsm")
            {
                System.Data.DataTable dataTable = new System.Data.DataTable();

                using (var memoryStream = new MemoryStream()) {
                    file.CopyTo(memoryStream);

                    using (var workbook = new XLWorkbook(memoryStream)) {
                        var worksheet = workbook.Worksheet(1);

                        int i;
                        for (i = 1; i <= worksheet.Columns().Count(); i++) {
                            dataTable.Columns.Add(worksheet.Cell(1, i).Value.ToString());
                        }

                        for (i = 2; i <= worksheet.Rows().Count(); i++) {
                            DataRow dataRow = dataTable.NewRow();

                            int j;
                            for (j = 1; j <= worksheet.Columns().Count(); j++) {
                                dataRow[j - 1] = worksheet.Cell(i, j).Value;
                            }

                            dataTable.Rows.Add(dataRow);
                        }
                    }
                }

                var dataTableJson = JsonConvert.SerializeObject(dataTable);
                var excelContents= JsonConvert.DeserializeObject<List<ExcelFirstColumnNamesModel>>(dataTableJson);
                
                List<string> userEmails = new List<string>();

                foreach (var item in excelContents)
                {
                    userEmails.Add(item.Email);
                }
                
                ViewBag.UserEmails = userEmails;
                return View();
            }
            else
            {
                return NotFound();
            }
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SendEmail(EmailModel model, string[] recipientemails)
        {
            if (String.IsNullOrEmpty(model.EmailSubject) || String.IsNullOrEmpty(model.EmailMessage) || recipientemails.Count() < 1)
            {
                return NotFound();
            }

            var emailMessage = System.Web.HttpUtility.HtmlEncode(model.EmailMessage.Trim());
            var emailSubject = System.Web.HttpUtility.HtmlEncode(model.EmailSubject.Trim());

            string htmlMessage = "<!DOCTYPE html>" +
                                "<html lang=\"en\">" +
                                    "<head>" +
                                        "<meta charset=\"UTF-8\">" +
                                        "<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">" +
                                        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">" +
                                        "<title>Email Sender</title>" +
                                    "</head>" +
                                    "<body style=\"margin: 0; padding: 0; border: 0;\">" +
                                        "<div style=\"min-height: 100%; max-width: 800px; background-color: #DDDDDD; margin-left: auto; margin-right: auto; text-align: center;\">" +
                                            "<header style=\"padding-top: 50px; padding-bottom: 50px;\">" +
                                                "<h1 style=\"font-size: 40px; color: #A76C00;\">Email Title</h1>" +
                                                $"<p style=\"font-size: 20px; padding-bottom: 10px; color: #000000;\">{emailMessage}</p>" +
                                            "</header>" +
                                            "<footer style=\"background-color: #8E8E8E; padding-top: 10px; padding-bottom: 10px;\">" +
                                                "<p style=\"font-size: 14px; color: #000000; opacity: 0.7;\">Designed by Semih POLAT</p>" +
                                                "<p style=\"font-size: 14px; color: #000000;\">For your questions and offers, contact <a href=\"mailto:polatsemih@protonmail.com\" style=\"font-size: 14px; color: #A76C00;\">polatsemih@protonmail.com</a></p>" +
                                            "</footer>" +
                                        "</div>" +
                                    "</body>" +
                                "</html>";

            await _smtpemailSender.SendEmailAsync(recipientemails, emailSubject, htmlMessage);

            return RedirectToAction("Index", "Home");
        }
    }
}
