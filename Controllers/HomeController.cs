using Aspose.Pdf;
using ExcelToPdf.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Threading.Tasks;

namespace ExcelToPdf.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Index(FileUpload upload)
        {
            if (upload.File.Length <= 0)
                return View();

            string extension = Path.GetExtension(upload.File.FileName);
            if (extension != ".jpg" && extension != ".png")
                return View();

            string originalfileName = upload.File.FileName;
            string pdFileName = originalfileName.Replace(Path.GetExtension(upload.File.FileName), ".pdf");
            string OriginalfilePath = Path.Combine("UploadedImages", "Original", originalfileName);
            string ConvertedfilePath = Path.Combine("UploadedImages", "Converted", pdFileName);

            using (var stream = System.IO.File.Create(OriginalfilePath))
            {
                upload.File.CopyTo(stream);
            }

            System.Drawing.Image srcImage = System.Drawing.Image.FromFile(OriginalfilePath);
            Aspose.Pdf.Document doc = new();

            Page page = doc.Pages.Add();
            Image image = new()
            {
                File = OriginalfilePath
            };

            page.PageInfo.Height = (srcImage.Height);
            page.PageInfo.Width = (srcImage.Width);
            page.PageInfo.Margin.Bottom = (0);
            page.PageInfo.Margin.Top = (0);
            page.PageInfo.Margin.Right = (0);
            page.PageInfo.Margin.Left = (0);
            page.Paragraphs.Add(image);

            doc.Save(ConvertedfilePath);
            byte[] fileBytes = System.IO.File.ReadAllBytes(ConvertedfilePath);
            return File(fileBytes, "application/force-download", pdFileName);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
