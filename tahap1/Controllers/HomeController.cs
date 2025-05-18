using Microsoft.AspNetCore.Mvc;
using Rotativa.AspNetCore;
using ClosedXML.Excel;
using System.IO;

namespace tahap1.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View(GetSampleData());
        }

        public IActionResult ExportPdf()
        {
            var data = GetSampleData();
            return new ViewAsPdf("pdftemplate", data)
            {
                FileName = "Data.pdf"
            };
        }

        public IActionResult ExportExcel()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Data");

            worksheet.Cell(1, 1).Value = "Nama";
            worksheet.Cell(1, 2).Value = "Umur";

            var data = GetSampleData();
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cell(i + 2, 1).Value = data[i].Nama;
                worksheet.Cell(i + 2, 2).Value = data[i].Umur;
            }

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;

            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Data.xlsx");
        }

        private List<Person> GetSampleData() => new()
        {
            new Person { Nama = "Andi", Umur = 22 },
            new Person { Nama = "Budi", Umur = 23 }
        };

        public class Person
        {
            public string Nama { get; set; }
            public int Umur { get; set; }
        }
    }
}
