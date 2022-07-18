using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Text;
using WebExcel2.Models;
using IronXL;
using System.Data;

namespace WebExcel2.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public FileResult DownloadXLS()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("1,2,3\r\n");
            sb.Append("4,5,6\r\n");
            return File(Encoding.UTF8.GetBytes(sb.ToString()), "text/csv", "Grid.csv");
        }

        [HttpGet]
        public FileResult DownloadXLS2()
        {

            String[] names = new string[] { "Travis Johnson", "Kyle Johnson", "Rocky Johnson", "Storm Johnson" };
            DataTable people = new DataTable();
            people.Columns.Add("First");
            people.Columns.Add("Last");

            foreach (string name in names)
            {
                var pieces = name.Split(" ");
                var row = people.NewRow();
                row["First"] = pieces[0];
                row["Last"] = pieces[1];
                people.Rows.Add(row);
            }

            DataSet ds = new DataSet("mydataset");
            ds.Tables.Add(people);
            WorkBook wb = WorkBook.Load(ds);

            // wb.SaveAs(@"c:\work\people.xls");

            return File(wb.ToStream(), "application/vnd.ms-excel", "People.xlsx");
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}