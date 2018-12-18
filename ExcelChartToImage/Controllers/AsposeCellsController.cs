using Aspose.Cells;
using ExcelChartToImage.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;

namespace ExcelChartToImage.Controllers
{
    public class AsposeCellsController : Controller
    {
        // GET: AsposeCells
        public ActionResult Index()
        {
            HomeModel model = new HomeModel();
            model.ListaExcelChartImg = new List<byte[]>();
            //Pegar o caminho do projeto
            string path = Server.MapPath("~");
            //Abrir arquivo excel com Aspose Cells
            Workbook workbook = new Workbook(path + "\\Content\\column-chart.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            foreach (var grafico in worksheet.Charts)
            {
                MemoryStream ms = new MemoryStream();
                grafico.ToImage().Save(ms, ImageFormat.Jpeg);
                byte[] bmpBytes = ms.ToArray();
                model.ListaExcelChartImg.Add(bmpBytes);

            }

            return View(model);
        }
    }
}