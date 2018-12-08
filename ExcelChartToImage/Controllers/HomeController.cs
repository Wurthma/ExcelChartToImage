using ExcelChartToImage.Models;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelChartToImage.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            HomeModel model = new HomeModel();
            model.ListaExcelChartImg = new List<byte[]>();
            //Pegar o caminho do projeto
            string path = Server.MapPath("~");
            //Carregar o arquivo excel
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(path + "\\Content\\column-chart.xlsx");
            //Carregar a planilha 1
            Worksheet sheet = workbook.Worksheets[0];
            Image[] imgs = workbook.SaveChartAsImage(sheet);
            //Caso queira grava em MemoryStream
            var ms = new MemoryStream();

            for (int i = 0; i < imgs.Length; i++)
            {
                //Salva arquivo em uma pasta
                imgs[i].Save(string.Format(path + "\\Images\\img-{0}.jpeg", i), ImageFormat.Jpeg);
                //Para MemoryStream instanciado anteriormente
                imgs[i].Save(ms, ImageFormat.Jpeg);
                //Jogando para model para imprimir imagens na view
                model.ListaExcelChartImg.Add(ms.ToArray());
            }

            return View(model);
        }
    }
}