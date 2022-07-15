using OfficeOpenXml;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web.Mvc;

namespace Epplus_Test.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult GenerarExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            String connectionString = ConfigurationManager.ConnectionStrings["cnn"].ConnectionString;
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                String sql = "SELECT * FROM PERSONA";

                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(sql.ToString(), con);
                da.Fill(dt);

                String nombreArchivo = DateTime.Now.ToShortDateString().Replace('-', '_').Replace('/', '_') + "_" + DateTime.Now.ToShortTimeString().Replace(' ', '_').Replace(':', '_') + ".xlsx";
                String rutaConArchivo = AppDomain.CurrentDomain.BaseDirectory + @"\Files\" + nombreArchivo.ToString();
                Stream documento = System.IO.File.Create(rutaConArchivo);

                using (ExcelPackage package = new ExcelPackage(documento))
                {
                    using (ExcelWorksheet ws = package.Workbook.Worksheets.Add("Informe"))
                    {
                        ws.Cells["A1"].LoadFromDataTable(dt, true);
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        package.SaveAs(documento);
                    }
                }
                documento.Close();
                return File(new FileStream(rutaConArchivo, FileMode.Open), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreArchivo);
            }
        }
    }
}