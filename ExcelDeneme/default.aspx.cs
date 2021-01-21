using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;

namespace ExcelDeneme
{
    public partial class _default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string[] malzemeler = { "Tuğla", "Alçı","Kiremit","Alçıpan" };

            DataTable dt = new DataTable("Malzemeler");

            dt.Columns.Add("MALZEME ID", typeof(int));
            dt.Columns.Add("MALZEME ADI", typeof(string));
            dt.Columns.Add("MALZEME YAPISI", typeof(string));
            dt.Columns.Add("MALZEME FİYATI", typeof(double));
            dt.Columns.Add("MALZEME DİNAMİĞİ", typeof(string));

            dt.Rows.Add(1, "MALZEME 1", "DEMİR", 65.99,"Güçlü");
            dt.Rows.Add(2, "MALZEME 2", "TAHTA", 99.99, "Zayıf");
            dt.Rows.Add(3, "MALZEME 3", "AHŞAP", 59.99, "Orta");
            dt.Rows.Add(4, "MALZEME 4", "ALİMÜNYUM", 49.99, "Güçlü");
            dt.Rows.Add(5, "MALZEME 5", "TAŞ", 14.99, "Güçlü");
            dt.Rows.Add(6, "MALZEME 6", "DEMİR", 64.99, "Güçlü");
            dt.Rows.Add(7, "MALZEME 7", "AHŞAP", 19.99, "Orta");
            dt.Rows.Add(8, "MALZEME 8", "KİREMİT", 69.99, "İyi");
            dt.Rows.Add(9, "MALZEME 8", "KİREMİT", 69.99, "İyi");


            ToExcel.ExcelClient excel = new ToExcel.ExcelClient();

            byte[] veri = excel.datatableToExcel(malzemeler, dt, ToExcel.Tema.Kırmızı);
            using (var memoryStream = new MemoryStream())
            {
                Response.Charset = "utf-8";
                Response.Buffer = true;
                Response.ContentEncoding = System.Text.Encoding.UTF8;
                Response.AddHeader("content-disposition", "attachment; filename=excel.xlsx");
                Response.BinaryWrite(veri);
                memoryStream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.Close();
            }
        }
    }
}