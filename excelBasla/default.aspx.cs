using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace excelBasla
{
    public partial class _default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                Excel.ExcelClient ex = new Excel.ExcelClient();

                //byte[] veri= ex.ExcelPivot();

                string[] baslik = { "baslik1", "baslik2", "baslik-3", "Kişi Sayısı" };

                DataTable dt = new DataTable();

                dt.TableName = "MySampleTable";
                dt.WriteXmlSchema("sample.xsd");

                dt.Columns.Add("ID", typeof(int));
                dt.Columns.Add("EARTH/COUNTRIES", typeof(string));
                dt.Columns.Add("CITIES", typeof(string));
                dt.Columns.Add("MONEY", typeof(double));
                dt.Columns.Add("PERSON COUNT", typeof(int));

                //add some rows
                dt.Rows.Add(0, "Country", "Netherlands", 56.65, 4);
                dt.Rows.Add(1, "Country", "Japan", 12.5, 4);
                dt.Rows.Add(2, "Country", "America", 0.1, 94);
                dt.Rows.Add(3, "State", "Gelderland", 40.1, 5);
                dt.Rows.Add(4, "State", "Texas", 50.1, 7);
                dt.Rows.Add(5, "State", "Echizen", 60.1, 5);
                dt.Rows.Add(6, "City", "Amsterdam", 90.1, 4);
                dt.Rows.Add(7, "City", "Tokyo", 100.1, 3);
                dt.Rows.Add(8, "City", "New York", 560.1, 1);

                byte[] veri = ex.Basla(baslik, dt);

                using (var memoryStream = new MemoryStream())
                {
                    Response.Charset = "utf-8";
                    Response.Buffer = true;
                    Response.ContentEncoding = System.Text.Encoding.UTF8;
                    Response.AddHeader("content-disposition", "attachment; filename=dene.xlsx");
                    Response.BinaryWrite(veri);
                    memoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.Close();
                }
            }
            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }
        }
    }
}