using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace wcfExel
{
    public enum Tema { Mavi = 1, Yeşil = 2, Kırmızı = 3 }

    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    public class Service1 : IExcel
    {
        public string GetData(int value)
        {
            return string.Format("You entered: {0}", value);
        }

        public byte[] datatableToExcel(string[] Baslik, DataTable dt, Tema t = Tema.Mavi)
        {
            try
            {
                byte[] r;

                Color arkaplan = Color.Blue;

                switch (t)
                {
                    case Tema.Mavi:
                        arkaplan = Color.FromArgb(219, 229, 241);
                        break;
                    case Tema.Yeşil:
                        arkaplan = Color.FromArgb(234, 241, 221);
                        break;
                    case Tema.Kırmızı:
                        arkaplan = Color.FromArgb(242, 221, 220);
                        break;
                }

                using (ExcelPackage excel = new ExcelPackage())
                {
                    ExcelWorksheet ws = excel.Workbook.Worksheets.Add("Sayfa");
                    ws.View.ShowGridLines = false;

                    for (int i = 1; i <= Baslik.Length; i++)
                    {
                        ws.Cells[i+1, 1, i+1, dt.Columns.Count].Merge = true;
                        ws.Cells[i+1, 1].Value = Baslik[i-1];
                        ws.Cells[i+1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Column(i).Style.HorizontalAlignment = (ExcelHorizontalAlignment)ExcelVerticalAlignment.Center;
                        ws.Cells[i+1, 1].Style.Font.Size = 13;
                    }//Başlıkları kolon uzunluğu kadar başıkları ekliyoruz.


                    ws.Cells[Baslik.Length + 3, 1].LoadFromDataTable(dt, true);

                    ws.Row(Baslik.Length + 3).Height = 24;
                    ws.Row(Baslik.Length + 3).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Row(Baslik.Length + 3).Style.Font.Size = 11;
                    ws.Row(Baslik.Length + 3).Style.Font.Color.SetColor(Color.Red);
                    ws.Row(Baslik.Length + 3).Style.Font.Bold = true;
                    ws.Cells[2,1,Baslik.Length + 1,dt.Columns.Count].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    ws.Cells[2, 1, 1 + Baslik.Length, dt.Columns.Count].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[2, 1, 1 + Baslik.Length, dt.Columns.Count].Style.Fill.BackgroundColor.SetColor(arkaplan);

                    ws.Cells[Baslik.Length + 3, 1, Baslik.Length + 3, dt.Columns.Count].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells[Baslik.Length + 3, 1, Baslik.Length + 3, dt.Columns.Count].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    ws.Cells[Baslik.Length + 3, 1, Baslik.Length + 3, dt.Columns.Count].Style.Border.Top.Color.SetColor(Color.Black);
                    ws.Cells[Baslik.Length + 3, 1, Baslik.Length + 3, dt.Columns.Count].Style.Border.Bottom.Color.SetColor(Color.Black);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (i%2 == 0)
                        {
                            ws.Cells[Baslik.Length + 4 + i, 1, Baslik.Length + 4 + i, dt.Columns.Count].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[Baslik.Length + 4 + i, 1, Baslik.Length + 4 + i, dt.Columns.Count].Style.Fill.BackgroundColor.SetColor(arkaplan);
                        }
                    }

                    for (int i = 1; i <= dt.Columns.Count; i++)
                    {
                        ws.Column(i).AutoFit(8.32);
                        ws.Column(i).Style.WrapText = false;
                        if (dt.Columns[i - 1].DataType == typeof(int))
                        {
                            ws.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        if (dt.Columns[i-1].DataType == typeof(double))
                        {
                            ws.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                    }

                    r = excel.GetAsByteArray();
                }
                return r;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
