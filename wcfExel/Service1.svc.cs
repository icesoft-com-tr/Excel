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

        /// <summary>
        /// Datatable top Excel
        /// </summary>
        /// <param name="Baslik">Excel Başlık metni</param>
        /// <param name="dt">Datatable</param>
        /// <param name="t">Tema</param>
        /// <param name="toplam">Footer satırına toplamları alınacak sütunların listesi, null ise footer olmaz</param>
        /// <returns></returns>
        public byte[] datatableToExcel(string[] Baslik, DataTable dt, Tema t = Tema.Mavi, int[] toplam = null)
        {
            try
            {
                byte[] r;

                Color arkaplan = Color.FromArgb(219, 229, 241);

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
                        //ws.Column(i).Style.HorizontalAlignment = (ExcelHorizontalAlignment)ExcelVerticalAlignment.Center;
                        ws.Cells[i+1, 1].Style.Font.Size = 12;
                        ws.Cells[i + 1, 1].Style.Font.Bold = true;
                    }//Başlıkları kolon uzunluğu kadar başıkları ekliyoruz.

                    ws.Cells[1 + Baslik.Length + 1, 1, 1 + Baslik.Length + 1, dt.Columns.Count].Merge = true;
                    ws.Cells[1 + Baslik.Length + 1, 1].Value = DateTime.Now.ToString("dddd, dd MMMM yyyy hh:mm");
                    ws.Cells[1 + Baslik.Length + 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


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
                    ws.Cells[1 + Baslik.Length + 1 + dt.Rows.Count + 1 ,1, 1 + Baslik.Length + 1 + dt.Rows.Count + 1, dt.Columns.Count].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    ws.Cells[Baslik.Length + 3, 1, Baslik.Length + 3, dt.Columns.Count].Style.Border.Top.Color.SetColor(Color.Black);
                    ws.Cells[Baslik.Length + 3, 1, Baslik.Length + 3, dt.Columns.Count].Style.Border.Bottom.Color.SetColor(Color.Black);
                    ws.Cells[1 + Baslik.Length + 1 + dt.Rows.Count + 1, 1, 1 + Baslik.Length + 1 + dt.Rows.Count + 1, dt.Columns.Count].Style.Border.Bottom.Color.SetColor(Color.Black);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (i%2 == 0)
                        {
                            ws.Cells[Baslik.Length + 4 + i, 1, Baslik.Length + 4 + i, dt.Columns.Count].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[Baslik.Length + 4 + i, 1, Baslik.Length + 4 + i, dt.Columns.Count].Style.Fill.BackgroundColor.SetColor(arkaplan);
                        }
                    }//dataTable'ın arkaplan renginin uygulama bölümü

                    for (int i = 1; i <= dt.Columns.Count; i++)
                    {
                        ws.Column(i).AutoFit(8.32);
                        ws.Column(i).Style.WrapText = false;
                        if (dt.Columns[i - 1].DataType == typeof(int))
                        {
                            ws.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        if (dt.Columns[i-1].DataType == typeof(double) || dt.Columns[i - 1].DataType == typeof(decimal))
                        {
                            ws.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(i).Style.Numberformat.Format = "#,##0.00";
                        }
                    }//integer değerlerin ortalanıp, double ve decimal değerlerinin sağa yaslanıp formatlanma bölümü

                    if (toplam != null)
                    {
                        Array.Sort(toplam);//Diziyi küçükten büyüğe doğru sıraladık.

                        if (toplam[0] != 1)
                        {
                            ws.Cells[1 + Baslik.Length + 1 + dt.Rows.Count + 2, 1, 1 + Baslik.Length + 1 + dt.Rows.Count + 2, toplam[0]-1].Merge = true;
                            ws.Cells[1 + Baslik.Length + 1 + dt.Rows.Count + 2, 1, 1 + Baslik.Length + 1 + dt.Rows.Count + 2, 1].Value = "TOPLAM";
                            ws.Cells[1 + Baslik.Length + 1 + dt.Rows.Count + 2, 1, 1 + Baslik.Length + 1 + dt.Rows.Count + 2, 1].Style.Font.Bold = true;
                            ws.Cells[1 + Baslik.Length + 1 + dt.Rows.Count + 2, 1, 1 + Baslik.Length + 1 + dt.Rows.Count + 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }

                        for (int i = 0; i < toplam.Count(); i++)
                        {
                            double sum = 0;
                            for (int j = 0; j < dt.Rows.Count; j++)
                            {
                                if (dt.Columns[toplam[i]-1].DataType == typeof(double) || dt.Columns[toplam[i]-1].DataType == typeof(decimal))
                                {
                                    sum += Convert.ToDouble(dt.Rows[j][toplam[i]-1].ToString());
                                }
                            }
                            ws.Cells[1 + Baslik.Length + 1 + dt.Rows.Count + 2, toplam[i], 1 + Baslik.Length + 1 + dt.Rows.Count + 2, toplam[i]].Value = sum;
                        }
                    }//toplam satırını oluşturulup hesaplamala bölümü

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
