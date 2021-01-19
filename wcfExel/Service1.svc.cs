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
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    public class Service1 : IExcel
    {
        public string GetData(int value)
        {
            return string.Format("You entered: {0}", value);
        }

        public byte[] GetExcel(string[] Baslik, DataTable dt)
        {
            try
            {
                byte[] r;

                using (ExcelPackage excel = new ExcelPackage())
                {
                    ExcelWorksheet ws = excel.Workbook.Worksheets.Add("Sayfa");

                    for (int i = 0; i < Baslik.Length; i++)
                    {
                        ws.Cells[i + 1, 1].Value = Baslik[i];
                        ws.Cells[i + 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Column(i + 1).Style.HorizontalAlignment = (ExcelHorizontalAlignment)ExcelVerticalAlignment.Center;
                        ws.Cells[i + 1, 1].Style.Font.Size = 13;
                        ws.Cells[i + 1, 1].Value = Baslik[i];
                        ws.Cells[i + 1, 1, i + 1, dt.Columns.Count].Merge = true;
                    }//Başlıkları kolon uzunluğu kadar başıkları ekliyoruz.

                    ws.Cells[Baslik.Length + 2, 1].LoadFromDataTable(dt, true);

                    ws.Row(Baslik.Length + 2).Height = 30;
                    ws.Row(Baslik.Length + 2).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Row(Baslik.Length + 2).Style.Font.Size = 15;
                    ws.Row(Baslik.Length + 2).Style.Font.Bold = true;

                    ws.Cells[Baslik.Length + 2, 1, Baslik.Length + 2 + dt.Rows.Count, dt.Columns.Count].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[Baslik.Length + 2, 1, Baslik.Length + 2 + dt.Rows.Count, dt.Columns.Count].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                    using (ExcelRange Rng = ws.Cells[Baslik.Length + 2, 1, Baslik.Length + 2 + dt.Rows.Count, dt.Columns.Count])
                    {
                        for (int i = Baslik.Length; i <= (Baslik.Length+2+dt.Rows.Count); i++)
                        {
                            if (i%2 == 0)
                            {
                                ws.Row(i).Style.Fill.PatternType = ExcelFillStyle.None;
                            }
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
