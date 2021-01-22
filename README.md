# datatableToExcel Metodunun Kullanımı
datatableToExcel metodu string bir başlık dizisi , DataTable türünde bir tablo, tema ve integer bir toplam dizisi alır.
Bu metoda gönderdiğiniz veriler ile geriye byte dizisi döner.Oluşan Excel'de 1. satırı atlayarak başlıkları alt alta yazar seçtiğiniz ve seçtiğiniz temaya uygun bir arkaplanı rengi atar ve fontunu bold yapar.
Başlığın veya başlıkların bir alt satırın altına oluşturulan tarihi yazar.
Oluşturulan tarihin bir alt satırına ise DataTable'den oluşturulan tabloyu yazar.

# wcfExcel Solution'un Kullanımı

try
{
  string[] Baslik = { "KAHVALTI", "SULU YEMEK", "SEBZE YEMEKLERİ"}; // String düründe bir başlık dizisi tanımladım.
  
  DataTable dt = new DataTable("Tablo"); // dataTable'mı oluşturdum.
  dt.Columns.Add("ID", typeof(int));
  dt.Columns.Add("ÜRÜN ADI", typeof(string));
  dt.Columns.Add("FİYAT", typeof(double));
  
  dt.Rows.Add(1, "PEYNİR", 20.65);
  dt.Rows.Add(2, "KIYMALI PATATES YEMEĞİ TARİFİ", 60.99);
  dt.Rows.Add(3, "FIRINDA BÜTÜN TAVUK TARİFİ", 70.25);
  dt.Rows.Add(4, "DİYET ALİNAZİK TARİFİ", 89.99); //sütunları ve satırları oluşturdum.

  int[] toplam = new int[] { 3 }; //  toplam dizisini oluşturduk. Toplamasını istediğiz sütunları buraya yazarsanız double veya decimal olan sayıları toplayıp son satıra toplam olarak eklenecektir.
  
  toExcel.ExcelClient excel = new toExcel.ExcelClient();
  
  byte[] veri = excel.datatableToExcel(Baslik, dt, toExcel.Tema.Kırmızı, toplam); // Geriye size bir byte dizisi dönecektir. Bundan sonra ise byte dizisini kullanarak excel tablosunu indirebilirsiniz.
  
}
catch (Exception ex)
{
  ScriptManager.RegisterStartupScript(this, GetType(), "scriptKodu", "alert('Hata " + ex.Message + "');", true);
}
