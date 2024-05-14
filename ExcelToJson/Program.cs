// bu program Medyasoft için yapılmıştır. Carbon Emisyon modülü için LoadSheet.xlsx dosyasının sene de bir güncellenmesi durumuna karşın excelin istenilen json formatına dönüşümü yapılmıştır.
// yeni gelen exceli de fonksiyona dosya yolunu vererek çıktıyı elde edebilrisiniz.


using ExcelToJson;

ExcelToJsonHelper.FromExcelToJson("LoadSheet.xlsx", "denemeLoadSheet");

Console.WriteLine("Tamamlandı");


// created by Mustafa Gönültaş.