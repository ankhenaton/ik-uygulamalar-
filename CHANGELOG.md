# CHANGELOG

## 2026-10-08 — Operasyon ve Soru Odaklı Üretim
- Şubat Bordro Sihirbazı eklendi.
- Hafta Tatili + Rapor kontrolü eklendi.
- UBGT Ücret Kontrolü eklendi.
- İzin Süresi Hesaplama eklendi.
- 2026 Ücret Beyan Kontrolü eklendi.
- İş Kazası Bildirim Süresi ön kontrolü eklendi.
- 8 SEO rehber sayfası eklendi.
- Dinamik sol menü yeni modülleri mevcut sayfalara da ekleyecek şekilde güçlendirildi.
- Sitemap yeniden oluşturulacaktır.


## 2026-10-08 — Üretim QA ve Hata Düzeltmeleri
- 70 HTML sayfasında tarayıcı tabanlı duman testi yapıldı; sayfa yükleme ve etkileşimlerde JavaScript hatası kalmadı.
- `sgk95.js` ortak yardımcı fonksiyonlarının bazı sayfalarda ilk hesaplamadan önce kullanılmasını engelleyen `defer` kaynaklı zamanlama problemi giderildi.
- `out` yardımcı fonksiyonu ile HTML `id="out"` global adı çakışması giderildi.
- Bordro Denetimi içindeki `esc` isim çakışması `escapeHtml` olarak düzeltildi.
- Dosya tabanlı mutabakat araçlarında kırılgan global element referansları açık DOM referanslarına çevrildi.
- İK Takvimi localStorage hatasına karşı dayanıklı hale getirildi.
- Excel Veri Temizleme'de boş satır sayımı düzeltildi.
- 69 indexlenebilir HTML için canonical ↔ sitemap.xml ↔ sitemap.txt eşleşmesi doğrulandı.
- 8 Excel şablonu açılarak sayfa yapıları ve görünür formül-hata metinleri kontrol edildi.
