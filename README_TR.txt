# JEOGAZ Araç Muayene Takip — Windows EXE

Bu paket, Windows üzerinde Python kurmadan çalışacak tek dosyalık EXE üretmek için hazırlanmıştır.
Bu çalışma ortamında Windows EXE derleyicisi olmadığı için EXE doğrudan burada derlenmedi; GitHub Actions Windows sunucusunda oluşturulur.

## Program özellikleri
- Excel'deki `Sayfa1` sayfasını okur.
- Plaka: C sütunu, marka: D sütunu, muayene tarihi: E sütunu, T9/sigorta tarihi: H sütunu.
- Muayene için 20 gün, T9/sigorta için 30 gün kala uyarır.
- Süresi geçmiş ve bugün bitecek kayıtları da bildirir.
- Sistem tepsisinde çalışır; sağ tık menüsünden şimdi kontrol edebilir, Excel'i açabilir veya çıkabilirsiniz.
- İlk açılışta takip Excel dosyasını seçmenizi ister ve seçimi kaydeder.

## EXE'yi oluşturma (bir kez internet üzerinden)
1. GitHub.com'da hesabınız yoksa oluşturun.
2. Yeni bir **private repository** oluşturun (ör. `JEOGAZ-Takip`).
3. Bu klasördeki tüm dosyaları ve `.github/workflows/build.yml` dosyasını repository'ye yükleyin. `.github` klasörü ve içindeki `workflows/build.yml` yolu aynen korunmalı.
4. GitHub repository sayfasında **Actions** sekmesine girin.
5. `JEOGAZ Windows EXE oluştur` iş akışını seçip **Run workflow** düğmesine basın.
6. İşlem tamamlanınca çalışmanın altındaki **Artifacts** bölümünden `JEOGAZ_Takip_Windows` dosyasını indirin.
7. ZIP'i çıkarın. İçindeki `JEOGAZ_Takip.exe` dosyasını çalıştırın.
8. İlk açılışta `JEOGAZ MUAYENE TAKİP 2024.xlsx` dosyanızı seçin.

## Önemli
- Bu uygulama Windows 10/11 x64 için hedeflenmiştir.
- Bildirimlerin görünmesi için Windows bildirimleri/odaklanma ayarları izin vermelidir.
- Program kapalıysa bildirim veremez; otomatik başlatmak için EXE kısayolunu `Win + R` -> `shell:startup` ile açılan klasöre koyabilirsiniz.
- Excel dosyası taşınır veya yeniden adlandırılırsa programı ilk kez açarken kayıtlı dosya yolunu silmeniz gerekir: `%LOCALAPPDATA%\JEOGAZTakip\excelpath.txt`.
- Araç tarihleri için orijinal Excel'deki tarih hücreleri gerçek Excel tarih biçiminde olmalıdır.
