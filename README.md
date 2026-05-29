# Metasezgisel Yöntemler Dersi Final Ödevi

## Proje Hakkında

Bu repo, Endüstri Mühendisliği Yüksek Lisans programı kapsamında alınan Metasezgisel Yöntemler dersi için hazırlanmıştır.

Çalışma, **El-Ghazali Talbi tarafından yazılan "Metaheuristics: From Design to Implementation"** kitabında yer alan metasezgisel optimizasyon yöntemleri ve uygulama örneklerinden yararlanılarak gerçekleştirilmiştir.

Bu projede Kapasiteli Araç Rotalama Problemi (CVRP), Karınca Kolonisi Optimizasyonu (Ant Colony Optimization - ACO) algoritması kullanılarak Excel VBA ortamında modellenmiş ve çözülmüştür.
Çalışmanın amacı, müşteri talepleri ve araç kapasite kısıtları dikkate alınarak toplam rota mesafesini minimize eden uygun dağıtım rotalarının oluşturulmasıdır.

---

## Dosyalar

### CVRP_ACO_zuleyhakoc.xlsx

Bu dosya proje uygulamasını içeren Excel çalışma kitabıdır.İçerisinde kullanıcıdan veri alan giriş ekranı, algoritmayı çalıştırmak için kullanılan butonlar 
ve sonuçların görüntülendiği alanlar bulunmaktadır. Butona tıklandığında VBA makrosu çalışarak CVRP problemi için ACO algoritmasını uygular ve 
elde edilen araç rotalarını toplam mesafe değeriyle birlikte kullanıcıya sunar.

Excel Sayfaları:

- Problem verilerinin girildiği sayfa
- Müşteri koordinatları ve talepleri
- Araç kapasitesi bilgileri
- ACO algoritma parametreleri
- Oluşturulan araç rotaları
- Toplam mesafe ve çözüm sonuçları

### CVRP_ACO_VBA_Modul.bas

Karınca Kolonisi Optimizasyonu algoritmasının VBA kaynak kodlarını içeren modül dosyasıdır.

Modül içerisinde:
Müşteri bilgileri  müşteri numarası şeklinde , X-Y koordinatları(sütunlarda bulunur) ve talep miktarı(sütunlarda bulunur) excel üzerinden okunur. 
Her satır bir müşteriyi temsil etmektedir.Müşteri koordinatları okunduktan sonra her müşteri noktası arasındaki uzaklık hesaplanır.Mesafe matrisi boyutu nxn'liktir. Buradaki kordinatlar ve talep miktarları müşteri(satır)/depo noktalarını(sütun) temsil etmektedir.
**Araç kapasitesi:** Excel’de kullanıcı tarafından girilen araç kapasitesi değeri okunarak(Aracımızın rota boyunca taşıyabileceği en fazla toplam talep miktarı)Rota oluşturulurken müşterilerin talep miktarları sırayla toplam yüke eklenir.
Eğer yeni müşteri eklendiğinde toplam yük araç kapasitesini aşmıyorsa müşteri mevcut rotaya dahil ediliyor. Kapasite aşılırsa mevcut rota tamamlanır ve yeni araç rotası başlatılır.Kısaca; Rota oluşturulurken müşterilerin talepleri sırayla toplanır.
Ancak toplam talep araç kapasitesini aşarsa o müşteri mevcut rotaya eklenmez, yeni araç rotası başlatılır.
Toplam yük, rotaya dahil edilen müşterilerin talep miktarlarının toplamını; araç kapasitesi ise aracın taşıyabileceği maksimum yük miktarını ifade eder.
Rota oluşturulurken toplam yükün araç kapasitesini aşmaması hedeflenir. Toplam yük kapasite sınırını geçtiğinde mevcut rota sonlandırılır ve yeni bir araç rotası oluşturulur.
**Rotanın içeriği ise**  depodan hangi müşteriye gideceği kısmı mesafe matrisyle alakalıdır. Mesafe matrisi oluşturulduktan sonra feromon matrisi(**Önceki çözümlerden elde edilen yol tercih bilgilerinin tutulduğu matris**) başlatılır. 
Rota oluşturma aşamasında algoritma, mesafe matrisi ve feromon matrisi bilgilerini birlikte kullanarak bir sonraki ziyaret edilecek müşteri belirlenir. 
Tüm müşteriler ziyaret edildiğinde algoritma rota oluşturmayı bırakır.Tüm müşteriler ziyaret edilerek rotalar tamamlandıktan sonra her rotanın toplam mesafesi, mesafe matrisindeki uzaklık değerleri kullanılarak hesaplanır.
Daha sonra tüm rotaların mesafeleri toplanarak çözümün toplam mesafesi elde edilir.İlk iterasyonda elde edilen toplam mesafe başlangıç çözümü olarak saklanır.
Sonraki iterasyonlarda elde edilen toplam mesafeler bu değerle karşılaştırılır ve daha düşük mesafeye sahip çözüm bulunursa en iyi çözüm güncellenir.En iyi çözüm saklandıktan sonra feromon matrisi güncellenir.
Feromon matrisi, yolların tercih edilme seviyelerini tutan yapıydı. Daha kısa toplam mesafeye sahip rotalarda kullanılan müşteri geçişlerinin feromon değerleri artırılır. 
Kullanılmayan veya daha az başarılı olan yolların feromon değerleri ise buharlaşma oranı(**feromon matrisi üzerindeki değerler belirli bir katsayı ile çarpılarak azaltılır.**) ile azaltılır.
Böylece eski rotaların etkisi zamanla azalırken, yeni bulunan başarılı rotaların daha fazla öne çıkması sağlanır.
Algoritma tekrar rota oluşturmaya başlıyor. Her iterasyonda tüm müşteriler için gerekli rotalar baştan oluşturuluyor.
Rota oluşturma sırasında müşteri talepleri tekrar değerlendirilir, kapasite kontrolleri yeniden yapılır ve tüm müşteriler ziyaret edilene kadar yeni rotalar oluşturulmaya devam edilir.
Oluşturulan tüm rotaların toplam mesafesi hesaplandıktan sonra çözüm değerlendirilir ve feromon matrisi güncellenir.
Sonraki iterasyonlarda ise aynı işlemler güncellenmiş feromon değerleri kullanılarak tekrar gerçekleştirilir.
Belirlenen iterasyon sayısına ulaşıldığında algoritma sonlandırılır. Her iterasyonda oluşturulan rota setlerinin toplam mesafeleri karşılaştırılır. 
**En düşük toplam mesafeye sahip çözüm en iyi çözüm olarak saklanır. Belirlenen iterasyon sayısına ulaşıldığında saklanan en iyi çözüm ve bu çözüme ait rotalar Excel sayfasına yazdırılarak kullanıcıya çıktı olarak veriliyor.**





## Kullanılan Yöntem

Karınca Kolonisi Optimizasyonu (Ant Colony Optimization - ACO), karıncaların yiyecek ararken oluşturdukları feromon izlerinden esinlenerek geliştirilmiş bir metasezgisel optimizasyon yöntemidir.

Bu çalışmada ACO algoritması;

* Araç kapasite kısıtlarını,
* Müşteri taleplerini,
* Depo başlangıç noktasını,
* Müşteriler arasındaki mesafeleri

dikkate alarak uygun araç rotaları üretmektedir.



## Projenin Amacı

Projenin temel amaçları şunlardır:

* Toplam taşıma mesafesini azaltmak
* Araç kapasite kısıtlarını sağlamak
* Tüm müşteri taleplerini karşılamak
* Etkin ve uygulanabilir dağıtım rotaları oluşturmak

---

## Kullanım

1. Excel dosyasını açın.
2. Problem verilerini ilgili alanlara girin.
3. VBA makrolarını etkinleştirin.
4. Algoritmayı çalıştırın.


---

## Ders Bilgisi

Bu çalışma, Metasezgisel Yöntemler dersi kapsamında hazırlanan final ödevi olarak geliştirilmiştir. Projede Excel VBA kullanılarak ACO tabanlı bir CVRP çözüm modeli oluşturulmuştur.
