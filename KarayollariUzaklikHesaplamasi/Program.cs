using Microsoft.Office.Interop.Excel;
using System.Collections;

int[][] DosyaOkumaArrayOlusturma(string filePath)
{
    //bu fonkisyonda excel dosyasından okuma yaparak jaggedArray yapısında şehirler
    //arasındaki mesafe değerlerinin tutulduğu bir yapı oluşturulması hedeflenmktedir.
    
    int[][] jaggedArray = new int[82][];
    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
    Workbook wb;
    Worksheet ws;

    wb = excel.Workbooks.Open(filePath);
    ws = wb.Worksheets[1];
    
    //bu for döngülerinde il kodu index degeri olacak sekilde excelden okunan degerler jaggedArraye ataniyor.
    for(int satirIndex = 1; satirIndex <= 81; satirIndex++)
    {
        jaggedArray[satirIndex] = new int[82];

        for (int sutunIndex = 1; sutunIndex <=81;sutunIndex++)
        {
            if (ws.Cells[satirIndex + 2, sutunIndex + 2].Value2 == null) {
                continue;
            }
            else
            {
                int mesafe = (int)ws.Cells[satirIndex + 2, sutunIndex + 2].Value2;
                jaggedArray[satirIndex][sutunIndex] = mesafe;
            }
        }
    }
    return jaggedArray;

}

string[] ilAdlariniOkumaArrayeAtama(string filePath) 
{
    string[] ilAdlari = new string[82];

    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
    Workbook wb;
    Worksheet ws;

    wb = excel.Workbooks.Open(filePath);
    ws = wb.Worksheets[1];

    int sutunIndex = 2;

    for (int satirIndex = 1; satirIndex <= 81; satirIndex++) 
    {
        string ilAdi = ws.Cells[satirIndex + 2, sutunIndex].Value2;
        ilAdlari[satirIndex] = ilAdi;
    }
    return ilAdlari;
}

void fonksiyon2(int[][] jaggedArray, string[] ilAdlari) {
    // Numarası verilen bir ilden verilen bir mesafeye kadar olan illerin sayısı, adları ve mesafelerini bulup yazdıran metot.
    
    ArrayList sonucIller = new ArrayList();

    Console.WriteLine("İl kodu giriniz:");
    string ilKoduInput = Console.ReadLine();
    int ilKodu = int.Parse(ilKoduInput);

    Console.WriteLine("Mesafe giriniz:");
    string mesafeInput = Console.ReadLine();
    int mesafe = int.Parse(mesafeInput);

    int[] mesafeArray = jaggedArray[ilKodu];
    for (int i = 1;i<=81;i++) 
    {
        //i plaka kodunu temsil ediyor.
        if (mesafeArray[i] > 0 && mesafeArray[i] <= mesafe) 
        {
            Console.WriteLine("İl Adı: " + ilAdlari[i] + " Mesafe: " + mesafeArray[i]);

        }
    }
}

void fonksiyon3(int[][]jaggedArray, string[] ilAdlari) {
    //Tüm Türkiye’deki il çiftleri arasındaki mesafelere bakılarak en küçüğünü ve en büyüğünü bulduran metot.
    int minMesafe = 100000;
    int minMesafeIl1 =0;
    int minMesafeIl2=0;

    int maxMesafe = 0;
    int maxMesafeIl1=0;
    int maxMesafeIl2 = 0;

    for (int satirIndex = 1; satirIndex <= 81; satirIndex++) 
    {
        for (int sutunIndex = 1; sutunIndex <= 81; sutunIndex++) 
        {
            if (jaggedArray[satirIndex][sutunIndex] < minMesafe && jaggedArray[satirIndex][sutunIndex] != 0) 
            {
                minMesafe = jaggedArray[satirIndex][sutunIndex];
                minMesafeIl1 = satirIndex;
                minMesafeIl2 = sutunIndex;
            }
            if (jaggedArray[satirIndex][sutunIndex] > maxMesafe && jaggedArray[satirIndex][sutunIndex] != 0) 
            {
                maxMesafe = jaggedArray[satirIndex][sutunIndex];
                maxMesafeIl1 = satirIndex;
                maxMesafeIl2 = sutunIndex;
            }
        }
    }

    Console.WriteLine("Aralarında en fazla mesafe bulunan iller: " + ilAdlari[maxMesafeIl1]+" " + ilAdlari[maxMesafeIl2]+" Mesafe: "+maxMesafe);
    Console.WriteLine("Aralarında en az mesafe bulunan iller: " + ilAdlari[minMesafeIl1] + " " + ilAdlari[minMesafeIl2] + " Mesafe: " + minMesafe);
}

void fonksiyon4(int[][] jaggedArray, string[] ilAdlari)
{
    //bu fonksiyonun amacı belirtilen baslangıc ilinden baslayarak verilen mesafeyi aşmadan en fazla kaç şehir gezilebileceğini hesaplamaktır.
    Console.Write("İl kodu giriniz:");
    string ilKoduInput = Console.ReadLine();
    int ilKodu = int.Parse(ilKoduInput);

    Console.Write("Mesafe giriniz:");
    string mesafeInput = Console.ReadLine();
    int maxMesafe = int.Parse(mesafeInput);

    int toplamMesafe = 0;
    ArrayList gezilenIlKodlari = new ArrayList();
    gezilenIlKodlari.Add(ilKodu);
    ArrayList gezilenMesafeler = new ArrayList();

    while (toplamMesafe < maxMesafe) 
    {
        int[] ilMesafeleri = jaggedArray[ilKodu];
        int minMesafe = 100000;
        int tempIlKodu = 0;
        for (int index = 1; index <= 81; index++) 
        {   
            bool gezildiMi = gezilenIlKodlari.Contains(index);
            if (index == ilKodu||gezildiMi) {
                continue;
            }
            else
            {
                if (ilMesafeleri[index] < minMesafe) {
                    if (toplamMesafe + ilMesafeleri[index] > maxMesafe) 
                    { continue; }
                    minMesafe = ilMesafeleri[index];
                    tempIlKodu = index;
                }
            }
        }
        if (minMesafe == 100000) {
            //liste gezilmesine rağmen gidilebilecek il kalmamış demektir.
            break;
        }
        ilKodu = tempIlKodu;
        gezilenIlKodlari.Add(ilKodu);
        toplamMesafe += minMesafe;
        gezilenMesafeler.Add(minMesafe);
    }

    ArrayList gezilenIller = new ArrayList();
    foreach (int gezilenIlKodu in gezilenIlKodlari) {
        gezilenIller.Add(ilAdlari[gezilenIlKodu]);
    }

    Console.Write("Gezilebilecek Toplam İl Sayısı: "+(gezilenIller.Count-1)+"\nGezilecek İller: ");
    for(int index = 0; index < gezilenIller.Count-1; index++)
    {   
        Console.Write(gezilenIller[index]+" ");
        if(index != gezilenIller.Count - 2)
        {
            Console.Write(gezilenMesafeler[index] + " ");
        }
        
    }
    Console.WriteLine("Toplam Mesafe: " + toplamMesafe+" km");
}

void fonksiyon5(int[][] jaggedArray, string[] ilAdlari) 
{   
    Random random = new Random();
    ArrayList ilKodlari = new ArrayList();
    ArrayList seciliIller = new ArrayList();
    int[][] mesafeMatrisi = new int[5][];

    for (int i = 0; i < 5; i++) 
    { 
        int ilKodu = random.Next(1,82);
        ilKodlari.Add(ilKodu);
    }
    ilKodlari.Sort();

    foreach (int ilKodu in ilKodlari) 
    {
        seciliIller.Add(ilAdlari[ilKodu]);
    }

    int satirIndex = 0;
    foreach (int ilKodu1 in ilKodlari)
    {
        int sutunIndex = 0;
        int[] mesafeArray = new int[5];
        foreach (int ilKodu2 in ilKodlari)
        {
            mesafeArray[sutunIndex] = jaggedArray[ilKodu1][ilKodu2];
            sutunIndex++;
        }
        mesafeMatrisi[satirIndex] = mesafeArray;
        satirIndex++;
    }

    Console.Write("\t");
    foreach (string ilAdi in seciliIller) 
    { 
        Console.Write(ilAdi+" ");
    }
    Console.WriteLine();
    for (int i = 0; i < 5; i++) 
    {
        Console.Write(seciliIller[i]+"  ");
        for (int j = 0; j < 5; j++) 
        {
            Console.Write(mesafeMatrisi[i][j]+"\t");
        }
        Console.WriteLine(" ");
    }
}

Console.Write("Mesafe verilerinin çekileceği excel dosyasının konumunu giriniz: ");
string filepath = Console.ReadLine();

Console.WriteLine("Excelden veriler çekiliyor lütfen bekleyiniz...");
int[][] jaggedArray = DosyaOkumaArrayOlusturma(filepath);
string[] ilAdlari = ilAdlariniOkumaArrayeAtama(filepath);

string devam = "e";

while(devam == "e" || devam == "E")
{
    Console.WriteLine("1 - verilen ilden belli bir uzaklığa kadar olan illerin listelenmesi.\n" +
        "2 - Türkiye'deki birbirine en yakın ve en uzak olan iki ilin belirlenmesi.\n" +
        "3 - Verilen ilden verilen mesafe kullanılarak en fazla ne kadar ilin gezilebildiğinin bulunması:\n" +
        "4 - rastgele seçilen 5 ilin birbirlerine olan uzaklıklarının mateisinin yazdırılması.");
    Console.Write("Seçiminiz: ");
    string input = Console.ReadLine();
    int secim = Convert.ToInt32(input);

    if(secim == 1)
    {
        fonksiyon2(jaggedArray, ilAdlari);
    }else if(secim == 2)
    {
        fonksiyon3(jaggedArray, ilAdlari);
    }else if (secim == 3)
    {
        fonksiyon4(jaggedArray, ilAdlari);
    }else if (secim == 4)
    {
        fonksiyon5(jaggedArray, ilAdlari);
    }

    Console.Write("Devam mı?(e/E)");
    devam = Console.ReadLine();
}