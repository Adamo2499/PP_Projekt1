using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
namespace PP_Projekt1 {
    class Program {
        static Double DlugoscTworzacej(Double promien, Double wysokosc) {
            Double tworzaca;
            if (promien <= 0 || wysokosc <= 0) {
                throw new ArgumentException("Promień podstawy oraz wysokość muszą być większe od 0");
            }
            else {
                tworzaca = Math.Sqrt(promien * promien + wysokosc * wysokosc);
                return tworzaca;
            }
            
        }
        static Double PoleStozka(Double promien, Double tworzaca) {
            Double pStozka;
            if (promien <= 0 || tworzaca <= 0) {
                throw new ArgumentException("Promień podstawy oraz tworząca stożka musi być wieksza od 0");
            }
            else {
                pStozka = (Math.PI * promien * promien) + (Math.PI * promien * tworzaca);
                return pStozka;
            }
        }
        static Double PoleSzescianu(Double a) {
            Double pSzescianu;
            if (a <= 0) {
                throw new ArgumentException("Długość boku sześcianu musi być większa od 0");
            }
            else {
                pSzescianu = 6 * a * a;
                return pSzescianu;
            }
        }
        public static void WygenerujSciezki(out String xls, out String xlsx) {
            String SciezkaDoDokumentow = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            xls = Path.Combine(SciezkaDoDokumentow, "Projekt_AB.xls");
            xlsx = Path.Combine(SciezkaDoDokumentow, "Projekt_AB.xlsx");
        }
        static void EksportujDoExcela(Double promienStozka, Double wysokoscStozka, Double dlugoscTworzacej, Double pStożka, Double bokSzescianu, Double pSześcianu, Double sumaPol) {            
            WygenerujSciezki(out String xls,out String xlsx);
            Excel.Application excelApp = new Excel.Application();
            Excel._Workbook excelWorkBook = excelApp.Workbooks.Add();
            Excel._Worksheet excelWorkSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            if (excelApp == null) {
                Console.WriteLine("Program Microsoft Excel nie jest zainstalowany na tym komputerze!");
                Console.WriteLine("Możesz dostać go tutaj: https://www.ceneo.pl/oferty/office-2019");
            }
            else {
                excelApp.Visible = true;
                #region stylizowanie arkusza
                Excel.Range formatRange1;
                formatRange1 = excelWorkSheet.get_Range("a1");
                formatRange1.EntireRow.Font.Bold = true;
                formatRange1.EntireRow.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                formatRange1.EntireRow.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                formatRange1.EntireRow.WrapText = true;
                Excel.Range formatRange2;
                formatRange2 = excelWorkSheet.get_Range("a2");
                formatRange2.EntireRow.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                formatRange2.EntireRow.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                formatRange2.EntireRow.WrapText = true;
                #endregion
                #region wartosci
                excelWorkSheet.Cells[1, "A"] = "Promień podstawy stożka";
                excelWorkSheet.Cells[1, "B"] = "Długość wysokości stożka";
                excelWorkSheet.Cells[1, "C"] = "Długość tworzącej stożka";
                excelWorkSheet.Cells[1, "D"] = "Pole stożka";
                excelWorkSheet.Cells[1, "E"] = "Długość boku sześcianu";
                excelWorkSheet.Cells[1, "F"] = "Pole sześcianu";
                excelWorkSheet.Cells[1, "G"] = "Suma pól";

                excelWorkSheet.Cells[2, "A"] = promienStozka;
                excelWorkSheet.Cells[2, "B"] = wysokoscStozka;
                excelWorkSheet.Cells[2, "C"] = dlugoscTworzacej;
                excelWorkSheet.Cells[2, "D"] = pStożka;
                excelWorkSheet.Cells[2, "E"] = bokSzescianu;
                excelWorkSheet.Cells[2, "F"] = pSześcianu;
                excelWorkSheet.Cells[2, "G"] = sumaPol;
                #endregion
                try {
                    if (!File.Exists(xls) || !File.Exists(xlsx)) {
                        excelWorkBook.SaveAs("Projekt_AB.xlsx");
                        excelWorkBook.SaveCopyAs("Projekt_AB.xls");
                    }
                    else {
                        excelWorkBook.Save();
                    }
                }
                catch (System.Runtime.InteropServices.COMException) {
                    Console.WriteLine("Brak uprawnień do zapisywania w tym pliku!");
                    Environment.Exit(0);
                    excelApp.Quit();
                }
                excelWorkBook.Close();
                excelApp.Quit();
            }
        }
        static void Main() {
            Console.Title = "Program do obliczania pól figur przestrzennych oraz eksportu danych do Excela";
            Double a, r, H, l, poleStozka, poleSzescianu, sumaPol;
        //pole stożka
        daneDoStozka:
            Console.WriteLine("Podaj promień podstawy stożka: ");
            try{
                r = Double.Parse(Console.ReadLine());
            }
            catch (System.FormatException) {
                Console.WriteLine("Podano niepoprawne dane!");
                goto daneDoStozka;
            }
            Console.WriteLine("Podaj wysokość stożka: ");
            try {
                H = Double.Parse(Console.ReadLine());
            }
            catch (System.FormatException) {
                Console.WriteLine("Podano niepoprawne dane!");
                goto daneDoStozka;
            }
        daneDoSzescianu:
            Console.WriteLine("Podaj długość boku sześcianu: ");
            try {
                a = Double.Parse(Console.ReadLine());
            }
            catch (System.FormatException) {
                Console.WriteLine("Podano niepoprawne dane!");
                goto daneDoSzescianu;
            }
            Console.WriteLine();

                // metody dotyczące pola stożka
            try {
                l = DlugoscTworzacej(r, H);
                Console.WriteLine($"Długość tworzącej wynosi: {l:n2}");
            }
            catch(ArgumentException exp) {
                Console.WriteLine(exp.Message);
                Console.WriteLine();
                goto daneDoStozka;
            }
            try {
                poleStozka = PoleStozka(r, l);
                Console.Write($"Pole stożka wynosi {poleStozka:n2} \n");
            }
            catch (ArgumentException exp) {
                Console.WriteLine(exp.Message);
                Console.WriteLine();
                goto daneDoStozka;
            }
            
                //metoda dotycząca pola sześcianu
            try {
                poleSzescianu = PoleSzescianu(a);
                Console.WriteLine($"Pole sześcianu wynosi: {poleSzescianu:n2}");
            }
            catch (ArgumentException exp) {
                Console.WriteLine(exp.Message);
                Console.WriteLine();
                goto daneDoSzescianu;
            }
            Console.WriteLine();

            // Suma pól
            sumaPol = poleStozka + poleSzescianu;
            Console.WriteLine($"Suma pól wynosi {sumaPol:n2}");

            // eksportowanie danych do Excela
            Console.WriteLine("Teraz dokonuję eksportu podanych danych do Excela...");
            EksportujDoExcela(r, H, l, poleStozka, a, poleSzescianu,sumaPol);
            Console.WriteLine("Eksport danych zakończył się sukcesem!");
            Console.WriteLine();
            Console.WriteLine("Wyeksportowane dane można znaleźć w folderze Dokumenty (pliki Projekt_AB.xls i Projekt_AB.xlsx; w przypadku ponownego wywołania programu dane będą zapisywane w pliku Zeszyt1.xlsx)");
            Console.WriteLine();
            Console.WriteLine("Wcisnij dowolny klawisz, aby zakończyć działanie programu!");
            Console.ReadKey(true);
        }
    }
}
