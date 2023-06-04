using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public class HanteringAvExcel
{
    public void ImporteraDataFranExcel(string filSokvag)
    {
        try
        {
            // Skapar en filström för att läsa Excel-filen
            using (FileStream filstrom = new FileStream(filSokvag, FileMode.Open, FileAccess.Read))
            {
                // Skapar en arbetsbok från filströmmen
                IWorkbook arbetsbok = new XSSFWorkbook(filstrom);

                // Hämtar det första arket i arbetsboken
                ISheet ark = arbetsbok.GetSheetAt(0);

                // Loopar igenom varje rad i arket
                for (int radIndex = 0; radIndex <= ark.LastRowNum; radIndex++)
                {
                    IRow rad = ark.GetRow(radIndex);
                    if (rad != null)
                    {
                        // Hämtar värdet i varje cell för varje rad
                        for (int cellIndex = 0; cellIndex < rad.LastCellNum; cellIndex++)
                        {
                            ICell cell = rad.GetCell(cellIndex);
                            if (cell != null)
                            {
                                string cellvärde = cell.ToString();
                                Console.WriteLine("Importerat värde: " + cellvärde);
                            }
                        }
                    }
                }
            }

            Console.WriteLine("Data importerades från Excel-filen.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ett fel inträffade vid importen av data från Excel: " + ex.Message);
        }
    }

    public void ExporteraDataTillExcel(string filSokvag)
    {
        try
        {
            // Skapar en arbetsbok
            IWorkbook arbetsbok = new XSSFWorkbook();

            // Skapra ett ark i arbetsboken
            ISheet ark = arbetsbok.CreateSheet("DataSheet");

            // Skapar data att exportera
            string[,] data = new string[,]
            {
                {"Namn", "Yrke", "Stad"},
                {"Berk", "Fotbollsspelare", "Stockholm"},
                {"Olle", "Basketbollspelare", "Stockholm"},
                {"Leo", "Dansare", "Stockholm"}
            };

            for (int radIndex = 0; radIndex < data.GetLength(0); radIndex++)
            {
                IRow rad = ark.CreateRow(radIndex);
                for (int cellIndex = 0; cellIndex < data.GetLength(1); cellIndex++)
                {
                    ICell cell = rad.CreateCell(cellIndex);
                    cell.SetCellValue(data[radIndex, cellIndex]);
                }
            }

            using (FileStream filstrom = new FileStream(filSokvag, FileMode.Create, FileAccess.Write))
            {   // Lägger till "true" för att lämna filen öppen efter skrivning
                arbetsbok.Write(filstrom, true); 
            }

            Console.WriteLine("Data exporterades till Excel-filen.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ett fel inträffade vid exporten av data till Excel: " + ex.Message);
        }
    }

}
public class Program
{
    public static void Main()
    {
        // Anger sökvägen till din Excel-fil
        string filSökvag = "Namnlöst kalkylark.xlsx";

        // Skapar en instans av ExcelHanteraren
        HanteringAvExcel excelHanteraren = new HanteringAvExcel();

        // Importerar data från Excel
        excelHanteraren.ImporteraDataFranExcel(filSökvag);

        // Exporterar data till Excel
        excelHanteraren.ExporteraDataTillExcel(filSökvag);
    }
}

