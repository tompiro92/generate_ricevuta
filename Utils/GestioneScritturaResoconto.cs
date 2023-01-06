using Spire.Xls;
using System;
using System.Drawing;
using System.IO;
using System.Linq;


namespace Genera_Fatture.Utils
{
    public class GestioneScritturaResoconto
    {
        private Workbook workbook;
        private Worksheet worksheet;
        private bool enabled = false;

        private int lastRow;
        private int COL_DATA_FATTURA = 1;
        private int COL_AMMINISTRATORE = 2;
        private int COL_CONDOMINIO = 3;
        private int COL_DESCRIZIONE = 4;
        private int COL_COSTO = 5;

        public GestioneScritturaResoconto(String pathFile)
        {

            try
            {
                this.workbook = new Workbook();
                this.workbook.LoadFromFile(pathFile);
                //only on sheet 0 for now
                this.worksheet = workbook.Worksheets[0];
                this.lastRow = calculateLastRow();
            }
            catch (Exception ex)
            {
                throw new Exception("Impossibile aprire l'excel resoconto. Error: " + ex.Message + ". \nStackTrace: "+ex.StackTrace);
            }

        }

        private int calculateLastRow()
        {
            int rowCount = worksheet.Rows.Count();
            int indexLastRow = 0;

            for (int i = 1; i <= rowCount; ++i)
            {
                if (!worksheet[i, COL_DATA_FATTURA].Value.Equals("") && !worksheet[i, COL_AMMINISTRATORE].Value.Equals(""))
                {
                    indexLastRow++;
                }
            }

            return indexLastRow;
        }

        public void SalvataggioFile(String pathFolder, String nomeFile)
        {

            Directory.CreateDirectory(pathFolder);
            workbook.SaveToFile(pathFolder + "\\" + nomeFile);
            workbook = null;
            worksheet = null;

        }

        public void writeInFile(String dataFattura, String amministratore, String condominio, String descrizione, Double costo)
        {
            if (worksheet != null)
            {
                lastRow += 1;
                worksheet[lastRow, COL_DATA_FATTURA].Value = dataFattura;
                worksheet[lastRow, COL_AMMINISTRATORE].Value = amministratore;
                worksheet[lastRow, COL_CONDOMINIO].Value = condominio;
                worksheet[lastRow, COL_DESCRIZIONE].Value = descrizione;
                worksheet[lastRow, COL_COSTO].NumberFormat = ("0.00 €");
                worksheet[lastRow, COL_COSTO].Value2 = costo; //0 se non è calcolabile
                if (costo <= 0)
                {
                    worksheet[lastRow, COL_COSTO].Style.Color = Color.Red;
                }
            }
        }
    }
}
