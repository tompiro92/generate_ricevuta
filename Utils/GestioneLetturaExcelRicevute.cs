using Spire.Xls;
using System;
using System.Linq;


namespace Genera_Fatture.Utils
{
    public class GestioneLetturaExcelRicevute
    {

        private Workbook workbook;
        private Worksheet worksheet;

        private int indexColAmministratore;
        private int indexColCondominio;
        private int indexColCosto;
        private int indexColFattura;

        public GestioneLetturaExcelRicevute(String pathFile)
        {
            workbook = new Workbook();
            workbook.LoadFromFile(pathFile);
            //only on sheet 0 for now
            worksheet = workbook.Worksheets[0];
            indexColAmministratore = 1;
            indexColCondominio = 8;
            indexColCosto = 10;
            indexColFattura = 4;
            validazioneHeader();
        }


        private void validazioneHeader()
        {
            //check header of file.
            String headerAmministratore = retrieveAmministratore(1);
            String headerFattura = retrieveFattura(1);
            String headerCantiere = retrieveCondominio(1);
            String headerCosto = retrieveCosto(1);

            if (!headerAmministratore.Equals("AMMINISTRATORE")
                || (!headerFattura.Equals("FATT.") && !headerFattura.Equals("FATTURA"))
                || (!headerCantiere.Equals("CANTIERE") && !headerCantiere.Equals("CONDOMINIO"))
                || (!headerCosto.Equals("COSTO") && !headerCosto.Equals("€")))
            {
                throw new Exception("Il file clienti attivi selezionato non è valido. L'intestazione dovrebbe contenere:\n" +
                    "Colonna A: AMMINISTRATORE\n" +
                    "Colonna D: FATTURA O FATT.\n" +
                    "Colonna H: CANTIERE O CONDOMINIO\n" +
                    "Colonna J: € O COSTO");
            }
        }

        public String retrieveAmministratore(int row)
        {
            if (row < worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColAmministratore] != null ? worksheet[row, indexColAmministratore].Value.ToUpper() : "";
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public String retrieveCondominio(int row)
        {
            if (row < worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColCondominio] != null ? worksheet[row, indexColCondominio].Value.ToUpper() : "";
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public String retrieveCosto(int row)
        {
            if (row < worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColCosto] != null ? worksheet[row, indexColCosto].Value.ToUpper() : "";
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public String retrieveFattura(int row)
        {
            if (row < worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColFattura] != null ? worksheet[row, indexColFattura].Value.ToUpper() : "";
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public int getRowCount()
        {
            return worksheet.Rows.Count();
        }
    }
}
