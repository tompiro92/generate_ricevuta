using Spire.Xls;
using System;
using System.Linq;

namespace Genera_Fatture.Utils
{
    public class GestioneLetturaExcelAnagrafica
    {
        private Workbook workbook;
        private Worksheet worksheet;

        private int indexColCondominio;
        private int indexColIndirizzo;
        private int indexColCap;
        private int indexColProvincia;
        private int indexColComune;
        private SingletonFileInizializzazione singletonFileInizializzazione;
        public GestioneLetturaExcelAnagrafica(String pathFile)
        {
            singletonFileInizializzazione = SingletonFileInizializzazione.getIstance();

            workbook = new Workbook();
            workbook.LoadFromFile(pathFile);
            //only on sheet 0 for now
            worksheet = workbook.Worksheets[0];
            indexColCondominio = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_CONDOMINIO));
            indexColIndirizzo = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_INDIRIZZO));
            indexColCap = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_CAP));
            indexColProvincia = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_PROVINCIA));
            indexColComune = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_COMUNE));
            validazioneHeader();
        }
        private void validazioneHeader()
        {
            if(worksheet!= null)
            {
                String headerCondominio = retrieveCondominio(1);
                String headerIndirizzo = retrieveIndirizzo(1);
                String headerCap = retrieveCap(1);
                String headerComune = retrieveComune(1);
                String headerProvincia = retrieveProvincia(1);

                if ((!headerCondominio.Equals("CONDOMINIO") && !headerCondominio.Equals("PARCO") && !headerCondominio.Equals("CONDOMINIO/PARCO"))
                    || !headerIndirizzo.Equals("INDIRIZZO")
                    || !headerCap.Equals("CAP")
                    || !headerComune.Equals("COMUNE")
                    || (!headerProvincia.Equals("PROV.") && !headerProvincia.Equals("PROVINCIA")))
                {
                    throw new Exception("Il file anagrafica clienti selezionato non è valido. L'intestazione dovrebbe contenere:\n" +
                        $"Colonna {indexColCondominio}: CONDOMINIO O CONDOMINIO/PARCO\n" +
                        $"Colonna {indexColIndirizzo}: INDIRIZZO\n" +
                        $"Colonna {indexColCap}: CAP\n" +
                        $"Colonna {indexColComune}: COMUNE\n" +
                        $"Colonna {indexColProvincia}: PROVINCIA o PROV.");
                }
            }
            else
            {
                throw new Exception("NullPointer: Workseeh null");
            }
           
        }

        public String retrieveCondominio(int row)
        {
            if(row <= worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColCondominio] != null ? worksheet[row, indexColCondominio].Value.ToUpper() : "";
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public String retrieveIndirizzo(int row)
        {
            if (row <= worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColIndirizzo] != null ? worksheet[row, indexColIndirizzo].Value.ToUpper() : "";
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public String retrieveCap(int row)
        {
            if (row <= worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColCap] != null ? worksheet[row, indexColCap].Value.ToUpper() : "";
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public String retrieveComune(int row)
        {
            if (row <= worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColComune] != null ? worksheet[row, indexColComune].Value.ToUpper() : "";
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public String retrieveProvincia(int row)
        {
            if (row <= worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColProvincia] != null ? worksheet[row, indexColProvincia].Value.ToUpper() : "";
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
