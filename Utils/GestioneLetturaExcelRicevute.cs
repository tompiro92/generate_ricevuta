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
        private int indexColSospesi;
        private int indexColCostoPulizie;
        private int indexColCostoBidoni;
        private int indexColCostoGiardini;
        private int indexColCostoPorta;

        private SingletonFileInizializzazione singletonFileInizializzazione;
        public GestioneLetturaExcelRicevute(String pathFile)
        {
            singletonFileInizializzazione = SingletonFileInizializzazione.getIstance();

            workbook = new Workbook();
            workbook.LoadFromFile(pathFile);
            //only on sheet 0 for now
            worksheet = workbook.Worksheets[0];
            indexColAmministratore = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_AMMINISTRATORE));
            indexColCondominio = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_CONDOMINIO));
            indexColCosto = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO));
            indexColFattura = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_FATTURA));
            indexColSospesi = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_SOSPESI));
            indexColCostoPulizie = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_PULIZIE));
            indexColCostoBidoni = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_BIDONI));
            indexColCostoGiardini = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_GIARDINI));
            indexColCostoPorta = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_PORTA_A_PORTA));
            validazioneHeader();
        }


        private void validazioneHeader()
        {
            //check header of file.
            String headerAmministratore = retrieveAmministratore(1);
            String headerFattura = retrieveFattura(1);
            String headerCantiere = retrieveCondominio(1);
            String headerCosto = retrieveCosto(1);
            String headerSospesi = retrieveSospesi(1);

            if (!headerAmministratore.Equals("AMMINISTRATORE")
                || (!headerFattura.Equals("FATT.") && !headerFattura.Equals("FATTURA"))
                || (!headerCantiere.Equals("CANTIERE") && !headerCantiere.Equals("CONDOMINIO"))
                || (!headerCosto.Equals("COSTO") && !headerCosto.Equals("€"))
                || (!headerSospesi.Equals("SOSPESI")))
            {
                throw new Exception("Il file clienti attivi selezionato non è valido. L'intestazione dovrebbe contenere:\n" +
                    $"Colonna {indexColAmministratore}: AMMINISTRATORE\n" +
                    $"Colonna {indexColFattura}: FATTURA O FATT.\n" +
                    $"Colonna {indexColSospesi}: SOSPESI\n" +
                    $"Colonna {indexColCondominio}: CANTIERE O CONDOMINIO\n" +
                    $"Colonna {indexColCosto}: € O COSTO");
            }
        }

        public String retrieveAmministratore(int row)
        {
            if (row <= worksheet.Rows.Count() && row > 0)
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
            if (row <= worksheet.Rows.Count() && row > 0)
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
            if (row <= worksheet.Rows.Count() && row > 0)
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
            if (row <= worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColFattura] != null ? worksheet[row, indexColFattura].Value.ToUpper() : "";
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public String retrieveSospesi(int row)
        {
            if (row <= worksheet.Rows.Count() && row > 0)
            {
                return worksheet[row, indexColSospesi] != null ? worksheet[row, indexColSospesi].Value.Trim().ToUpper() : "";
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public Double retrieveCostoPulizie(int row)
        {
            String costoPulizie;

            if (row <= worksheet.Rows.Count() && row > 0)
            {
                costoPulizie = worksheet[row, indexColCostoPulizie] != null ? worksheet[row, indexColCostoPulizie].Value.Trim().ToUpper() : "";
                try
                {
                    if (costoPulizie.Equals(""))
                    {
                        return 0;
                    }
                   return Double.Parse(costoPulizie);
                }
                catch (Exception e)
                {
                    return 0;
                }
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }


        public Double retrieveCostoBidoni(int row)
        {
            String costoBidoni;

            if (row <= worksheet.Rows.Count() && row > 0)
            {
                costoBidoni = worksheet[row, indexColCostoBidoni] != null ? worksheet[row, indexColCostoBidoni].Value.Trim().ToUpper() : "";
                try
                {
                    if (costoBidoni.Equals(""))
                    {
                        return 0;
                    }
                    return Double.Parse(costoBidoni);
                }
                catch (Exception e)
                {
                    return 0;
                }
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public Double retrieveCostoGiardini(int row)
        {
            String costoGiardini;

            if (row <= worksheet.Rows.Count() && row > 0)
            {
                costoGiardini = worksheet[row, indexColCostoGiardini] != null ? worksheet[row, indexColCostoGiardini].Value.Trim().ToUpper() : "";
                try
                {
                    if (costoGiardini.Equals(""))
                    {
                        return 0;
                    }
                    return Double.Parse(costoGiardini);
                }
                catch (Exception e)
                {
                    return 0;
                }
            }
            else
            {
                throw new Exception("La riga non esiste: riga non compresa tra 1 e " + worksheet.Rows.Count());
            }
        }

        public Double retrieveCostoPorta(int row)
        {
            String costoPortaAporta;

            if (row <= worksheet.Rows.Count() && row > 0)
            {
                costoPortaAporta = worksheet[row, indexColCostoPorta] != null ? worksheet[row, indexColCostoPorta].Value.Trim().ToUpper() : "";
                try
                {
                    if (costoPortaAporta.Equals(""))
                    {
                        return 0;
                    }
                    return Double.Parse(costoPortaAporta);
                }
                catch (Exception e)
                {
                    return 0;
                }
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
