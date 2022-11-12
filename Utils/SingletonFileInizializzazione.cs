using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Genera_Fatture.Utils
{
    public class SingletonFileInizializzazione
    {
        private static SingletonFileInizializzazione instance;
        private static readonly object padlock = new object();
        private Workbook workbook = null;
        private String path= "./Data/Inizializzazione.xlsx";
        private SingletonFileInizializzazione()

        {
            workbook = new Workbook();
            workbook.LoadFromFile(path);
        }

        public static SingletonFileInizializzazione getIstance()
        {        
            lock (padlock)
            {
                if (instance == null)
                {
                    instance = new SingletonFileInizializzazione();
                }
                return instance;
            }       
        }

        public void SaveFile()
        {
            lock (padlock)
            {
                workbook.Save();
            }
        }

        public String getNumeroUltimaFattura()
        {
            lock (padlock)
            {
                return workbook.Worksheets[0][1, 2] != null ? workbook.Worksheets[0][1, 2].Value.ToString() : "";
            }
        }
      
        public void setNumeroUltimaFattura(string numeroUltimaFattura)
        {
            lock (padlock)
            {
                workbook.Worksheets[0][1, 2].Value = numeroUltimaFattura.ToString();
                this.SaveFile();
            }
        }

        public String getIndexOf(ValueInizializzazioneEnum valueEnum)
        {
            lock (padlock)
            {
                if (valueEnum.Equals(ValueInizializzazioneEnum.CLIENTI_ATTIVI_AMMINISTRATORE))
                {
                    return workbook.Worksheets[0][4, 2] != null ? workbook.Worksheets[0][4, 2].Value.Trim().ToString() : "";
                }
                else if (valueEnum.Equals(ValueInizializzazioneEnum.CLIENTI_ATTIVI_FATTURA))
                {
                    return workbook.Worksheets[0][5, 2] != null ? workbook.Worksheets[0][5, 2].Value.Trim().ToString() : "";
                }
                else if (valueEnum.Equals(ValueInizializzazioneEnum.CLIENTI_ATTIVI_SOSPESI))
                {
                    return workbook.Worksheets[0][6, 2] != null ? workbook.Worksheets[0][6, 2].Value.Trim().ToString() : "";
                }
                else if (valueEnum.Equals(ValueInizializzazioneEnum.CLIENTI_ATTIVI_CONDOMINIO))
                {
                    return workbook.Worksheets[0][7, 2] != null ? workbook.Worksheets[0][7, 2].Value.Trim().ToString() : "";
                }
                else if (valueEnum.Equals(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO))
                {
                    return workbook.Worksheets[0][8, 2] != null ? workbook.Worksheets[0][8, 2].Value.Trim().ToString() : "";
                }
                else if (valueEnum.Equals(ValueInizializzazioneEnum.ANAGRAFICA_CONDOMINIO))
                {
                    return workbook.Worksheets[0][11, 2] != null ? workbook.Worksheets[0][11, 2].Value.Trim().ToString() : "";
                }
                else if (valueEnum.Equals(ValueInizializzazioneEnum.ANAGRAFICA_INDIRIZZO))
                {
                    return workbook.Worksheets[0][12, 2] != null ? workbook.Worksheets[0][12, 2].Value.Trim().ToString() : "";
                }
                else if (valueEnum.Equals(ValueInizializzazioneEnum.ANAGRAFICA_CAP))
                {
                    return workbook.Worksheets[0][13, 2] != null ? workbook.Worksheets[0][13, 2].Value.Trim().ToString() : "";
                }
                else if (valueEnum.Equals(ValueInizializzazioneEnum.ANAGRAFICA_COMUNE))
                {
                    return workbook.Worksheets[0][14, 2] != null ? workbook.Worksheets[0][14, 2].Value.Trim().ToString() : "";
                }
                else if (valueEnum.Equals(ValueInizializzazioneEnum.ANAGRAFICA_PROVINCIA))
                {
                    return workbook.Worksheets[0][15, 2] != null ? workbook.Worksheets[0][15, 2].Value.Trim().ToString() : "";
                }
                else {
                    throw new Exception("VALORE NON VALIDO");
                }
            }
        }

        public void setIndexOf(ValueInizializzazioneEnum valueEnum, String value)
        {
            if (valueEnum.Equals(ValueInizializzazioneEnum.CLIENTI_ATTIVI_AMMINISTRATORE))
            {
                workbook.Worksheets[0][4,2].Value = value.Trim();
            }
            else if (valueEnum.Equals(ValueInizializzazioneEnum.CLIENTI_ATTIVI_FATTURA))
            {
                workbook.Worksheets[0][5, 2].Value = value.Trim();

            }
            else if (valueEnum.Equals(ValueInizializzazioneEnum.CLIENTI_ATTIVI_SOSPESI))
            {
                workbook.Worksheets[0][6, 2].Value = value.Trim();

            }
            else if (valueEnum.Equals(ValueInizializzazioneEnum.CLIENTI_ATTIVI_CONDOMINIO))
            {
                workbook.Worksheets[0][7, 2].Value = value.Trim();

            }
            else if (valueEnum.Equals(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO))
            {
                workbook.Worksheets[0][8, 2].Value = value.Trim();

            }
            else if (valueEnum.Equals(ValueInizializzazioneEnum.ANAGRAFICA_CONDOMINIO))
            {
                workbook.Worksheets[0][11, 2].Value = value.Trim();

            }
            else if (valueEnum.Equals(ValueInizializzazioneEnum.ANAGRAFICA_INDIRIZZO))
            {
                workbook.Worksheets[0][12, 2].Value = value.Trim();
            }
            else if (valueEnum.Equals(ValueInizializzazioneEnum.ANAGRAFICA_CAP))
            {
                workbook.Worksheets[0][13, 2].Value = value.Trim();
            }
            else if (valueEnum.Equals(ValueInizializzazioneEnum.ANAGRAFICA_COMUNE))
            {
                workbook.Worksheets[0][14, 2].Value = value.Trim();
            }
            else if (valueEnum.Equals(ValueInizializzazioneEnum.ANAGRAFICA_PROVINCIA))
            {
                workbook.Worksheets[0][15, 2].Value = value.Trim();
            }
            else
            {
                throw new Exception("VALORE NON VALIDO");
            }
        }


    }
}
