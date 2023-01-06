using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Genera_Fatture.Utils
{
    public class GestioneScritturaExcelTemplateFattura
    {
        private Workbook workbook;
        private Worksheet worksheet;

        private String NUMERO_FATTURA = "D9";
        private String DATA_FATTURA = "D10";
        private String MESE_FATTURA = "C18";
        private String NOME_CONDOMINIO_FATTURA = "F7";
        private String VIA_CONDOMINIO_FATTURA = "F10";
        private String COMUNE_CAP_PROVINCIA_CONDOMINIO_FATTURA = "F11";
        private String COSTO_FATTURA = "I16";
        private String ANNO_FATTURA = "E9";
        private String DESCRIZIONE_FATTURA = "C16";

        public GestioneScritturaExcelTemplateFattura(String pathFile)
        {
            workbook = new Workbook();
            workbook.LoadFromFile(pathFile);
            //only on sheet 0 for now
            worksheet = workbook.Worksheets[0];
        }

        public void SalvataggioFile(String pathFolder, String nomeFile)
        {
            Directory.CreateDirectory(pathFolder);
            worksheet.CalculateAllValue();
            workbook.SaveToFile(pathFolder + "\\" + nomeFile);
            workbook = null;
            worksheet = null;       
        }

        public void writeInFile(String text, TipologiaWriteTemplateExcel type)
        {
            if (type.Equals(TipologiaWriteTemplateExcel.NUMERO_FATTURA))
            {
                worksheet[NUMERO_FATTURA].Text = text;
            }
            else if (type.Equals(TipologiaWriteTemplateExcel.DATA_FATTURA))
            {
                worksheet[DATA_FATTURA].Text = text;
            }
            else if (type.Equals(TipologiaWriteTemplateExcel.COSTO_FATTURA))
            {
                if (!text.Equals(""))
                {
                    try
                    {
                        worksheet[COSTO_FATTURA].Value2 = Math.Round(float.Parse(text), 2);
                    }
                    catch (FormatException ex)
                    {
                        worksheet[COSTO_FATTURA].Value2 = "";
                    }
                }
            }
            else if(type.Equals(TipologiaWriteTemplateExcel.MESE_FATTURA))
            {
                worksheet[MESE_FATTURA].Text = text;
            }
            else if(type.Equals(TipologiaWriteTemplateExcel.NOME_CONDOMINIO_FATTURA))
            {
                worksheet[NOME_CONDOMINIO_FATTURA].Text = text;
            }
            else if (type.Equals(TipologiaWriteTemplateExcel.COMUNE_CAP_PROVINCIA_CONDOMINIO_FATTURA))
            {
                worksheet[COMUNE_CAP_PROVINCIA_CONDOMINIO_FATTURA].Text = text;
            }
            else if (type.Equals(TipologiaWriteTemplateExcel.VIA_CONDOMINIO_FATTURA))
            {
                worksheet[VIA_CONDOMINIO_FATTURA ].Text = text;
            }
            else if (type.Equals(TipologiaWriteTemplateExcel.ANNO_FATTURA))
            {
                worksheet[ANNO_FATTURA].Text = text;
            }
            else if (type.Equals(TipologiaWriteTemplateExcel.DESCRIZIONE))
            {
                worksheet[DESCRIZIONE_FATTURA].Text = text;
            }
        }  
    }
}
