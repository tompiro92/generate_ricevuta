
using Genera_Fatture.Utils;
using Spire.Xls;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Genera_Fatture
{
    public partial class form : Form
    {
        private String inputFilePathCosti="";
        private String inputFilePathAnagrafica="";
        public string InputFilePathCosti { get => inputFilePathCosti; set => inputFilePathCosti = value; }
        public string InputFilePathAnagrafica { get => inputFilePathAnagrafica; set => inputFilePathAnagrafica = value; }

        private Delegates delegates;
    

        private String basePathOutputFile =  $"C:\\Fatture\\";
        Workbook workbookUltimaFattura;
        private int progressivo;

        public form()
        {
            InitializeComponent();
            CustomInizializeComponent();
            
        }

        private void CustomInizializeComponent()
        {
            //Iniziazzazione delegates
            delegates = new Delegates();

            //Inizializzazione openFileDialog
            this.openFileDialog.Filter = "Excel Files |*.xlsx";
            this.openFileDialog.Multiselect = false;
            this.openFileDialog.FileName = "";
            this.openFileDialog.InitialDirectory = $"C:\\users\\{System.Environment.UserName}\\Desktop";

            //Inizializzazione Button
            this.buttonGeneraFatture.Enabled = false;
            this.buttonGeneraFatture.ButtonColor = Color.Gray;

            //Inizializzazione Data -> Default primo del mese
            DateTime now = DateTime.Now;
            this.dateTimePicker1.Value = new DateTime(now.Year, now.Month, 1);

            //inizializzazione ultima fattura.
            this.numericUpDownNumeroFattura.Minimum = 0;
            this.numericUpDownNumeroFattura.Maximum = 100000000;

            workbookUltimaFattura = new Workbook();
            workbookUltimaFattura.LoadFromFile("./Data/NumeroUltimaFattura.xlsx");
            Worksheet worksheetUltimaFattura = workbookUltimaFattura.Worksheets[0];
            int progressivo = 0 ;
            if (DateTime.Now.Month == 1)
            {
                progressivo = 0;
            }
            else
            {
                progressivo = int.Parse(worksheetUltimaFattura[1, 1].Value);
                if (progressivo < 0)
                {
                    progressivo = 0;
                }
            }
            this.numericUpDownNumeroFattura.Value = progressivo;
        }

        private void buttonFileCosti_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog.ShowDialog() == DialogResult.OK)
            {              
                this.textBoxFileCosti.Text = openFileDialog.FileName.ToString();
                this.InputFilePathCosti = this.textBoxFileCosti.Text;
               
                if (!inputFilePathAnagrafica.Equals("")) {
                    //UI
                    this.buttonGeneraFatture.Enabled = true;
                    this.buttonGeneraFatture.ButtonColor = Color.DodgerBlue;
                }
                //clear open dialog
                this.openFileDialog.FileName = "";
                this.openFileDialog.InitialDirectory = $"C:\\users\\{System.Environment.UserName}\\Desktop";
            }
            else
            {
                this.textBoxFileCosti.Text = "";
                this.InputFilePathCosti = "";
                this.buttonGeneraFatture.Enabled = false;
                this.buttonGeneraFatture.ButtonColor = Color.Gray;
                //clear open dialog
                this.openFileDialog.FileName = "";
                this.openFileDialog.InitialDirectory = $"C:\\users\\{System.Environment.UserName}\\Desktop";
                MessageBox.Show(this, "Nessun file selezionato", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void buttonFileAnagrafica_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxFileAnagrafica.Text = openFileDialog.FileName.ToString();
                this.inputFilePathAnagrafica = this.textBoxFileAnagrafica.Text;

                if (!inputFilePathCosti.Equals(""))
                {
                    //UI
                    this.buttonGeneraFatture.Enabled = true;
                    this.buttonGeneraFatture.ButtonColor = Color.DodgerBlue;
                }
                //clear open dialog
                this.openFileDialog.FileName = "";
                this.openFileDialog.InitialDirectory = $"C:\\users\\{System.Environment.UserName}\\Desktop";
            }
            else
            {
                this.textBoxFileAnagrafica.Text = "";
                this.inputFilePathAnagrafica = "";
                this.buttonGeneraFatture.Enabled = false;
                this.buttonGeneraFatture.ButtonColor = Color.Gray;
                //clear open dialog
                this.openFileDialog.FileName = "";
                this.openFileDialog.InitialDirectory = $"C:\\users\\{System.Environment.UserName}\\Desktop";
                MessageBox.Show(this, "Nessun file selezionato", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void buttonGeneraFatture_Click(object sender, EventArgs e)
        {

            delegates.disableEnableButtonDelegate(buttonGeneraFatture, false);
            delegates.disableEnableButtonDelegate(buttonFileCosti, false);
            delegates.disableEnableButtonDelegate(buttonFileAnagrafica, false);     

            Thread t = new Thread(
                () => generafatture(InputFilePathCosti));
            t.Start();
        }

        private void generafatture(string file)
        {
            progressivo = (int) numericUpDownNumeroFattura.Value;
            delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, null);
            Thread.Sleep(500);
            delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Generazione Fatture Iniziata");
            try
            {
                // Carica file Fatture Input
                GestioneLetturaExcelRicevute excelRicevute = new GestioneLetturaExcelRicevute(inputFilePathCosti);

                // Carica file anagrafiche clienti Input
                GestioneLetturaExcelAnagrafica excelAnagrafica = new GestioneLetturaExcelAnagrafica(inputFilePathAnagrafica);


                int rowsExcelCosti = excelRicevute.getRowCount();
                int rowsExcelAnagrafica = excelAnagrafica.getRowCount();

                if (rowsExcelCosti > 1 && rowsExcelAnagrafica > 1)
                {
                    bool validazioneGenerale = true;
                    String folderOutput = basePathOutputFile + dateTimePicker1.Value.ToString("yyyy") + "\\" + dateTimePicker1.Value.ToString("MMMM");

                    if (!Directory.Exists(folderOutput)) {

                        for (int i = 2; i <= rowsExcelCosti; i++)
                        {
                            Console.WriteLine("ROW COSTI: " + i);

                            // Carica file template
                            GestioneScritturaExcelTemplateFattura excelTemplateFattura = new GestioneScritturaExcelTemplateFattura("./Data/Template.xlsx");

                            //Lettura file input
                            String amministratore = excelRicevute.retrieveAmministratore(i);
                            String prezzo = excelRicevute.retrieveCosto(i);
                            String nomeCondominio = excelRicevute.retrieveCondominio(i);
                            String fattura = excelRicevute.retrieveFattura(i);
                            String data = dateTimePicker1.Value.ToString("dd/MM/yyyy");
                            String mese = "NEL MESE DI " + dateTimePicker1.Value.ToString("MMMM");

                            String indirizzo = "";
                            String provincia = "";
                            String comune = "";
                            String cap = "";

                            
                            if (!fattura.Trim().Equals("N") && !fattura.Trim().Equals("NO"))
                            {
                                for (int j = 2; j <= rowsExcelAnagrafica; ++j)
                                {
                                    String nomeCondominioAnagrafica = excelAnagrafica.retrieveCondominio(j);

                                    if (!nomeCondominioAnagrafica.Trim().Equals("") && nomeCondominioAnagrafica.Trim().Equals(nomeCondominio.Trim()))
                                    {
                                        indirizzo = excelAnagrafica.retrieveIndirizzo(j);
                                        cap = excelAnagrafica.retrieveCap(j);
                                        comune = excelAnagrafica.retrieveComune(j);
                                        provincia = excelAnagrafica.retrieveProvincia(j);
                                    }
                                }
                                ++progressivo;
                                bool valid = CheckFieldAndLog(i, progressivo, amministratore, prezzo, nomeCondominio, cap, provincia, indirizzo, comune);
                                excelTemplateFattura.writeInFile(progressivo.ToString(), TipologiaWriteTemplateExcel.NUMERO_FATTURA);
                                excelTemplateFattura.writeInFile(data.ToUpper(), TipologiaWriteTemplateExcel.DATA_FATTURA);
                                excelTemplateFattura.writeInFile(prezzo, TipologiaWriteTemplateExcel.COSTO_FATTURA);
                                excelTemplateFattura.writeInFile(mese.ToUpper(), TipologiaWriteTemplateExcel.MESE_FATTURA);
                                excelTemplateFattura.writeInFile(nomeCondominio.ToUpper(), TipologiaWriteTemplateExcel.NOME_CONDOMINIO_FATTURA);
                                excelTemplateFattura.writeInFile(indirizzo.ToUpper(), TipologiaWriteTemplateExcel.VIA_CONDOMINIO_FATTURA);
                                excelTemplateFattura.writeInFile("/" + dateTimePicker1.Value.ToString("yy") + " S.R.L", TipologiaWriteTemplateExcel.ANNO_FATTURA);
                                excelTemplateFattura.writeInFile(comune + " (" + provincia.ToUpper() + "), " + cap.ToUpper(), TipologiaWriteTemplateExcel.COMUNE_CAP_PROVINCIA_CONDOMINIO_FATTURA);

                                String nomeFile = progressivo + "_" + amministratore + ".xlsx";
                                excelTemplateFattura.SalvataggioFile(folderOutput, nomeFile);
                                excelTemplateFattura = null;


                                if (!valid)
                                {
                                    validazioneGenerale = false;
                                }

                                delegates.changeNumberInNumericUpAndDownDelegate(numericUpDownNumeroFattura, progressivo);
                            }
                            else
                            {
                                delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Cliente alla riga " + i + " è stato saltato perché il campo fattura è stato impostato su 'N' o 'NO'");
                            }
                        }

                        delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Generazione Fatture Terminata. (Vedi in C:\\Fatture)");

                        workbookUltimaFattura.Worksheets[0][1, 1].Text = progressivo.ToString();
                        workbookUltimaFattura.Save();
                        this.ClearUI();
                        if (validazioneGenerale == false)
                        {
                            this.Invoke(new Action(() => MessageBox.Show(this, "Attenzione alcune fatture non sono state generate correttamente perché avevano dei campi mancanti. Nel log applicativo trovi maggiori informazioni per correggerle manualmente", "Warn", MessageBoxButtons.OK, MessageBoxIcon.Warning)));
                        }
                    }
                    else
                    {
                        delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Generazione Fatture Terminata con errore.");
                        this.Invoke(new Action(() => MessageBox.Show(this, "Esiste già una cartella per il mese corrente: eliminala/rinominala e riprova.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                        this.ClearUI();
                    }
                }
                else
                {
                    if(rowsExcelCosti < 2)
                    {
                        delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Il file con le ricevute dei clienti è vuoto. Impossibile generare documenti");
                        this.Invoke(new Action(() => MessageBox.Show(this, "Il file con le ricevute dei clienti è vuoto. Impossibile generare documenti", "Warn", MessageBoxButtons.OK, MessageBoxIcon.Warning)));
                    }
                    else if(rowsExcelAnagrafica < 2)
                    {
                        delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Il file con l'anagrafica dei clienti è vuoto. Impossibile generare documenti");
                        this.Invoke(new Action(() => MessageBox.Show(this, "Il file con l'anagrafica dei clienti è vuoto. Impossibile generare documenti", "Warn", MessageBoxButtons.OK, MessageBoxIcon.Warning)));
                    }
                }
            }
            catch (Exception ex)
            {
                delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Generazione Fatture Terminata con errore.");
                this.Invoke(new Action(() => MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                ClearUI();
            }

        }

        private void ClearUI()
        {

            delegates.disableEnableButtonDelegate(buttonFileCosti, true);
            delegates.disableEnableButtonDelegate(buttonFileAnagrafica, true);
            delegates.disableEnableButtonDelegate(buttonGeneraFatture, false);
            delegates.changeTextInTextBoxDelegate(textBoxFileCosti, "");
            delegates.changeTextInTextBoxDelegate(textBoxFileAnagrafica, "");
            this.InputFilePathCosti = "";
            this.InputFilePathAnagrafica = "";
        }

        private bool CheckFieldAndLog(int riga, int numeroFattura, String amministratore, String prezzo, String nomeCondominio, String cap, String provincia, String indirizzo, String comune)
        {
            bool valid = true;
            String log = $"Cliente alla riga {riga} Numero Fattura [{ numeroFattura }] - Non sono stati trovati i seguenti campi: |";

            if (amministratore.Equals(""))
            {
                valid = false;
                log += " AMMINISTRATORE |";
            }

            if (prezzo.Equals(""))
            {
                valid = false;
                log += " COSTO |";
            }

            if (nomeCondominio.Equals(""))
            {
                valid = false;
                log += " NOME CONDOMINIO |";
            }

            if (indirizzo.Equals(""))
            {
                valid = false;
                log += " INDIRIZZO |";
            }

            if (cap.Equals(""))
            {
                valid = false;
                log += " COMUNE |";
            }

            if (provincia.Equals(""))
            {
                valid = false;
                log += " PROVINCIA |";
            }

            if (cap.Equals(""))
            {
                valid = false;
                log += " CAP |";
            }

            if (valid == false)
            {
                delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, log);
            }

            return valid;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
