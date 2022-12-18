
using Genera_Fatture.PersonalUI;
using Genera_Fatture.Utils;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace Genera_Fatture
{
    public partial class form : Form
    {
        private String inputFilePathCosti = "";
        private String inputFilePathAnagrafica = "";
        public string InputFilePathCosti { get => inputFilePathCosti; set => inputFilePathCosti = value; }
        public string InputFilePathAnagrafica { get => inputFilePathAnagrafica; set => inputFilePathAnagrafica = value; }

        private SingletonFileInizializzazione singletonFile;

        private Delegates delegates;

        private String basePathOutputFile; // = $"C:\\Users\\{Environment.UserName}\\Desktop\\Fatture\\";
        private int progressivo;

        private ToolTip toolTipInputClientiAttivi = new ToolTip();
        private ToolTip toolTipInputAnagrafica = new ToolTip();


        public form()
        {
            this.basePathOutputFile = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)+"\\Fatture\\";
            InitializeComponent();
            CustomInizializeComponent();
            Console.WriteLine("LOADING");
        }

        private void CustomInizializeComponent()
        {
            try
            {
                //Iniziazzazione delegates
                delegates = new Delegates();

                //Inizializzazione openFileDialog
                this.openFileDialog.Filter = "Excel Files |*.xlsx";
                this.openFileDialog.Multiselect = false;
                this.openFileDialog.FileName = "";
                //$"C:\\users\\{System.Environment.UserName}\\Desktop";
                this.openFileDialog.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

                //Inizializzazione Button
                this.buttonGeneraFatture.Enabled = false;
                this.buttonGeneraFatture.ButtonColor = Color.Gray;

                //tool tip genera fatture
                ToolTip toolTip = new ToolTip();
                toolTip.ToolTipTitle = "Info";
                toolTip.SetToolTip(this.buttonGeneraFatture, "Non saranno generate le fatture dove non è presente il condominio nel file clienti attivi");
              
                //Inizializzazione checkbox
                this.checkBoxLog.Checked = false;
                this.checkBoxlog2.Checked = true;

                //Inizializzazione Data -> Default primo del mese
                DateTime now = DateTime.Now;
                this.dateTimePicker1.Value = new DateTime(now.Year, now.Month, 1);

                //inizializzazione ultima fattura.
                this.numericUpDownNumeroFattura.Minimum = 0;
                this.numericUpDownNumeroFattura.Maximum = 100000000;

                singletonFile = SingletonFileInizializzazione.getIstance();
                //workbookInizializzazione = new Workbook();
                //workbookInizializzazione.LoadFromFile("./Data/Inizializzazione.xlsx");
                //Worksheet worksheetUltimaFattura = workbookInizializzazione.Worksheets[0];            
                int progressivo = 0;
                if (DateTime.Now.Month == 1)
                {
                    progressivo = 0;
                }
                else
                {
                    progressivo = int.Parse(singletonFile.getNumeroUltimaFattura());
                    if (progressivo < 0)
                    {
                        progressivo = 0;
                    }
                }
                this.numericUpDownNumeroFattura.Value = progressivo;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore inaspettato. Se il problema persiste contattare il proprietario del software: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw ex;
            }
        }


        private void buttonFileCosti_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    this.textBoxFileCosti.Text = openFileDialog.FileName.ToString();
                    this.InputFilePathCosti = this.textBoxFileCosti.Text;
                    toolTipInputClientiAttivi.Active = true;
                    toolTipInputClientiAttivi.SetToolTip(this.textBoxFileCosti, this.InputFilePathCosti);
                    if (!inputFilePathAnagrafica.Equals(""))
                    {
                        //UI
                        this.buttonGeneraFatture.Enabled = true;
                        this.buttonGeneraFatture.ButtonColor = Color.DodgerBlue;
                    }
                    //clear open dialog
                    this.openFileDialog.FileName = "";
                    //$"C:\\users\\{System.Environment.UserName}\\Desktop";
                    this.openFileDialog.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

                }
                else
                {
                    toolTipInputClientiAttivi.Active = false;
                    this.textBoxFileCosti.Text = "";
                    this.InputFilePathCosti = "";
                    this.buttonGeneraFatture.Enabled = false;
                    this.buttonGeneraFatture.ButtonColor = Color.Gray;
                    //clear open dialog
                    this.openFileDialog.FileName = "";
                    //$"C:\\users\\{System.Environment.UserName}\\Desktop"
                    this.openFileDialog.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);            
                    MessageBox.Show(this, "Nessun file selezionato", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(this, "Errore inaspettato. Se il problema persiste contattare il proprietario del software: " + ex.Message, "Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void buttonFileAnagrafica_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    this.textBoxFileAnagrafica.Text = openFileDialog.FileName.ToString();
                    this.inputFilePathAnagrafica = this.textBoxFileAnagrafica.Text;
                    toolTipInputAnagrafica.Active = true;
                    toolTipInputAnagrafica.SetToolTip(this.textBoxFileAnagrafica, inputFilePathAnagrafica);

                    if (!inputFilePathCosti.Equals(""))
                    {
                        
                        //UI
                        this.buttonGeneraFatture.Enabled = true;
                        this.buttonGeneraFatture.ButtonColor = Color.DodgerBlue;
                    }
                    else
                    {
                        toolTipInputAnagrafica = null;
                    }
                    //clear open dialog
                    this.openFileDialog.FileName = "";
                    //$"C:\\users\\{System.Environment.UserName}\\Desktop"
                    this.openFileDialog.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

                }
                else
                {
                    toolTipInputAnagrafica.Active = false;
                    this.textBoxFileAnagrafica.Text = "";
                    this.inputFilePathAnagrafica = "";
                    this.buttonGeneraFatture.Enabled = false;
                    this.buttonGeneraFatture.ButtonColor = Color.Gray;
                    //clear open dialog
                    this.openFileDialog.FileName = "";
                    //$"C:\\users\\{System.Environment.UserName}\\Desktop"
                    this.openFileDialog.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    MessageBox.Show(this, "Nessun file selezionato", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(this, "Errore inaspettato. Se il problema persiste contattare il proprietario del software: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonGeneraFatture_Click(object sender, EventArgs e)
        {

            delegates.disableEnableButtonDelegate(buttonGeneraFatture, false);
            delegates.disableEnableButtonDelegate(buttonFileCosti, false);
            delegates.disableEnableButtonDelegate(buttonFileAnagrafica, false);
            delegates.disableEnableButtonSettingDelegate(buttonSettings, false);
            delegates.disableEnableCheckBox(checkBoxLog, false);
            delegates.disableEnableCheckBox(checkBoxLog, false);

            Thread t = new Thread(
                () => generafatture(InputFilePathCosti));
            t.Start();
        }

        private void generafatture(string file)
        {
            progressivo = (int)numericUpDownNumeroFattura.Value;
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
                int countRowEmpty = 0;

                if (rowsExcelCosti > 1 && rowsExcelAnagrafica > 1)
                {
                    bool validazioneGenerale = true;
                    String folderOutput = basePathOutputFile + dateTimePicker1.Value.ToString("yyyy") + "\\" + dateTimePicker1.Value.ToString("MMMM");

                    if (!Directory.Exists(folderOutput))
                    {
                       
                        for (int i = 2; i <= rowsExcelCosti; i++)
                        {
                            //100 righe vuote consecutive => Esci -> Limite per l'EOF di excel errato.
                            if(countRowEmpty == 100)
                            {
                                break;
                            }
                            String nomeCondominio = excelRicevute.retrieveCondominio(i);
                            if (!nomeCondominio.Equals(""))
                            {
                                countRowEmpty = 0;
                                Console.WriteLine("ROW COSTI: " + i);

                                // Carica file template
                                GestioneScritturaExcelTemplateFattura excelTemplateFattura = new GestioneScritturaExcelTemplateFattura("./Data/Template.xlsx");

                                //Lettura file input
                                String amministratore = excelRicevute.retrieveAmministratore(i);
                                String prezzo = excelRicevute.retrieveCosto(i);
                                String fattura = excelRicevute.retrieveFattura(i);
                                String data = dateTimePicker1.Value.ToString("dd/MM/yyyy");
                                String mese = "NEL MESE DI " + dateTimePicker1.Value.ToString("MMMM");
                                String sospeso = excelRicevute.retrieveSospesi(i);
                                String indirizzo = "";
                                String provincia = "";
                                String comune = "";
                                String cap = "";
                                String anno = dateTimePicker1.Value.ToString("yy");

                                if (!fattura.Trim().Equals("N") && !fattura.Trim().Equals("NO") && sospeso.Equals(""))
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
                                    excelTemplateFattura.writeInFile("/" + anno + " S.R.L", TipologiaWriteTemplateExcel.ANNO_FATTURA);
                                    excelTemplateFattura.writeInFile(comune + " (" + provincia.ToUpper() + "), " + cap.ToUpper(), TipologiaWriteTemplateExcel.COMUNE_CAP_PROVINCIA_CONDOMINIO_FATTURA);

                                    String nomeFile = progressivo + "_" + nomeCondominio + "_" + anno + ".xlsx";
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
                                    if (!sospeso.Equals("") && this.checkBoxLog.Checked == true)
                                    {
                                        delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, $"({nomeCondominio}) - riga [{i}] - è stato saltato perché il campo sospesi non è vuoto");
                                    }
                                    else
                                    {
                                        if (this.checkBoxlog2.Checked == true)
                                        {
                                            delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, $"({nomeCondominio}) - riga [{i}] - è stato saltato perché il campo fattura è stato impostato su 'N' o 'NO'");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("ROW COSTI: " + i + " SALTATA PERCHÈ VUOTA");
                                countRowEmpty++;
                            }
                        }

                        delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Generazione Fatture Terminata. (Vedi fatture in " + basePathOutputFile+ " )");

                        singletonFile.setNumeroUltimaFattura(progressivo.ToString());

                        this.ClearUI();
                        if (validazioneGenerale == false)
                        {
                            this.Invoke(new Action(() => MessageBox.Show(this, "Attenzione alcune fatture non sono state generate correttamente perché avevano dei campi mancanti o non validi. Nel log applicativo trovi maggiori informazioni per correggerle manualmente", "Warn", MessageBoxButtons.OK, MessageBoxIcon.Warning)));
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
                    if (rowsExcelCosti < 2)
                    {
                        delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Il file con le ricevute dei clienti è vuoto. Impossibile generare documenti");
                        this.Invoke(new Action(() => MessageBox.Show(this, "Il file con le ricevute dei clienti è vuoto. Impossibile generare documenti", "Warn", MessageBoxButtons.OK, MessageBoxIcon.Warning)));
                    }
                    else if (rowsExcelAnagrafica < 2)
                    {
                        delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Il file con l'anagrafica dei clienti è vuoto. Impossibile generare documenti");
                        this.Invoke(new Action(() => MessageBox.Show(this, "Il file con l'anagrafica dei clienti è vuoto. Impossibile generare documenti", "Warn", MessageBoxButtons.OK, MessageBoxIcon.Warning)));
                    }
                }

            }
            catch (Exception ex)
            {
                delegates.appendTextWithDateTimeInRichTextBoxLogDelegate(textBoxLog, "Generazione Fatture Terminata con errore.");
                this.Invoke(new Action(() => MessageBox.Show(this, "Errore inaspettato. Se il problema persiste contattare il proprietario del software: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                ClearUI();
            }

        }

        private void ClearUI()
        {
            try
            {
                delegates.disableEnableButtonDelegate(buttonFileCosti, true);
                delegates.disableEnableButtonDelegate(buttonFileAnagrafica, true);
                delegates.disableEnableButtonSettingDelegate(buttonSettings, true);
                delegates.disableEnableButtonDelegate(buttonGeneraFatture, false);
                delegates.disableEnableCheckBox(checkBoxLog, true);
                delegates.disableEnableCheckBox(checkBoxLog, true);
                delegates.changeTextInTextBoxDelegate(textBoxFileCosti, "");
                delegates.changeTextInTextBoxDelegate(textBoxFileAnagrafica, "");
                toolTipInputAnagrafica.Active = false;
                toolTipInputClientiAttivi.Active = false;
                this.InputFilePathCosti = "";
                this.InputFilePathAnagrafica = "";
            }
            catch(Exception ex)
            {
                MessageBox.Show(this, "Errore inaspettato. Se il problema persiste contattare il proprietario del software: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool CheckFieldAndLog(int riga, int numeroFattura, String amministratore, String prezzo, String nomeCondominio, String cap, String provincia, String indirizzo, String comune)
        {
            bool valid = true;
            String log = $"({nomeCondominio}) - riga [{riga}] - Numero Fattura [{ numeroFattura }] - Valori non trovati o non validi per i seguenti campi: |";

            //if (amministratore.Equals(""))
            //{
            //    valid = false;
            //    log += " AMMINISTRATORE |";
            //}

            if (prezzo.Equals(""))
            {
                valid = false;
                log += " COSTO |";
            }
            else
            {
                try
                {
                    double prezzoValido = Math.Round(float.Parse(prezzo), 2);
                }
                catch (Exception ex)
                {
                    prezzo = "";
                    valid = false;
                    log += " COSTO |";
                }
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


        private void buttonSettings_Click(object sender, EventArgs e)
        {
            try
            {
                DialogSettings dialogSettings = new DialogSettings();
                DialogResult dialogResult = dialogSettings.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Errore inaspettato. Se il problema persiste contattare il proprietario del software: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void onLoad(object sender, EventArgs e)
        {
            Console.WriteLine("LOADED");
            this.Opacity = 1;
        }
    }
}
