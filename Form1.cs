
using Spire.Xls;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Genera_Fatture
{
    public partial class Form1 : Form
    {
        private String inputFilePathCosti="";
        private String inputFilePathAnagrafica="";
        public string InputFilePathCosti { get => inputFilePathCosti; set => inputFilePathCosti = value; }
        public string InputFilePathAnagrafica { get => inputFilePathAnagrafica; set => inputFilePathAnagrafica = value; }

        private const String NUMERO_FATTURA = "D9";
        private const String DATA_FATTURA = "D10";
        private const String MESE_FATTURA = "C18";
        private const String NOME_CONDOMINIO_FATTURA = "F7";
        private const String VIA_CONDOMINIO_FATTURA = "F10";
        private const String CAP_PROVINCIA_CONDOMINIO_FATTURA = "F11";
        private const String COSTO_FATTURA = "I16";

        private String basePathOutputFile =  $"C:\\Fatture\\";
        Workbook workbookUltimaFattura;
        int progressivo;

        public Form1()
        {
            InitializeComponent();
            CustomInizializeComponent();
        }

        private void CustomInizializeComponent()
        {
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
            this.numericUpDownNumeroFattura.Minimum = 1;
            this.numericUpDownNumeroFattura.Maximum = 100000000;

            workbookUltimaFattura = new Workbook();
            workbookUltimaFattura.LoadFromFile("./Data/NumeroUltimaFattura.xlsx");
            Worksheet worksheetUltimaFattura = workbookUltimaFattura.Worksheets[0];
            if (DateTime.Now.Month == 1)
            {
                progressivo = 1;
            }
            else
            {
                progressivo = int.Parse(worksheetUltimaFattura[1, 1].Value);
                if (progressivo < 1)
                {
                    MessageBox.Show(this, "Il numero progressivo non può essere minore di 1", "Warn", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    progressivo = 1;
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
                this.InputFilePathCosti = this.textBoxFileAnagrafica.Text;

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
            updateButtonGenerate(false);
            updateButtonAnagrafica(false);
            updateButtonCosti(false);

            Thread t = new Thread(
                () => generafatture(InputFilePathCosti));
            t.Start();
        }

        private void generafatture(string file)
        {
            updateRichTextBoxLog(null);
            updateRichTextBoxLog("Generazione Fatture Iniziata");
            // Carica file Fatture Input
            Workbook workbookInputCosti = new Workbook();
            workbookInputCosti.LoadFromFile(file);
            Worksheet worksheetInputCosti = workbookInputCosti.Worksheets[0];

            // Carica file anagrafiche clienti Input
            Workbook workbookInputAnagrafica = new Workbook();
            workbookInputAnagrafica.LoadFromFile(file);
            Worksheet worksheetInputAnagrafica = workbookInputAnagrafica.Worksheets[0];

            // Carica file template
            Workbook workbookTemplate = new Workbook();
            workbookTemplate.LoadFromFile("./Data/Template.xlsx");
            Worksheet worksheetTemplate = workbookTemplate.Worksheets[0];


            int rowsExcelCosti = worksheetInputCosti.Rows.Count();
            int rowsExcelAnagrafica = worksheetInputAnagrafica.Rows.Count();

            if (rowsExcelCosti > 1 && rowsExcelAnagrafica > 1)
            {
                bool validazioneGenerale = true;
                validazioneHeaderFileCosti(worksheetInputCosti); //can trows exception ( Invalid file ).
                validazioneHeaderFileAnagrafica(worksheetInputAnagrafica);

                for (int i = 2; i < rowsExcelCosti; i++)
                {
                    //Lettura file input
                    String amministratore = worksheetInputCosti[i, 1] != null ? worksheetInputCosti[i, 1].Value : "";
                    String prezzo = worksheetInputCosti[i, 10] != null ? worksheetInputCosti[i, 10].Value : "";
                    String nomeCondominio = worksheetInputCosti[i, 8] != null ? worksheetInputCosti[i, 8].Value : "";
                    String fattura = worksheetInputCosti[i, 4] != null ? worksheetInputCosti[i, 4].Value : "";
                    String data = dateTimePicker1.Value.ToString("dd/MM/yyyy");
                    String mese = "NEL MESE DI " + dateTimePicker1.Value.ToString("MMMM");

                    String indirizzo = "";
                    String provincia = "";
                    String comune = "";
                    String cap = "";

                    for(int j = 2; j < rowsExcelAnagrafica; ++ j)
                    {
                        String nomeCondominioAnagrafica = worksheetInputCosti[j, 2] != null ? worksheetInputCosti[j, 2].Value : "";
                        if (!nomeCondominioAnagrafica.Equals("") && nomeCondominioAnagrafica.Equals(nomeCondominio))
                        {
                            indirizzo = worksheetInputAnagrafica[i, 5] != null ? worksheetInputCosti[i, 5].Value : "";
                            cap = worksheetInputAnagrafica[i, 6] != null ? worksheetInputCosti[i, 6].Value : "";
                            comune = worksheetInputAnagrafica[i, 7] != null ? worksheetInputCosti[i, 7].Value : "";
                            provincia = worksheetInputAnagrafica[i, 8] != null ? worksheetInputCosti[i, 8].Value : "";
                        }
                    }
              

                    bool valid = CheckFieldAndLog(progressivo, amministratore, prezzo, nomeCondominio, cap, provincia, indirizzo, comune);

                    if (!fattura.Equals("N") && !fattura.Equals("NO"))
                    {
                        worksheetTemplate[NUMERO_FATTURA].Text = progressivo.ToString();
                        worksheetTemplate[DATA_FATTURA].Text = data.ToUpper();
                        worksheetTemplate[COSTO_FATTURA].Text = prezzo;
                        worksheetTemplate[MESE_FATTURA].Text = mese.ToUpper();
                        worksheetTemplate[NOME_CONDOMINIO_FATTURA].Text = nomeCondominio.ToString().ToUpper();
                        worksheetTemplate[VIA_CONDOMINIO_FATTURA].Text = indirizzo.ToUpper();
                        worksheetTemplate[CAP_PROVINCIA_CONDOMINIO_FATTURA].Text = provincia.ToUpper() + " " + cap.ToUpper();
                        String folderOutput = basePathOutputFile + dateTimePicker1.Value.ToString("yyyy") + "\\" + dateTimePicker1.Value.ToString("MMMM") + "\\";
                        Directory.CreateDirectory(folderOutput);
                        workbookTemplate.SaveToFile(folderOutput + progressivo + "_" + amministratore + ".xlsx");

                        if (!valid)
                        {
                            validazioneGenerale = false;
                        }

                        ++progressivo;
                        this.updateProgressiveNumberUpAndDown(progressivo);
                    }
                    else
                    {
                        updateRichTextBoxLog("Cliente alla riga " + i + " è stato saltato perché il campo fattura è stato impostato su 'N' o 'NO'");
                    }
                }
                if(validazioneGenerale == false)
                {
                    updateRichTextBoxLog("Attenzione alcune fatture non sono state generate correttamente perché avevano dei campi mancanti. Sopra trovi maggiori informazioni per correggerli manualmente");

                }
                updateRichTextBoxLog("Generazione Fatture Terminata.");
                workbookUltimaFattura.Worksheets[0][1, 1].Text = progressivo.ToString();
                workbookUltimaFattura.Save();
                workbookUltimaFattura.Dispose();
                workbookTemplate.Dispose();
                worksheetInputCosti.Dispose();
                this.ClearUI();

            }
            else
            {
                updateRichTextBoxLog("I file in input sono vuoti. Impossibile generare le fatture.");
            }

        }


        private void ClearUI()
        {
            
            updateButtonCosti(true);
            updateButtonGenerate(false);
            updateButtonAnagrafica(true);
            updateTextBoxFile("");
            this.InputFilePathCosti = "";
            this.InputFilePathAnagrafica = "";
        }

        private bool CheckFieldAndLog(int numeroFattura, String amministratore, String prezzo, String nomeCondominio, String cap, String provincia, String indirizzo, String comune)
        {
            bool valid = true;
            String log = $"***Numero Fattura [{ numeroFattura }] - Non sono stati trovati i seguenti campi: |";

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

           

            return true;
        }

        private void validazioneHeaderFileCosti(Worksheet worksheet)
        {
            //check header of file.
            String headerAmministratore = worksheet[1, 1] != null ? worksheet[1, 1].Value.ToUpper() : "";
            String headerFattura = worksheet[1, 4] != null ? worksheet[1, 4].Value.ToUpper() : "";
            String headerCantiere = worksheet[1, 8] != null ? worksheet[1, 8].Value.ToUpper() : "";
            String headerCosto = worksheet[1, 10] != null ? worksheet[1, 10].Value.ToUpper() : "";

            if ( !headerAmministratore.Equals("AMMINISTRATORE") 
                || (!headerFattura.Equals("FATT.") && !headerFattura.Equals("FATTURA")) 
                || (!headerCantiere.Equals("CANTIERE") && !headerCantiere.Equals("CONDOMINIO"))
                || (!headerCosto.Equals("COSTO") && !headerCosto.Equals("&euro;")) )
            {
                throw new Exception("Il file selezionato non è valido. L'intestazione dovrebbe contenere:\n" +
                    "Colonna A: AMMINISTRATORE\n" +
                    "Colonna D: FATTURA O FATT.\n" +
                    "Colonna H: CANTIERE O CONDOMINIO\n" +
                    "Colonna J: &euro; O COSTO");
            }
        }

        private void validazioneHeaderFileAnagrafica(Worksheet worksheetInputAnagrafica)
        {
            String headerCondominio = worksheetInputAnagrafica[1, 2] != null ? worksheetInputAnagrafica[1, 2].Value.ToUpper() : "";
            String headerIndirizzo = worksheetInputAnagrafica[1, 5] != null ? worksheetInputAnagrafica[1, 5].Value.ToUpper() : "";
            String headerCap = worksheetInputAnagrafica[1, 7] != null ? worksheetInputAnagrafica[1, 7].Value.ToUpper() : "";
            String headerComune = worksheetInputAnagrafica[1, 6] != null ? worksheetInputAnagrafica[1, 6].Value.ToUpper() : "";
            String headerProvincia = worksheetInputAnagrafica[1, 8] != null ? worksheetInputAnagrafica[1, 8].Value.ToUpper() : "";

            if (( !headerCondominio.Equals("CONDOMINIO") && !headerCondominio.Equals("PARCO") && !headerCondominio.Equals("CONDOMINIO/PARCO")) 
                || !headerIndirizzo.Equals("INDIRIZZO") 
                || !headerCap.Equals("CAP")
                || !headerComune.Equals("COMUNE")
                || (!headerProvincia.Equals("PROV") && !headerProvincia.Equals("PROVINCIA")))
            {
                throw new Exception("Il file selezionato non è valido. L'intestazione dovrebbe contenere:\n" +
                    "Colonna A: CONDOMINIO O CONDOMINIO/PARCO\n" +
                    "Colonna E: INDIRIZZO\n" +
                    "Colonna F: COMUNE\n" +
                    "Colonna G: CAP\n" + 
                    "Colonna H: PROVINCIA o PROV.");
            }
        }

        private void updateRichTextBoxLog(string logText)
        {
            try
            {
                if (this.textBoxLog.InvokeRequired)
                {
                    Action safeWrite = delegate { updateRichTextBoxLog(logText); };
                    this.textBoxLog.Invoke(safeWrite);
                }
                else
                {
                    if (logText != null)
                    {
                        this.textBoxLog.AppendText(DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZfff") + "> " + logText + "\n");
                    }
                    else
                    {
                        this.textBoxLog.Text = "";
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }


        private void updateButtonCosti(bool enable)
        {
            try
            {
                if (this.buttonFileCosti.InvokeRequired)
                {
                    Action safeWrite = delegate { updateButtonCosti(enable); };
                    this.buttonFileCosti.Invoke(safeWrite);
                }
                else
                {
                    if (enable)
                    {
                        this.buttonFileCosti.Enabled = true;
                        this.buttonFileCosti.ButtonColor = Color.DodgerBlue;
                    }
                    else
                    {
                        this.buttonFileCosti.Enabled = false;
                        this.buttonFileCosti.ButtonColor = Color.Gray;
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void updateButtonAnagrafica(bool enable)
        {
            try
            {
                if (this.buttonFileAnagrafica.InvokeRequired)
                {
                    Action safeWrite = delegate { updateButtonAnagrafica(enable); };
                    this.buttonFileAnagrafica.Invoke(safeWrite);
                }
                else
                {
                    if (enable)
                    {
                        this.buttonFileAnagrafica.Enabled = true;
                        this.buttonFileAnagrafica.ButtonColor = Color.DodgerBlue;
                    }
                    else
                    {
                        this.buttonFileAnagrafica.Enabled = false;
                        this.buttonFileAnagrafica.ButtonColor = Color.Gray;
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }


        private void updateButtonGenerate(bool enable)
        {
            try
            {
                if (this.buttonGeneraFatture.InvokeRequired)
                {
                    Action safeWrite = delegate { updateButtonGenerate(enable); };
                    this.buttonGeneraFatture.Invoke(safeWrite);
                }
                else
                {
                    if (enable)
                    {
                        this.buttonGeneraFatture.Enabled = true;
                        this.buttonGeneraFatture.ButtonColor = Color.DodgerBlue;
                    }
                    else
                    {
                        this.buttonGeneraFatture.Enabled = false;
                        this.buttonGeneraFatture.ButtonColor = Color.Gray;
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void updateProgressiveNumberUpAndDown(int x)
        {
            try
            {
                if (this.numericUpDownNumeroFattura.InvokeRequired)
                {
                    Action safeWrite = delegate { updateProgressiveNumberUpAndDown(x); };
                    this.numericUpDownNumeroFattura.Invoke(safeWrite);
                }
                else
                {
                    this.numericUpDownNumeroFattura.Value = x;
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void updateTextBoxFile(String path)
        {
            try
            {
                if (this.textBoxFileCosti.InvokeRequired)
                {
                    Action safeWrite = delegate { updateTextBoxFile(path); };
                    this.textBoxFileCosti.Invoke(safeWrite);
                }
                else
                {
                    this.textBoxFileCosti.Text = path;
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {

        }
    }
}
