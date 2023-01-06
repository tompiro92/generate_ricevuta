using Genera_Fatture.Utils;
using Spire.Xls;
using System;
using System.Windows.Forms;

namespace Genera_Fatture.PersonalUI
{
    public partial class DialogSettings : Form
    {
        private SingletonFileInizializzazione singletonFileInizializzazione;
        public DialogSettings()
        {
            InitializeComponent();
            InitializeCustom();
        }

        private void InitializeCustom()
        {

            singletonFileInizializzazione = SingletonFileInizializzazione.getIstance();

            numericUpAmministratore.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_AMMINISTRATORE));
            numericUpDownFattura.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_FATTURA));
            numericUpDownSospesi.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_SOSPESI));
            numericUpDownCondominioAttivi.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_CONDOMINIO));
            numericUpDownCosto.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO));

            numericUpDownPulizie.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_PULIZIE));
            numericUpDownBidoni.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_BIDONI));
            numericUpDownGiardini.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_GIARDINI));
            numericUpDownPaP.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_PORTA_A_PORTA));

            numericUpDownCondominioAnagrafica.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_CONDOMINIO));
            numericUpDownIndirizzo.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_INDIRIZZO));
            numericUpDownCAP.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_CAP));
            numericUpDownComune.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_COMUNE));
            numericUpDownProvincia.Value = int.Parse(singletonFileInizializzazione.getIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_PROVINCIA));              

        }

        private void buttonFileCosti_Click(object sender, EventArgs e)
        {
            try
            {
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_AMMINISTRATORE, numericUpAmministratore.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_FATTURA, numericUpDownFattura.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_SOSPESI, numericUpDownSospesi.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_CONDOMINIO, numericUpDownCondominioAttivi.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO, numericUpDownCosto.Value.ToString());
                
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_PULIZIE, numericUpDownPulizie.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_BIDONI, numericUpDownBidoni.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_GIARDINI, numericUpDownGiardini.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.CLIENTI_ATTIVI_COSTO_PORTA_A_PORTA, numericUpDownPaP.Value.ToString());

                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_CONDOMINIO, numericUpDownCondominioAnagrafica.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_INDIRIZZO, numericUpDownIndirizzo.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_CAP, numericUpDownCAP.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_COMUNE, numericUpDownComune.Value.ToString());
                singletonFileInizializzazione.setIndexOf(ValueInizializzazioneEnum.ANAGRAFICA_PROVINCIA, numericUpDownProvincia.Value.ToString());

                singletonFileInizializzazione.SaveFile();

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Errore inaspettato. Se il problema persiste contattare il proprietario del software: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
