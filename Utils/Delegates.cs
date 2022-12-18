using Prototipo_Denso.PersonalUI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Genera_Fatture
{
    internal class Delegates
    {
        public void disableEnableButtonDelegate(CustomButton customButton, bool enable)
        {
            try
            {
                if (customButton.InvokeRequired)
                {
                    Action safeWrite = delegate { disableEnableButtonDelegate(customButton,enable); };
                    customButton.Invoke(safeWrite);
                }
                else
                {
                    if (enable)
                    {
                        customButton.Enabled = true;
                        customButton.ButtonColor = Color.DodgerBlue;
                    }
                    else
                    {
                        customButton.Enabled = false;
                        customButton.ButtonColor = Color.Gray;
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }


        public void disableEnableButtonSettingDelegate(CustomButton customButton, bool enable)
        {
            try
            {
                if (customButton.InvokeRequired)
                {
                    Action safeWrite = delegate { disableEnableButtonSettingDelegate(customButton, enable); };
                    customButton.Invoke(safeWrite);
                }
                else
                {
                    if (enable)
                    {
                        customButton.Enabled = true;
                    }
                    else
                    {
                        customButton.Enabled = false;
                    }
                }
            }

            catch (Exception ex)
            {
            }
        }
        public void changeTextInTextBoxDelegate(TextBox textBox, String text)
        {
            try
            {
                if (textBox.InvokeRequired)
                {
                    Action safeWrite = delegate { changeTextInTextBoxDelegate(textBox, text); };
                    textBox.Invoke(safeWrite);
                }
                else
                {
                    textBox.Text = text;
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void changeNumberInNumericUpAndDownDelegate(NumericUpDown numericUpDown ,int number)
        {
            try
            {
                if (numericUpDown.InvokeRequired)
                {
                    Action safeWrite = delegate { changeNumberInNumericUpAndDownDelegate(numericUpDown, number); };
                    numericUpDown.Invoke(safeWrite);
                }
                else
                {
                    numericUpDown.Value = number;
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void appendTextWithDateTimeInRichTextBoxLogDelegate(RichTextBox richTextBox,string logText)
        {
            try
            {
                if (richTextBox.InvokeRequired)
                {
                    Action safeWrite = delegate { appendTextWithDateTimeInRichTextBoxLogDelegate(richTextBox, logText); };
                    richTextBox.Invoke(safeWrite);
                }
                else
                {
                    if (logText != null)
                    {
                        richTextBox.AppendText(DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZfff") + "> " + logText + "\n");
                    }
                    else
                    {
                        richTextBox.Text = "";
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void disableEnableCheckBox(CheckBox checkbox, bool enable)
        {
            try
            {
                if (checkbox.InvokeRequired)
                {
                    Action safeWrite = delegate { disableEnableCheckBox(checkbox, enable); };
                    checkbox.Invoke(safeWrite);
                }
                else
                {
                    checkbox.Enabled = enable;
                }
            }
            catch (Exception ex)
            {
            }
        }       
    }
}
