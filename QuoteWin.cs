using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;

namespace QuotingSkype
{
    public partial class QuoteWin : Form
    {
        private static readonly DateTime Epoch = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);

        public QuoteWin()
        {
            InitializeComponent();
            DateTextBox.Text = DateTime.Now.ToString();
        }

        private void CopyButton_Click(object sender, EventArgs e)
        {
            var dateTime = DateTime.Parse(DateTextBox.Text, DateTextBox.Culture, DateTimeStyles.AssumeLocal);

            var timeSpan = dateTime.ToUniversalTime() - Epoch;
            var dataObject = new DataObject();

            var messagePlain = string.Format(
                    "[{0:hh:mm:ss}] {1}: {2}",
                    dateTime,
                    AuthorFullNameTextBox.Text,
                    MessageTextBox.Text);
            var messageXml = string.Format(
                    "<quote author=\"{0}\" timestamp=\"{1}\">{2}</quote>",
                    AuthorTextBox.Text,
                    timeSpan.TotalSeconds,
                    MessageTextBox.Text);

            dataObject.SetData("System.String", messagePlain);
            dataObject.SetData("UnicodeText", messagePlain);
            dataObject.SetData("Text", messagePlain);
            dataObject.SetData("SkypeMessageFragment", new MemoryStream(Encoding.UTF8.GetBytes(messageXml)));
            dataObject.SetData("Locale", new MemoryStream(BitConverter.GetBytes(CultureInfo.CurrentCulture.LCID)));
            dataObject.SetData("OEMText", messagePlain);

            Clipboard.SetDataObject(dataObject, true);
        }

        private void TimeResetButton_Click(object sender, EventArgs e)
        {
            DateTextBox.Text = DateTime.Now.ToString();
        }

        private void ImportButton_Click(object sender, EventArgs e)
        {
            try
            {
                const string DataKey = "SkypeMessageFragment";
                Dictionary<string, object> boardData = new Dictionary<string, object>();
                string xmlMessage;
                string uniMessage = Clipboard.GetText();

                string refinedUser;
                string refinedName;
                string refinedMessage;
                string refinedTime;

                IDataObject clipData = Clipboard.GetDataObject();
                foreach (var format in clipData.GetFormats())
                {
                    boardData[format] = clipData.GetData(format);
                }

                using (StreamReader dataReader = new StreamReader(boardData[DataKey] as MemoryStream))
                {
                    xmlMessage = dataReader.ReadToEnd();
                    (boardData[DataKey] as MemoryStream).Seek(0, SeekOrigin.Begin);
                }
                xmlMessage = xmlMessage.Remove(0,(xmlMessage.IndexOf("author=\"") + 8));
                xmlMessage = xmlMessage.Substring(0,xmlMessage.LastIndexOf("<legacyquote>"));
                refinedUser = xmlMessage.Substring(0, xmlMessage.IndexOf("\""));
                xmlMessage = xmlMessage.Remove(0, (xmlMessage.IndexOf("authorname=\"") + 12));
                refinedName = xmlMessage.Substring(0, xmlMessage.IndexOf("\""));
                xmlMessage = xmlMessage.Remove(0,(xmlMessage.IndexOf("timestamp=\"") + 11));
                refinedTime = xmlMessage.Substring(0, xmlMessage.IndexOf("\""));
                xmlMessage = xmlMessage.Remove(0, xmlMessage.IndexOf("</legacyquote>") + 14);
                refinedMessage = xmlMessage.Substring(0, xmlMessage.Length);

                AuthorTextBox.Text = refinedUser;
                AuthorFullNameTextBox.Text = refinedName;
                MessageTextBox.Text = refinedMessage;

                DateTextBox.Text = Epoch.AddSeconds(double.Parse(refinedTime)).ToLocalTime().ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Quote not found.  Exception:\r\n" + ex);
            }
        }
    }
}
