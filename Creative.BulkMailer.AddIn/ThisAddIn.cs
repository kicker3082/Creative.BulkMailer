using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace Creative.BulkMailer.AddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        /// <summary>
        /// Inspect the Bcc recipients for the mail item and create a set of new
        /// emails cloned from the original, and replace the To: recipient with
        /// the Bcc recipient in each new email.
        /// </summary>
        /// <param name="original"></param>
        /// <returns></returns>
        IEnumerable<Outlook.MailItem> transposeBccs(Outlook.MailItem original)
        {
            return new List<Outlook.MailItem>();
        }

        /// <summary>
        /// Intercept outgoing mail and generate and send individual emails
        /// from the Bcc recipients. Send the original email with the Bcc
        /// recipients removed.
        /// </summary>
        /// <param name="Item"></param>
        /// <param name="Cancel"></param>
        void Application_ItemSend(object Item, ref bool Cancel)
        {

            var newRecipients = new List<Outlook.Recipient>();

            Outlook.Recipient recipient = null;
            Outlook.Recipients recipients = null;
            Outlook.MailItem mail = Item as Outlook.MailItem;
            if (mail != null)
            {
                foreach (Outlook.Recipient r in mail.Recipients)
                {

                    if (r.Type == (int)Outlook.OlMailRecipientType.olBCC)
                    {
                        newRecipients.Add(r); 
                    }
                }

                string addToSubject = " !IMPORTANT";
                string addToBody = "Sent from my Outlook 2010";
                if (!mail.Subject.Contains(addToSubject))
                    mail.Subject += addToSubject;
                if (!mail.Body.EndsWith(addToBody))
                    mail.Body += addToBody;
                recipients = mail.Recipients;
                recipient = recipients.Add("Eugene Astafiev");
                recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;
                recipient.Resolve();
                if (recipient != null) Marshal.ReleaseComObject(recipient);
                if (recipients != null) Marshal.ReleaseComObject(recipients);
            }
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
