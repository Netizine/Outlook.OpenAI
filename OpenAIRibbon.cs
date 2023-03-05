using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OpenAI
{
    public partial class OpenAIRibbon
    {
        private void OpenAIRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void OpenAIRespond_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrEmpty(Globals.ThisAddIn.OpenAIApiKey))
            {
                MessageBox.Show(Properties.Resources.OpenAIKeyNotSet);
                return;
            }
            try
            {
                if (!(e.Control.Context is Inspector inspector) ||
                    !(inspector.CurrentItem is MailItem mailItem)) return;
                string body;
                switch (mailItem.BodyFormat)
                {
                        case OlBodyFormat.olFormatHTML:

                        body = StripHtml(mailItem.HTMLBody);
                        break;
                    case OlBodyFormat.olFormatRichText:
                        body = StripRtf(mailItem.HTMLBody);
                        break;
                    default:
                        body = mailItem.Body;
                        break;
                }
                if (body.Length > 0)
                {
                    AddressEntry addrEntry = null;
                    // Get the Store for CurrentFolder
                    if (Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder is Folder folder)
                    {
                        var store = folder.Store;
                        var accounts = Globals.ThisAddIn.Application.Session.Accounts;
                        // Enumerate accounts to find account.DeliveryStore for store.
                        foreach (Account account in accounts)
                        {
                            if (account.DeliveryStore.StoreID !=
                                store.StoreID) continue;
                            addrEntry =
                                account.CurrentUser.AddressEntry;
                            break;
                        }
                    }

                    var chatMessage = new ChatCompletionMessage
                    {
                        Role = ChatRoles.User,
                        Content = "Please could you write a suitable email body reply to " + mailItem.SenderName + " for this email i received on the " + mailItem.SentOn.ToString("MM/dd/yyyy h:mm tt") + "?" +
                                  Environment.NewLine + "The email subject line is:" + mailItem.Subject +
                                  Environment.NewLine + "The email body is " + body + Environment.NewLine
                    };
                    var chatMessageList = new List<ChatCompletionMessage>
                    {
                        chatMessage
                    };
                    var chatCompletionOptions = new ChatGPT3CompletionCreateOptions
                    {
                        Model = "gpt-3.5-turbo",
                        Messages = chatMessageList,
                        Temperature = 0,
                    };
                    var chatCompletion = Globals.ThisAddIn.ChatGPT3Service.Create(chatCompletionOptions);
                    if (chatCompletion.Choices.Count > 0)
                    {
                        CreateMailItem(mailItem.SenderEmailAddress, mailItem.Subject,
                            chatCompletion.Choices[0].Message.Content,
                            addrEntry != null ? addrEntry.Name : "[Your Name]");
                    }
                }
            }
            catch (System.Exception ex)
            {
                Globals.ThisAddIn.LogMessage(ex.Message, ex.StackTrace);
                MessageBox.Show("An error has occurred. Please check the eventlog for detailed information." +
                                Environment.NewLine
                                + "Your event log can be found at " + Globals.ThisAddIn.LogFileLocation);
            }

        }


        /// <summary>
        /// Remove HTML from string.
        /// </summary>
        private static string StripHtml(string source)
        {
            return Dangl.TextConverter.Html.HtmlToText.ConvertHtmlToPlaintext(source);
        }

        /// <summary>
        /// Remove RTF from string.
        /// </summary>
        private static string StripRtf(string source)
        {
            return Dangl.TextConverter.Rtf.RtfToText.ConvertRtfToText(source);
        }

        private void CreateMailItem(string to, string subject, string body, string yourName)
        {
            var mailFooter =
                "This e-mail communication and any attachments are confidential and for the sole use of the intended recipient(s). Any review, reliance, dissemination, distribution, copying or other use is strictly prohibited and may be illegal. Any opinions expressed in this communication are personal and are not attributable to Atom Supplies Limited or any of its affiliated or parent companies.\r\n\r\nThe reliability of this method of communication cannot be guaranteed. It can be intercepted, corrupted, delayed, may arrive incomplete, contain viruses or be affected by other interference. We have taken reasonable steps to reduce risks against viruses but cannot accept liability for any damages sustained as a result of this transmission.\r\n\r\nIf you are not the intended recipient please delete this e-mail and notify the sender immediately by replying to this e-mail.\r\n\r\nAtom Group, Atom Brands, Master of Malt and Maverick Drinks are trading names of Atom Supplies Limited.  Registered office:  Unit 1, Ton Business Park, 2 -8 Morley Road, Tonbridge, Kent, TN9 1RA. Registered in England & Wales. Company number 3193057. VAT number GB 662241553.\r\n";
            body = body + Environment.NewLine + Environment.NewLine + mailFooter;
            var mailBody = Properties.Resources.HtmlBody.Replace("YourMessageBody", body);
            var mailItem = (MailItem)
                Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
            mailItem.Subject = "RE: " + subject;
            mailItem.To = to;
            mailItem.HTMLBody = "";
            mailItem.Body = body.TrimStart().Replace("[Your Name]", yourName);
            mailItem.Importance = OlImportance.olImportanceLow;
            mailItem.Display(false);
        }

    }
}
