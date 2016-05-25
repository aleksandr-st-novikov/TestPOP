using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TestPOP
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            List<string> seenUids = new List<string>();
            seenUids.Add("531d2cb32c09c3a8c5c6c36ba1b88cae");
            List<OpenPop.Mime.Message> mes = FetchUnseenMessages("pop.yandex.ru", 995, true, "test@e-tiande.by", "test0101", seenUids);

            foreach (OpenPop.Mime.Message msg in mes)
            {
                var att = msg.FindAllAttachments();
                foreach (var ado in att)
                {
                    ado.Save(new System.IO.FileInfo(System.IO.Path.Combine("d:\\att", ado.FileName)));
                }
            }

            string ff = "bfaac90cc38811c94cf67738372c87e6";
        }

        public static List<OpenPop.Mime.Message> FetchUnseenMessages(
            string hostname, 
            int port, 
            bool useSsl, 
            string username, 
            string password, 
            List<string> seenUids)
        {
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);

                // Fetch all the current uids seen
                List<string> uids = client.GetMessageUids();

                // Create a list we can return with all new messages
                List<OpenPop.Mime.Message> newMessages = new List<OpenPop.Mime.Message>();

                // All the new messages not seen by the POP3 client
                for (int i = 0; i < uids.Count; i++)
                {
                    string currentUidOnServer = uids[i];
                    if (!seenUids.Contains(currentUidOnServer))
                    {
                        // We have not seen this message before.
                        // Download it and add this new uid to seen uids

                        // the uids list is in messageNumber order - meaning that the first
                        // uid in the list has messageNumber of 1, and the second has 
                        // messageNumber 2. Therefore we can fetch the message using
                        // i + 1 since messageNumber should be in range [1, messageCount]
                        OpenPop.Mime.Message unseenMessage = client.GetMessage(i + 1);

                        // Add the message to the new messages
                        newMessages.Add(unseenMessage);

                        // Add the uid to the seen uids, as it has now been seen
                        seenUids.Add(currentUidOnServer);
                    }
                }

                // Return our new found messages
                return newMessages;
            }
        }


        public static void DeleteMessageOnServer(string hostname, int port, bool useSsl, string username, string password, string Uid)
        {
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);

                // Mark the message as deleted
                // Notice that it is only MARKED as deleted
                // POP3 requires you to "commit" the changes
                // which is done by sending a QUIT command to the server
                // You can also reset all marked messages, by sending a RSET command.

                int messageCount = client.GetMessageCount();

                // Run trough each of these messages and download the headers
                for (int messageItem = messageCount; messageItem > 0; messageItem--)
                {
                    var ddd = client.GetMessageHeaders(messageItem);
                    var dfd = client.GetMessageUid(messageItem);
                    // If the Message ID of the current message is the same as the parameter given, delete that message
                    if (client.GetMessageUid(messageItem) == Uid)
                    //if (client.GetMessageHeaders(messageItem).MessageId == Uid)
                    {
                        // Delete
                        client.DeleteMessage(messageItem);
                    }
                }
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            DeleteMessageOnServer("pop.yandex.ru", 995, true, "test@e-tiande.by", "test0101", "bfaac90cc38811c94cf67738372c87e6");
        }
    }
}
