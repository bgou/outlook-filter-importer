using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;
using System.Collections;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
using System.Threading;

namespace OutlookAddIn1
{
    public partial class ImportRibbon
    {
        private static string xmlPath = string.Empty;

        private void ImportRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Debug.WriteLine("Ribbon loaded");
        }

        private void SelectXml_Click(object sender, RibbonControlEventArgs e)
        {
            Debug.WriteLine("SelectXml_Click");
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Filter = "xml files (*.xml)|*.xml";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Debug.WriteLine("File: " + openFileDialog1.FileName);
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            var reader = XmlReader.Create(myStream, new XmlReaderSettings() { ConformanceLevel = ConformanceLevel.Document });
                            feed fe = new XmlSerializer(typeof(feed)).Deserialize(reader) as feed;
                            int i = 0;
                            foreach (feedEntry en in fe.entry)
                            {
                                Debug.WriteLine("Entry {0}", ++i);
                                foreach (property p in en.property)
                                {
                                    Debug.WriteLine("{0}: {1}", p.name, p.value);
                                }
                                Debug.WriteLine("");
                            }
                        }
                    }
                    xmlPath = openFileDialog1.FileName;
                    MessageBox.Show("XML file is okay. Click Start to begin.");
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void Import_Click(object sender, RibbonControlEventArgs e)
        {
            Debug.WriteLine("Import_Click");
            if (string.IsNullOrEmpty(xmlPath) || !System.IO.File.Exists(xmlPath))
            {
                MessageBox.Show("Please select a filter first", "Invalid input");
                return;
            }

            Debug.WriteLine("File: " + xmlPath);
            StartBackground(AddAllRules);
        }

        private void Debug_Click(object sender, RibbonControlEventArgs e)
        {
            Folder defaultFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;
            if (defaultFolder != null)
            {
                //enumerateFolder(defaultFolder);

                DialogResult dialogResult = MessageBox.Show("This will delete all rules from client and server, are you sure?", "Confirm", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    StartBackground(DeleteAllRules);
                }
            }
        }

        private void AddAllRules()
        {
            Rules rules = GetExchangeRules();

            if (rules == null)
                return;

            int count = 0;

            try
            {
                using (FileStream myStream = new FileStream(xmlPath, FileMode.Open))
                {
                    var reader = XmlReader.Create(myStream, new XmlReaderSettings() { ConformanceLevel = ConformanceLevel.Document });
                    feed fe = new XmlSerializer(typeof(feed)).Deserialize(reader) as feed;

                    foreach (feedEntry entry in fe.entry)
                    {
                        Rule rule = rules.Create(entry.id, OlRuleType.olRuleReceive);
                        rule.Enabled = true;
                        string destinationFolder = string.Empty;
                        bool skipInbox = false;

                        foreach (property p in entry.property)
                        {
                            if (p.name.Equals("label"))
                            {
                                destinationFolder = p.value;
                            }
                            else if (p.name.Equals("shouldArchive"))
                            {
                                bool.TryParse(p.value, out skipInbox);
                            }
                            else if (p.name.Equals("from"))
                            {
                                string value = CleanForAddress(p.value);
                                rule.Conditions.From.Recipients.Add(value);
                                //rule.Conditions.From.Recipients.ResolveAll();
                                rule.Conditions.From.Enabled = true;
                            }
                            else if (p.name.Equals("to"))
                            {
                                string value = CleanForAddress(p.value);
                                rule.Conditions.SentTo.Recipients.Add(value);
                                //rule.Conditions.SentTo.Recipients.ResolveAll();
                                rule.Conditions.SentTo.Enabled = true;
                            }
                            else if (p.name.Equals("subject"))
                            {
                                string[] words = CleanForSubjectOrBody(p.value);
                                rule.Conditions.Subject.Text = words;
                                rule.Conditions.Subject.Enabled = true;
                            }
                            else if (p.name.Equals("hasTheWord"))
                            {
                                string[] words = CleanForSubjectOrBody(p.value);
                                rule.Conditions.BodyOrSubject.Text = words;
                                rule.Conditions.BodyOrSubject.Enabled = true;
                            }
                            else if (p.name.Equals("doesNotHaveTheWord"))
                            {
                                // don't know/have the matching option in outlook
                            }
                            else if (p.name.Equals("hasAttachment"))
                            {
                                rule.Conditions.HasAttachment.Enabled = true;
                            }
                            else if (p.name.Equals("excludeChats"))
                            {
                                // don't know/have the matching option in outlook
                            }
                            else if (p.name.Equals("shouldMarkAsRead"))
                            {
                                // not supported while creating rules programmatically
                                // http://msdn.microsoft.com/en-us/library/bb206764.aspx
                            }
                            else if (p.name.Equals("shouldStar"))
                            {
                                // not supported while creating rules programmatically
                                // http://msdn.microsoft.com/en-us/library/bb206764.aspx
                            }
                            else if (p.name.Equals("shouldTrash"))
                            {
                                rule.Actions.Delete.Enabled = true;
                            }
                            else if (p.name.Equals("shouldNeverSpam"))
                            {
                                // don't know/have the matching option in outlook
                            }
                            else if (p.name.Equals("shouldAlwaysMarkAsImportant"))
                            {
                                // not supported while creating rules programmatically
                                // http://msdn.microsoft.com/en-us/library/bb206764.aspx
                            }
                            else if (p.name.Equals("smartLabelToApply"))
                            {
                                // not supported while creating rules programmatically
                                // http://msdn.microsoft.com/en-us/library/bb206764.aspx
                            }
                            else if (p.name.Equals("size"))
                            {
                                // not supported while creating rules programmatically
                                // http://msdn.microsoft.com/en-us/library/bb206764.aspx
                            }
                            else if (p.name.Equals("sizeOperator"))
                            {
                                // not supported while creating rules programmatically
                                // http://msdn.microsoft.com/en-us/library/bb206764.aspx
                            }
                            else if (p.name.Equals("sizeUnit"))
                            {
                                // not supported while creating rules programmatically
                                // http://msdn.microsoft.com/en-us/library/bb206764.aspx
                            }
                        }

                        var customFolder = getOrCreateMailFolder(destinationFolder);

                        if (skipInbox)
                        {
                            rule.Actions.MoveToFolder.Folder = customFolder;
                            rule.Actions.MoveToFolder.Enabled = true;
                        }
                        else
                        {
                            rule.Actions.CopyToFolder.Folder = customFolder;
                            rule.Actions.CopyToFolder.Enabled = true;
                        }

                        count++;
                        Debug.WriteLine("Adding rule " + count);
                    }
                    Debug.WriteLine("Saving rule changes to server");
                    rules.Save();
                    Debug.WriteLine("Changes uploaded to server");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error: Could not create all rules. Cause:\n" + ex.Message, "Error");
            }
            finally
            {
                MessageBox.Show(string.Format("Successfully created {0} rules.", count), "Success");
            }
        }

        private string CleanForAddress(string value)
        {
            value = value.Trim();
            value = Regex.Replace(value, "^.*?(\\w.*?),+$", "$1");
            return value;
        }

        private string[] CleanForSubjectOrBody(string value)
        {
            value = value.Trim();
            string[] words = value.Split(new string[] { " OR " }, StringSplitOptions.RemoveEmptyEntries);
            return words;
        }

        private void DeleteAllRules()
        {
            int count = 0;

            Rules rules = GetExchangeRules();
            if (rules != null)
            {
                for (int i = 1; i <= rules.Count; i++)
                {
                    Debug.WriteLine("Removing rule " + i);
                    rules.Remove(i);
                    count++;
                }
                Debug.WriteLine("Saving rule changes to server");
                rules.Save();
                Debug.WriteLine("Changes uploaded to server");
                MessageBox.Show(string.Format("{0} rules have been deleted.", count), "Success");
            }
        }

        private Rules GetExchangeRules()
        {
            Rules rules = null;
            try
            {
                rules = Globals.ThisAddIn.Application.Session.DefaultStore.GetRules();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            return rules;
        }

        private void StartBackground(System.Action action)
        {
            Thread worker = new Thread(new ThreadStart(action));
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
        }

        #region Helpers

        /// <summary>
        /// Enumerate folder by depth-first search
        /// </summary>
        /// <param name="folder"></param>
        private void enumerateFolder(Folder folder)
        {
            if (folder != null)
            {
                Debug.WriteLine(string.Format("Folder: {0}", folder.FolderPath));
                foreach (Folder f in folder.Folders)
                {
                    enumerateFolder(f);
                }
            }
        }

        /// <summary>
        /// The path of the folder under Inbox. e.g. "dev/inbox1". Will create recursively
        /// </summary>
        /// <param name="folderName">the folder name separated by '/'</param>
        /// <returns>An existing or a new Folder</returns>
        private Folder getOrCreateMailFolder(string folderName)
        {
            // Default the targetFolder to the inbox
            Folder inboxFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;

            string[] paths = folderName.Trim().Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            return getOrCreateMailFolderHelper(paths, 0, inboxFolder);
        }

        private Folder getOrCreateMailFolderHelper(string[] paths, int index, Folder parentFolder)
        {
            if (paths == null || index == paths.Length)
            {
                return parentFolder;
            }

            string curDir = paths[index];
            if (string.IsNullOrWhiteSpace(curDir))
                return parentFolder;

            Folder curFolder = null;
            // See if one exists
            foreach (Folder f in parentFolder.Folders)
            {
                if (f.Name.Equals(curDir))
                {
                    curFolder = f;
                }
            }
            // if curFolder doesn't exist
            if (curFolder == null)
            {
                curFolder = (Folder)parentFolder.Folders.Add(curDir);
            }
            return getOrCreateMailFolderHelper(paths, index+1, curFolder);
        }

        #endregion
    }
}
