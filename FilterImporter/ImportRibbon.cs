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
            Debug.WriteLine("File: " + xmlPath);

            try
            {
                using (FileStream myStream = new FileStream(xmlPath, FileMode.Open))
                {
                    var reader = XmlReader.Create(myStream, new XmlReaderSettings() { ConformanceLevel = ConformanceLevel.Document });
                    feed fe = new XmlSerializer(typeof(feed)).Deserialize(reader) as feed;

                    addRules(fe.entry);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Please wait until you're connected to server");
            }
        }

        private void Debug_Click(object sender, RibbonControlEventArgs e)
        {
            Folder defaultFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;
            if (defaultFolder != null)
            {
                //getOrCreateMailFolder("level1/level2");
                //enumerateFolder(defaultFolder);
                //DeleteAllRules();
            }
        }

        private void DeleteAllRules()
        {
            Rules rules = Globals.ThisAddIn.Application.Session.DefaultStore.GetRules();
            for(int i=1; i <= rules.Count; i++)
            {
                Debug.WriteLine("Removing rule");
                rules.Remove(i);
            }
            rules.Save();
            MessageBox.Show("All rules have been deleted.");
        }

        public void addRules(feedEntry[] entries)
        {
            Rules rules = Globals.ThisAddIn.Application.Session.DefaultStore.GetRules();
            int count = 0;
            try
            {
                foreach (feedEntry entry in entries)
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
                            rule.Conditions.From.Recipients.Add(p.value);
                            //rule.Conditions.From.Recipients.ResolveAll();
                            rule.Conditions.From.Enabled = true;
                        }
                        else if (p.name.Equals("to"))
                        {
                            rule.Conditions.SentTo.Recipients.Add(p.value);
                            //rule.Conditions.SentTo.Recipients.ResolveAll();
                            rule.Conditions.SentTo.Enabled = true;
                        }
                        else if (p.name.Equals("subject"))
                        {
                            rule.Conditions.Subject.Text = new string[] { p.value };
                            rule.Conditions.Subject.Enabled = true;
                        }
                        else if (p.name.Equals("hasTheWord"))
                        {
                            rule.Conditions.BodyOrSubject.Text = new string[] { p.value };
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
                }
                rules.Save();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error: Could not create all rules. Original error: " + ex.Message);
            }
            finally
            {
                MessageBox.Show(string.Format("Successfully created {0} rules.", count));
            }
        }

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
            
            return getOrCreateMailFolderHelper(folderName.Trim().Split('/'), 0, inboxFolder);
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
    }
}
