using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using SharpCompress.Common;
using SharpCompress.Readers;
using System.Text.RegularExpressions;
using System.Xml;

namespace SkypeParser {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        OpenFileDialog ofd;
        dynamic skypeJson;
        string bkpPath;
        public MainWindow() {
            InitializeComponent();
#if DEBUG
            MessageBox.Show("Skies is running in debug mode. Debug mode is made for testing Skies and its features.",
                            "Info",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
#endif
            // Initialize OpenFileDialog
            ofd = new OpenFileDialog();
            ofd.Title = "Select Skype export file to read";
            ofd.DefaultExt = "*.tar";
            ofd.FileOk += Ofd_FileOk;
            ofd.Filter = "TAR archives|*.tar";
            // Show it
            ofd.ShowDialog();
        }

        private void Ofd_FileOk(object sender, System.ComponentModel.CancelEventArgs e) {
            // If the file is over 300 MB:
            if (new FileInfo(ofd.FileName).Length >= 300000000) {
                // Show a message box
                MessageBox.Show("Loading of this file might take a while and the app might appear as not responding. This is normal, please wait for the loading to finish.",
                                "Info",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
            }
            // Create a temp folder name
            var skiesFolder = Environment.GetEnvironmentVariable("temp") + "\\skies\\";
            var folder = skiesFolder + Guid.NewGuid().ToString();
            bkpPath = folder;
            // Create the skies folder, if needed
            if (!Directory.Exists(skiesFolder)) Directory.CreateDirectory(skiesFolder);
            // Create the temp folder
            Directory.CreateDirectory(folder);
            // Decompress the TAR file
            using (Stream stream = File.OpenRead(ofd.FileName))
            using (var reader = ReaderFactory.Open(stream)) {
                while (reader.MoveToNextEntry()) {
                    if (!reader.Entry.IsDirectory) {
                        reader.WriteEntryToDirectory(folder, new ExtractionOptions() {
                            ExtractFullPath = true,
                            Overwrite = true
                        });
                    }
                }
            }
#if DEBUG
            // Open explorer
            System.Diagnostics.Process.Start("explorer", folder);
#endif
            // If there is a messages.json:
            if (File.Exists(folder + "\\messages.json")) {
                // Read the JSON
                var json = File.ReadAllText(folder + "\\messages.json");
                skypeJson = JObject.Parse(json);
#if DEBUG
                // Show a message box with the JSON length
                MessageBox.Show(json.Length.ToString());
#endif
                // For each conversation:
                foreach (var c in skypeJson.conversations) {
                    // Get the display name and ID
                    string disp = c.displayName.ToString();
                    string id = c.id.ToString();
                    // Display the display name (or ID if there is no display name)
                    chatlist.Items.Add(string.IsNullOrEmpty(disp) ? "[" + id + "]" : disp);
                }
            }
            // Else:
            else {
                // Complain
                var res = MessageBox.Show("This file isn't a valid Skype export. Please select another.",
                            "Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                // If user clicked OK:
                if (res == MessageBoxResult.OK) {
                    // Close all dialogs created by Skies
                    DialogCloser.Execute();
                    // Clear the filename
                    ofd.FileName = "";
                    // Show the open file dialog again
                    ofd.ShowDialog();
                }
            }
        }

        private void chatlist_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            // For each conversation:
            foreach (var c in skypeJson.conversations) {
                // If the name or ID match the selected contact's name or ID:
                if (chatlist.SelectedItem.ToString() ==  c.displayName.ToString() || chatlist.SelectedItem.ToString() == "[" + c.id.ToString() + "]") {
                    // Clear the message list
                    msglist.Items.Clear();
                    // For each message:
                    foreach (var m in c.MessageList) {
                        // Get the content
                        string content = m.content.ToString();
                        // If it's a call:
                        if (content.Contains("partlist")) {
                            // Create an XML document
                            // Yes, you heard it correctly, XML in a JSON file! Thanks Microsoft, that'll make parsing your stupid exports
                            // a whole lot harder! (Skies parses the XML though, unlike the official web-based parser)
                            XmlDocument xmlDoc = new XmlDocument();
                            // Load the XML content
                            xmlDoc.LoadXml(content.ToString());
                            // If there is a call length:
                            if (xmlDoc["partlist"]["part"]["duration"] != null) {
                                // Get call length from the first part
                                string lengthStr = xmlDoc["partlist"]["part"]["duration"].InnerText;
                                // Round up the length
                                int length = (int)Math.Round(float.Parse(lengthStr));
                                // Turn it into minutes
                                double lengthMins = Math.Round((double)(length / 60), 2);
                                // Create a new content
                                content = "[" + lengthMins.ToString() + " minute call]";
                            }
                            // Else:
                            else {
                                // Just make the content "call"
                                content = "[Call]";
                            }
                        }
                        // If it's a file:
                        else if (content.Contains("To view this")) {
                            // Load the content XML
                            var doc = new XmlDocument();
                            doc.LoadXml(content.ToString().Split(new string[]{"<c_i id"}, StringSplitOptions.RemoveEmptyEntries)[0]);
                            string fileName = "";
                            // If it's an image:
                            if (doc.SelectSingleNode("URIObject").Attributes["type"].Value.Contains("Picture")) {
                                // Get the file name
                                if (doc.SelectSingleNode("URIObject").Attributes["doc_id"] != null)
                                    fileName = doc.SelectSingleNode("URIObject").Attributes["doc_id"].Value;
                                else if (doc.SelectSingleNode("URIObject").Attributes["ams_id"] != null)
                                    fileName = doc.SelectSingleNode("URIObject").Attributes["ams_id"].Value;
                                // Get the additional number
                                fileName += "." + doc.SelectSingleNode("URIObject").Attributes["type"].Value.Split('.')[1];
                                // Get the extension
                                fileName += "." + doc.SelectSingleNode("URIObject/OriginalName").Attributes["v"].Value.Split('.')[1];
                                // If the filename isn't ".1.png", set the content to explanation text
                                if (fileName != ".1.png")
                                content = "[File, double click to view]                   (" + fileName + ")";
                                // Else set the content to "File"
                                content = "[File]";
                            }
                            // Else:
                            else {
                                // Set filename to URL
                                fileName = doc.SelectSingleNode("URIObject/a").Attributes["href"].Value;
                                // Set the content to simply "File"
                                content = "[File]";
                            }
                        }
                        // If it's an emoji:
                        else if (content.ToString().Contains("ss")) {
                            // Set the content to the emoji's code
                            content = Regex.Replace(content, "<[^>]+>", string.Empty);
                        }
                        // If we can load it as an XML:
                        else if (!string.IsNullOrEmpty(content.ToString()) && content.ToString().TrimStart().StartsWith("<")) {
                            // Set the content to explanation text
                            content = "[Unsupported message type]";
                        }
                        // Fix HTML escapes
                        content = content.Replace("&apos;", "'")
                                         .Replace("&lt;", "<")
                                         .Replace("&gt;", ">");
                        // Display the name and arrival time
                        msglist.Items.Add(m.from.ToString().Split(':')[1] + ", " + m.originalarrivaltime);
                        // Show the content
                        msglist.Items.Add(content);
                        // Add a blank space
                        msglist.Items.Add("");
                    }
                }
            }
        }

        private void msglist_MouseDoubleClick(object sender, MouseButtonEventArgs e) {
            // If this is a file:
            var si = msglist.SelectedItem.ToString();
            if (si.Contains("File, double")) {
                
                // Get the filename
                var filename = si.Substring(si.IndexOf("(") + 1, (si.IndexOf(")") - si.IndexOf("(")) - 1);
                // If it's a link:
                if (filename.Contains("http")) {
                    // Do nothing. Skype's links are broken.
                }
                // Else:
                else {
                    // Start the default handler for this file
                    System.Diagnostics.Process.Start(bkpPath + "\\media\\" + filename);
                }
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
            // Delete the backup folder
            Directory.Delete(bkpPath, true);
        }
    }
}
