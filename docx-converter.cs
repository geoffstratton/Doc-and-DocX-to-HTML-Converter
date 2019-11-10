using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using HtmlAgilityPack;
 
namespace DocConverter
{
    public partial class docForm : Form
    {
        public docForm()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(docForm_DragEnter);
            this.DragDrop += new DragEventHandler(docForm_DragDrop);
        }
 
        void docForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }
 
        void docForm_DragDrop(object sender, DragEventArgs e)
        {
            // Gives us the path to the file
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
             
            // Invoke Word, open doc by path, do doc.SaveAs to generate HTML
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            
            Document doc = application.Documents.Open(files[0]);
            string result = Path.GetTempPath();
            //More "complete" but worse HTML
            //doc.SaveAs(result + "temp.html", WdSaveFormat.wdFormatHTML);
            doc.SaveAs(result + "temp.html", WdSaveFormat.wdFormatFilteredHTML);
            doc.Close();
 
            // Close Word
            application.Quit();
 
            // Now, clean up Word's HTML using Html Agility Pack
            HtmlAgilityPack.HtmlDocument mangledHTML = new HtmlAgilityPack.HtmlDocument();
            mangledHTML.Load(result + "temp.html");
             
            //Uncomment to see results so far
            //Process.Start("notepad.exe", result + "temp.html");
             
            //"Blacklisted" tags and all inclusive data will be removed completely
            //"Stripped" tags will have all attributes removed, so <p class="someclass"> becomes <p>
            string[] blacklistedTags = { "span", "head" };
            string[] strippedTags = { "body", "div", "p", "strong", "ul", "li", "table", "tr", "td" };
             
            foreach(var blackTag in blacklistedTags) 
            {
                try
                {
                    foreach (HtmlNode item in mangledHTML.DocumentNode.SelectNodes("//" + blackTag))
                    {
                        item.ParentNode.RemoveChild(item);
                    }
                }
                catch (NullReferenceException)
                {
                    // No tags of that type; skip it and move on
                    continue;
                }
            }
 
            foreach(var stripTag in strippedTags)
            {
                try
                {
                    foreach (HtmlNode item in mangledHTML.DocumentNode.SelectNodes("//" + stripTag))
                    {
                        item.Attributes.RemoveAll();
                    }
                }
                catch (NullReferenceException)
                {
                    // No tags of that type; skip it and move on
                    continue;
                }
            }
 
            mangledHTML.Save(result + "newtemp.html");
 
            // Remove standalone CRLF 
            string badHTML = File.ReadAllText(result + "newtemp.html");
            badHTML = badHTML.Replace("\r\n\r\n", "ackThbbtt");
            badHTML = badHTML.Replace("\r\n", "");
            badHTML = badHTML.Replace("ackThbbtt", "\r\n");
            File.WriteAllText(result + "finaltemp.html", badHTML);
    
            // Clean up temp files, show the finished result in Notepad
            File.Delete(result + "temp.html");
            File.Delete(result + "newtemp.html");
            Process.Start("notepad.exe", result + "finaltemp.html");
        }
 
    }
 
}