﻿/*Copyright 2009 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
 * 
 * UserControl1.cs - the api called from MarkLogicPowerPointAddin.js.  The methods here map directly to functions in the .js.
 * 
*/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PwrPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;
//using DocumentFormat.OpenXml.Packaging; //OpenXML sdk
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using PPT = Microsoft.Office.Interop.PowerPoint;
using OX = DocumentFormat.OpenXml.Packaging;



namespace MarkLogic_PowerPointAddin
{
    [ComVisible(true)]
    public partial class UserControl1 : UserControl
    {
        private AddinConfiguration ac = AddinConfiguration.GetInstance();
        private string webUrl = "";
        private bool debug = false;
        private string color = "";
        private string addinVersion = "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION";
        HtmlDocument htmlDoc;

        public UserControl1()
        {
            InitializeComponent();
           // bool regEntryExists = checkUrlInRegistry();
            webUrl = ac.getWebURL();

            if (webUrl.Equals(""))
            {
                //MessageBox.Show("Unable to find configuration info. Please insure OfficeProperties.txt exists in your system temp directory.  If problems persist, please contact your system administrator.");
                MessageBox.Show("                                   Unable to find configuration info. \n\r " +
                                " Please see the README for how to add configuration info for your system. \n\r " +
                                "           If problems persist, please contact your system administrator.");
            }
            else
            {
                color = TryGetColorScheme().ToString();
                webBrowser1.AllowWebBrowserDrop = false;
                webBrowser1.IsWebBrowserContextMenuEnabled = false;
                webBrowser1.WebBrowserShortcutsEnabled = false;
                webBrowser1.ObjectForScripting = this;
                webBrowser1.Navigate(webUrl);
                webBrowser1.ScriptErrorsSuppressed = true;

                this.webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted);

            }

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser1.Document != null)
            {
                htmlDoc = webBrowser1.Document;

                htmlDoc.Click += htmlDoc_Click;

            }
        }
       
       private void htmlDoc_Click(object sender, HtmlElementEventArgs e)
       {
            if (!(webBrowser1.Parent.Focused))
            {

                webBrowser1.Parent.Focus();
                webBrowser1.Document.Focus();
               
            }

        }
        /*
        private bool checkUrlInRegistry()
        {
            RegistryKey regKey1 = Registry.CurrentUser;
            regKey1 = regKey1.OpenSubKey(@"MarkLogicAddinConfiguration\PowerPoint");
            bool keyExists = false;
            if (regKey1 == null)
            {
                if (debugMsg)
                    MessageBox.Show("KEY IS NULL");

            }
            else
            {
                if (debugMsg)
                    MessageBox.Show("KEY IS: " + regKey1.GetValue("URL"));

                webUrl = (string)regKey1.GetValue("URL");
                if (!((webUrl.Equals("")) || (webUrl == null)))
                    keyExists = true;
            }
            return keyExists;
        }
       */
        public enum ColorScheme : int
        {
            Blue = 1,
            Silver = 2,
            Black = 3,
            Unknown = 4
        };

        public ColorScheme TryGetColorScheme()
        {

            ColorScheme CurrentColorScheme = 0;
            try
            {
                Microsoft.Win32.RegistryKey rootKey = Microsoft.Win32.Registry.CurrentUser;
                Microsoft.Win32.RegistryKey registryKey = rootKey.OpenSubKey("Software\\Microsoft\\Office\\12.0\\Common");
                if (registryKey == null) return ColorScheme.Unknown;

                CurrentColorScheme =
                    (ColorScheme)Enum.Parse(typeof(ColorScheme), registryKey.GetValue("Theme").ToString());
            }
            catch
            { }

            return CurrentColorScheme;
        }

        public String getOfficeColor()
        {
            return color;
        }

        public String getAddinVersion()
        {
            return addinVersion;
        }

        public String getBrowserUrl()
        {
            return webUrl;
        }
        public String getCustomXMLPartIds()
        {

            string ids = "";

            try
            {
                PwrPt.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
                int count = pres.CustomXMLParts.Count;

                foreach (Office.CustomXMLPart c in pres.CustomXMLParts)
                {
                    if (c.BuiltIn.Equals(false))
                    {
                        ids += c.Id + " ";// "U+016000";


                    }
                }

                char[] space = { ' ' };
                ids = ids.TrimEnd(space);

                //char[] tengwar = { 'U', '+', '0', '1', '6', '0', '0', '0' };
                //ids = ids.TrimEnd(tengwar);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                ids = "error: " + errorMsg;
            }

            if (debug)
                ids = "error";

            return ids;
        }


        public String getCustomXMLPart(string id)
        {

            string custompiecexml = "";

            try
            {
                PwrPt.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
                Office.CustomXMLPart cx = pres.CustomXMLParts.SelectByID(id);

                if (cx != null)
                    custompiecexml = cx.XML;

                /*another way (used until I discovered SelectByID(id) above)
                  keeping here for notes, but this is for Word, translate to XL
                  foreach (Office.CustomXMLPart c in doc.CustomXMLParts)
                  {
                      if (c.BuiltIn.Equals(false) && c.Id.Equals(id))
                      {
                          Office.CustomXMLNode x = c.DocumentElement;
                          custompiecexml = x.XML;
                      }
                
                  }
                 */
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                custompiecexml = "error: " + errorMsg;
            }

            if (debug)
                custompiecexml = "error";

            return custompiecexml;

        }

        public String addCustomXMLPart(string custompiecexml)
        {
            string newid = "";
            try
            {
                PwrPt.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
                Office.CustomXMLPart cx = pres.CustomXMLParts.Add(String.Empty, new Office.CustomXMLSchemaCollectionClass());
                cx.LoadXML(custompiecexml);
                newid = cx.Id;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                newid = "error: " + errorMsg;
            }
            if (debug)
                newid = "error";

            return newid;

        }

        public String deleteCustomXMLPart(string id)
        {
            string message = "";
            try
            {
                PwrPt.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
                foreach (Office.CustomXMLPart c in pres.CustomXMLParts)
                {
                    if (c.BuiltIn.Equals(false) && c.Id.Equals(id))
                    {
                        //Office.CustomXMLNode x = c.DocumentElement;
                        c.Delete();
                    }

                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            if (debug)
                message = "error";

            return message;

        }

        public Image byteArrayToImage(byte[] byteArrayIn)
        {
            try
            {
                MemoryStream ms = new MemoryStream(byteArrayIn);
                Image returnImage = Image.FromStream(ms);
                return returnImage;
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        public byte[] imageToByteArray(System.Drawing.Image imageIn)
        {
            try
            {
                MemoryStream ms = new MemoryStream();
                imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Gif);
                return ms.ToArray();
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        public String insertImage(string imageuri, string uname, string pwd)
        {
            object missing = Type.Missing;
            string message = "";

            try
            {
                byte[] bytearray = downloadData(imageuri, uname, pwd);
                Image img = byteArrayToImage(bytearray);

                PPT.Slide slide = (PPT.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                Clipboard.SetImage(img);
                slide.Shapes.Paste();
                Clipboard.Clear();
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }


            return message;
        }

        public String getFileName()
        {
            string filename = "";
            filename = Globals.ThisAddIn.Application.ActivePresentation.Name;
            return filename;
        }

        public String getPath()
        {
            string path = "";
            path = Globals.ThisAddIn.Application.ActivePresentation.Path;
            return path;
        }

        public String getTempPath()
        {
            string tmpPath = "";
            try
            {
                tmpPath = System.IO.Path.GetTempPath();
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                tmpPath = "error: " + errorMsg;
            }

            return tmpPath;
        }

        static bool FileInUse(string path)
        {
            string __message = "";
            try
            {
                //Just opening the file as open/create
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    //If required we can check for read/write by using fs.CanRead or fs.CanWrite
                }
                return false;
            }
            catch (IOException ex)
            {
                //check if message is for a File IO
                __message = ex.Message.ToString();
                if (__message.Contains("The process cannot access the file"))
                    return true;
                else
                    throw;
            }
        }
        /*==========================================*/
        /*==========================================*/
        /*==========================================*/
        public string embedXLSX(string path, string title, string url, string user, string pwd)
        {
            string message="foo";
            string tmpdoc = "";
            MessageBox.Show("test");
            // MessageBox.Show("In addin");
            object missing = System.Type.Missing;
            int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;

                         title=title.Replace("/","");
                         MessageBox.Show("title" + title);

                         try
                         {
                             tmpdoc = path + title;
                             downloadFile(url, tmpdoc, user, pwd);

                            /* System.Net.WebClient Client = new System.Net.WebClient();
                             Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                             tmpdoc = path + title;
                             //works thought path ends with / and doc starts with \ so you have C:tmp/\foo.xslx
                             //may need to fix
                             //MessageBox.Show("Tempdoc"+tmpdoc);
                             //Client.DownloadFile("http://w2k3-32-4:8000/test.xqy?uid=/Default.xlsx", tmpdoc);//@"C:\test2.xlsx");
                             Client.DownloadFile(url, tmpdoc);//@"C:\test2.xlsx");
                             * */
                         }
                         catch (Exception e)
                         {
                             MessageBox.Show("ERROR: "+e.Message);
                         }

                         try
                         {
                             Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddOLEObject(21, 105, 250, 250, "", tmpdoc, Microsoft.Office.Core.MsoTriState.msoFalse, "", 0, "", Microsoft.Office.Core.MsoTriState.msoFalse);
                         }
                         catch (Exception e)
                         {
                             MessageBox.Show("Error" + e.Message);
                         }

            /*
            PPT.Shape s = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddTable(2, 3, 50, 50, 450, 70);
            try
            {
                Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddOLEObject(21, 105, 250, 250, "", @"C:\Workflow101.docx", Microsoft.Office.Core.MsoTriState.msoFalse, "", 0, "", Microsoft.Office.Core.MsoTriState.msoFalse);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error" + e.Message);
            }



            PPT.Table tbl = s.Table;

            // MessageBox.Show(tbl.Rows.Count + "here");
            PPT.Cell cell = tbl.Rows[1].Cells[1];
            cell.Shape.TextFrame.TextRange.Text = "Foo";
            // PPT.Shapes s = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes;


            return "foo"; */
            return message;
        }

        public String openPPTX(string path, string title, string url, string user, string pwd)
        {
            //MessageBox.Show("in the addin path:"+path+  "      title:"+title+ "   uri: "+url+"user"+user+"pwd"+pwd);
            
            string message = "";
            object missing = Type.Missing;
            string tmpdoc = "";

            try
            {
                tmpdoc = path + title;
                downloadFile(url, tmpdoc, user, pwd);
                PPT.Presentation ppt = Globals.ThisAddIn.Application.Presentations.Open(tmpdoc, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue);
            }
            catch (Exception e)
            {
                //not always true, need to improve error handling or message or both
                string origmsg = "A presentation with the name '" + title + "' is already open. You cannot open two documents with the same name, even if the documents are in different \nfolders. To open the second document, either close the document that's currently open, or rename one of the documents.";
                MessageBox.Show(origmsg);
                string errorMsg = e.Message;
                message = "error: " + errorMsg;

            }

            return message;
        }

        public string copyPasteSlideToActive(string tmpPath, string filename, string slideidx,string url,string user, string pwd, string retain)
        {

            string message = "";
            object missing = Type.Missing;
            string sourcefile = "";
            string path = getTempPath();
            bool retainformat = false;
            bool proceed = false;

            PPT.Slides ss = Globals.ThisAddIn.Application.ActivePresentation.Slides;

            if (retain.ToLower().Equals("true"))
                retainformat = true;

            try
            {
                sourcefile = path + filename;
                if (FileInUse(sourcefile))
                {
                    string origmsg = "A presentation with the name '" + filename + "' is already open. You cannot open two documents with the same name, even if the documents are in different \nfolders. To open the second document, either close the document that's currently open, or rename one of the documents.";
                    MessageBox.Show(origmsg);
                    
                }
                else
                {
                    downloadFile(url, sourcefile, user, pwd);
                    proceed = true;
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show("issue with download"+e.Message+e.StackTrace);
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            
            try
            {
                if (proceed)
                {
                    PPT.Presentation sourcePres = Globals.ThisAddIn.Application.Presentations.Open(sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
                    int num = Convert.ToInt32(slideidx);
                    copyPasteSlideToActiveSupport(sourcePres, num, retainformat);
                    sourcePres.Close();
                    sourcePres = null;
                }
            }
            catch(Exception e)
            {
                //MessageBox.Show("Unable to open: "+e.Message);
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                    
            }

            return message;
        }
        public string copyPasteSlideToActiveSupport(PPT.Presentation sourcePres, int slideidx, bool retain)
        {
            //MessageBox.Show("retain=" + retain);
            //arguments need to include slide(s) to be inserted ..
            // user, pwd, url, title, tmpath(?), retainsourceformatting 

            //get index of starter slide and reset at end of function?
            //don't have to worry about if just inserting one slide at a time.
            //string sourcefile = @"C:\Aven_MarkLogicUserConference2009Exceling.pptx";

            PPT.Presentation activePres = Globals.ThisAddIn.Application.ActivePresentation;
            PPT.Slides activeSlides = activePres.Slides;
            PPT.Slides sourceSlides = sourcePres.Slides;

            for (int x = 1; x <= sourceSlides.Count; x++)
            {
                int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                int id = sourceSlides[x].SlideID;

                if (sourceSlides[x].SlideIndex == slideidx)
                {
                    sourceSlides.FindBySlideID(id).Copy();

                    try
                    {
      //Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid].Select();
      //activeSlides.Paste(sid).FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
                        if (retain)
                        {
                            //Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid].Select();
                            activeSlides.Paste(sid).FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
                            Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid].Select();
                            PPT.SlideRange sr = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;
                            sr.Design = sourcePres.SlideMaster.Design;
                           //Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid + 1].Select();
                        }
                        else
                        {
                            //Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid].Select();
                            activeSlides.Paste(sid).FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
                            Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid].Select();
                            PPT.SlideRange sr = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;
                            sr.Design = Globals.ThisAddIn.Application.ActivePresentation.SlideMaster.Design;
                        }
                     }
                    catch (Exception e)
                    {
                        MessageBox.Show("FAIL" + e.Message + "   " + e.StackTrace);
                    }
                }


            }

            return "foo";
        }

        public string convertFilenameToImageDir(string filename)
        {
            string imgDir = "";
            string tmpDir = "";
            string fname = "";

            string[] split = filename.Split(new Char[] { '\\' });
            fname = split.Last();
            tmpDir = filename.Replace(fname, "");
            fname = fname.Replace(".pptx", "_GIF"); //"_pptx_parts_GIF");

            //imgDir = tmpDir + fname;
            imgDir = getTempPath() + fname;
            //MessageBox.Show("imgdir: "+imgDir);
            return imgDir;

        }

  /*      public string useSaveFileDialog()
        {
            Prompt p = new Prompt();
            p.ShowDialog();
            string filename = p.pfilename;
            MessageBox.Show(filename);
            return filename;
        }
   * */
       

        public string useSaveFileDialogOrig()
        {

            SaveFileDialog s = new SaveFileDialog();
           
            s.Filter = "PowerPoint Presentation (*.pptx)|*.pptx|All files (*.*)|*.*";
            s.DefaultExt = "pptx";
            s.AddExtension = true;
           
            s.ShowDialog();

            return s.FileName;
        }

        public string downloadFile(string url, string sourcefile, string user, string pwd)
        {
            string message = "";
            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                Client.DownloadFile(url, sourcefile);
                Client.Dispose();
            }
            catch (Exception e)
            {
                throw (e);
            }

            return message;
        }


        public string uploadData(string url, byte[] content)
        {

            System.Net.WebClient Client = new System.Net.WebClient();
            Client.Headers.Add("enctype", "multipart/form-data");
            Client.Headers.Add("Content-Type", "application/octet-stream");
            Client.Credentials = new System.Net.NetworkCredential("oslo", "oslo");

            Client.UploadData(url, "POST", content);
            Client.Dispose();


            return "foo";

        }

        private byte[] downloadData(string url, string user, string pwd)
        {
            //MessageBox.Show("downloading data");
            byte[] bytearray;

            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                bytearray = Client.DownloadData(url);
                Client.Dispose();
            }
            catch (Exception e)
            {
                throw (e);
            }

            return bytearray;
        }

        public string saveToML(string filename, string url)
        {
            string message = "";

            try
            {

                FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                int length = (int)fs.Length;
                byte[] content = new byte[length];
                fs.Read(content, 0, length);

                try
                {
                    uploadData(url, content);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                   // MessageBox.Show("message1 :" + message);
                }

                fs.Dispose();
                fs.Close();

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                MessageBox.Show("message2 :" + message);
            }

            return message;
        }

        public string saveWithImages(string saveasname, bool saveas)
        {
            //dir parameter?  make optional in the javascript.  So you can save anywhere in ML.
            //remember to tied to filenames and mapping
           
            bool insuresaveas = saveas;
           // if (saveas.Equals("true"))
             //   insuresaveas = true;

            string message = "";
            PPT.Presentation pptx = Globals.ThisAddIn.Application.ActivePresentation;

            string url = "http://localhost:8023/ppt/api/upload.xqy?uid=";


            string path = pptx.Path;
            //dir parameter might be used here
            string filename = pptx.Name;
            string fullfilenamewithpath = "";
            string imgdirwithpath = "";
            string imgdir = "";

            if ((pptx.Name == null || pptx.Name.Equals("") || pptx.Path == null || pptx.Path.Equals(""))
                 ||insuresaveas)
            {
                fullfilenamewithpath = getTempPath() + saveasname + ".pptx"; // useSaveFileDialog()+".pptx";
                //MessageBox.Show("fullnamewithpath is now  " + fullfilenamewithpath);
                //here's where dir parameter might come in
                filename = fullfilenamewithpath.Split(new Char[] { '\\' }).Last();

                pptx.SaveAs(fullfilenamewithpath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoFalse);
                url = url + "/" + filename;

                saveToML(fullfilenamewithpath, url);  //rename saveActivePresentation() - see excel


                imgdirwithpath = convertFilenameToImageDir(fullfilenamewithpath);
                //dir parameter?
                imgdir = imgdirwithpath.Split(new Char[] { '\\' }).Last();

                saveImages(imgdirwithpath);
               // pptx.SaveAs(imgdir, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsGIF, Microsoft.Office.Core.MsoTriState.msoFalse);

            }
            else
            {
               // MessageBox.Show("In the else");
                fullfilenamewithpath = path + "\\" + filename;
               // MessageBox.Show("Saving " + fullfilenamewithpath);
                pptx.Save();

                url = url + "/" + filename;
                try
                {
                    saveToML(fullfilenamewithpath, url);
                    //save to ML

                    imgdirwithpath = convertFilenameToImageDir(fullfilenamewithpath);
                    imgdir = imgdirwithpath.Split(new Char[] { '\\' }).Last();

                    saveImages(imgdirwithpath);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error" + e.Message);
                }
                
                //pptx.SaveAs(imgdir, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsGIF, Microsoft.Office.Core.MsoTriState.msoFalse);

            }


           // MessageBox.Show("fullnamewithpath:  "+fullfilenamewithpath + " imgdir: "+imgdirwithpath );

            return message;
        }

        public string saveImages(string imgdirwithpath)
        {
            string message = "";
            string imgdir = imgdirwithpath.Split(new Char[] { '\\' }).Last();

            //name of folder with images, prepend with optional dir?
           // MessageBox.Show("IMGDIRWITHPATH.SPLIT.LAST: " + imgdir);
            imgdir = "/" + imgdir; // +"/";
            PPT.Presentation ppt = Globals.ThisAddIn.Application.ActivePresentation;

            //need some try/catch action here ( and all over the place)
            if (Directory.Exists(imgdirwithpath))
            {
                string[] files = Directory.GetFiles(imgdirwithpath);
                foreach (string s in files)
                {
                    File.Delete(s);
                }
                Directory.Delete(imgdirwithpath);
            }

            ppt.SaveAs(imgdirwithpath, PPT.PpSaveAsFileType.ppSaveAsGIF,Office.MsoTriState.msoFalse);

            string[] imgfiles = Directory.GetFiles(imgdirwithpath);

            foreach (string i in imgfiles)
            {
               // MessageBox.Show("filename: " + i);
                string fname = i.Split(new Char[] { '\\' }).Last();
                string fileuri = imgdir + "/" + fname;
                //convert this uri to .pptx slide.xml
                //als get index from here
                // add as parameters for upload.xqy doc properties

              //  MessageBox.Show("fileuri to save :" + fileuri);

                string parentprop = imgdir.Replace("_GIF", ".pptx");

                string slideprop = fname.Replace(".GIF", ".xml");
                slideprop = imgdir+"/ppt/slides/" + slideprop;
                slideprop = slideprop.Replace("_GIF", "_pptx_parts");
                slideprop = slideprop.Replace("Slide", "slide");

               // string slideprop = fileuri.Replace(".GIF", ".xml");
               // slideprop = slideprop.Replace("_GIF", "");


                string slideidx = fname.Replace("Slide", "");
                slideidx = slideidx.Replace(".GIF", "");

              //  MessageBox.Show("properties: parent: " + parentprop + " slide: " + slideprop + " idx: " + slideidx);



                //save to ml, pass imagesurl
                //link in .xqy to .pptx
                string url = "http://localhost:8023/ppt/api/upload.xqy?uid=" + fileuri+"&source="+parentprop+"&slide="+slideprop+"&idx="+slideidx;

                try
                {
                   
                    FileStream fs = new FileStream(i, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    int length = (int)fs.Length;
                    byte[] content = new byte[length];
                    fs.Read(content, 0, length);

                    try
                    {
                        uploadData(url, content);
                    }
                    catch (Exception e)
                    {
                        string errorMsg = e.Message;
                        message = "error: " + errorMsg;
                        MessageBox.Show("message1 :" + message);
                    }
                    
                    fs.Dispose();
                    fs.Close();
                   
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    MessageBox.Show("message2 :" + message);
                }
            }

            //have the images, now have to get files and upload to ML
            //don't delete til we've copied to ML


            //Directory.Delete(imgdir);
            return message;
        }

        public string insertExcel()
        {
            string message = "";
            return message;
        }

        public string insertText(string txt)
        {
            //int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;

            //PPT.Shapes s = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes;

            try
            {
                string orig =  Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.Text;
                Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.Text = orig + txt;
            }
            catch (Exception e)
            {
                MessageBox.Show("Please place select text insertion point with cursor.");
            }

           // PPT.TextRange tr = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes[1].TextFrame.TextRange;
           // tr.Text = "FOOO";
            return "Foo";
           
        }

        public string insertTable() //parameterize rows, columns, vals
        {
           // MessageBox.Show("In addin");
            object missing = System.Type.Missing;
            int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
            PPT.Shape s = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddTable(2, 3,50,50,450,70);
            try
            {
                Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddOLEObject(21, 105, 250, 250, "", @"C:\Workflow101.docx", Microsoft.Office.Core.MsoTriState.msoFalse, "", 0, "", Microsoft.Office.Core.MsoTriState.msoFalse);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error" + e.Message);
            }
            
            

            PPT.Table tbl = s.Table;
          
           // MessageBox.Show(tbl.Rows.Count + "here");
            PPT.Cell cell = tbl.Rows[1].Cells[1];
            cell.Shape.TextFrame.TextRange.Text = "Foo";
           // PPT.Shapes s = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes;
         

            return "foo";
        }


//====================================================================================================
//====================================================================================================
//====================================================================================================
    //  public static bool CopySlidesFromPPT(string sourcefile, string dstfile, out string exmsg)
        public string copyPasteSlideToActiveSupportBACKUP(PPT.Presentation sourcePres)
        {
            MessageBox.Show("Copy Pasting files  --");
            //MessageBox.Show("1: "+GC.MaxGeneration);

            //try getting this from server
// string sourcefile = @"C:\Aven_MarkLogicUserConference2009Exceling.pptx";

             PPT.Presentation activePres = Globals.ThisAddIn.Application.ActivePresentation;
            //MessageBox.Show("3: " + GC.MaxGeneration + "activepresegen: " + GC.GetGeneration(activePres));
// PPT.Presentation sourcePres = Globals.ThisAddIn.Application.Presentations.Open(sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);

            //activePres.SlideMaster.BackgroundStyle = sourcePres.SlideMaster.BackgroundStyle;
          
            PPT.Slides activeSlides = activePres.Slides;
            PPT.Slides sourceSlides = sourcePres.Slides;

            for (int x = 1; x < sourceSlides.Count; x++)
            {
                int id = sourceSlides[x].SlideID;
                //MessageBox.Show(id+"");
                sourceSlides.FindBySlideID(id).Copy();
                //sourcePres.SlideMaster.Background.
                //activePres.Application.ActiveWindow.View.PasteSpecial();
                //activeSlides.Paste(x);
                try
                {
                    int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                   // MessageBox.Show("Idx before:  " + Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex);

                    activeSlides.Paste(sid).FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
          //if need to pull in master, then (also don't set follow master background above
         
             Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid].Select();
             PPT.SlideRange sr = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;
             sr.Design = sourcePres.SlideMaster.Design;
             Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid+1].Select();
                    ///sr.BackgroundStyle = sourceSlides.FindBySlideID(id).BackgroundStyle;//sourcePres.SlideMaster.Background;
                  //  sr.ColorScheme = sourceSlides.FindBySlideID(id).ColorScheme;//sourcePres.SlideMaster.ColorScheme;
                   // sr.DisplayMasterShapes = //Microsoft.Office.Core.MsoTriState.msoTrue;

                 //activeSlides[x].Background.BackgroundStyle = sourceSlides.FindBySlideID(id).Background.BackgroundStyle;
                }
                catch (Exception e)
                {
                    MessageBox.Show("FAIL"+e.Message+"   "+e.StackTrace);
                }
                

            }

  

            MessageBox.Show("returning foo");
            return "foo";

        }

        //missing template style BLERG!
        public string copySlideToActive()
        {
            MessageBox.Show("Saving files 1");
            string sourcefile = @"C:\Aven_MarkLogicUserConference2009Exceling.pptx";
            //PPT.Presentation p = Globals.ThisAddIn.Application.ActivePresentation;

            PPT.Application ppa = new PPT.ApplicationClass();
            PPT.Presentations ppp = ppa.Presentations;
            //PPT.Presentation ppmp = null;

            PPT.Presentation ppmp = Globals.ThisAddIn.Application.ActivePresentation;

            PPT.Slides ppms = ppmp.Slides;
            
            

            PPT.Presentation ppps = ppp.Open( sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
            
            ppms.InsertFromFile( sourcefile, ppms.Count, 1, ppps.Slides.Count);
           

            //ppmp.SlideMaster.CustomLayouts.Add( ppps.SlideMaster.CustomLayouts);
            //ppmp.SlideMaster.
            ////ppmp.SlideMaster.Background.BackgroundStyle = ppps.SlideMaster.Background.BackgroundStyle;
           // ppmp.SlideMaster.ColorScheme = ppps.SlideMaster.ColorScheme;
           // ppmp.HandoutMaster.BackgroundStyle = ppps.HandoutMaster.BackgroundStyle;
            try
            {
                ppmp.SlideMaster.BackgroundStyle = ppps.SlideMaster.BackgroundStyle;
            }
            catch (Exception e)
            {
                MessageBox.Show("FAIL");
            }
            finally
            {

                ppps.Close();
            }
            

                                //ppmp.Close();
                    //
                    // Release the COM object holding the merged presentation
                    //
                    //Marshal.ReleaseComObject(ppmp);
                   // ppmp = null;
                    //
                    // Release the COM object holding the presentations
                    //
            Marshal.ReleaseComObject(ppp);
            ppp = null;
                    //
                    // Release the COM object holding the powerpoint application
                    //
            Marshal.ReleaseComObject(ppa);
            ppa = null;
                

            return "foo";

            
        }


        public string CopySlidesFromPPT()
        {
            MessageBox.Show("Saving files 2");
            string sourcefile = @"C:\Aven_MarkLogicUserConference2009Exceling.pptx";
            string dstfile = @"C:\JetBlue case study r6.pptx";
            string exmsg="";
            bool success = false;

            //
            // Initialise the exception message
            //
            exmsg = "";
            //
            // Create a link to the PowerPoint object model
            //
            PPT.Application ppa = new PPT.ApplicationClass();
            PPT.Presentations ppp = ppa.Presentations;
            PPT.Presentation ppmp = null;
            //
            // If the destination presentation exists on disk, load it so
            // that we can append the new slides
            //
            if (File.Exists(dstfile) == true)
            {
                try
                {
                    //
                    // Try and open the destination presentation
                    //
                    ppmp = ppp.Open(dstfile, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
                }
                catch (Exception ex)
                {
                    ppmp = null;
                    exmsg = ex.Message;
                }
            }
            else
            {
                //
                // Create a new presentation
                //
                try
                {
                    ppmp = ppp.Add(Microsoft.Office.Core.MsoTriState.msoFalse);
                }
                catch (Exception ex)
                {
                    ppmp = null;
                    exmsg = ex.Message;
                }
            }
            //
            // Do we have a master presentation ?
            //
            if (ppmp != null)
            {
                //
                // Point to the slides in the master presentation
                //
               PPT.Slides ppms = ppmp.Slides;
                try
                {
                    try
                    {
                        //
                        // Open the source presentation
                        //
                        PPT.Presentation ppps = ppp.Open(sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
                        try
                        {
                            //
                            // Insert the source slides onto the end of the merge presentation
                            //
                            ppms.InsertFromFile(sourcefile, ppms.Count, 1, ppps.Slides.Count);
                            //
                            // Save the merged presentation back to disk
                            //
                            ppmp.SaveAs(dstfile, PPT.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
                            //
                            // Signal success
                            success = true;
                        }
                        finally
                        {
                            //
                            // Close the source presentation
                            //
                            ppps.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        exmsg = ex.Message;
                    }
                }
                finally
                {
                    //
                    // Ensure the merge presentation is closed
                    //
                    ppmp.Close();
                    //
                    // Release the COM object holding the merged presentation
                    //
                    Marshal.ReleaseComObject(ppmp);
                    ppmp = null;
                    //
                    // Release the COM object holding the presentations
                    //
                    Marshal.ReleaseComObject(ppp);
                    ppp = null;
                    //
                    // Release the COM object holding the powerpoint application
                    //
                    Marshal.ReleaseComObject(ppa);
                    ppa = null;
                    ppa.Quit();
                }
            }
            return "TEST";
        }

        /*
        public String insertImage(string imageuri, string imagename)
        {
            object missing = Type.Missing;
            MessageBox.Show("Adding Image");
            string message = "";

            System.Net.WebClient Client = new System.Net.WebClient();
            Client.Credentials = new System.Net.NetworkCredential("zeke", "zeke");
            byte[] bytearray = Client.DownloadData(imageuri);
            Image img = byteArrayToImage(bytearray);
            //Image img = Image.FromFile(@"C:\gijoe_destro.jpg");


            PPT.Slide slide = (PPT.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            Clipboard.SetImage(img);
            slide.Shapes.Paste();
            Clipboard.Clear();
            return message;
        }
         * */

        public String addSlide()
        {
      
            MessageBox.Show("IN ADDIN");
            string message = "foo";
            object missing = Type.Missing;
            //string message="";
            string filename = @"C:\MarkLogic Connector for SharePoint r1,1.pptx";
         //   PPT.Application ppa = new PPT.ApplicationClass();
         //   ppa.Presentations.Open(filename,Microsoft.Office.Core.MsoTriState.msoFalse,Microsoft.Office.Core.MsoTriState.msoFalse,Microsoft.Office.Core.MsoTriState.msoFalse);
         //   ppa.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
         //   ppa.Activate();

            try
            { 
              
                PPT.Presentation actP = Globals.ThisAddIn.Application.ActivePresentation;
                

                PPT.Application ppa = new PPT.ApplicationClass();
              
                ppa.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                ppa.Presentations.Open(filename, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);

                PPT.Presentation sourceP = ppa.ActivePresentation;
                



                PPT.Slide slide = actP.Slides[1];
                //PPT.Slide slide = actP.Slides.Add(actP.Slides.Count, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);
              //  PPT.Slide s2 = actP.Slides[slide.SlideIndex];
                
                slide.Application.Activate();
                /*
                 * objSourcePresentation.Slides(SlideID).Copy()
                   objDestinationPresentation.Slides.Paste
                 * */
                MessageBox.Show(sourceP.Slides.Count+"");
                sourceP.Slides[1].Copy();
                //slide.BackgroundStyle = sourceP.SlideMaster.BackgroundStyle;

                actP.Slides.Paste(slide.SlideIndex);
                PPT.Slide s2 = actP.Slides[slide.SlideIndex - 1];
                s2.CustomLayout = sourceP.Slides[1].CustomLayout;

               
                
             //   actP.Application.Activate();
                //sourceP.Close();
                //slide.CustomLayout = sourceP.Slides[1].CustomLayout;
         //       actP.Slides.Paste(slide.SlideIndex);
                //slide.BackgroundStyle = sourceP.SlideMaster.BackgroundStyle;

                
                //actP.SlideMaster.BackgroundStyle = sourceP.SlideMaster.BackgroundStyle;
                //sourceP.Close();


                //slide = ppa.Presentations[1].Slides[4];
                //PPT.Application ppa = Globals.ThisAddIn.Application;
                //PPT.Presentations ppts = Globals.ThisAddIn.Application.Presentations;
                //ppts.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                //ppts.Open(filename, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);
                
               

               // PPT.Application ppa = new PPT.ApplicationClass();
               // ppa.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
               // PPT.Presentation ppts = Globals.ThisAddIn.Application.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
              
               //PPT.Presentations ppp = ppa.Presentations;
                //PPT.Presentation ppmp = null;
                //ppa.Presentations.Open(filename, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);
               // ppts.Open(filename, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);
               // PPT.Presentation p1 = Globals.ThisAddIn.Application.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
               // p1 = ppa.Presentations[1];
               // p1.Application.Activate();
                //ppa.Activate();
            }
            catch (Exception e)
            {
                MessageBox.Show("ERROR" + e.Message + "==================" + e.StackTrace);
            }
            

//PPTApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue; ;

          //  PPT.Presentations ppp = ppa.Presentations;
          //  ppp.Open(@"http://localhost:8011/MarkLogic Connector for SharePoint r1,1.pptx"

            //PPT.Application objSourcePresentation;
            
           // = File.Open(@"http://localhost:8011/MarkLogic Connector for SharePoint r1,1.pptx", FileMode.Open);


  //          'copies the source slide to the clipboard
//objSourcePresentation.Slides(SlideID).Copy()


//'appends the slide from the clipboard to the end of the other presentation
//objDestinationPresentation.Slides.Paste


            return message;

        }

        //TODO:
        //pass filename, imagename, uri - want to use client download to tmp file
        //insert image
        //delete tmp file
 /*      public String insertImageORIG(string imageuri, string imagename)
        {
            object missing = Type.Missing;
            MessageBox.Show("Adding Image");
            string message = "";
          //PPT.Slide s = Globals.ThisAddIn.Application.ActivePresentation.Slides[Globals.ThisAddIn.Application.ActivePresentation.Slides.];
//ONE WAY
           // try this Current slide? gijoe too
 PPT.Slide slide = (PPT.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;// app.ActiveWindow.View.Slide;

//ADDING SLIDE
          PPT.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
//ADDING SLIDE PPT.Slide slide =
//ADDING SLIDE presentation.Slides.Add(
//ADDING SLIDE presentation.Slides.Count + 1,
//ADDING SLIDE PPT.PpSlideLayout.ppLayoutPictureWithCaption);

            //can get this from byte array
          //Image img = Image.FromFile(@"C:\test.png");
          Image img = Image.FromFile(@"C:\gijoe_destro.jpg");
         
          Clipboard.SetImage(img);
          //slide.Shapes.Paste();
            //how to add to current slide?
          slide.Shapes.Paste();
         // presentation.Slides[1].Shapes.Paste();

          //  richTextBox1.SelectionStart = 0;
          //  richTextBox1.Paste();

          Clipboard.Clear();

//ONE WAY PPT.Shape shape = slide.Shapes[2];
// ONE WAY slide.Shapes.AddPicture(@"C:\test.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
//shape.Left, shape.Top, shape.Width, shape.Height);
//ONE WAY slide.Select();
            
           

                //shape.Left, shape.Top, shape.Width, shape.Height);
             

           // slide.Shapes.AddPicture();
            //PPT.Presentation s = Globals.ThisAddIn.Application.ActivePresentation;
            //s.Application.ActiveWindow.ActivePane;
           // PPT.Slide s = Globals.ThisAddIn.Application.ActivePresentation;
           // Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

   //         System.Net.WebClient Client = new System.Net.WebClient();
   //         Client.Credentials = new System.Net.NetworkCredential("zeke", "zeke");

   //         byte[] bytearray = Client.DownloadData(imageuri);
   //         Image img = byteArrayToImage(bytearray);


            //place on clipboard
   //         System.Windows.Forms.Clipboard.SetImage(img);
           
           // Globals.ThisAddIn.Application.Selection.Range.Paste();
 
            return message;
        }

  */



    }
}
