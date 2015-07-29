using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.IO;
using System.Globalization;
using System.Security;
using System.Runtime.InteropServices;

namespace PDFCONVERTER
{
    public partial class Form1 : Form
    {
        //for single folder
        DirectoryInfo folder;
        FileInfo[] images;
        


        String spath;
        String dpath;

        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
        }
        [SuppressUnmanagedCodeSecurity]
        internal static class SafeNativeMethods
        {
            [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
            public static extern int StrCmpLogicalW(string psz1, string psz2);
        }

        public sealed class NaturalStringComparer : IComparer<string>
        {
            public int Compare(string a, string b)
            {
                return SafeNativeMethods.StrCmpLogicalW(a, b);
            }
        }

        public sealed class NaturalFileInfoNameComparer : IComparer<FileInfo>
        {
            public int Compare(FileInfo a, FileInfo b)
            {
                return SafeNativeMethods.StrCmpLogicalW(a.Name, b.Name);
            }
        }
        Boolean convertsinglefolder(String path,String Destination)
        {
           

            folder = new DirectoryInfo(path);
            images = folder.GetFiles("*."+comboBox1.SelectedItem);
           
            if (images.Length > 0)
            {
                Array.Sort<FileInfo>(images, new NaturalFileInfoNameComparer());
              

                PdfDocument doc = new PdfDocument();
                doc.PageLayout = PdfPageLayout.SinglePage;

                foreach (FileInfo f in images)
                {
                    doc.Pages.Add(new PdfPage());
                    XGraphics xgr = XGraphics.FromPdfPage(doc.Pages[doc.Pages.Count - 1]);
                    XImage img = null;
                    Image testimage = Image.FromFile(f.Directory.ToString() + Path.DirectorySeparatorChar + f.Name);
                    
                    if (Array.IndexOf(testimage.PropertyIdList, 274) > -1)
                    {
                        var orientation = (int)testimage.GetPropertyItem(274).Value[0];
                        
                        switch (orientation)
                        {
                            case 1:
                                // No rotation required.
                                break;
                            case 2:
                                testimage.RotateFlip(RotateFlipType.RotateNoneFlipX);
                                break;
                            case 3:
                                testimage.RotateFlip(RotateFlipType.Rotate180FlipNone);
                                break;
                            case 4:
                                 testimage.RotateFlip(RotateFlipType.Rotate180FlipX);
                                break;
                            case 5:
                                testimage.RotateFlip(RotateFlipType.Rotate90FlipX);
                                break;
                            case 6:
                                 testimage.RotateFlip(RotateFlipType.Rotate90FlipNone);
                                break;
                            case 7:
                                testimage.RotateFlip(RotateFlipType.Rotate270FlipX);
                                break;
                            case 8:
                                testimage.RotateFlip(RotateFlipType.Rotate270FlipNone);
                                break;
                        }
                        // This EXIF data is now invalid and should be removed.
                        testimage.RemovePropertyItem(274);
                       
                        
                    }
                    img = XImage.FromGdiPlusImage(testimage);
                    foreach (var prop in testimage.PropertyItems)
                    {
                        if (prop.Id == 112 || prop.Id == 5029)
                        {
                            // do my rotate code - e.g "RotateFlip"
                            // Never get's in here - can't find these properties.
                        }
                    }
                    
                    

                    doc.Pages[doc.Pages.Count - 1].Width = img.PointWidth;
                    doc.Pages[doc.Pages.Count - 1].Height = img.PointHeight;

                    xgr.DrawImage(img, 0, 0);
                    
                }

                string foldername = path.Substring(1 + path.LastIndexOf(Path.DirectorySeparatorChar));
                
                if (Directory.Exists(Destination) == false)
                {
                    Directory.CreateDirectory(Destination);
                }

                doc.Save(Destination + Path.DirectorySeparatorChar + foldername + ".pdf");
                
                return true;
            }
            else
                return false;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();


            DialogResult r = fbd.ShowDialog();

            if (r == DialogResult.OK)
            {
                label2.Text = "Source : " + fbd.SelectedPath;
                spath = fbd.SelectedPath;
                //label3.Text = "Destination : " + fbd.SelectedPath;
                //dpath = fbd.SelectedPath;

            }
        }
        private Boolean dirhassubdir(String Path)
        {
            if (Directory.GetDirectories(Path).Length > 0)
                return true;
            else
                return false;
        }

        private void traverse(String source_root,String Current_dir_path, String Dest_Root)
        {
            if( dirhassubdir(Current_dir_path) ==true)  // no pdf conversion to be performed ,traversal required
            {
                String[] directories = Directory.GetDirectories(Current_dir_path);

                foreach (String selctd_dir in directories)
                    traverse(source_root, selctd_dir, Dest_Root);
            
            }
            else//leaf node
            {
                convertsinglefolder(Current_dir_path, Dest_Root + Current_dir_path.Substring(source_root.Length));
             
            }


        }
        private void button2_Click(object sender, EventArgs e)
        {
            label4.Text = "Please Wait";
            PleaseWait pw = new PleaseWait();
            pw.Show();
            pw.Refresh();
           

            if (spath != null)
            {
                if (dpath != null)
                {
                    if (spath != dpath)      //source and destination can not be same
                    {
                        traverse(spath, spath, dpath);
                    
                    }
                    else
                        MessageBox.Show("Destination can not be same as source, please choose a different directory.");
                
                }
                else
                    MessageBox.Show("ERROR-No Destination Path.");
                  
            }
            else
                MessageBox.Show("ERROR-No source Path");


            label4.Text = "";
            pw.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();


            DialogResult r = fbd.ShowDialog();

            if (r == DialogResult.OK)
            {
               
                label3.Text = "Destination : " + fbd.SelectedPath;
                dpath = fbd.SelectedPath;

            }
        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Program helps in converting a Directory full of Images to a Directory full of PDF Files.\n\nFor further information contact : gswebspace@gmail.com");
        }

      
    }
}
