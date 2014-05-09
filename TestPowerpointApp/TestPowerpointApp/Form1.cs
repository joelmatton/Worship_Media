﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using PPt = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Diagnostics;
using System.IO;


namespace TestPowerpointApp
{



    public partial class frmTestPowerPoint : Form
    {

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(IntPtr ZeroOnly, string lpWindowName);


        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);


        PPt._Application pptApplication;
        PPt.Presentation presentation;

        PPt.Slides slides;
        int slidescount;
        PPt.Slide slide;
        int slideIndex;

        System.Windows.Forms.ImageList SlideImages;

        public frmTestPowerPoint()
        {
            InitializeComponent();
        }

        private void frmTestPowerPoint_FormClosing(Object sender, FormClosingEventArgs e)
        {
            foreach (PPt.Presentation x in pptApplication.Presentations)
            {
                try { x.Close(); }
                catch (Exception ex) { };
            }

        }
        private void frmTestPowerPoint_Load(object sender, EventArgs e)
        {
            int sCounter=0;
            foreach (Screen x  in Screen.AllScreens)
            {
                sCounter+=1;
                txtScreens.Text += "(" + sCounter + ")" + x.DeviceName + Environment.NewLine;
                
                if (x.Primary == true)
                { txtScreens.Text += "\t" + " ISPRIMARY " + Environment.NewLine; }
                
                txtScreens.Text += "\t" + x.WorkingArea.ToString() + Environment.NewLine;
                txtScreens.Text += "\t" + x.Bounds.ToString() + Environment.NewLine;

                txtScreens.Text += Environment.NewLine + Environment.NewLine;
            }
        }

        private void btnCheckIsRunning_Click(object sender, EventArgs e)
        {


            pptApplication = new PPt.Application();
            try
            {
                pptApplication = null;
                // Get Running PowerPoint Application object 
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;
                
                // Get PowerPoint application successfully, then set control button enable 
                this.btnFirst.Enabled = true;
                this.btnNext.Enabled = true;
                this.btnPrevious.Enabled = true;
                this.btnLast.Enabled = true;
            if (pptApplication != null)
            {

                // Get Presentation Object 
                presentation = pptApplication.ActivePresentation;
                // Get Slide collection object 
                slides = presentation.Slides;
                // Get Slide count 
                slidescount = slides.Count;
                // Get current selected slide  
                try
                {
                    // Get selected slide object in normal view 
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch
                {
                    // Get selected slide object in reading view 
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                } 
            } 
                 
             }
            catch
            {
                MessageBox.Show("Please Run PowerPoint Firstly", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            } 

        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            try
            {
                // Call Select method to select first slide in normal view 
                slides[1].Select();
                slide = slides[1];
            }
            catch
            {
                // Transform to first page in reading view 
                pptApplication.SlideShowWindows[1].View.First();
                slide = pptApplication.SlideShowWindows[1].View.Slide;
            } 
        }

        private void btnLast_Click(object sender, EventArgs e)
        {
            try
            {
                slides[slidescount].Select();
                slide = slides[slidescount];
            }
            catch
            {
                pptApplication.SlideShowWindows[1].View.Last();
                slide = pptApplication.SlideShowWindows[1].View.Slide;
            } 
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            slideIndex = slide.SlideIndex + 1;
            if (slideIndex > slidescount)
            {
                MessageBox.Show("It is already last page");
            }
            else
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Next();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            } 
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            slideIndex = slide.SlideIndex - 1;
            if (slideIndex >= 1)
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Previous();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            else
            {
                MessageBox.Show("It is already Fist Page");
            } 
        }

        private void btnOpenPptDoc_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.SlideShowView SSView;

            pptApplication = new PPt.Application();
            try
            {
                pptApplication = null;
                // Get Running PowerPoint Application object 
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;

                // Get PowerPoint application successfully, then set control button enable 
                this.btnFirst.Enabled = true;
                this.btnNext.Enabled = true;
                this.btnPrevious.Enabled = true;
                this.btnLast.Enabled = true;
                if (pptApplication != null)
                {

                    // Get Presentation Object 
                    presentation = pptApplication.Presentations.Open("C:\\CROSSWAY\\12step_leaf_tree.ppt", MsoTriState.msoTrue,MsoTriState.msoFalse, MsoTriState.msoFalse);
                    SlideImages = GetSlideImages();
                    SlideImages.ImageSize = new Size(128, 128);
                    lstSlides.View = View.LargeIcon;
                    lstSlides.LargeImageList = SlideImages;
                    for (int j = 0; j < this.SlideImages.Images.Count; j++)
                    {
                        ListViewItem item = new ListViewItem();
                        item.ImageIndex = j;
                        this.lstSlides.Items.Add(item);
                    }
                    // Get Slide collection object 
                    slides = presentation.Slides;
                    // Get Slide count 
                    slidescount = slides.Count;
                    // Get current selected slide  
                    try
                    {
                        // Get selected slide object in normal view 
                        slide = presentation.Slides.FindBySlideID(slides[1].SlideID);
                        IntPtr screenClassWnd = (IntPtr)0;
                        IntPtr x = (IntPtr)0;
                        panel1.Controls.Add(pptApplication as Control);
                        PPt.SlideShowSettings sst1 = presentation.SlideShowSettings;

                        sst1.StartingSlide = 1;
                        sst1.EndingSlide = slides.Count;
                        //panel1.Dock = DockStyle.Bottom;
                        pptApplication.Height=panel1.Height;
                        sst1.ShowType = PPt.PpSlideShowType.ppShowTypeWindow;
                        sst1.Application.Width = panel1.Height;
                        sst1.Application.Height = panel1.Width;
                      //  sst1.ShowType = PPt.PpSlideShowType.ppShowTypeSpeaker;


                        PPt.SlideShowWindow sw = sst1.Run();
                        
                        IntPtr pptptr = (IntPtr)sw.HWND;
                        SetParent(pptptr, panel1.Handle);


                    }
                    catch
                    {
                        // Get selected slide object in reading view 
                        slide = pptApplication.SlideShowWindows[1].View.Slide;

                    }
                }

            }
            catch
            {
                MessageBox.Show("Please Run PowerPoint Firstly", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            } 


        }

        public System.Windows.Forms.ImageList GetSlideImages()
        {
            if (presentation == null)
                return null;

            var imageList = new System.Windows.Forms.ImageList();

            foreach (PPt.Slide slide in presentation.Slides)
            {
                var fileName = Path.Combine(
                    Path.GetTempPath(),
                    string.Format("Slide{0:00}.jpg", slide.SlideNumber));

                slide.Export(fileName, "JPG", 800, 600);

                imageList.Images.Add(Image.FromFile(fileName));
            }

            return imageList;
        }

        private void btnClosePPT_Click(object sender, EventArgs e)
        {
            foreach (PPt.Presentation x in pptApplication.Presentations)
            {
                try { x.Close();}catch (Exception ex){};
            }
        }
    }
}
