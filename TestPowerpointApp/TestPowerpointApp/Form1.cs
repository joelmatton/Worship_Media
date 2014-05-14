using System;
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
using MediaManagers;

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
            if (pptApplication != null)
            {
                foreach (PPt.Presentation x in pptApplication.Presentations)
                {
                    try { x.Close(); }
                    catch  { };
                }
                pptApplication.Quit();
            }

            try
            {
                PPTMediaApplication.DisposeEverything();
            }
            catch  { }
            try
            {
                PPTMediaApplication=null;
            }
            catch { }

        }
        private void frmTestPowerPoint_Load(object sender, EventArgs e)
        {
            int sCounter = 0;
            foreach (Screen x in Screen.AllScreens)
            {
                sCounter += 1;
                txtScreens.Text += "(" + sCounter + ")" + x.DeviceName + Environment.NewLine;

                if (x.Primary == true)
                { txtScreens.Text += "\t" + " ISPRIMARY " + Environment.NewLine; }

                txtScreens.Text += "\t" + x.WorkingArea.ToString() + Environment.NewLine;
                txtScreens.Text += "\t" + x.Bounds.ToString() + Environment.NewLine;

                txtScreens.Text += Environment.NewLine + Environment.NewLine;
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



            try
            {
                // pptApplication = null;
                // Get Running PowerPoint Application object 

                Type powerpointType = Type.GetTypeFromProgID("PowerPoint.Application");
                object instance1 = Activator.CreateInstance(powerpointType);
                pptApplication = (PPt._Application)instance1;

                // Get PowerPoint application successfully, then set control button enable 
                this.btnFirst.Enabled = true;
                this.btnNext.Enabled = true;
                this.btnPrevious.Enabled = true;
                this.btnLast.Enabled = true;
                if (pptApplication != null)
                {

                    // Get Presentation Object 
                    presentation = pptApplication.Presentations.Open("C:\\CROSSWAY\\12step_leaf_tree.ppt", MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                    try { pptApplication.Visible = MsoTriState.msoFalse; }
                    catch { }
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
                        pptApplication.Height = panel1.Height;
                        sst1.ShowType = PPt.PpSlideShowType.ppShowTypeWindow;
                        sst1.Application.Width = panel1.Height;
                        sst1.Application.Height = panel1.Width;

                        try { pptApplication.Visible = MsoTriState.msoFalse; }
                        catch { }
                        sst1.ShowScrollbar = MsoTriState.msoFalse;



                        PPt.SlideShowWindow sw = sst1.Run();
                        try { pptApplication.Visible = MsoTriState.msoFalse; }
                        catch { }
                        IntPtr pptptr = (IntPtr)sw.HWND;
                        SetParent(pptptr, panel1.Handle);
                        try { pptApplication.Visible = MsoTriState.msoFalse; }
                        catch { }

                    }
                    catch
                    {
                        // Get selected slide object in reading view 
                        slide = pptApplication.SlideShowWindows[1].View.Slide;

                    }
                }

            }
            catch (Exception eMAin)
            {
                MessageBox.Show(eMAin.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(pptApplication);
            }

            this.BringToFront();
            this.WindowState = FormWindowState.Maximized;
            this.MinimumSize = this.Size;
            this.MaximumSize = this.Size;

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



        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }



        MediaManagers.PowerPointApplication PPTMediaApplication;
        MediaManagers.clsMediaListItem ActiveMediaItem;
        System.Windows.Forms.ImageList PreviewStripImageList = new System.Windows.Forms.ImageList();
        private void EnablePPTButtons()
        {
            this.btnFirst.Enabled = true;
            this.btnNext.Enabled = true;
            this.btnPrevious.Enabled = true;
            this.btnLast.Enabled = true;

        }



        private void btnPPTDll_Click(object sender, EventArgs e)
        {
            MediaGuiComponents.DialogFormCalls dfc = new MediaGuiComponents.DialogFormCalls();
            string pptFileName= dfc.GetPowerpointFile(this);
            if (pptFileName == "") { return; }
            if (PPTMediaApplication == null)
            {
                PPTMediaApplication = new MediaManagers.PowerPointApplication();
            } 
            PPTMediaApplication.AddPresentationFromFile(pptFileName);
            EnablePPTButtons();

          
        }

        private void btnNextPPTdll_Click(object sender, EventArgs e)
        {
            ActiveMediaItem = PPTMediaApplication.GetNextMediaListItem();
            DisplayFilmstripPreview(ActiveMediaItem);
            ShowSlides();
            
        }

        public void DisplayFilmstripPreview(clsMediaListItem ActiveMediaItem  )
        {
            PreviewStripImageList.Images.Clear();
            this.lstSlides.Items.Clear();
            switch (ActiveMediaItem.mediaType)
            {
                case MediaTypeEnum.PowerPointFile:

                    lstSlides.View = View.LargeIcon;
                    lstSlides.MultiSelect = false;
                    PreviewStripImageList.ImageSize = new Size(256, 256);
                    foreach (MediaManagers.clsSlidePictureItem x in ActiveMediaItem.SlideImages)
                    {
                        PreviewStripImageList.Images.Add(x.SlideImage);
                        ListViewItem item = new ListViewItem();
                        item.ImageIndex = x.ID; //slide id
                        this.lstSlides.Items.Add(item);



                    }
                    lstSlides.LargeImageList = PreviewStripImageList;

                    return;
                case MediaTypeEnum.PowerPointDynamic:
                    return;
                case MediaTypeEnum.Movie:
                    return;

            }

        }

        private void btnPrevPPTdll_Click(object sender, EventArgs e)
        {
            ActiveMediaItem = PPTMediaApplication.GetPrevMediaListItem();
            DisplayFilmstripPreview(ActiveMediaItem);
            ShowSlides();
            
        }


        public void ShowSlides()
        {


            // Get Slide collection object 
            slides = ActiveMediaItem.PPTPresentation.Slides;
            
            // Get Slide count 
            slidescount = slides.Count;
            
            // Get current selected slide  
            try
            {
                // Get selected slide object in normal view                  
                //if (lstSlides.SelectedItems.Count > 0)
                //{
                //    slide = presentation.Slides.FindBySlideID(lstSlides.SelectedItems[0].ImageIndex);
                //}
                //else
                //{
                //    slide = presentation.Slides.FindBySlideID(lstSlides.SelectedItems[0].ImageIndex);
                //}
                IntPtr screenClassWnd = (IntPtr)0;
                IntPtr x = (IntPtr)0;

                panel1.Controls.Add(PPTMediaApplication.PPTApplication as Control);

                ActiveMediaItem.PPTPresentation.SlideShowSettings.StartingSlide = 1;
                ActiveMediaItem.PPTPresentation.SlideShowSettings.EndingSlide = slides.Count;
                //panel1.Dock = DockStyle.Bottom;
                PPTMediaApplication.PPTApplication.Height = panel1.Height;
                ActiveMediaItem.PPTPresentation.SlideShowSettings.ShowType = PPt.PpSlideShowType.ppShowTypeWindow;
                ActiveMediaItem.PPTPresentation.SlideShowSettings.Application.Width = panel1.Height;
                ActiveMediaItem.PPTPresentation.SlideShowSettings.Application.Height = panel1.Width;

                try { PPTMediaApplication.PPTApplication.Visible = MsoTriState.msoFalse; }
                catch { }
                ActiveMediaItem.PPTPresentation.SlideShowSettings.ShowScrollbar = MsoTriState.msoFalse;
                PPt.SlideShowWindow sw = ActiveMediaItem.PPTPresentation.SlideShowSettings.Run();
            
                IntPtr pptptr = (IntPtr)sw.HWND;
                SetParent(pptptr, panel1.Handle);

            }
            catch
            {
            //    // Get selected slide object in reading view 
            //    slide = pptApplication.SlideShowWindows[1].View.Slide;

            }
        }




    }

}
