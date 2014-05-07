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

namespace TestPowerpointApp
{
    public partial class frmTestPowerPoint : Form
    {
        PPt._Application pptApplication;
        PPt.Presentation presentation;
        PPt.Presentation New_Presentation;
        PPt.Slides slides;
        int slidescount;
        PPt.Slide slide;
        int slideIndex;

        public frmTestPowerPoint()
        {
            InitializeComponent();
        }

        private void frmTestPowerPoint_Load(object sender, EventArgs e)
        {

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

                    // Get Slide collection object 
                    slides = presentation.Slides;
                    // Get Slide count 
                    slidescount = slides.Count;
                    // Get current selected slide  
                    try
                    {
                        // Get selected slide object in normal view 
                        slide = presentation.Slides.FindBySlideID(slides[1].SlideID);
                        presentation.SlideShowSettings.Run();
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
    }
}
