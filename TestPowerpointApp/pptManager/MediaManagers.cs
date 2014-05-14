using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using MOIP = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Drawing;

namespace MediaManagers
{

    public class PowerPointApplication : IDisposable
    {
        public MOIP.Application PPTApplication;
        MOIP.Presentation PrimaryPresentation;

        int PictureImportIterator = 1000;

        int ActiveMediaItem = -1;

       // Type pptApplicationType;
        bool disposed = false;
        List<clsMediaListItem> MediaQueue = new List<clsMediaListItem>();

        public PowerPointApplication()
        {
          //  pptApplicationType = Type.GetTypeFromProgID("PowerPoint.Application");
           // _pptApplication = Activator.CreateInstance(pptApplicationType);
            PPTApplication = new MOIP.Application();
            foreach (MOIP.Presentation x in PPTApplication.Presentations)
            {
                try { x.Close(); }
                catch { };
            }

        }



        // Public implementation of Dispose pattern callable by consumers. 
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // Protected implementation of Dispose pattern. 
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here. 
                //
            }

            try
            {
                if (PPTApplication != null)
                {
                    foreach (MOIP.Presentation x in PPTApplication.Presentations)
                    {
                        try { x.Close(); }
                        catch { };
                    }
                    PPTApplication.Quit();
                }
            }
            catch { }
            

            disposed = true;
        }

        public int AddPresentationFromFile(string filename) 
        {
            //interop 15
//            return _pptApplication.OpenPresentation(filename);

            //interop 14

            MOIP.Presentation x = PPTApplication.Presentations.Open(filename, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                //create image list of slides
            PowerPointUtilites ppu = new PowerPointUtilites();
            List<clsSlidePictureItem> PictureList = ppu.GetSlideImages(x, PictureImportIterator * (MediaQueue.Count+1));

            MediaQueue.Add(new clsMediaListItem(MediaTypeEnum.PowerPointFile, filename, MediaQueue.Count, PictureList,x));



            return PPTApplication.Presentations.Count;
        }

        public int AddBlankPresentation()
        {
            MOIP.Presentation x = PPTApplication.Presentations.Add(MsoTriState.msoFalse);
            List<clsSlidePictureItem> PictureList = new List<clsSlidePictureItem>();
            MediaQueue.Add(new clsMediaListItem(MediaTypeEnum.PowerPointDynamic, "", MediaQueue.Count, PictureList,x));
            return PPTApplication.Presentations.Count;

        }

        public void ClosePresentation(int PresentationID)
        {
            try
            {
                PPTApplication.Presentations[PresentationID].Close();
            }
            catch { }
               
                
        }


        public clsMediaListItem GetNextMediaListItem()
        {
            if (ActiveMediaItem + 1 <= MediaQueue.Count-1)
            {
                ActiveMediaItem += 1;
                return MediaQueue[ActiveMediaItem];
            }
            else
            { return MediaQueue[MediaQueue.Count-1]; }              
        }

        public clsMediaListItem GetPrevMediaListItem()
        {
            if (ActiveMediaItem -1 > 0)
            {
                ActiveMediaItem -= 1;
                return MediaQueue[ActiveMediaItem-1];
            }
            else
            { return MediaQueue[0]; }
        }

        public void DisposeEverything()
        {
            if (PPTApplication != null)
            {
                foreach (MOIP.Presentation x in PPTApplication.Presentations)
                {
                    try { x.Close(); }
                    catch { };
                }
                PPTApplication.Quit();
            }
        
        }

    }

    public enum MediaTypeEnum
    {
        PowerPointFile=1,
        PowerPointDynamic=2,
        Movie=3,
        Picture=4
    }


    public class clsMediaListItem
    {
        public bool isActive;
        public MediaTypeEnum mediaType;
        public string filename;
        public int ID;
        public int PresentationID;
        public List<clsSlidePictureItem> SlideImages;
        public MOIP.Presentation PPTPresentation;


        public clsMediaListItem(MediaTypeEnum m, string filename, int MediaID, List<clsSlidePictureItem> slideImages, MOIP.Presentation xPres)
        {
            this.ID = MediaID;
            this.mediaType = m;
            this.filename = filename;
            this.SlideImages = slideImages;
            this.PPTPresentation = xPres;
        }
    }


    public class clsSlidePictureItem
    {
        public string filename;
        public int ID;  //slide id
        public Image SlideImage;

        public clsSlidePictureItem(string filename, int slideID)
        {
            this.ID = slideID;
            this.filename = filename;
        }
        public clsSlidePictureItem(string filename, int slideID, Image slideImage)
        {
            this.ID = slideID;
            this.filename = filename;
            this.SlideImage = slideImage;
        }

    }

    public class PowerPointUtilites{


        public List<clsSlidePictureItem> GetSlideImages(MOIP.Presentation picPresentation, int PictureImportIterator)
            {
                PictureImportIterator += 1000;
                List<clsSlidePictureItem> PictureList = new List<clsSlidePictureItem>();

                if (picPresentation == null)
                    return PictureList;

                foreach (MOIP.Slide slide in picPresentation.Slides)
                {
                    PictureImportIterator += 1;
                    var fileName = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        string.Format("Slide_" + PictureImportIterator + "{0:00}.jpg", slide.SlideNumber));

                     slide.Export(fileName, "JPG", 1024, 768);
                     Image slImage= Image.FromFile(fileName);
                     PictureList.Add(new clsSlidePictureItem(fileName, slide.SlideID, slImage));
                }

                return PictureList;
            }

    }



}
