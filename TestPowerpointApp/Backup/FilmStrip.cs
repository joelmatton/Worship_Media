using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FilmStrip
{
	/// <summary>
	/// Summary description for FilmStrip.
	/// </summary>
	public class FilmStrip : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Panel FilmStripPanel;
		
		public event System.EventHandler OnImageClicked;
		
		int pictureBoxNumber=0,pictureBoxOffset=5;
		int width=0,height=0;
		string[] folderparts;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FilmStrip()
		{
			
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			
			// TODO: Add any initialization after the InitializeComponent call

		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}


		// write code for resize event of the control to resize the pictureBoxes too;
 


		


		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.FilmStripPanel = new System.Windows.Forms.Panel();
			this.SuspendLayout();
			// 
			// FilmStripPanel
			// 
			this.FilmStripPanel.AutoScroll = true;
			this.FilmStripPanel.BackColor = System.Drawing.Color.White;
			this.FilmStripPanel.Location = new System.Drawing.Point(0, 0);
			this.FilmStripPanel.Name = "FilmStripPanel";
			this.FilmStripPanel.Size = new System.Drawing.Size(384, 136);
			this.FilmStripPanel.TabIndex = 0;
			this.FilmStripPanel.Validating += new System.ComponentModel.CancelEventHandler(this.FilmStripPanel_Validating);
			this.FilmStripPanel.Paint += new System.Windows.Forms.PaintEventHandler(this.FilmStripPanel_Paint);
			// 
			// FilmStrip
			// 
			this.Controls.Add(this.FilmStripPanel);
			this.Name = "FilmStrip";
			this.Size = new System.Drawing.Size(384, 136);
			this.Resize += new System.EventHandler(this.FilmStrip_Resize);
			this.ResumeLayout(false);

		}
		#endregion

		protected void Image_Click(object sender,EventArgs e)
		{
			int PictureBoxNumber=0;
			PictureBox pb= (PictureBox)sender;
			PictureBoxNumber =Convert.ToInt32(pb.Name);
			if(OnImageClicked!=null)
			{
				OnImageClicked((object)pb.Name,e);
			}

		}

		public bool RemoveAllPictures()
		{
			ControlCollection controlCollection = this.FilmStripPanel.Controls;
			Control pb;
			for(int i=pictureBoxNumber-1;i>-1;i--)
			{
				pb = (Control)controlCollection[i*2];
				controlCollection.Remove(pb);
				pb = (Control)controlCollection[i*2];
				controlCollection.Remove(pb);
			}
			pictureBoxNumber = 0;
			pictureBoxOffset=5;
			return true;
		}

		public bool AddPicture(string ImagePath)
		{

			Bitmap bm = new Bitmap(ImagePath);
			Label lb =new Label();

			width=0;height=0;
			folderparts= ImagePath.Split('\\'); 
			PictureBox pb = new PictureBox();
			pb.Name = pictureBoxNumber .ToString();
			pb.BorderStyle = BorderStyle.FixedSingle;
			height=this.Height-50;
			width = height;//Convert.ToInt32(height*0.70);
			pb.Size = new Size(width,height);
			pb.Location = new Point(pictureBoxOffset,5);
			lb.Text = folderparts[folderparts.GetLength(0)-1];
			lb.Size = new Size(width,15);
			lb.Location = new Point(pictureBoxOffset,height+2);
			pb.SizeMode = PictureBoxSizeMode.CenterImage;
			if(bm.Height>bm.Width)
			{
				pb.Image = (Image)bm.GetThumbnailImage(Convert.ToInt32(((float)height/(float)bm.Height)*bm.Width),height,null,IntPtr.Zero);
			}
			else
			{
				pb.Image = (Image)bm.GetThumbnailImage(width,Convert.ToInt32(((float)width/(float)bm.Width)*bm.Height),null,IntPtr.Zero);
			}
			pb.Click +=new EventHandler(Image_Click);
			pictureBoxOffset = pictureBoxOffset + width + 21;
			this.FilmStripPanel.Controls.Add(pb);
			this.FilmStripPanel.Controls.Add(lb);
			pictureBoxNumber++;
			return true;
		}

		private void FilmStrip_Resize(object sender, System.EventArgs e)
		{
			this.FilmStripPanel.Size = new System.Drawing.Size(this.Width, this.Height);
			this.FilmStripPanel.AutoScroll = true;
		}

		private void FilmStripPanel_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
//			for(int i=0;i< pictureBoxNumber-1;i++)
//			{
//				this.FilmStripPanel.Controls[i].Invalidate();
//				this.FilmStripPanel.Controls[i].Update();
//			}
		}

		private void FilmStripPanel_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
//			for(int i=0;i< pictureBoxNumber-1;i++)
//			{
//				//this.FilmStripPanel.Controls[i].Invalidate();
//				this.FilmStripPanel.Controls[i].Update();
//			}
		
		}

	}
}
