namespace TestPowerpointApp
{
    partial class frmTestPowerPoint
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {

            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnFirst = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnPrevious = new System.Windows.Forms.Button();
            this.btnLast = new System.Windows.Forms.Button();
            this.btnOpenPptDoc = new System.Windows.Forms.Button();
            this.txtScreens = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnClosePPT = new System.Windows.Forms.Button();
            this.btnPPTDll = new System.Windows.Forms.Button();
            this.btnNextPPTdll = new System.Windows.Forms.Button();
            this.btnPrevPPTdll = new System.Windows.Forms.Button();
            this.MediaStrip = new FilmStrip.ImageStrip();
            this.SuspendLayout();
            // 
            // btnFirst
            // 
            this.btnFirst.Location = new System.Drawing.Point(12, 54);
            this.btnFirst.Name = "btnFirst";
            this.btnFirst.Size = new System.Drawing.Size(75, 23);
            this.btnFirst.TabIndex = 1;
            this.btnFirst.Text = "First Page";
            this.btnFirst.UseVisualStyleBackColor = true;
            this.btnFirst.Click += new System.EventHandler(this.btnFirst_Click);
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(12, 83);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(75, 23);
            this.btnNext.TabIndex = 2;
            this.btnNext.Text = "Next Page";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnPrevious
            // 
            this.btnPrevious.Location = new System.Drawing.Point(93, 84);
            this.btnPrevious.Name = "btnPrevious";
            this.btnPrevious.Size = new System.Drawing.Size(75, 23);
            this.btnPrevious.TabIndex = 3;
            this.btnPrevious.Text = "Previous Page";
            this.btnPrevious.UseVisualStyleBackColor = true;
            this.btnPrevious.Click += new System.EventHandler(this.btnPrevious_Click);
            // 
            // btnLast
            // 
            this.btnLast.Location = new System.Drawing.Point(93, 55);
            this.btnLast.Name = "btnLast";
            this.btnLast.Size = new System.Drawing.Size(75, 23);
            this.btnLast.TabIndex = 4;
            this.btnLast.Text = "Last Page";
            this.btnLast.UseVisualStyleBackColor = true;
            this.btnLast.Click += new System.EventHandler(this.btnLast_Click);
            // 
            // btnOpenPptDoc
            // 
            this.btnOpenPptDoc.Location = new System.Drawing.Point(12, 12);
            this.btnOpenPptDoc.Name = "btnOpenPptDoc";
            this.btnOpenPptDoc.Size = new System.Drawing.Size(192, 23);
            this.btnOpenPptDoc.TabIndex = 5;
            this.btnOpenPptDoc.Text = "open a ppt";
            this.btnOpenPptDoc.UseVisualStyleBackColor = true;
            this.btnOpenPptDoc.Click += new System.EventHandler(this.btnOpenPptDoc_Click);
            // 
            // txtScreens
            // 
            this.txtScreens.Location = new System.Drawing.Point(12, 272);
            this.txtScreens.Multiline = true;
            this.txtScreens.Name = "txtScreens";
            this.txtScreens.Size = new System.Drawing.Size(390, 146);
            this.txtScreens.TabIndex = 6;
            // 
            // panel1
            // 
            this.panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.panel1.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel1.Location = new System.Drawing.Point(635, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(843, 825);
            this.panel1.TabIndex = 7;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // btnClosePPT
            // 
            this.btnClosePPT.Location = new System.Drawing.Point(12, 141);
            this.btnClosePPT.Name = "btnClosePPT";
            this.btnClosePPT.Size = new System.Drawing.Size(75, 23);
            this.btnClosePPT.TabIndex = 8;
            this.btnClosePPT.Text = "Close Ppt";
            this.btnClosePPT.UseVisualStyleBackColor = true;
            this.btnClosePPT.Click += new System.EventHandler(this.btnClosePPT_Click);
            // 
            // btnPPTDll
            // 
            this.btnPPTDll.Location = new System.Drawing.Point(361, 12);
            this.btnPPTDll.Name = "btnPPTDll";
            this.btnPPTDll.Size = new System.Drawing.Size(249, 23);
            this.btnPPTDll.TabIndex = 10;
            this.btnPPTDll.Text = "open a ppt from dll";
            this.btnPPTDll.UseVisualStyleBackColor = true;
            this.btnPPTDll.Click += new System.EventHandler(this.btnPPTDll_Click);
            // 
            // btnNextPPTdll
            // 
            this.btnNextPPTdll.Location = new System.Drawing.Point(478, 41);
            this.btnNextPPTdll.Name = "btnNextPPTdll";
            this.btnNextPPTdll.Size = new System.Drawing.Size(103, 23);
            this.btnNextPPTdll.TabIndex = 11;
            this.btnNextPPTdll.Text = "Select Next PPT";
            this.btnNextPPTdll.UseVisualStyleBackColor = true;
            this.btnNextPPTdll.Click += new System.EventHandler(this.btnNextPPTdll_Click);
            // 
            // btnPrevPPTdll
            // 
            this.btnPrevPPTdll.Location = new System.Drawing.Point(361, 41);
            this.btnPrevPPTdll.Name = "btnPrevPPTdll";
            this.btnPrevPPTdll.Size = new System.Drawing.Size(111, 23);
            this.btnPrevPPTdll.TabIndex = 12;
            this.btnPrevPPTdll.Text = "Select Prev PPT";
            this.btnPrevPPTdll.UseVisualStyleBackColor = true;
            this.btnPrevPPTdll.Click += new System.EventHandler(this.btnPrevPPTdll_Click);
            // 
            // MediaStrip
            // 
            this.MediaStrip.Location = new System.Drawing.Point(12, 439);
            this.MediaStrip.Name = "MediaStrip";
            this.MediaStrip.Size = new System.Drawing.Size(390, 136);
            this.MediaStrip.TabIndex = 13;
            this.MediaStrip.OnImageClicked += new System.EventHandler(this.MediaStrip_OnImageClicked);
            // 
            // frmTestPowerPoint
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1478, 825);
            this.Controls.Add(this.MediaStrip);
            this.Controls.Add(this.btnPrevPPTdll);
            this.Controls.Add(this.btnNextPPTdll);
            this.Controls.Add(this.btnPPTDll);
            this.Controls.Add(this.btnClosePPT);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.txtScreens);
            this.Controls.Add(this.btnOpenPptDoc);
            this.Controls.Add(this.btnLast);
            this.Controls.Add(this.btnPrevious);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.btnFirst);
            this.Name = "frmTestPowerPoint";
            this.Text = "PowerPointTester";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmTestPowerPoint_FormClosing);
            this.Load += new System.EventHandler(this.frmTestPowerPoint_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnFirst;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnPrevious;
        private System.Windows.Forms.Button btnLast;
        private System.Windows.Forms.Button btnOpenPptDoc;
        private System.Windows.Forms.TextBox txtScreens;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnClosePPT;
        private System.Windows.Forms.Button btnPPTDll;
        private System.Windows.Forms.Button btnNextPPTdll;
        private System.Windows.Forms.Button btnPrevPPTdll;
        private FilmStrip.ImageStrip MediaStrip;
    }
}

