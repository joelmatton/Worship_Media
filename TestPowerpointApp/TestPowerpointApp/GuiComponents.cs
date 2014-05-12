using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows;


    namespace MediaGuiComponents
    {
        public class DialogFormCalls
        {
            public string GetPowerpointFile(System.Windows.Forms.IWin32Window parentForm)
            {   
                System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
                ofd.Filter = "PowerPoint Files|*.pptx;*.pptm;*.ppt" +
                             "|All Files|*.*";
                ofd.CheckFileExists = true;
                ofd.Multiselect = false;
               System.Windows.Forms.DialogResult result = ofd.ShowDialog(parentForm);
               if (result == System.Windows.Forms.DialogResult.OK) // Test result.
                {
                    return ofd.FileName;
                }
                else
                {
                    return "";
                }
    
            }
            
        }
    
    }
