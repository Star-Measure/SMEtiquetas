using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SMEtiquetasBrother
{
    public partial class FormMessage : Form
    {
        public FormMessage()
        {
            InitializeComponent();
        }

        private void FormMessage_VisibleChanged(object sender, EventArgs e)
        {
            if(this.Visible == true)
            {
                labelMessage.Text = clGlobal.MessageText;
                this.Text = clGlobal.MessageWarning;
                this.Width = labelMessage.Text.Length * 13;
                this.CenterToScreen();
            }
        }
    }
}
