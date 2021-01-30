using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyAuto
{
    public partial class AttributeValueDialog : Form
    {
        public AttributeValueDialog()
        {
            InitializeComponent();
            AcceptButton = buttonOk;
            CancelButton = buttonCancel;
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;
        }
        public string AttributeValue => textBoxAttValue.Text;
    }
}
