using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ComAddinTest
{
    [Guid("31036541-2B6E-4D70-806A-1C60396AEE04")]
    [ProgId("ComAddinTest.IDEForm")]
    [ComVisible(true)]
    public partial class IDEForm : UserControl
    {
        public IDEForm()
        {
            InitializeComponent();
        }
    }
}
