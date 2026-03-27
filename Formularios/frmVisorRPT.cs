using CrystalDecisions.Windows.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CertificadosRetencion.Formularios
{
    public partial class frmVisorRPT : Form
    {
        public CrystalReportViewer crystalReportViewer1;
        public frmVisorRPT()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;

            crystalReportViewer1 = new CrystalReportViewer();
            crystalReportViewer1.Dock = DockStyle.Fill;
            crystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None;

            this.Controls.Add(crystalReportViewer1);
        }

        private void frmVisorRPT_Load(object sender, EventArgs e)
        {

        }
    }
}
