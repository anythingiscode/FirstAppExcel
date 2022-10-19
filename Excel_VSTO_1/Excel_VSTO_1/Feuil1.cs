using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using static Excel_VSTO_1.Class1;

namespace Excel_VSTO_1
{
    public partial class Feuil1
    {
        private void Feuil1_Startup(object sender, System.EventArgs e)
        {
        }

        private void Feuil1_Shutdown(object sender, System.EventArgs e)
        {
        }


        private void ButtonFrm_Click(object sender, EventArgs e)
        {
            Form miFormulario = new frmTest();
            miFormulario.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Class1.Tester();
            //Class1.Tester2();

            //double resultado = Class1.Suma(9, 6);
            double resultado = Suma(3, 8);

            MessageBox.Show($"El resultado es : {resultado}");
        }


        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);

        }
    }
}
