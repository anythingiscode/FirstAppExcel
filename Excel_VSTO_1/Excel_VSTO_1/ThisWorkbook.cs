using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Excel_VSTO_1
{
    public partial class ThisWorkbook
    {
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            MessageBox.Show("Hola !");

            var hoja1 = Globals.Feuil1;
            var hojaActiva = (Excel.Worksheet)Globals.ThisWorkbook.ActiveSheet;
            //var hojaActiva = (Excel.Worksheet)Globals.ThisWorkbook.ActiveSheet;

            hoja1.Range["A1"].Value = "Variable hoja1";

            hojaActiva.Range["A1"].Value = "Var Hoja Activa";
            //hojaActiva.Range["A1"].Value = "Variable hojaActiva";

        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            //MessageBox.Show("Adios !");
        }

        #region Code généré par le Concepteur VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
