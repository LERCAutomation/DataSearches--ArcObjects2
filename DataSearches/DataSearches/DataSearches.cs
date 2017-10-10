using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DataSearches
{
    public class DataSearches : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public DataSearches()
        {
        }

        protected override void OnClick()
        {
            frmDataSearches myForm = new frmDataSearches();
            //myForm.Show();
            myForm.ShowDialog();
            ArcMap.Application.CurrentTool = null;
        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }

}
