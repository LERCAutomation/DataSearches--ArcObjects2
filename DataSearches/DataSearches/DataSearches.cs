// DataSearches is an ArcGIS add-in used to extract biodiversity
// and conservation area information from ArcGIS based on a radius around a feature.
//
// Copyright © 2016-2017 SxBRC, 2017-2018 TVERC
//
// This file is part of DataSearches.
//
// DataSearches is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataSearches is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with DataSearches.  If not, see <http://www.gnu.org/licenses/>.

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
            //
            //  Click simply launches the Data Searches Tool
            //
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
