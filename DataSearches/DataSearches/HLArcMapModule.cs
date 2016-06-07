// DataSearches is an ArcGIS add-in used to extract biodiversity
// and conservation area information from ArcGIS based on a radius around a feature.
//
// Copyright © 2016 Sussex Biodiversity Record Centre
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
// along with DataSelector.  If not, see <http://www.gnu.org/licenses/>.


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;

using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Desktop.AddIns;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.GeoDatabaseUI;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.AnalysisTools;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.DataSourcesOleDB;
using ESRI.ArcGIS.Display;

using ESRI.ArcGIS.Catalog;
using ESRI.ArcGIS.CatalogUI;

using HLFileFunctions;

namespace HLArcMapModule
{
    class ArcMapFunctions
    {
        #region Constructor
        private IApplication thisApplication;
        private FileFunctions myFileFuncs;
        // Class constructor.
        public ArcMapFunctions(IApplication theApplication)
        {
            // Set the application for the class to work with.
            // Note the application can be got at from a command / tool by using
            // IApplication pApp = ArcMap.Application - then pass pApp as an argument.
            this.thisApplication = theApplication;
            myFileFuncs = new FileFunctions();
        }
        #endregion

        public IMxDocument GetIMXDocument()
        {
            ESRI.ArcGIS.ArcMapUI.IMxDocument mxDocument = ((ESRI.ArcGIS.ArcMapUI.IMxDocument)(thisApplication.Document));
            return mxDocument;
        }

        public void UpdateTOC()
        {
            IMxDocument mxDoc = GetIMXDocument();
            mxDoc.UpdateContents();
        }

        public bool SaveMXD()
        {
            IMxDocument mxDoc = GetIMXDocument();
            IMapDocument pDoc = (IMapDocument)mxDoc;
            try
            {
                pDoc.Save(true, true);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public IActiveView GetActiveView()
        {
            IMxDocument mxDoc = GetIMXDocument();
            return mxDoc.ActiveView;
        }

        public ESRI.ArcGIS.Carto.IMap GetMap()
        {
            if (thisApplication == null)
            {
                return null;
            }
            ESRI.ArcGIS.ArcMapUI.IMxDocument mxDocument = ((ESRI.ArcGIS.ArcMapUI.IMxDocument)(thisApplication.Document)); // Explicit Cast
            ESRI.ArcGIS.Carto.IActiveView activeView = mxDocument.ActiveView;
            ESRI.ArcGIS.Carto.IMap map = activeView.FocusMap;

            return map;
        }

        public void RefreshTOC()
        {
            IMxDocument theDoc = GetIMXDocument();
            theDoc.CurrentContentsView.Refresh(null);
        }

        public IWorkspaceFactory GetWorkspaceFactory(string aFilePath, bool aTextFile = false, bool Messages = false)
        {
            // This function decides what type of feature workspace factory would be best for this file.
            // it is up to the user to decide whether the file path and file names exist (or should exist).

            IWorkspaceFactory pWSF;
            // What type of output file it it? This defines what kind of workspace factory.
            if (aFilePath.Substring(aFilePath.Length - 4, 4) == ".gdb")
            {
                // It is a file geodatabase file.
                pWSF = new FileGDBWorkspaceFactory();
            }
            else if (aFilePath.Substring(aFilePath.Length - 4, 4) == ".mdb")
            {
                // Personal geodatabase.
                pWSF = new AccessWorkspaceFactory();
            }
            else if (aFilePath.Substring(aFilePath.Length - 4, 4) == ".sde")
            {
                // ArcSDE connection
                pWSF = new SdeWorkspaceFactory();
            }
            else if (aTextFile == true)
            {
                // Text file
                pWSF = new TextFileWorkspaceFactory();
            }
            else
            {
                pWSF = new ShapefileWorkspaceFactory();
            }
            return pWSF;
        }


        #region FeatureclassExists
        public bool FeatureclassExists(string aFilePath, string aDatasetName)
        {
            
            if (aDatasetName.Substring(aDatasetName.Length - 4, 1) == ".")
            {
                // it's a file.
                if (myFileFuncs.FileExists(aFilePath + @"\" + aDatasetName))
                    return true;
                else
                    return false;
            }
            else if (aFilePath.Substring(aFilePath.Length - 3, 3) == "sde")
            {
                // It's an SDE class
                // Not handled. We know the table exists.
                return true;
            }
            else // it is a geodatabase class.
            {
                IWorkspaceFactory pWSF = GetWorkspaceFactory(aFilePath);
                IWorkspace2 pWS = (IWorkspace2)pWSF.OpenFromFile(aFilePath, 0);
                bool blReturn = false;
                if (pWS.get_NameExists(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTFeatureClass, aDatasetName))
                    blReturn = true;
                Marshal.ReleaseComObject(pWS);
                return blReturn;


            }
        }

        public bool FeatureclassExists(string aFullPath)
        {
            return FeatureclassExists(myFileFuncs.GetDirectoryName(aFullPath), myFileFuncs.GetFileName(aFullPath));
        }
        #endregion

        #region GetFeatureClass
        public IFeatureClass GetFeatureClass(string aFilePath, string aDatasetName, bool Messages = false)
        // This is incredibly quick.
        {
            // Check input first.
            string aTestPath = aFilePath;
            if (aFilePath.Contains(".sde"))
            {
                aTestPath = myFileFuncs.GetDirectoryName(aFilePath);
            }
            if (myFileFuncs.DirExists(aTestPath) == false || aDatasetName == null)
            {
                if (Messages) MessageBox.Show("Please provide valid input", "Get Featureclass");
                return null;
            }
            

            IWorkspaceFactory pWSF = GetWorkspaceFactory(aFilePath);
            IFeatureWorkspace pWS = (IFeatureWorkspace)pWSF.OpenFromFile(aFilePath, 0);
            if (FeatureclassExists(aFilePath, aDatasetName))
            {
                IFeatureClass pFC = pWS.OpenFeatureClass(aDatasetName);
                Marshal.ReleaseComObject(pWS);
                pWS = null;
                pWSF = null;
                GC.Collect();
                return pFC;
            }
            else
            {
                if (Messages) MessageBox.Show("The file " + aDatasetName + " doesn't exist in this location", "Open Feature Class from Disk");
                Marshal.ReleaseComObject(pWS);
                pWS = null;
                pWSF = null;
                GC.Collect();
                return null;
            }

        }


        public IFeatureClass GetFeatureClass(string aFullPath, bool Messages = false)
        {
            string aFilePath = myFileFuncs.GetDirectoryName(aFullPath);
            string aDatasetName = myFileFuncs.GetFileName(aFullPath);
            IFeatureClass pFC = GetFeatureClass(aFilePath, aDatasetName, Messages);
            return pFC;
        }

        #endregion

        public IFeatureLayer GetFeatureLayerFromString(string aFeatureClassName, bool Messages = false)
        {
            // as far as I can see this does not work for geodatabase files.
            // firstly get the Feature Class
            // Does it exist?
            if (!myFileFuncs.FileExists(aFeatureClassName))
            {
                if (Messages)
                {
                    MessageBox.Show("The featureclass " + aFeatureClassName + " does not exist");
                }
                return null;
            }
            string aFilePath = myFileFuncs.GetDirectoryName(aFeatureClassName);
            string aFCName = myFileFuncs.GetFileName(aFeatureClassName);

            IFeatureClass myFC = GetFeatureClass(aFilePath, aFCName);
            if (myFC == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot open featureclass " + aFeatureClassName);
                }
                return null;
            }

            // Now get the Feature Layer from this.
            FeatureLayer pFL = new FeatureLayer();
            pFL.FeatureClass = myFC;
            pFL.Name = myFC.AliasName;
            return pFL;
        }

        public ILayer GetLayer(string aName, bool Messages = false)
        {
            // Gets existing layer in map.
            // Check there is input.
           if (aName == null)
           {
               if (Messages)
               {
                   MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
               }
               return null;
            }
        
            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages)
                {
                    MessageBox.Show("No map found", "Find Layer By Name");
                }
                return null;
            }
            IEnumLayer pLayers = pMap.Layers;
            Boolean blFoundit = false;
            ILayer pTargetLayer = null;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while ((pLayer != null) && !blFoundit)
            {
                if (!(pLayer is ICompositeLayer))
                {
                    if (pLayer.Name == aName)
                    {
                        pTargetLayer = pLayer;
                        blFoundit = true;
                    }
                }
                pLayer = pLayers.Next();
            }

            if (pTargetLayer == null)
            {
                if (Messages) MessageBox.Show("The layer " + aName + " doesn't exist", "Find Layer");
                return null;
            }
            return pTargetLayer;
        }

        public bool FieldExists(string aFilePath, string aDatasetName, string aFieldName, bool Messages = false)
        {
            // This function returns true if a field (or a field alias) exists, false if it doesn (or the dataset doesn't)
            IFeatureClass myFC = GetFeatureClass(aFilePath, aDatasetName, Messages);
            ITable myTab;
            if (myFC == null)
            {
                myTab = GetTable(aFilePath, aDatasetName, Messages);

                if (myTab == null)
                {
                    if (Messages)
                        MessageBox.Show("Cannot check for field in dataset " + aFilePath + @"\" + aDatasetName + ". Dataset does not exist");
                    return false; // Dataset doesn't exist.
                }
            }
            else
            {
                myTab = (ITable)myFC;
            }

            int aTest;
            IFields theFields = myTab.Fields;
            aTest = theFields.FindField(aFieldName);
            if (aTest == -1)
            {
                aTest = theFields.FindFieldByAliasName(aFieldName);
            }

            if (aTest == -1) return false;
            return true;
        }

        public bool FieldExists(IFeatureClass aFeatureClass, string aFieldName, bool Messages = false)
        {

            int aTest;
            IFields theFields = aFeatureClass.Fields;
            aTest = theFields.FindField(aFieldName);
            if (aTest == -1)
            {
                aTest = theFields.FindFieldByAliasName(aFieldName);
            }

            if (aTest == -1) return false;
            return true;
        }

        public bool FieldExists(string aFeatureClass, string aFieldName, bool Messages = false)
        {
            string aFilePath = myFileFuncs.GetDirectoryName(aFeatureClass);
            string aDatasetName = myFileFuncs.GetFileName(aFeatureClass);
            return FieldExists(aFilePath, aDatasetName, aFieldName, Messages);
        }

        public bool FieldExists(ILayer aLayer, string aFieldName, bool Messages = false)
        {
            IFeatureLayer pFL = null;
            try
            {
                pFL = (IFeatureLayer)aLayer;
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("The layer is not a feature layer");
                return false;
            }
            IFeatureClass pFC = pFL.FeatureClass;
            return FieldExists(pFC, aFieldName);
        }

        public bool FieldIsNumeric(string aFeatureClass, string aFieldName, bool Messages = false)
        {
            // Check the obvious.
            if (!FeatureclassExists(aFeatureClass))
            {
                if (Messages)
                    MessageBox.Show("The featureclass " + aFeatureClass + " doesn't exist");
                return false;
            }

            if (!FieldExists(aFeatureClass, aFieldName))
            {
                if (Messages)
                    MessageBox.Show("The field " + aFieldName + " does not exist in featureclass " + aFeatureClass);
                return false;
            }

            IField pField = GetFCField(aFeatureClass, aFieldName);
            if (pField == null)
            {
                if (Messages) MessageBox.Show("The field " + aFieldName + " does not exist in this layer", "Field Is Numeric");
                return false;
            }

            if (pField.Type == esriFieldType.esriFieldTypeDouble |
                pField.Type == esriFieldType.esriFieldTypeInteger |
                pField.Type == esriFieldType.esriFieldTypeSingle |
                pField.Type == esriFieldType.esriFieldTypeSmallInteger) return true;
            
            return false;

        }

        public bool AddField(IFeatureClass aFeatureClass, string aFieldName, esriFieldType aFieldType, int aLength, bool Messages = false)
        {
            // Validate input.
            if (aFeatureClass == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Please pass a valid feature class", "Add Field");
                }
                return false;
            }
            if (aLength <= 0)
            {
                if (Messages)
                {
                    MessageBox.Show("Please enter a valid field length", "Add Field");
                }
                return false;
            }
            IFields pFields = aFeatureClass.Fields;
            int i = pFields.FindField(aFieldName);
            if (i > -1)
            {
                if (Messages)
                {
                    MessageBox.Show("This field already exists", "Add Field");
                }
                return false;
            }

            ESRI.ArcGIS.Geodatabase.Field aNewField = new ESRI.ArcGIS.Geodatabase.Field();
            IFieldEdit anEdit = (IFieldEdit)aNewField;

            anEdit.AliasName_2 = aFieldName;
            anEdit.Name_2 = aFieldName;
            anEdit.Type_2 = aFieldType;
            anEdit.Length_2 = aLength;

            aFeatureClass.AddField(aNewField);
            return true;
        }

        public bool AddField(string aFeatureClass, string aFieldName, esriFieldType aFieldType, int aLength, bool Messages = false)
        {
            IFeatureClass pFC = GetFeatureClass(aFeatureClass, Messages);
            return AddField(pFC, aFieldName, aFieldType, aLength, Messages);
        }

        public bool AddLayerField(string aLayer, string aFieldName, esriFieldType aFieldType, int aLength, bool Messages = false)
        {
            if (!LayerExists(aLayer))
            {
                if (Messages)
                    MessageBox.Show("The layer " + aLayer + " could not be found in the map");
                return false;
            }

            ILayer pLayer = GetLayer(aLayer);
            IFeatureLayer pFL;
            try
            {
                pFL = (IFeatureLayer)pLayer;
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("Layer " + aLayer + " is not a feature layer.");
                return false;
            }

            IFeatureClass pFC = pFL.FeatureClass;
            AddField(pFC, aFieldName, aFieldType, aLength, Messages);

            return true;
        }

        public bool DeleteLayerField(string aLayer, string aFieldName, bool Messages = false)
        {
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();


            // Populate the variant array with parameter values.
            parameters.Add(aLayer);
            parameters.Add(aFieldName);

            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("DeleteField_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);
                // Wait for 1 second.
                if (Messages)
                {
                    MessageBox.Show("Process complete");
                }
                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                gp = null;
                return false;
            }
        }

        public bool AddLayerFromFClass(IFeatureClass theFeatureClass, bool Messages = false)
        {
            // Check we have input
            if (theFeatureClass == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Please pass a feature class", "Add Layer From Feature Class");
                }
                return false;
            }
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages)
                {
                    MessageBox.Show("No map found", "Add Layer From Feature Class");
                }
                return false;
            }
            FeatureLayer pFL = new FeatureLayer();
            pFL.FeatureClass = theFeatureClass;
            pFL.Name = theFeatureClass.AliasName;
            pMap.AddLayer(pFL);

            return true;
        }

        public bool AddFeatureLayerFromString(string aFeatureClassName, bool Messages = false)
        {
            // firstly get the Feature Class
            // Does it exist?
            if (!myFileFuncs.FileExists(aFeatureClassName))
            {
                if (Messages)
                {
                    MessageBox.Show("The featureclass " + aFeatureClassName + " does not exist");
                }
                return false;
            }
            string aFilePath = myFileFuncs.GetDirectoryName(aFeatureClassName);
            string aFCName = myFileFuncs.GetFileName(aFeatureClassName);

            IFeatureClass myFC = GetFeatureClass(aFilePath, aFCName);
            if (myFC == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot open featureclass " + aFeatureClassName);
                }
                return false;
            }

            // Now add it to the view.
            bool blResult = AddLayerFromFClass(myFC);
            if (blResult)
            {
                return true;
            }
            else
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot add featureclass " + aFeatureClassName);
                }
                return false;
            }
        }

        #region TableExists
        public bool TableExists(string aFilePath, string aDatasetName)
        {

            if (aDatasetName.Substring(aDatasetName.Length - 4, 1) == ".")
            {
                // it's a file.
                if (myFileFuncs.FileExists(aFilePath + @"\" + aDatasetName))
                    return true;
                else
                    return false;
            }
            else if (aFilePath.Substring(aFilePath.Length - 3, 3) == "sde")
            {
                // It's an SDE class
                // Not handled. We know the table exists.
                return true;
            }
            else // it is a geodatabase class.
            {
                IWorkspaceFactory pWSF = GetWorkspaceFactory(aFilePath);
                IWorkspace2 pWS = (IWorkspace2)pWSF.OpenFromFile(aFilePath, 0);
                if (pWS.get_NameExists(ESRI.ArcGIS.Geodatabase.esriDatasetType .esriDTTable, aDatasetName))
                    return true;
                else
                    return false;
            }
        }

        public bool TableExists(string aFullPath)
        {
            return TableExists(myFileFuncs.GetDirectoryName(aFullPath), myFileFuncs.GetFileName(aFullPath));
        }
        #endregion

        #region GetTable
        public ITable GetTable(string aFilePath, string aDatasetName, bool Messages = false)
        {
            // Check input first.
            string aTestPath = aFilePath;
            if (aFilePath.Contains(".sde"))
            {
                aTestPath = myFileFuncs.GetDirectoryName(aFilePath);
            }
            if (myFileFuncs.DirExists(aTestPath) == false || aDatasetName == null)
            {
                if (Messages) MessageBox.Show("Please provide valid input", "Get Table");
                return null;
            }
            bool blText = false;
            string strExt = aDatasetName.Substring(aDatasetName.Length - 4, 4);
            if (strExt == ".txt" || strExt == ".csv" || strExt == ".tab")
            {
                blText = true;
            }

            IWorkspaceFactory pWSF = GetWorkspaceFactory(aFilePath, blText);
            IFeatureWorkspace pWS = (IFeatureWorkspace)pWSF.OpenFromFile(aFilePath, 0);
            ITable pTable = pWS.OpenTable(aDatasetName);
            if (pTable == null)
            {
                if (Messages) MessageBox.Show("The file " + aDatasetName + " doesn't exist in this location", "Open Table from Disk");
                Marshal.ReleaseComObject(pWS);
                pWSF = null;
                pWS = null;
                GC.Collect();
                return null;
            }
            Marshal.ReleaseComObject(pWS);
            pWSF = null;
            pWS = null;
            GC.Collect();
            return pTable;
        }

        public ITable GetTable(string aTableLayer, bool Messages = false)
        {
            IMap pMap = GetMap();
            IStandaloneTableCollection pColl = (IStandaloneTableCollection)pMap;
            IStandaloneTable pThisTable = null;

            for (int I = 0; I < pColl.StandaloneTableCount; I++)
            {
                pThisTable = pColl.StandaloneTable[I];
                if (pThisTable.Name == aTableLayer)
                {
                    ITable myTable = pThisTable.Table;
                    return myTable;
                }
            }
            if (Messages)
            {
                MessageBox.Show("The table layer " + aTableLayer + " could not be found in this map");
            }
            return null;
        }
        #endregion

        public bool AddTableLayerFromString(string aTableName, string aLayerName, bool Messages = false)
        {
            // firstly get the Table
            // Does it exist? // Does not work for GeoDB tables!!
            if (!myFileFuncs.FileExists(aTableName))
            {
                if (Messages)
                {
                    MessageBox.Show("The table " + aTableName + " does not exist");
                }
                return false;
            }
            string aFilePath = myFileFuncs.GetDirectoryName(aTableName);
            string aTabName = myFileFuncs.GetFileName(aTableName);

            ITable myTable = GetTable(aFilePath, aTabName);
            if (myTable == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot open table " + aTableName);
                }
                return false;
            }

            // Now add it to the view.
            bool blResult = AddLayerFromTable(myTable, aLayerName);
            if (blResult)
            {
                return true;
            }
            else
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot add table " + aTabName);
                }
                return false;
            }
        }

        public bool AddLayerFromTable(ITable theTable, string aName, bool Messages = false)
        {
            // check we have nput
            if (theTable == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Please pass a table", "Add Layer From Table");
                }
                return false;
            }
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages)
                {
                    MessageBox.Show("No map found", "Add Layer From Table");
                }
                return false;
            }
            IStandaloneTableCollection pStandaloneTableCollection = (IStandaloneTableCollection)pMap;
            IStandaloneTable pTable = new StandaloneTable();
            IMxDocument mxDoc = GetIMXDocument();

            pTable.Table = theTable;
            pTable.Name = aName;

            // Remove if already exists
            if (TableLayerExists(aName))
                RemoveStandaloneTable(aName);

            mxDoc.UpdateContents();
            
            pStandaloneTableCollection.AddStandaloneTable(pTable);
            mxDoc.UpdateContents();
            return true;
        }

        public bool TableLayerExists(string aLayerName, bool Messages = false)
        {
            // Check there is input.
            if (aLayerName == null)
            {
                if (Messages) MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return false;
            }

            // Get map, and layer names.
            IMxDocument mxDoc = GetIMXDocument();
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return false;
            }

            IStandaloneTableCollection pColl = (IStandaloneTableCollection)pMap;
            IStandaloneTable pThisTable = null;
            for (int I = 0; I < pColl.StandaloneTableCount; I++)
            {
                pThisTable = pColl.StandaloneTable[I];
                if (pThisTable.Name == aLayerName)
                {
                    return true;
                    //pColl.RemoveStandaloneTable(pThisTable);
                   // mxDoc.UpdateContents();
                    //break; // important: get out now, the index is no longer valid
                }
            }
            return false;
        }

        public bool RemoveStandaloneTable(string aTableName, bool Messages = false)
        {
            // Check there is input.
            if (aTableName == null)
            {
                if (Messages) MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return false;
            }

            // Get map, and layer names.
            IMxDocument mxDoc = GetIMXDocument();
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return false;
            }

            IStandaloneTableCollection pColl = (IStandaloneTableCollection)pMap;
            IStandaloneTable pThisTable = null;
            for (int I = 0; I < pColl.StandaloneTableCount; I++)
            {
                pThisTable = pColl.StandaloneTable[I];
                if (pThisTable.Name == aTableName)
                {
                    try
                    {
                        pColl.RemoveStandaloneTable(pThisTable);
                        mxDoc.UpdateContents();
                        return true; // important: get out now, the index is no longer valid
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }
            return false;
        }

        public bool LayerExists(string aLayerName, bool Messages = false)
        {
            // Check there is input.
            if (aLayerName == null)
            {
                if (Messages) MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return false;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return false;
            }
            IEnumLayer pLayers = pMap.Layers;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while (pLayer != null)
            {
                if (!(pLayer is IGroupLayer))
                {
                    if (pLayer.Name == aLayerName)
                    {
                        return true;
                    }

                }
                pLayer = pLayers.Next();
            }
            return false;
        }

        public bool GroupLayerExists(string aGroupLayerName, bool Messages = false)
        {
            // Check there is input.
            if (aGroupLayerName == null)
            {
                if (Messages) MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return false;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return false;
            }
            IEnumLayer pLayers = pMap.Layers;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while (pLayer != null)
            {
                if (pLayer is IGroupLayer)
                {
                    if (pLayer.Name == aGroupLayerName)
                    {
                        return true;
                    }

                }
                pLayer = pLayers.Next();
            }
            return false;
        }

        public ILayer GetGroupLayer(string aGroupLayerName, bool Messages = false)
        {
            // Check there is input.
            if (aGroupLayerName == null)
            {
                if (Messages) MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return null;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return null;
            }
            IEnumLayer pLayers = pMap.Layers;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while (pLayer != null)
            {
                if (pLayer is IGroupLayer)
                {
                    if (pLayer.Name == aGroupLayerName)
                    {
                        return pLayer;
                    }

                }
                pLayer = pLayers.Next();
            }
            return null;
        }      
        
        public bool MoveToGroupLayer(string theGroupLayerName, ILayer aLayer,  bool Messages = false)
        {
            bool blExists = false;
            IGroupLayer myGroupLayer = new GroupLayer(); 
            // Does the group layer exist?
            if (GroupLayerExists(theGroupLayerName))
            {
                myGroupLayer = (IGroupLayer)GetGroupLayer(theGroupLayerName);
                blExists = true;
            }
            else
            {
                myGroupLayer.Name = theGroupLayerName;
            }
            string theOldName = aLayer.Name;

            // Remove the original instance, then add it to the group.
            RemoveLayer(aLayer);
            myGroupLayer.Add(aLayer);
            
            if (!blExists)
            {
                // Add the layer to the map.
                IMap pMap = GetMap();
                pMap.AddLayer(myGroupLayer);
            }
            RefreshTOC();
            return true;
        }

        public bool MoveToSubGroupLayer(string theGroupLayerName, string theSubGroupLayerName, ILayer aLayer, bool Messages = false)
        {
            bool blGroupLayerExists = false;
            bool blSubGroupLayerExists = false;
            IGroupLayer myGroupLayer = new GroupLayer();
            IGroupLayer mySubGroupLayer = new GroupLayer();
            // Does the group layer exist?
            if (GroupLayerExists(theGroupLayerName))
            {
                myGroupLayer = (IGroupLayer)GetGroupLayer(theGroupLayerName);
                blGroupLayerExists = true;
            }
            else
            {
                myGroupLayer.Name = theGroupLayerName;
            }


            if (GroupLayerExists(theSubGroupLayerName))
            {
                mySubGroupLayer = (IGroupLayer)GetGroupLayer(theSubGroupLayerName);
                blSubGroupLayerExists = true;
            }
            else
            {
                mySubGroupLayer.Name = theSubGroupLayerName;
            }

            // Remove the original instance, then add it to the group.
            string theOldName = aLayer.Name; 
            RemoveLayer(aLayer);
            mySubGroupLayer.Add(aLayer);

            if (!blSubGroupLayerExists)
            {
                // Add the subgroup layer to the group layer.
                myGroupLayer.Add(mySubGroupLayer);
            }
            if (!blGroupLayerExists)
            {
                // Add the layer to the map.
                IMap pMap = GetMap();
                pMap.AddLayer(myGroupLayer);
            }
            RefreshTOC();
            return true;
        }

        #region RemoveLayer
        public bool RemoveLayer(string aLayerName, bool Messages = false)
        {
            // Check there is input.
            if (aLayerName == null)
            {
                MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return false;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return false;
            }
            IEnumLayer pLayers = pMap.Layers;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while (pLayer != null)
            {
                if (!(pLayer is IGroupLayer))
                {
                    if (pLayer.Name == aLayerName)
                    {
                        pMap.DeleteLayer(pLayer);
                        return true;
                    }

                }
                pLayer = pLayers.Next();
            }
            return false;
        }

        public bool RemoveLayer(ILayer aLayer, bool Messages = false)
        {
            // Check there is input.
            if (aLayer == null)
            {
                MessageBox.Show("Please pass a valid layer ", "Remove Layer");
                return false;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Remove Layer");
                return false;
            }
            pMap.DeleteLayer(aLayer);
            return true;
        }
        #endregion


        public string GetOutputFileName(string aFileType, string anInitialDirectory = @"C:\")
        {
            // This would be done better with a custom type but this will do for the momment.
            IGxDialog myDialog = new GxDialogClass();
            myDialog.set_StartingLocation(anInitialDirectory);
            IGxObjectFilter myFilter;


            switch (aFileType)
            {
                case "Geodatabase FC":
                    myFilter = new GxFilterFGDBFeatureClasses();
                    break;
                case "Geodatabase Table":
                    myFilter = new GxFilterFGDBTables();
                    break;
                case "Shapefile":
                    myFilter = new GxFilterShapefiles();
                    break;
                case "DBASE file":
                    myFilter = new GxFilterdBASEFiles();
                    break;
                case "Text file":
                    myFilter = new GxFilterTextFiles();
                    break;
                default:
                    myFilter = new GxFilterDatasets();
                    break;
            }

            myDialog.ObjectFilter = myFilter;
            myDialog.Title = "Save Output As...";
            myDialog.ButtonCaption = "OK";

            string strOutFile = "None";
            if (myDialog.DoModalSave(thisApplication.hWnd))
            {
                strOutFile = myDialog.FinalLocation.FullName + @"\" + myDialog.Name;
                
            }
            myDialog = null;
            return strOutFile; // "None" if user pressed exit
            
        }

        #region CopyFeatures
        public bool CopyFeatures(string InFeatureClassOrLayer, string OutFeatureClass, bool Messages = false)
        {
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = true;
            IGeoProcessorResult myresult = new GeoProcessorResultClass();
            object sev = null;

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(InFeatureClassOrLayer);
            parameters.Add(OutFeatureClass);

            // Execute the tool.
            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("CopyFeatures_management", parameters, null);
                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);
                    // Wait for 1 second.
                if (Messages)
                {
                    MessageBox.Show("Process complete");
                }
                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                if (Messages)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(gp.GetMessages(ref sev));
                }
                gp = null;
                return false;
            }
        }

        public bool CopyFeatures(string InWorkspace, string InDatasetName, string OutFeatureClass, bool Messages = false)
        {
            string inFeatureClass = InWorkspace + @"\" + InDatasetName;
            return CopyFeatures(inFeatureClass, OutFeatureClass, Messages);
        }

        public bool CopyFeatures(string InWorkspace, string InDatasetName, string OutWorkspace, string OutDatasetName, bool Messages = false)
        {
            string inFeatureClass = InWorkspace + @"\" + InDatasetName;
            string outFeatureClass = OutWorkspace + @"\" + OutDatasetName;
            return CopyFeatures(inFeatureClass, outFeatureClass, Messages);
        }
        #endregion

        public bool CopyTable(string InTable, string OutTable, bool Messages = false)
        {
            // This works absolutely fine for dbf and geodatabase but does not export to CSV.

            // Note the csv export already removes ghe geometry field; in this case it is not necessary to check again.

            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = true;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(InTable);
            parameters.Add(OutTable);

            // Execute the tool.
            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("CopyRows_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);
                    // Wait for 1 second.

                
                if (Messages)
                {
                    MessageBox.Show("Process complete");
                }
                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                gp = null;
                return false;
            }
        }

        public bool AlterFieldAliasName(string aDatasetName, string aFieldName, string theAliasName, bool Messages = false)
        {
            // This script changes the field alias of a the named field in the layer.
            IObjectClass myObject = (IObjectClass)GetFeatureClass(aDatasetName);
            IClassSchemaEdit myEdit = (IClassSchemaEdit)myObject;
            try
            {
                myEdit.AlterFieldAliasName(aFieldName, theAliasName);
                myObject = null;
                myEdit = null;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                myObject = null;
                myEdit = null;
                return false;
            }
        }

        public IField GetFCField(string InputDirectory, string FeatureclassName, string FieldName, bool Messages = false)
        {
            IFeatureClass featureClass = GetFeatureClass(InputDirectory, FeatureclassName);
            // Find the index of the requested field.
            int fieldIndex = featureClass.FindField(FieldName);

            // Get the field from the feature class's fields collection.
            if (fieldIndex > -1)
            {
                IFields fields = featureClass.Fields;
                IField field = fields.get_Field(fieldIndex);
                return field;
            }
            else
            {
                if (Messages)
                {
                    MessageBox.Show("The field " + FieldName + " was not found in the featureclass " + FeatureclassName);
                }
                return null;
            }
        }

        public IField GetFCField(string aFeatureClass, string FieldName, bool Messages = false)
        {
            string strInputDir = myFileFuncs.GetDirectoryName(aFeatureClass);
            string strInputShape = myFileFuncs.GetFileName(aFeatureClass);
            return GetFCField(strInputDir, strInputShape, FieldName, Messages);
        }

        public IField GetTableField(string TableName, string FieldName, bool Messages = false)
        {
            ITable theTable = GetTable(myFileFuncs.GetDirectoryName(TableName), myFileFuncs.GetFileName(TableName), Messages);
            int fieldIndex = theTable.FindField(FieldName);

            // Get the field from the feature class's fields collection.
            if (fieldIndex > -1)
            {
                IFields fields = theTable.Fields;
                IField field = fields.get_Field(fieldIndex);
                return field;
            }
            else
            {
                if (Messages)
                {
                    MessageBox.Show("The field " + FieldName + " was not found in the table " + myFileFuncs.GetFileName(TableName));
                }
                return null;
            }
        }

        public bool AppendTable(string InTable, string TargetTable, bool Messages = false)
        {
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = true;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();


            // Populate the variant array with parameter values.
            parameters.Add(InTable);
            parameters.Add(TargetTable);

            // Execute the tool. Note this only works with geodatabase tables.
            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("Append_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);
                // Wait for 1 second.
                if (Messages)
                {
                    MessageBox.Show("Process complete");
                }
                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                gp = null;
                return false;
            }
        }

        public int CopyToCSV(string InTable, string OutTable, string aWhereClause, string Columns, string OrderByColumns, bool Spatial, bool Append, bool ExcludeHeader = false, bool Messages = false)
        {
            // This sub copies the input table to CSV.
            string aFilePath = myFileFuncs.GetDirectoryName(InTable);
            string aTabName = myFileFuncs.GetFileName(InTable);

            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.WhereClause = aWhereClause; // This works.
            ITable pTable = GetTable(myFileFuncs.GetDirectoryName(InTable), myFileFuncs.GetFileName(InTable));

            ICursor myCurs = null;
            IFields fldsFields = null;
            if (Spatial)
            {
                
                IFeatureClass myFC = GetFeatureClass(aFilePath, aTabName, true); 
                myCurs = (ICursor)myFC.Search(queryFilter, false);
                fldsFields = myFC.Fields;
            }
            else
            {
                ITable myTable = GetTable(aFilePath, aTabName, true);
                myCurs = myTable.Search(queryFilter, false);
                fldsFields = myTable.Fields;
            }

            if (myCurs == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot open table " + InTable);
                }
                return -1;
            }

            // Align the columns with what actually exists in the layer.
            // Return if there are no columns left.
            if (Columns != "")
            {
                List<string> strColumns = Columns.Split(',').ToList();
                Columns = "";
                foreach (string strCol in strColumns)
                {
                    string aColNameTr = strCol.Trim();
                    if ((aColNameTr.Substring(0, 1) == "\"") || (FieldExists(InTable, aColNameTr)))
                        Columns = Columns + aColNameTr + ",";
                }
                if (Columns != "")
                    Columns = Columns.Substring(0, Columns.Length - 1);
                else
                    return 0;
            }
            else
                return 0; // Technically we're finished as there is nothing to write.

            if (OrderByColumns != "")
            {
                List<string> strOrderColumns = OrderByColumns.Split(',').ToList();
                OrderByColumns = "";
                foreach (string strCol in strOrderColumns)
                {
                    if (FieldExists(InTable, strCol.Trim()))
                        OrderByColumns = OrderByColumns + strCol.Trim() + ",";
                }
                if (OrderByColumns != "")
                {
                    OrderByColumns = OrderByColumns.Substring(0, OrderByColumns.Length - 1);

                    ITableSort pTableSort = new TableSortClass();
                    pTableSort.Table = pTable;
                    pTableSort.Cursor = myCurs; 
                    pTableSort.Fields = OrderByColumns;

                    pTableSort.Sort(null);

                    myCurs = pTableSort.Rows;
                    Marshal.ReleaseComObject(pTableSort);
                    pTableSort = null;
                    GC.Collect();
                }
            }

            // Open output file.
            StreamWriter theOutput = new StreamWriter(OutTable, Append);
            List<string> ColumnList = Columns.Split(',').ToList();
            int intLineCount = 0;
            if (!Append && !ExcludeHeader)
            {
                string strHeader = Columns;
                theOutput.WriteLine(strHeader);
            }
            // Now write the file.
            IRow aRow = myCurs.NextRow();

            while (aRow != null)
            {
                string strRow = "";
                intLineCount++;
                foreach (string aColName in ColumnList)
                {
                    string aColNameTr = aColName.Trim();
                    if (aColNameTr.Substring(0, 1) != "\"")
                    {
                        int i = fldsFields.FindField(aColNameTr);
                        var theValue = aRow.get_Value(i);
                        // Wrap value if quotes if it is a string that contains a comma
                        if ((theValue is string) &&
                           (theValue.ToString().Contains(","))) theValue = "\"" + theValue.ToString() + "\"";
                        // Format distance to the nearest metre
                        if (theValue is double && aColNameTr == "Distance")
                        {
                            double dblValue = double.Parse(theValue.ToString());
                            int intValue = Convert.ToInt32(dblValue);
                            theValue = intValue;
                        }
                        strRow = strRow + theValue.ToString() + ",";
                    }
                    else
                    {
                        strRow = strRow + aColNameTr +",";
                    }
                    
                }

                strRow = strRow.Substring(0, strRow.Length - 1); // Remove final comma.

                theOutput.WriteLine(strRow);
                aRow = myCurs.NextRow();
            }

            theOutput.Close();
            theOutput.Dispose();
            aRow = null;
            pTable = null;
            Marshal.ReleaseComObject(myCurs);
            myCurs = null;
            GC.Collect();
            return intLineCount;
        }

        public bool WriteEmptyCSV(string OutTable, string theHeader)
        {
            // Open output file.
            try
            {
                StreamWriter theOutput = new StreamWriter(OutTable, false);
                theOutput.WriteLine(theHeader);
                theOutput.Close();
                theOutput.Dispose();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open " + OutTable + ". Please ensure this is not open in another window. System error: " + ex.Message);
                return false;
            }

        }

        public void ShowTable(string aTableName, bool Messages = false)
        {
            if (aTableName == null)
            {
                if (Messages) MessageBox.Show("Please pass a table name", "Show Table");
                return;
            }

            ITable myTable = GetTable(aTableName);
            if (myTable == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Table " + aTableName + " not found in map");
                    return;
                }
            }

            ITableWindow myWin = new TableWindow();
            myWin.Table = myTable;
            myWin.Application = thisApplication;
            myWin.Show(true);
        }

        public bool BufferFeature(string aLayer, string anOutputName, string aBufferDistance, string AggregateFields, bool Overwrite = true, bool Messages = false)
        {
            // Firstly check if the output feature exists.
            if (FeatureclassExists(anOutputName))
            {
                if (!Overwrite)
                {
                    if (Messages)
                        MessageBox.Show("The feature class " + anOutputName + " already exists. Cannot overwrite");
                    return false;
                }
            }
            if (!LayerExists(aLayer))
            {
                if (Messages)
                    MessageBox.Show("The layer " + aLayer + " does not exist in the map");
                return false;
            }

            if (GroupLayerExists(aLayer))
            {
                if (Messages)
                    MessageBox.Show("The layer " + aLayer + " is a group layer and cannot be buffered.");
                return false;
            }
            ILayer pLayer = GetLayer(aLayer);
            try
            {
                IFeatureLayer pTest = (IFeatureLayer)pLayer;
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("The layer " + aLayer + " is not a feature layer");
                return false;
            }

            // Check if all fields in the aggregate fields exist. If not, ignore.
            List<string> strAggColumns = AggregateFields.Split(';').ToList();
            AggregateFields = "";
            foreach (string strField in strAggColumns)
            {
                if (FieldExists(pLayer, strField))
                {
                    AggregateFields = AggregateFields + strField + ";";
                }
            }
            string strDissolveOption = "ALL";
            if (AggregateFields != "")
            {
                AggregateFields = AggregateFields.Substring(0, AggregateFields.Length - 1);
                strDissolveOption = "LIST";
            }


            // a different approach using the geoprocessor object.
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = Overwrite;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(aLayer);
            parameters.Add(anOutputName);
            parameters.Add(aBufferDistance);
            parameters.Add("FULL");
            parameters.Add("ROUND");
            parameters.Add(strDissolveOption);
            if (AggregateFields != "")
                parameters.Add(AggregateFields);

            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("Buffer_analysis", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);
                // Wait for 1 second.
                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                gp = null;
                return false;
            }

        }

        public bool BufferFeature(IFeature anInputFeature, string anOutputName, ISpatialReference aSpatialReference, double aBufferDistance, string aBufferUnit, bool Overwrite = false, bool Messages = false)
        {
            // While this works it is tricky to get the spatial reference correct.
            // Firstly check if the output feature exists.
            if (FeatureclassExists(anOutputName))
            {
                if (!Overwrite)
                {   
                    if (Messages)
                        MessageBox.Show("The feature class " + anOutputName + " already exists. Cannot overwrite");
                    return false;
                }
                else
                {
                    // Try to delete the output featureclass.
                   
                    bool blResult =  DeleteFeatureclass(anOutputName);
                    if (!blResult)
                    {
                        if (Messages)
                            MessageBox.Show("Cannot delete feature class " + anOutputName + ". Please check if it is open in another location");
                        return false;
                    }
                }
            }


            // All set up, now buffer the feature.
            IGeometry ptheGeometry = anInputFeature.Shape;
            ptheGeometry.SpatialReference = aSpatialReference;
           
            //ptheGeometry.SpatialReference = esriSRGeoCSType.esriSRGeoCS_OSGB1936;
            ITopologicalOperator5 pTopoOperator = (ITopologicalOperator5)ptheGeometry;
            IGeometry pResultPoly = pTopoOperator.Buffer(aBufferDistance); // in map units. Think about this.

            // create new featureclass
            string strWSName = myFileFuncs.GetDirectoryName(anOutputName);
            string strFCName = myFileFuncs.GetFileName(anOutputName);
            IFeatureClass pNewFC;
            try
            {
                pNewFC = CreateFeatureClass(strFCName, strWSName, esriGeometryType.esriGeometryPolygon);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not create new feature class " + anOutputName + ". System error: " + ex.Message);
                return false;
            }

            // Store the resulting polygon. This is the simplest implementation.
            try
            {
                IFeature newFeature = pNewFC.CreateFeature();
                newFeature.Shape = pResultPoly;
                newFeature.Store();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not insert new feature. System error: " + ex.Message);
                return false;
            }
            return true;
            
        }

        public IFeatureClass CreateFeatureClass(String featureClassName, string featureWorkspaceName, esriGeometryType aGeometryType, esriSRGeoCSType aSpatialReferenceSystem = esriSRGeoCSType.esriSRGeoCS_OSGB1936)
        {
            IWorkspaceFactory pWSF = GetWorkspaceFactory(featureWorkspaceName);
            IFeatureWorkspace featureWorkspace = (IFeatureWorkspace)pWSF.OpenFromFile(featureWorkspaceName, 0);
 
            // Assume we are always in Great Britain Grid.
            ISpatialReferenceFactory spatialReferenceFactory = new SpatialReferenceEnvironmentClass();
            ISpatialReference spatialReference = spatialReferenceFactory.CreateGeographicCoordinateSystem((int)aSpatialReferenceSystem);

            // Instantiate a feature class description to get the required fields.
            IFeatureClassDescription fcDescription = new FeatureClassDescriptionClass();
            IObjectClassDescription ocDescription = (IObjectClassDescription)fcDescription;
            IFields fields = ocDescription.RequiredFields;

            // Find the shape field in the required fields and modify its GeometryDef to
            // use relevant geometry and to set the spatial reference.

            int shapeFieldIndex = fields.FindField(fcDescription.ShapeFieldName);
            IField field = fields.get_Field(shapeFieldIndex);
            IGeometryDef geometryDef = field.GeometryDef;
            IGeometryDefEdit geometryDefEdit = (IGeometryDefEdit)geometryDef;
            geometryDefEdit.GeometryType_2 = aGeometryType;
            geometryDefEdit.SpatialReference_2 = spatialReference;

            // In this example, only the required fields from the class description are used as fields
            // for the feature class. If additional fields are added, IFieldChecker should be used to
            // validate them.

            // Create the feature class.

            IFeatureClass featureClass = featureWorkspace.CreateFeatureClass(featureClassName, fields,
              ocDescription.InstanceCLSID, ocDescription.ClassExtensionCLSID, esriFeatureType.esriFTSimple,
              fcDescription.ShapeFieldName, "");

            Marshal.ReleaseComObject(featureWorkspace);
            featureWorkspace = null;
            pWSF = null;
            GC.Collect();
                
            return featureClass;

        }

        public bool DeleteFeatureclass(string aFeatureclassName, bool Messages = false)
        {

            // a different approach using the geoprocessor object.
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = true;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();


            // Populate the variant array with parameter values.
            parameters.Add(aFeatureclassName);

            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("Delete_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);
                // Wait for 1 second.
                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                gp = null;
                return false;
            }

        }

        public IFeature GetFeatureFromLayer(string aFeatureLayer, string aQuery, bool Messages = false)
        {
            // This function returns a feature from the FeatureLayer. If there is more than one feature, it returns the LAST one.
            // Check if the layer exists.
            if (!LayerExists(aFeatureLayer))
            {
                if (Messages)
                    MessageBox.Show("Cannot find feature layer " + aFeatureLayer);
                return null;
            }

            ILayer pLayer = GetLayer(aFeatureLayer);
            IFeatureLayer pFL;
            try
            {
                pFL = (IFeatureLayer)pLayer;
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("Layer " + aFeatureLayer + " is not a feature layer.");
                return null;
            }

            IFeatureClass pFC = pFL.FeatureClass;
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = aQuery;
            IFeatureCursor pCurs = pFC.Search(pQueryFilter, false);

            int aCount = 0;
            IFeature feature = null;
            IFeature pResult = null;
            int nameFieldIndex = pFC.FindField("Shape");
            try
            {
                while ((feature = pCurs.NextFeature()) != null)
                {
                    aCount = aCount + 1;
                    pResult = feature;
                }
                
            }
            catch (COMException comExc)
            {
                // Handle any errors that might occur on NextFeature().
                if (Messages)
                    MessageBox.Show("Error: " + comExc.Message);
                Marshal.ReleaseComObject(pCurs);
                return null;
            }

            // Release the cursor.
            Marshal.ReleaseComObject(pCurs);

            if (aCount == 0)
            {
                if (Messages)
                    MessageBox.Show("There was no feature found matching those criteria");
                return null;
            }
            else if (aCount > 1)
            {
                if (Messages)
                    MessageBox.Show("There were " + aCount.ToString() + " features found matching those criteria");
                return null;
            }
            else
                return pResult;

        }

        public ISpatialReference GetSpatialReference(string aFeatureLayer, bool Messages = false)
        {
            // This falls over for reasons unknown.

            // Check if the layer exists.
            if (!LayerExists(aFeatureLayer))
            {
                if (Messages)
                    MessageBox.Show("Cannot find feature layer " + aFeatureLayer);
                return null;
            }

            ILayer pLayer = GetLayer(aFeatureLayer);
            IFeatureLayer pFL;
            try
            {
                pFL = (IFeatureLayer)pLayer;
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("Layer " + aFeatureLayer + " is not a feature layer.");
                return null;
            }

            IFeatureClass pFC = pFL.FeatureClass;
            IDataset pDataSet = pFC.FeatureDataset;
            IGeoDataset pDS = (IGeoDataset)pDataSet;
            MessageBox.Show(pDS.SpatialReference.ToString());
            ISpatialReference pRef = pDS.SpatialReference;
            return pRef;
        }

        public bool SelectLayerByAttributes(string aFeatureLayerName, string aWhereClause, string aSelectionType = "NEW_SELECTION", bool Messages = false)
        {
            ///<summary>Select features in the IActiveView by an attribute query using a SQL syntax in a where clause.</summary>
            /// 
            ///<param name="featureLayer">An IFeatureLayer interface to select upon</param>
            ///<param name="whereClause">A System.String that is the SQL where clause syntax to select features. Example: "CityName = 'Redlands'"</param>
            ///  
            ///<remarks>Providing and empty string "" will return all records.</remarks>
            if (!LayerExists(aFeatureLayerName))
                return false;

            IActiveView activeView = GetActiveView();
            IFeatureLayer featureLayer = null;
            try
            {
                featureLayer = (IFeatureLayer)GetLayer(aFeatureLayerName);
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("The layer " + aFeatureLayerName + " is not a feature layer");
                return false;
            }

            if (activeView == null || featureLayer == null || aWhereClause == null)
            {
                if (Messages)
                    MessageBox.Show("Please check input for this tool");
                return false;
            }


            // do this with Geoprocessor.

            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();


            // Populate the variant array with parameter values.
            parameters.Add(aFeatureLayerName);
            parameters.Add(aSelectionType);
            parameters.Add(aWhereClause);

            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("SelectLayerByAttribute_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);
                // Wait for 1 second.
                if (Messages)
                {
                    MessageBox.Show("Process complete");
                }
                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                gp = null;
                return false;
            }
            //ESRI.ArcGIS.Carto.IFeatureSelection featureSelection = featureLayer as ESRI.ArcGIS.Carto.IFeatureSelection; // Dynamic Cast

            //// Set up the query
            //ESRI.ArcGIS.Geodatabase.IQueryFilter queryFilter = new ESRI.ArcGIS.Geodatabase.QueryFilterClass();
            //queryFilter.WhereClause = whereClause;

            //// Invalidate only the selection cache. Flag the original selection
            //activeView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, null, null);

            //// Perform the selection
            //featureSelection.SelectFeatures(queryFilter, ESRI.ArcGIS.Carto.esriSelectionResultEnum.esriSelectionResultNew, false);

            //// Flag the new selection
            //activeView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, null, null);
        }

        public bool SelectLayerByLocation(string aTargetLayer, string aSearchLayer, string anOverlapType = "INTERSECT", string aSearchDistance = "", string aSelectionType = "NEW_SELECTION", bool Messages = false)
        {
            // Implementation of python SelectLayerByLocation_management.

            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();


            // Populate the variant array with parameter values.
            parameters.Add(aTargetLayer);
            parameters.Add(anOverlapType);
            parameters.Add(aSearchLayer);
            parameters.Add(aSearchDistance);
            parameters.Add(aSelectionType);

            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("SelectLayerByLocation_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);
                // Wait for 1 second.
                if (Messages)
                {
                    MessageBox.Show("Process complete");
                }
                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                gp = null;
                return false;
            }
        }

        public int CountSelectedLayerFeatures(string aFeatureLayerName, bool Messages = false)
        {
            // Check input.
            if (aFeatureLayerName == null)
            {
                if (Messages) MessageBox.Show("Please pass valid input string", "Feature Layer Has Selection");
                return -1;
            }

            if (!LayerExists(aFeatureLayerName))
            {
                if (Messages) MessageBox.Show("Feature layer " + aFeatureLayerName + " does not exist in this map");
                return -1;
            }

            IFeatureLayer pFL = null;
            try
            {
                pFL = (IFeatureLayer)GetLayer(aFeatureLayerName);
            }
            catch
            {
                if (Messages)
                    MessageBox.Show(aFeatureLayerName + " is not a feature layer");
                return -1;
            }

            IFeatureSelection pFSel = (IFeatureSelection)pFL;
            if (pFSel.SelectionSet.Count > 0) return pFSel.SelectionSet.Count;
            return 0;
        }

        public void ClearSelectedMapFeatures(string aFeatureLayerName, bool Messages = false)
        {
            ///<summary>Clear the selected features in the IActiveView for a specified IFeatureLayer.</summary>
            /// 
            ///<param name="activeView">An IActiveView interface</param>
            ///<param name="featureLayer">An IFeatureLayer</param>
            /// 
            ///<remarks></remarks>
            if (!LayerExists(aFeatureLayerName))
                return;

            IActiveView activeView = GetActiveView();
            IFeatureLayer featureLayer = null;
            try
            {
                featureLayer = (IFeatureLayer)GetLayer(aFeatureLayerName);
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("The layer " + aFeatureLayerName + " is not a feature layer");
                return;
            }
            if (activeView == null || featureLayer == null)
            {
                return;
            }
            ESRI.ArcGIS.Carto.IFeatureSelection featureSelection = featureLayer as ESRI.ArcGIS.Carto.IFeatureSelection; // Dynamic Cast

            // Invalidate only the selection cache. Flag the original selection
            activeView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, null, null);

            // Clear the selection
            featureSelection.Clear();

            // Flag the new selection
            activeView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, null, null);
        }

        public void ZoomToLayer(string aLayerName, bool Messages = false)
        {
            if (!LayerExists(aLayerName))
            {
                if (Messages)
                    MessageBox.Show("The layer " + aLayerName + " does not exist in the map");
                return;
            }
            IActiveView activeView = GetActiveView();
            ILayer pLayer = GetLayer(aLayerName);
            IEnvelope pEnv = pLayer.AreaOfInterest;
            pEnv.Expand(1.05, 1.05, true);
            activeView.Extent = pEnv;
            activeView.Refresh();
        }

        public void ChangeLegend(string aLayerName, string aLayerFile, bool Messages = false)
        {
            if (!LayerExists(aLayerName))
            {
                if (Messages)
                    MessageBox.Show("The layer " + aLayerName + " does not exist in the map");
                return;
            }
            if (!myFileFuncs.FileExists(aLayerFile))
            {
                if (Messages)
                    MessageBox.Show("The layer file " + aLayerFile + " does not exist");
                return;
            }

            ILayer pLayer = GetLayer(aLayerName);
            IGeoFeatureLayer pTargetLayer = null;
            try
            {
                pTargetLayer = (IGeoFeatureLayer)pLayer;
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("The input layer " + aLayerName + " is not a feature layer");
                return;
            }
            ILayerFile pLayerFile = new LayerFileClass();
            pLayerFile.Open(aLayerFile);

            IGeoFeatureLayer pTemplateLayer = (IGeoFeatureLayer)pLayerFile.Layer;
            IFeatureRenderer pTemplateSymbology = pTemplateLayer.Renderer;
            pLayerFile.Close();

            IObjectCopy pCopy = new ObjectCopyClass();
            pTargetLayer.Renderer = (IFeatureRenderer)pCopy.Copy(pTemplateSymbology);

        }

        public bool CalculateField(string aLayerName, string aFieldName, string aCalculate, bool Messages = false)
        {
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();


            // Populate the variant array with parameter values.
            parameters.Add(aLayerName);
            parameters.Add(aFieldName);
            parameters.Add(aCalculate);
            parameters.Add("VB");

            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("CalculateField_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);
                // Wait for 1 second.
                if (Messages)
                {
                    MessageBox.Show("Process complete");
                }
                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                gp = null;
                return false;
            }
        }

        public bool ExportSelectionToShapefile(string aLayerName, string anOutShapefile, string OutputColumns, string TempShapeFile, string GroupColumns = "",
            string StatisticsColumns = "", bool IncludeDistance = false, string aTargetLayer = null, bool Overwrite = true, bool CheckForSelection = false, bool RenameColumns = false, bool Messages = false)
        {
            // Some sanity tests.
            if (!LayerExists(aLayerName))
            {
                if (Messages)
                    MessageBox.Show("The layer " + aLayerName + " does not exist in the map");
                return false;
            }
            if (CountSelectedLayerFeatures(aLayerName, Messages) <= 0 && CheckForSelection)
            {
                if (Messages)
                    MessageBox.Show("The layer " + aLayerName + " does not have a selection");
                return false;
            }

            // Does the output file exist?
            if (FeatureclassExists(anOutShapefile))
            {
                if (!Overwrite)
                {
                    if (Messages)
                        MessageBox.Show("The output feature class " + anOutShapefile + " already exists. Cannot overwrite");
                    return false;
                }
            }

            ILayer pLayer = GetLayer(aLayerName);
            IFeatureLayer pFL = (IFeatureLayer)pLayer;
            IFeatureClass pFC = pFL.FeatureClass;

            // Check all the requested group by and statistics fields exist.
            // Only pass those that do.
            if (GroupColumns != "")
            {
                List<string> strColumns = GroupColumns.Split(';').ToList();
                GroupColumns = "";
                foreach (string strCol in strColumns)
                {
                    if (FieldExists(pFC, strCol.Trim()))
                        GroupColumns = GroupColumns + strCol.Trim() + ";";
                }
                if (GroupColumns != "")
                    GroupColumns = GroupColumns.Substring(0, GroupColumns.Length - 1);

            }

            if (StatisticsColumns != "")
            {
                List<string> strStatsColumns = StatisticsColumns.Split(';').ToList();
                StatisticsColumns = "";
                foreach (string strColDef in strStatsColumns)
                {
                    List<string> strComponents = strColDef.Split(' ').ToList();
                    string strField = strComponents[0]; // The field name.
                    if (FieldExists(pFC, strField.Trim()))
                        StatisticsColumns = StatisticsColumns + strColDef + ";";
                }
                if (StatisticsColumns != "")
                    StatisticsColumns = StatisticsColumns.Substring(0, StatisticsColumns.Length - 1);
            }



            string strTempLayer = myFileFuncs.ReturnWithoutExtension(myFileFuncs.GetFileName(TempShapeFile)); // Temporary layer.

            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = Overwrite;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // If we are including distance, the process is slighly different.
            if (GroupColumns != null) // include group columns.
            {
                string strOutFile = TempShapeFile;
                if (!IncludeDistance)
                    // We are ONLY performing a group by. Go straight to final shapefile.
                    strOutFile = anOutShapefile;
        

                // Do the dissolve as requested.
                IVariantArray DissolveParams = new VarArrayClass();
                DissolveParams.Add(aLayerName);
                DissolveParams.Add(strOutFile);
                DissolveParams.Add(GroupColumns);
                DissolveParams.Add(StatisticsColumns); // These should be set up to be as required beforehand.

                try
                {
                    myresult = (IGeoProcessorResult)gp.Execute("Dissolve_management", DissolveParams, null);

                    // Wait until the execution completes.
                    while (myresult.Status == esriJobStatus.esriJobExecuting)
                        Thread.Sleep(1000);
                    // Wait for 1 second.
                    string strNewLayer = myFileFuncs.ReturnWithoutExtension(myFileFuncs.GetFileName(strOutFile));

                    ILayer pInLayer = GetLayer(aLayerName);
                    IFeatureLayer pInFLayer = (IFeatureLayer)pInLayer;
                    IFeatureClass pInFC = pInFLayer.FeatureClass;

                    ILayer pOutLayer = GetLayer(strNewLayer);
                    IFeatureLayer pOutFLayer = (IFeatureLayer)pOutLayer;
                    IFeatureClass pOutFC = pOutFLayer.FeatureClass;

                    // Now rejig the statistics fields if required because they will look like FIRST_SAC which is no use.
                    if (StatisticsColumns != "" && RenameColumns)
                    {
                        List<string> strFieldNames = StatisticsColumns.Split(';').ToList();
                        int intIndexCount = 0;
                        foreach (string strField in strFieldNames)
                        {
                            List<string> strFieldComponents = strField.Split(' ').ToList();
                            // Let's find out what the new field is called - could be anything.
                            int intNewIndex = 2; // FID = 1; Shape = 2.
                            intNewIndex = intNewIndex + GroupColumns.Split(';').ToList().Count + intIndexCount; // Add the number of columns uses for grouping
                            IField pNewField = pOutFC.Fields.get_Field(intNewIndex);
                            string strInputField = pNewField.Name;
                            // Note index stays the same, since we're deleting the fields. 
                            
                            string strNewField = strFieldComponents[0]; // The original name of the field.
                            // Get the definition of the original field from the original feature class.
                            int intIndex = pInFC.Fields.FindField(strNewField);
                            IField pField = pInFC.Fields.get_Field(intIndex);

                            // Add the field to the new FC.
                            AddLayerField(strNewLayer, strNewField, pField.Type, pField.Length, Messages);
                            // Calculate the new field.
                            string strCalc = "[" + strInputField + "]";
                            CalculateField(strNewLayer, strNewField, strCalc);
                            DeleteLayerField(strNewLayer, strInputField);
                        }
                        
                    }

                    aLayerName = strNewLayer;
                    
                }
                catch (COMException ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    gp = null;
                    return false;
                }

            }
            if (IncludeDistance)
            {
                // Now add the distance field by joining if required. This will take all fields.

                IVariantArray params1 = new VarArrayClass();
                params1.Add(aLayerName);
                params1.Add(aTargetLayer);
                params1.Add(anOutShapefile);
                params1.Add("JOIN_ONE_TO_ONE");
                params1.Add("KEEP_ALL");
                params1.Add("");
                params1.Add("CLOSEST");
                params1.Add("0");
                params1.Add("Distance");

                try
                {
                    myresult = (IGeoProcessorResult)gp.Execute("SpatialJoin_analysis", params1, null);

                    // Wait until the execution completes.
                    while (myresult.Status == esriJobStatus.esriJobExecuting)
                        Thread.Sleep(1000);
                    // Wait for 1 second.
                    
                }
                catch (COMException ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    gp = null;
                    return false;
                }
            }

            
            if (GroupColumns == null && !IncludeDistance) 
                // Only run a straight copy if neither a group nor a distance has been requested
                // Because the data won't have been processed yet.
            {

                // Create a variant array to hold the parameter values.
                IVariantArray parameters = new VarArrayClass();

                // Populate the variant array with parameter values.
                parameters.Add(aLayerName);
                parameters.Add(anOutShapefile);

                try
                {
                    myresult = (IGeoProcessorResult)gp.Execute("CopyFeatures_management", parameters, null);

                    // Wait until the execution completes.
                    while (myresult.Status == esriJobStatus.esriJobExecuting)
                        Thread.Sleep(1000);
                    // Wait for 1 second.
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    gp = null;
                    return false;
                }
            }

            // Remove all temporary layers.
            bool blFinished = false;
            while (!blFinished)
            {
                if (LayerExists(strTempLayer))
                    RemoveLayer(strTempLayer);
                else
                    blFinished = true;
            }

            if (FeatureclassExists(TempShapeFile))
            {
                IVariantArray DelParams = new VarArrayClass();
                DelParams.Add(TempShapeFile);
                try
                {

                    myresult = (IGeoProcessorResult)gp.Execute("Delete_management", DelParams, null);

                    // Wait until the execution completes.
                    while (myresult.Status == esriJobStatus.esriJobExecuting)
                        Thread.Sleep(1000);
                    // Wait for 1 second.
                }
                catch (Exception ex)
                {
                    if (Messages)
                        MessageBox.Show("Cannot delete temporary layer " + TempShapeFile + ". System error: " + ex.Message);
                }
            }

            // Now drop any fields from the output that we don't want.
            pFC = GetFeatureClass(anOutShapefile);
            IFields pFields = pFC.Fields;
            List<string> strDeleteFields = new List<string>();

            // Make a list of fields to delete.
            for (int i = 0; i < pFields.FieldCount; i++)
            {
                IField pField = pFields.get_Field(i);
                if (OutputColumns.IndexOf(pField.Name) == -1 && !pField.Required) 
                    // Does it exist in the 'keep' list or is it required?
                {
                    // If not, add to te delete list.
                    strDeleteFields.Add(pField.Name);
                }
            }

            //Delete the listed fields.
            foreach (string strField in strDeleteFields)
            {
                DeleteField(pFC, strField);
            }
            
            pFC = null;
            gp = null;

            UpdateTOC();
            GC.Collect(); // Just in case it's hanging onto anything.

            return true;
        }


        public void AnnotateLayer(string thisLayer, String LabelExpression, string aFont = "Arial",double aSize = 10, int Red = 0, int Green = 0, int Blue = 0, string OverlapOption = "OnePerShape", bool annotationsOn = true, bool showMapTips = false, bool Messages = false)
        {
            // Options: OnePerShape, OnePerName, OnePerPart and NoRestriction.
            ILayer pLayer = GetLayer(thisLayer);
            try
            {
                IFeatureLayer pFL = (IFeatureLayer)pLayer;
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("Layer " + thisLayer + " is not a feature layer");
                return;
            }

            esriBasicNumLabelsOption esOverlapOption;
            if (OverlapOption == "NoRestriction")
                esOverlapOption = esriBasicNumLabelsOption.esriNoLabelRestrictions;
            else if (OverlapOption == "OnePerName")
                esOverlapOption = esriBasicNumLabelsOption.esriOneLabelPerName;
            else if (OverlapOption == "OnePerPart")
                esOverlapOption = esriBasicNumLabelsOption.esriOneLabelPerPart;
            else
                esOverlapOption = esriBasicNumLabelsOption.esriOneLabelPerShape;

            stdole.IFontDisp fnt = (stdole.IFontDisp)new stdole.StdFontClass();
            fnt.Name = aFont;
            fnt.Size = Convert.ToDecimal(aSize);

            RgbColor annotationLabelColor = new RgbColorClass();
            annotationLabelColor.Red = Red;
            annotationLabelColor.Green = Green;
            annotationLabelColor.Blue = Blue;

            IGeoFeatureLayer geoLayer = pLayer as IGeoFeatureLayer;
            if (geoLayer != null)
            {
                geoLayer.DisplayAnnotation = annotationsOn;
                IAnnotateLayerPropertiesCollection propertiesColl = geoLayer.AnnotationProperties;
                IAnnotateLayerProperties labelEngineProperties = new LabelEngineLayerProperties() as IAnnotateLayerProperties;
                IElementCollection placedElements = new ElementCollectionClass();
                IElementCollection unplacedElements = new ElementCollectionClass();
                propertiesColl.QueryItem(0, out labelEngineProperties, out placedElements, out unplacedElements);
                ILabelEngineLayerProperties lpLabelEngine = labelEngineProperties as ILabelEngineLayerProperties;
                lpLabelEngine.Expression = LabelExpression;
                lpLabelEngine.Symbol.Color = annotationLabelColor;
                lpLabelEngine.Symbol.Font = fnt;
                lpLabelEngine.BasicOverposterLayerProperties.NumLabelsOption = esOverlapOption;
                IFeatureLayer thisFeatureLayer = pLayer as IFeatureLayer;
                IDisplayString displayString = thisFeatureLayer as IDisplayString;
                IDisplayExpressionProperties properties = displayString.ExpressionProperties;
                
                properties.Expression = LabelExpression; //example: "[OWNER_NAME] & vbnewline & \"$\" & [TAX_VALUE]";
                thisFeatureLayer.ShowTips = showMapTips;
            }
        }

        public bool DeleteField(IFeatureClass aFeatureClass, string aFieldName)
        {
            // Get the fields collection
            int intIndex = aFeatureClass.Fields.FindField(aFieldName);
            IField pField = aFeatureClass.Fields.get_Field(intIndex);
            try
            {
                aFeatureClass.DeleteField(pField);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot delete field " + aFieldName + ". System error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

        }

        public int AddIncrementalNumbers(string aFeatureClass, string aFieldName, string aKeyField, int aStartNumber = 1, bool Messages = false)
        {
            // Check the obvious.
            if (!FeatureclassExists(aFeatureClass))
            {
                if (Messages)
                    MessageBox.Show("The featureclass " + aFeatureClass + " doesn't exist");
                return -1;
            }

            if (!FieldExists(aFeatureClass, aFieldName))
            {
                if (Messages)
                    MessageBox.Show("The field " + aFieldName + " does not exist in featureclass " + aFeatureClass);
                return -1;
            }

            if (!FieldIsNumeric(aFeatureClass, aFieldName))
            {
                if (Messages)
                    MessageBox.Show("The field " + aFieldName + " is not numeric");
                return -1;
            }

            // All hurdles passed - let's do this.
            // Firstly make the list of labels in the correct order.
            // Get the search cursor
            IQueryFilter pQFilt = new QueryFilterClass();
            pQFilt.SubFields = aFieldName + "," + aKeyField;
            IFeatureClass pFC = GetFeatureClass(aFeatureClass);
            IFeatureCursor pSearchCurs = pFC.Search(pQFilt, false);

            // Sort the cursor
            ITableSort pTableSort = new TableSortClass();
            pTableSort.Table = (ITable)pFC;
            pTableSort.Fields = aKeyField;
            pTableSort.Cursor = (ICursor)pSearchCurs;
            pTableSort.Sort(null);
            pSearchCurs = (IFeatureCursor)pTableSort.Rows;
            Marshal.ReleaseComObject(pTableSort); // release the sort object.

            // Extract the lists of values.
            IFields pFields = pFC.Fields;
            int intFieldIndex = pFields.FindField(aFieldName);
            int intKeyFieldIndex = pFields.FindField(aKeyField);
            List<string> KeyList = new List<string>();
            List<int> ValueList = new List<int>(); // These lists are in sync.

            IFeature feature = null;
            int intMax = aStartNumber - 1;
            int intValue = intMax;
            string strKey = "";
            while ((feature = pSearchCurs.NextFeature()) != null)
            {
                string strTest = feature.get_Value(intKeyFieldIndex).ToString();
                if (strTest != strKey) // Different key value
                {
                    // Do we know about it?
                    if (KeyList.IndexOf(strTest) != -1)
                    {
                        intValue = ValueList[KeyList.IndexOf(strTest)];
                        strKey = strTest;
                    }
                    else
                    {
                        intMax++;
                        intValue = intMax;
                        strKey = strTest;
                        KeyList.Add(strKey);
                        ValueList.Add(intValue);
                    }
                }
            }
            Marshal.ReleaseComObject(pSearchCurs);
            pSearchCurs = null;

            // Now do the update.
            IFeatureCursor pUpdateCurs = pFC.Update(pQFilt, false);
            strKey = "";
            intValue = -1;
            try
            {
            while ((feature = pUpdateCurs.NextFeature()) != null)
                {
                    string strTest = feature.get_Value(intKeyFieldIndex).ToString();
                    if (strTest != strKey) // Different key value
                    {
                        // Find out all about it
                        intValue = ValueList[KeyList.IndexOf(strTest)];
                        strKey = strTest;
                    }
                    feature.set_Value(intFieldIndex, intValue);
                    pUpdateCurs.UpdateFeature(feature);
                }
            }
            catch (Exception ex)
            {
                if (Messages)
                    MessageBox.Show("Error: " + ex.Message, "Error");
                Marshal.ReleaseComObject(pUpdateCurs);
            }

            // If the cursor is no longer needed, release it.
            Marshal.ReleaseComObject(pUpdateCurs);
            pUpdateCurs = null;
            return intMax; // Return the maximum value for info.
        }

        public void ToggleDrawing(bool AllowDrawing)
        {
            IMxApplication2 thisApp = (IMxApplication2)thisApplication;
            thisApp.PauseDrawing = !AllowDrawing;
            if (AllowDrawing)
            {
                IActiveView activeView = GetActiveView();
                activeView.Refresh();
            }
        }


        public void ToggleTOC()
        {
            IApplication m_app = thisApplication;

            IDockableWindowManager pDocWinMgr = m_app as IDockableWindowManager;
            UID uid = new UIDClass();
            uid.Value = "{368131A0-F15F-11D3-A67E-0008C7DF97B9}";
            IDockableWindow pTOC = pDocWinMgr.GetDockableWindow(uid);
            if (pTOC.IsVisible())
                pTOC.Show(false); 
            else pTOC.Show(true);

           
            IActiveView activeView = GetActiveView();
            activeView.Refresh();

        }

    }
}
