// Copyright (c) 2016 Autodesk, Inc. All rights reserved.
// Author: paolo.serra@autodesk.com
// 
// Licensed under the Apache License, Version 2.0 (the "License").
// You may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//    http://www.apache.org/licenses/LICENSE-2.0
// 
//  Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or
// implied.  See the License for the specific language governing
// permissions and limitations under the License.
using Autodesk.AECC.Interop.UiRoadway;
using Autodesk.AutoCAD.Interop;
using Autodesk.DesignScript.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;


namespace CivilConnection
{

    /// <summary>
    /// CivilApplication object type.
    /// </summary>
    public class CivilApplication
    {

        /// <summary>
        /// The documents in Civil 3D.
        /// </summary>
        private IList<CivilDocument> Documents;
        /// <summary>
        /// The land XML path.
        /// </summary>
        public string LandXMLPath;

        /// <summary>
        /// The active document
        /// </summary>
        AcadDocument ActiveDocument;
        /// <summary>
        /// The active application
        /// </summary>
        AeccRoadwayApplication mApp;
        /// <summary>
        /// Gets the internal element.
        /// </summary>
        /// <value>
        /// The internal element.
        /// </value>
        internal object InternalElement { get { return this.mApp; } }

        /// <summary>
        /// Gets the application.
        /// </summary>
        /// <returns></returns>
        internal AeccRoadwayApplication GetApplication()
        {
            Utils.Log(string.Format("GetApplication started...", ""));

            string m_sAcadProdID = "AutoCAD.Application";

            string[] progids = new string[] {
                "AeccXUiRoadway.AeccRoadwayApplication.13.2", // 2020
                "AeccXUiRoadway.AeccRoadwayApplication.13.0", // 2019
                "AeccXUiRoadway.AeccRoadwayApplication.12.0", // 2018
                "AeccXUiRoadway.AeccRoadwayApplication.11.0", // 2017
                "AeccXUiRoadway.AeccRoadwayApplication.10.5" // 2016
            };

#if C2021
            progids = new string[] {"AeccXUiRoadway.AeccRoadwayApplication.13.3"};  // 2021
#endif

            AcadApplication m_oAcadApp = System.Runtime.InteropServices.Marshal.GetActiveObject(m_sAcadProdID) as AcadApplication;

            // Roadway application

            dynamic output = null;

            foreach (string r_sAeccAppProgId in progids)
            {
                try
                {
                    output = m_oAcadApp.GetInterfaceObject(r_sAeccAppProgId);

                    if (output != null)
                    {
                        break;
                    }
                }
                catch 
                {
                    continue;
                }
            }
            //return m_oAcadApp.GetInterfaceObject(r_sAeccAppProgId);

            Utils.Log(string.Format("GetApplication completed.", ""));

            return output;
        }

        /// <summary>
        /// Creates the connection with the running session of Civil 3D.
        /// </summary>
        public CivilApplication()
        {
            Utils.InitializeLog();

            Utils.Log(string.Format("CivilApplication.CivilApplication started...", ""));

            this.mApp = this.GetApplication();

            IList<CivilDocument> documents = new List<CivilDocument>();

            foreach (var doc in this.mApp.Documents)
            {
                documents.Add(new CivilDocument(doc as AeccRoadwayDocument));
            }

            this.Documents = documents;
            this.ActiveDocument = mApp.ActiveDocument;

            var revitDoc = RevitServices.Persistence.DocumentManager.Instance.CurrentDBDocument;

            RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(revitDoc);

            Autodesk.Revit.DB.Units units = new Autodesk.Revit.DB.Units(Autodesk.Revit.DB.UnitSystem.Metric);

            var du = this.Documents.First()._document.Settings.DrawingSettings.UnitZoneSettings.DrawingUnits;

            Utils.Log(string.Format("CivilApplication.Units started...", ""));

            // 1.1.0 Change Revit Document untis to match the Civil 3D Units
            if (du == Autodesk.AECC.Interop.Land.AeccDrawingUnitType.aeccDrawingUnitMeters)
            {
                units.SetFormatOptions(Autodesk.Revit.DB.UnitType.UT_Length, new Autodesk.Revit.DB.FormatOptions(Autodesk.Revit.DB.DisplayUnitType.DUT_METERS));
            }
            else if (du == Autodesk.AECC.Interop.Land.AeccDrawingUnitType.aeccDrawingUnitDecimeters)
            {
                units.SetFormatOptions(Autodesk.Revit.DB.UnitType.UT_Length, new Autodesk.Revit.DB.FormatOptions(Autodesk.Revit.DB.DisplayUnitType.DUT_DECIMETERS));
            }
            else if (du == Autodesk.AECC.Interop.Land.AeccDrawingUnitType.aeccDrawingUnitCentimeters)
            {
                units.SetFormatOptions(Autodesk.Revit.DB.UnitType.UT_Length, new Autodesk.Revit.DB.FormatOptions(Autodesk.Revit.DB.DisplayUnitType.DUT_CENTIMETERS));
            }
            else if (du == Autodesk.AECC.Interop.Land.AeccDrawingUnitType.aeccDrawingUnitMillimeters)
            {
                units.SetFormatOptions(Autodesk.Revit.DB.UnitType.UT_Length, new Autodesk.Revit.DB.FormatOptions(Autodesk.Revit.DB.DisplayUnitType.DUT_MILLIMETERS));
            }
            else if (du == Autodesk.AECC.Interop.Land.AeccDrawingUnitType.aeccDrawingUnitFeet)
            {
                units.SetFormatOptions(Autodesk.Revit.DB.UnitType.UT_Length, new Autodesk.Revit.DB.FormatOptions(Autodesk.Revit.DB.DisplayUnitType.DUT_DECIMAL_FEET));
            }          
            else if (du == Autodesk.AECC.Interop.Land.AeccDrawingUnitType.aeccDrawingUnitInches)
            {
                units.SetFormatOptions(Autodesk.Revit.DB.UnitType.UT_Length, new Autodesk.Revit.DB.FormatOptions(Autodesk.Revit.DB.DisplayUnitType.DUT_DECIMAL_INCHES));
            }
            else
            {
                throw new Exception("UNITS ERROR\nThe Civil 3D units of the Active Document are not supported in Revit.\nChange the Civil 3D Units to continue");
            }

            revitDoc.SetUnits(units);

            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

            Utils.Log(string.Format("CivilApplication.Units completed.", ""));
            
            SessionVariables.LandXMLPath = System.IO.Path.GetTempPath();
            SessionVariables.IsLandXMLExported = false;
            SessionVariables.CivilApplication = this;
            SessionVariables.ParametersCreated = false;
            SessionVariables.DocumentTotalTransform = null;
            RevitUtils.DocumentTotalTransform();
        }

        /// <summary>
        /// Returns the list of Civil Documents opened in Civil 3D.
        /// </summary>
        /// <returns></returns>
        public IList<CivilDocument> GetDocuments()
        {

            //foreach (CivilDocument doc in this.Documents)
            //{
            //    //Utils.DumpLandXML(doc._document);
            //}

            return this.Documents;
        }

        /// <summary>
        /// Returns the Civil Documents opened in Civil 3D with the same name.
        /// </summary>
        /// <param name="name">The Document name</param>
        /// <returns></returns>
        public CivilDocument GetDocumentByName(string name)
        {
            return this.Documents.First(x => x.Name == name);
        }

     
        /// <summary>
        /// Enables the Run Periodically mode and updates the connection with Civil 3D.
        /// </summary>
        /// <returns></returns>
        [CanUpdatePeriodicallyAttribute(true)]
        public CivilApplication UpdatePeriodically()
        {
            Utils.Log(string.Format("CivilApplication.UpdatePeriodically started...", ""));

            return new CivilApplication();
        }


        /// <summary>
        /// Writes a message to the log file
        /// </summary>
        /// <param name="data">The data that is passed through</param>
        /// <param name="message">An optional message to write to the log.</param>
        /// <returns></returns>
        public static object WriteToLog(object data, string message = "")
        {
            Utils.Log(string.Format("{0}{1}", message.Length > 0 ? message + " " : "", data));

            return data;
        }

        /// <summary>
        /// Public textual representation of the Dynamo node preview.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("CivilApplication(ActiveDocument = {0})", this.ActiveDocument.Name);
        }
    }
}
