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
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Revit.GeometryConversion;
using RevitServices.Persistence;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;


namespace CivilConnection
{
    /// <summary>
    /// Collection of utilities for obejct location.
    /// </summary>
    [DynamoServices.RegisterForTrace()]
    [SupressImportIntoVM()]
    public class UtilsObjectsLocation
    {
        /// <summary>
        /// Feets to mm.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        public static double FeetToMm(double d)
        {
            return d * 304.8;
        }

        /// <summary>
        /// Feets to m.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        public static double FeetToM(double d)
        {
            return d * 0.3048;
        }

        /// <summary>
        /// ms to feet.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        public static double MToFeet(double d)
        {
            return d / 0.3048;
        }

        /// <summary>
        /// Mms to feet.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        public static double MmToFeet(double d)
        {
            return d / 304.8;
        }

        /// <summary>
        /// Degs to radians.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        public static double DegToRadians(double d)
        {
            return d / 180 * Math.PI;
        }

        /// <summary>
        /// Radianses to deg.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        public static double RadiansToDeg(double d)
        {
            return d / Math.PI * 180;
        }

        /// <summary>
        /// Xyzms to feet.
        /// </summary>
        /// <param name="xyz">The xyz.</param>
        /// <returns></returns>
        private static XYZ XYZMToFeet(XYZ xyz)
        {
            return new XYZ(MToFeet(xyz.X), MToFeet(xyz.Y), MToFeet(xyz.Z));
        }

        /// <summary>
        /// Xyzs the feet to m.
        /// </summary>
        /// <param name="xyz">The xyz.</param>
        /// <returns></returns>
        private static XYZ XYZFeetToM(XYZ xyz)
        {
            return new XYZ(FeetToM(xyz.X), FeetToM(xyz.Y), FeetToM(xyz.Z));
        }

        /// <summary>
        /// Locations the XML.
        /// </summary>
        /// <returns></returns>
        public static string LocationXML()
        {
            Utils.Log(string.Format("UtilsObjectsLocation.LocationXML started...", ""));

            XmlDocument xmlDoc = new XmlDocument();

            string folderPath = Path.GetDirectoryName(DocumentManager.Instance.CurrentDBDocument.PathName);

            if (folderPath.StartsWith("BIM 360:"))
            {
                folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            }

            Utils.Log(string.Format("Folder: {0}", folderPath));

            string docName = DocumentManager.Instance.CurrentDBDocument.Title;
            docName = Path.ChangeExtension(docName, "xml");

            Utils.Log(string.Format("XML Document: {0}", docName));

            string path = Path.Combine(folderPath, docName);

            if (File.Exists(path))
            {
                Utils.Log(string.Format("Existing File to Override: {0}", path));
                File.Delete(path);
            }

            var project = xmlDoc.CreateElement("Project");
            var objects = xmlDoc.CreateElement("Objects");
            var pointObjects = xmlDoc.CreateElement("PointObjects");
            var lineObjects = xmlDoc.CreateElement("LineObjects");
            var multiPointObjects = xmlDoc.CreateElement("MultiPointObjects");

            objects.AppendChild(pointObjects);
            objects.AppendChild(lineObjects);
            objects.AppendChild(multiPointObjects);
            project.AppendChild(objects);
            xmlDoc.AppendChild(project);

            var name = xmlDoc.CreateAttribute("name");
            var user = xmlDoc.CreateAttribute("user");
            var date = xmlDoc.CreateAttribute("date");
            project.SetAttribute("name", DocumentManager.Instance.CurrentDBDocument.PathName);
            project.SetAttribute("user", DocumentManager.Instance.CurrentUIApplication.Application.Username);
            project.SetAttribute("date", DateTime.Now.ToString());

            foreach (Element e in new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)  // 1.1.0
           .OfClass(typeof(FamilyInstance))
           .WhereElementIsNotElementType()
           .Where(x => x.Category.CategoryType == CategoryType.Model))  // 1.1.1
            {
                var par = e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.MultiPoint.Name);

                if (par != null)  // 1.1.0
                {
                    if (!par.HasValue || par.AsString() == "")
                    {
                        continue;
                    }
                }

                var fi = e as FamilyInstance;

                if (AdaptiveComponentInstanceUtils.IsAdaptiveComponentInstance(fi))
                {
                    var multiPointObject = xmlDoc.CreateElement("MultiPointObject");
                    multiPointObject.InnerText = e.UniqueId;
                    multiPointObjects.AppendChild(multiPointObject);
                    Utils.Log(string.Format("MultiPoint: {0}", fi.Id));
                }
            }

            foreach (Element e in new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
            .OfClass(typeof(FamilyInstance))
            .WhereElementIsNotElementType()
            .Where(x => x.Category.CategoryType == CategoryType.Model))
            {
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Corridor.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.BaselineIndex.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Code.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Side.Name).HasValue)  // 1.1.0
                {
                    continue;
                }

                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Station.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Offset.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Elevation.Name).HasValue)  // 1.1.0
                {
                    continue;
                }

                var fi = e as FamilyInstance;

                if (fi.Location is LocationPoint)
                {
                    if (e.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Columns) || e.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_StructuralColumns))
                    {
                        var lineObject = xmlDoc.CreateElement("LineObject");
                        lineObject.InnerText = e.UniqueId;
                        lineObjects.AppendChild(lineObject);
                        Utils.Log(string.Format("Point Object: {0}", fi.Id));
                    }
                    else
                    {
                        if (!e.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_PipeFitting) ||
                            !e.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_DuctFitting) ||
                            !e.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_ConduitFitting) ||
                            !e.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_CableTrayFitting))  // 1.1.1
                        {
                            var pointObject = xmlDoc.CreateElement("PointObject");
                            pointObject.InnerText = e.UniqueId;
                            pointObjects.AppendChild(pointObject);
                            Utils.Log(string.Format("Point Object: {0}", fi.Id));
                        }
                    }
                }
                else if (fi.Location is LocationCurve)
                {
                    var lineObject = xmlDoc.CreateElement("LineObject");
                    lineObject.InnerText = e.UniqueId;
                    lineObjects.AppendChild(lineObject);
                    Utils.Log(string.Format("Linear Object: {0}", fi.Id));
                }
            }

            foreach (Element e in new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
                        .OfClass(typeof(MEPCurve))
                        .WhereElementIsNotElementType())
            {
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Corridor.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.BaselineIndex.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Code.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Side.Name).HasValue)  // 1.1.0
                {
                    continue;
                }

                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Station.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Offset.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.Elevation.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.EndStation.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.EndOffset.Name).HasValue)  // 1.1.0
                {
                    continue;
                }
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.EndElevation.Name).HasValue)  // 1.1.0
                {
                    continue;
                }

                var lineObject = xmlDoc.CreateElement("LineObject");
                lineObject.InnerText = e.UniqueId;
                lineObjects.AppendChild(lineObject);
                Utils.Log(string.Format("Linear Object: {0}", e.Id));
            }

            foreach (Element e in new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
                    .OfClass(typeof(Floor))
                    .WhereElementIsNotElementType())
            {
                if (!e.Parameters.Cast<Parameter>().First(x => x.Definition.Name == ADSK_Parameters.Instance.MultiPoint.Name).HasValue)  // 1.1.0
                {
                    continue;
                }

                var multiPointObject = xmlDoc.CreateElement("MultiPointObject");
                multiPointObject.InnerText = e.UniqueId;
                multiPointObjects.AppendChild(multiPointObject);
                Utils.Log(string.Format("MultiPoint: {0}", e.Id));
            }

            xmlDoc.Save(path);

            Utils.Log(string.Format("XML Saved...", ""));

            Utils.Log(string.Format("UtilsObjectsLocation.LocationXML completed.", ""));

            return path;
        }

        /// <summary>
        /// Udpates the document from XML.
        /// </summary>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="elements">The elements.</param>
        /// <returns></returns>
        public static IList<Revit.Elements.Element> UdpateDocumentFromXML(CivilDocument civilDocument, Revit.Elements.Element[] elements = null)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.UdpateDocumentFromXML started...", ""));

            IList<Revit.Elements.Element> output = new List<Revit.Elements.Element>();

            if (elements == null)
            {
                XmlDocument xmlDoc = new XmlDocument();

                string folderPath = Path.GetDirectoryName(DocumentManager.Instance.CurrentDBDocument.PathName);

                if (folderPath.StartsWith("BIM 360:"))
                {
                    folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                }

                string docName = DocumentManager.Instance.CurrentDBDocument.Title;
                docName = Path.ChangeExtension(docName, "xml");

                string path = Path.Combine(folderPath, docName);

                if (!File.Exists(path))
                {
                    Utils.Log(string.Format("File to be created: {0}", path));

                    LocationXML();
                }

                Utils.Log(string.Format("File created: {0}", path));

                try
                {
                    xmlDoc.Load(path);
                }
                catch (Exception ex)
                {
                    string message = string.Format("Error loading XML. {0}\n{1}", path, ex.Message);

                    Utils.Log(message);

                    throw new Exception(message);
                }

                if (!xmlDoc.HasChildNodes)
                {
                    string message = string.Format("Error loading XML. {0}", path);

                    Utils.Log(message);

                    throw new Exception(message);
                }

                Utils.Log("XML Loaded...");

                // Collect existing objects
                IList<Revit.Elements.Element> pointObjects = new List<Revit.Elements.Element>();
                IList<Revit.Elements.Element> lineObjects = new List<Revit.Elements.Element>();

                foreach (XmlElement n in xmlDoc.GetElementsByTagName("PointObject"))
                {
                    try
                    {
                        var uid = n.InnerText;
                        var e = Revit.Elements.ElementSelector.ByUniqueId(uid, true);
                        if (e != null)
                        {
                            pointObjects.Add(e);
                        }
                    }
                    catch
                    {

                        continue;
                    }
                }

                foreach (XmlElement n in xmlDoc.GetElementsByTagName("LineObject"))
                {
                    try
                    {
                        var uid = n.InnerText;
                        var e = Revit.Elements.ElementSelector.ByUniqueId(uid, true);
                        if (e != null)
                        {
                            lineObjects.Add(e);
                        }
                    }
                    catch
                    {

                        continue;
                    }
                }

                if (pointObjects.Count > 0)
                {

                    var a = OptimizedUdpateObjectFromXML(civilDocument, pointObjects);

                    output.Concat(a);

                }

                if (lineObjects.Count > 0)
                {
                    output.Concat(OptimizedUdpateObjectFromXML(civilDocument, lineObjects));
                }
            }
            else
            {
                foreach (var e in elements)
                {
                    output.Add(TestUdpateObject(civilDocument, e));
                }
            }

            Utils.Log(string.Format("UtilsObjectsLocation.UdpateDocumentFromXML completed.", ""));

            return output;
        }

        internal class Selectable
        {
            string _corridor;
            int _bi;
            string _code;
            string _side;
            int _ri;  // 1.1.0
            double _rr; // 1.1.0
            double _rn;  // 1.1.0
            double _err;  // 1.1.0
            double _ern;  // 1.1.0
            Revit.Elements.Element _e;

            public string Corridor { get { return this._corridor; } set { this._corridor = value; } }
            public int BaselineIndex { get { return this._bi; } set { this._bi = value; } }
            public string Code { get { return this._code; } set { this._code = value; } }
            public string Side { get { return this._side; } set { this._side = value; } }
            public int RegionIndex { get { return this._ri; } set { this._ri = value; } }  // 1.1.0
            public double RegionRelative { get { return this._rr; } set { this._rr = value; } }  // 1.1.0
            public double RegionNormalized { get { return this._rn; } set { this._rn = value; } }  // 1.1.0
            public double EndRegionRelative { get { return this._err; } set { this._err = value; } }  // 1.1.0
            public double EndRegionNormalized { get { return this._ern; } set { this._ern = value; } }  // 1.1.0
            public Revit.Elements.Element Element { get { return _e; } set { this._e = value; } }

            internal Selectable() { }

            internal Selectable(Revit.Elements.Element e)
            {
                this._corridor = (string)e.GetParameterValueByName(ADSK_Parameters.Instance.Corridor.Name);
                this._bi = Convert.ToInt32(e.GetParameterValueByName(ADSK_Parameters.Instance.BaselineIndex.Name));
                this._code = (string)e.GetParameterValueByName(ADSK_Parameters.Instance.Code.Name);
                this._side = (string)e.GetParameterValueByName(ADSK_Parameters.Instance.Side.Name);
                this._ri = (int)e.GetParameterValueByName(ADSK_Parameters.Instance.RegionIndex.Name);  // 1.1.0
                this._rr = (double)e.GetParameterValueByName(ADSK_Parameters.Instance.RegionRelative.Name);  // 1.1.0
                this._rn = (double)e.GetParameterValueByName(ADSK_Parameters.Instance.RegionNormalized.Name);  // 1.1.0
                this._err = this._rr;  // 1.1.0
                try
                {
                    this._err = (double)e.GetParameterValueByName(ADSK_Parameters.Instance.EndRegionRelative.Name);  // 1.1.0
                }
                catch { }

                this._ern = this._rn;
                try
                {
                    this._ern = (double)e.GetParameterValueByName(ADSK_Parameters.Instance.EndRegionNormalized.Name);  // 1.1.0
                }
                catch { }

                this._e = e;
            }
        }

        /// <summary>
        /// Udpates the object from XML.
        /// </summary>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="elements">The element.</param>
        /// <param name="normalized">Boolean value to proces based on the normalized values.</param>
        /// <returns></returns>
        /// <remarks>If the code is "None" the Baseline will be used to calculate the location.</remarks>
        public static IList<Revit.Elements.Element> OptimizedUdpateObjectFromXML(CivilDocument civilDocument, IList<Revit.Elements.Element> elements, bool normalized = false)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.OptimizedUdpateObjectFromXML started...", ""));

            #region SETUP
            XmlDocument xmlDoc = new XmlDocument();
            Document doc = DocumentManager.Instance.CurrentDBDocument;

            var totalTransform = RevitUtils.DocumentTotalTransform();

            Dictionary<Featureline, IList<Revit.Elements.Element>> dictionary = new Dictionary<Featureline, IList<Revit.Elements.Element>>();

            IList<Revit.Elements.Element> excluded = new List<Revit.Elements.Element>();

            string folderPath = Path.GetDirectoryName(DocumentManager.Instance.CurrentDBDocument.PathName);

            if (folderPath.StartsWith("BIM 360:"))
            {
                folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            }

            string docName = DocumentManager.Instance.CurrentDBDocument.Title;
            docName = Path.ChangeExtension(docName, "xml");

            //if (folderPath.StartsWith("BIM 360:\\"))
            //{
            //    folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            //}
            // TODO Add desktop connector

            string path = Path.Combine(folderPath, docName);

            if (!File.Exists(path))
            {
                Utils.Log(string.Format("1. File to be created: {0}", path));
                LocationXML();
            }

            // Use only the items in the XML file

            try
            {
                xmlDoc.Load(path);
            }
            catch (Exception ex)
            {
                string message = string.Format("Error loading XML. {0}\n{1}", path, ex.Message);

                Utils.Log(message);

                throw new Exception(message);
            }

            if (!xmlDoc.HasChildNodes)
            {
                string message = string.Format("Error loading XML. {0}", path);

                Utils.Log(message);

                throw new Exception(message);
            }

            Utils.Log("XML Loaded...");

            string multipoint = "<None>";

            IList<Revit.Elements.Element> currentElements = new List<Revit.Elements.Element>();
            IList<Revit.Elements.Element> currentLineElements = new List<Revit.Elements.Element>();  // 1.1.0
            IList<Revit.Elements.Element> currentMPElements = new List<Revit.Elements.Element>();

            IList<string> guids = new List<string>();
            IList<string> line_guids = new List<string>();  // 1.1.0
            IList<string> mp_guids = new List<string>();

            foreach (XmlElement node in xmlDoc.GetElementsByTagName("PointObject"))
            {
                guids.Add(node.InnerText);
            }

            foreach (XmlElement node in xmlDoc.GetElementsByTagName("LineObject"))
            {
                guids.Add(node.InnerText);
                line_guids.Add(node.InnerText);  // 1.1.0
            }

            foreach (XmlElement node in xmlDoc.GetElementsByTagName("MultiPointObject"))
            {
                mp_guids.Add(node.InnerText);
            }

            foreach (Revit.Elements.Element e in elements)
            {
                if (guids.Contains(e.UniqueId))
                {
                    currentElements.Add(e);
                }
               
            }

            foreach (Revit.Elements.Element e in elements)
            {
                if (line_guids.Contains(e.UniqueId))
                {
                    currentLineElements.Add(e);
                }
                
            }

            foreach (Revit.Elements.Element e in elements)
            {
                if (mp_guids.Contains(e.UniqueId))
                {
                    currentMPElements.Add(e);
                }
                
            }
            #endregion

            // DEBUG
            //System.Windows.Forms.MessageBox.Show(string.Format("Guids: {0}", currentElements.Count));

            // No element in the XML Location file
            if (currentElements.Count == 0 && currentMPElements.Count == 0)
            {
                return elements;
            }

            #region OPTIMIZE
            IList<Selectable> selectables = new List<Selectable>();

            // Create Selectable obejcts
            foreach (var e in currentElements)
            {
                selectables.Add(new Selectable(e));
            }

            if (selectables.Count > 0)
            {
                var collection = selectables.GroupBy(sel => new { sel.Corridor, sel.BaselineIndex, sel.RegionIndex, sel.Code, sel.Side }).Select(x => x.ToList()).ToList();

                // Calculate the featureline fore each group
                foreach (var list in collection)
                {
                    #region GET_FEATURELINE

                    Corridor corridor = null;

                    try
                    {
                        corridor = civilDocument.GetCorridors().First(x => x.Name == list[0].Corridor);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(string.Format("Corridor {0} could not be found.\n{1}", list[0].Corridor, ex.Message));  // 1.1.0
                    }

                    if (corridor == null)
                    {
                        excluded.Concat(list.Select(x => x.Element).ToList());
                        continue;  // 1.1.0 
                    }

                    int bi = list[0].BaselineIndex;

                    Baseline bl = null;

                    try
                    {
                        bl = corridor.Baselines[bi];
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(string.Format("Corridor {0} doesn't have a Baseline with the specified index {1}.\n{2}", list[0].Corridor, bi, ex.Message));  // 1.1.0
                    }

                    if (bl == null)
                    {
                        excluded.Concat(list.Select(x => x.Element).ToList());
                        continue;  // 1.1.0 TODO Add all lists to excluded 
                    }

                    double[] stations = bl.Stations;

                    double min = stations.Min();
                    double max = stations.Max();

                    string code = list[0].Code;
                    string side = list[0].Side;
                    int ri = 0; // 1.1.0

                    try
                    {
                        ri = list[0].RegionIndex;  // 1.1.0 this is to make it work on older versions
                    }
                    catch { }

                    BaselineRegion blr = null;

                    foreach (BaselineRegion reg in bl.GetBaselineRegions())
                    {
                        if (reg != null)
                        {
                            if (reg.Index == ri)
                            {
                                blr = reg;
                                Utils.Log(string.Format("BaselineRegion ok...", ""));
                                break;
                            }
                        }
                    }

                    double eleStation = (double)list[0].Element.GetParameterValueByName(ADSK_Parameters.Instance.Station.Name);  // 1.1.0

                    // for each group it was used one featureline
                    Featureline tempFeat = null;

                    if (blr == null)  // the region index isnot valid anymore -> use the closest region
                    {
                        blr = bl.GetBaselineRegions().Where(x => x != null).OrderBy(r => Math.Abs(eleStation - 0.5 * (r.End + r.Start))).First(g => g.Stations.Count() > 0); // 1.1.0
                    }

                    // Scenario 1: the featureline is compatible with the parameters of the list and all the elements are in range
                    // Scenario 2: the featureline is compatible with the parameters of the list and some elements are in range
                    // Scenario 3: the featureline is compatible with the parameters of the list and no elements are in range
                    // Scenario 4: the featureline is not fully compatible with the parameters of the list
                    tempFeat = bl.GetFeaturelinesByCodeStation(code, blr.Start + 0.5 * blr.RelativeEnd).First(x => x.Side.ToString() == side);

                    Utils.Log(string.Format("Featureline : {0}", tempFeat == null ? "is null !!!" : tempFeat.Code));

                    if (tempFeat != null)
                    {
                        foreach (var element in list)
                        {
                            double tempStation = (double)element.Element.GetParameterValueByName(ADSK_Parameters.Instance.Station.Name);  // 1.1.0

                            Utils.Log(string.Format("Station: {0}", tempStation));

                            if (element.RegionNormalized > 1 || element.RegionNormalized < 0)
                            {
                                tempFeat = bl.GetFeaturelinesByCodeStation(code, tempStation).First(x => x.Side.ToString() == side);
                            }

                            bool found = false;

#pragma warning disable CS0219 // The variable 'inRange' is assigned but its value is never used
                            bool inRange = false;
#pragma warning restore CS0219 // The variable 'inRange' is assigned but its value is never used

#pragma warning disable CS0219 // The variable 'regSystem' is assigned but its value is never used
                            bool regSystem = false;
#pragma warning restore CS0219 // The variable 'regSystem' is assigned but its value is never used

                            bool skip = false;

                            if (tempFeat.Baseline.CorridorName == element.Corridor &&
                                 tempFeat.Baseline.Index == element.BaselineIndex &&
                                 tempFeat.Code == element.Code &&
                                 tempFeat.Side.ToString() == element.Side &&
                                 tempFeat.BaselineRegionIndex == element.RegionIndex)  // 1.1.0
                            {
                                // Scenarios 1, 2, 3
                                found = true;

                                if (tempFeat.Start <= tempStation || Math.Abs(tempFeat.Start - tempStation) < 0.00001 || tempFeat.End >= tempStation || Math.Abs(tempFeat.End - tempStation) < 0.00001)
                                {
                                    inRange = true;  // Scenarios 1 and 2

                                    Utils.Log(string.Format("Station in range...", ""));

                                    if (normalized)
                                    {
                                        double newStation = tempFeat.Start + element.RegionNormalized * (tempFeat.End - tempFeat.Start);

                                        element.Element.SetParameterByName(ADSK_Parameters.Instance.Station.Name, newStation);

                                        Utils.Log(string.Format("New station normalized: {0}", newStation));
                                    }
                                }
                                else
                                {
                                    if (!normalized)
                                    {
                                        if (element.RegionRelative <= blr.RelativeEnd) // Relative is in range
                                        {
                                            regSystem = true;
                                            element.Element.SetParameterByName(ADSK_Parameters.Instance.Station.Name, tempFeat.Start + element.RegionRelative);

                                            Utils.Log(string.Format("New station relative: {0}", tempFeat.Start + element.RegionRelative));
                                        }
                                        else  // Scenarios  2 and 3
                                        {
                                            excluded.Add(element.Element);
                                            skip = true;
                                            Utils.Log(string.Format("Excluded: {0}", element.Element.Id));
                                        }
                                    }
                                    else
                                    {
                                        // Scenarios 1 and 2
                                        // TODO Normalized
                                        double newStation = tempFeat.Start + element.RegionNormalized * (tempFeat.End - tempFeat.Start);

                                        element.Element.SetParameterByName(ADSK_Parameters.Instance.Station.Name, newStation);

                                        Utils.Log(string.Format("New station normalized: {0}", newStation));
                                    }
                                }

                                if (!skip)
                                {
                                    if (currentLineElements.Contains(element.Element))  // The element is Linear so it has End parameters to check as well.
                                    {
                                        double tempEndStation = (double)element.Element.GetParameterValueByName(ADSK_Parameters.Instance.EndStation.Name);  // 1.1.0

                                        if (element.RegionNormalized > 1 || element.RegionNormalized < 0)  // More strict for Linear Objects as they could break systems in multiple places
                                        {
                                            excluded.Add(element.Element);
                                            skip = true;
                                        }

                                        if (tempFeat.Start <= tempEndStation || Math.Abs(tempFeat.Start - tempEndStation) < 0.00001 || tempFeat.End >= tempEndStation || Math.Abs(tempFeat.End - tempEndStation) < 0.00001)
                                        {
                                            inRange = true;  // Scenarios 1 and 2

                                            if (normalized)
                                            {
                                                element.Element.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, tempFeat.Start + element.EndRegionNormalized * (tempFeat.End - tempFeat.Start));
                                            }
                                        }
                                        else
                                        {
                                            if (!normalized)
                                            {
                                                if (element.EndRegionRelative <= blr.RelativeEnd) // Relative is in range
                                                {
                                                    regSystem = true;
                                                    element.Element.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, tempFeat.Start + element.EndRegionRelative);
                                                }
                                                else  // Scenarios  2 and 3
                                                {
                                                    excluded.Add(element.Element);
                                                    skip = true;
                                                }
                                            }
                                            else
                                            {
                                                // Scenarios 1 and 2
                                                // TODO Normalized
                                                element.Element.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, tempFeat.Start + element.EndRegionNormalized * (tempFeat.End - tempFeat.Start));
                                            }
                                        }
                                    }
                                }
                            }
                            else  // The region parameters are not the fully compatible 
                            {
                                Utils.Log(string.Format("Featureline is not fully compatible...", ""));

                                // in range
                                if (tempFeat.Start <= tempStation || Math.Abs(tempFeat.Start - tempStation) < 0.00001 || tempFeat.End >= tempStation || Math.Abs(tempFeat.End - tempStation) < 0.00001)
                                {
                                    inRange = true;  // Scenarios 1 and 2
                                    found = true;

                                    if (normalized)
                                    {
                                        double newStation = tempFeat.Start + element.RegionNormalized * (tempFeat.End - tempFeat.Start);

                                        element.Element.SetParameterByName(ADSK_Parameters.Instance.Station.Name, newStation);

                                        Utils.Log(string.Format("New station normalized: {0}", newStation));
                                    }

                                    element.Element.SetParameterByName(ADSK_Parameters.Instance.RegionIndex.Name, tempFeat.BaselineRegionIndex);

                                    Utils.Log(string.Format("New region index: {0}", tempFeat.BaselineRegionIndex));
                                }
                                else  // not in range
                                {
                                    if (!normalized)
                                    {
                                        if (element.RegionRelative <= blr.RelativeEnd) // Relative is in range
                                        {
                                            regSystem = true;
                                            found = true;
                                            element.Element.SetParameterByName(ADSK_Parameters.Instance.Station.Name, tempFeat.Start + element.RegionRelative);
                                            element.Element.SetParameterByName(ADSK_Parameters.Instance.RegionIndex.Name, tempFeat.BaselineRegionIndex);

                                            Utils.Log(string.Format("New station relative: {0}", tempFeat.Start + element.RegionRelative));
                                            Utils.Log(string.Format("New region index: {0}", tempFeat.BaselineRegionIndex));
                                        }
                                        else  // Scenarios  2 and 3
                                        {
                                            excluded.Add(element.Element);
                                            skip = true;
                                        }
                                    }
                                    else
                                    {
                                        found = true;
                                        element.Element.SetParameterByName(ADSK_Parameters.Instance.RegionIndex.Name, tempFeat.BaselineRegionIndex);

                                        double newStation = tempFeat.Start + element.RegionNormalized * (tempFeat.End - tempFeat.Start);

                                        element.Element.SetParameterByName(ADSK_Parameters.Instance.Station.Name, newStation);

                                        Utils.Log(string.Format("New station normalized: {0}", newStation));
                                        Utils.Log(string.Format("New region index: {0}", tempFeat.BaselineRegionIndex));
                                        // Scenarios 1 and 2
                                        // TODO Normalized
                                    }
                                }

                                if (!skip)
                                {
                                    if (currentLineElements.Contains(element.Element))  // The element is Linear so it has End parameters to check as well.
                                    {
                                        double tempEndStation = (double)element.Element.GetParameterValueByName(ADSK_Parameters.Instance.EndStation.Name);  // 1.1.0

                                        if (tempFeat.Start <= tempEndStation || Math.Abs(tempFeat.Start - tempEndStation) < 0.00001 || tempFeat.End >= tempEndStation || Math.Abs(tempFeat.End - tempEndStation) < 0.00001)
                                        {
                                            inRange = true;  // Scenarios 1 and 2
                                            found = true;

                                            if (normalized)
                                            {
                                                element.Element.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, tempFeat.Start + element.EndRegionNormalized * (tempFeat.End - tempFeat.Start));
                                            }
                                        }
                                        else
                                        {
                                            if (!normalized)
                                            {
                                                if (element.EndRegionRelative <= blr.RelativeEnd) // Relative is in range
                                                {
                                                    regSystem = true;
                                                    found = true;
                                                    element.Element.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, tempFeat.Start + element.EndRegionRelative);
                                                }
                                                else  // Scenarios  2 and 3
                                                {
                                                    excluded.Add(element.Element);
                                                    skip = true;
                                                }
                                            }
                                            else
                                            {
                                                // Scenarios 1 and 2
                                                // TODO Normalized
                                                found = true;
                                                element.Element.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, tempFeat.Start + element.EndRegionNormalized * (tempFeat.End - tempFeat.Start));
                                            }
                                        }
                                    }
                                }
                            }

                            if (found && !skip)
                            {
                                if (!dictionary.ContainsKey(tempFeat))
                                {
                                    dictionary.Add(tempFeat, new List<Revit.Elements.Element>() { element.Element });
                                }
                                else
                                {
                                    dictionary[tempFeat].Add(element.Element);
                                }
                            }
                        }
                    }
                    else
                    {
                        excluded.Concat(list.Select(x => x.Element).ToList());
                    }

                    // Check if the featureline is still valid
                    #endregion
                }
            }
            #endregion

            // Send the exluded Ids to the clipboard to select them in Revit
            if (excluded.Count > 0)
            {
                string excludedId = string.Join(";", excluded.Select(x => x.Id.ToString()));
                Utils.Log(string.Format("Excluded objects:\n{0}", excludedId));

                System.Windows.Forms.Clipboard.SetText(excludedId);
            }

            #region MULTIPOINT
            // if there are multi point objects

            // Featureline optimization for MultiPoint objects
            Dictionary<string, Featureline> multiPointFeaturelines = new Dictionary<string, Featureline>();

            Dictionary<string, Tuple<string, int, int, double, double, double, string, Tuple<string>>> tuples = new Dictionary<string, Tuple<string, int, int, double, double, double, string, Tuple<string>>>();


            if (currentMPElements.Count > 0)
            {
                var parameters = currentMPElements.Select(x => (string)x.GetParameterValueByName(ADSK_Parameters.Instance.MultiPoint.Name)).ToList();

                foreach (var s in parameters)
                {
                    if (s != "" && s != null)
                    {
                        Newtonsoft.Json.Linq.JObject obj = JObject.Parse(s);

                        foreach (ShapePoint sp in obj["ShapePoints"].ToObject<ShapePointArray>().Points)
                        {
                            // Complete the optimization
                            Tuple<string, int, int, double, double, double, string, Tuple<string>> tuple = new Tuple<string, int, int, double, double, double, string, Tuple<string>>(
                                sp.Corridor, sp.BaselineIndex, sp.RegionIndex, sp.Station, sp.RegionRelative, sp.RegionNormalized, sp.Code, new Tuple<string>(sp.Side.ToString()));

                            tuples[sp.UniqueId] = tuple;
                        }
                    }
                }
            }

            var uniquevalues = tuples.Values.Distinct().ToList();

            for (int i = 0; i < uniquevalues.Count; ++i)
            {
                var set = uniquevalues[i];

                Corridor corridor = null;

                try
                {
                    corridor = civilDocument.GetCorridors().First(x => x.Name == set.Item1);
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Corridor {0} could not be found.\n{1}", set.Item1, ex.Message));  // 1.1.0
                }

                if (corridor == null)
                {
                    continue;  // 1.1.0 TODO Add all lists to excluded 
                }


                Baseline bl = null;

                try
                {
                    bl = corridor.Baselines[set.Item2];
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Corridor {0} doesn't have a Baseline with the specified index {1}.\n{2}", set.Item1, set.Item2, ex.Message));  // 1.1.0
                }

                if (bl == null)
                {
                    continue;  // 1.1.0 TODO Add all lists to excluded 
                }

                BaselineRegion blr = null;

                foreach (BaselineRegion reg in bl.GetBaselineRegions())
                {
                    if (reg != null)
                    {
                        if (reg.Index == set.Item3)
                        {
                            blr = reg;
                            Utils.Log(string.Format("BaselineRegion ok...", ""));
                            break;
                        }
                    }
                }

                Featureline tempFeat = null;

                if (blr == null)  // the region index isnot valid anymore -> use the closest region
                {
                    blr = bl.GetBaselineRegions().Where(x => x != null).OrderBy(r => Math.Abs(set.Item4 - 0.5 * (r.End + r.Start))).First(g => g.Stations.Count() > 0); // 1.1.0
                }

                tempFeat = bl.GetFeaturelinesByCodeStation(set.Item7, blr.Start + 0.5 * blr.RelativeEnd).First(x => x.Side.ToString() == set.Rest.Item1);  // 1.1.0

                if (tempFeat != null)
                {
                    foreach (var pair in tuples)
                    {
                        if (pair.Value.Equals(set))
                        {
                            multiPointFeaturelines[pair.Key] = tempFeat;
                        }
                    }
                }

            }

            foreach (Revit.Elements.Element element in currentMPElements)
            {
                // Delete objects to be removed
                if ((int)element.GetParameterValueByName(ADSK_Parameters.Instance.Delete.Name) == 1)
                {
                    RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

                    DocumentManager.Instance.CurrentDBDocument.Delete(element.InternalElement.Id);

                    RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                    return null;
                }

                else if ((int)element.GetParameterValueByName(ADSK_Parameters.Instance.Update.Name) == 1)
                {
                    // Update the remaining objects
                    try
                    {
                        multipoint = (string)element.GetParameterValueByName(ADSK_Parameters.Instance.MultiPoint.Name);
                    }
                    catch { }

                    if (multipoint != "<None>" && multipoint != "" && multipoint != null)
                    {
                        // Adaptive Component
                        // Floor
                        // Potentially walls with edited profile
                        try
                        {
                            MultiPoint mp = Newtonsoft.Json.JsonConvert.DeserializeObject<MultiPoint>(multipoint);

                            if (null == mp)
                            {
                                throw new Exception(string.Format("The MultiPoint cannot be updated\n\n{0}", element.Id));
                            }

                            foreach (ShapePoint sp in mp.ShapePoints.Points)
                            {
                                Featureline f = multiPointFeaturelines[sp.UniqueId];

                                sp.UpdateByFeatureline(f);  // 1.1.0 TODO Check the normalized and relative region logic in here

                            }

                            Element internalElement = element.InternalElement;

                            if (internalElement is Floor)
                            {
                                Revit.Elements.Floor floor = element as Revit.Elements.Floor;
                                Revit.Elements.FloorType ft = Revit.Elements.FloorType.ByName(floor.Name);
                                var levelId = element.InternalElement.Parameters.Cast<Parameter>().First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.LEVEL_PARAM)).AsElementId();  // 1.1.0
                                var levelRvt = doc.GetElement(levelId) as Level;  // 1.1.0
                                var level = Revit.Elements.InternalUtilities.ElementQueries.OfElementType(typeof(Level)).Cast<Revit.Elements.Level>().First(x => x.Name == levelRvt.Name);  // 1.1.0

                                // It's not possible to update the boundary of a floor via API, only create a new one

                                RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

                                var fl = mp.ToFloor(ft, level);

                                RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                                // Copy all instance parameters on the new floor
                                foreach (var par in element.Parameters)
                                {
                                    if (par.Name != ADSK_Parameters.Instance.MultiPoint.Name)  // 1.1.0
                                    {
                                        if (!par.IsReadOnly)
                                        {
                                            try
                                            {
                                                fl.SetParameterByName(par.Name, element.GetParameterValueByName(par.Name));
                                            }
                                            catch (Exception) { }
                                        }
                                    }
                                }

                                // For some reason when processing a list of floors only the last one is passed to the Revit document
                                // I've created a copy of and deleted the original and the process
                                // the only one that stays in the Revit document is the last copy
                                RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

                                // without this it just deletes the new floors and doesn't create the copy, not sure why
                                var todelete = ElementTransformUtils.CopyElement(DocumentManager.Instance.CurrentDBDocument, internalElement.Id, XYZ.Zero);

                                DocumentManager.Instance.CurrentDBDocument.Delete(internalElement.Id);

                                ElementTransformUtils.CopyElement(DocumentManager.Instance.CurrentDBDocument, fl.InternalElement.Id, XYZ.Zero);

                                DocumentManager.Instance.CurrentDBDocument.Delete(todelete);

                                DocumentManager.Instance.CurrentDBDocument.Delete(fl.InternalElement.Id);

                                RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                            }
                            else if (internalElement is FamilyInstance)
                            {
                                FamilyInstance fi = internalElement as FamilyInstance;

                                if (AdaptiveComponentInstanceUtils.IsAdaptiveComponentInstance(fi))
                                {
                                    var pnts = AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(fi);
                                    IList<XYZ> refPts = mp.ShapePoints.Points.OrderBy(p => p.Id).Select(x => x.RevitPoint.ToXyz()).ToList();

                                    if (refPts.Count != pnts.Count)
                                    {
                                        throw new Exception("Point number are mismatching");
                                    }

                                    RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

                                    for (int i = 0; i < pnts.Count; ++i)
                                    {
                                        ReferencePoint rp = doc.GetElement(pnts[i]) as ReferencePoint;

                                        rp.Position = refPts[i];
                                    }

                                    element.SetParameterByName(ADSK_Parameters.Instance.MultiPoint.Name, mp.SerializeJSON());  // 1.1.0  If the ShapePoint objects have been updated they need to be stored on the element
                                    element.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);  // 1.1.0  
                                    element.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);  // 1.1.0  

                                    RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("CivilConnection", string.Format("{0}", ex.Message));
                        }
                    }
                    else
                    {
                        continue;  // 1.1.0 When selecting objects that don't have this parameter compiled.
                    }
                }
            }

            #endregion

            // Process the objects
            foreach (var pair in dictionary)
            {
                Featureline featureline = pair.Key;

                foreach (Revit.Elements.Element element in pair.Value)
                {
                    // Delete objects to be removed
                    if ((int)element.GetParameterValueByName(ADSK_Parameters.Instance.Delete.Name) == 1)
                    {
                        RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

                        DocumentManager.Instance.CurrentDBDocument.Delete(element.InternalElement.Id);

                        RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                        return null;
                    }

                    else if ((int)element.GetParameterValueByName(ADSK_Parameters.Instance.Update.Name) == 1)
                    {
                        multipoint = "";

                        if (multipoint != "<None>" && multipoint != "" && multipoint != null)
                        {
                           
                        }

                        else
                        {
                            #region GET_PARAMETERS
                            double station = 0;
                            try
                            {
                                station = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Station.Name);
                            }
                            catch { }

                            double offset = 0;
                            try
                            {
                                offset = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Offset.Name);
                            }
                            catch { }

                            double elevation = 0;
                            try
                            {
                                elevation = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Elevation.Name);
                            }
                            catch { }

                            double angle = 0;

                            try
                            {
                                angle = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.AngleZ.Name);
                            }
                            catch { }

                            double endStation = station;
                            try
                            {
                                endStation = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndStation.Name);
                            }
                            catch { }

                            double endOffset = offset;
                            try
                            {
                                endOffset = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndOffset.Name);
                            }
                            catch { }

                            double endElevation = elevation;
                            try
                            {
                                endElevation = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndElevation.Name);
                            }
                            catch { }
                            #endregion

                            if (Math.Abs(station - endStation) < 0.0001 && Math.Abs(offset - endOffset) < 0.0001 && Math.Abs(elevation - endElevation) < 0.0001)
                            {
                                #region POINT_OBJECTS

                                if (station > featureline.End || station < featureline.Start)  // [20190923] 1.1.0 
                                {
                                    if (!normalized)
                                    {
                                        station = featureline.Start + (double)element.GetParameterValueByName(ADSK_Parameters.Instance.RegionRelative.Name);
                                    }
                                    else
                                    {
                                        station = featureline.Start + (double)element.GetParameterValueByName(ADSK_Parameters.Instance.RegionNormalized.Name) * (featureline.End - featureline.Start);
                                    }
                                }

                                CoordinateSystem cs = featureline.CoordinateSystemByStation(station);

                                var localPoint = featureline.PointByStationOffsetElevation(station, offset, elevation, false);  // WCS
                                var point = localPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;  // LCS

                                if (element is Revit.Elements.FamilyInstance)
                                {
                                    Revit.Elements.FamilyInstance fi = element as Revit.Elements.FamilyInstance;

                                    var lp = fi.InternalElement.Location as LocationPoint;  // LCS

                                    RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(doc);

                                    lp.Move(point.ToXyz() - lp.Point);

                                    var currentRotation = lp.Rotation;

                                    var xAxis = cs.XAxis.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Vector;

                                    lp.Rotate(Autodesk.Revit.DB.Line.CreateBound(lp.Point, lp.Point + XYZ.BasisZ),
                                        -currentRotation
                                        + DegToRadians(angle
                                        
                                        - xAxis.AngleAboutAxis(Vector.XAxis(), Vector.ZAxis())
                                        ));

                                    RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                                    double regRelative = station - featureline.Start;
                                    double regNormalized = regRelative / (featureline.End - featureline.Start);

                                    fi.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(localPoint.X, 3));
                                    fi.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(localPoint.Y, 3));
                                    fi.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(localPoint.Z, 3));
                                    fi.SetParameterByName(ADSK_Parameters.Instance.RegionIndex.Name, featureline.BaselineRegionIndex);  // 1.1.0
                                    fi.SetParameterByName(ADSK_Parameters.Instance.RegionRelative.Name, regRelative);  // 1.1.0
                                    fi.SetParameterByName(ADSK_Parameters.Instance.RegionNormalized.Name, regNormalized);  // 1.1.0
                                    fi.SetParameterByName(ADSK_Parameters.Instance.BaselineIndex.Name, featureline.Baseline.Index);  // 1.1.0 These following parameters enable an assignment in case the featurline is not the original used to create the object
                                    fi.SetParameterByName(ADSK_Parameters.Instance.Code.Name, featureline.Code);  // 1.1.0
                                    fi.SetParameterByName(ADSK_Parameters.Instance.Side.Name, featureline.Side.ToString());  // 1.1.0
                                    fi.SetParameterByName(ADSK_Parameters.Instance.Corridor.Name, featureline.Baseline.CorridorName);  // 1.1.0
                                    fi.SetParameterByName(ADSK_Parameters.Instance.Station.Name, Math.Round(station, 3));  // 1.1.0
                                } //Family Instances

                                cs.Dispose();
                                localPoint.Dispose();
                                point.Dispose();

                                #endregion
                            } // PointObjects

                            else
                            {
                                #region LINEAR_OBJECTS

                                if (station > featureline.End || station < featureline.Start)  // [20180923] 1.1.0 
                                {
                                    if (!normalized)
                                    {
                                        station = featureline.Start + (double)element.GetParameterValueByName(ADSK_Parameters.Instance.RegionRelative.Name);
                                    }
                                    else
                                    {
                                        station = featureline.Start + (double)element.GetParameterValueByName(ADSK_Parameters.Instance.RegionNormalized.Name) * (featureline.End - featureline.Start);
                                    }
                                }


                                if (endStation > featureline.End || endStation < featureline.Start)  // [20180923] 1.1.0 
                                {
                                    if (!normalized)
                                    {
                                        endStation = featureline.Start + (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndRegionRelative.Name);
                                    }
                                    else
                                    {
                                        endStation = featureline.Start + (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndRegionNormalized.Name) * (featureline.End - featureline.Start);
                                    }
                                }

                                var localPoint = featureline.PointByStationOffsetElevation(station, offset, elevation, false);
                                var point = localPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                                var end = featureline.PointByStationOffsetElevation(endStation, endOffset, endElevation, false).Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;

                                RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(doc);

                                if (element.InternalElement.Location is LocationPoint) // Vertical column
                                {
                                    var lp = element.InternalElement.Location as LocationPoint;

                                    lp.Move(point.ToXyz() - lp.Point);

                                    ElementId baseId = element.InternalElement.Parameters.Cast<Parameter>().First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)).AsElementId();

                                    Level bl = (Level)doc.GetElement(baseId);

                                    double bh = bl.Elevation;

                                    ElementId topId = element.InternalElement.Parameters.Cast<Parameter>().First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)).AsElementId();

                                    Level tl = (Level)doc.GetElement(topId);

                                    double th = tl.Elevation;

                                    element.InternalElement.Parameters.Cast<Parameter>()
                                        .First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)).Set(Utils.MToFeet(elevation) - bh);

                                    element.InternalElement.Parameters.Cast<Parameter>()
                                        .First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)).Set(Utils.MToFeet(endElevation) - th);
                                } // Vertical columns
                                else
                                {

                                    var locCurve = element.InternalElement.Location as LocationCurve;

                                    var newCurve = Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(point, end).ToRevitType();

                                    if (element.InternalElement is MEPCurve)
                                    {
                                        UpdateMEPCurve(element.InternalElement, newCurve);
                                    }
                                    else
                                    {
                                        locCurve.Curve = newCurve;
                                    }

                                    element.SetParameterByName(ADSK_Parameters.Instance.EndRegionRelative.Name, endStation - featureline.Start);  // 1.1.0
                                    element.SetParameterByName(ADSK_Parameters.Instance.EndRegionNormalized.Name, (endStation - featureline.Start) / (featureline.End - featureline.Start));  // 1.1.0
                                } // MEP Curves or Walls or Structural Beams, etc.

                                RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                                element.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(localPoint.X, 3));
                                element.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(localPoint.Y, 3));
                                element.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(localPoint.Z, 3));
                                element.SetParameterByName(ADSK_Parameters.Instance.RegionIndex.Name, featureline.BaselineRegionIndex);  // 1.1.0
                                element.SetParameterByName(ADSK_Parameters.Instance.RegionRelative.Name, station - featureline.Start);  // 1.1.0
                                element.SetParameterByName(ADSK_Parameters.Instance.RegionNormalized.Name, (station - featureline.Start) / (featureline.End - featureline.Start));  // 1.1.0
                                element.SetParameterByName(ADSK_Parameters.Instance.BaselineIndex.Name, featureline.Baseline.Index);  // 1.1.0 These following parameters enable an assignment in case the featurline is not the original used to create the object
                                element.SetParameterByName(ADSK_Parameters.Instance.Code.Name, featureline.Code);  // 1.1.0
                                element.SetParameterByName(ADSK_Parameters.Instance.Side.Name, featureline.Side.ToString());  // 1.1.0
                                element.SetParameterByName(ADSK_Parameters.Instance.Corridor.Name, featureline.Baseline.CorridorName);  // 1.1.0

                                localPoint.Dispose();
                                point.Dispose();
                                end.Dispose();

                                #endregion
                            } // LineObjects

                        }
                    }

                    //++j;
                }

                //++i;
            }

            RevitUtils.ExportXML(true);  // 1.1.0

            Utils.Log(string.Format("UtilsObjectsLocation.OptimizedUdpateObjectFromXML completed.", ""));

            return elements;
        }

        /// <summary>
        /// Updates hte location of a ShapePoint object
        /// </summary>
        /// <param name="civilDocument">The CivilDocument in Civil 3D</param>
        /// <param name="sp">The ShapePoint</param>
        /// <returns></returns>
        private static ShapePoint UpdateShapePoint(CivilDocument civilDocument, ShapePoint sp)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.UpdateShapePoint started...", ""));

            Corridor corridor = null;
            int baselineIndex = sp.BaselineIndex;
            int regionIndex = sp.RegionIndex;  // 1.1.0
            string code = sp.Code;
            Featureline.SideType side = sp.Side;
            double station = sp.Station;
            double offset = sp.Offset;
            double elevation = sp.Elevation;
            double regionRelative = sp.RegionRelative;  // 1.1.0
            double regionNormalized = sp.RegionRelative;  // 1.1.0

            ShapePoint sp1 = sp;

            try
            {
                corridor = civilDocument.GetCorridors().First(x => x.Name == sp.Corridor);
            }
            catch { }

            if (corridor != null)
            {
                // TODO Add Relative and/or Normalized check
                Featureline f = corridor.Baselines[baselineIndex].GetFeaturelinesByCodeStation(code, station).First(x => x.Side == side);  // 1.1.0

                if (f != null)
                {
                    sp1 = sp.Copy(sp.Id);
                    sp1 = sp1.UpdateByFeatureline(f);  // 1.1.0
                }
            }

            Utils.Log(string.Format("UtilsObjectsLocation.UpdateShapePoint completed.", ""));

            return sp1;
        }

        /// <summary>
        /// Not in use
        /// </summary>
        /// <param name="elements"></param>
        private static void GroupElements(IList<Revit.Elements.Element> elements)
        {
            var corridors = elements.GroupBy(e => (string)e.GetParameterValueByName(ADSK_Parameters.Instance.Corridor.Name)).Select(cg => cg.Key).ToList();
            var codes = elements.GroupBy(e => (string)e.GetParameterValueByName(ADSK_Parameters.Instance.Code.Name)).Select(cg => cg.Key).ToList();
            var sides = elements.GroupBy(e => (string)e.GetParameterValueByName(ADSK_Parameters.Instance.Side.Name)).Select(cg => cg.Key).ToList();
            var indices = elements.GroupBy(e => (int)e.GetParameterValueByName(ADSK_Parameters.Instance.BaselineIndex.Name)).Select(cg => cg.Key).ToList();
            var stations = elements.GroupBy(e => (double)e.GetParameterValueByName(ADSK_Parameters.Instance.Station.Name)).Select(cg => cg.Key).ToList();
            var offsets = elements.GroupBy(e => (double)e.GetParameterValueByName(ADSK_Parameters.Instance.Offset.Name)).Select(cg => cg.Key).ToList();
            var elevations = elements.GroupBy(e => (double)e.GetParameterValueByName(ADSK_Parameters.Instance.Elevation.Name)).Select(cg => cg.Key).ToList();
        }

        private static Dictionary<string, object> GetElementFeatureline(CivilDocument civilDocument, Revit.Elements.Element element)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.GetElementFeatureline started...", ""));

            Featureline featureline = null;
            double station = 0;
            double offset = 0;
            double elevation = 0;
            double angle = 0;

            double endStation = 0;
            double endOffset = 0;
            double endElevation = 0;

            if ((int)element.GetParameterValueByName(ADSK_Parameters.Instance.Update.Name) == 1)
            {
                var corridor = (string)element.GetParameterValueByName(ADSK_Parameters.Instance.Corridor.Name);
                var baselineIndex = (int)element.GetParameterValueByName(ADSK_Parameters.Instance.BaselineIndex.Name);
                var code = (string)element.GetParameterValueByName(ADSK_Parameters.Instance.Code.Name);
                var side = (string)element.GetParameterValueByName(ADSK_Parameters.Instance.Side.Name);
                station = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Station.Name);
                offset = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Offset.Name);
                elevation = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Elevation.Name);

                try
                {
                    angle = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.AngleZ.Name);
                }
                catch
                {

                }

                // TODO: Check for parameters different from null

                Corridor corr = null;

                foreach (var c in civilDocument.GetCorridors())
                {
                    if (c.Name == corridor)
                    {
                        corr = c;
                        break;
                    }
                }

                if (corr == null)
                {
                    throw new Exception("It was not possible to find corridor");
                }

                featureline = corr.GetFeaturelinesByCodeStation(code, station)[baselineIndex].First(x => x.Side.ToString() == side);

                if (featureline == null)
                {
                    List<double> startStations = new List<double>();

                    foreach (var region in corr.GetFeaturelinesByCode(code)[baselineIndex])
                    {
                        foreach (var f in region)
                        {
                            if (!startStations.Contains(f.Start))
                            {
                                startStations.Add(f.Start);
                            }
                        }
                    }

                    foreach (var region in corr.GetFeaturelinesByCode(code)[baselineIndex])
                    {
                        foreach (var f in region)
                        {
                            if (!startStations.Contains(f.End))
                            {
                                startStations.Add(f.End);
                            }
                        }
                    }

                    if (station < startStations.Min())
                    {
                        station = startStations.Min();
                    }
                    else if (station > startStations.Max())
                    {
                        station = startStations.Max();
                    }

                    int regionIndex = corr
                        .Baselines[baselineIndex]
                        .GetBaselineRegionIndexByStation(station);

                    foreach (var c in corr
                    .GetFeaturelinesByCode(code)[baselineIndex][regionIndex]
                    .Where(f => f.Side.ToString() == side && station >= f.Start && station <= f.End))
                    {
                        featureline = c;
                        break;
                    }

                    if (featureline == null)
                    {
                        throw new Exception("It was not possible to find featureline");
                    }
                }

                try
                {
                    endStation = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndStation.Name);
                    endOffset = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndOffset.Name);
                    endElevation = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndElevation.Name);

                    if (endStation < featureline.Start)
                    {
                        endStation = featureline.Start;
                    }
                    if (endStation > featureline.End)
                    {
                        endStation = featureline.End;
                    }
                }
                catch
                {
                    endStation = station;
                    endOffset = offset;
                    endElevation = elevation;
                }
            }

            Utils.Log(string.Format("UtilsObjectsLocation.GetElementFeatureline completed.", ""));

            return new Dictionary<string, object> { 
            {"featureline", featureline}, 
            {"station", station}, 
            {"offset", offset}, 
            {"elevation", elevation}, 
            {"angle", angle},
            {"endstation", endStation}, 
            {"endoffset", endOffset}, 
            {"endelevation", endElevation} 
            };
        }

        /// <summary>
        /// Tests the udpate object.
        /// </summary>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        /// <exception cref="Exception">
        /// It was not possible to find the specified corridor
        /// or
        /// It was not possible to find the specified featureline
        /// or
        /// It was not possible to find corridor
        /// or
        /// It was not possible to find featureline
        /// </exception>
        public static Revit.Elements.Element TestUdpateObject(CivilDocument civilDocument, Revit.Elements.Element element)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.TestUdpateObject started...", ""));

            var totalTransform = RevitUtils.DocumentTotalTransform();

            if (element is Revit.Elements.FamilyInstance)
            {
                var fi = element as Revit.Elements.FamilyInstance;
                bool found = false;
                if (fi.Location is Autodesk.DesignScript.Geometry.Point)
                {
                    
                    found = true;

                    if (found)
                    {
                        if ((int)element.GetParameterValueByName(ADSK_Parameters.Instance.Delete.Name) == 1)
                        {
                            RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

                            DocumentManager.Instance.CurrentDBDocument.Delete(element.InternalElement.Id);

                            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                            return null;
                        }
                        else if ((int)element.GetParameterValueByName(ADSK_Parameters.Instance.Update.Name) == 1)
                        {
                            var corridor = (string)element.GetParameterValueByName(ADSK_Parameters.Instance.Corridor.Name);
                            var baselineIndex = (int)element.GetParameterValueByName(ADSK_Parameters.Instance.BaselineIndex.Name);
                            var code = (string)element.GetParameterValueByName(ADSK_Parameters.Instance.Code.Name);
                            var side = (string)element.GetParameterValueByName(ADSK_Parameters.Instance.Side.Name);
                            var station = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Station.Name);
                            var offset = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Offset.Name);
                            var elevation = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Elevation.Name);
                            var angle = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.AngleZ.Name);

                            // TODO: Check for parameters different from null

                            Corridor corr = null;

                            foreach (var c in civilDocument.GetCorridors())
                            {
                                if (c.Name == corridor)
                                {
                                    corr = c;
                                    break;
                                }
                            }

                            if (corr == null)
                            {
                                var message = "It was not possible to find the specified corridor";

                                Utils.Log(string.Format("Error: UtilsObjectsLocation.TestUdpateObject {0}", message));

                                throw new Exception(message);
                            }

                            List<double> startStations = new List<double>();

                            Featureline featureline = null;

                            var featurelines = corr.GetFeaturelinesByCode(code) as List<IList<Featureline>>;

                            featureline = featurelines[baselineIndex].ToList().First(x => x.Side.ToString() == side);

                            if (featureline == null)
                            {
                                var message = "It was not possible to find the specified featureline";

                                Utils.Log(string.Format("Error: UtilsObjectsLocation.TestUdpateObject {0}", message));

                                throw new Exception(message);
                            }

                            CoordinateSystem cs = featureline.CoordinateSystemByStation(station);

                            var localPoint = featureline.PointByStationOffsetElevation(station, offset, elevation, false);
                            var point = localPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;

                            RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

                            var lp = fi.InternalElement.Location as LocationPoint;
                            lp.Move(point.ToXyz() - lp.Point);

                            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                            if (angle != 0)
                            {
                                fi.SetRotation(-angle);
                            }

                            var ang = fi.FacingOrientation.AngleAboutAxis(Vector.XAxis(), Vector.ZAxis());

                            if (ang != 0)
                            {
                                fi.SetRotation(-ang);
                            }

                            var xAxis = cs.XAxis.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Vector;

                            fi.SetRotation(xAxis.AngleAboutAxis(Vector.XAxis(), cs.ZAxis) + angle);
                            fi.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(localPoint.X, 3));
                            fi.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(localPoint.Y, 3));
                            fi.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(localPoint.Z, 3));

                            cs.Dispose();
                        }
                    }
                }
            }
            else
            {
                bool found = false;
                if (element.InternalElement.Location is LocationCurve)
                {

                    found = true;

                    if (found)
                    {
                        if ((int)element.GetParameterValueByName(ADSK_Parameters.Instance.Delete.Name) == 1)
                        {
                            RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

                            DocumentManager.Instance.CurrentDBDocument.Delete(element.InternalElement.Id);

                            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                            return null;
                        }
                        else if ((int)element.GetParameterValueByName(ADSK_Parameters.Instance.Update.Name) == 1)
                        {
                            var corridor = (string)element.GetParameterValueByName(ADSK_Parameters.Instance.Corridor.Name);
                            var baselineIndex = (int)element.GetParameterValueByName(ADSK_Parameters.Instance.BaselineIndex.Name);
                            var code = (string)element.GetParameterValueByName(ADSK_Parameters.Instance.Code.Name);
                            var side = (string)element.GetParameterValueByName(ADSK_Parameters.Instance.Side.Name);
                            var station = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Station.Name);
                            var offset = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Offset.Name);
                            var elevation = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.Elevation.Name);

                            var endStation = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndStation.Name);
                            var endOffset = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndOffset.Name);
                            var endElevation = (double)element.GetParameterValueByName(ADSK_Parameters.Instance.EndElevation.Name);

                            Corridor corr = null;

                            foreach (var c in civilDocument.GetCorridors())
                            {
                                if (c.Name == corridor)
                                {
                                    corr = c;
                                    break;
                                }
                            }

                            if (corr == null)
                            {
                                var message = "It was not possible to find corridor";

                                Utils.Log(string.Format("Error: UtilsObjectsLocation.TestUdpateObject {0}", message));

                                throw new Exception(message);
                            }

                            int regionIndex = corr
                                .Baselines[baselineIndex]
                                .GetBaselineRegionIndexByStation(station);

                            Featureline featureline = null;

                            var featurelines = corr.GetFeaturelinesByCode(code) as List<IList<IList<Featureline>>>;

                            featureline = featurelines[baselineIndex][regionIndex][0];

                            if (featureline == null)
                            {
                                var message = "It was not possible to find featureline";

                                Utils.Log(string.Format("Error: UtilsObjectsLocation.TestUdpateObject {0}", message));

                                throw new Exception(message);
                            }

                            if (station < featureline.Start)
                            {
                                station = featureline.Start;
                            }
                            if (station > featureline.End)
                            {
                                station = featureline.End;
                            }
                            if (endStation < featureline.Start)
                            {
                                endStation = featureline.Start;
                            }
                            if (endStation > featureline.End)
                            {
                                endStation = featureline.End;
                            }

                            var localPoint = featureline.PointByStationOffsetElevation(station, offset, elevation, false);
                            var point = localPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                            var end = featureline.PointByStationOffsetElevation(endStation, endOffset, endElevation, false).Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;

                            Document doc = DocumentManager.Instance.CurrentDBDocument;

                            RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(doc);

                            var locCurve = element.InternalElement.Location as LocationCurve;

                            var newCurve = Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(point, end).ToRevitType();

                            UpdateMEPCurve(element.InternalElement, newCurve);

                            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                            element.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(localPoint.X, 3));
                            element.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(localPoint.Y, 3));
                            element.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(localPoint.Z, 3));

                            localPoint.Dispose();
                            point.Dispose();
                            end.Dispose();
                        }
                    }
                }
            }

            Utils.Log(string.Format("UtilsObjectsLocation.TestUdpateObject completed.", ""));

            return element;
        }

        /// <summary>
        /// Return the given element's connector manager,
        /// using either the family instance MEPModel or
        /// directly from the MEPCurve connector manager
        /// for ducts and pipes.
        /// </summary>
        /// <param name="e">The e.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">Element is neither an MEP curve nor a fitting.</exception>
        private static ConnectorManager GetConnectorManager(Element e)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.GetConnectorManager started...", ""));

            MEPCurve mc = e as MEPCurve;
            FamilyInstance fi = e as FamilyInstance;

            if (null == mc && null == fi)
            {
                var message = "Element is neither an MEP curve nor a fitting.";

                Utils.Log(string.Format("ERROR: UtilsObjectsLocation.GetConnectorManager {0}", message));

                throw new ArgumentException(message);
            }

            Utils.Log(string.Format("UtilsObjectsLocation.GetConnectorManager completed.", ""));

            return null == mc
                ? fi.MEPModel.ConnectorManager
                : mc.ConnectorManager;
        }

        /// <summary>
        /// Updates the mep curve.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="newCurve">The new curve.</param>
        public static void UpdateMEPCurve(Element element, Autodesk.Revit.DB.Curve newCurve)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.UpdateMEPCurve started...", ""));

            MEPCurve p = null;

            if (element is MEPCurve)
            {
                p = element as MEPCurve;
            }
            else
            {
                return;
            }

            Element next = null;

            LocationCurve pLocCurve = p.Location as LocationCurve;

            Autodesk.Revit.DB.Curve pCurve = pLocCurve.Curve;

            pCurve = newCurve;

            ConnectorManager cm = GetConnectorManager(p);

            foreach (Connector connector in cm.Connectors)
            {
                foreach (Connector eCon in connector.AllRefs)
                {
                    if (eCon.Owner.Id.IntegerValue != p.Id.IntegerValue)
                    {
                        if (eCon.Owner is MEPCurve)
                        {
                            next = eCon.Owner;

                            LocationCurve nLocCurve = next.Location as LocationCurve;

                            Autodesk.Revit.DB.Curve nCurve = nLocCurve.Curve;

                            double minDist = double.MaxValue;

                            Autodesk.Revit.DB.Curve temp = null;

                            for (int i = 0; i < 2; ++i)
                            {
                                for (int j = 0; j < 2; ++j)
                                {
                                    double d = pCurve.GetEndPoint(i).DistanceTo(nCurve.GetEndPoint(j));

                                    if (d < minDist)
                                    {
                                        minDist = d;

                                        if (j == 0)
                                        {
                                            if (pCurve.GetEndPoint(i).DistanceTo(nCurve.GetEndPoint(1)) > p.Document.Application.ShortCurveTolerance)
                                            {
                                                temp = Autodesk.Revit.DB.Line.CreateBound(pCurve.GetEndPoint(i), nCurve.GetEndPoint(1));
                                            }
                                        }
                                        else
                                        {
                                            if (pCurve.GetEndPoint(i).DistanceTo(nCurve.GetEndPoint(0)) > p.Document.Application.ShortCurveTolerance)
                                            {
                                                temp = Autodesk.Revit.DB.Line.CreateBound(nCurve.GetEndPoint(0), pCurve.GetEndPoint(i));
                                            }
                                        }
                                    }
                                }
                            }

                            nLocCurve.Curve = temp;
                        }

                        if (eCon.Owner is FamilyInstance)
                        {
                            foreach (Connector qCon in GetConnectorManager(eCon.Owner).Connectors)
                            {
                                if (qCon.Owner.Id.IntegerValue != eCon.Owner.Id.IntegerValue &&
                                   qCon.Owner.Id.IntegerValue != p.Id.IntegerValue)
                                {
                                    if (qCon.Owner is MEPCurve)
                                    {
                                        next = qCon.Owner;

                                        LocationCurve nLocCurve = next.Location as LocationCurve;

                                        Autodesk.Revit.DB.Curve nCurve = nLocCurve.Curve;

                                        double minDist = double.MaxValue;

                                        Autodesk.Revit.DB.Curve temp = null;

                                        for (int i = 0; i < 2; ++i)
                                        {
                                            for (int j = 0; j < 2; ++j)
                                            {
                                                double d = pCurve.GetEndPoint(i).DistanceTo(nCurve.GetEndPoint(j));

                                                if (d < minDist)
                                                {
                                                    minDist = d;

                                                    if (j == 0)
                                                    {
                                                        if (pCurve.GetEndPoint(i).DistanceTo(nCurve.GetEndPoint(1)) > p.Document.Application.ShortCurveTolerance)
                                                        {
                                                            temp = Autodesk.Revit.DB.Line.CreateBound(pCurve.GetEndPoint(i), nCurve.GetEndPoint(1));
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (pCurve.GetEndPoint(i).DistanceTo(nCurve.GetEndPoint(0)) > p.Document.Application.ShortCurveTolerance)
                                                        {
                                                            temp = Autodesk.Revit.DB.Line.CreateBound(nCurve.GetEndPoint(0), pCurve.GetEndPoint(i));
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        nLocCurve.Curve = temp;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Utils.Log(string.Format("UtilsObjectsLocation.UpdateMEPCurve completed.", ""));

            pLocCurve.Curve = pCurve;
        }

        // TODO: test on a large Revit model how long it takes to read the data
        /// <summary>
        /// Reads the family instances point based.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="categoryId">The category identifier.</param>
        /// <returns></returns>
        public static object[][] ReadFamilyInstancesPointBased(Document doc, CivilDocument civilDocument, int categoryId)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.ReadFamilyInstancesPointBased started...", ""));

            // Looking only at single insertion point FamilyInstances
            var instances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(x => x.ViewSpecific == false &&
                    x.Location is LocationPoint &&
                    x.Category.Id.IntegerValue.Equals(categoryId));

            object[][] data = new object[instances.Count()][];

            int count = 0;

            //TODO: Check Shared Parameters first

            if (!SessionVariables.ParametersCreated)
            {
                CheckParameters(doc); 
            }

            Transform tr = doc.ActiveProjectLocation.GetTotalTransform().Inverse;

            foreach (FamilyInstance fi in instances)
            {

                // Shared Parameters                
                Parameter pCorridorName = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.Corridor.Name);
                Parameter pBaselineIndex = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.BaselineIndex.Name);
                Parameter pCode = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.Code.Name);
                Parameter pSide = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.Side.Name);
                Parameter pX = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.X.Name);
                Parameter pY = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.Y.Name);
                Parameter pZ = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.Z.Name);
                Parameter pStation = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.Station.Name);
                Parameter pOffset = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.Offset.Name);
                Parameter pElevation = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.Elevation.Name);
                Parameter pAngleZ = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.AngleZ.Name);
                Parameter pUpdate = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.Update.Name);
                Parameter pDelete = fi.Parameters.Cast<Parameter>().First(g => g.Definition.Name == ADSK_Parameters.Instance.Delete.Name);

                Location location = fi.Location;
                LocationPoint locPoint = location as LocationPoint;
                XYZ point = tr.OfPoint(locPoint.Point);
                Autodesk.DesignScript.Geometry.Point dsPoint = point.ToPoint();

                double x = 0;
                double y = 0;
                double z = 0;

                if (pX.HasValue)
                {
                    x = pX.AsDouble();
                }
                else
                {
                    x = FeetToM(point.X);
                }

                if (pY.HasValue)
                {
                    y = pY.AsDouble();
                }
                else
                {
                    y = FeetToM(point.Y);
                }

                if (pZ.HasValue)
                {
                    z = pZ.AsDouble();
                }
                else
                {
                    z = FeetToM(point.Z);
                }

                string uniqueId = fi.UniqueId;
                ElementId elementId = fi.Id;
                ElementId typeId = fi.Symbol.Id;
                string familyName = fi.Symbol.FamilyName;
                string typeName = fi.Symbol.Name;
                Parameter pMark = fi.Parameters.Cast<Parameter>().First(g => g.Id.IntegerValue.Equals(Convert.ToInt32(BuiltInParameter.ALL_MODEL_MARK)));

                string mark = "";
                string corridorName = "";
                int baselineIndex = 0;
                string code = "";
                string side = "";
                double station = 0;
                double offset = 0;
                double elevation = 0;
                double angleZ = 0;
                int update = 1;
                int delete = 0;

                // Check the mark
                if (!pMark.HasValue || pMark.AsString() == "")
                {
                    mark = "<None>"; // TODO: define numbering based on station values?
                }
                else
                {
                    mark = pMark.AsString();
                }

                //Check the corridor name

                if (pCorridorName.AsString() == "" || !pCorridorName.HasValue)
                {
                    corridorName = civilDocument.GetCorridors().First().Name;
                }
                else if (pCorridorName.HasValue && pCorridorName.AsString() != "")
                {
                    bool found = false;

                    foreach (Corridor c in civilDocument.GetCorridors())
                    {
                        if (c.Name == pCorridorName.AsString())
                        {
                            found = true;
                            corridorName = pCorridorName.AsString();
                            break;
                        }
                    }

                    if (!found)
                    {
                        corridorName = "The specified corridor name is not the current Civil Document";
                    }
                }

                Corridor corridor = civilDocument.GetCorridors().First();

                foreach (Corridor c in civilDocument.GetCorridors())
                {
                    if (c.Name == pCorridorName.AsString())
                    {
                        corridor = c;
                        break;
                    }
                }

                if (pBaselineIndex.HasValue)
                {
                    baselineIndex = pBaselineIndex.AsInteger();
                }

                if (!pCode.HasValue || pCode.AsString() == "")
                {
                    code = "*";
                }
                else
                {
                    bool found = false;
                    foreach (string c in corridor.GetCodes())
                    {
                        if (pCode.AsString() == c)
                        {
                            found = true;
                            code = pCode.AsString();
                            break;
                        }
                    }

                    if (!found)
                    {
                        code = pCode.AsString();
                    }
                }

                if (pStation.HasValue)
                {
                    station = pStation.AsDouble();
                }
                if (pOffset.HasValue)
                {
                    offset = pOffset.AsDouble();
                }
                if (pElevation.HasValue)
                {
                    elevation = pElevation.AsDouble();
                }

                int regIndex = corridor.Baselines[baselineIndex].GetBaselineRegionIndexByStation(station);


                if (!pStation.HasValue || !pOffset.HasValue || !pElevation.HasValue)
                {
                    if (code != "*" && code != "")
                    {
                        Featureline fl = corridor.GetFeaturelinesByCode(code)[baselineIndex][regIndex].First(f => f.Side.ToString() == pSide.AsString());

                        CoordinateSystem cs = fl.CoordinateSystemByStation(station);

                        angleZ = 0;
                        var soe = fl.GetStationOffsetElevationByPoint(dsPoint);
                        station = (double)soe["Station"];
                        offset = (double)soe["Offset"];
                        elevation = (double)soe["Elevation"];

                        cs.Dispose();
                    }
                    else if (code == "*")
                    {
                        Baseline baseline = corridor.Baselines[baselineIndex];
                        CoordinateSystem cs = baseline.CoordinateSystemByStation(station);

                        angleZ = 0;
                        station = baseline.GetArrayStationOffsetElevationByPoint(dsPoint)[0];
                        offset = baseline.GetArrayStationOffsetElevationByPoint(dsPoint)[1];
                        elevation = baseline.GetArrayStationOffsetElevationByPoint(dsPoint)[2];

                        cs.Dispose();
                    }
                }

                if (!pSide.HasValue || pSide.AsString() == "")
                {
                    if (offset <= 0)
                    {
                        side = "Left"; // by default 
                    }
                    else
                    {
                        side = "Right";
                    }
                }
                else
                {
                    side = pSide.AsString();
                }

                if (pUpdate.HasValue)
                {
                    update = pUpdate.AsInteger();
                }
                if (pDelete.HasValue)
                {
                    delete = pDelete.AsInteger();
                }

                data[count] = new object[] { uniqueId, elementId, typeId, familyName, typeName, mark, corridorName, baselineIndex, code, side, x, y, z, station, offset, elevation, angleZ, update, delete };

                ++count;
            }

            Utils.Log(string.Format("UtilsObjectsLocation.ReadFamilyInstancesPointBased completed.", ""));

            return data;
        }

        /// <summary>
        /// Object location parameters.
        /// </summary>
        /// <param name="familyInstance">The family instance.</param>
        /// <returns></returns>
        public static object[] ObjectLocationParameters(Revit.Elements.FamilyInstance familyInstance)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.ObjectLocationParameters started...", ""));

            var totalTransform = RevitUtils.DocumentTotalTransform();

            var totalTransformInverse = RevitUtils.DocumentTotalTransformInverse();

            if (!SessionVariables.ParametersCreated)
            {
                CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
            }
            Autodesk.DesignScript.Geometry.Point locationPBP = familyInstance.Location;
            Autodesk.DesignScript.Geometry.Point locationWCS = locationPBP.Transform(totalTransformInverse) as Autodesk.DesignScript.Geometry.Point;

            object[] data = new object[19];

            var uniqueId = familyInstance.UniqueId;
            var elementId = familyInstance.Id;
            var typeId = familyInstance.ElementType.Id;
            var familyName = familyInstance.GetFamily.Name;
            var typeName = familyInstance.ElementType.Name;
            var mark = familyInstance.GetParameterValueByName("Mark");

            var corridorName = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.Corridor.Name);

            var baselineIndex = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.BaselineIndex.Name);

            var code = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.Code.Name);

            var side = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.Side.Name);

            var x = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.X.Name);

            var y = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.Y.Name);

            var z = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.Z.Name);

            var station = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.Station.Name);

            var offset = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.Offset.Name);

            var elevation = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.Elevation.Name);

            var angleZ = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.AngleZ.Name);

            var update = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.Update.Name);

            var delete = familyInstance.GetParameterValueByName(ADSK_Parameters.Instance.Delete.Name);

            locationPBP.Dispose();
            locationWCS.Dispose();

            Utils.Log(string.Format("UtilsObjectsLocation.ObjectLocationParameters completed.", ""));


            return new object[] { uniqueId, elementId, typeId, familyName, typeName, mark, corridorName, baselineIndex, code, side, x, y, z, station, offset, elevation, 
                angleZ, update, delete };
        }

        /// <summary>
        /// Linear objects location parameters.
        /// </summary>
        /// <param name="linearMEPCurve">The linear mep curve.</param>
        /// <returns></returns>
        public static object[] LinearObjectLocationParameters(object linearMEPCurve)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.LinearObjectLocationParameters started...", ""));

            var totalTransform = RevitUtils.DocumentTotalTransform();

            var totalTransformInverse = RevitUtils.DocumentTotalTransformInverse();


            AbstractMEPCurve mep = linearMEPCurve as AbstractMEPCurve;
            if (!SessionVariables.ParametersCreated)
            {
                CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
            }
            var lc = mep.InternalMEPCurve.Location as LocationCurve;
            Autodesk.DesignScript.Geometry.Point startPBP = lc.Curve.ToProtoType().StartPoint;
            Autodesk.DesignScript.Geometry.Point endPBP = lc.Curve.ToProtoType().EndPoint;
            Autodesk.DesignScript.Geometry.Point startWCS = startPBP.Transform(totalTransformInverse) as Autodesk.DesignScript.Geometry.Point;
            Autodesk.DesignScript.Geometry.Point endWCS = endPBP.Transform(totalTransformInverse) as Autodesk.DesignScript.Geometry.Point;

            object[] data = new object[22];

            var uniqueId = mep.UniqueId;
            var elementId = mep.Id;
            var typeId = ElementId.InvalidElementId;
            if (mep.InternalMEPCurve is Pipe)
            {
                var obj = mep.InternalMEPCurve as Pipe;
                typeId = obj.PipeType.Id;
            }
            if (mep.InternalMEPCurve is Duct)
            {
                var obj = mep.InternalMEPCurve as Duct;
                typeId = obj.DuctType.Id;
            }
            if (mep.InternalMEPCurve is Conduit)
            {
                var obj = mep.InternalMEPCurve as Conduit;
                typeId = obj.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsElementId();
            }
            if (mep.InternalMEPCurve is CableTray)
            {
                var obj = mep.InternalMEPCurve as CableTray;
                typeId = obj.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsElementId();
            }

            var systemName = "";
            var domain = mep.Connectors()[0].Domain;
            if (domain == Domain.DomainCableTrayConduit)
            {
                systemName = "CableTrayConduit";
            }
            if (domain == Domain.DomainElectrical)
            {
                systemName = new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
                    .OfClass(typeof(Autodesk.Revit.DB.Electrical.ElectricalSystemType))
                    .WhereElementIsElementType()
                    .Cast<Autodesk.Revit.DB.Electrical.ElectricalSystemType>()
                    .First(e => e.Equals(mep.Connectors()[0].InternalConnector.ElectricalSystemType))
                    .ToString();
            }
            if (domain == Domain.DomainHvac)
            {
                systemName = new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
                    .OfClass(typeof(Autodesk.Revit.DB.Mechanical.MechanicalSystemType))
                    .WhereElementIsElementType()
                    .Cast<Autodesk.Revit.DB.Electrical.ElectricalSystemType>()
                    .First(e => e.Equals(mep.Connectors()[0].InternalConnector.DuctSystemType))
                    .ToString();
            }
            if (domain == Domain.DomainPiping)
            {
                systemName = new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
                    .OfClass(typeof(Autodesk.Revit.DB.Plumbing.PipingSystemType))
                    .WhereElementIsElementType()
                    .Cast<Autodesk.Revit.DB.Plumbing.PipingSystemType>()
                    .First(e => e.Equals(mep.Connectors()[0].InternalConnector.PipeSystemType))
                    .ToString();
            }

            var typeName = Revit.Elements.ElementSelector.ByElementId(typeId.IntegerValue).Name;
            var mark = mep.GetParameterValueByName("Mark");

            var corridorName = mep.GetParameterValueByName(ADSK_Parameters.Instance.Corridor.Name);

            var baselineIndex = mep.GetParameterValueByName(ADSK_Parameters.Instance.BaselineIndex.Name);

            var code = mep.GetParameterValueByName(ADSK_Parameters.Instance.Code.Name);

            var side = mep.GetParameterValueByName(ADSK_Parameters.Instance.Side.Name);

            var sx = mep.GetParameterValueByName(ADSK_Parameters.Instance.X.Name);

            var sy = mep.GetParameterValueByName(ADSK_Parameters.Instance.Y.Name);

            var sz = mep.GetParameterValueByName(ADSK_Parameters.Instance.Z.Name);


            var ex = endWCS.X;
            var ey = endWCS.Y;
            var ez = endWCS.Z;


            var startStation = mep.GetParameterValueByName(ADSK_Parameters.Instance.Station.Name);

            var endStation = mep.GetParameterValueByName(ADSK_Parameters.Instance.EndStation.Name);

            var startOffset = mep.GetParameterValueByName(ADSK_Parameters.Instance.Offset.Name);

            var endOffset = mep.GetParameterValueByName(ADSK_Parameters.Instance.EndOffset.Name);

            var startElevation = mep.GetParameterValueByName(ADSK_Parameters.Instance.Elevation.Name);

            var endElevation = mep.GetParameterValueByName(ADSK_Parameters.Instance.EndElevation.Name);

            var update = mep.GetParameterValueByName(ADSK_Parameters.Instance.Update.Name);

            var delete = mep.GetParameterValueByName(ADSK_Parameters.Instance.Delete.Name);

            startPBP.Dispose();
            startWCS.Dispose();

            endPBP.Dispose();
            endWCS.Dispose();

            Utils.Log(string.Format("UtilsObjectsLocation.LinearObjectLocationParameters completed.", ""));

            return new object[] { uniqueId, elementId, typeId, systemName, typeName, mark, corridorName, baselineIndex, code, 
                side, sx, sy, sz, startStation, startOffset, startElevation, ex, ey, ez, endStation, endOffset, endElevation, update, delete };
        }

        /// <summary>
        /// Gets the station offset elevation.
        /// </summary>
        /// <param name="corridor">The corridor.</param>
        /// <param name="baselineIndex">Index of the baseline.</param>
        /// <param name="point">The point.</param>
        /// <returns></returns>
        private static double[] GetStationOffsetElevation(Corridor corridor, int baselineIndex, XYZ point)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.GetStationOffsetElevation started...", ""));

            Baseline baseline = corridor.Baselines[0];

            if (baselineIndex >= 0 && baselineIndex <= corridor.Baselines.Count - 1)
            {
                baseline = corridor.Baselines[baselineIndex];
            }

            Utils.Log(string.Format("UtilsObjectsLocation.GetStationOffsetElevation completed.", ""));

            return baseline.GetArrayStationOffsetElevationByPoint(point.ToPoint());
        }

        /// <summary>
        /// Parameter to collect Name and GUID
        /// </summary>
        public class ADSK_Parameter
        {
            string _name;
            string _guid;

            /// <summary>
            /// The name of the shared parameter
            /// </summary>
            public string Name { get { return "ADSK_" + this._name; } set { this._name = value; } }

            /// <summary>
            /// The GUID of the shared parameter
            /// </summary>
            public string GUID { get { return this._guid; } set { this._guid = value; } }

            /// <summary>
            /// Constrictor
            /// </summary>
            /// <param name="name">The name of the shared parameter.</param>
            /// <param name="guid">The guid of the shared parameter.</param>
            public ADSK_Parameter(string name, string guid)
            {
                this._name = name;
                this._guid = guid;
            }
        }

        /// <summary>
        /// Class that given a parameter name returns the same guid.
        /// This is to have the same CivilConnection parameters in Revit files and tansfer project standards.
        /// </summary>
        public class ADSK_Parameters
        {
            ///<excluded/>
            public ADSK_Parameter Corridor { get { return new ADSK_Parameter("Corridor", "ED990980-BF6B-4B0D-8AA4-7AF00B5E518E"); } }
            ///<excluded/>
            public ADSK_Parameter BaselineIndex { get { return new ADSK_Parameter("BaselineIndex", "CA4C73E1-774E-46DE-A2FA-7CB38B0EBAE4"); } }
            ///<excluded/>
            public ADSK_Parameter RegionIndex { get { return new ADSK_Parameter("RegionIndex", "BB94E73A-9ABE-4196-8D9E-FB28BF3E2A00"); } }
            ///<excluded/>
            public ADSK_Parameter RegionRelative { get { return new ADSK_Parameter("RegionRelative", "2302AE08-18E0-4ECE-90A4-DED9C7B8FB13"); } }
            ///<excluded/>
            public ADSK_Parameter RegionNormalized { get { return new ADSK_Parameter("RegionNormalized", "736E8182-152E-4590-A35E-E8756C771F6C"); } }
            ///<excluded/>
            public ADSK_Parameter Code { get { return new ADSK_Parameter("Code", "EC3E64DF-6359-4FB7-A369-D6E39D0FA7AC"); } }
            ///<excluded/>
            public ADSK_Parameter Side { get { return new ADSK_Parameter("Side", "07546E01-5E3F-4EE2-BABC-FC016DD3BDDE"); } }
            ///<excluded/>
            public ADSK_Parameter X { get { return new ADSK_Parameter("X", "FF0594B6-064F-472D-AD93-1A3D0378B27E"); } }
            ///<excluded/>
            public ADSK_Parameter Y { get { return new ADSK_Parameter("Y", "D5005157-70F1-47A6-BAC1-635AB0D187AC"); } }
            ///<excluded/>
            public ADSK_Parameter Z { get { return new ADSK_Parameter("Z", "0FF9D026-D49A-42BF-9D2F-70070E92EC4E"); } }
            ///<excluded/>
            public ADSK_Parameter Station { get { return new ADSK_Parameter("Station", "B435EAF4-10CF-4AF5-95C7-5193406CC22E"); } }
            ///<excluded/>
            public ADSK_Parameter Offset { get { return new ADSK_Parameter("Offset", "90C3A519-46C0-411F-AF3C-7BCFFB64ABA4"); } }
            ///<excluded/>
            public ADSK_Parameter Elevation { get { return new ADSK_Parameter("Elevation", "E1790F04-ED60-4961-9E1D-5BBA1EDC9D7C"); } }
            ///<excluded/>
            public ADSK_Parameter EndStation { get { return new ADSK_Parameter("EndStation", "C4C35E24-B674-4F34-9236-B65A5304A1B2"); } }
            ///<excluded/>
            public ADSK_Parameter EndOffset { get { return new ADSK_Parameter("EndOffset", "CDAD01E4-0D0C-4939-9430-54D510ED6D0A"); } }
            ///<excluded/>
            public ADSK_Parameter EndElevation { get { return new ADSK_Parameter("EndElevation", "2FF228B5-6348-4C18-925C-CC4F071F289A"); } }
            ///<excluded/>
            public ADSK_Parameter AngleZ { get { return new ADSK_Parameter("AngleZ", "0DBFD02F-7F1A-4FB2-87F6-1FAACED72868"); } }
            ///<excluded/>
            public ADSK_Parameter Update { get { return new ADSK_Parameter("Update", "A8875117-6353-465C-B6A8-49E22E558C9E"); } }
            ///<excluded/>
            public ADSK_Parameter Delete { get { return new ADSK_Parameter("Delete", "492626E0-6B69-4164-BD2D-4E8F3D909C29"); } }
            ///<excluded/>
            public ADSK_Parameter MultiPoint { get { return new ADSK_Parameter("MultiPoint", "589DC507-77FD-4020-A954-31B9632094F7"); } }
            ///<excluded/>
            public ADSK_Parameter EndRegionRelative { get { return new ADSK_Parameter("EndRegionRelative", "A0FE8E88-21C2-4288-9B23-90757CF35A58"); } }
            ///<excluded/>
            public ADSK_Parameter EndRegionNormalized { get { return new ADSK_Parameter("EndRegionNormalized", "414EC335-04A0-4A98-B2F6-E0028445748E"); } }

            private static ADSK_Parameters _instance;

            /// <summary>
            /// The instance to create the Shared Parameters
            /// </summary>
            public static ADSK_Parameters Instance
            {
                get
                {
                    if (_instance == null)
                    {
                        _instance = new ADSK_Parameters();
                    }

                    return _instance;
                }
            }

            ///<excluded/>
            protected ADSK_Parameters()
            {
            }
        }

        /// <summary>
        /// Returns a raw of parameters to be used for creation.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="parameter">The parameter.</param>
        /// <param name="type">The type.</param>
        /// <param name="visible">if set to <c>true</c> [visible].</param>
        /// <param name="userModifiable">if set to <c>true</c> [user modifiable].</param>
        /// <param name="cats">The cats.</param>
        /// <param name="group">The group.</param>
        /// <param name="inst">if set to <c>true</c> [inst].</param>
        /// <returns></returns>
        public static ExternalDefinition RawCreateProjectParameter(Document doc, ADSK_Parameter parameter, ParameterType type, bool visible, bool userModifiable, CategorySet cats, BuiltInParameterGroup group, bool inst)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.RawCreateProjectParameter started...", ""));

            ExternalDefinition def = null;

            CategorySet cs2 = new CategorySet();

            string oriFile = doc.Application.SharedParametersFilename;

            string tempFile = Path.Combine(Path.GetTempPath(), "SharedParameters_CivilConnection.txt");

            if (!File.Exists(tempFile))
            {
                using (File.Create(tempFile)) { }
            }

            doc.Application.SharedParametersFilename = tempFile;

            string name = parameter.Name;

            ExternalDefinitionCreationOptions edco = new ExternalDefinitionCreationOptions(name, type);

            edco.GUID = new Guid(parameter.GUID);

            edco.UserModifiable = userModifiable;

            edco.Visible = visible;

            def = doc.Application.OpenSharedParameterFile().Groups.Create("TemporaryDefintionGroup").Definitions.Create(edco) as ExternalDefinition;

            BindingMap bm = doc.ParameterBindings;

            foreach (Category cat in cats)
            {
                // Loop all Binding Definitions
                // IMPORTANT NOTE: Categories.Size is ALWAYS 1 !?
                // For multiple categories, there is really one
                // pair per each category, even though the
                // Definitions are the same...

                DefinitionBindingMapIterator iter
                    = doc.ParameterBindings.ForwardIterator();

                bool found = true;

                while (iter.MoveNext())
                {
                    if (iter.Key.Name == name)
                    {
                        ElementBinding eb = (ElementBinding)iter.Current;

                        foreach (Category catEB in eb.Categories)
                        {
                            if (catEB.Id.IntegerValue.Equals(cat.Id.IntegerValue))
                            {
                                if (type == iter.Key.ParameterType)
                                {
                                    if (group == iter.Key.ParameterGroup)
                                    {
                                        found = false;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                if (found)
                    cs2.Insert(cat);
            }

            if (cs2.Size > 0)
            {
                doc.Application.SharedParametersFilename = oriFile;

                Autodesk.Revit.DB.Binding binding = doc.Application.Create.NewTypeBinding(cs2);

                if (inst) binding = doc.Application.Create.NewInstanceBinding(cs2);

                BindingMap map = (new UIApplication(doc.Application)).ActiveUIDocument.Document.ParameterBindings;

                map.Insert(def, binding, group);
            }

            File.Delete(tempFile);

            Utils.Log(string.Format("UtilsObjectsLocation.RawCreateProjectParameter completed.", ""));

            return def;
        }

        /// <summary>
        /// Checks the parameters.
        /// </summary>
        /// <param name="doc">The document.</param>
        public static void CheckParameters(Document doc)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.CheckParameters started...", ""));

            // TODO check if the docuemnt already has the parameters

            if (!SessionVariables.ParametersCreated)
            {
                CategorySet cs = new CategorySet();

                foreach (Category cat in doc.Settings.Categories.Cast<Category>().Where(c => c.AllowsBoundParameters && c.CategoryType == CategoryType.Model))
                {
                    cs.Insert(cat);
                }

                RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(doc);

                ADSK_Parameters par = ADSK_Parameters.Instance;
                RawCreateProjectParameter(doc, par.Corridor, ParameterType.Text, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.BaselineIndex, ParameterType.Integer, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.RegionIndex, ParameterType.Integer, true, false, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.RegionRelative, ParameterType.Length, true, false, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.RegionNormalized, ParameterType.Number, true, false, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.Code, ParameterType.Text, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.Side, ParameterType.Text, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.X, ParameterType.Length, true, false, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.Y, ParameterType.Length, true, false, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.Z, ParameterType.Length, true, false, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.Station, ParameterType.Length, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.Offset, ParameterType.Length, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.Elevation, ParameterType.Length, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.AngleZ, ParameterType.Angle, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.Update, ParameterType.YesNo, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.Delete, ParameterType.YesNo, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.MultiPoint, ParameterType.Text, true, true, cs, BuiltInParameterGroup.PG_DATA, true);


                cs.Clear();

                foreach (Category cat in doc.Settings.Categories.Cast<Category>().Where(c => c.AllowsBoundParameters && c.CategoryType == CategoryType.Model))
                {

                    if (//cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_PipeSegments) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_PipeCurves) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_FlexPipeCurves) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_DuctCurves) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_FlexDuctCurves) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_PlaceHolderDucts) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_PlaceHolderPipes) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_CableTray) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Conduit) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_StructuralFraming) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_StructuralColumns) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Columns) ||
                        cat.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Walls))
                    {
                        cs.Insert(cat);
                    }
                }

                RawCreateProjectParameter(doc, par.EndStation, ParameterType.Length, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.EndOffset, ParameterType.Length, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.EndElevation, ParameterType.Length, true, true, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.EndRegionRelative, ParameterType.Length, true, false, cs, BuiltInParameterGroup.PG_DATA, true);
                RawCreateProjectParameter(doc, par.EndRegionNormalized, ParameterType.Number, true, false, cs, BuiltInParameterGroup.PG_DATA, true);

                RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

                doc.Regenerate();

                SessionVariables.ParametersCreated = true;
            }

            Utils.Log(string.Format("UtilsObjectsLocation.CheckParameters completed.", ""));
        }

        /// <summary>
        /// Families the instances point based.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "Created", "Updated", "Deleted" })]
        public static Dictionary<string, object> FamilyInstancesPointBased(Document doc, CivilDocument civilDocument, object[][] data)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.FamilyInstancesPointBased started...", ""));

            IList<Revit.Elements.Element> updated = new List<Revit.Elements.Element>();
            IList<object> created = new List<object>();

            // Delete
            IList<ElementId> deleteIds = new List<ElementId>();

            // Update
            IList<int> updateIds = new List<int>();

            IList<int> createIds = new List<int>();

            Transform tr = doc.ActiveProjectLocation.GetTotalTransform();

            Corridor corridor = civilDocument.GetCorridors().First();
            string corridorName = "";

            Baseline baseline = corridor.Baselines.First();
            int baselineIndex = baseline.Index;

            Featureline fl = null;

            CoordinateSystem cs = CoordinateSystem.Identity();

            for (int i = 0; i < data.Count(); ++i)
            {
                int j = data[i].Count();

                int valDelete = Convert.ToInt32(data[i][j - 1]);
                int valUpdate = Convert.ToInt32(data[i][j - 2]);
                int valCreate = Convert.ToInt32(data[i][j - 3]);

                if (valDelete == 1)
                {
                    try
                    {
                        deleteIds.Add(doc.GetElement(Convert.ToString(data[i][0])).Id);
                    }
                    catch { }
                }
                else if (valUpdate == 1)
                {
                    updateIds.Add(i);
                }
                else if (valCreate == 1)
                {
                    createIds.Add(i);
                }
            }

            #region DELETE
            if (deleteIds.Count > 0)
            {
                RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(doc);
                doc.Delete(deleteIds);
                RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();
            }
            #endregion

            #region UPDATE
            if (updateIds.Count > 0)
            {
                RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(doc);
                // try to find existing object with matching Ids
                foreach (int i in updateIds)
                {
                    if (null != data[i][0])
                    {
                        Revit.Elements.FamilyInstance familyInstance = Revit.Elements.ElementSelector.ByUniqueId(Convert.ToString(data[i][0])) as Revit.Elements.FamilyInstance;

                        FamilyInstance fi = familyInstance.InternalElement as FamilyInstance;

                        Location location = fi.Location;

                        // Location
                        if (location is LocationPoint)
                        {
                            string code = "";
                            string side = "";
                            bool useFeatureLine = false;

                            string familyName = "";
                            string typeName = "";
                            string mark = "";

                            Family family = fi.Symbol.Family;
                            FamilySymbol type = fi.Symbol;
                            Parameter pMark = fi.Parameters.Cast<Parameter>().First(p => p.Id.IntegerValue.Equals(Convert.ToInt32(BuiltInParameter.ALL_MODEL_MARK)));

                            if (null != data[i][3])
                            {
                                familyName = Convert.ToString(data[i][3]);
                            }

                            if (null != data[i][4])
                            {
                                typeName = Convert.ToString(data[i][4]);
                            }

                            if (null != data[i][5])
                            {
                                mark = Convert.ToString(data[i][5]);
                            }

                            if (fi.Symbol.FamilyName != familyName && familyName != "")
                            {
                                var temp = new FilteredElementCollector(doc)
                                    .OfCategoryId(fi.Category.Id)
                                    .Cast<Family>()
                                    .First(w => w.Name == familyName);
                                if (temp != null)
                                {
                                    family = temp;
                                }
                            }

                            if (type.Name != typeName && typeName != "")
                            {
                                var temp = new FilteredElementCollector(doc)
                                    .OfCategoryId(fi.Category.Id)
                                    .WhereElementIsElementType()
                                    .Cast<FamilySymbol>()
                                    .Where(w => w.Family == family)
                                    .First(q => q.Name == typeName);
                                if (temp != null)
                                {
                                    fi.Symbol = temp;
                                }
                            }

                            Revit.Elements.FamilyType familyType = Revit.Elements.ElementSelector.ByUniqueId(type.UniqueId) as Revit.Elements.FamilyType;


                            if (pMark.AsString() != mark && mark != "")
                            {
                                pMark.Set(mark);
                            }

                            if (null != data[i][6])
                            {
                                corridorName = Convert.ToString(data[i][6]);
                            }

                            // Corridor
                            foreach (Corridor corr in civilDocument.GetCorridors())
                            {
                                if (corr.Name == corridorName)
                                {
                                    corridor = corr;
                                    break;
                                }
                            }

                            // Baseline index
                            if (null != data[i][7])
                            {
                                try
                                {
                                    baselineIndex = Convert.ToInt32(data[i][7]);
                                }
                                catch
                                {
                                    //System.Windows.Forms.MessageBox.Show("The specified baseline index is not valid");
                                }
                            }

                            // Baseline
                            if (baselineIndex >= 0 && baselineIndex < corridor.Baselines.Count)
                            {
                                baseline = corridor.Baselines[baselineIndex];
                            }

                            // Code
                            if (null != data[i][8])
                            {
                                foreach (string c in corridor.GetCodes())
                                {
                                    if (c == Convert.ToString(data[i][8]))
                                    {
                                        code = c;
                                        break;
                                    }
                                }
                            }

                            // Side
                            if (null != data[i][9])
                            {
                                string val = Convert.ToString(data[i][9]);
                                if (val.ToLower() == "left" || val.ToLower() == "l")
                                {
                                    side = "Left";
                                }

                                if (val.ToLower() == "right" || val.ToLower() == "r")
                                {
                                    side = "Right";
                                }
                            }




                            double station = 0;
                            double offset = 0;
                            double elevation = 0;

                            if (null != data[i][13])
                            {
                                station = Convert.ToDouble(data[i][13], System.Globalization.CultureInfo.InvariantCulture);
                            }

                            if (station < baseline.Start || station > baseline.End)
                            {
                                continue;
                            }

                            int regIndex = corridor.Baselines[baselineIndex].GetBaselineRegionIndexByStation(station);

                            // Featureline
                            if (side != "" && code != "")
                            {
                                Featureline.SideType sideType = Featureline.SideType.None;

                                if (side == "Left")
                                {
                                    sideType = Featureline.SideType.Left;
                                }
                                else
                                {
                                    sideType = Featureline.SideType.Right;
                                }

                                fl = corridor.GetFeaturelinesByCode(code)[baselineIndex][regIndex].First(f => f.Side == sideType);
                                useFeatureLine = true;
                            }

                            // TODO: Add check for station and featureline based on the baselineregion

                            if (null != data[i][14])
                            {
                                offset = Convert.ToDouble(data[i][14], System.Globalization.CultureInfo.InvariantCulture);
                            }

                            if (null != data[i][15])
                            {
                                elevation = Convert.ToDouble(data[i][15], System.Globalization.CultureInfo.InvariantCulture);
                            }

                            double angleZ = 0;

                            if (null != data[i][18])
                            {
                                angleZ = Convert.ToDouble(data[i][18], System.Globalization.CultureInfo.InvariantCulture);
                            }

                            UpdateFamilyInstance(familyInstance, familyType, fl, !useFeatureLine, station, offset, elevation, angleZ);

                            updated.Add(familyInstance);
                        }
                    }
                }
            }
            #endregion

            #region CREATE
            if (createIds.Count > 0)
            {
                foreach (int i in createIds)
                {
                    string familyName = "";
                    string typeName = "";
                    string mark = "";

                    double station = 0;
                    double offset = 0;
                    double elevation = 0;

                    double angleZ = 0;

                    string code = "";
                    string side = "";
                    bool useFeatureLine = false;

                    FamilySymbol type = null;

                    if (null != data[i][3])
                    {
                        familyName = Convert.ToString(data[i][3]);
                    }

                    if (null != data[i][4])
                    {
                        typeName = Convert.ToString(data[i][4]);
                    }

                    if (null != data[i][5])
                    {
                        mark = Convert.ToString(data[i][5]);
                    }

                    if (typeName != "")
                    {
                        var temp = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilySymbol))
                            .WhereElementIsElementType()
                            .Cast<FamilySymbol>()
                            .First(q => q.FamilyName == familyName && q.Name == typeName);
                        if (temp != null)
                        {
                            type = temp;
                        }
                    }

                    // TODO: add report here
                    if (type == null)
                    {
                        continue;
                    }

                    // Corridor
                    foreach (Corridor corr in civilDocument.GetCorridors())
                    {
                        if (corr.Name == corridorName)
                        {
                            corridor = corr;
                            break;
                        }
                    }

                    // Baseline index
                    if (null != data[i][7])
                    {
                        try
                        {
                            baselineIndex = Convert.ToInt32(data[i][7]);
                        }
                        catch
                        {
                            //System.Windows.Forms.MessageBox.Show("The specified baseline index is not valid");
                        }
                    }

                    // Baseline
                    if (baselineIndex >= 0 && baselineIndex < corridor.Baselines.Count)
                    {
                        baseline = corridor.Baselines[baselineIndex];
                    }

                    // Code
                    if (null != data[i][8])
                    {
                        foreach (string c in corridor.GetCodes())
                        {
                            if (c == Convert.ToString(data[i][8]))
                            {
                                code = c;
                                break;
                            }
                        }
                    }

                    // Side
                    if (null != data[i][9])
                    {
                        string val = Convert.ToString(data[i][9]);
                        if (val.ToLower() == "left" || val.ToLower() == "l")
                        {
                            side = "Left";
                        }

                        if (val.ToLower() == "right" || val.ToLower() == "r")
                        {
                            side = "Right";
                        }
                    }



                    XYZ newLocPoint = new XYZ();

                    if (null != data[i][13])
                    {
                        station = Convert.ToDouble(data[i][13], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (station < baseline.Start || station > baseline.End)
                    {
                        continue;
                    }

                    int regIndex = corridor.Baselines[baselineIndex].GetBaselineRegionIndexByStation(station);

                    // Featureline
                    if (side != "" && code != "")
                    {
                        Featureline.SideType sideType = Featureline.SideType.None;

                        if (side == "Left")
                        {
                            sideType = Featureline.SideType.Left;
                        }
                        else
                        {
                            sideType = Featureline.SideType.Right;
                        }

                        fl = corridor.GetFeaturelinesByCode(code)[baselineIndex][regIndex].First(f => f.Side == sideType);
                        useFeatureLine = true;
                    }

                    // TODO: Add check for station and featureline based on the baselineregion

                    if (null != data[i][14])
                    {
                        offset = Convert.ToDouble(data[i][14], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (null != data[i][15])
                    {
                        elevation = Convert.ToDouble(data[i][15], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (null != data[i][5])
                    {
                        mark = Convert.ToString(data[i][5]);
                    }

                    Revit.Elements.FamilyType familyType = Revit.Elements.ElementSelector.ByUniqueId(type.UniqueId) as Revit.Elements.FamilyType;

                    created.Add(CreateFamilyInstance(familyType, fl, !useFeatureLine, station, offset, elevation, angleZ));
                }
            }
            #endregion
            // Read 

            //TODO: new data output to overwrite the original

            cs.Dispose();

            Utils.Log(string.Format("UtilsObjectsLocation.FamilyInstancesPointBased completed.", ""));

            return new Dictionary<string, object> { { "Created", created }, { "Updated", updated }, { "Deleted", deleteIds } };
        }

        /// <summary>
        /// Reads the family instance location parameters for update.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "FamilyInstance", "FamilyType", "Mark", "Featureline", "UseBaseline", "Station", "Offset", "Elevation", "AngleZ" })]
        public static Dictionary<string, object> ReadFamilyInstanceLocationParametersForUpdate(Document doc, CivilDocument civilDocument, object[][] data)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.ReadFamilyInstanceLocationParametersForUpdate started...", ""));

            IList<int> updateIds = new List<int>();

            Transform tr = doc.ActiveProjectLocation.GetTotalTransform();

            Corridor corridor = civilDocument.GetCorridors().First();
            string corridorName = "";

            Baseline baseline = corridor.Baselines.First();
            int baselineIndex = baseline.Index;

            Featureline fl = null;

            // TODO: group entries by Corridor
            // TODO: group entries by BaselineIndex
            // TODO: group entries by Code
            // TODO: group entries by Side

            IList<Revit.Elements.FamilyInstance> familyInstances = new List<Revit.Elements.FamilyInstance>();
            IList<Revit.Elements.FamilyType> familyTypes = new List<Revit.Elements.FamilyType>();
            IList<string> marks = new List<string>();
            IList<Featureline> featurelines = new List<Featureline>();
            IList<bool> useFeaturelines = new List<bool>();
            IList<double> stations = new List<double>();
            IList<double> offsets = new List<double>();
            IList<double> elevations = new List<double>();
            IList<double> zAngles = new List<double>();

            for (int i = 0; i < data.Count(); ++i)
            {
                int j = data[i].Count();

                int valUpdate = Convert.ToInt32(data[i][j - 2]);
                int valDelete = Convert.ToInt32(data[i][j - 1]);

                if (valUpdate == 1 && valDelete != 1)
                {
                    updateIds.Add(i);
                }
            }

            if (updateIds.Count > 0)
            {
                foreach (int i in updateIds)
                {
                    string familyName = "";
                    string typeName = "";
                    string mark = "";

                    double station = 0;
                    double offset = 0;
                    double elevation = 0;

                    double angleZ = 0;
                    string code = "";
                    string side = "";
                    bool useFeatureLine = false;

                    Revit.Elements.FamilyInstance fi = Revit.Elements.ElementSelector.ByUniqueId(Convert.ToString(data[i][0])) as Revit.Elements.FamilyInstance;

                    FamilySymbol type = fi.ElementType.InternalElement as FamilySymbol;

                    if (null != data[i][3])
                    {
                        familyName = Convert.ToString(data[i][3]);
                    }

                    if (null != data[i][4])
                    {
                        typeName = Convert.ToString(data[i][4]);
                    }

                    if (null != data[i][5])
                    {
                        mark = Convert.ToString(data[i][5]);
                    }
                    if (mark == "")
                    {
                        mark = i.ToString();
                    }

                    if (typeName != "")
                    {
                        var temp = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilySymbol))
                            .WhereElementIsElementType()
                            .Cast<FamilySymbol>()
                            .First(q => q.FamilyName == familyName && q.Name == typeName);
                        if (temp != null)
                        {
                            type = temp;
                        }
                    }

                    // Corridor
                    foreach (Corridor corr in civilDocument.GetCorridors())
                    {
                        if (corr.Name == corridorName)
                        {
                            corridor = corr;
                            break;
                        }
                    }

                    // Baseline index
                    if (null != data[i][7])
                    {
                        try
                        {
                            baselineIndex = Convert.ToInt32(data[i][7]);
                        }
                        catch
                        {
                            //System.Windows.Forms.MessageBox.Show("The specified baseline index is not valid");
                        }
                    }

                    // Baseline
                    if (baselineIndex >= 0 && baselineIndex < corridor.Baselines.Count)
                    {
                        baseline = corridor.Baselines[baselineIndex];
                    }

                    // Code
                    if (null != data[i][8])
                    {
                        foreach (string c in corridor.GetCodes())
                        {
                            if (c == Convert.ToString(data[i][8]))
                            {
                                code = c;
                                break;
                            }
                        }
                    }

                    // Side
                    if (null != data[i][9])
                    {
                        string val = Convert.ToString(data[i][9]);
                        if (val.ToLower() == "left" || val.ToLower() == "l")
                        {
                            side = "Left";
                        }

                        if (val.ToLower() == "right" || val.ToLower() == "r")
                        {
                            side = "Right";
                        }
                    }



                    if (null != data[i][13])
                    {
                        station = Convert.ToDouble(data[i][13], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (station < baseline.Start || station > baseline.End)
                    {
                        station = baseline.Start;
                    }

                    int regIndex = corridor.Baselines[baselineIndex].GetBaselineRegionIndexByStation(station);

                    // Featureline
                    if (side != "" && code != "")
                    {
                        Featureline.SideType sideType = Featureline.SideType.None;

                        if (side == "Left")
                        {
                            sideType = Featureline.SideType.Left;
                        }
                        else
                        {
                            sideType = Featureline.SideType.Right;
                        }

                        fl = corridor.GetFeaturelinesByCode(code)[baselineIndex][regIndex].First(f => f.Side == sideType);

                        if (fl == null)
                        {
                            fl = corridor.GetFeaturelines()[baselineIndex][regIndex][0];
                        }
                        useFeatureLine = true;
                    }

                    if (null != data[i][14])
                    {
                        offset = Convert.ToDouble(data[i][14], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (null != data[i][15])
                    {
                        elevation = Convert.ToDouble(data[i][15], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (null != data[i][18])
                    {
                        angleZ = Convert.ToDouble(data[i][18], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    familyInstances.Add(fi);
                    familyTypes.Add(Revit.Elements.ElementSelector.ByUniqueId(type.UniqueId) as Revit.Elements.FamilyType);
                    marks.Add(mark);
                    featurelines.Add(fl);
                    useFeaturelines.Add(!useFeatureLine);
                    stations.Add(station);
                    offsets.Add(offset);
                    elevations.Add(elevation);
                    zAngles.Add(angleZ);
                }
            }

            Utils.Log(string.Format("UtilsObjectsLocation.ReadFamilyInstanceLocationParametersForUpdate completed.", ""));

            return new Dictionary<string, object> { {"FamilyInstance", familyInstances},
                {"FamilyType", familyTypes},
                {"Mark", marks},
                {"Featureline", featurelines},
                {"UseBaseline", useFeaturelines},
                {"Station", stations},
                {"Offset", offsets},
                {"Elevation", elevations},
                {"AngleZ",zAngles} };
        }

        /// <summary>
        /// Reads the family instance location parameters for create.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "FamilyType", "Mark", "Featureline", "UseBaseline", "Station", "Offset", "Elevation", "AngleZ" })]
        public static Dictionary<string, object> ReadFamilyInstanceLocationParametersForCreate(Document doc, CivilDocument civilDocument, object[][] data)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.ReadFamilyInstanceLocationParametersForCreate started...", ""));

            IList<int> createIds = new List<int>();

            Transform tr = doc.ActiveProjectLocation.GetTotalTransform();

            Corridor corridor = civilDocument.GetCorridors().First();
            string corridorName = "";

            Baseline baseline = corridor.Baselines.First();
            int baselineIndex = baseline.Index;

            Featureline fl = null;

            // TODO: group entries by Corridor
            // TODO: group entries by BaselineIndex
            // TODO: group entries by Code
            // TODO: group entries by Side

            IList<Revit.Elements.FamilyType> familyTypes = new List<Revit.Elements.FamilyType>();
            IList<string> marks = new List<string>();
            IList<Featureline> featurelines = new List<Featureline>();
            IList<bool> useFeaturelines = new List<bool>();
            IList<double> stations = new List<double>();
            IList<double> offsets = new List<double>();
            IList<double> elevations = new List<double>();
            IList<double> zAngles = new List<double>();

            for (int i = 0; i < data.Count(); ++i)
            {
                int j = data[i].Count();

                int valCreate = Convert.ToInt32(data[i][j - 3]);
                int valUpdate = Convert.ToInt32(data[i][j - 2]);
                int valDelete = Convert.ToInt32(data[i][j - 1]);

                if (valCreate == 1 && valDelete != 1 && valUpdate != 1)
                {
                    createIds.Add(i);
                }
            }

            if (createIds.Count > 0)
            {
                foreach (int i in createIds)
                {
                    string familyName = "";
                    string typeName = "";
                    string mark = "";

                    double station = 0;
                    double offset = 0;
                    double elevation = 0;

                    double angleZ = 0;
                    string code = "";
                    string side = "";
                    bool useFeatureLine = false;

                    FamilySymbol type = null;

                    if (null != data[i][3])
                    {
                        familyName = Convert.ToString(data[i][3]);
                    }

                    if (null != data[i][4])
                    {
                        typeName = Convert.ToString(data[i][4]);
                    }

                    if (null != data[i][5])
                    {
                        mark = Convert.ToString(data[i][5]);
                    }
                    if (mark == "")
                    {
                        mark = i.ToString();
                    }

                    if (typeName != "")
                    {
                        var temp = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilySymbol))
                            .WhereElementIsElementType()
                            .Cast<FamilySymbol>()
                            .First(q => q.FamilyName == familyName && q.Name == typeName);
                        if (temp != null)
                        {
                            type = temp;
                        }
                    }

                    // Corridor
                    foreach (Corridor corr in civilDocument.GetCorridors())
                    {
                        if (corr.Name == corridorName)
                        {
                            corridor = corr;
                            break;
                        }
                    }

                    // Baseline index
                    if (null != data[i][7])
                    {
                        try
                        {
                            baselineIndex = Convert.ToInt32(data[i][7]);
                        }
                        catch
                        {
                            //System.Windows.Forms.MessageBox.Show("The specified baseline index is not valid");
                        }
                    }

                    // Baseline
                    if (baselineIndex >= 0 && baselineIndex < corridor.Baselines.Count)
                    {
                        baseline = corridor.Baselines[baselineIndex];
                    }

                    // Code
                    if (null != data[i][8])
                    {
                        foreach (string c in corridor.GetCodes())
                        {
                            if (c == Convert.ToString(data[i][8]))
                            {
                                code = c;
                                break;
                            }
                        }
                    }

                    // Side
                    if (null != data[i][9])
                    {
                        string val = Convert.ToString(data[i][9]);
                        if (val.ToLower() == "left" || val.ToLower() == "l")
                        {
                            side = "Left";
                        }

                        if (val.ToLower() == "right" || val.ToLower() == "r")
                        {
                            side = "Right";
                        }
                    }



                    if (null != data[i][13])
                    {
                        station = Convert.ToDouble(data[i][13], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (station < baseline.Start || station > baseline.End)
                    {
                        station = baseline.Start;
                    }

                    int regIndex = corridor.Baselines[baselineIndex].GetBaselineRegionIndexByStation(station);

                    // Featureline
                    if (side != "" && code != "")
                    {
                        Featureline.SideType sideType = Featureline.SideType.None;

                        if (side == "Left")
                        {
                            sideType = Featureline.SideType.Left;
                        }
                        else
                        {
                            sideType = Featureline.SideType.Right;
                        }

                        fl = corridor.GetFeaturelinesByCode(code)[baselineIndex][regIndex].First(f => f.Side == sideType);

                        if (fl == null)
                        {
                            fl = corridor.GetFeaturelines()[baselineIndex][regIndex][0];
                        }
                        useFeatureLine = true;
                    }

                    if (null != data[i][14])
                    {
                        offset = Convert.ToDouble(data[i][14], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (null != data[i][15])
                    {
                        elevation = Convert.ToDouble(data[i][15], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (null != data[i][18])
                    {
                        angleZ = Convert.ToDouble(data[i][18], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    familyTypes.Add(Revit.Elements.ElementSelector.ByUniqueId(type.UniqueId) as Revit.Elements.FamilyType);
                    marks.Add(mark);
                    featurelines.Add(fl);
                    useFeaturelines.Add(!useFeatureLine);
                    stations.Add(station);
                    offsets.Add(offset);
                    elevations.Add(elevation);
                    zAngles.Add(angleZ);
                }
            }

            Utils.Log(string.Format("UtilsObjectsLocation.ReadFamilyInstanceLocationParametersForCreate completed.", ""));

            return new Dictionary<string, object> { {"FamilyType", familyTypes},
                {"Mark", marks},
                {"Featureline", featurelines},
                {"UseBaseline", useFeaturelines},
                {"Station", stations},
                {"Offset", offsets},
                {"Elevation", elevations},
                {"AngleZ",zAngles} };
        }

        /// <summary>
        /// Updates the linear object location.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "MEPCurves" })]
        private static Dictionary<string, object> UpdateLinearObjectLocation(Document doc, CivilDocument civilDocument, object[][] data)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.UpdateLinearObjectLocation started...", ""));

            IList<int> updateIds = new List<int>();

            Transform tr = doc.ActiveProjectLocation.GetTotalTransform();

            Corridor corridor = civilDocument.GetCorridors().First();
            string corridorName = "";

            Baseline baseline = corridor.Baselines.First();
            int baselineIndex = baseline.Index;

            Featureline fl = null;

            IList<AbstractMEPCurve> MEPCurves = new List<AbstractMEPCurve>();

            for (int i = 0; i < data.Count(); ++i)
            {
                int j = data[i].Count();

                int valUpdate = Convert.ToInt32(data[i][j - 2]);
                int valDelete = Convert.ToInt32(data[i][j - 1]);

                if (valUpdate == 1 && valDelete != 1)
                {
                    updateIds.Add(i);
                }
            }

            if (updateIds.Count > 0)
            {
                foreach (int i in updateIds)
                {
                    string systemName = "";
                    string typeName = "";
                    string mark = "";

                    double startStation = 0;
                    double startOffset = 0;
                    double startElevation = 0;

                    double endStation = 0;
                    double endOffset = 0;
                    double endElevation = 0;

#pragma warning disable CS0219 // The variable 'angleZ' is assigned but its value is never used
                    double angleZ = 0;
#pragma warning restore CS0219 // The variable 'angleZ' is assigned but its value is never used
                    string code = "";
                    string side = "";
#pragma warning disable CS0219 // The variable 'useFeatureLine' is assigned but its value is never used
                    bool useFeatureLine = false;
#pragma warning restore CS0219 // The variable 'useFeatureLine' is assigned but its value is never used

                    AbstractMEPCurve mep = Revit.Elements.ElementSelector.ByUniqueId(Convert.ToString(data[i][0])) as AbstractMEPCurve;

                    if (mep == null)
                    {
                        return null;
                    }

                    var systemTypes = new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
                    .OfClass(typeof(MEPSystemType))
                    .WhereElementIsElementType();

                    MEPSystemType type = null;

                    if (null != data[i][3])
                    {
                        systemName = Convert.ToString(data[i][3]);
                    }

                    if (null != data[i][4])
                    {
                        typeName = Convert.ToString(data[i][4]);
                    }

                    if (null != data[i][5])
                    {
                        mark = Convert.ToString(data[i][5]);
                    }
                    if (mark == "")
                    {
                        mark = i.ToString();
                    }

                    // TODO: Filter by MEPSystemType Domain
                    if (systemName != "" && systemName != "CableTrayConduit")
                    {
                        var temp = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilySymbol))
                            .WhereElementIsElementType()
                            .Cast<MEPSystemType>()
                            .First(q => q.Name == systemName);
                        if (temp != null)
                        {
                            type = temp;
                        }
                    }

                    // Corridor
                    foreach (Corridor corr in civilDocument.GetCorridors())
                    {
                        if (corr.Name == corridorName)
                        {
                            corridor = corr;
                            break;
                        }
                    }

                    // Baseline index
                    if (null != data[i][7])
                    {
                        try
                        {
                            baselineIndex = Convert.ToInt32(data[i][7]);
                        }
                        catch
                        {
                            //System.Windows.Forms.MessageBox.Show("The specified baseline index is not valid");
                        }
                    }

                    // Baseline
                    if (baselineIndex >= 0 && baselineIndex < corridor.Baselines.Count)
                    {
                        baseline = corridor.Baselines[baselineIndex];
                    }

                    // Code
                    if (null != data[i][8])
                    {
                        foreach (string c in corridor.GetCodes())
                        {
                            if (c == Convert.ToString(data[i][8]))
                            {
                                code = c;
                                break;
                            }
                        }
                    }

                    // Side
                    if (null != data[i][9])
                    {
                        string val = Convert.ToString(data[i][9]);
                        if (val.ToLower() == "left" || val.ToLower() == "l")
                        {
                            side = "Left";
                        }

                        if (val.ToLower() == "right" || val.ToLower() == "r")
                        {
                            side = "Right";
                        }
                    }

                    // Featureline
                    if (side != "" && code != "")
                    {
                        Featureline.SideType sideType = Featureline.SideType.None;

                        if (side == "Left")
                        {
                            sideType = Featureline.SideType.Left;
                        }
                        else
                        {
                            sideType = Featureline.SideType.Right;
                        }

                        int regIndex = corridor.Baselines[baselineIndex].GetBaselineRegionIndexByStation(startStation);

                        fl = corridor.GetFeaturelinesByCode(code)[baselineIndex][regIndex].First(f => f.Side == sideType);

                        if (fl == null)
                        {
                            fl = corridor.GetFeaturelines()[baselineIndex][regIndex][0];
                        }

                        useFeatureLine = true;
                    }

                    if (null != data[i][13])
                    {
                        startStation = Convert.ToDouble(data[i][13], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (startStation < baseline.Start || startStation > baseline.End)
                    {
                        startStation = baseline.Start;
                    }

                    if (null != data[i][14])
                    {
                        startOffset = Convert.ToDouble(data[i][14], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (null != data[i][15])
                    {
                        startElevation = Convert.ToDouble(data[i][15], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (null != data[i][19])
                    {
                        endStation = Convert.ToDouble(data[i][19], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (endStation < baseline.Start || endStation > baseline.End)
                    {
                        endStation = baseline.End;
                    }

                    if (null != data[i][20])
                    {
                        endOffset = Convert.ToDouble(data[i][20], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    if (null != data[i][21])
                    {
                        endElevation = Convert.ToDouble(data[i][21], System.Globalization.CultureInfo.InvariantCulture);
                    }

                    Autodesk.DesignScript.Geometry.Line line = Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(
                        Autodesk.DesignScript.Geometry.Point.ByCoordinates(startElevation, startOffset, startElevation),
                        Autodesk.DesignScript.Geometry.Point.ByCoordinates(endElevation, endOffset, endElevation));

                    mep.Location = line;
                    mep.SetParameterByName("Mark", mark);
                }
            }

            Utils.Log(string.Format("UtilsObjectsLocation.UpdateLinearObjectLocation completed.", ""));

            return new Dictionary<string, object> { { "FamilyInstance", MEPCurves } };
        }

        /// <summary>
        /// Updates the family instance.
        /// </summary>
        /// <param name="familyInstance">The family instance.</param>
        /// <param name="familyType">Type of the family.</param>
        /// <param name="featureline">The featureline.</param>
        /// <param name="useBaseline">if set to <c>true</c> [use baseline].</param>
        /// <param name="station">The station.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="elevation">The elevation.</param>
        /// <param name="angleZ">The angle z.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static Revit.Elements.FamilyInstance UpdateFamilyInstance(Revit.Elements.FamilyInstance familyInstance, Revit.Elements.FamilyType familyType, Featureline featureline, bool useBaseline = false, double station = 0, double offset = 0, double elevation = 0, double angleZ = 0)
        {
            // TODO There could be a BIG improvement if we handle a single Revit Transaction for all the objects.
            // The caveat is to porvide also a list for each of all the other parameters (e.g. stations, offsets, elevations, etc.)
            Utils.Log(string.Format("UtilsObjectsLocation.UpdateFamilyInstance started...", ""));

            Document doc = DocumentManager.Instance.CurrentDBDocument;

            var totalTransform = RevitUtils.DocumentTotalTransform();

            var totalTransformInverse = RevitUtils.DocumentTotalTransformInverse();

            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

            if (!SessionVariables.ParametersCreated)
            {
                CheckParameters(doc); 
            }

            CoordinateSystem cs = CoordinateSystem.Identity();
            Autodesk.DesignScript.Geometry.Point newLocPoint = Autodesk.DesignScript.Geometry.Point.Origin();


            if (!useBaseline)
            {
                cs = featureline.CoordinateSystemByStation(station);
            }
            else
            {
                cs = featureline.Baseline.CoordinateSystemByStation(station);
            }

            newLocPoint = Autodesk.DesignScript.Geometry.Point.ByCoordinates(offset, 0, elevation).Transform(cs) as Autodesk.DesignScript.Geometry.Point;

            // newLocPoint is in WCS like in Civil 3D, before using it it must be transformed in the Project Base Coordinates
            //Transform transform = doc.ActiveProjectLocation.GetTotalTransform();

            RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            var lp = familyInstance.InternalElement.Location as LocationPoint;
            if (null != lp)
            {
                var temp = newLocPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                lp.Point = temp.ToXyz();

                temp.Dispose();
            }

            FamilyInstance fi = familyInstance.InternalElement as FamilyInstance;

            familyInstance.SetRotation(cs.XAxis.AngleAboutAxis(Vector.XAxis(), cs.ZAxis) + angleZ);

            familyInstance.SetParameterByName("Type", familyType);
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Corridor.Name, featureline.Baseline.CorridorName);
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.BaselineIndex.Name, featureline.Baseline.Index);
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Code.Name, featureline.Code);
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Side.Name, featureline.Side.ToString());
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Station.Name, Math.Round(station, 3));
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Offset.Name, Math.Round(offset, 3));
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Elevation.Name, Math.Round(elevation, 3));
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(newLocPoint.X, 3));
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(newLocPoint.Y, 3));
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(newLocPoint.Z, 3));
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.AngleZ.Name, Math.Round(angleZ, 3));
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);

            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

            cs.Dispose();

            newLocPoint.Dispose();

            Utils.Log(string.Format("UtilsObjectsLocation.UpdateFamilyInstance completed.", ""));

            return familyInstance;
        }



        /// <summary>
        /// Creates the family instance.
        /// </summary>
        /// <param name="familyType">Type of the family.</param>
        /// <param name="featureline">The featureline.</param>
        /// <param name="useBaseline">if set to <c>true</c> [use baseline].</param>
        /// <param name="station">The station.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="elevation">The elevation.</param>
        /// <param name="angleZ">The angle z.</param>
        /// <returns></returns>
        public static Revit.Elements.FamilyInstance CreateFamilyInstance(Revit.Elements.FamilyType familyType, Featureline featureline, bool useBaseline = false, double station = 0, double offset = 0, double elevation = 0, double angleZ = 0)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.CreateFamilyInstance started...", ""));

            Document doc = DocumentManager.Instance.CurrentDBDocument;

            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

            if (!SessionVariables.ParametersCreated)
            {
                CheckParameters(doc); 
            }

            CoordinateSystem cs = CoordinateSystem.Identity();
            Autodesk.DesignScript.Geometry.Point newLocPoint = Autodesk.DesignScript.Geometry.Point.Origin();

            if (!useBaseline)
            {
                cs = featureline.CoordinateSystemByStation(station);
            }
            else
            {
                cs = featureline.Baseline.CoordinateSystemByStation(station);
            }

            newLocPoint = Autodesk.DesignScript.Geometry.Point.ByCoordinates(offset, 0, elevation).Transform(cs) as Autodesk.DesignScript.Geometry.Point;

            Utils.Log(string.Format("newLocPoint: {0}", newLocPoint));

            // totalTransform : WCS to PBP
            // totalTransform.Inverse : PBP to WCS

            RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            var totalTransform = RevitUtils.DocumentTotalTransform();

            Revit.Elements.FamilyInstance familyInstance = Revit.Elements.FamilyInstance.ByPoint(familyType, newLocPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point);

            Utils.Log(string.Format("familyInstance: {0} {1}", familyInstance, familyInstance.Id));

            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

            var xAxis = cs.XAxis.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Vector;

            familyInstance.SetRotation(xAxis.AngleAboutAxis(Vector.XAxis(), cs.ZAxis) + angleZ);  // [20181007]

            Utils.Log(string.Format("SetRotation: {0}", xAxis.AngleAboutAxis(Vector.XAxis(), cs.ZAxis) + angleZ));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Corridor.Name, featureline.Baseline.CorridorName);

            Utils.Log(string.Format("ADSK_Corridor: {0}", featureline.Baseline.CorridorName));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.BaselineIndex.Name, featureline.Baseline.Index);

            Utils.Log(string.Format("ADSK_BaselineIndex: {0}", featureline.Baseline.Index));

            // 1.1.0 : BaselineRegionSystem to prevent errors from changes in the station during the update.
            // The Revit element is referenced to the Region Index in the corridor
            // If the relative station still makes sense for the updated region, the element can be updated too.
            // If the station is not so important there is a normalized value that will scale the positioning of the Revit element along the updated region.
            // if the reltaive station is larger than the updated RelativeEnd of the region the Revit element will not be updated.
            familyInstance.SetParameterByName(ADSK_Parameters.Instance.RegionIndex.Name, featureline.BaselineRegionIndex);  // 1.1.0

            Utils.Log(string.Format("ADSK_RegionIndex: {0}", featureline.BaselineRegionIndex));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.RegionRelative.Name, station - featureline.Start);  // 1.1.0

            Utils.Log(string.Format("ADSK_RegionRelative: {0}", station - featureline.Start));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.RegionNormalized.Name, (station - featureline.Start) / (featureline.End - featureline.Start));  // 1.1.0

            Utils.Log(string.Format("ADSK_RegionNormalized: {0}", (station - featureline.Start) / (featureline.End - featureline.Start)));

            // 1.1.0

            if (!useBaseline)
            {
                familyInstance.SetParameterByName(ADSK_Parameters.Instance.Code.Name, featureline.Code);
                familyInstance.SetParameterByName(ADSK_Parameters.Instance.Side.Name, featureline.Side.ToString());
            }
            else
            {
                familyInstance.SetParameterByName(ADSK_Parameters.Instance.Code.Name, "<None>");
                familyInstance.SetParameterByName(ADSK_Parameters.Instance.Side.Name, Featureline.SideType.None.ToString());
            }

            Utils.Log(string.Format("ADSK_Code: {0}", featureline.Code));
            Utils.Log(string.Format("ADSK_Side: {0}", featureline.Side));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Station.Name, Math.Round(station, 3));

            Utils.Log(string.Format("ADSK_Station: {0}", Math.Round(station, 3)));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Offset.Name, Math.Round(offset, 3));

            Utils.Log(string.Format("ADSK_Offset: {0}", Math.Round(offset, 3)));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Elevation.Name, Math.Round(elevation, 3));

            Utils.Log(string.Format("ADSK_Elevation: {0}", Math.Round(elevation, 3)));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(newLocPoint.X, 3));

            Utils.Log(string.Format("ADSK_X: {0}", Math.Round(newLocPoint.X, 3)));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(newLocPoint.Y, 3));

            Utils.Log(string.Format("ADSK_Y: {0}", Math.Round(newLocPoint.Y, 3)));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(newLocPoint.Z, 3));

            Utils.Log(string.Format("ADSK_Z: {0}", Math.Round(newLocPoint.Z, 3)));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.AngleZ.Name, Math.Round(angleZ, 3));

            Utils.Log(string.Format("ADSK_AngleZ: {0}", Math.Round(angleZ, 3)));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);

            Utils.Log(string.Format("ADSK_Update: {0}", true));

            familyInstance.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);

            Utils.Log(string.Format("ADSK_Delete: {0}", false));

            cs.Dispose();
            newLocPoint.Dispose();
            xAxis.Dispose();

            Utils.Log(string.Format("UtilsObjectsLocation.CreateFamilyInstance completed.", ""));

            return familyInstance;
        }

        /// <summary>
        /// Revit link by station offset elevation.
        /// </summary>
        /// <param name="revitLinkType">Type of the revit link.</param>
        /// <param name="featureline">The featureline.</param>
        /// <param name="station">The station.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="elevation">The elevation.</param>
        /// <param name="rotate">if set to <c>true</c> [rotate].</param>
        /// <param name="rotation">The rotation.</param>
        /// <returns></returns>
        public static string RevitLinkByStationOffsetElevation(Revit.Elements.Element revitLinkType, Featureline featureline, double station, double offset = 0, double elevation = 0, bool rotate = true, double rotation = 0)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.RevitLinkByStationOffsetElevation started...", ""));

            string side = "";

            Autodesk.DesignScript.Geometry.Point origin = featureline.PointByStationOffsetElevation(station, offset, elevation, false);

            if (featureline.Side == Featureline.SideType.None)
            {
                if (featureline.Baseline.GetArrayStationOffsetElevationByPoint(origin)[1] <= 0)
                {
                    side = "Left";
                }
                else
                {
                    side = "Right";
                }
            }
            else
            {
                side = featureline.Side.ToString();
            }

            // TODO: Since Revit 2018 Revit Links can have parameters assigned...
            string name = string.Format("{0}_{1}_{2}_{3}_{4:0.00}_{5:0.00}_{6:0.00}_{7:0.00}", featureline.Baseline.CorridorName, featureline.Baseline.Index.ToString(), featureline.Code, side, station, offset, elevation, rotation);

            Document doc = DocumentManager.Instance.CurrentDBDocument;

            RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(doc);

            RevitLinkInstance rli = null;

            bool found = false;

            foreach (RevitLinkInstance i in new FilteredElementCollector(doc)
                 .OfClass(typeof(RevitLinkInstance))
                 .WhereElementIsNotElementType()
                 .Cast<RevitLinkInstance>()
                 .Where(x => x.GetTypeId().IntegerValue.Equals(revitLinkType.InternalElement.Id.IntegerValue)))
            {
                if (i.Parameters.Cast<Parameter>().First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.RVT_LINK_INSTANCE_NAME)).AsString() == name)
                {
                    found = true;
                    rli = i;
                    break;
                }
            }

            if (!found)
            {
                rli = RevitLinkInstance.Create(DocumentManager.Instance.CurrentDBDocument, revitLinkType.InternalElement.Id);
            }

            rli.MoveOriginToHostOrigin(true);

            Location location = rli.Location;

            var totalTransform = RevitUtils.DocumentTotalTransform();

            origin = origin.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;

            location.Move(origin.ToXyz() - XYZ.Zero);

            Vector xLocal = featureline.CoordinateSystemByStation(station).XAxis;

            double newAngle = ProjectPositionUtils.Instance.Angle; // 0;

            double angle = DegToRadians(xLocal.AngleAboutAxis(Vector.XAxis(), Vector.ZAxis()) + rotation) + newAngle;

            if (side == "Left" && rotate)
            {
                angle += Math.PI;
            }

            location.Rotate(Autodesk.Revit.DB.Line.CreateBound(origin.ToXyz(), origin.ToXyz() + XYZ.BasisZ), -angle);

            rli.Parameters.Cast<Parameter>().First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.RVT_LINK_INSTANCE_NAME)).Set(name);

            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

            origin.Dispose();
            xLocal.Dispose();

            Utils.Log(string.Format("UtilsObjectsLocation.RevitLinkByStationOffsetElevation completed.", ""));

            return name;
        }

        /// <summary>
        /// Create a Named site by station offset elevation.
        /// </summary>
        /// <param name="featureline">The featureline.</param>
        /// <param name="station">The station.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="elevation">The elevation.</param>
        /// <param name="rotate">if set to <c>true</c> [rotate].</param>
        /// <param name="rotation">The rotation.</param>
        /// <returns></returns>
        public static string NamedSiteByStationOffsetElevation(Featureline featureline, double station, double offset = 0, double elevation = 0, bool rotate = true, double rotation = 0)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.NamedSiteByStationOffsetElevation started...", ""));

            string side = "";

            Autodesk.DesignScript.Geometry.Point origin = featureline.PointByStationOffsetElevation(station, offset, elevation, false);

            if (featureline.Side == Featureline.SideType.None)
            {
                if (featureline.Baseline.GetArrayStationOffsetElevationByPoint(origin)[1] <= 0)
                {
                    side = "Left";
                }
                else
                {
                    side = "Right";
                }
            }
            else
            {
                side = featureline.Side.ToString();
            }

            string name = string.Format("{0}_{1}_{2}_{3}_{4:0.00}_{5:0.00}_{6:0.00}_{7:0.00}", featureline.Baseline.CorridorName, featureline.Baseline.Index.ToString(), featureline.Code, side, station, offset, elevation, rotation);

            Document doc = DocumentManager.Instance.CurrentDBDocument;

            RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(doc);

            bool found = false;

            ProjectLocation projectLocation;

            foreach (ProjectLocation pl in doc.ProjectLocations)
            {
                if (pl.Name == name)
                {
                    found = true;
                    break;
                }
            }

            if (found)
            {
                //update
                projectLocation = doc.ProjectLocations.Cast<ProjectLocation>().First(x => x.Name == name);
            }
            else
            {
                // create new Project Location
                projectLocation = doc.ProjectLocations.Cast<ProjectLocation>().First().Duplicate(name);
            }

            ProjectPosition position = ProjectPositionUtils.Instance.ProjectPosition;  // null;

            var totalTransform = RevitUtils.DocumentTotalTransform();

            origin = origin.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;

            Vector xLocal = featureline.CoordinateSystemByStation(station).XAxis;

            double newAngle = ProjectPositionUtils.Instance.Angle; // 0;

            double angle = DegToRadians(xLocal.AngleAboutAxis(Vector.XAxis(), Vector.ZAxis()) + rotation) + newAngle;

            if (side == "Left" && rotate)
            {
                angle += Math.PI;
            }

            position.Angle = -angle;
            position.EastWest = origin.ToXyz().X;
            position.NorthSouth = origin.ToXyz().Y;
            position.Elevation = origin.ToXyz().Z;

            ProjectPositionUtils.Instance.SetProjectPosition(projectLocation, position);

            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

            xLocal.Dispose();
            origin.Dispose();

            Utils.Log(string.Format("UtilsObjectsLocation.NamedSiteByStationOffsetElevation completed.", ""));

            return name;
        }

        /// <summary>
        /// Given a Point it returns the first the Featureline that meets the argument values.
        /// </summary>
        /// <param name="p">The p.</param>
        /// <param name="corridor">The corridor.</param>
        /// <param name="baselineIndex">Index of the baseline.</param>
        /// <param name="code">The code.</param>
        /// <param name="side">The side.</param>
        /// <returns></returns>
        public static Featureline ClosestFeaturelineByPoint(Autodesk.DesignScript.Geometry.Point p, Corridor corridor, int baselineIndex, string code, string side)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.ClosestFeaturelineByPoint started...", ""));

            Featureline f = null;

            Baseline b = corridor.Baselines[baselineIndex];

            double station = 0;
            double offset = 0;

            b._baseline.Alignment.StationOffset(p.X, p.Y, out station, out offset);

            double[] stations = b.Stations;

            double min = stations.Min();
            double max = stations.Max();

            if ((station < min && Math.Abs(station - min) > 0.00001) || (station > max && Math.Abs(station - max) > 0.00001))
            {
                // don't look for the featurline
                Utils.Log(string.Format("ERROR: the point is outside of the Baseline station range.", ""));
                return f;
            }

            f = b.GetFeaturelinesByCodeStation(code, station).First(x => x.Side.ToString() == side);

            Utils.Log(string.Format("UtilsObjectsLocation.ClosestFeaturelineByPoint completed.", ""));

            return f;
        }

        /// <summary>
        /// Returns the Featureline that meets the argument values.
        /// </summary>
        /// <param name="corridor">The corridor.</param>
        /// <param name="baselineIndex">Index of the baseline.</param>
        /// <param name="regionIndex">Index of the region.</param>
        /// <param name="code">The code.</param>
        /// <param name="side">The side.</param>
        /// <returns></returns>
        public static Featureline FeaturelineByParameter(Corridor corridor, int baselineIndex, int regionIndex, string code, string side)  // 1.1.0
        {
            Utils.Log(string.Format("UtilsObjectsLocation.FeaturelineByParameter started...", ""));

            Featureline f = null;

            f = corridor.GetFeaturelinesByCode(code)[baselineIndex][regionIndex].Cast<Featureline>().First(x => x.Side.ToString() == side);

            Utils.Log(string.Format("UtilsObjectsLocation.FeaturelineByParameter completed.", ""));

            return f;
        }

        /// <summary>
        /// Given a Revit Element it returns the first Featureline that meets the argument values.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="corridor">The corridor.</param>
        /// <param name="baselineIndex">Index of the baseline.</param>
        /// <param name="code">The code.</param>
        /// <param name="side">The side.</param>
        /// <returns></returns>
        public static Featureline ClosestFeaturelineByElement(Revit.Elements.Element element, Corridor corridor, int baselineIndex, string code, string side)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.ClosestFeaturelineByElement started...", ""));

            Autodesk.DesignScript.Geometry.Point p = null;

            Geometry location = element.GetLocation();

            var totalTransform = RevitUtils.DocumentTotalTransform();

            var totalTransformInverse = RevitUtils.DocumentTotalTransformInverse();

            if (location is Autodesk.DesignScript.Geometry.Point)
            {
                p = location.Transform(totalTransformInverse) as Autodesk.DesignScript.Geometry.Point;
            }
            else if (location is Autodesk.DesignScript.Geometry.Curve)
            {
                Autodesk.DesignScript.Geometry.Curve curve = location as Autodesk.DesignScript.Geometry.Curve;

                p = curve.StartPoint.Transform(totalTransformInverse) as Autodesk.DesignScript.Geometry.Point; ;
            }
            else
            {
                Utils.Log(string.Format("UtilsObjectsLocation.ClosestFeaturelineByElement The element has to be a single point or based on curve object.", ""));

                return null;
            }

            location.Dispose();

            Utils.Log(string.Format("UtilsObjectsLocation.ClosestFeaturelineByElement completed.", ""));

            return ClosestFeaturelineByPoint(p, corridor, baselineIndex, code, side);
        }


        /// <summary>
        /// Create walls from surface.
        /// </summary>
        /// <param name="surface"></param>
        /// <param name="wallType"></param>
        /// <param name="structural"></param>
        /// <returns></returns>
        /// <remarks>The wall is recreated but not updated. The input surface must be planar and its normal must be orthogonal to the world Z Axis.</remarks>
        public static Revit.Elements.Wall WallBySurface(Autodesk.DesignScript.Geometry.Surface surface, Revit.Elements.WallType wallType, bool structural = false)
        {
            Utils.Log(string.Format("UtilsObjectsLocation.WallBySurface started...", ""));

            if (!PolyCurve.ByJoinedCurves(surface.PerimeterCurves()).IsPlanar)
            {
                Utils.Log("ERROR: Surface is not planar!");

                throw new Exception("Surface must be planar");
            }

            var doc = RevitServices.Persistence.DocumentManager.Instance.CurrentDBDocument;

            if (!SessionVariables.ParametersCreated)
            {
                CheckParameters(doc); 
            }

            var bb = BoundingBox.ByGeometry(new List<Geometry>() { surface });

            Utils.Log("Bounding Box created...");

            Wall newWall = null;


            // This creates a new wall and deletes the old one
            RevitServices.Transactions.TransactionManager.Instance.EnsureInTransaction(doc);

            //There was a modelcurve, try and set sketch plane
            // if you can't, rebuild 

            try
            {
                Utils.Log("Looking for existing wall...");

                var existingWall = ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.Wall>(doc);

                if (existingWall != null && existingWall.Location is Autodesk.Revit.DB.LocationCurve)
                {
                    Utils.Log("Wall Found...");

                    var wallLocation = existingWall.Location as Autodesk.Revit.DB.LocationCurve;

                    if (wallLocation.Curve is Autodesk.Revit.DB.Line)
                    {
                        Utils.Log("Wall Location Curve Found...");

                        var plane = surface.CoordinateSystemAtParameter(0.5, 0.5).XYPlane;
                        var p = wallLocation.Curve.ToProtoType().PullOntoPlane(plane).PointAtParameter(0.5);
                        p = Autodesk.DesignScript.Geometry.Point.ByCoordinates(p.X, p.Y);

                        var q = wallLocation.Curve.ToProtoType().PullOntoPlane(plane).PullOntoPlane(Autodesk.DesignScript.Geometry.Plane.XY());

                        if (p.DistanceTo(q) < 0.001)
                        {
                            doc.Delete(existingWall.Id);
                        }
                    }
                }
            }
            catch
            {
                Utils.Log("ERROR: Issues with the trace");

                throw new Exception("Issues with the trace");
            }

            // new create the geometry
            var max = bb.MaxPoint.ToXyz().Z;
            var min = bb.MinPoint.ToXyz().Z;

            var curves = new List<Autodesk.Revit.DB.Curve>();

            foreach (var c in surface.PerimeterCurves())
            {
                curves.Add(c.ToRevitType());
            }

            Utils.Log("New profile acquired...");

            WallType rvtWallType = wallType.InternalElement as WallType;

            if (rvtWallType.Kind == WallKind.Curtain)
            {
                if (structural == true)
                {
                    structural = false;
                }
            }

            try
            {
                newWall = Wall.Create(doc, curves, structural);
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: {0}", ex.Message));
            }

            Utils.Log(string.Format("New wall created {0}...", newWall.Id.IntegerValue));

            newWall.WallType = rvtWallType;

            Utils.Log(string.Format("Wall type assigned...", ""));

            var level = doc.GetElement(newWall.LevelId) as Level;

            Utils.Log(string.Format("Level acquired...", ""));
            try
            {

                newWall.Parameters.Cast<Parameter>().First(x => x.Id.IntegerValue == (int)BuiltInParameter.WALL_BASE_OFFSET).Set(min - level.Elevation);

                Utils.Log(string.Format("WALL_BASE_OFFSET Set...", ""));
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: {0}", ex.Message));
            }

            try
            {
                newWall.Parameters.Cast<Parameter>().First(x => x.Id.IntegerValue == (int)BuiltInParameter.WALL_HEIGHT_TYPE).Set(ElementId.InvalidElementId);

                Utils.Log(string.Format("WALL_HEIGHT_TYPE Set...", ""));
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: {0}", ex.Message));
            }

            try
            {
                newWall.Parameters.Cast<Parameter>().First(x => x.Id.IntegerValue == (int)BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(max - min);

                Utils.Log(string.Format("WALL_USER_HEIGHT_PARAM Set...", ""));
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: {0}", ex.Message));
            }

            try
            {
                newWall.Parameters.Cast<Parameter>().First(x => x.Id.IntegerValue == (int)BuiltInParameter.WALL_KEY_REF_PARAM).Set(0);

                Utils.Log(string.Format("WALL_KEY_REF_PARAM Set...", ""));
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: {0}", ex.Message));
            }

            var revitWall = newWall;

            // delete the element stored in trace and add this new one
            ElementBinder.CleanupAndSetElementForTrace(doc, revitWall);

            Utils.Log(string.Format("Cleanup successful...", ""));

            RevitServices.Transactions.TransactionManager.Instance.TransactionTaskDone();

            var wall = Revit.Elements.ElementWrapper.Wrap(revitWall, true);

            Utils.Log(string.Format("UtilsObjectsLocation.WallBySurface completed.", ""));

            return wall;
        }
    }

    /// <summary>
    /// Json Converter fot MultiPoint objects
    /// </summary>
    [SupressImportIntoVM()]
    public class MultiPointConverter : JsonConverter
    {
        /// <summary>
        /// Can Convert
        /// </summary>
        /// <param name="objectType"></param>
        /// <returns></returns>
        public override bool CanConvert(Type objectType)
        {
            return typeof(MultiPoint).IsAssignableFrom(objectType);
        }
        /// <summary>
        /// Read Json
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="objectType"></param>
        /// <param name="existingValue"></param>
        /// <param name="serializer"></param>
        /// <returns></returns>
        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            JObject obj = JObject.Load(reader);
            return new MultiPoint() { ShapePoints = obj["ShapePoints"].ToObject<ShapePointArray>() };
        }
        /// <summary>
        /// Can write
        /// </summary>
        public override bool CanWrite
        {
            get
            {
                return false;
            }
        }
        /// <summary>
        /// Write Json
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="value"></param>
        /// <param name="serializer"></param>
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Json Converter for ShapePoints objects
    /// </summary>
    [SupressImportIntoVM()]
    public class ShapePointArrayConverter : JsonConverter
    {
        /// <summary>
        /// Can Convert
        /// </summary>
        /// <param name="objectType"></param>
        /// <returns></returns>
        public override bool CanConvert(Type objectType)
        {
            return typeof(ShapePointArray).IsAssignableFrom(objectType);
        }

        /// <summary>
        /// Read Json
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="objectType"></param>
        /// <param name="existingValue"></param>
        /// <param name="serializer"></param>
        /// <returns></returns>
        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            JObject obj = JObject.Load(reader);
            return new ShapePointArray() { Points = obj["Points"].ToObject<IList<ShapePoint>>() };
        }

        /// <summary>
        /// Can Write
        /// </summary>
        public override bool CanWrite
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// Write Json
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="value"></param>
        /// <param name="serializer"></param>
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Json Converter for ShapePoint objects
    /// </summary>
    [SupressImportIntoVM()]
    public class ShapePointConverter : JsonConverter
    {
        /// <summary>
        /// Can Convert
        /// </summary>
        /// <param name="objectType"></param>
        /// <returns></returns>
        public override bool CanConvert(Type objectType)
        {
            return typeof(ShapePoint).IsAssignableFrom(objectType);
        }


        /// <summary>
        /// Read Json
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="objectType"></param>
        /// <param name="existingValue"></param>
        /// <param name="serializer"></param>
        /// <returns></returns>
        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            JObject obj = JObject.Load(reader);

            var corridor = obj["Corridor"].Value<string>();
            var code = obj["Code"].Value<string>();
            var uid = obj["UniqueId"].Value<string>();
            var id = obj["Id"].Value<int>();
            var bi = obj["BaselineIndex"].Value<int>();
            var regionid = obj["RegionIndex"].Value<int>();
            var regionRelative = obj["RegionRelative"].Value<double>();
            var regionNormalized = obj["RegionNormalized"].Value<double>();
            var side_id = obj["Side"].Value<int>();
            var station = obj["Station"].Value<double>();
            var offset = obj["Offset"].Value<double>();
            var elevation = obj["Elevation"].Value<double>();

            var sp = new ShapePoint(uid, id, corridor, bi, regionid, regionRelative, regionNormalized, code, side_id, station, offset, elevation);

            return sp;
        }

        /// <summary>
        /// Can Write
        /// </summary>
        public override bool CanWrite
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// Write Json
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="value"></param>
        /// <param name="serializer"></param>
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }
}

