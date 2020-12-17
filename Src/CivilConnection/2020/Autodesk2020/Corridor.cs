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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime;
using System.Runtime.InteropServices;

using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AECC.Interop.UiRoadway;
using Autodesk.AECC.Interop.Roadway;
using Autodesk.AECC.Interop.Land;
using Autodesk.AECC.Interop.UiLand;
using System.Reflection;

using Autodesk.DesignScript.Runtime;
using Autodesk.DesignScript.Geometry;

//using Dynamo.Wpf.Nodes;

using ProtoCore.Properties;
using System.Xml;
using System.Globalization;

namespace CivilConnection
{
    /// <summary>
    /// Corridor obejct type.
    /// </summary>
    public class Corridor
    {
        #region PRIVATE PROPERTIES
        /// <summary>
        /// The corridor
        /// </summary>
        internal AeccCorridor _corridor;
        /// <summary>
        /// The baselines
        /// </summary>
        internal IList<Baseline> _baselines;
        /// <summary>
        /// The document
        /// </summary>
        internal AeccRoadwayDocument _document;
        /// <summary>
        /// Corridor Applied Subassembly Shapes
        /// </summary>
        private IList<IList<IList<AppliedSubassemblyShape>>> _shapes = new List<IList<IList<AppliedSubassemblyShape>>>();
        /// <summary>
        /// Corridor Applied Subassembly Links
        /// </summary>
        private IList<IList<IList<AppliedSubassemblyLink>>> _links = new List<IList<IList<AppliedSubassemblyLink>>>();
        /// <summary>
        /// Indicates if the corridor feature lines have been already extracted
        /// </summary>
        internal bool _corridorFeaturelinesXMLExported;
        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the baselines.
        /// </summary>
        /// <value>
        /// The baselines.
        /// </value>
        public IList<Baseline> Baselines { get { return _baselines; } }
        /// <summary>
        /// Gets the Corridor name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get { return _corridor.DisplayName; } }
        /// <summary>
        /// Gets the internal element.
        /// </summary>
        /// <value>
        /// The internal element.
        /// </value>
        internal AeccCorridor InternalElement { get { return this._corridor; } }

        /// <summary>
        /// Gets the corridor applied subassembly shapes.
        /// </summary>
        public IList<IList<IList<AppliedSubassemblyShape>>> Shapes
        {
            get
            {
                if (this._shapes.Count == 0)
                {
                    this._shapes = GetShapesFromXML();
                }

                return this._shapes;
            }
        }

        /// <summary>
        /// Gets the corridor applied subassembly links.
        /// </summary>
        public IList<IList<IList<AppliedSubassemblyLink>>> Links
        {
            get
            {
                if (this._links.Count == 0)
                {
                    this._links = GetLinksFromXML();
                }

                return this._links;
            }
        }
        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Initializes a new instance of the <see cref="Corridor"/> class.
        /// </summary>
        /// <param name="corridor">The corridor.</param>
        /// <param name="doc">The document.</param>
        internal Corridor(AeccCorridor corridor, AeccRoadwayDocument doc)
        {
            this._corridor = corridor;
            this._document = doc;
            IList<Baseline> bls = new List<Baseline>();

            int index = 0;
            foreach (AeccBaseline b in corridor.Baselines)
            {
                bls.Add(new Baseline(b, index, this));
                ++index;
            }

            this._baselines = bls;
            this._corridorFeaturelinesXMLExported = false;
        }

        /// <summary>
        /// Rebuilds the Corridor in Civil 3D.
        /// </summary>
        /// <returns></returns>
        public Corridor Rebuild()
        {
            this.InternalElement.Rebuild();
            return this;
        }

        #endregion

        #region PRIVATE METHODS

        /// <summary>
        /// Returns the points that define the subassemblies in a corridor organized by:
        /// Corridor &gt; Baseline &gt; Region &gt; Assembly &gt; Subassembly
        /// </summary>
        /// <returns>
        /// The list of points that define each subassembly in the corridor
        /// </returns>
        /// <search> Subassembly, section</search>
        private IList<IList<IList<IList<IList<IList<Point>>>>>> GetPointsBySubassembly()
        {
            Utils.Log(string.Format("Corridor.GetPointsBySubassembly started...", ""));

            IList<IList<IList<IList<IList<IList<Point>>>>>> output = new List<IList<IList<IList<IList<IList<Point>>>>>>();

            foreach (AeccBaseline b in this._corridor.Baselines)
            {
                IList<IList<IList<IList<IList<Point>>>>> baselineColl = new List<IList<IList<IList<IList<Point>>>>>();

                foreach (AeccBaselineRegion blr in b.BaselineRegions)
                {
                    IList<IList<IList<IList<Point>>>> regionColl = new List<IList<IList<IList<Point>>>>();

                    foreach (AeccAppliedAssembly assembly in blr.AppliedAssemblies)
                    {
                        IList<IList<IList<Point>>> assemblyColl = new List<IList<IList<Point>>>();

                        foreach (AeccAppliedSubassembly sub in assembly.AppliedSubassemblies)
                        {
                            IList<IList<Point>> subColl = new List<IList<Point>>();

                            foreach (AeccCalculatedShape shape in sub.CalculatedShapes)
                            {
                                IList<Point> shapeColl = new List<Point>();

                                foreach (IAeccCalculatedLink link in shape.CalculatedLinks)
                                {
                                    foreach (AeccCalculatedPoint p in link.CalculatedPoints)
                                    {
                                        dynamic xyz = b.StationOffsetElevationToXYZ(p.GetStationOffsetElevationToBaseline());

                                        Point point = Point.ByCoordinates(xyz[0], xyz[1], xyz[2]);

                                        shapeColl.Add(point);
                                    }
                                }

                                subColl.Add(Point.PruneDuplicates(shapeColl));
                            }

                            assemblyColl.Add(subColl);
                        }

                        regionColl.Add(assemblyColl);
                    }

                    baselineColl.Add(regionColl);
                }

                output.Add(baselineColl);
            }

            Utils.Log(string.Format("Corridor.GetPointsBySubassembly completed.", ""));

            return output;
        }

        /// <summary>
        /// Gets the corridor surfaces.
        /// </summary>
        /// <returns></returns>
        private IList<IList<Surface>> GetCorridorSurfaces()
        {
            Utils.Log(string.Format("Corridor.GetCorridorSurfaces started...", ""));

            IList<IList<Surface>> output = new List<IList<Surface>>();

            if (null != this._corridor.CorridorSurfaces)
            {
                foreach (AeccCorridorSurface s in this._corridor.CorridorSurfaces)
                {
                    IList<Surface> surfaces = new List<Surface>();

                    foreach (AeccCorridorSurfaceMask b in s.Masks)
                    {
                        IList<Point> dSpoints = new List<Point>();

                        foreach (double[] point in b.GetPolygonPoints())
                        {
                            dSpoints.Add(Point.ByCoordinates(point[0], point[1], point[2]));
                        }

                        surfaces.Add(Surface.ByPerimeterPoints(Point.PruneDuplicates(dSpoints)));  // [20181009]
                    }

                    foreach (AeccCorridorSurfaceBoundary b in s.Boundaries)
                    {
                        IList<Point> dSpoints = new List<Point>();

                        foreach (double[] point in b.GetPolygonPoints())
                        {
                            dSpoints.Add(Point.ByCoordinates(point[0], point[1], point[2]));
                        }

                        surfaces.Add(Surface.ByPerimeterPoints(Point.PruneDuplicates(dSpoints)));  // [20181009]
                    }

                    if (surfaces.Count > 0)
                    {
                        output.Add(surfaces);
                    }
                }
            }

            Utils.Log(string.Format("Corridor.GetCorridorSurfaces completed.", ""));

            //TODO raise exception for no corridor surfaces
            return output;
        }

        /// <summary>
        /// Returns a collection of AppliedSubassemblyShapes in the Corridor.
        /// </summary>
        /// <returns></returns>
        private IList<IList<IList<AppliedSubassemblyShape>>> GetShapesFromXML()
        {
            Utils.Log(string.Format("Corridor.GetShapesFromXML started...", ""));

            IList<IList<IList<AppliedSubassemblyShape>>> corridorShapes = new List<IList<IList<AppliedSubassemblyShape>>>();

            string xmlPath = System.IO.Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "CorridorShapes.xml");  // Revit 2020 changed the path to the temp at a session level

            Utils.Log(xmlPath);

            this._document.SendCommand(string.Format("-ExportSubassemblyShapesToXML\n{0}\n{1}\n{2}\n", this._corridor.Handle, -1, -1));

            DateTime start = DateTime.Now;


            while (true)
            {
                if (System.IO.File.Exists(xmlPath))
                {
                    if (System.IO.File.GetLastWriteTime(xmlPath) > start)
                    {
                        start = System.IO.File.GetLastWriteTime(xmlPath);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            Utils.Log("XML acquired.");

            if (System.IO.File.Exists(xmlPath))
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlPath);

                foreach (XmlElement corridor in xmlDoc.GetElementsByTagName("Corridor").Cast<XmlElement>().First(x => x.Attributes["Name"].Value == this.Name))
                {
                    foreach (XmlElement baseline in corridor.GetElementsByTagName("Baseline"))
                    {
                        IList<IList<AppliedSubassemblyShape>> baselineShapes = new List<IList<AppliedSubassemblyShape>>();

                        foreach (XmlElement region in baseline.GetElementsByTagName("Region"))
                        {
                            IList<AppliedSubassemblyShape> regionShapes = new List<AppliedSubassemblyShape>();

                            foreach (XmlElement shape in region.GetElementsByTagName("Shape"))
                            {
                                IList<Point> points = new List<Point>();

                                string corrName = shape.Attributes["Corridor"].Value;
                                string baselineIndex = shape.Attributes["BaselineIndex"].Value;
                                string regionIndex = shape.Attributes["RegionIndex"].Value;
                                string assembly = shape.Attributes["AssemblyName"].Value;
                                string subassembly = shape.Attributes["SubassemblyName"].Value;
                                string handle = shape.Attributes["Handle"].Value;
                                string index = shape.Attributes["ShapeIndex"].Value;
                                double station = Convert.ToDouble(shape.Attributes["Station"].Value, CultureInfo.InvariantCulture);

                                string name = string.Join("_", corrName, baselineIndex, regionIndex, assembly, subassembly, handle, index);

                                foreach (XmlElement p in shape.GetElementsByTagName("Point"))
                                {
                                    double x = Convert.ToDouble(p.Attributes["X"].Value, CultureInfo.InvariantCulture);
                                    double y = Convert.ToDouble(p.Attributes["Y"].Value, CultureInfo.InvariantCulture);
                                    double z = Convert.ToDouble(p.Attributes["Z"].Value, CultureInfo.InvariantCulture);

                                    points.Add(Point.ByCoordinates(x, y, z));
                                }

                                IList<string> codes = new List<string>();

                                // 20201025 - START

                                var xml_codes = shape.GetElementsByTagName("Codes");

                                if (xml_codes != null)
                                {
                                    var xml_code = shape.GetElementsByTagName("Code");

                                    if (xml_code != null)
                                    {
                                        foreach (XmlElement c in xml_code)
                                        {
                                            string code = c.Attributes["Name"].Value;
                                            if (!codes.Contains(code))
                                            {
                                                codes.Add(code);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        codes.Add("_NoCode_");
                                    }
                                }
                                else
                                {
                                    codes.Add("_NoCodes_");
                                }

                                //foreach (XmlElement c in shape.GetElementsByTagName("Code"))
                                //{
                                //    string code = c.Attributes["Name"].Value;
                                //    if (!codes.Contains(code))
                                //    {
                                //        codes.Add(code);
                                //    }
                                //}

                                // 20201025 - END

                                points = Point.PruneDuplicates(points);

                                if (points.Count < 2)
                                {
                                    Utils.Log(string.Format("ERROR: Not enough points to make a closed loop: {0} {1}", name, station));
                                    continue;
                                }

                                PolyCurve pc = PolyCurve.ByPoints(points, true);

                                AppliedSubassemblyShape appSubShape = null;

                                try
                                {
                                    appSubShape = new AppliedSubassemblyShape(name, pc, codes, station);
                                }
                                catch (Exception ex)
                                {
                                    Utils.Log(string.Format("ERROR: {0} {1} {2}", name, station, ex.Message));
                                }

                                if (appSubShape != null)
                                {
                                    regionShapes.Add(appSubShape);
                                }
                            }

                            baselineShapes.Add(regionShapes);
                        }

                        corridorShapes.Add(baselineShapes);
                    }
                }
            }
            else
            {
                Utils.Log("ERROR: Failed to locate CorridorShapes.xml in the Temp folder.");
            }

            Utils.Log(string.Format("Corridor.GetShapesFromXML completed.", ""));

            return corridorShapes;
        }

        /// <summary>
        /// Returns a collection of AppliedSubassemblyLinks in the Corridor.
        /// </summary>
        /// <returns></returns>
        private IList<IList<IList<AppliedSubassemblyLink>>> GetLinksFromXML()
        {
            Utils.Log(string.Format("Corridor.GetLinksFromXML started...", ""));

            IList<IList<IList<AppliedSubassemblyLink>>> corridorLinks = new List<IList<IList<AppliedSubassemblyLink>>>();

            string xmlPath = System.IO.Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "CorridorLinks.xml");  // Revit 2020 changed the path to the temp at a session level

            Utils.Log(xmlPath);

            this._document.SendCommand(string.Format("-ExportSubassemblyLinksToXML\n{0}\n{1}\n{2}\n", this._corridor.Handle, -1, -1));

            DateTime start = DateTime.Now;


            while (true)
            {
                if (System.IO.File.Exists(xmlPath))
                {
                    if (System.IO.File.GetLastWriteTime(xmlPath) > start)
                    {
                        start = System.IO.File.GetLastWriteTime(xmlPath);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            Utils.Log("XML acquired.");

            if (System.IO.File.Exists(xmlPath))
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlPath);

                foreach (XmlElement corridor in xmlDoc.GetElementsByTagName("Corridor").Cast<XmlElement>().First(x => x.Attributes["Name"].Value == this.Name))
                {
                    foreach (XmlElement baseline in corridor.GetElementsByTagName("Baseline"))
                    {
                        IList<IList<AppliedSubassemblyLink>> baselineLinks = new List<IList<AppliedSubassemblyLink>>();

                        foreach (XmlElement region in baseline.GetElementsByTagName("Region"))
                        {
                            IList<AppliedSubassemblyLink> regionLinks = new List<AppliedSubassemblyLink>();

                            foreach (XmlElement link in region.GetElementsByTagName("Link"))
                            {
                                //Utils.Log($"Processing Link...");

                                try
                                {
                                    IList<Point> points = new List<Point>();

                                    string corrName = link.Attributes["Corridor"].Value;
                                    string baselineIndex = link.Attributes["BaselineIndex"].Value;
                                    string regionIndex = link.Attributes["RegionIndex"].Value;
                                    string assembly = link.Attributes["AssemblyName"].Value;
                                    string subassembly = link.Attributes["SubassemblyName"].Value;
                                    string handle = link.Attributes["Handle"].Value;
                                    string index = link.Attributes["LinkIndex"].Value;
                                    double station = Convert.ToDouble(link.Attributes["Station"].Value, CultureInfo.InvariantCulture);

                                    string name = string.Join("_", corrName, baselineIndex, regionIndex, assembly, subassembly, handle, index);

                                    //Utils.Log($"Name: {name}");

                                    foreach (XmlElement p in link.GetElementsByTagName("Point"))
                                    {
                                        double x = Convert.ToDouble(p.Attributes["X"].Value, CultureInfo.InvariantCulture);
                                        double y = Convert.ToDouble(p.Attributes["Y"].Value, CultureInfo.InvariantCulture);
                                        double z = Convert.ToDouble(p.Attributes["Z"].Value, CultureInfo.InvariantCulture);

                                        points.Add(Point.ByCoordinates(x, y, z));
                                    }

                                    points = Point.PruneDuplicates(points);

                                    //Utils.Log($"Points: {points.Count}");

                                    IList<string> codes = new List<string>();

                                    // 20201025 - START

                                    var xml_codes = link.GetElementsByTagName("Codes");

                                    if (xml_codes != null)
                                    {
                                        var xml_code = link.GetElementsByTagName("Code");

                                        if (xml_code != null)
                                        {
                                            foreach (XmlElement c in xml_code)
                                            {
                                                string code = c.Attributes["Name"].Value;
                                                if (!codes.Contains(code))
                                                {
                                                    codes.Add(code);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            codes.Add("_NoCode_");
                                        }
                                    }
                                    else
                                    {
                                        codes.Add("_NoCodes_");
                                    }

                                    //Utils.Log($"Codes: {codes}");

                                    //foreach (XmlElement c in link.GetElementsByTagName("Code"))
                                    //{
                                    //    string code = c.Attributes["Name"].Value;
                                    //    if (!codes.Contains(code))
                                    //    {
                                    //        codes.Add(code);
                                    //    }
                                    //}

                                    // 20201025 - END


                                    if (points.Count > 1)
                                    {
                                        PolyCurve pc = PolyCurve.ByPoints(points);

                                        AppliedSubassemblyLink appSubLink = null;

                                        try
                                        {
                                            appSubLink = new AppliedSubassemblyLink(name, pc, codes, station);
                                        }
                                        catch (Exception ex)
                                        {
                                            Utils.Log(string.Format("ERROR: {0} {1} {2}", name, station, ex.Message));
                                        }

                                        if (appSubLink != null)
                                        {
                                            regionLinks.Add(appSubLink);
                                        }
                                        else
                                        {
                                            Utils.Log(string.Format("ERROR: The AppliedSubassemblyLink is null, Station: {0}", station));
                                        }
                                    }
                                    else
                                    {
                                        Utils.Log(string.Format("ERROR: Not enough points to make a link: {0} {1}", name, station));
                                    }

                                    //Utils.Log("Completed");
                                }
                                catch (Exception ex)
                                {
                                    Utils.Log(ex);
                                }
                            }

                            baselineLinks.Add(regionLinks);
                        }

                        corridorLinks.Add(baselineLinks);
                    }
                }
            }
            else
            {
                Utils.Log("ERROR: Failed to locate CorridorLinks.xml in the Temp folder.");
            }

            Utils.Log(string.Format("Corridor.GetLinksFromXML completed.", ""));

            return corridorLinks;
        }
        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Returns a Point by station offset elevation.
        /// </summary>
        /// <param name="baseline">The baseline.</param>
        /// <param name="station">The station.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="elevation">The elevation.</param>
        /// <returns></returns>
        public Point PointByStationOffsetElevation(Baseline baseline, double station = 0, double offset = 0, double elevation = 0)
        {
            return baseline.PointByStationOffsetElevation(station, offset, elevation);
        }

        /// <summary>
        /// Returns a CoordinateSystem by station.
        /// </summary>
        /// <param name="baseline">The baseline.</param>
        /// <param name="station">The station.</param>
        /// <returns></returns>
        public CoordinateSystem CoordinateSystemByStation(Baseline baseline, double station = 0)
        {
            return baseline.CoordinateSystemByStation(station);
        }

        /// <summary>
        /// Returns a CoordinateSystem by point.
        /// </summary>
        /// <param name="baseline">The baseline.</param>
        /// <param name="point">The point.</param>
        /// <returns></returns>
        public CoordinateSystem CoordinateSystemByPoint(Baseline baseline, Point point)
        {
            return baseline.GetCoordinateSystemByPoint(point);
        }

        /// <summary>
        /// Gets the PointCodes.
        /// </summary>
        /// <returns></returns>
        public IList<string> GetCodes()
        {
            Utils.Log(string.Format("Corridor.GetCodes started...", ""));

            IList<string> output = new List<string>();

            IList<AeccCorridorCodes> codeList = new List<AeccCorridorCodes>();

            try
            {
                foreach (AeccBaseline b in this._corridor.Baselines)
                {
                    foreach (string code in b.MainBaselineFeatureLines.CodeNames)
                    {
                        if (!output.Contains(code))
                        {
                            output.Add(code);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: {0}", ex.Message));
            }

            Utils.Log(string.Format("Corridor.GetCodes completed.", ""));

            return output.OrderBy(x => x).ToList();
        }

        /// <summary>
        /// Gets the corridor Featurelies organized by Corridor-Baseline-Code-Region
        /// </summary>
        /// <returns></returns>
        public IList<IList<IList<Featureline>>> GetFeaturelines()
        {
            return Utils.GetFeaturelines(this);
        }

        /// <summary>
        /// Gets the subassembly points organized by: Corridor &gt; Baseline &gt; Region &gt; Assembly &gt; Subassembly.
        /// </summary>
        /// <param name="dumpXML">If true dumps a LandXML in the Temp folder.</param>
        /// <returns></returns>
        public IList<IList<IList<IList<IList<Point>>>>> GetSubassemblyPoints(bool dumpXML = false)
        {
            return Utils.GetCorridorSubAssembliesFromLandXML(this, dumpXML);
        }

        /// <summary>
        ///  Gets the subassembly points organized by: Corridor &gt; Baseline &gt; Region &gt; Code.
        ///  It requires to export a LandXML to the %Temp% folder, named like the Civil 3D Document, containing only the corridor data.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <returns></returns>
        public IList<IList<IList<IList<IList<Point>>>>> GetPointsByCode(string code)
        {
            return Utils.GetCorridorPointsByCodeFromLandXML(this, code);
        }


        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("Corridor(Name = {0})", this.Name);
        }

        /// <summary>
        /// Gets the closest featureline by point code side.
        /// </summary>
        /// <param name="point">The point.</param>
        /// <param name="baselineIndex">Index of the baseline.</param>
        /// <param name="code">The code.</param>
        /// <param name="side">The side.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "Featureline" })]
        public Dictionary<string, object> GetFeaturelineByPointCodeSide(Point point, int baselineIndex, string code, string side)
        {
            return new Dictionary<string, object>() { { "Featureline", UtilsObjectsLocation.ClosestFeaturelineByPoint(point, this, baselineIndex, code, side) } };
        }

        /// <summary>
        /// Gets the featurelines by Code &gt; Baseline &gt; Region.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <returns></returns>
        public IList<IList<IList<Featureline>>> GetFeaturelinesByCode(string code)  // 1.1.0
        {
            Utils.Log(string.Format("Corridor.GetFeaturelinesByCode started...", ""));

            IList<IList<IList<Featureline>>> corridorFeaturelines = new List<IList<IList<Featureline>>>();

            foreach (Baseline bl in this.Baselines)
            {
                IList<IList<Featureline>> blFeaturelines = bl.GetFeaturelinesByCode(code);

                corridorFeaturelines.Add(blFeaturelines);
            }

            Utils.Log(string.Format("Corridor.GetFeaturelinesByCode completed.", ""));

            return corridorFeaturelines;
        }

        /// <summary>
        /// Gets the featurelines by Code &gt; Baseline &gt; Region.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <param name="station">The station.</param>
        /// <returns></returns>
        public IList<IList<Featureline>> GetFeaturelinesByCodeStation(string code, double station)  // 1.1.0
        {
            Utils.Log(string.Format("Corridor.GetFeaturelinesByCodeStation started...", ""));

            IList<IList<Featureline>> corridorFeaturelines = new List<IList<Featureline>>();

            foreach (Baseline bl in this.Baselines)
            {
                IList<Featureline> blFeaturelines = bl.GetFeaturelinesByCodeStation(code, station);

                corridorFeaturelines.Add(blFeaturelines);
            }

            Utils.Log(string.Format("Corridor.GetFeaturelinesByCodeStation completed.", ""));

            return corridorFeaturelines;
        }

        #endregion
    }
}
