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
using Autodesk.AECC.Interop.UiLand;
using Autodesk.AECC.Interop.Land;
using System.Reflection;

using Autodesk.DesignScript.Runtime;

using Autodesk.DesignScript.Geometry;

using System.Xml;
using System.Globalization;

namespace CivilConnection
{
    /// <summary>
    /// The CivilDocument class
    /// </summary>
    public class CivilDocument
    {
        #region PRIVATE PROPERTIES
        /// <summary>
        /// The document
        /// </summary>
        internal AeccRoadwayDocument _document;
        /// <summary>
        /// The document name.
        /// </summary>
        public string Name { get { return _document.Name; } }
        /// <summary>
        /// The corridors
        /// </summary>
        private AeccCorridors _corridors;
        /// <summary>
        /// The alignments
        /// </summary>
        private AeccAlignmentsSiteless _alignments;
        /// <summary>
        /// The Surfaces
        /// </summary>
        private AeccSurfaces _surfaces;
        /// <summary>
        /// Gets the internal element.
        /// </summary>
        /// <value>
        /// The internal element.
        /// </value>
        internal object InternalElement { get { return this._document; } }
        #endregion

        #region CONSTRUCTORS
        /// <summary>
        /// Initializes a new instance of the <see cref="CivilDocument"/> class.
        /// </summary>
        /// <param name="_doc">The document.</param>
        internal CivilDocument(AeccRoadwayDocument _doc)
        {
            this._document = _doc;
            _corridors = _doc.Corridors;
            _alignments = _doc.AlignmentsSiteless;
            _surfaces = _doc.Surfaces;
        }
        #endregion

        #region PRIVATE METHODS
        /// <summary>
        /// Creates a LandXML from the Civil Document
        /// </summary>
        /// <returns></returns>
        private string DumpLandXML()
        {
            return Utils.DumpLandXML(this._document);
        }

        /// <summary>
        /// Gets the land featurelines.
        /// </summary>
        /// <param name="xmlPath">The XML path for the LandFeaturerline properties.</param>
        /// <returns></returns>
        /// 
        [IsVisibleInDynamoLibraryAttribute(false)]
        public IList<LandFeatureline> GetLandFeaturelines(string xmlPath = "")
        {
            Utils.Log(string.Format("CivilDocument.GetLandFeaturelines started...", ""));

            if (string.IsNullOrEmpty(xmlPath))
            {
                xmlPath = System.IO.Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "LandFeatureLinesReport.xml");
            }

            this.SendCommand("-ExportLandFeatureLinesToXml\n");

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

            bool result = true;

            IList<LandFeatureline> output = new List<LandFeatureline>();

            if (result)
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlPath);

                AcadDatabase db = this._document as AcadDatabase;
                AcadModelSpace ms = db.ModelSpace;

                foreach (AcadEntity e in ms)
                {
                    if (e.EntityName.ToLower().Contains("featureline"))
                    {
                        AeccLandFeatureLine f = e as AeccLandFeatureLine;

                        XmlElement fe = xmlDoc.GetElementsByTagName("FeatureLine").Cast<XmlElement>().First(x => x.Attributes["Handle"].Value == f.Handle.ToString());

                        IList<Point> points = new List<Point>();

                        foreach (XmlElement p in fe.GetElementsByTagName("Point"))
                        {
                            double x = Convert.ToDouble(p.Attributes["X"].Value, CultureInfo.InvariantCulture);
                            double y = Convert.ToDouble(p.Attributes["Y"].Value, CultureInfo.InvariantCulture);
                            double z = Convert.ToDouble(p.Attributes["Z"].Value, CultureInfo.InvariantCulture);

                            points.Add(Point.ByCoordinates(x, y, z));
                        }

                        points = Point.PruneDuplicates(points);

                        if (points.Count > 1)
                        {
                            PolyCurve pc = PolyCurve.ByPoints(points);
                            string style = fe.Attributes["Style"].Value;

                            output.Add(new LandFeatureline(f, pc, style));
                        }

                        foreach (var item in points)
                        {
                            if (item != null)
                            {
                                item.Dispose();
                            }
                        }
                    }
                }

                Utils.Log(string.Format("CivilDocument.GetLandFeaturelines completed.", ""));
            }

            return output;
        }
        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// Gets the corridors.
        /// </summary>
        /// <returns></returns>
        public IList<Corridor> GetCorridors()
        {
            Utils.Log(string.Format("CivilDocument.GetCorridors started...", ""));

            IList<Corridor> output = new List<Corridor>();

            foreach (AeccCorridor corridor in this._document.Corridors)
            {
                output.Add(new Corridor(corridor, this._document));
            }

            Utils.Log(string.Format("CivilDocument.GetCorridors completed.", ""));

            return output;
        }

        /// <summary>
        /// Get the corridor by name.
        /// </summary>
        /// <param name="name">The corridor name.</param>
        /// <returns></returns>
        public Corridor GetCorridorByName(string name)
        {
            return this.GetCorridors().First(x => x.Name == name);
        }

        /// <summary>
        /// Gets the alignments.
        /// </summary>
        /// <returns></returns>
        public IList<Alignment> GetAlignments()
        {
            Utils.Log(string.Format("CivilDocument.GetAlignments started...", ""));

            IList<Alignment> output = new List<Alignment>();

            foreach (AeccAlignment a in this._alignments)
            {
                output.Add(new Alignment(a));
            }

            foreach (AeccSite site in this._document.Sites)
            {
                foreach (AeccAlignment a in site.Alignments)
                {
                    output.Add(new Alignment(a));
                }
            }

            Utils.Log(string.Format("CivilDocument.GetAlignments completed.", ""));

            return output;
        }

        /// <summary>
        /// Gets alignment by name.
        /// </summary>
        /// <param name="name">The alignment name.</param>
        /// <returns></returns>
        public Alignment GetAlignmentByName(string name)
        {
            return this.GetAlignments().First(x => x.Name == name);
        }

        /// <summary>
        /// Gets all surfaces in the document
        /// </summary>
        /// <returns>
        /// List of surfaces
        /// </returns>
        public IList<CivilSurface> GetSurfaces()
        {
            Utils.Log(string.Format("CivilDocument.GetSurfaces started...", ""));

            IList<CivilSurface> output = new List<CivilSurface>();

            foreach (AeccSurface s in this._surfaces)
            {
                output.Add(new CivilSurface(s));
            }

            Utils.Log(string.Format("CivilDocument.GetSurfaces completed.", ""));

            return output;
        }

        /// <summary>
        /// Gets surface by name.
        /// </summary>
        /// <param name="name">The name of the surface</param>
        /// <returns>
        /// Civil Surface
        /// </returns>
        public CivilSurface GetSurfaceByName(string name)
        {
            return this.GetSurfaces().First(x => x.Name == name);
        }

        #region AUTOCAD METHODS
        /// <summary>
        /// Adds the arc to the document.
        /// </summary>
        /// <param name="arc">The arc.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public string AddArc(Arc arc, string layer)
        {
            return Utils.AddArcByArc(this._document, arc, layer);
        }

        /// <summary>
        /// Adds the circle to the document.
        /// </summary>
        /// <param name="circle">The circle.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public string AddCircle(Circle circle, string layer)
        {
            return Utils.AddCircleByCircle(this._document, circle, layer);
        }

        /// <summary>
        /// Adds the 2D polyline by points.
        /// </summary>
        /// <param name="points">The points.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public string AddLWPolylineByPoints(IList<Point> points, string layer)
        {
            return Utils.AddLWPolylineByPoints(this._document, points, layer);
        }

        /// <summary>
        /// Adds the 3D polyline by curve.
        /// </summary>
        /// <param name="curve">The curve.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public string AddPolylineByCurve(Curve curve, string layer)
        {
            return Utils.AddPolylineByCurve(this._document, curve, layer);
        }

        /// <summary>
        /// Adds the 3D polylines by curve.
        /// </summary>
        /// <param name="curves">The curves.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public string AddPolylineByCurve(IList<Curve> curves, string layer)
        {
            return Utils.AddPolylineByCurves(this._document, curves, layer);
        }

        /// <summary>
        /// Adds the region by closed profile.
        /// </summary>
        /// <param name="closedCurve">The closed curve.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public string AddRegionByPatch(Curve closedCurve, string layer)
        {
            return Utils.AddRegionByPatch(this._document, closedCurve, layer);
        }

        /// <summary>
        /// Creates a closed profile form the points and adds the extruded solid.
        /// </summary>
        /// <param name="points">The points.</param>
        /// <param name="height">The height. By Default is equal to 1.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public string AddExtrudedSolidByPoints(IList<Point> points, double height = 1, string layer = "_CivilConnectionSolid")
        {
            return Utils.AddExtrudedSolidByPoints(this._document, points, height, layer);
        }

        /// <summary>
        /// Adds the extruded solid by closed profile.
        /// </summary>
        /// <param name="closedCurve">The closed curve.</param>
        /// <param name="height">The height. By Default is equal to 1.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public string AddExtrudedSolidByPatch(Curve closedCurve, double height = 1, string layer = "_CivilConnectionSolid")
        {
            return Utils.AddExtrudedSolidByPatch(this._document, closedCurve, height, layer);
        }

        /// <summary>
        /// Adds multiple extruded solid by closed curves.
        /// </summary>
        /// <param name="curves">The curves.</param>
        /// <param name="height">The height. By Default is equal to 1.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public string AddExtrudedSolidByCurves(IList<Curve> curves, double height = 1, string layer = "_CivilConnectionSolid")
        {
            return Utils.AddExtrudedSolidByCurves(this._document, curves, height, layer);
        }

        /// <summary>
        /// Adds a new layer to the Civil Document by name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        public string AddLayer(string name)
        {
            Utils.AddLayer(this._document, name);
            return name;
        }

        /// <summary>
        /// Creates a text in the CivilDocument and rotates it to match the plane.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="point">The point.</param>
        /// <param name="textHeight">Height of the text.</param>
        /// <param name="layer">The layer.</param>
        /// <param name="cs">The cs.</param>
        /// <returns></returns>
        public string AddText(string text, Point point, double textHeight, string layer, CoordinateSystem cs)
        {
            return Utils.AddText(this._document, text, point, textHeight, layer, cs);
        }

        /// <summary>
        /// Creates an extrusion based on a closed curve (Polycurve, Polygon or Rectangle) profile to be used to cut the solids in the Civil Document.
        /// </summary>
        /// <param name="closedCurve">The closed curve.</param>
        /// <param name="height">The height. By Default is equal to 1.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public bool CutSolidsByPatch(Curve closedCurve, double height = 1, string layer = "_CivilConnectionSolid")
        {
            return Utils.CutSolidsByPatch(this._document, closedCurve, height, layer);
        }

        /// <summary>
        /// Creates an extrusion based on a profile defined by a set of curves profile to be used to cut the solids in the Civil Document.
        /// </summary>
        /// <param name="closedCurves">The closed curves.</param>
        /// <param name="height">The height. By Default is equal to 1.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public bool CutSolidsByCurves(IList<Curve> closedCurves, double height = 1, string layer = "_CivilConnectionSolid")
        {
            return Utils.CutSolidsByCurves(this._document, closedCurves, height, layer);
        }

        /// <summary>
        /// Creates a solid to be used to cut the solids in the Civil 3D Document.
        /// </summary>
        /// <param name="geometry">The solid geometry.</param>
        /// <param name="layer">The layer where to crerate the cutting solid.</param>
        /// <returns></returns>
        public bool CutSolidsByGeometry(Geometry[] geometry, string layer)
        {
            return Utils.CutSolidsByGeometry(this._document, geometry, layer);
        }

        /// <summary>
        /// Import the geometry from Dynamo into the Civil 3D Document.
        /// </summary>
        /// <param name="geometry">The geometry.</param>
        /// <param name="layer">The layer where to crerate the solid.</param>
        /// <returns></returns>
        public IList<string> ImportGeometry(Geometry[] geometry, string layer)
        {
            return Utils.ImportGeometry(this._document, geometry, layer);
        }

        /// <summary>
        /// Links the geometry associated to a Revit object into Civil 3D.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="parameter">The parameter.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        public string LinkElement(Revit.Elements.Element element, string parameter, string layer)
        {
            return Utils.ImportElement(this._document, element, parameter, layer);
        }

        /// <summary>
        /// Send Command to the Civil Document.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <returns></returns>
        public bool SendCommand(string command)
        {
            Utils.Log(string.Format("CivilDocument.SendCommand started...", ""));

            bool output = true;

            try
            {
                AcadDocument doc = this.InternalElement as AcadDocument;
                doc.SendCommand(command);
            }
            catch (Exception ex)
            {
                output = false;

                Utils.Log(string.Format("ERROR: {0}", ex.Message));

                throw ex;
            }

            Utils.Log(string.Format("CivilDocument.SendCommand completed.", ""));

            return output;
        }

        /// <summary>
        /// Slices the solids in Civil 3D using a Dynamo Plane.
        /// </summary>
        /// <param name="plane">The plane.</param>
        /// <returns></returns>
        public bool SliceSolidsByPlane(Plane plane)
        {
            return Utils.SliceSolidsByPlane(this._document, plane);
        }

        #endregion

        #region CIVIL 3D METHODS

        /// <summary>
        /// Adds the civil point.
        /// </summary>
        /// <param name="point">The point.</param>
        /// <returns></returns>
        public string AddCivilPoint(Point point)
        {
            return Utils.AddCivilPointByPoint(this._document, point);
        }

        /// <summary>
        /// Adds the civil point group.
        /// </summary>
        /// <param name="points">The points.</param>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        public string AddCivilPointGroup(Point[] points, string name)
        {
            return Utils.AddPointGroupByPoint(this._document, points, name);
        }

        /// <summary>
        /// Gets the Civil point groups.
        /// </summary>
        /// <returns></returns>
        [MultiReturn(new string[] { "PointGroupNames", "Points" })]
        public Dictionary<string, object> GetPointGroups()
        {
            var dict = Utils.GetPointGroups(this._document);

            return new Dictionary<string, object>() { { "PointGroupNames", dict.Keys }, { "Points", dict.Values } };
        }


        /// <summary>
        /// Adds the tin surface by points.
        /// </summary>
        /// <param name="points">The points.</param>
        /// <param name="name">The name.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        //[IsVisibleInDynamoLibrary(false)]
        public string AddTINSurfaceByPoints(Point[] points, string name, string layer)
        {
            return Utils.AddTINSurfaceByPoints(this._document, points, name, layer);
        }

        #endregion

        /// <summary>
        /// Public textual representation of the Dynamo node preview
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("CivilDocument(Name = {0})", this.Name);
        }
        #endregion
    }
}
