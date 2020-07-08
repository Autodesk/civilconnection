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

using System.IO;
using System.Xml;
using System.Windows.Forms;

using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AECC.Interop.UiRoadway;
using Autodesk.AECC.Interop.Roadway;
using Autodesk.AECC.Interop.Land;
using Autodesk.AECC.Interop.UiLand;
using System.Reflection;

using Autodesk.DesignScript.Runtime;
using Autodesk.DesignScript.Geometry;

namespace CivilConnection
{
    /// <summary>
    /// Collection of utilities.
    /// </summary>
    [SupressImportIntoVM()]
    internal class Utils
    {
        #region PRIVATE PROPERTIES


        #endregion

        #region PUBLIC PROPERTIES


        #endregion

        #region CONSTRUCTOR


        #endregion

        #region PRIVATE METHODS


        #endregion

        #region PUBLIC METHODS


        /// <summary>
        /// Checks if two values are almost equal
        /// </summary>
        /// <param name="a">The first value.</param>
        /// <param name="b">The second value.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static bool AlmostEqual(double a, double b)
        {
            return Math.Abs(a - b) <= 0.00001;
        }


        /// <summary>
        /// Feets to mm.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static double FeetToMm(double d)
        {
            return d * 304.8;
        }


        /// <summary>
        /// Mms to feet.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static double MmToFeet(double d)
        {
            return d / 304.8;
        }


        /// <summary>
        /// Feets to m.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static double FeetToM(double d)
        {
            return d * 0.3048;
        }


        /// <summary>
        /// ms to feet.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static double MToFeet(double d)
        {
            return d / 0.3048;
        }


        /// <summary>
        /// Degs to RAD.
        /// </summary>
        /// <param name="angle">The angle.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static double DegToRad(double angle)
        {
            return angle / 180 * Math.PI;
        }


        /// <summary>
        /// RADs to deg.
        /// </summary>
        /// <param name="d">The d.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static double RadToDeg(double d)
        {
            return d * 180 / Math.PI;
        }


        /// <summary>
        /// Adds the layer.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="layerName">Name of the layer.</param>
        [IsVisibleInDynamoLibrary(false)]
        public static void AddLayer(AeccRoadwayDocument doc, string layerName)
        {
            Utils.Log(string.Format("Utils.AddLayer {0} started...", layerName));

            AcadDatabase db = doc as AcadDatabase;

            bool found = false;

            foreach (AcadLayer l in db.Layers)
            {
                if (l.Name == layerName)
                {
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                db.Layers.Add(layerName);
            }

            Utils.Log(string.Format("Utils.AddLayer completed.", ""));
        }


        /// <summary>
        /// Freezes the layers.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="layer">the name of the layer.</param>
        [IsVisibleInDynamoLibrary(false)]
        public static void FreezeLayers(AeccRoadwayDocument doc, string layer)
        {
            Utils.Log(string.Format("Utils.FreezeLayers started...", ""));

            AcadDatabase db = doc as AcadDatabase;
            IList<AcadLayer> dynLayers = db.Layers.Cast<AcadLayer>().Where(l => l.Name.Equals(layer)).ToList();

            foreach (AcadLayer l in dynLayers)
            {
                l.Freeze = true;
            }

            Utils.Log(string.Format("Utils.FreezeLayers completed.", ""));
        }


        /// <summary>
        /// Adds the text.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="text">The text.</param>
        /// <param name="point">The point.</param>
        /// <param name="height">The height.</param>
        /// <param name="layer">The layer.</param>
        /// <param name="cs">The cs.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddText(AeccRoadwayDocument doc, string text, Point point, double height, string layer, CoordinateSystem cs)
        {
            Utils.Log(string.Format("Utils.AddText started...", ""));

            // TODO: different orientation of curves from Dynamo to AutoCAD
            AddLayer(doc, layer);

            AcadDatabase db = doc as AcadDatabase;
            AcadModelSpace ms = db.ModelSpace;
            var vlist = new double[] { point.X, point.Y, point.Z };
            AcadText a = ms.AddText(text, vlist, height);
            a.Layer = layer;

            double rotationZ = 0;
            double rotationX = 0;

            Vector prX = Vector.ByCoordinates(cs.XAxis.X, cs.XAxis.Y, 0).Normalized();
            Vector normalZ = Vector.ZAxis().Cross(cs.ZAxis).Normalized();

            rotationZ = Vector.ZAxis().AngleAboutAxis(cs.ZAxis, normalZ);
            rotationX = Vector.XAxis().AngleAboutAxis(prX, Vector.ZAxis());

            if (rotationX != 0)
            {
                Utils.Log(string.Format("Rotation X: {0}", rotationX));
                var p1 = new double[] { point.X, point.Y, point.Z + 1 };
                a.Rotate3D(vlist, p1, DegToRad(rotationX));
            }

            if (rotationZ != 0)
            {
                Utils.Log(string.Format("Rotation Z: {0}", rotationZ));
                var p1 = new double[] { point.X + normalZ.X, point.Y + normalZ.Y, point.Z + normalZ.Z };
                a.Rotate3D(vlist, p1, DegToRad(rotationZ));
            }

            prX.Dispose();
            normalZ.Dispose();

            Utils.Log(string.Format("Utils.AddText completed.", ""));

            return text;
        }


        /// <summary>
        /// Adds the arc by arc.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="arc">The arc.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddArcByArc(AeccRoadwayDocument doc, Arc arc, string layer)
        {
            Utils.Log(string.Format("Utils.AddArcByArc started...", ""));

            // Arcs in AutoCAD are created horizontal
            // Create the arc and then rotate in 3D to match Dynamo input

            Point center = arc.CenterPoint;
            Plane curvePlane = Plane.ByOriginNormal(center, arc.Normal);

            // (1) Create horizontal Dynamo Arc from Arc input
            CoordinateSystem curveCSInverse = curvePlane.ToCoordinateSystem().Inverse();

            Arc ha = arc.Transform(curveCSInverse) as Arc;  // this arc is the transformed copy of the input in the origin

            // Do not trust the StartAngle and SweepAngle properties..
            Vector cs = Vector.ByTwoPoints(ha.CenterPoint, ha.StartPoint);
            Vector ce = Vector.ByTwoPoints(ha.CenterPoint, ha.EndPoint);
            double start = Vector.XAxis().AngleAboutAxis(cs, Vector.ZAxis());
            double end = Vector.XAxis().AngleAboutAxis(ce, Vector.ZAxis());
            double radius = arc.Radius;

            // (2) Create the Arc in AutoCAD
            AddLayer(doc, layer);
            AcadDatabase db = doc as AcadDatabase;
            AcadModelSpace ms = db.ModelSpace;
            var vlist = new double[] { ha.CenterPoint.X, ha.CenterPoint.Y, ha.CenterPoint.Z };
            AcadArc a = ms.AddArc(vlist, radius, DegToRad(start), DegToRad(end));
            a.Layer = layer;

            curveCSInverse = curveCSInverse.Inverse();

            a.TransformBy(new double[,] 
            {
                {curveCSInverse.XAxis.X / curveCSInverse.XScaleFactor, curveCSInverse.YAxis.X / curveCSInverse.XScaleFactor, curveCSInverse.ZAxis.X / curveCSInverse.XScaleFactor, curveCSInverse.Origin.X},
                {curveCSInverse.XAxis.Y / curveCSInverse.YScaleFactor, curveCSInverse.YAxis.Y / curveCSInverse.YScaleFactor, curveCSInverse.ZAxis.Y / curveCSInverse.YScaleFactor, curveCSInverse.Origin.Y},
                {curveCSInverse.XAxis.Z / curveCSInverse.ZScaleFactor, curveCSInverse.YAxis.Z / curveCSInverse.ZScaleFactor, curveCSInverse.ZAxis.Z / curveCSInverse.ZScaleFactor, curveCSInverse.Origin.Z},
                {0, 0, 0, 1}
            });

            // Dispose Dynamo geometry objects
            center.Dispose();
            curvePlane.Dispose();
            curveCSInverse.Dispose();
            ha.Dispose();
            cs.Dispose();
            ce.Dispose();

            Utils.Log(string.Format("Utils.AddArcByArc completed.", ""));

            return a.Handle;
        }

        /// <summary>
        /// Adds the point to the document.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="point">The point.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddPointByPoint(AeccRoadwayDocument doc, Point point, string layer)
        {
            Utils.Log(string.Format("Utils.AddPointByPoint started...", ""));

            AddLayer(doc, layer);

            AcadDatabase db = doc as AcadDatabase;
            AcadModelSpace ms = db.ModelSpace;
            var coordinates = new double[] { point.X, point.Y, point.Z };
            var p = ms.AddPoint(coordinates);
            p.Layer = layer;

            Utils.Log(string.Format("Utils.AddPointByPoint completed.", ""));

            return p.Handle;
        }


        /// <summary>
        /// Adds the land point by point.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="point">The point.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddCivilPointByPoint(AeccRoadwayDocument doc, Point point)
        {
            Utils.Log(string.Format("Utils.AddCivilPointByPoint started...", ""));

            var points = doc.Points;
            var coordinates = new double[] { point.X, point.Y, point.Z };
            var p = points.Add(coordinates);

            Utils.Log(string.Format("Utils.AddCivilPointByPoint completed.", ""));

            return p.Handle;
        }


        /// <summary>
        /// Adds the point group by point.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="points">The points.</param>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddPointGroupByPoint(AeccRoadwayDocument doc, Point[] points, string name)
        {
            Utils.Log(string.Format("Utils.AddPointGroupByPoint started...", ""));

            AeccPointGroups groups = null;
            AeccPointGroup group = null;

            groups = doc.PointGroups;

            if (groups != null)
            {
                if (groups.Count > 0)
                {
                    foreach (AeccPointGroup g in groups)
                    {
                        if (g.Name == name)
                        {
                            group = g;
                            break;
                        }
                    }
                }

                if (group == null)
                {
                    group = groups.Add(name);
                }

                var docPoints = doc.Points;

                IList<string> numbers = new List<string>();

                for (int i = 0; i < points.Length; ++i)
                {
                    var coordinates = new double[] { points[i].X, points[i].Y, points[i].Z };
                    var p = docPoints.Add(coordinates);

                    numbers.Add(p.Number.ToString());
                }

                string formula = group.QueryBuilder.IncludeNumbers;

                if ("" == formula)
                {
                    formula = numbers[0] + "-" + numbers[numbers.Count - 1];
                }
                else
                {
                    formula += ", " + numbers[0] + "-" + numbers[numbers.Count - 1];
                }

                group.QueryBuilder.IncludeNumbers = formula;

                Utils.Log(string.Format("Utils.AddPointGroupByPoint completed.", ""));

                return group.Handle;
            }

            Utils.Log(string.Format("Utils.AddPointGroupByPoint completed.", ""));

            return "";
        }


        /// <summary>
        /// Returns the point groups in Civil 3D as Dynamo point lists.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static Dictionary<string, IList<Point>> GetPointGroups(AeccRoadwayDocument doc)
        {
            Utils.Log(string.Format("Utils.GetPointGroups started...", ""));

            AeccPointGroups groups = null;

            Dictionary<string, IList<Point>> output = new Dictionary<string, IList<Point>>();

            groups = doc.PointGroups;

            if (groups != null)
            {
                Utils.Log("Processing Point Groups...");

                if (groups.Count > 0)
                {
                    foreach (AeccPointGroup g in groups)
                    {
                        Utils.Log(string.Format("Processing Point Group {0}...", g.Name));

                        IList<Point> group = new List<Point>();

                        foreach (int i in g.Points)
                        {
                            AeccPoint p = doc.Points.Item(i - 1);

                            Utils.Log(string.Format("Processing Point {0}...", i));

                            Point pt = Point.ByCoordinates(p.Easting, p.Northing, p.Elevation);

                            Utils.Log(string.Format("{0} acquired.", pt));

                            group.Add(pt);
                        }

                        if (group.Count > 0)
                        {
                            output.Add(g.Name, group);

                            Utils.Log(string.Format("Processing Point Group {0} completed.", g.Name));
                        }
                    }
                }
            }

            Utils.Log(string.Format("Utils.GetPointGroups completed.", ""));

            return output;
        }

        // Atul Tegar  -- 20190917
        private class TinCreationData : AeccTinCreationData
        {
            public string BaseLayer { get; set; }
            public string Description { get; set; }
            public string Layer { get; set; }
            public string Name { get; set; }
            public string Style { get; set; }
        }

        /// <summary>
        /// Adds a TIN surface by points.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="points">The points.</param>
        /// <param name="name">The name.</param>
        /// <param name="layer">The name of the layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(true)]
        public static string AddTINSurfaceByPoints(AeccRoadwayDocument doc, Point[] points, string name, string layer)
        {
            Utils.Log(string.Format("Utils.AddTINSurfaceByPoints started...", ""));

            List<Point> pts = new List<Point>();

            foreach (Point p in points)
            {
                if (double.IsInfinity(p.X) || double.IsInfinity(p.Y) || double.IsInfinity(p.Z) ||
                    double.MaxValue < p.X || double.MaxValue < p.Y || double.MaxValue < p.Z ||
                    double.MinValue > p.X || double.MinValue > p.Y || double.MinValue > p.Z ||
                    double.NaN == p.X || double.NaN == p.Y || double.NaN == p.Z)
                {
                    Utils.Log(string.Format("Discarded {0} ", p));
                    continue;
                }

                pts.Add(p);
            }

            string handle = "";

            try
            {
                AddLayer(doc, layer);

                AeccPointGroup group = doc.HandleToObject(AddPointGroupByPoint(doc, pts.ToArray(), name)) as AeccPointGroup;

                AeccSurfaces surfaces = doc.Surfaces;

                Type surfacesType = surfaces.GetType();

                if (surfacesType != null)
                {
                    TinCreationData data = new TinCreationData()
                    {
                        BaseLayer = layer,
                        Description = "Created by Autodesk CivilConnection",
                        Layer = layer,
                        Name = name,
                        Style = doc.SurfaceStyles.Cast<AeccSurfaceStyle>().First().Name
                    };

                    AeccTinSurface surface = surfaces.AddTinSurface(data);

                    surface.PointGroups.Add(group);

                    handle = surface.Handle;
                }
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: Utils.AddTINSurfaceByPoints {0}", ex.Message));

                throw ex;
            }

            Utils.Log(string.Format("Utils.AddTINSurfaceByPoints completed.", ""));

            return handle;
        }


        /// <summary>
        /// Adds a polyline by points.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="points">The points.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddPolylineByPoints(AeccRoadwayDocument doc, IList<Point> points, string layer)
        {
            Utils.Log(string.Format("Utils.AddPolylineByPoints started...", ""));

            //string layer = "DYN-Shapes";
            AddLayer(doc, layer);

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            double[] vlist = new double[3 * points.Count];

            for (int i = 0; i < points.Count; ++i)
            {
                vlist[3 * i] = points[i].X;
                vlist[3 * i + 1] = points[i].Y;
                vlist[3 * i + 2] = points[i].Z;
            }

            var pl = ms.Add3DPoly(vlist);
            pl.Layer = layer;
            pl.Closed = true;

            Utils.Log(string.Format("Utils.AddPolylineByPoints completed.", ""));

            return pl.Handle;
        }


        /// <summary>
        /// Adds a circle entity in Civil 3D by circle.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="c">The c.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddCircleByCircle(AeccRoadwayDocument doc, Circle c, string layer)
        {
            Utils.Log(string.Format("Utils.AddCircleByCircle started...", ""));

            AddLayer(doc, layer);

            Point center = c.CenterPoint;

            double radius = c.Radius;

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            double[] vlist = new double[] { center.X, center.Y, center.Z };

            var circle = ms.AddCircle(vlist, radius);
            circle.Layer = layer;

            if (Math.Abs(Math.Abs(c.Normal.Dot(Vector.ZAxis())) - 1) > 0.001)
            {
                Rotate3DByCurveNormal(doc, circle.Handle, c);
            }

            Utils.Log(string.Format("Utils.AddCircleByCircle completed.", ""));

            return circle.Handle;
        }


        /// <summary>
        /// Adds a light weigth polyline by points.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="points">The points.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddLWPolylineByPoints(AeccRoadwayDocument doc, IList<Point> points, string layer)
        {
            Utils.Log(string.Format("Utils.AddLWPolylineByPoints started...", ""));

            AddLayer(doc, layer);

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            double[] vlist = new double[2 * points.Count];

            for (int i = 0; i < points.Count; ++i)
            {
                vlist[2 * i] = points[i].X;
                vlist[2 * i + 1] = points[i].Y;
            }

            var pl = ms.AddLightWeightPolyline(vlist);
            pl.Layer = layer;

            Utils.Log(string.Format("Utils.AddLWPolylineByPoints completed.", ""));

            return pl.Handle;
        }


        /// <summary>
        /// Adds a light weight polyline by poly curve.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="polycurve">The polycurve.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddLWPolylineByPolyCurve(AeccRoadwayDocument doc, PolyCurve polycurve, string layer)
        {
            Utils.Log(string.Format("Utils.AddLWPolylineByPolyCurve started...", ""));

            var totalTransform = RevitUtils.DocumentTotalTransform();

            polycurve = polycurve.Transform(totalTransform.Inverse()) as PolyCurve;

            IList<Point> points = polycurve.Curves().Select<Curve, Point>(c => c.StartPoint).ToList();

            points.Add(polycurve.EndPoint);

            Utils.Log(string.Format("Utils.AddLWPolylineByPolyCurve completed.", ""));

            return AddLWPolylineByPoints(doc, points, layer);
        }


        /// <summary>
        /// Rotates in 3D by curve normal.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="handle">The handle.</param>
        /// <param name="dynCurve">The dyn curve.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string Rotate3DByCurveNormal(AeccRoadwayDocument doc, string handle, Curve dynCurve)
        {
            Utils.Log(string.Format("Utils.Rotate3DByCurveNormal started...", ""));

            dynamic curve = doc.HandleToObject(handle);

            CoordinateSystem cs = dynCurve.ContextCoordinateSystem;

            double[,] transform = new double[4, 4];

            transform[0, 0] = cs.XAxis.X / cs.XScaleFactor;
            transform[0, 1] = cs.YAxis.X / cs.XScaleFactor;
            transform[0, 2] = cs.ZAxis.X / cs.XScaleFactor;
            transform[0, 3] = cs.Origin.X;

            transform[1, 0] = cs.XAxis.Y / cs.YScaleFactor;
            transform[1, 1] = cs.YAxis.Y / cs.YScaleFactor;
            transform[1, 2] = cs.ZAxis.Y / cs.YScaleFactor;
            transform[1, 3] = cs.Origin.Y;

            transform[2, 0] = cs.XAxis.Z / cs.ZScaleFactor;
            transform[2, 1] = cs.YAxis.Z / cs.ZScaleFactor;
            transform[2, 2] = cs.ZAxis.Z / cs.ZScaleFactor;
            transform[2, 3] = cs.Origin.Z / cs.ZScaleFactor;

            transform[3, 0] = 0;
            transform[3, 1] = 0;
            transform[3, 2] = 0;
            transform[3, 3] = 1;

            curve.TransformBy(transform);

            Utils.Log(string.Format("Utils.Rotate3DByCurveNormal completed.", ""));

            return handle;
        }


        /// <summary>
        /// Rotates the by vector.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="handle">The handle.</param>
        /// <param name="vector">The vector.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string RotateByVector(AeccRoadwayDocument doc, string handle, Vector vector)
        {
            Utils.Log(string.Format("Utils.RotateByVector started...", ""));

            Point p1 = null;
            Vector v = null;

            dynamic curve = doc.HandleToObject(handle);

            v = Vector.ByCoordinates(vector.X, vector.Y, 0);
            p1 = Point.ByCoordinates(curve.InsertionPoint[0], curve.InsertionPoint[1], curve.InsertionPoint[2]);

            double[] a1 = new double[] { p1.X, p1.Y, p1.Z };

            double rotation = Vector.XAxis().AngleAboutAxis(v, Vector.ZAxis());

            curve.Rotate(a1, DegToRad(rotation));

            p1.Dispose();
            v.Dispose();

            Utils.Log(string.Format("Utils.RotateByVector completed.", ""));

            return handle;
        }


        /// <summary>
        /// Rotate3s the d by plane.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="handle">The handle.</param>
        /// <param name="plane">The plane.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string Rotate3DByPlane(AeccRoadwayDocument doc, string handle, Plane plane)
        {
            Utils.Log(string.Format("Utils.Rotate3DByPlane started...", ""));

            Point p1 = null;
            Point p2 = null;
            Vector v = null;

            dynamic curve = doc.HandleToObject(handle);

            v = Vector.ByCoordinates(curve.Normal[0], curve.Normal[1], curve.Normal[2]);
            p1 = Point.ByCoordinates(curve.InsertionPoint[0], curve.InsertionPoint[1], curve.InsertionPoint[2]);
            p2 = p1.Translate(plane.Normal.Cross(v)) as Point;

            double[] a1 = new double[] { p1.X, p1.Y, p1.Z };
            double[] a2 = new double[] { p2.X, p2.Y, p2.Z };

            double rotation = Math.Acos(v.Dot(plane.Normal));

            curve.Rotate3D(a1, a2, DegToRad(180) - rotation);

            p1.Dispose();
            p2.Dispose();
            v.Dispose();

            Utils.Log(string.Format("Utils.Rotate3DByPlane completed.", ""));

            return handle;
        }


        /// <summary>
        /// Adds the polyline by curve.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="curve">The curve.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddPolylineByCurve(AeccRoadwayDocument doc, Curve curve, string layer)
        {
            Utils.Log(string.Format("Utils.AddPolylineByCurve started...", ""));

            IList<string> temp = new List<string>();

            if (curve.ToString().Contains("Line"))
            {
                temp.Add(AddPolylineByPoints(doc, new List<Point>() { curve.StartPoint, curve.EndPoint }, layer));
            }
            else if (curve.ToString().Contains("Circle"))
            {
                Circle circle = curve as Circle;

                temp.Add(AddCircleByCircle(doc, circle, layer));

                Rotate3DByCurveNormal(doc, temp.Last(), circle);
            }
            else if (curve.ToString().Contains("Arc"))
            {
                Arc arc = curve as Arc;

                temp.Add(AddArcByArc(doc, arc, layer));
            }
            else if (curve.ToString().Contains("PolyCurve") ||
                curve.ToString().Contains("Rectangle") ||
                curve.ToString().Contains("Polygon"))
            {
                PolyCurve polycurve = curve as PolyCurve;

                Acad3DPolyline pl = doc.HandleToObject(AddPolylineByPoints(doc, new List<Point>() { polycurve.CurveAtIndex(0).StartPoint, polycurve.CurveAtIndex(0).EndPoint }, layer));

                if (polycurve.IsClosed)
                {
                    for (int i = 1; i < polycurve.NumberOfCurves - 1; ++i)
                    {
                        Point end = polycurve.CurveAtIndex(i).EndPoint;

                        pl.AppendVertex(new double[] { end.X, end.Y, end.Z });
                    }

                    pl.Closed = true;
                }
                else
                {
                    for (int i = 1; i < polycurve.NumberOfCurves; ++i)
                    {
                        Point end = polycurve.CurveAtIndex(i).EndPoint;

                        pl.AppendVertex(new double[] { end.X, end.Y, end.Z });
                    }

                    pl.Closed = false;
                }

                temp.Add(pl.Handle);

            }           

            Utils.Log(string.Format("Utils.AddPolylineByCurve completed.", ""));

            // TODO: handle nurbs curves
            return temp[0];

        }


        /// <summary>
        /// Adds the polyline by curves.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="curves">The curves.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddPolylineByCurves(AeccRoadwayDocument doc, IList<Curve> curves, string layer)
        {
            Utils.Log(string.Format("Utils.AddPolylineByCurves started...", ""));

            IList<Point> points = new List<Point>();

            foreach (Curve crv in curves)
            {
                points.Add(crv.StartPoint);
            }

            if (curves.First().StartPoint.DistanceTo(curves.Last().EndPoint) < 0.001)
            {
                points.Add(points.First());
            }

            Utils.Log(string.Format("Utils.AddPolylineByCurves completed.", ""));

            return AddPolylineByPoints(doc, points, layer);
        }


        /// <summary>
        /// Adds the extruded solid by points.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="points">The points.</param>
        /// <param name="height">The height.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddExtrudedSolidByPoints(AeccRoadwayDocument doc, IList<Point> points, double height, string layer)
        {
            Utils.Log(string.Format("Utils.AddExtrudedSolidByPoints started...", ""));

            AddLayer(doc, layer);

            Acad3DSolid solid = null;

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            Acad3DPolyline pl = doc.HandleToObject(AddPolylineByPoints(doc, points, layer)) as Acad3DPolyline;

            if (pl.Closed)
            {
                var collection = pl.Explode();

                AcadEntity[] obj = new AcadEntity[collection.Length];

                for (int i = 0; i < collection.Length; ++i)
                {
                    obj[i] = collection[i] as AcadEntity;
                }

                var region = ms.AddRegion(obj)[0];
                region.Layer = layer;

                pl.Delete();

                foreach (var l in obj)
                {
                    l.Delete();
                }

                solid = ms.AddExtrudedSolid(region, height, 0);
                solid.Layer = layer;

                region.Delete();
            }

            Utils.Log(string.Format("Utils.AddExtrudedSolidByPoints completed.", ""));

            return solid.Handle;
        }


        /// <summary>
        /// Adds the region by patch.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="curve">The curve.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddRegionByPatch(AeccRoadwayDocument doc, Curve curve, string layer)
        {
            Utils.Log(string.Format("Utils.AddRegionByPatch started...", ""));

            AddLayer(doc, layer);

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            IList<Acad3DPolyline> polylines = new List<Acad3DPolyline>();

            string id = AddPolylineByCurve(doc, curve, layer);

            Acad3DPolyline pl = doc.HandleToObject(id) as Acad3DPolyline;

            var collection = pl.Explode();

            AcadEntity[] obj = new AcadEntity[collection.Length];

            for (int i = 0; i < collection.Length; ++i)
            {
                obj[i] = collection[i] as AcadEntity;
            }

            var region = ms.AddRegion(obj)[0];
            region.Layer = layer;

            pl.Delete();

            foreach (var l in obj)
            {
                l.Delete();
            }

            Utils.Log(string.Format("Utils.AddRegionByPatch completed.", ""));

            return region.Handle;
        }


        /// <summary>
        /// Adds the extruded solid by patch.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="curve">The curve.</param>
        /// <param name="height">The height.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddExtrudedSolidByPatch(AeccRoadwayDocument doc, Curve curve, double height, string layer)
        {
            Utils.Log(string.Format("Utils.AddExtrudedSolidByPatch started...", ""));

            AddLayer(doc, layer);

            Acad3DSolid solid = null;

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            var r = doc.HandleToObject(AddRegionByPatch(doc, curve, layer));
            r.Layer = layer;

            solid = ms.AddExtrudedSolid(r, height, 0);
            solid.Layer = layer;

            r.Delete();

            Utils.Log(string.Format("Utils.AddExtrudedSolidByPatch completed.", ""));

            return solid.Handle;
        }


        /// <summary>
        /// Adds the extruded solid by curves.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="curves">The curves.</param>
        /// <param name="height">The height.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string AddExtrudedSolidByCurves(AeccRoadwayDocument doc, IList<Curve> curves, double height, string layer)
        {
            Utils.Log(string.Format("Utils.AddExtrudedSolidByCurves started...", ""));

            AddLayer(doc, layer);

            Acad3DSolid solid = null;

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            IList<Acad3DPolyline> polylines = new List<Acad3DPolyline>();

            string id = AddPolylineByCurves(doc, curves, layer);

            if (id != null)
            {
                Acad3DPolyline pl = doc.HandleToObject(id) as Acad3DPolyline;
                if (pl.Closed)
                {
                    polylines.Add(pl);
                }
            }

            foreach (Acad3DPolyline pl in polylines)
            {
                var collection = pl.Explode();

                AcadEntity[] obj = new AcadEntity[collection.Length];

                for (int i = 0; i < collection.Length; ++i)
                {
                    obj[i] = collection[i] as AcadEntity;
                }

                var region = ms.AddRegion(obj)[0];
                region.Layer = layer;

                pl.Delete();

                foreach (var l in obj)
                {
                    l.Delete();
                }

                solid = ms.AddExtrudedSolid(region, height, 0);
                solid.Layer = layer;

                region.Delete();
            }

            Utils.Log(string.Format("Utils.AddExtrudedSolidByCurves completed.", ""));

            return solid.Handle;
        }


        /// <summary>
        /// Cuts the solids by patch.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="closedCurve">The closed curve.</param>
        /// <param name="height">The height.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static bool CutSolidsByPatch(AeccRoadwayDocument doc, Curve closedCurve, double height, string layer)
        {
            Utils.Log(string.Format("Utils.CutSolidsByPatch started...", ""));

            bool result = false;

            IList<Acad3DSolid> cSolids = new List<Acad3DSolid>();

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            foreach (AcadEntity s in ms)
            {
                if (s.EntityName.Contains("Solid"))
                {
                    if (!s.Layer.Equals(layer))
                    {
                        cSolids.Add(s as Acad3DSolid);
                    }
                }
            }

            var solid = doc.HandleToObject(AddExtrudedSolidByPatch(doc, closedCurve, height, layer));

            var operation = AcBooleanType.acSubtraction;

            if (cSolids.Count > 0)
            {
                foreach (Acad3DSolid cs in cSolids)
                {
                    bool interference = false;

                    Acad3DSolid interf = cs.CheckInterference(solid, true, out interference);

                    if (interference)
                    {
                        cs.Boolean(operation, interf);
                        result = true;
                    }
                }
            }

            FreezeLayers(doc, layer);

            Utils.Log(string.Format("Utils.CutSolidsByPatch completed.", ""));

            return result;
        }

        /// <summary>
        /// Cuts the solids by curves.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="closedCurves">The closed curves.</param>
        /// <param name="height">The height.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static bool CutSolidsByCurves(AeccRoadwayDocument doc, IList<Curve> closedCurves, double height, string layer)
        {
            Utils.Log(string.Format("Utils.CutSolidsByCurves started...", ""));

            bool result = false;

            IList<Acad3DSolid> cSolids = new List<Acad3DSolid>();

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            foreach (AcadEntity s in ms)
            {
                if (s.EntityName.Contains("Solid") && !s.Layer.Equals(layer))
                {
                    cSolids.Add((Acad3DSolid)s);
                }
            }

            var solid = doc.HandleToObject(AddExtrudedSolidByCurves(doc, closedCurves, height, layer));

            var operation = AcBooleanType.acSubtraction;

            foreach (Acad3DSolid cs in cSolids)
            {
                bool interference = false;

                Acad3DSolid interf = cs.CheckInterference(solid, true, out interference);

                if (interference)
                {
                    cs.Boolean(operation, interf);
                    result = true;
                }
            }

            FreezeLayers(doc, layer);

            Utils.Log(string.Format("Utils.CutSolidsByCurves completed.", ""));

            return result;
        }


        /// <summary>
        /// Cuts the solids by geometry.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="geometry">The geometry.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static bool CutSolidsByGeometry(AeccRoadwayDocument doc, Geometry[] geometry, string layer)
        {
            Utils.Log(string.Format("Utils.CutSolidsByGeometry started...", ""));

            bool result = false;

            var handles = ImportGeometry(doc, geometry, layer);

            IList<Acad3DSolid> cSolids = new List<Acad3DSolid>();

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            foreach (AcadEntity s in ms)
            {
                if (s.EntityName.Contains("Solid") && s.Layer != layer)
                {
                    cSolids.Add((Acad3DSolid)s);
                }
            }

            var operation = AcBooleanType.acSubtraction;

            foreach (var handle in handles)
            {
                var solid = doc.HandleToObject(handle);

                foreach (Acad3DSolid cs in cSolids)
                {
                    bool interference = false;

                    Acad3DSolid interf = cs.CheckInterference(solid, true, out interference);

                    if (interference)
                    {
                        cs.Boolean(operation, interf);
                        result = true;
                    }
                }
            }

            FreezeLayers(doc, layer);

            Utils.Log(string.Format("Utils.CutSolidsByGeometry completed.", ""));

            return result;
        }


        /// <summary>
        /// Slices the solids by plane.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="plane">The plane..</param>     
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static bool SliceSolidsByPlane(AeccRoadwayDocument doc, Plane plane)
        {
            Utils.Log(string.Format("Utils.SliceSolidsByPlane started...", ""));

            bool result = false;

            IList<Acad3DSolid> cSolids = new List<Acad3DSolid>();

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            foreach (AcadEntity s in ms)
            {
                if (s.EntityName.Contains("Solid"))
                {
                    cSolids.Add((Acad3DSolid)s);
                }
            }

            Point a = plane.Origin;
            Point b = a.Add(plane.XAxis);
            Point c = a.Add(plane.YAxis);

            Acad3DSolid solid = null;

            foreach (Acad3DSolid cs in cSolids)
            {
                try
                {
                    solid = cs.SliceSolid(new double[] { a.X, a.Y, a.Z },
                                  new double[] { b.X, b.Y, b.Z },
                                  new double[] { c.X, c.Y, c.Z },
                                  true); // If set to true keeps both parts of the solid

                    result = true;
                }
                catch (Exception ex)
                {
                    Utils.Log(string.Format("ERROR: Utils.SliceSolidsByPlane {0}", ex.Message));

                    result = false;
                }
            }

            Utils.Log(string.Format("Utils.SliceSolidsByPlane completed.", ""));

            return result;
        }


        /// <summary>
        /// Imports the geometry.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="geometry">The geometry.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<string> ImportGeometry(AeccRoadwayDocument doc, Geometry[] geometry, string layer)
        {
            Utils.Log(string.Format("Utils.ImportGeometry started...", ""));

            if (geometry.Length == 0)
            {
                Utils.Log(string.Format("No geometry!", ""));

                return null;
            }

            AddLayer(doc, layer);

            IList<string> currentHandles = new List<string>();
            IList<string> newHandles = new List<string>();

            Dictionary<Geometry, string> dict = new Dictionary<Geometry, string>();

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            // this is to make sure that the geometries are processed in the same order as they are coming in the geometry inputs

            IList<Geometry> solids = new List<Geometry>();  // 20191028

            foreach (Geometry g in geometry)
            {
                if (g is Solid)
                {
                    solids.Add(g);
                }

                else if (g is Arc)
                {
                    Arc arc = g as Arc;

                    dict.Add(g, AddArcByArc(doc, arc, layer));  // 20191028
                }
                else if (g is Curve)
                {
                    Curve c = g as Curve;

                    dict.Add(g, AddPolylineByCurve(doc, c, layer));  // 20191028
                }
                else if (g is Point)
                {
                    Point p = g as Point;

                    dict.Add(g, AddPointByPoint(doc, p, layer));  // 20191028
                }
            }

            // 20191028
            if (solids.Count > 0)
            {
                var solidsArray = solids.ToArray();

                string path = Path.Combine(Path.GetTempPath(), "CivilConnection.sat");

                Geometry.ExportToSAT(solidsArray, path);

                int start = ms.Count;

                doc.Import(path, new double[] { 0, 0, 0 }, 1);

                int end = ms.Count;

                IList<AcadEntity> added = new List<AcadEntity>();

                for (int i = start; i < end; ++i)
                {
                    var s = ms.Item(i);

                    if (s.EntityName.Contains("Solid") || s.EntityName.Contains("Surface"))  // The assumption is that the import of the SAT file respects the order of the solids
                    {
                        if (!currentHandles.Contains(s.Handle))
                        {
                            s.Layer = layer;

                            added.Add(s);
                        }
                    }
                }

                for (int i = 0; i < added.Count; ++i)
                {
                    dict.Add(solids[i], added[i].Handle);
                }

                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }

            // reorder handles to match intial input order

            foreach (var g in geometry)
            {
                newHandles.Add(dict[g]);
            }

            Utils.Log(string.Format("Utils.ImportGeometry completed.", ""));

            return newHandles;
        }

        /// <summary>
        /// Imports the geometry of a SAT file
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="path"></param>
        /// <param name="layer"></param>
        /// <returns></returns>
        public static IList<string> ImportGeometryByPath(AeccRoadwayDocument doc, string path, string layer)
        {
            if (Path.GetExtension(path).ToLower() != ".sat")
            {
                Utils.Log(string.Format("ERROR: File path is not a SAT file", ""));

                throw new Exception("File path is not a SAT file");
            }

            Utils.Log(string.Format("Utils.ImportGeometryByPath started...", ""));

            AddLayer(doc, layer);

            IList<string> currentHandles = new List<string>();
            IList<string> newHandles = new List<string>();

            Dictionary<Geometry, string> dict = new Dictionary<Geometry, string>();

            AcadDatabase db = doc as AcadDatabase;

            AcadModelSpace ms = db.ModelSpace;

            int start = ms.Count;

            doc.Import(path, new double[] { 0, 0, 0 }, 1);

            int end = ms.Count;

            for (int i = start; i < end; ++i)
            {
                var s = ms.Item(i);

                if (s.EntityName.Contains("Solid") || s.EntityName.Contains("Surface"))  // The assumption is that the import of the SAT file respects the order of the solids
                {
                    if (!currentHandles.Contains(s.Handle))
                    {
                        s.Layer = layer;

                        newHandles.Add(s.Handle);
                    }
                }
            }

            Utils.Log(string.Format("Utils.ImportGeometryByPath completed.", ""));

            return newHandles;
        }


        //TODO: this node is not working when the geometry extraction from one of the solids returns a null or raise an exception

        /// <summary>
        /// Imports the element.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="element">The element.</param>
        /// <param name="parameter">The parameter.</param>
        /// <param name="layer">The layer.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string ImportElement(AeccRoadwayDocument doc, Revit.Elements.Element element, string parameter, string layer)
        {
            Utils.Log(string.Format("Utils.ImportElement started...", ""));

            var totalTransform = RevitUtils.DocumentTotalTransform();

            var totalTransformInverse = RevitUtils.DocumentTotalTransformInverse();

            string result = "";
            string handles = "";
            Acad3DSolid existent = null;
            bool update = false;

            try
            {
                handles = Convert.ToString(element.GetParameterValueByName(parameter));

                foreach (var id in handles.Split(new string[] { "," }, StringSplitOptions.None))
                {
                    var temp = doc.HandleToObject(id) as Acad3DSolid;

                    if (temp != null)
                    {
                        update = true;

                        if (null != existent)
                        {
                            existent.Boolean(AcBooleanType.acUnion, temp);
                        }
                        else
                        {
                            existent = temp;
                        }
                    }
                }
            }
            catch
            { }

            try
            {
                IList<Solid> temp = element.Solids.ToList();
                IList<Solid> solids = new List<Solid>();

                Solid s = temp[0] as Solid;
                s = s.Transform(totalTransformInverse) as Solid;

                temp.RemoveAt(0);

                if (temp.Count > 0)
                {
                    foreach (Geometry g in temp)
                    {
                        Solid gs = null;

                        try
                        {
                            gs = g as Solid;
                            gs = gs.Transform(totalTransformInverse) as Solid;
                            s = Solid.ByUnion(new Solid[] { s, gs });
                        }
                        catch
                        {
                            if (null != gs)
                            {
                                solids.Add(gs);
                            }
                            continue;
                        }
                    }
                }

                solids.Add(s);

                IList<string> handlesList = new List<string>();

                Acad3DSolid newSolid = null;

                foreach (Solid i in solids)
                {
                    var ids = ImportGeometry(doc, new Geometry[] { i }, layer);

                    if (ids.Count > 0)
                    {
                        handlesList.Add(ids[0]);

                        Acad3DSolid tempSolid = doc.HandleToObject(ids[0]) as Acad3DSolid;

                        if (newSolid != null)
                        {
                            newSolid.Boolean(AcBooleanType.acUnion, tempSolid);
                        }
                        else
                        {
                            newSolid = tempSolid;
                        }
                    }
                }

                if (newSolid != null)
                {
                    result = newSolid.Handle;
                }
                else
                {
                    result = string.Join(",", handlesList);
                }

                s.Dispose();
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }

            if (!update || existent == null)
            {
                handles = result;
            }
            else
            {
                handles = Convert.ToString(existent.Handle);
            }

            if (update)
            {
                try
                {

                    doc.SendCommand(string.Format("-ReplaceSolid \"{0}\"\n\"{1}\"\n\n", handles, result));
                }
                catch (Exception ex)
                {
                    Utils.Log(string.Format("ERROR: Utils.ImportElement {0}", ex.Message));

                    MessageBox.Show(string.Format("PythonScript Failed\n\n{0}", ex.Message));
                }
            }

            element.SetParameterByName(parameter, handles);

            Utils.Log(string.Format("Utils.ImportElement completed.", ""));

            return handles;
        }


        /// <summary>
        /// Dumps the land XML.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static string DumpLandXML(AeccRoadwayDocument doc)
        {
            Utils.Log(string.Format("Utils.DumpLandXML started...", ""));

            string landxml = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(doc.Name) + ".xml");

            if (!File.Exists(landxml))  // 1.1.0
            {
                doc.SendCommand("-aecclandxmlout\n" + landxml + "\n");  // 1.1.0
                SessionVariables.IsLandXMLExported = true;  // 1.1.0
            }
            else if (File.Exists(landxml) && !SessionVariables.IsLandXMLExported)
            {
                File.Delete(landxml);
                doc.SendCommand("-aecclandxmlout\n" + landxml + "\n");  // 1.1.0
                SessionVariables.IsLandXMLExported = true;  // 1.1.0
            }

            // asynchronous task

            while (!File.Exists(landxml))  // 1.1.0
            {
                // HACK: wait until the file is ready
                int i = 0;
            }

            Utils.Log(string.Format("Utils.DumpLandXML completed.", ""));

            return landxml;
        }


        /// <summary>
        /// Gets the XML document.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <returns></returns>
        /// <exception cref="Exception">Error in Loading XML</exception>
        [IsVisibleInDynamoLibrary(false)]
        public static XmlDocument GetXmlDocument(AeccRoadwayDocument doc)
        {
            Utils.Log(string.Format("Utils.GetXmlDocument started...", ""));

            XmlDocument xmlDoc = new XmlDocument();

            try
            {
                string landxml = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(doc.Name) + ".xml");

                if (!File.Exists(landxml))
                {
                    SessionVariables.IsLandXMLExported = false;  // 1.1.0

                    DumpLandXML(doc);  // 1.1.0
                }

                xmlDoc.Load(landxml);
            }
            catch (Exception ex)
            {
                var message = string.Format("ERROR: Utils.GetXmlDocument {0} {1}", "Error in Loading XML", ex.Message);

                Utils.Log(message);

                throw new Exception(message);
            }

            Utils.Log(string.Format("Utils.GetXmlDocument completed.", ""));

            return xmlDoc;
        }


        /// <summary>
        /// Gets the XML namespace manager.
        /// </summary>
        /// <param name="xmlDoc">The XML document.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static XmlNamespaceManager GetXmlNamespaceManager(XmlDocument xmlDoc)
        {
            Utils.Log(string.Format("Utils.GetXmlNamespaceManager started...", ""));

            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);

            if (xmlDoc.DocumentElement.NamespaceURI == "http://www.landxml.org/schema/LandXML-1.2")
            {
                nsmgr.AddNamespace("lx", "http://www.landxml.org/schema/LandXML-1.2");
            }
            else if (xmlDoc.DocumentElement.NamespaceURI == "http://www.landxml.org/schema/LandXML-1.1")
            {
                nsmgr.AddNamespace("lx", "http://www.landxml.org/schema/LandXML-1.1");
            }
            else
            {
                nsmgr.AddNamespace("lx", "http://www.landxml.org/schema/LandXML-1.0");
            }

            Utils.Log(string.Format("Utils.GetXmlNamespaceManager completed.", ""));

            return nsmgr;
        }


        /// <summary>
        /// Function that writes an entry to the log file
        /// </summary>
        /// <param name="message"></param>
        [IsVisibleInDynamoLibrary(false)]
        public static void Log(string message)
        {
            string path = System.IO.Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "CivilConnection_temp.log");

            using (StreamWriter sw = new StreamWriter(path, true))
            {
                sw.WriteLine(string.Format("[{0}] {1}", DateTime.Now, message));
            }
        }

        /// <summary>
        /// Finalizes the Log file.
        /// </summary>
        [IsVisibleInDynamoLibrary(false)]
        public static void InitializeLog()
        {
            string path = System.IO.Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "CivilConnection_temp.log");

            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        /// <summary>
        /// Gets the corridor subassemblies shapes from LandXML.
        /// </summary>
        /// <param name="corridor">The corridor.</param>
        /// <param name="dumpXML">If True exports a LandXML in the Temp folder.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<IList<IList<IList<IList<Point>>>>> GetCorridorSubAssembliesFromLandXML(Corridor corridor, bool dumpXML = false)
        {
            Utils.Log(string.Format("Utils.GetCorridorSubAssembliesFromLandXML started...", ""));

            IList<IList<IList<IList<IList<Point>>>>> corrPoints = new List<IList<IList<IList<IList<Point>>>>>();

            SessionVariables.IsLandXMLExported = !dumpXML;  // 1.1.0

            AeccRoadwayDocument doc = corridor._document;

            if (dumpXML)
            {
                DumpLandXML(doc);
                Log(string.Format("Create LandXML", ""));
            }

            // XmlDocument xmlDoc = GetXmlDocument(doc, corridor.Name);
            XmlDocument xmlDoc = GetXmlDocument(doc);

            Log(string.Format("Create XML document", ""));

            XmlNamespaceManager nsmgr = GetXmlNamespaceManager(xmlDoc);

            Log(string.Format("LandXML namespace ok", ""));

            CoordinateSystem cs = CoordinateSystem.Identity();

            foreach (Baseline b in corridor.Baselines)
            {
                IList<IList<IList<IList<Point>>>> baseline = new List<IList<IList<IList<Point>>>>();

                Log(string.Format("Processing Baseline {0}...", b.Index));

                foreach (AeccBaselineRegion blr in b._baseline.BaselineRegions)  // 1.1.0
                {
                    double start = blr.StartStation;
                    double end = blr.EndStation;

                    IList<IList<IList<Point>>> baselineRegion = new List<IList<IList<Point>>>();

                    Log(string.Format("Processing Baseline Region {0} - {1}...", start, end));

                    string[] separator = new string[] { " " };

                    string alName = b.Alignment.Name.Replace(' ', '_');   // this replacement happens when exporting to LandXML from Civil 3D

                    Log(string.Format("Processing Alignment {0}...", alName));

                    foreach (XmlNode alignmentXml in xmlDoc.SelectNodes(string.Format("//lx:Alignment[@name = '{0}']", alName), nsmgr))
                    {
                        Log(string.Format("Alignment {0} found!", alName));

                        foreach (XmlNode assembly in alignmentXml.SelectNodes(".//lx:CrossSect", nsmgr))
                        {
                            IList<IList<Point>> assPoints = new List<IList<Point>>();

                            double station = Convert.ToDouble(assembly.Attributes["sta"].Value, System.Globalization.CultureInfo.InvariantCulture);

                            Log(string.Format("Processing Station {0}...", station));

                            if (Math.Abs(station - start) < 0.001)
                            {
                                station = start;
                            }
                            if (Math.Abs(station - end) < 0.001)
                            {
                                station = end;
                            }

                            if (station >= start && station <= end)
                            {
                                cs = b.CoordinateSystemByStation(station);

                                foreach (XmlNode subassembly in assembly.SelectNodes("lx:DesignCrossSectSurf", nsmgr))
                                {
                                    IList<Point> subPoints = new List<Point>();

                                    Log(string.Format("Processing Subassembly {0} points...", subassembly.ChildNodes.Count));

                                    if (subassembly.ChildNodes.Count > 2)  // 20180810 - Changed to skip Links in the processing
                                    {
                                        foreach (XmlNode calcPoint in subassembly.SelectNodes("lx:CrossSectPnt", nsmgr))
                                        {
                                            string[] coords = calcPoint.InnerText.Split(separator, StringSplitOptions.None);

                                            subPoints.Add(Point.ByCoordinates(Convert.ToDouble(coords[0], System.Globalization.CultureInfo.InvariantCulture), 0, Convert.ToDouble(coords[1], System.Globalization.CultureInfo.InvariantCulture)).Transform(cs) as Point);

                                            Log(string.Format("Processing Coordinates...", ""));
                                        }

                                        var temp = Point.PruneDuplicates(subPoints).ToList();

                                        // discard links
                                        if (temp.Count > 2)
                                        {
                                            assPoints.Add(temp);

                                            Log(string.Format("Subassembly Points added!", ""));
                                        }
                                    }
                                }
                            }

                            if (assPoints.Count > 0)
                            {
                                baselineRegion.Add(assPoints);

                                Log(string.Format("Assembly Points added!", ""));
                            }
                        }
                    }

                    // 20180810 - Changed it throws some errors needs investigation

                    baseline.Add(baselineRegion);

                    Log(string.Format("Region Points added!", ""));
                }

                corrPoints.Add(baseline);

                Log(string.Format("Baseline Points added!", ""));
            }

            cs.Dispose();

            Utils.Log(string.Format("Utils.GetCorridorSubAssembliesFromLandXML completed.", ""));

            return corrPoints;
        }

        /// <summary>
        /// Gets the corridor points by code from land XML.
        /// </summary>
        /// <param name="corridor">The corridor.</param>
        /// <param name="code">The code.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<IList<IList<IList<IList<Point>>>>> GetCorridorPointsByCodeFromLandXML(Corridor corridor, string code)
        {
            Utils.Log(string.Format("Utils.GetCorridorPointsByCodeFromLandXML started...", ""));

            Log(string.Format("Code: {0}", code));

            IList<IList<IList<IList<IList<Point>>>>> corrPoints = new List<IList<IList<IList<IList<Point>>>>>();

            AeccRoadwayDocument doc = corridor._document;

            Log(string.Format("Create LandXML", ""));

            XmlDocument xmlDoc = GetXmlDocument(doc);

            XmlNamespaceManager nsmgr = GetXmlNamespaceManager(xmlDoc);

            CoordinateSystem cs = CoordinateSystem.Identity();

            foreach (Baseline b in corridor.Baselines)
            {
                Log(string.Format("Processing Baseline {0}...", b.Index));

                IList<IList<IList<IList<Point>>>> baseline = new List<IList<IList<IList<Point>>>>();

                foreach (AeccBaselineRegion blr in b._baseline.BaselineRegions)  // 1.1.0
                {
                    double start = blr.StartStation;
                    double end = blr.EndStation;

                    Log(string.Format("Processing Baseline Region {0} - {1}...", start, end));

                    IList<IList<IList<Point>>> baselineRegion = new List<IList<IList<Point>>>();

                    string[] separator = new string[] { " " };

                    string alName = b.Alignment.Name.Replace(' ', '_');   // this replacement happens when exporting to LandXML form Civil 3D

                    foreach (XmlNode alignmentXml in xmlDoc.SelectNodes(string.Format("//lx:Alignment[@name = '{0}']", alName), nsmgr))
                    {
                        Log(string.Format("Processing Alignment {0}...", alName));

                        foreach (XmlNode assembly in alignmentXml.SelectNodes(".//lx:CrossSect", nsmgr))
                        {
                            IList<IList<Point>> assPoints = new List<IList<Point>>();

                            double station = Convert.ToDouble(assembly.Attributes["sta"].Value, System.Globalization.CultureInfo.InvariantCulture);

                            Log(string.Format("Processing Station {0}...", station));

                            if (Math.Abs(station - start) < 0.001)
                            {
                                station = start;
                            }
                            if (Math.Abs(station - end) < 0.001)
                            {
                                station = end;
                            }

                            if (station >= start && station <= end)
                            {
                                cs = b.CoordinateSystemByStation(station);

                                IList<Point> left = new List<Point>();

                                foreach (XmlNode calcPoint in assembly.SelectNodes(string.Format(".//lx:CrossSectPnt[@code = '{0}']", code), nsmgr))
                                {


                                    string[] coords = calcPoint.InnerText.Split(separator, StringSplitOptions.None);

                                    left.Add(Point.ByCoordinates(Convert.ToDouble(coords[0], System.Globalization.CultureInfo.InvariantCulture), 0, Convert.ToDouble(coords[1], System.Globalization.CultureInfo.InvariantCulture)).Transform(cs) as Point);
                                }

                                Log(string.Format("Processed {0} Calculated Points...", Point.PruneDuplicates(left).Length));

                                assPoints.Add(Point.PruneDuplicates(left));
                            }

                            if (assPoints.Count > 0)
                            {
                                baselineRegion.Add(assPoints);
                            }
                        }
                    }

                    try // 1.1.0
                    {
                        baselineRegion = baselineRegion.OrderBy(p => b.GetArrayStationOffsetElevationByPoint(p[0][0])[0]).ToList();
                    }
                    catch (Exception ex)
                    {
                        Log(string.Format("Error occured: {0}", ex.Message));
                    }

                    baseline.Add(baselineRegion);

                    Log(string.Format("Region Points added!", ""));
                }

                corrPoints.Add(baseline);

                Log(string.Format("Baseline Points added!", ""));
            }

            cs.Dispose();

            Utils.Log(string.Format("Utils.GetCorridorPointsByCodeFromLandXML completed.", ""));

            return corrPoints;
        }

        /// <summary>
        /// Gets the feature lines by code from land XML.
        /// </summary>
        /// <param name="corridor">The corridor.</param>
        /// <param name="code">The code.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<IList<IList<Featureline>>> GetFeatureLinesByCodeFromLandXML(Corridor corridor, string code)
        {
            Utils.Log(string.Format("Utils.GetFeatureLinesByCodeFromLandXML started...", ""));

            AeccRoadwayDocument doc = corridor._document;

            XmlDocument xmlDoc = GetXmlDocument(doc);

            XmlNamespaceManager nsmgr = GetXmlNamespaceManager(xmlDoc);

            CoordinateSystem cs = CoordinateSystem.Identity();

            string[] separator = new string[] { " " };

            IList<IList<IList<Featureline>>> corridorFeaturelines = new List<IList<IList<Featureline>>>();

            foreach (Baseline b in corridor.Baselines)
            {
                IList<IList<Featureline>> baselineColl = new List<IList<Featureline>>();

                string alName = b.Alignment.Name.Replace(' ', '_');   // this replacement happens when exporting to LandXML form Civil 3D

                foreach (XmlNode alignmentXml in xmlDoc.SelectNodes(string.Format("//lx:Alignment[@name = '{0}']", alName), nsmgr))
                {
                    foreach (BaselineRegion blr in b.GetBaselineRegions())
                    {
                        IList<Featureline> featurelines = new List<Featureline>();

                        IList<Point> right = new List<Point>();

                        IList<Point> left = new List<Point>();

                        IList<Point> none = new List<Point>();

                        foreach (XmlNode assembly in alignmentXml.SelectNodes(".//lx:CrossSect", nsmgr))
                        {
                            double station = Convert.ToDouble(assembly.Attributes["sta"].Value, System.Globalization.CultureInfo.InvariantCulture);

                            if (Math.Abs(station - blr.Start) < 0.001)
                            {
                                station = blr.Start;
                            }
                            if (Math.Abs(station - blr.End) < 0.001)
                            {
                                station = blr.End;
                            }

                            if (station >= blr.Start && station <= blr.End)
                            {
                                cs = b.CoordinateSystemByStation(station);

                                foreach (XmlNode subassembly in assembly.SelectNodes("lx:DesignCrossSectSurf[@side = 'left' and @closedArea]", nsmgr))
                                {

                                    foreach (XmlNode calcPoint in subassembly.SelectNodes(string.Format("lx:CrossSectPnt[@code = '{0}']", code), nsmgr))
                                    {
                                        string[] coords = calcPoint.InnerText.Split(separator, StringSplitOptions.None);

                                        left.Add(Point.ByCoordinates(Convert.ToDouble(coords[0], System.Globalization.CultureInfo.InvariantCulture), 0, Convert.ToDouble(coords[1], System.Globalization.CultureInfo.InvariantCulture)).Transform(cs) as Point);
                                        break;
                                    }

                                }

                                foreach (XmlNode subassembly in assembly.SelectNodes("lx:DesignCrossSectSurf[@side = 'right' and @closedArea]", nsmgr))
                                {

                                    foreach (XmlNode calcPoint in subassembly.SelectNodes(string.Format("lx:CrossSectPnt[@code = '{0}']", code), nsmgr))
                                    {
                                        string[] coords = calcPoint.InnerText.Split(separator, StringSplitOptions.None);

                                        right.Add(Point.ByCoordinates(Convert.ToDouble(coords[0], System.Globalization.CultureInfo.InvariantCulture), 0, Convert.ToDouble(coords[1], System.Globalization.CultureInfo.InvariantCulture)).Transform(cs) as Point);
                                        break;
                                    }
           
                                }

                                foreach (XmlNode calcPoint in assembly.SelectNodes(string.Format(".//lx:CrossSectPnt[@code = '{0}']", code), nsmgr))
                                {
                                    string[] coords = calcPoint.InnerText.Split(separator, StringSplitOptions.None);

                                    none.Add(Point.ByCoordinates(Convert.ToDouble(coords[0], System.Globalization.CultureInfo.InvariantCulture), 0, Convert.ToDouble(coords[1], System.Globalization.CultureInfo.InvariantCulture)).Transform(cs) as Point);
                                }
                            }
                        }

                        if (left.Count > 1)
                        {
                            left = Point.PruneDuplicates(left);
                            left = left.OrderBy(p => b.GetArrayStationOffsetElevationByPoint(p)[0]).ToList();
                            if (left.Count > 1)
                            {
                                featurelines.Add(new Featureline(b, PolyCurve.ByPoints(left), code, Featureline.SideType.Left));
                            }
                        }

                        if (right.Count > 1)
                        {
                            right = Point.PruneDuplicates(right);
                            right = right.OrderBy(p => b.GetArrayStationOffsetElevationByPoint(p)[0]).ToList();
                            if (right.Count > 1)
                            {
                                featurelines.Add(new Featureline(b, PolyCurve.ByPoints(right), code, Featureline.SideType.Right));
                            }
                        }

                        if (none.Count > 1)
                        {
                            none = Point.PruneDuplicates(none);
                            none = none.OrderBy(p => b.GetArrayStationOffsetElevationByPoint(p)[0]).ToList();
                            var pc = PolyCurve.ByPoints(none);
                            var offset = b.GetArrayStationOffsetElevationByPoint(pc.PointAtParameter(0.5))[1];
                            var side = Featureline.SideType.Right;

                            if (offset < 0)
                            {
                                side = Featureline.SideType.Left;
                            }

                            if (none.Count > 1)
                            {
                                featurelines.Add(new Featureline(b, pc, code, side));
                            }
                        }

                        baselineColl.Add(featurelines);
                    }
                }

                corridorFeaturelines.Add(baselineColl);
            }

            cs.Dispose();

            Utils.Log(string.Format("Utils.GetFeatureLinesByCodeFromLandXML completed.", ""));

            return corridorFeaturelines;
        }

        /// <summary>
        /// Gets the featurelines from land XML.
        /// </summary>
        /// <param name="corridor">The corridor.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<IList<IList<Featureline>>> GetFeaturelinesFromLandXML(Corridor corridor)
        {
            Utils.Log(string.Format("Utils.GetFeaturelinesFromLandXML started...", ""));

            AeccRoadwayDocument doc = corridor._document;

            Log(string.Format("Create LandXML", ""));

            XmlDocument xmlDoc = GetXmlDocument(doc);

            XmlNamespaceManager nsmgr = GetXmlNamespaceManager(xmlDoc);

            CoordinateSystem cs = CoordinateSystem.Identity();

            string[] separator = new string[] { " " };

            IList<IList<IList<Featureline>>> corridorFeaturelines = new List<IList<IList<Featureline>>>();

            // TODO: what happens to the corridor in the LandXML if you have more than one baseline with different profiles?
            foreach (Baseline b in corridor.Baselines)
            {
                Log(string.Format("Processing Baseline {0}...", b.Index));

                IList<IList<Featureline>> baselineFeaturelines = new List<IList<Featureline>>();

                string alName = b.Alignment.Name.Replace(' ', '_');   // this replacement happens when exporting to LandXML form Civil 3D

                foreach (XmlNode alignmentXml in xmlDoc.SelectNodes(string.Format("//lx:Alignment[@name = '{0}']", alName), nsmgr))
                {
                    Log(string.Format("Processing Alignment {0}...", alName));

                    foreach (AeccBaselineRegion blr in b._baseline.BaselineRegions)  // 1.1.0
                    {
                        double start = blr.StartStation;
                        double end = blr.EndStation;

                        Log(string.Format("Processing Baseline Region {0} - {1}...", start, end));

                        IList<Featureline> featurelines = new List<Featureline>();

                        foreach (var code in corridor.GetCodes())
                        {
                            Log(string.Format("Code: {0}", code));

                            IList<Point> right = new List<Point>();

                            IList<Point> left = new List<Point>();

                            IList<Point> none = new List<Point>();

                            foreach (XmlNode assembly in alignmentXml.SelectNodes(".//lx:CrossSect", nsmgr))
                            {
                                double station = Convert.ToDouble(assembly.Attributes["sta"].Value, System.Globalization.CultureInfo.InvariantCulture);

                                Log(string.Format("Processing Station {0}...", station));

                                if (Math.Abs(station - start) < 0.001)
                                {
                                    station = start;
                                }
                                if (Math.Abs(station - end) < 0.001)
                                {
                                    station = end;
                                }

                                if (station >= start && station <= end)
                                {
                                    cs = b.CoordinateSystemByStation(station);

                                    foreach (XmlNode subassembly in assembly.SelectNodes("lx:DesignCrossSectSurf[@side = 'left' and @closedArea]", nsmgr))
                                    {
                                        Log(string.Format("Processing Subassembly {0} points...", subassembly.ChildNodes.Count));

                                        foreach (XmlNode calcPoint in subassembly.SelectNodes(string.Format("lx:CrossSectPnt[@code = '{0}']", code), nsmgr))
                                        {
                                            string[] coords = calcPoint.InnerText.Split(separator, StringSplitOptions.None);

                                            left.Add(Point.ByCoordinates(Convert.ToDouble(coords[0], System.Globalization.CultureInfo.InvariantCulture), 0, Convert.ToDouble(coords[1], System.Globalization.CultureInfo.InvariantCulture)).Transform(cs) as Point);

                                            Log(string.Format("Processing Coordinates Left...", ""));
                                        }
            
                                    }

                                    foreach (XmlNode subassembly in assembly.SelectNodes("lx:DesignCrossSectSurf[@side = 'right' and @closedArea]", nsmgr))
                                    {
                       
                                        foreach (XmlNode calcPoint in subassembly.SelectNodes(string.Format("lx:CrossSectPnt[@code = '{0}']", code), nsmgr))
                                        {
                                            string[] coords = calcPoint.InnerText.Split(separator, StringSplitOptions.None);

                                            right.Add(Point.ByCoordinates(Convert.ToDouble(coords[0], System.Globalization.CultureInfo.InvariantCulture), 0, Convert.ToDouble(coords[1], System.Globalization.CultureInfo.InvariantCulture)).Transform(cs) as Point);

                                            Log(string.Format("Processing Coordinates Right...", ""));
                                        }
                            
                                    }

                                    foreach (XmlNode calcPoint in assembly.SelectNodes(string.Format(".//lx:CrossSectPnt[@code = '{0}']", code), nsmgr))
                                    {
                                        string[] coords = calcPoint.InnerText.Split(separator, StringSplitOptions.None);

                                        none.Add(Point.ByCoordinates(Convert.ToDouble(coords[0], System.Globalization.CultureInfo.InvariantCulture), 0, Convert.ToDouble(coords[1], System.Globalization.CultureInfo.InvariantCulture)).Transform(cs) as Point);

                                        Log(string.Format("Processing Coordinates Centered...", ""));
                                    }
                                }
                            }

                            if (left.Count > 1)
                            {

                                left = left.OrderBy(p => b.GetArrayStationOffsetElevationByPoint(p)[0]).ToList();
                                if (left.Count > 1)
                                {
                                    featurelines.Add(new Featureline(b, PolyCurve.ByPoints(left), code, Featureline.SideType.Left));

                                    Log(string.Format("Left Featureline Created!", ""));
                                }
                            }

                            if (right.Count > 1)
                            {
                                right = Point.PruneDuplicates(right);
                                right = right.OrderBy(p => b.GetArrayStationOffsetElevationByPoint(p)[0]).ToList();
                                if (right.Count > 1)
                                {
                                    featurelines.Add(new Featureline(b, PolyCurve.ByPoints(right), code, Featureline.SideType.Right));

                                    Log(string.Format("Right Featureline Created!", ""));
                                }
                            }

                            if (none.Count > 1)
                            {
                                none = Point.PruneDuplicates(none);
                                none = none.OrderBy(p => b.GetArrayStationOffsetElevationByPoint(p)[0]).ToList();

                                var pc = PolyCurve.ByPoints(none);
                                var offset = b.GetArrayStationOffsetElevationByPoint(pc.PointAtParameter(0.5))[1];
                                var side = Featureline.SideType.Right;

                                if (offset < 0)
                                {
                                    side = Featureline.SideType.Left;
                                }

                                if (none.Count > 1)
                                {
                                    featurelines.Add(new Featureline(b, pc, code, side));

                                    Log(string.Format("Featureline Created!", ""));
                                }
                            }
                        }

                        baselineFeaturelines.Add(featurelines);

                        Log(string.Format("Region Featurelines added!", ""));
                    }
                }

                corridorFeaturelines.Add(baselineFeaturelines);

                Log(string.Format("Baseline Featurelines added!", ""));
            }

            cs.Dispose();

            Utils.Log(string.Format("Utils.GetFeaturelinesFromLandXML completed.", ""));

            return corridorFeaturelines;
        }

        /// <summary>
        /// Gets the featurelies from the corridor organized by Corridor-Baseline-Code-Region
        /// </summary>
        /// <param name="corridor">The corridor.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<IList<IList<Featureline>>> GetFeaturelines(Corridor corridor)  // 20190125
        {
            Utils.Log(string.Format("Utils.GetFeaturelines started...", ""));

            IList<IList<IList<Featureline>>> corridorFeaturelines = new List<IList<IList<Featureline>>>();

            IList<string> codes = corridor.GetCodes();

            foreach (Baseline b in corridor.Baselines)
            {
                foreach (string code in codes)
                {
                    corridorFeaturelines.Add(b.GetFeaturelinesByCode(code));
                }
            }

            corridor._corridorFeaturelinesXMLExported = true;

            Utils.Log(string.Format("Utils.GetFeaturelines completed.", ""));

            return corridorFeaturelines;
        }

        /// <summary>
        /// Gets the triangles from a CivilSurface
        /// </summary>
        /// <param name="surface">The CivilSurface.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<Surface> GetSurfaceTriangles(CivilSurface surface)  // 20190922
        {
            Utils.Log(string.Format("Utils.GetSurfaceTriangles ({0}) started...", surface.Name));

            IList<Surface> triangles = new List<Surface>();

            string xmlPath = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "Surface.xml");  // Revit 2020 changed the path to the temp at a session level

            Utils.Log(xmlPath);

            surface.InternalElement.Document.SendCommand(string.Format("-ExportSurfaceToXML\n{0}\n", surface.InternalElement.Handle));

            DateTime start = DateTime.Now;

            while (true)
            {
                if (File.Exists(xmlPath))
                {
                    if (File.GetLastWriteTime(xmlPath) > start)
                    {
                        start = File.GetLastWriteTime(xmlPath);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            Utils.Log("XML acquired.");

            if (File.Exists(xmlPath))
            {
                IList<Featureline> output = new List<Featureline>();

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlPath);

                foreach (XmlElement se in xmlDoc.GetElementsByTagName("Surface")
                    .Cast<XmlElement>()
                    .Where(x => x.Attributes["Name"].Value == surface.Name))
                {
                    Dictionary<string, Point> points = new Dictionary<string, Point>();

                    foreach (XmlElement ve in se.GetElementsByTagName("Vertex"))
                    {
                        double x = 0;
                        double y = 0;
                        double z = 0;
                        string id = "";

                        try
                        {
                            x = Convert.ToDouble(ve.Attributes["X"].Value, System.Globalization.CultureInfo.InvariantCulture);
                        }
                        catch (Exception ex)
                        {
                            Utils.Log(string.Format("ERROR: X {0}", ex.Message));
                        }

                        try
                        {
                            y = Convert.ToDouble(ve.Attributes["Y"].Value, System.Globalization.CultureInfo.InvariantCulture);
                        }
                        catch (Exception ex)
                        {
                            Utils.Log(string.Format("ERROR: Y {0}", ex.Message));
                        }

                        try
                        {
                            z = Convert.ToDouble(ve.Attributes["Z"].Value, System.Globalization.CultureInfo.InvariantCulture);
                        }
                        catch (Exception ex)
                        {
                            Utils.Log(string.Format("ERROR: Z {0}", ex.Message));
                        }

                        try
                        {
                            id = ve.Attributes["id"].Value;
                        }
                        catch (Exception ex)
                        {
                            Utils.Log(string.Format("ERROR: Id {0}", ex.Message));
                        }

                        if (!string.IsNullOrEmpty(id) && !string.IsNullOrWhiteSpace(id))
                        {
                            points.Add(id, Point.ByCoordinates(x, y, z));
                        }
                    }

                    Point a = null;
                    Point b = null;
                    Point c = null;

                    // create surfaces
                    foreach (XmlElement te in se.GetElementsByTagName("Triangle"))
                    {
                        a = points[te.Attributes["V0"].Value];
                        b = points[te.Attributes["V1"].Value];
                        c = points[te.Attributes["V2"].Value];

                        triangles.Add(Surface.ByPerimeterPoints(new Point[] { a, b, c }));
                    }

                    a.Dispose();
                    b.Dispose();
                    c.Dispose();
                }
            }

            Utils.Log(string.Format("Utils.GetSurfaceTriangles completed.", ""));

            return triangles;
        }


        /// <summary>
        /// Gets the triangles from a CivilSurface
        /// </summary>
        /// <param name="surface">The CivilSurface.</param>
        /// <param name="path">The path to the LandXML that contains the surface export.</param>
        /// <param name="onlyVisible">If true processes the visible triangles in the XML.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<Surface> GetSurfaceTrianglesByLandXML(CivilSurface surface, string path, bool onlyVisible = true)  // 20191008
        {
            Utils.Log(string.Format("Utils.GetSurfaceTrianglesByLandXML ({0}) started...", surface.Name));

            IList<Surface> triangles = new List<Surface>();

            string xmlPath = path;

            Utils.Log(xmlPath);

            if (File.Exists(xmlPath))
            {
               
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlPath);

                foreach (XmlElement se in xmlDoc.GetElementsByTagName("Surface")
                    .Cast<XmlElement>()
                    .Where(x => x.Attributes["name"].Value == surface.Name))
                {
                    Dictionary<string, Point> points = new Dictionary<string, Point>();

                    foreach (XmlElement ve in se.GetElementsByTagName("P"))
                    {
                        string[] inner = ve.InnerText.Split(new string[] { " " }, StringSplitOptions.None);

                        double x = Convert.ToDouble(inner[1], System.Globalization.CultureInfo.InvariantCulture);
                        double y = Convert.ToDouble(inner[0], System.Globalization.CultureInfo.InvariantCulture);
                        double z = Convert.ToDouble(inner[2], System.Globalization.CultureInfo.InvariantCulture);
                        string id = "";

                        try
                        {
                            id = ve.Attributes["id"].Value;
                        }
                        catch (Exception ex)
                        {
                            Utils.Log(string.Format("ERROR: Id {0}", ex.Message));
                        }

                        if (!string.IsNullOrEmpty(id) && !string.IsNullOrWhiteSpace(id))
                        {
                            try
                            {
                                points.Add(id, Point.ByCoordinates(x, y, z));
                            }
                            catch (Exception ex)
                            {
                                Utils.Log(string.Format("ERROR: Id {0}", ex.Message));
                            }
                        }
                    }

                    Point a = null;
                    Point b = null;
                    Point c = null;

                    // create surfaces
                    foreach (XmlElement te in se.GetElementsByTagName("F"))
                    {
                        if (onlyVisible)
                        {
                            // Process only visible faces
                            if (te.HasAttribute("i"))
                            {
                                if (te.Attributes["i"].Value == "1")
                                {
                                    continue;
                                }
                            }
                        }

                        string[] inner = te.InnerText.Split(new string[] { " " }, StringSplitOptions.None);

                        a = points[inner[0]];
                        b = points[inner[1]];
                        c = points[inner[2]];

                        triangles.Add(Surface.ByPerimeterPoints(new Point[] { a, b, c }));
                    }

                    a.Dispose();
                    b.Dispose();
                    c.Dispose();
                }
            }

            Utils.Log(string.Format("Utils.GetSurfaceTrianglesByLandXML completed.", ""));

            return triangles;
        }


        /// <summary>
        /// Gets the triangles from a CivilSurface
        /// </summary>
        /// <param name="surface">The CivilSurface.</param>
        /// <param name="path">The path to the LandXML that contains the surface export.</param>
        /// <param name="onlyVisible">If true processes the visible triangles in the XML.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static Dictionary<string, object> GetFacesLandXML(CivilSurface surface, string path, bool onlyVisible = true)  // 20191125
        {
            Utils.Log(string.Format("Utils.GetFacesLandXML ({0}) started...", surface.Name));

            IList<Surface> triangles = new List<Surface>();

            string xmlPath = path;

            Dictionary<string, Point> points = new Dictionary<string, Point>();
            IList<IList<int>> faces = new List<IList<int>>();

            if (File.Exists(xmlPath))
            {
               
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlPath);

                Utils.Log(xmlPath);

                string ns = "http://www.landxml.org/schema/LandXML-1.2";

                foreach (XmlElement se in xmlDoc.GetElementsByTagName("Surface")
                    .Cast<XmlElement>()
                    .Where(x => x.Attributes["name"].Value == surface.Name))
                {
                    Utils.Log("Surface found");

                    foreach (XmlElement ve in se.GetElementsByTagName("P"))
                    {
                        string[] inner = ve.InnerText.Split(new string[] { " " }, StringSplitOptions.None);

                        double x = Convert.ToDouble(inner[1], System.Globalization.CultureInfo.InvariantCulture);
                        double y = Convert.ToDouble(inner[0], System.Globalization.CultureInfo.InvariantCulture);
                        double z = Convert.ToDouble(inner[2], System.Globalization.CultureInfo.InvariantCulture);
                        string id = "";

                        Utils.Log(ve.InnerText);

                        try
                        {
                            id = ve.Attributes["id"].Value;
                        }
                        catch (Exception ex)
                        {
                            Utils.Log(string.Format("ERROR: Id {0}", ex.Message));
                        }

                        if (!string.IsNullOrEmpty(id) && !string.IsNullOrWhiteSpace(id))
                        {
                            try
                            {
                                points.Add(id, Point.ByCoordinates(x, y, z));
                            }
                            catch (Exception ex)
                            {
                                Utils.Log(string.Format("ERROR: Id {0}", ex.Message));
                            }
                        }
                    }


                    // create surfaces
                    foreach (XmlElement te in se.GetElementsByTagName("F"))
                    {
                        if (onlyVisible)
                        {
                            // Process only visible faces
                            if (te.HasAttribute("i"))
                            {
                                if (te.Attributes["i"].Value == "1")
                                {
                                    continue;
                                }
                            }
                        }

                        string[] inner = te.InnerText.Split(new string[] { " " }, StringSplitOptions.None);

                        IList<int> indices = inner.Select(i => Convert.ToInt32(i)).ToList();

                        faces.Add(indices);
                    }

                    break;
                }
            }

            Utils.Log(string.Format("Utils.GetFacesLandXML completed.", ""));

            return new Dictionary<string, object>() { {"Points", points}, {"Faces", faces}};
        }

        /// <summary>
        /// Recursive function to join surfaces into a PolySurface
        /// </summary>
        /// <param name="surfaces">The surface list to process</param>
        /// <param name="limit">The amount of surfaces to join together</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<Surface> JoinSurfaces(IList<Surface> surfaces, int limit = 100)  // 20190922
        {
            Utils.Log(string.Format("Utils.JoinSurfaces started on {0} surfaces", surfaces.Count));

            if (surfaces.Count == 1)
            {
                Utils.Log(string.Format("Utils.JoinSurfaces completed.", ""));

                return surfaces;
            }
            else
            {
                IList<Surface> result = new List<Surface>();

                for (int i = 0; i < surfaces.Count; i = i + limit)
                {
                    IList<Surface> temp = new List<Surface>();

                    for (int j = i; j < i + limit; ++j)
                    {
                        if (j < surfaces.Count)
                        {
                            temp.Add(surfaces[j]);
                        }
                        else
                        {
                            break;
                        }
                    }

                    PolySurface ps = PolySurface.ByJoinedSurfaces(temp);

                    result.Add(ps);
                }

                result = JoinSurfaces(result);

                Utils.Log(string.Format("Utils.JoinSurfaces completed.", ""));

                return result;
            }
        }

        // TODO : Create a set of nodes to process directly LandXML files to extract:
        // Surfaces
        // Alignments
        // Corridors
        // Pipe Networks

        #endregion
    }
}
