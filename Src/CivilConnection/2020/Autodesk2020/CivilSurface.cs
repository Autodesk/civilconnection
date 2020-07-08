// Copyright (c) 2019 Autodesk, Inc. All rights reserved.
// Authors: Atul Tegar - ATTE@cowi.com, paolo.serra@autodesk.com
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
using Autodesk.AECC.Interop.Land;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.DesignScript.Runtime;
using Autodesk.DesignScript.Geometry;

namespace CivilConnection
{
    /// <summary>
    /// Civil 3D Surface object
    /// </summary>
    public class CivilSurface
    {
        #region PRIVATE PROPERTIES
        private AeccSurface _surface;
        private AeccSurfaceBoundaries _boundaries;
        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the name of surface
        /// </summary>
        /// <value>
        /// Name of surface
        /// </value>
        public string Name { get { return _surface.Name; } }
        /// <summary>
        /// Gets the surface type
        /// </summary>
        public string SurfaceType { get { return _surface.Type.ToString(); } }
        #endregion

        #region INTERNAL CONSTRUCTOR
        /// <summary>
        /// Gets the internal element
        /// </summary>
        internal AeccSurface InternalElement { get { return this._surface; } }
        /// <summary>
        /// Internal constructor
        /// </summary>
        /// <param name="surface">the internal AeccSurface</param>
        internal CivilSurface(AeccSurface surface)
        {
            this._surface = surface;
            this._boundaries = surface.Boundaries;  // What to do with boudanries?
        }
        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// Get elevation of surface at points
        /// </summary>
        /// <param name="points">The points to process</param>
        /// <returns>
        /// The List of Elevations
        /// </returns>
        public List<object> GetElevationAtPoint(List<Point> points)  // author: Atul Tegar
        {
            Utils.Log(string.Format("CivilSurface.GetElevationAtPoint started...", ""));

            List<object> elevations = new List<object>();

            foreach (Point point in points)
            {
                try
                {
                    double elevation = Math.Round(this._surface.FindElevationAtXY(point.X, point.Y), 5);

                    elevations.Add(elevation);
                }

                catch (Exception ex)
                {
                    Utils.Log(string.Format("ERROR: CivilSurface.GetElevationAtPoint: The surface elevation at selected point is not found, {0}", ex.Message));

                    elevations.Add(null);
                }
            }

            Utils.Log(string.Format("CivilSurface.GetElevationAtPoint completed.", ""));

            return elevations;
        }

        /// <summary>
        /// Gets all surface points along line
        /// </summary>
        /// <param name="lines">The lines to process</param>
        /// <returns>
        /// The List of Points
        /// </returns>
        public List<List<Point>> GetElevationsAlongLine(List<Line> lines)  // author: Atul Tegar
        {
            Utils.Log(string.Format("CivilSurface.GetElevationsAlongLine started...", ""));

            List<List<Point>> pointsOnLine = new List<List<Point>>();

            Point startPoint = null;

            Point endPoint = null;

            foreach (Line line in lines)
            {
                List<Point> pointList = new List<Point>();

                startPoint = line.StartPoint;

                endPoint = line.EndPoint;

                double[] surfacePoints = this._surface.SampleElevations(startPoint.X, startPoint.Y, endPoint.X, endPoint.Y);

                for (int i = 0; i < surfacePoints.Length; i += 3)
                {
                    double pX = surfacePoints[i];

                    double pY = surfacePoints[i + 1];

                    double pZ = surfacePoints[i + 2];

                    Point point = Point.ByCoordinates(pX, pY, pZ);

                    pointList.Add(point);
                }

                pointsOnLine.Add(pointList);
            }

            startPoint.Dispose();
            endPoint.Dispose();

            Utils.Log(string.Format("CivilSurface.GetElevationsAlongLine completed.", ""));

            return pointsOnLine;
        }

        /// <summary>
        /// Gets all surface points along line
        /// </summary>
        /// <param name="lines">The lines to process</param>
        /// <returns>
        /// The List of Points
        /// </returns>
        public List<List<Point>> GetPointsAlongLine(List<Line> lines)  // Renamed as it confuses people
        {
            Utils.Log(string.Format("CivilSurface.GetPointsAlongLine started...", ""));

            List<List<Point>> pointsOnLine = new List<List<Point>>();

            Point startPoint = null;

            Point endPoint = null;

            foreach (Line line in lines)
            {
                List<Point> pointList = new List<Point>();

                startPoint = line.StartPoint;

                endPoint = line.EndPoint;

                double[] surfacePoints = this._surface.SampleElevations(startPoint.X, startPoint.Y, endPoint.X, endPoint.Y);

                for (int i = 0; i < surfacePoints.Length; i += 3)
                {
                    double pX = surfacePoints[i];

                    double pY = surfacePoints[i + 1];

                    double pZ = surfacePoints[i + 2];

                    Point point = Point.ByCoordinates(pX, pY, pZ);

                    pointList.Add(point);
                }

                pointsOnLine.Add(pointList);
            }

            startPoint.Dispose();
            endPoint.Dispose();

            Utils.Log(string.Format("CivilSurface.GetPointsAlongLine completed.", ""));

            return pointsOnLine;
        }

        /// <summary>
        /// Get all surface points
        /// </summary>
        /// <returns>
        /// The List of Points
        /// </returns>
        public List<Point> GetSurfacePoints()  // author: Atul Tegar
        {
            Utils.Log(string.Format("CivilSurface.GetSurfacePoints started...", ""));

            List<Point> pointList = new List<Point>();

            double[] points = this._surface.Points;

            for (int i = 0; i < points.Length; i += 3)
            {
                Point p = Point.ByCoordinates(points[i], points[i + 1], points[i + 2]);

                pointList.Add(p);
            }

            Utils.Log(string.Format("CivilSurface.GetSurfacePoints completed.", ""));

            return pointList;
        }

        /// <summary>
        /// Gets all surface points between lower and upper limit points.
        /// </summary>
        /// <param name="lowerLeftPoint">The minmum point.</param>
        /// <param name="upperRightPoint">The maximum pont.</param>
        /// <returns>
        /// The List of Points
        /// </returns>
        private List<Point> GetPointsBetweenLowerUpperLimits(Point lowerLeftPoint, Point upperRightPoint)  // author: Atul Tegar
        {
            Utils.Log(string.Format("CivilSurface.GetPointsBetweenLowerUpperLimits started...", ""));

            List<Point> points = new List<Point>();

            double minX = lowerLeftPoint.X;
            double minY = lowerLeftPoint.Y;

            double maxX = upperRightPoint.X;
            double maxY = upperRightPoint.Y;

            List<Point> surfacePoints = this.GetSurfacePoints();

            foreach (Point point in surfacePoints)
            {
                if (point.X >= minX && point.Y >= minY && point.X <= maxX && point.Y <= maxY)
                {
                    points.Add(point);
                }
            }

            Utils.Log(string.Format("CivilSurface.GetPointsBetweenLowerUpperLimits completed.", ""));

            return points;
        }

        /// <summary>
        /// Gets all surface points in the BoundingBox.
        /// </summary>
        /// <param name="boundingBox">The BoundingBox used for the containment test.</param>
        /// <returns>
        /// The List of Points
        /// </returns>
        public List<Point> GetPointsInBoundingBox(BoundingBox boundingBox)
        {
            Utils.Log(string.Format("CivilSurface.GetPointsInBoundingBox started...", ""));

            List<Point> points = this.GetSurfacePoints().Where(p => boundingBox.Contains(p)).ToList();

            Utils.Log(string.Format("CivilSurface.GetPointsInBoundingBox completed.", ""));

            return points;
        }

        /// <summary>
        /// Get surface points inside a closed boundary
        /// </summary>
        /// <param name="boundary">A closed curve</param>
        /// <param name="tolerance">A value between 0 and 1 to define the precision of the tessellation along non-straight curves.</param>
        /// <returns>
        /// The List of Points
        /// </returns>
        public List<Point> GetPointsInBoundary(Curve boundary, double tolerance = 0.1)
        {
            Utils.Log(string.Format("CivilSurface.GetPointsInBoundary started...", ""));

            if (!boundary.IsClosed)
            {
                string message = "The Curve provided is not closed.";

                Utils.Log(string.Format("ERROR: CivilSurface.GetPointsInBoundary {0}", message));

                throw new Exception(message);
            }

            Plane xy = Plane.XY();

            Curve boundaryXY = boundary.PullOntoPlane(xy);

            Surface surf = null;

            List<Point> pointsInside = new List<Point>();

            Polygon polygon = null;

            try
            {
                surf = Surface.ByPatch(boundaryXY);
            }
            catch 
            {
                string message = "Unable to create surface form boundary.";

                Utils.Log(string.Format("ERROR: CivilSurface.GetPointsInBoundary {0}", message));
            }

            if (surf != null)
            {
                pointsInside = this.GetSurfacePoints().Where(p => Point.ByCoordinates(p.X, p.Y).DoesIntersect(surf)).ToList();
            }
            else
            {
                if (tolerance <= 0 || tolerance > 1)
                {
                    tolerance = 1;
                }

                if (boundary is PolyCurve)
                {
                   
                    PolyCurve pc = boundary as PolyCurve;

                    List<Point> points = new List<Point>();

                    for (int i = 0; i < pc.NumberOfCurves; ++i)
                    {
                        Curve c = pc.CurveAtIndex(i);

                        try
                        {
                            Line line = c as Line;
                            points.Add(c.StartPoint);
                        }
                        catch
                        {
                            for (double j = 0; j < 1; j = j + tolerance)
                            {
                                points.Add(c.PointAtParameter(j));
                            }
                        }

                        c.Dispose();
                    }

                    polygon = Polygon.ByPoints(points.Select(p => Point.ByCoordinates(p.X, p.Y)));

                    foreach (var item in points)
                    {
                        if (item != null)
                        {
                            item.Dispose();
                        }
                    }

                    points.Clear();
                }
                else if (boundary is Circle)
                {
                    List<Point> points = new List<Point>();

                    for (double j = 0; j < 1; j = j + tolerance)
                    {
                        points.Add(boundary.PointAtParameter(j));
                    }

                    polygon = Polygon.ByPoints(points.Select(p => Point.ByCoordinates(p.X, p.Y)));

                    foreach (var item in points)
                    {
                        if (item != null)
                        {
                            item.Dispose();
                        }
                    }

                    points.Clear();

                }
                else if (boundary is Rectangle)
                {
                    polygon = Polygon.ByPoints(((Rectangle)boundary).Curves().Select(c => c.StartPoint).Select(p => Point.ByCoordinates(p.X, p.Y)));
                }
                else if (boundary is Polygon)
                {
                    polygon = boundary as Polygon;
                }

                pointsInside = this.GetSurfacePoints().Where(p => polygon.ContainmentTest(Point.ByCoordinates(p.X, p.Y))).ToList();
            }

            if (polygon != null)
            {
                polygon.Dispose(); 
            }

            xy.Dispose();

            if (boundaryXY != null)
            {
                boundaryXY.Dispose(); 
            }

            if (surf != null)
            {
                surf.Dispose();
            }

            Utils.Log(string.Format("CivilSurface.GetPointsInBoundary completed.", ""));

            return pointsInside;
        }

        /// <summary>
        /// Gets all the triangle surfaces in a CivilSurface
        /// </summary>
        /// <returns></returns>
        public IList<Surface> GetTrianglesSurfaces()
        {
            return Utils.GetSurfaceTriangles(this);
        }

        /// <summary>
        /// Gets all the triangle surfaces in a CivilSurface via LandXML
        /// </summary>
        /// <param name="landXMLpath">The path to the LandXML that contains the surface export</param>
        /// <param name="onlyVisible">Processes only the visible faces</param>
        /// <returns></returns>
        public IList<Surface> GetTrianglesSurfaces(string landXMLpath, bool onlyVisible = true)
        {
            return Utils.GetSurfaceTrianglesByLandXML(this, landXMLpath, onlyVisible);
        }

        /// <summary>
        /// Gets all the triangle faces in a CivilSurface via LandXML
        /// </summary>
        /// <param name="landXMLpath">The path to the LandXML that contains the surface export</param>
        /// <param name="onlyVisible">Processes only the visible faces</param>
        /// <returns></returns>
        [MultiReturn(new string[]{"Points", "Faces"})]
        public Dictionary<string, object> GetFacesSurfaces(string landXMLpath, bool onlyVisible = true)
        {
            return Utils.GetFacesLandXML(this, landXMLpath, onlyVisible);
        }

        /// <summary>
        /// Joins the surfaces recursively into a Polysurface
        /// </summary>
        /// <param name="surfaces">The surfaces to join</param>
        /// <param name="limit">The amount of surfaces to join with recursion</param>
        /// <returns></returns>
        public static IList<Surface> JoinSurfaces(IList<Surface> surfaces, int limit = 100)
        {
            return Utils.JoinSurfaces(surfaces, limit);
        }

        /// <summary>
        /// Get intersection point between the line with start point and direction on the surface 
        /// </summary>
        /// <param name="point">The point to process</param>
        /// <param name="vector">The direction vector</param>
        /// <returns>
        /// The intersection point
        /// </returns>
        public Point GetIntersectionPoint(Point point, Vector vector) // author: Atul Tegar
        {
            Utils.Log(string.Format("CivilSurface.GetIntersectionPoint Started...", ""));

            Point point2 = null;
            double p1X = point.X;
            double p1Y = point.Y;
            double p1Z = point.Z;

            double[] p1 = { p1X, p1Y, p1Z };

            double vX = vector.X;
            double vY = vector.Y;
            double vZ = vector.Z;

            double[] direction = { vX, vY, vZ };

            try
            {
                var intPoint = this._surface.IntersectPointWithSurface(p1, direction);

                double p2X = intPoint[0];
                double p2Y = intPoint[1];
                double p2Z = intPoint[2];
                point2 = Point.ByCoordinates(p2X, p2Y, p2Z);

            }

            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: CivilSurface.GetIntersectionPoint: No intersection with vector and surface found, {0}", ex.Message));
                point2 = null;
            }

            Utils.Log(string.Format("CivilSurface.GetIntersectionPoint completed.", ""));

            return point2;
        }


        /// <summary>
        /// Public textual representation in the Dynamo node preview.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("CivilSurface(Name = {0})", this.Name);
        }

        #endregion
    }
}