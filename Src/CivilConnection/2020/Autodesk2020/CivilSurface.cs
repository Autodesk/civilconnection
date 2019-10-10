﻿// Copyright (c) 2019 Autodesk, Inc. All rights reserved.
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
                    // elevations.Add("The surface elevation at selected point is not found.");
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
        public List<List<object>> GetElevationsAlongLine(List<Line> lines)  // author: Atul Tegar
        {
            Utils.Log(string.Format("CivilSurface.GetElevationsAlongLine started...", ""));

            List<List<object>> pointsOnLine = new List<List<object>>();

            foreach (Line line in lines)
            {
                List<object> pointList = new List<object>();

                Point startPoint = line.StartPoint;

                Point endPoint = line.EndPoint;

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

            Utils.Log(string.Format("CivilSurface.GetElevationsAlongLine completed.", ""));

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
        public List<Point> GetPointsInBoundary(Curve boundary, double tolerance=0.1)
        {
            Utils.Log(string.Format("CivilSurface.GetPointsInBoundary started...", ""));

            if (!boundary.IsClosed)
            {
                string message = "The Curve provided is not closed.";

                Utils.Log(string.Format("ERROR: CivilSurface.GetPointsInBoundary {0}", message));

                throw new Exception(message);
            }

            if (tolerance <= 0 || tolerance > 1)
            {
                tolerance = 1;
            }

            Polygon polygon = null;

            List<Point> pointsInside = new List<Point>();

            if (boundary is PolyCurve)
            {
                // polygon = Polygon.ByPoints(((PolyCurve)boundary).Curves().Select(c => c.StartPoint).Select(p => Point.ByCoordinates(p.X, p.Y)));

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
                }

                polygon = Polygon.ByPoints(points.Select(p => Point.ByCoordinates(p.X, p.Y)));

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

             polygon.Dispose();

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
        /// Joins the surfaces recursively into a Polysurface
        /// </summary>
        /// <param name="surfaces">The surfaces to join</param>
        /// <param name="limit">The amount of surfaces to join with recursion</param>
        /// <returns></returns>
        public static IList<Surface> JoinSurfaces(IList<Surface> surfaces, int limit=100)
        {
            return Utils.JoinSurfaces(surfaces, limit);
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