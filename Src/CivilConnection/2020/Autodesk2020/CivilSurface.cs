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
        public AeccSurfaceType SurfaceType { get { return _surface.Type; } }
        #endregion

        #region INTERNAL CONSTRUCTOR
        /// <summary>
        /// Gets the internal element
        /// </summary>
        internal object Internal { get { return this._surface; } }
        /// <summary>
        /// Internal constructor
        /// </summary>
        /// <param name="surface">the internal AeccSurface</param>
        internal CivilSurface(AeccSurface surface)
        {
            this._surface = surface;
            this._boundaries = surface.Boundaries;
        }
        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// 
        /// </summary>
        /// <param name="points"></param>
        /// <returns></returns>
        public List<object> GetElevationAtPoint(List<Point> points)
        {
            List<object> elevations = new List<object>();
            foreach (Point point in points)
            {
                try
                {
                    double elevation = Math.Round(this._surface.FindElevationAtXY(point.X, point.Y), 3);
                    elevations.Add(elevation);
                }

                catch
                {
                    elevations.Add("The surface elevation at selected point is not found");
                }

            }
            return elevations;
        }

        /// <summary>
        /// Gets all surface points along line
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public List<List<object>> GetElevationsAlongLine(List<Line> lines)
        {
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
            return pointsOnLine;
        }

        /// <summary>
        /// Get all surface points
        /// </summary>
        /// <returns>
        /// Point(X,Y,Z)
        /// </returns>
        public List<Point> GetSurfacePoints()
        {
            List<Point> pointList = new List<Point>();
            double[] points = this._surface.Points;

            for (int i = 0; i < points.Length; i += 3)
            {
                Point p = Point.ByCoordinates(points[i], points[i + 1], points[i + 2]);
                pointList.Add(p);
            }

            return pointList;
        }

        /// <summary>
        /// Gets all surface points between lower and upper limit point
        /// </summary>
        /// <param name="LowerLeftPoint"></param>
        /// <param name="UpperRightPoint"></param>
        /// <returns>
        /// Point(X, Y, Z)
        /// </returns>
        public List<Point> GetPointsBetweenLowerUpperLimits(Point LowerLeftPoint, Point UpperRightPoint)
        {
            List<Point> points = new List<Point>();
            double minX = LowerLeftPoint.X;
            double minY = LowerLeftPoint.Y;
            double maxX = UpperRightPoint.X;
            double maxY = UpperRightPoint.Y;

            List<Point> surfacePoints = this.GetSurfacePoints();

            foreach (Point point in surfacePoints)
            {
                if (point.X >= minX && point.Y >= minY && point.X <= maxX && point.Y <= maxY)
                {
                    points.Add(point);
                }
            }

            return points;
        }

        /// <summary>
        /// Get surface points inside and along periphery of closed polygon
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public List<Point> GetPointsInsidePolygon(List<Line> lines)
        {
            List<Point> pointList = new List<Point>();
            List<Point> pointsInside = new List<Point>();
            foreach (Line line in lines)
            {
                Point startPoint = line.StartPoint;
                pointList.Add(startPoint);

            }
            double minX = pointList[0].X;
            double minY = pointList[0].Y;
            double maxX = pointList[0].X;
            double maxY = pointList[0].Y;

            for (int i = 0; i < pointList.Count; i++)
            {
                if (pointList[i].X < minX)
                {
                    minX = pointList[i].X;
                }
                if (pointList[i].Y < minY)
                {
                    minY = pointList[i].Y;
                }
                if (pointList[i].X > maxX)
                {
                    maxX = pointList[i].X;
                }
                if (pointList[i].Y > maxY)
                {
                    maxY = pointList[i].Y;
                }
            }
            Point minPoint = Point.ByCoordinates(minX, minY);
            Point maxPoint = Point.ByCoordinates(maxX, maxY);

            List<Point> pointsBounding = GetPointsBetweenLowerUpperLimits(minPoint, maxPoint);

            foreach (Point p in pointsBounding)
            {
                Point p1, p2;

                bool inside = false;

                if (pointList.Count < 3)
                {
                    pointsInside.Add(null);
                }
                else
                {
                    Point oldPoint = Point.ByCoordinates(pointList[pointList.Count - 1].X, pointList[pointList.Count - 1].Y);

                    for (int j = 0; j < pointList.Count; j++)
                    {
                        Point newPoint = Point.ByCoordinates(pointList[j].X, pointList[j].Y);

                        if (newPoint.X > oldPoint.X)
                        {
                            p1 = oldPoint;
                            p2 = newPoint;
                        }
                        else
                        {
                            p1 = newPoint;
                            p2 = oldPoint;
                        }

                        if ((newPoint.X < p.X) == (p.X <= oldPoint.X) && (p.Y - (long)p1.Y) * (p2.X - p1.X) < (p2.Y - (long)p1.Y) * (p.X - p1.X))
                        {
                            inside = !inside;
                        }

                        oldPoint = newPoint;

                    }
                    if (inside == true)
                    {
                        pointsInside.Add(p);
                    }
                }

            }


            return pointsInside;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return string.Format("Surface(Name = {0})", this._surface.Name);
        }

        #endregion
    }
}