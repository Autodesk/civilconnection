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

namespace CivilConnection
{
    /// <summary>
    /// Alignment object type.
    /// </summary>
    public class Alignment
    {
        #region PRIVATE PROPERTIES
        private AeccAlignment _alignment;
        private AeccAlignmentEntities _entities;
        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get { return _alignment.DisplayName; } }
        /// <summary>
        /// Gets the length.
        /// </summary>
        /// <value>
        /// The length.
        /// </value>
        public double Length { get { return _alignment.Length; } }
        /// <summary>
        /// Gets the start.
        /// </summary>
        /// <value>
        /// The start.
        /// </value>
        public double Start { get { return _alignment.StartingStation; } }
        /// <summary>
        /// Gets the end.
        /// </summary>
        /// <value>
        /// The end.
        /// </value>
        public double End { get { return _alignment.EndingStation; } }

        /// <summary>
        /// Gets the stations of the geometry points.
        /// </summary>
        /// <value>
        /// The GeometryStations.
        /// </value>
        public double[] GeometryStations { get { return _alignment.GetStations(AeccStationType.aeccGeometryPoint, this.Start, this.End).Cast<AeccAlignmentStation>().Select(x => x.Station).ToArray(); } }

        /// <summary>
        /// Gets the stations of the points of intersection.
        /// </summary>
        /// <value>
        /// The PIStations.
        /// </value>
        public double[] PIStations { get { return _alignment.GetStations(AeccStationType.aeccPIPoint, this.Start, this.End).Cast<AeccAlignmentStation>().Select(x => x.Station).ToArray(); } }

        /// <summary>
        /// Gets the stations of the points of superelevation transition.
        /// </summary>
        /// <value>
        /// The SuperTransStations.
        /// </value>
        public double[] SuperTransStations { get { return _alignment.GetStations(AeccStationType.aeccSuperTransPoint, this.Start, this.End).Cast<AeccAlignmentStation>().Select(x => x.Station).ToArray(); } }

        #endregion

        #region INTERNAL CONSTRUCTORS
        /// <summary>
        /// Gets the internal element.
        /// </summary>
        /// <value>
        /// The internal element.
        /// </value>
        internal object InternalElement { get { return this._alignment; } }

        /// <summary>
        /// Internal constructor.
        /// </summary>
        /// <param name="alignment">The internal AeccAlignment.</param>
        internal Alignment(AeccAlignment alignment)
        {
            this._alignment = alignment;
            this._entities = alignment.Entities;
        }
        #endregion

        #region PUBLIC CONSTRUCTORS
        /// <summary>
        /// Creates an Alignment in the Civil Document starting from a Dynamo polygonal PolyCurve.
        /// </summary>
        /// <param name="civilDocument">The CivilDocument.</param>
        /// <param name="name">The name of the alignment.</param>
        /// <param name="polyCurve">The source PolyCurve.</param>
        /// <param name="layer">The layer.</param>
        /// <returns>
        /// The new Alignment
        /// </returns>
        public static Alignment ByPolygonal(CivilDocument civilDocument, string name, PolyCurve polyCurve, string layer)
        {
            var pl = civilDocument._document.HandleToObject(Utils.AddLWPolylineByPolyCurve(civilDocument._document, polyCurve, "0"));

            Utils.Log(string.Format("Polyline 2D added: {0}", pl.Handle));

            var alignments = civilDocument._document.AlignmentsSiteless;

            AeccAlignmentStyle alignmentStyle = null;
            try
            {
                alignmentStyle = civilDocument._document.AlignmentStyles[0];
            }
            catch
            {
                alignmentStyle = civilDocument._document.AlignmentStyles.Add("CivilConnection_AlignmentStyle");
            }

            Utils.Log(string.Format("Alignment Style: {0}", alignmentStyle));

            AeccAlignmentLabelStyleSet alignmentLabelStyleSet = null;

            try
            {
                alignmentLabelStyleSet = civilDocument._document.AlignmentLabelStyleSets[0];
            }
            catch (Exception)
            {

                alignmentLabelStyleSet = civilDocument._document.AlignmentLabelStyleSets.Add("CivilConnection_AlignmentLabelStyle");
            }

            Utils.Log(string.Format("Alignment Label Style Set: {0}", alignmentLabelStyleSet));

            AeccAlignment al = alignments.AddFromPolylineEx(name, layer, pl, alignmentStyle, alignmentLabelStyleSet, true, false);

            Utils.Log(string.Format("Alignment Created: {0}, {1}", name, al.Handle));

            return new Alignment(al);
        }
        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// Returns the list of vertical Profiles associated to the Alignment.
        /// </summary>
        /// <returns>The list of associated Profiles.</returns>
        public IList<Profile> GetProfiles()
        {
            Utils.Log("Alignment.GetProfiles Started...");

            IList<Profile> output = new List<Profile>();

            foreach (AeccProfile profile in this._alignment.Profiles)
            {
                output.Add(new Profile(profile));
            }

            Utils.Log("Alignment.GetProfiles Completed.");

            return output;
        }

        /// <summary>
        /// Returns the list of vertical ProfileViews associated to the Alignment.
        /// </summary>
        /// <returns>The list of assocaited ProfileViews.</returns>
        public IList<ProfileView> GetProfileViews()
        {
            Utils.Log("Alignment.GetProfileViews Started...");

            IList<ProfileView> output = new List<ProfileView>();

            foreach (AeccProfileView profile in this._alignment.ProfileViews)
            {
                output.Add(new ProfileView(profile));
            }

            Utils.Log("Alignment.GetProfileViews Completed.");

            return output;
        }
        

        /// <summary>
        /// Factorial function. Returns a double to allow for values bigger than 20!
        /// </summary>
        /// <param name="f"></param>
        /// <returns></returns>
        private double Factorial(int f)
        {
            if (f < 0)
            {
                throw new Exception("Factorial is undefined for negative numbers.");
            }

            double output = 1;

            if (f == 0)
            {
                return output;
            }
            else 
            {
                for (int i = 1; i <= f; i++)
                {
                    output *= i;
                }
            }
            
            return output;
        }

        /// <summary>
        /// Returns the list of Dynamo curves that defines the Alignment.
        /// </summary>
        /// <returns>A list of lines and arcs that decribe the Alignment.</returns>
        /// <remarks>The tool returns only lines and arcs.</remarks>
        public IList<Curve> GetCurves()
        {
            Utils.Log("Alignment.GetCurves Started...");

            IList<Curve> output = new List<Curve>();

            foreach (AeccAlignmentEntity e in this._entities)
            {
                switch (e.Type)
                {
                    case AeccAlignmentEntityType.aeccTangent:
                        {
                            AeccAlignmentTangent a = e as AeccAlignmentTangent;

                            output.Add(Line.ByStartPointEndPoint(
                                Point.ByCoordinates(a.StartEasting, a.StartNorthing),
                                Point.ByCoordinates(a.EndEasting, a.EndNorthing)));
                            break;
                        }
                    case AeccAlignmentEntityType.aeccArc:
                        {
                            AeccAlignmentArc a = e as AeccAlignmentArc;

                            Arc arc = null;
                            if (!a.Clockwise)
                            {
                                arc = Arc.ByCenterPointStartPointEndPoint(
                                 Point.ByCoordinates(a.CenterEasting, a.CenterNorthing),
                                 Point.ByCoordinates(a.StartEasting, a.StartNorthing),
                                 Point.ByCoordinates(a.EndEasting, a.EndNorthing));
                            }
                            else 
                            {
                                arc = Arc.ByCenterPointStartPointEndPoint(
                                     Point.ByCoordinates(a.CenterEasting, a.CenterNorthing),
                                     Point.ByCoordinates(a.EndEasting, a.EndNorthing),
                                     Point.ByCoordinates(a.StartEasting, a.StartNorthing));
                            }
                            

                            output.Add(arc);

                            //var p1 = a.PassThroughPoint1;
                            //var p2 = a.PassThroughPoint2;
                            //var p3 = a.PassThroughPoint3;

                            //output.Add(Arc.ByThreePoints(Point.ByCoordinates(p1.X, p1.Y),
                            //   Point.ByCoordinates(p2.X, p2.Y),
                            //   Point.ByCoordinates(p3.X, p3.Y)));
                            break;
                        }
                    case AeccAlignmentEntityType.aeccSpiralCurveSpiralGroup:
                        {
                            AeccAlignmentSCSGroup a = e as AeccAlignmentSCSGroup;

                            AeccAlignmentEntity before = this._entities.Item(e.EntityBefore);
                            AeccAlignmentEntity after = this._entities.Item(e.EntityAfter);

                            Curve beforeCurve = null;
                            Curve afterCurve = null;

                            if (before.Type == AeccAlignmentEntityType.aeccTangent)
                            {
                                AeccAlignmentTangent b = before as AeccAlignmentTangent;

                                beforeCurve = Line.ByStartPointEndPoint(
                                Point.ByCoordinates(b.StartEasting, b.StartNorthing),
                                Point.ByCoordinates(b.EndEasting, b.EndNorthing));
                            }
                            else if (before.Type == AeccAlignmentEntityType.aeccArc)
                            {
                                AeccAlignmentArc b = before as AeccAlignmentArc;

                                beforeCurve = Arc.ByCenterPointStartPointEndPoint(
                                Point.ByCoordinates(b.CenterEasting, b.CenterNorthing),
                                Point.ByCoordinates(b.StartEasting, b.StartNorthing),
                                Point.ByCoordinates(b.EndEasting, b.EndNorthing));
                            }

                            if (after.Type == AeccAlignmentEntityType.aeccTangent)
                            {
                                AeccAlignmentTangent b = after as AeccAlignmentTangent;

                                afterCurve = Line.ByStartPointEndPoint(
                                Point.ByCoordinates(b.StartEasting, b.StartNorthing),
                                Point.ByCoordinates(b.EndEasting, b.EndNorthing));
                            }
                            else if (after.Type == AeccAlignmentEntityType.aeccArc)
                            {
                                AeccAlignmentArc b = after as AeccAlignmentArc;

                                afterCurve = Arc.ByCenterPointStartPointEndPoint(
                                Point.ByCoordinates(b.CenterEasting, b.CenterNorthing),
                                Point.ByCoordinates(b.StartEasting, b.StartNorthing),
                                Point.ByCoordinates(b.EndEasting, b.EndNorthing));
                            }

                            if (afterCurve != null && beforeCurve != null)
                            {
                                output.Add(Curve.ByBlendBetweenCurves(beforeCurve, afterCurve));
                            }

                            // Andrew Milford wrote the code for the spirals using Fourier approximation
                            // Get some parameters
                            int entCnt = a.SubEntityCount;

                            for (int i = 0; i < entCnt; i++)
                            {
                                AeccAlignmentEntity subEnt = a.SubEntity(i);
                                AeccAlignmentEntityType alignType = subEnt.Type;

                                switch (alignType)
                                {
                                    case AeccAlignmentEntityType.aeccTangent:
                                        break;
                                    case AeccAlignmentEntityType.aeccArc:
                                        AeccAlignmentArc arc = subEnt as AeccAlignmentArc;
                                        Arc dArc = null;
                                        if (!arc.Clockwise)
                                        {
                                            dArc = Arc.ByCenterPointStartPointEndPoint(
                                             Point.ByCoordinates(arc.CenterEasting, arc.CenterNorthing),
                                             Point.ByCoordinates(arc.StartEasting, arc.StartNorthing),
                                             Point.ByCoordinates(arc.EndEasting, arc.EndNorthing));
                                        }
                                        else
                                        {
                                            dArc = Arc.ByCenterPointStartPointEndPoint(
                                                 Point.ByCoordinates(arc.CenterEasting, arc.CenterNorthing),
                                                 Point.ByCoordinates(arc.EndEasting, arc.EndNorthing),
                                                 Point.ByCoordinates(arc.StartEasting, arc.StartNorthing));
                                        }

                                        output.Add(dArc);

                                        break;
                                    case AeccAlignmentEntityType.aeccSpiral:

                                        // calculate the spiral intervals
                                        AeccAlignmentSpiral s = subEnt as AeccAlignmentSpiral;
                                        double radIn = s.RadiusIn;
                                        double radOut = s.RadiusOut;
                                        double length = s.Length;
                                        AeccAlignmentSpiralDirectionType dir = s.Direction;
                                        
                                        // Identify the spiral is the start or end
                                        bool isStart = true;
                                        
                                        double radius;
                                        if (Double.IsInfinity(radIn))
                                        {
                                            // Start Spiral
                                            radius = radOut;
                                        }
                                        else
                                        {
                                            // End Spiral
                                            radius = radIn;
                                            isStart = false;
                                        }
                                        
                                        double A = s.A; // Flatness of spiral
                                        double RL = radius * length;
                                        //double A = Math.Sqrt(radius * length);

                                        List<Point> pts = new List<Point>();
                                        double n = 0;
                                        double interval = 2;  // 2 meters intervals TODO what if the curve is shorter than 2 meters?
                                        int num = Convert.ToInt32(Math.Ceiling(length / interval));
                                        double step = length / num;

                                        double dirStart = s.StartDirection;
                                        double dirEnd = s.EndDirection;

                                        Point ptPI = Point.ByCoordinates(s.PIEasting, s.PINorthing, 0);
                                        Point ptStart = Point.ByCoordinates(s.StartEasting, s.StartNorthing, 0);
                                        Point ptEnd = Point.ByCoordinates(s.EndEasting, s.EndNorthing, 0);
                                        Point pt;
                                        double angRot;

                                        /*
                                         * Paolo Serra - Implementation of Fourier's Transform for the clothoid for a given number of terms
                                         x = Sum[(-1)^t * n ^(4*t+1) / ((2*t)!*(4*t+1)*(2*RL)^(2*t))] for t = 0 -> N
                                         y = Sum[(-1)^t * n ^(4*t+3) / ((2^t + 1)!*(4*t+3)*(2*RL)^(2*t + 1))] for t = 0 -> N
                                         */

                                        for (int j = 0; j <= num; j++)
                                        {
                                            double x = 0;
                                            double y = 0;

                                            for (int t = 0; t < 8; t++)  // first 8 terms of the transform
                                            {
                                                x += Math.Pow(-1, t) * Math.Pow(n, 4 * t + 1) / (Factorial(2 * t) * (4 * t + 1) * Math.Pow(2 * RL, 2 * t));
                                                y += Math.Pow(-1, t) * Math.Pow(n, 4 * t + 3) / (Factorial(2 * t + 1) * (4 * t + 3) * Math.Pow(2 * RL, 2 * t + 1));
                                            }

                                            //double x = n - ((Math.Pow(n, 5)) / (40 * Math.Pow(RL, 2))) + ((Math.Pow(n, 9)) / (3456 * Math.Pow(RL, 4))) - ((Math.Pow(n, 13)) / (599040 * Math.Pow(RL, 6))) + ((Math.Pow(n, 17)) / (175472640 * Math.Pow(RL, 8)));
                                            //double y = ((Math.Pow(n, 3)) / (6 * RL)) - ((Math.Pow(n, 7)) / (336 * Math.Pow(RL, 3))) + (Math.Pow(n, 11) / (42240 * Math.Pow(RL, 5))) - (Math.Pow(n, 15) / (9676800 * Math.Pow(RL, 7)));

                                            // Flip the Y offset if start spiral is CW OR end spiral is CCW
                                            if ((isStart && dir == AeccAlignmentSpiralDirectionType.aeccAlignmentSpiralDirectionRight) ||
                                                    (!isStart && dir == AeccAlignmentSpiralDirectionType.aeccAlignmentSpiralDirectionLeft)
                                                )
                                            {
                                                y *= -1;
                                            }

                                            pts.Add(Point.ByCoordinates(x, y, 0));
                                            n += step;
                                        }
                                        
                                        // Create spiral at 0,0
                                        NurbsCurve spiralBase = NurbsCurve.ByControlPoints(pts, 3);  // Degree by default is 3
                                       
                                        if (isStart)
                                        {
                                            pt = ptStart;
                                            angRot = 90 - (dirStart * (180 / Math.PI));
                                        }
                                        else
                                        {
                                            pt = ptEnd;
                                            angRot = 270 - (dirEnd * (180 / Math.PI));
                                        }
                                           

                                        // Coordinate system on outer spiral point
                                        CoordinateSystem csOrig = CoordinateSystem.ByOrigin(pt);
                                        CoordinateSystem cs = csOrig.Rotate(pt, Vector.ZAxis(), angRot);

                                        NurbsCurve spiral = (NurbsCurve)spiralBase.Transform(cs);
                                      
                                        output.Add(spiral);

                                        break;
                                   
                                    default:
                                        break;
                                }

                            }

                            break;
                        }
                }
            }

            output = SortCurves(output);

            Utils.Log("Alignment.GetCurves Completed.");

            return output;
        }


        /// <summary>
        /// Sorts a list of Dynamo Curves.
        /// </summary>
        /// <param name="curves">THe list of Curves.</param>
        /// <returns></returns>
        private IList<Curve> SortCurves(IList<Curve> curves)
        {
            int n = curves.Count;

            for (int i = 0; i < n; i++)
            {
                Curve c = curves[i];

                Point endPoint = c.EndPoint;

                bool found = i + 1 >= n;

                Point p = null;

                for (int j = i + 1; j < n; j++)
                {
                    p = curves[j].StartPoint;

                    if (p.DistanceTo(endPoint) < 0.00001)
                    {
                        if (i + 1 != j)
                        {
                            Curve temp = curves[i + 1];

                            curves[i + 1] = curves[j];

                            curves[j] = temp;
                        }

                        found = true;

                        break;
                    }

                    p = curves[j].EndPoint;

                    if (p.DistanceTo(endPoint) < 0.00001)
                    {
                        if (i + 1 == j)
                        {
                            curves[i + 1] = curves[j].Reverse();
                        }
                        else
                        {
                            Curve temp = curves[i + 1];

                            curves[i + 1] = curves[j].Reverse();

                            curves[j] = temp;
                        }

                        found = true;

                        break;
                    }
                }
            }

            return curves;
        }

        /// <summary>
        /// Returns the Sample Lines parameters associated with the alignment.
        /// </summary>
        /// <returns></returns>
        [MultiReturn(new string[] { "station", "lengthLeft", "lengthRight", "elevationMin", "elevationMax" })]
        public IList<Dictionary<string, object>> SampleLinesParameters()
        {
            return RevitUtils.AlignmentSampleLinesParameters(this._alignment);
        }

        /// <summary>
        /// Returns the station, offset and elevation of a point from the alignment.
        /// </summary>
        /// <param name="point">The point.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "Station", "Offset", "Elevation"})]
        public Dictionary<string, object> GetStationOffsetElevation(Point point)
        {
            double station = 0;
            double offset = 0;

            ((AeccAlignment)this.InternalElement).StationOffset(point.X, point.Y, out station, out offset);

            return new Dictionary<string, object>() { {"Station", station}, {"Offset", offset}, {"Elevation", point.Z}};
        }

        /// <summary>
        /// Returns a CoordinateSystem along the Alignment at the specified station.
        /// </summary>
        /// <param name="station">The station value.</param>
        /// <param name="offset">The offset value.</param>
        /// <param name="elevation">The elevation value.</param>
        /// <returns></returns>
        public CoordinateSystem CoordinateSystemByStation(double station, double offset=0, double elevation=0)
        {
            Utils.Log("Alignment.CoordinateSystemByStation Started...");

            double northing = 0;
            double easting = 0;
            double northingX = 0;
            double eastingX = 0;

            this._alignment.PointLocation(station, offset, out easting, out northing);

            Point point = Point.ByCoordinates(easting, northing, elevation);

            this._alignment.PointLocation(station, offset + 1, out eastingX, out northingX);

            Point pointX = Point.ByCoordinates(eastingX, northingX, elevation);

            Vector x = Vector.ByTwoPoints(point, pointX).Normalized();

            Vector y = Vector.ZAxis().Cross(x).Normalized();

            CoordinateSystem cs = CoordinateSystem.ByOriginVectors(point, x, y);

            point.Dispose();
            pointX.Dispose();
            x.Dispose();
            y.Dispose();

            Utils.Log("Alignment.CoordinateSystemByStation Completed.");

            return cs;
        }

        /// <summary>
        /// Public textual representation in the Dynamo node preview.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("Alignment(Name = {0}, Length = {1}, Start = {2}, End = {3})",
                this.Name,
                Math.Round(this.Length, 2).ToString(),
                Math.Round(this.Start, 2).ToString(),
                Math.Round(this.End, 2).ToString());
        }
        #endregion
    }
}
