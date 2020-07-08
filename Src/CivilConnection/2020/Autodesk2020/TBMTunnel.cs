// Copyright (c) 2017 Autodesk, Inc. All rights reserved.
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
using Autodesk.DesignScript.Geometry;
using System.Collections.Generic;
using System.Linq;
using Autodesk.DesignScript.Runtime;

namespace CivilConnection.Tunnel
{
    /// <summary>
    /// Tunnel Bored Machine Tunnel object type.
    /// </summary>
    [IsVisibleInDynamoLibrary(true)]
    public class TbmTunnel
    {
        #region PRIVATE PROPERTIES
        Solid _ringSolid;
        # endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the baseline curve.
        /// </summary>
        /// <value>
        /// The baseline.
        /// </value>
        public Curve Baseline { get; private set; }
        /// <summary>
        /// Gets the length.
        /// </summary>
        /// <value>
        /// The length.
        /// </value>
        public double Length { get { return PolyCurve.ByPoints(this.Points).Length; } }
        /// <summary>
        /// Gets the rings.
        /// </summary>
        /// <value>
        /// The rings.
        /// </value>
        public IList<Ring> Rings { get; private set;}
        /// <summary>
        /// Gets the ring solids.
        /// </summary>
        /// <value>
        /// The ring solids.
        /// </value>
        public IList<Solid> RingSolids
        {
            get
            {
                IList<Solid> output = new List<Solid>();

                foreach (Ring r in this.Rings)
                {
                    Solid rs = this._ringSolid.Transform(r.StartCS) as Solid;

                    r.RingSolid = rs;

                    output.Add(rs);
                }
                return output;
            }
        }
        /// <summary>
        /// Gets the solid of the base ring for the tunnel.
        /// </summary>
        /// <value>
        /// The ring solid.
        /// </value>
        public Solid BaseRingSolid
        {
            get
            {
                return this._ringSolid;
            }
            set
            {
                this._ringSolid = value; 
            }
        }
        /// <summary>
        /// Gets the angles.
        /// </summary>
        /// <value>
        /// The angles.
        /// </value>
        public IList<double> Angles
        {
            get
            {
                return this.Rings.Select(x => x.Angle).ToList();
            }
        }
        /// <summary>
        /// Gets the points.
        /// </summary>
        /// <value>
        /// The points.
        /// </value>
        public IList<Point> Points
        {
            get
            {
                var pts = this.Rings.Select(x => x.Line.StartPoint).ToList();
                pts.Add(this.Rings.Last().Line.EndPoint);
                return pts;
            }
        }

        #endregion

        #region CONSTRUCTOR
        /// <summary>
        /// Initializes a new instance of the <see cref="TbmTunnel"/> class.
        /// </summary>
        internal TbmTunnel()
        {
            this.Baseline = null;
            this.Rings = new List<Ring>();
        }

        #endregion

        #region PUBLIC CONSTRUCTORS
        /// <summary>
        /// Bies the baseline ring.
        /// </summary>
        /// <param name="baseline">The baseline.</param>
        /// <param name="ring">The ring.</param>
        /// <returns></returns>
        public static TbmTunnel ByBaselineRing(Curve baseline, Ring ring)
        {
            Utils.Log(string.Format("TbmTunnel.ByBaselineRing started...", ""));

            var output = new TbmTunnel();

            try
            {
                output._ringSolid = new Ring(ring.InternalRadius, ring.Thickness, ring.Offset, ring.Depth, 0, CoordinateSystem.Identity()).RingSolid;

                output.Baseline = baseline;

                var pts = Featureline.PointsByChord(baseline, ring.Line.Length);

                Vector yaxis = Vector.ByTwoPoints(pts[0], pts[1]).Normalized();
                Vector xaxis = yaxis.Cross(Vector.ZAxis());

                Utils.Log(string.Format("Points: {0}", pts.Count));

                CoordinateSystem cs = CoordinateSystem.ByOriginVectors(pts[0], xaxis, yaxis);
                cs = cs.Rotate(cs.ZXPlane, -ring.Angle);

                Plane plane = null;
                Point intersection = null;
                Vector v = null;
                Point center = null;
                Circle circle = null;

                Utils.Log(string.Format("Ring {0}", 0));

                Ring newRing = null;

                newRing = ring.Transform(cs);

                output.Rings.Add(newRing);

                for (int i = 0; i < pts.Count; ++i )
                {
                    Utils.Log(string.Format("Ring {0}", output.Rings.Count));

                    try
                    {
                        Ring lastRing = output.Rings.Last();

                        Utils.Log(string.Format("Last Ring Found", ""));

                        plane = lastRing.BackCS.ZXPlane.Offset(ring.Lambda);
                        center = lastRing.BackCS.Origin.Add(lastRing.BackCS.YAxis.Scale(ring.Lambda));

                        if (center.DistanceTo(baseline) > baseline.Length / 2)
                        {
                            Utils.Log(string.Format("Projection too far away from baseline: {0}", plane.Origin));

                            break;
                        }

                        Utils.Log(string.Format("Center {0}", center));

                        plane = Plane.ByOriginNormal(center, lastRing.BackCS.YAxis);

                        Utils.Log(string.Format("Plane {0}", plane));

                        circle = Circle.ByPlaneRadius(plane, ring.Rho);

                        Utils.Log(string.Format("Circle {0}", circle));

                        intersection = circle.ClosestPointTo(baseline);

                        Utils.Log(string.Format("Intersection {0}", intersection));

                        if (null != intersection)
                        {
                            double angle = 0;                            

                            if (intersection.DistanceTo(center) > 0.0001)
                            {
                                v = Vector.ByTwoPoints(center, intersection);

                                Utils.Log(string.Format("Vector {0}", v));

                                angle = Math.Round(lastRing.BackCS.ZAxis.AngleAboutAxis(v, plane.Normal), 5);

                                Utils.Log(string.Format("Angle {0}", angle));
                            }

                            newRing = lastRing.Following(angle);
                        }
                        else
                        {
                            Utils.Log(string.Format("Ring calculation stopped!!!", ""));

                            break;
                        }

                        intersection = null;

                        if (null != newRing)
                        {
                            output.Rings.Add(newRing);

                            newRing = null;

                            Utils.Log(string.Format("Ring added.", ""));
                        }
                        else
                        {
                            Utils.Log(string.Format("ERROR 2: Ring null.", ""));
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Utils.Log(string.Format("ERROR 1: {0}", ex.Message));
                        break;
                    }
                }

                // Disposing these objects cause a failure at runtime

                //cs.Dispose();
                //plane.Dispose();
                //intersection.Dispose();
                //v.Dispose();
                //yaxis.Dispose();
                //xaxis.Dispose();
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR 0: {0}", ex.Message));
            }

            Utils.Log(string.Format("TbmTunnel.ByBaselineRing completed", ""));

            return output;
        }

        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("TbmTunnel(Baseline = {0}, Length = {1:0.000}, Rings = {2})", this.Baseline, this.Length, this.Rings.Count);
        }
        #endregion
    }
}
