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
using Autodesk.DesignScript.Runtime;

namespace CivilConnection.Tunnel
{
    /// <summary>
    /// Tunnel Ring object type.
    /// </summary>
    [IsVisibleInDynamoLibrary(true)]
    public class Ring
    {
        #region PRIVATE MEMEBERS
        double _angle;
        double _depth;
        double _internalRadius;
        double _externalRadius;
        double _offset;
        double _thickness;
        double _faceAngle;
        double _rho;
        double _lambda;

        CoordinateSystem _frontCS = CoordinateSystem.Identity();
        CoordinateSystem _backCS = CoordinateSystem.Identity();
        CoordinateSystem _startCS = CoordinateSystem.Identity();
        CoordinateSystem _endCS = CoordinateSystem.Identity();
        Line _line;

        Solid _solid;

        #endregion

        #region PUBLIC MEMBERS
        /// <summary>
        /// Gets or sets the angle.
        /// </summary>
        /// <value>
        /// The angle.
        /// </value>
        public double Angle 
        { 
            get 
            {
                if (this.StartCS != null)
                {
                    Utils.Log(string.Format("Ring.Angle started...", ""));

                    var x = this.StartCS.YAxis.Cross(Vector.ZAxis());
                    var z = x.Cross(this.StartCS.YAxis);
                    _angle = z.AngleAboutAxis(this.StartCS.ZAxis, this.StartCS.YAxis);

                    x.Dispose();
                    z.Dispose();

                    Utils.Log(string.Format("Ring.Angle completed.", ""));
                }

                return _angle;
            } 
            set 
            {
                _angle = value; 
            } 
        }
        /// <summary>
        /// Gets or sets the depth.
        /// </summary>
        /// <value>
        /// The depth.
        /// </value>
        public double Depth { get { return _depth; } set { _depth = value; } }
        /// <summary>
        /// Gets or sets the internal radius.
        /// </summary>
        /// <value>
        /// The internal radius.
        /// </value>
        public double InternalRadius { get { return _internalRadius; } set { _internalRadius = value; } }
        /// <summary>
        /// Gets or sets the external radius.
        /// </summary>
        /// <value>
        /// The external radius.
        /// </value>
        public double ExternalRadius { get { return _externalRadius; } set { _externalRadius = value; } }
        /// <summary>
        /// Gets the rho.
        /// </summary>
        /// <value>
        /// The rho.
        /// </value>
        public double Rho
        {
            get
            {
                _rho = (_depth - _offset) * Math.Sin(Math.Atan((_offset / (2 * _externalRadius))));
                return _rho;
            }
        }

        /// <summary>
        /// Gets the lambda.
        /// </summary>
        /// <value>
        /// The rho.
        /// </value>
        public double Lambda
        {
            get
            {
                _lambda = (_depth - _offset) * Math.Cos(Math.Atan((_offset / (2 * _externalRadius))));
                return _lambda;
            }
        }
        /// <summary>
        /// Gets or sets the offset.
        /// </summary>
        /// <value>
        /// The offset.
        /// </value>
        public double Offset { get { return _offset; } set { _offset = value; } }
        /// <summary>
        /// Gets or sets the thickness.
        /// </summary>
        /// <value>
        /// The thickness.
        /// </value>
        public double Thickness { get { return _thickness; } set { _thickness = value; } }
        /// <summary>
        /// Gets the front cs.
        /// </summary>
        /// <value>
        /// The front cs.
        /// </value>
        public CoordinateSystem FrontCS
        {
            get
            {
                return _frontCS;
            }
            set { _frontCS = value; }
        }
        /// <summary>
        /// Gets the back cs.
        /// </summary>
        /// <value>
        /// The back cs.
        /// </value>
        public CoordinateSystem BackCS
        {
            get
            {
                return _backCS;
            }
            set { _backCS = value; }
        }
        /// <summary>
        /// Gets the start cs.
        /// </summary>
        /// <value>
        /// The start cs.
        /// </value>
        public CoordinateSystem StartCS
        {
            get
            {
                return _startCS;
            }
            set { _startCS = value; }
        }
        /// <summary>
        /// Gets the end cs.
        /// </summary>
        /// <value>
        /// The end cs.
        /// </value>
        public CoordinateSystem EndCS
        {
            get
            {
                return _endCS;
            }
            set { _endCS = value; }
        }
        /// <summary>
        /// Gets the ring solid.
        /// </summary>
        /// <value>
        /// The ring solid.
        /// </value>
        public Solid RingSolid
        {
            get
            {
                if (this._solid == null)
                {
                    Utils.Log(string.Format("RingSolid started...", ""));

                    Circle ci = Circle.ByPlaneRadius(this.StartCS.ZXPlane.Offset(-this._offset / 2), this._internalRadius);
                    Circle ce = Circle.ByPlaneRadius(this.StartCS.ZXPlane.Offset(-this._offset / 2), this._externalRadius);

                    Surface ei = ci.Extrude(this.StartCS.YAxis, this._depth);
                    Surface ee = ce.Extrude(this.StartCS.YAxis, this._depth);
                    Curve internalProfileFront = ei.Intersect(this.FrontCS.ZXPlane)[0] as Curve;
                    Curve internalProfileBack = ei.Intersect(this.BackCS.ZXPlane)[0] as Curve;
                    Curve externalProfileFront = ee.Intersect(this.FrontCS.ZXPlane)[0] as Curve;
                    Curve externalProfileBack = ee.Intersect(this.BackCS.ZXPlane)[0] as Curve;

                    Solid si = Solid.ByLoft(new Curve[] { internalProfileFront, internalProfileBack });
                    Solid se = Solid.ByLoft(new Curve[] { externalProfileFront, externalProfileBack });

                    this._solid = se.Difference(si).Rotate(this.StartCS.ZXPlane, this._angle) as Solid;

                    ci.Dispose();
                    ce.Dispose();
                    ei.Dispose();
                    ee.Dispose();
                    internalProfileFront.Dispose();
                    internalProfileBack.Dispose();
                    externalProfileFront.Dispose();
                    externalProfileBack.Dispose();
                    si.Dispose();
                    se.Dispose();

                    Utils.Log(string.Format("RingSolid completed.", ""));
                }

                return this._solid;
            }

            set { _solid = value; }
        }
        /// <summary>
        /// Gets the line.
        /// </summary>
        /// <value>
        /// The line.
        /// </value>
        public Line Line
        {
            get
            {
                if (_line == null)
                {
                    _line = Line.ByStartPointDirectionLength(Point.Origin(), Vector.YAxis(), this._depth - this._offset);
                }
                return _line;
            }
            set { _line = value; }
        }
        /// <summary>
        /// Gets the face angle.
        /// </summary>
        /// <value>
        /// The face angle.
        /// </value>
        public double FaceAngle
        {
            get
            {
                _faceAngle = Utils.RadToDeg(Math.Atan(this._offset / (2 * this._externalRadius)));

                return _faceAngle;
            }
            set { _faceAngle = value; }
        }
        #endregion

        #region INTERNAL CONSTRUCTORS
        /// <summary>
        /// Initializes a new instance of the <see cref="Ring"/> class.
        /// </summary>
        internal Ring() { }

        /// <summary>
        /// Initializes a new isntance of the <see cref="Ring"/> class.
        /// </summary>
        /// <param name="radius">The internal radius.</param>
        /// <param name="thickness">The ring thickness.</param>
        /// <param name="offset">The ring offset in the narrow side.</param>
        /// <param name="depth">The ring depth.</param>
        /// <param name="angle">The ring rotation angle along the tunnel axis.</param>
        /// <param name="cs">The CoordinateSystem.</param>
        internal Ring(double radius, double thickness, double offset, double depth, double angle=0, CoordinateSystem cs = null)
        {
            Utils.Log(string.Format("Ring started...", ""));

            if (cs == null)
            {
                cs = CoordinateSystem.Identity();
            }

            this._internalRadius = radius;
            this._thickness = thickness;
            this._externalRadius = this._internalRadius + this._thickness;
            this._offset = offset;
            this._depth = depth;
            this._angle = Math.Abs(angle) > 360 ? (Math.Abs(angle) - Math.Floor(Math.Abs(angle) / 360) * 360) * Math.Sign(angle) : angle;

            double length = this._depth - this._offset;
            this._line = Line.ByStartPointDirectionLength(cs.Origin, cs.YAxis, length);

            this._faceAngle = Utils.RadToDeg(Math.Atan(this._offset / (2 * this._externalRadius)));

            

            Point p1 = Point.ByCoordinates(0, -this._offset / 2, -this._externalRadius).Transform(cs) as Point;
            Point p2 = Point.ByCoordinates(0, this._offset / 2, this._externalRadius).Transform(cs) as Point;
            Point p3 = Point.ByCoordinates(this._externalRadius, 0, 0).Transform(cs) as Point;
            Plane front = Plane.ByBestFitThroughPoints(new Point[] { p1, p2, p3 });

            Point p4 = Point.ByCoordinates(0, this._depth - this._offset / 2, -this._externalRadius).Transform(cs) as Point;
            Point p5 = Point.ByCoordinates(0, this._depth - 1.5 * this._offset, this._externalRadius).Transform(cs) as Point;
            Point p6 = Point.ByCoordinates(this._externalRadius, this._depth - this._offset, 0).Transform(cs) as Point;
            Plane back = Plane.ByBestFitThroughPoints(new Point[] { p4, p5, p6 });

            this._startCS = CoordinateSystem.ByOriginVectors(this._line.StartPoint,
                   cs.XAxis,
                   cs.YAxis,
                   cs.ZAxis)
                   .Rotate(cs.ZXPlane, this._angle);

            this._endCS = CoordinateSystem.ByOriginVectors(this._line.EndPoint,
                   cs.XAxis,
                   cs.YAxis,
                   cs.ZAxis)
                   .Rotate(cs.ZXPlane, this._angle);

            this._frontCS = CoordinateSystem.ByOriginVectors(this._startCS.Origin, this._startCS.XAxis, front.Normal.Reverse());

            this._backCS = CoordinateSystem.ByOriginVectors(this._endCS.Origin, this._endCS.XAxis, back.Normal.Reverse());


            p1.Dispose();
            p2.Dispose();
            p3.Dispose();
            p4.Dispose();
            p5.Dispose();
            p6.Dispose();
            front.Dispose();
            back.Dispose();

            Utils.Log(string.Format("Ring completed.", ""));
        }

        /// <summary>
        /// Creates a copy of the current Ring.
        /// </summary>
        /// <returns></returns>
        internal Ring Copy()
        {
            #region WIP
            //Ring copy = new Ring()
            //{
            //    _angle = this._angle,
            //    _depth = this._depth,
            //    _externalRadius = this._externalRadius,
            //    _internalRadius = this._internalRadius,
            //    _offset = this._offset,
            //    _thickness = this._thickness,
            //    //a = this.a,
            //    Angle = this.Angle,
            //    //back = this.back,
            //    //BackCS = this.BackCS,
            //    //ce = this.ce,
            //    //ci = this.ci,
            //    Depth = this.Depth,
            //    //ee = this.ee,
            //    //ei = this.ei,
            //    //EndCS = this.EndCS,
            //    //externalProfileBack = this.externalProfileBack,
            //    //externalProfileFront = this.externalProfileFront,
            //    ExternalRadius = this.ExternalRadius,
            //    FaceAngle = this.FaceAngle,
            //    //front = this.front,
            //    //FrontCS = this.FrontCS,
            //    //internalProfileBack = this.internalProfileBack,
            //    //internalProfileFront = this.internalProfileFront,
            //    InternalRadius = this.InternalRadius,
            //    //l = this.l,
            //    //Lambda = this.Lambda,
            //    //Line = this.Line,
            //    Offset = this.Offset,
            //    //p1 = this.p1,
            //    //p2 = this.p2,
            //    //p3 = this.p3,
            //    //p4 = this.p4,
            //    //p5 = this.p5,
            //    //p6 = this.p6,
            //    //Rho = this.Rho,
            //    //RingSolid = this.RingSolid,
            //    //se = this.se,
            //    //si = this.si,
            //    //solid = this.solid,
            //    //StartCS = this.StartCS,
            //    Thickness = this.Thickness
            //};
#endregion

            Ring copy = new Ring(this.InternalRadius, this.Thickness, this.Offset, this.Depth, this.Angle);

            Utils.Log("Copy created");

            return copy;
        }

        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// Reutrns a Ring by dimensions.
        /// </summary>
        /// <param name="radius">The radius.</param>
        /// <param name="thickness">The thickness.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="depth">The depth.</param>
        /// <param name="angle">The angle.</param>
        /// <param name="cs">The CoordinateSystem.</param>
        /// <returns></returns>
        public static Ring ByDimensions(double radius, double thickness, double offset, double depth, double angle=0, CoordinateSystem cs=null)
        {
            return new Ring(radius, thickness, offset, depth, angle, cs);
        }

        /// <summary>
        /// Transforms a Ring by the specified cs.
        /// </summary>
        /// <param name="cs">The cs.</param>
        /// <returns></returns>
        public Ring Transform(CoordinateSystem cs)
        {
            Utils.Log(string.Format("Ring.Transform started...", ""));

            Ring ring = new Ring(this.InternalRadius, this.Thickness, this.Offset, this.Depth, 0, cs);

            Utils.Log(string.Format("Ring.Transform completed.", ""));

            return ring;
        }

        /// <summary>
        /// Rotates the Ring on the starting face by the angle value.
        /// </summary>
        /// <param name="angle">The angle.</param>
        /// <returns></returns>
        public Ring RotateOnFace(double angle)
        {
            Utils.Log(string.Format("Ring.RotateOnFace started...", ""));

            var front = this.FrontCS;
            var start = this.StartCS;
            var a = this.Angle + angle;
            var ring = this.Transform(start.Rotate(front.ZXPlane, angle));

            Utils.Log(string.Format("Ring.RotateOnFace completed.", ""));

            return ring;
        }

        /// <summary>
        /// Returns the followings Ring with face rotated by the specified angle.
        /// </summary>
        /// <param name="angle">The angle.</param>
        /// <returns></returns>
        public Ring Following(double angle)
        {
            CoordinateSystem cs = this.BackCS.Rotate(this.BackCS.YZPlane, this.FaceAngle).Rotate(this.BackCS.ZXPlane, angle);
            Ring output = this.Transform(cs);
            cs.Dispose();
            return output;
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("Ring(InternalRadius = {0:0.000}, Thickness = {1:0.000}, Depth = {2:0.000}, Angle = {3:0.00})", this.InternalRadius, this.Thickness, this.Depth, this.Angle);
        }
        #endregion
    }
}
