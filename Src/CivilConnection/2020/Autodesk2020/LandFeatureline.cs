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
using Autodesk.DesignScript.Interfaces;

using System.Dynamic;

namespace CivilConnection
{
    /// <summary>
    /// LandFeatureline object type.
    /// </summary>
    public class LandFeatureline
    {
        #region PRIVATE PROPERTIES
        /// <summary>
        /// The name
        /// </summary>
        string _name;
        /// <summary>
        /// The polycurve
        /// </summary>
        PolyCurve _polycurve;
        /// <summary>
        /// The LandFeatureline
        /// </summary>
        internal AeccLandFeatureLine _featureline;
        /// <summary>
        /// The minimum grade
        /// </summary>
        double _minGrade;
        /// <summary>
        /// The minimum elevation
        /// </summary>
        double _minElevation;
        /// <summary>
        /// The maximum grade
        /// </summary>
        double _maxGrade;
        /// <summary>
        /// The maximum elevation
        /// </summary>
        double _maxElevation;
        /// <summary>
        /// The coordinates
        /// </summary>
        double[] _coordinates;
        /// <summary>
        /// The style
        /// </summary>
        string _style;
        /// <summary>
        /// The points of the LandFeatureline
        /// </summary>
        IList<Point> _points = new List<Point>();
        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get { return this._name; } set { this._name = value; } }
        /// <summary>
        /// Gets the curve.
        /// </summary>
        /// <value>
        /// The curve.
        /// </value>
        public PolyCurve Curve { get { return this._polycurve; } }
        /// <summary>
        /// Gets the style.
        /// </summary>
        /// <value>
        /// The style.
        /// </value>
        public string Style { get { return this._style; } }
        /// <summary>
        /// Gets the minimum grade.
        /// </summary>
        /// <value>
        /// The minimum grade.
        /// </value>
        public double MinGrade { get { return this._minGrade; } }
        /// <summary>
        /// Gets the maximum grade.
        /// </summary>
        /// <value>
        /// The maximum grade.
        /// </value>
        public double MaxGrade { get { return this._maxGrade; } }
        /// <summary>
        /// Gets the minimum elevation.
        /// </summary>
        /// <value>
        /// The minimum elevation.
        /// </value>
        public double MinElevation { get { return this._minElevation; } }
        /// <summary>
        /// Gets the maximum elevation.
        /// </summary>
        /// <value>
        /// The maximum elevation.
        /// </value>
        public double MaxElevation { get { return this._maxElevation; } }

        /// <summary>
        /// Gets the LandFeatureline points.
        /// </summary>
        public IList<Point> Points
        {
            get
            {
                if (this._points.Count == 0)
                {
                    foreach (Curve c in this._polycurve.Curves())
                    {
                        this._points.Add(c.StartPoint);
                    }

                    this._points.Add(this._polycurve.EndPoint);
                }

                return this._points;
            }
        }
        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Initializes a new instance of the <see cref="LandFeatureline"/> class.
        /// </summary>
        /// <param name="fl">The AeccLandFeatureline.</param>
        /// <param name="pc">The PolyCurve.</param>
        /// <param name="style">The style name.</param>
        internal LandFeatureline(AeccLandFeatureLine fl, PolyCurve pc, string style = "")
        {
            this._featureline = fl;
            this._name = fl.Name;
            this._minGrade = fl.MiniGrade;
            this._minElevation = fl.MiniElevation;
            this._maxElevation = fl.MaxElevation;
            this._maxGrade = fl.MaxGrade;
            this._polycurve = pc;
            this._style = style;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="LandFeatureline"/> class.
        /// </summary>
        /// <param name="fl">The AeccLandFeatureline.</param>
        /// <param name="style">The style name.</param>
        internal LandFeatureline(AeccLandFeatureLine fl, string style = "")
        {
            if (fl == null)
            {
                return;
            }

            this._featureline = fl;
            this._name = fl.Name;
            this._minGrade = fl.MiniGrade;
            this._minElevation = fl.MiniElevation;
            this._maxElevation = fl.MaxElevation;
            this._maxGrade = fl.MaxGrade;

            this._style = style;

            // Andrew Milford - Using reflection does not crash Dynamo
            Type fltype = fl.GetType();

            if (fltype != null)
            {
                try
                {
                    dynamic coord = fltype.InvokeMember("GetPoints",
                        BindingFlags.InvokeMethod,
                        System.Type.DefaultBinder,
                        fl,
                        new object[] { AeccLandFeatureLinePointType.aeccLandFeatureLinePointPI });

                    IList<Point> points = new List<Point>();

                    for (int i = 0; i < coord.Length; i = i + 3)
                    {
                        double x = coord[i];
                        double y = coord[i + 1];
                        double z = coord[i + 2];

                        points.Add(Point.ByCoordinates(x, y, z));
                    }

                    if (points.Count > 1)
                    {
                        try
                        {
                            PolyCurve pc = PolyCurve.ByPoints(Point.PruneDuplicates(points));
                            this._polycurve = pc;
                        }
                        catch
                        {
                            // Not all Polycurves are branching
                            this._polycurve = null;
                        }
                    }
                }
                catch { }
            }
        }
        #endregion

        #region PRIVATE METHODS


        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Gets all the LandFeaturelines from a CivilDocument.
        /// The Style will be empty by default.
        /// Not all the PolyCurves will be branching and it is to be expected to have null values.
        /// </summary>
        /// <param name="civilDocument">The CivilDocument</param>
        /// <returns></returns>
        public static IList<LandFeatureline> GetDocumentLandFeaturelines(CivilDocument civilDocument)
        {
            Utils.Log(string.Format("LandFeatureline.GetDocumentLandFeaturelines started...", ""));

            IList<LandFeatureline> output = civilDocument.GetLandFeaturelines();

            Utils.Log(string.Format("LandFeatureline.GetDocumentLandFeaturelines completed.", ""));

            return output;
        }

        /// <summary>
        /// Creates LandFeatureline
        /// </summary>
        /// <param name="fl">The featureline COM object</param>
        /// <param name="polyCurve">The polycurve</param>
        /// <returns></returns>
        /// 
        [IsVisibleInDynamoLibrary(false)]
        public static LandFeatureline ByObjectPolyCurve(AeccLandFeatureLine fl, PolyCurve polyCurve)
        {
            return new LandFeatureline(fl, polyCurve);
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("LandFeatureline(Name = {0}, Style = {1})", this.Name, this.Style);
        }

        #endregion
    }
}
