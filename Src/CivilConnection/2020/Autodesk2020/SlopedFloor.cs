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
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;
using Autodesk.Revit.DB;
using Revit.GeometryConversion;
using RevitServices.Persistence;
using RevitServices.Transactions;
using System;
using System.Linq;


namespace CivilConnection
{
    /// <summary>
    /// AbstratMEPCurve object Type. Base class for Revti MEP Curve objects.
    /// </summary>
    /// <seealso cref="Revit.Elements.Element" />
    [DynamoServices.RegisterForTrace()]
    //[IsVisibleInDynamoLibrary(false)]
    public class SlopedFloor : Revit.Elements.Element
    {
        #region PRIVATE PROPERTIES

        /// <summary>
        /// An internal handle on the Revit floor
        /// </summary>
        internal Autodesk.Revit.DB.Floor InternalFloor
        {
            get;
            private set;
        }

        /// <summary>
        /// Reference to the Element
        /// </summary>
        [SupressImportIntoVM]
        public override Autodesk.Revit.DB.Element InternalElement
        {
            get { return InternalFloor; }
        }

        #endregion

        #region PUBLIC PROPERTIES

        /// <summary>
        /// Gets and sets the FloorType
        /// </summary>
        public Revit.Elements.FloorType Floortype { get; set; }
        /// <summary>
        /// Gets and sets the Level
        /// </summary>
        public Revit.Elements.Level Level { get; set; }
        /// <summary>
        /// Gets and sets if the floor is structural
        /// </summary>
        public bool Structural { get; set; }

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Private constructor
        /// </summary>
        private SlopedFloor(Autodesk.Revit.DB.Floor floor)
        {
            SafeInit(() => InitFloor(floor));
        }

        /// <summary>
        /// Private constructor
        /// </summary>
        private SlopedFloor(CurveArray curveArray, Autodesk.Revit.DB.Line slopeArrow, double slope, Autodesk.Revit.DB.FloorType floorType, Autodesk.Revit.DB.Level level, bool structural)
        {
            SafeInit(() => InitFloor(curveArray, floorType, level, slopeArrow, slope, structural));
        }

        #endregion

        #region PRIVATE METHODS

        /// <summary>
        /// Set the InternalFloor property and the associated element id and unique id
        /// </summary>
        /// <param name="floor"></param>
        private void InternalSetFloor(Autodesk.Revit.DB.Floor floor)
        {
            InternalFloor = floor;
            InternalElementId = floor.Id;
            InternalUniqueId = floor.UniqueId;
        }

        /// <summary>
        /// Initialize a floor element
        /// </summary>
        private void InitFloor(Autodesk.Revit.DB.Floor floor)
        {
            InternalSetFloor(floor);
        }

        /// <summary>
        /// Initialize a floor element
        /// </summary>
        private void InitFloor(CurveArray curveArray, Autodesk.Revit.DB.FloorType floorType, Autodesk.Revit.DB.Level level, Autodesk.Revit.DB.Line slopeArrow, double slope, bool structural)
        {
            Document doc = DocumentManager.Instance.CurrentDBDocument;

            TransactionManager.Instance.EnsureInTransaction(doc);

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(doc); 
            }

            Autodesk.Revit.DB.Floor floor = null;

            if (floorType.IsFoundationSlab)
            {
                // Foundation Slabs require that the profile curves are planar and horizontal
                // The normal must be orthogonal to the profile, hence the only possible normal is the Z Axis
                floor = doc.Create.NewFoundationSlab(curveArray, floorType, level, structural, XYZ.BasisZ);
             
            }
            else
            {
                // we assume the floor is not structural here, this may be a bad assumption
                floor = doc.Create.NewSlab(curveArray, level, slopeArrow, slope, structural);
                floor.ChangeTypeId(floorType.Id);
            }

            InternalSetFloor(floor);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.CleanupAndSetElementForTrace(doc, InternalFloor);
        }

        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Create a Revit Floor given it's curve outline and Level
        /// </summary>
        /// <param name="outline">The outline.</param>
        /// <param name="floorType">Type of the floor.</param>
        /// <param name="level">The level.</param>
        /// <param name="structural">if set to <c>true</c> [structural].</param>
        /// <returns>
        /// The floor
        /// </returns>
        public static SlopedFloor ByOutlineTypeAndLevel(Autodesk.DesignScript.Geometry.PolyCurve outline, Revit.Elements.FloorType floorType, Revit.Elements.Level level, bool structural)
        {
            Utils.Log(string.Format("SlopedFloor.ByOutlineTypeAndLevel started...", ""));

            try
            {
                var profile = new CurveArray();

                Autodesk.DesignScript.Geometry.Plane plane = Autodesk.DesignScript.Geometry.Plane.ByBestFitThroughPoints(
                    outline.Curves().Cast<Autodesk.DesignScript.Geometry.Curve>().Select(x => x.StartPoint));

                Vector normal = plane.Normal;
                if (normal.Dot(Vector.ZAxis()) <= 0)
                {
                    normal = normal.Reverse();
                }

                Autodesk.DesignScript.Geometry.Point origin = plane.Origin;
                Autodesk.DesignScript.Geometry.Point end = origin.Add(normal);
                Autodesk.DesignScript.Geometry.Point projection = Autodesk.DesignScript.Geometry.Point.ByCoordinates(end.X, end.Y, -1000);
                end = Autodesk.DesignScript.Geometry.Point.ByCoordinates(end.X, end.Y, end.Z + 1000);
                Autodesk.DesignScript.Geometry.Point intersection = null;
                var result = plane.Intersect(Autodesk.DesignScript.Geometry
                    .Line.ByStartPointEndPoint(end, projection));

                if (result.Length > 0)
                {
                    intersection = result[0] as Autodesk.DesignScript.Geometry.Point;
                }
                else
                {
                    var message = "Couldn't find intersection";

                    Utils.Log(string.Format("ERROR: SlopedFloor.ByOutlineTypeAndLevel {0}", message));

                    throw new Exception(message);
                }

                Autodesk.DesignScript.Geometry.Curve temp = Autodesk.DesignScript.Geometry.Line.ByBestFitThroughPoints(new Autodesk.DesignScript.Geometry.Point[] { origin, intersection });

                PolyCurve flat = PolyCurve.ByJoinedCurves(outline.PullOntoPlane(Autodesk.DesignScript.Geometry.Plane.XY()
                    .Offset(temp.StartPoint.Z)).Explode().Cast<Autodesk.DesignScript.Geometry.Curve>().ToList());

                Autodesk.DesignScript.Geometry.Curve flatLine = temp.PullOntoPlane(Autodesk.DesignScript.Geometry.Plane.XY().Offset(temp.StartPoint.Z));

                if (Math.Abs(Math.Abs(plane.Normal.Dot(Vector.ZAxis())) - 1) < 0.00001)
                {
                    var f = Revit.Elements.Floor.ByOutlineTypeAndLevel(flat, floorType, level);
                    f.InternalElement.Parameters.Cast<Autodesk.Revit.DB.Parameter>()
                        .First(x => x.Id.IntegerValue.Equals(Autodesk.Revit.DB.BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL))
                        .Set(structural ? 1 : 0);

                    plane.Dispose();
                    flatLine.Dispose();
                    flat.Dispose();
                    origin.Dispose();
                    end.Dispose();
                    projection.Dispose();
                    intersection.Dispose();
                    temp.Dispose();

                    return new SlopedFloor(f.InternalElement as Autodesk.Revit.DB.Floor);
                }

                double slope = (temp.EndPoint.Z - temp.StartPoint.Z) / flatLine.Length;

                foreach (Autodesk.DesignScript.Geometry.Curve c in flat.Curves())
                {
                    profile.Append(c.ToRevitType());
                }

                Autodesk.Revit.DB.Line slopeArrow = flatLine.ToRevitType() as Autodesk.Revit.DB.Line;

                var ft = floorType.InternalElement as Autodesk.Revit.DB.FloorType;
                var lvl = level.InternalElement as Autodesk.Revit.DB.Level;

                var floor = new SlopedFloor(profile, slopeArrow, slope, ft, lvl, structural);

                floor.Level = level;
                floor.Floortype = floorType;
                floor.Structural = structural;

                plane.Dispose();
                flatLine.Dispose();
                flat.Dispose();
                origin.Dispose();
                end.Dispose();
                projection.Dispose();
                intersection.Dispose();
                temp.Dispose();

                Utils.Log(string.Format("SlopedFloor.ByOutlineTypeAndLevel completed.", ""));

                return floor;
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: SlopedFloor.ByOutlineTypeAndLevel {0}", ex.Message));

                throw ex;
            }
        }


        #endregion
    }
}
