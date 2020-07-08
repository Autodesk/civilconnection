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
using System.Xml;
using System.IO;

using Revit.Application;
using Revit.Elements.Views;
using Revit.Elements;
using Revit.Transaction;
using RevitServices.Transactions;
using RevitServices.Persistence;

using Autodesk.DesignScript.Runtime;
using Autodesk.DesignScript.Geometry;
using Revit.GeometryConversion;

using Autodesk.AECC.Interop.Roadway;
using Autodesk.AECC.Interop.UiRoadway;
using Autodesk.AECC.Interop.Land;
using Autodesk.AECC.Interop.UiLand;

using Autodesk.AutoCAD.Interop.Common;

using Autodesk.Revit.DB;

namespace CivilConnection
{
    /// <summary>
    /// OpeningUtils.
    /// </summary>
    public class OpeningUtils
    {
        #region PRIVATE PROPERTIES


        #endregion

        #region PUBLIC PROPERTIES


        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Initializes a new instance of the <see cref="OpeningUtils"/> class.
        /// </summary>
        internal OpeningUtils()
        { }

        #endregion

        #region PRIVATE METHODS


        #endregion

        #region PUBLIC METHODS



        /// <summary>
        /// Gets the rectangular openings.
        /// </summary>
        /// <returns></returns>
        /// <remarks>This method uses walls, doors, windows and generic models bounding boxes to determine the rectangles.
        /// These objects can be in the host file or in linked Revit files.</remarks>
        public static IList<Autodesk.DesignScript.Geometry.Rectangle> GetRectangularOpenings()
        {
            Utils.Log(string.Format("OpeningUtils.GetRectangularOpenings started...", ""));

            Autodesk.Revit.DB.Document doc = DocumentManager.Instance.CurrentDBDocument;

            //look for Walls, Doors, Windows, Generic Models in the current document and in the linked documents
            ElementCategoryFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            ElementCategoryFilter doorFilter = new ElementCategoryFilter(BuiltInCategory.OST_Doors);
            ElementCategoryFilter windowFilter = new ElementCategoryFilter(BuiltInCategory.OST_Windows);
            ElementCategoryFilter genericFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel);

            IList<ElementFilter> filterList = new List<ElementFilter>() { wallFilter, doorFilter, windowFilter, genericFilter };

            LogicalOrFilter orFilter = new LogicalOrFilter(filterList);

            IList<Autodesk.DesignScript.Geometry.Solid> solids = new List<Autodesk.DesignScript.Geometry.Solid>();

            IList<Autodesk.DesignScript.Geometry.Rectangle> output = new List<Autodesk.DesignScript.Geometry.Rectangle>();

            foreach (Autodesk.Revit.DB.Element e in new FilteredElementCollector(doc)
                                 .WherePasses(orFilter)
                                 .WhereElementIsNotElementType()
                                 .Where(x =>
                                        x.Parameters.Cast<Autodesk.Revit.DB.Parameter>()
                                        .First(p => p.Id.IntegerValue == (int)BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).HasValue))
            {
                string comments = e.Parameters.Cast<Autodesk.Revit.DB.Parameter>().First(p => p.Id.IntegerValue == (int)BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();

                if (comments.ToLower() != "opening")
                {
                    continue;
                }

                Transform tr = Transform.Identity;

                if (e is Instance)
                {
                    Instance instance = e as Instance;
                    tr = instance.GetTotalTransform();
                }

                IList<Autodesk.Revit.DB.Solid> temp = new List<Autodesk.Revit.DB.Solid>();

                foreach (GeometryObject go in e.get_Geometry(new Options()))
                {
                    if (go is GeometryInstance)
                    {
                        GeometryInstance geoInstance = go as GeometryInstance;

                        foreach (var gi in geoInstance.SymbolGeometry)
                        {
                            if (gi is Autodesk.Revit.DB.Solid)
                            {
                                Autodesk.Revit.DB.Solid s = gi as Autodesk.Revit.DB.Solid;
                                s = SolidUtils.CreateTransformed(s, tr);
                                temp.Add(s);
                            }
                        }
                    }
                    else
                    {
                        if (go is Autodesk.Revit.DB.Solid)
                        {
                            Autodesk.Revit.DB.Solid s = go as Autodesk.Revit.DB.Solid;
                            s = SolidUtils.CreateTransformed(s, tr);
                            temp.Add(s);
                        }
                    }
                }

                if (temp.Count > 0)
                {
                    Autodesk.Revit.DB.Solid s0 = temp[0];
                    for (int i = 1; i < temp.Count; ++i)
                    {
                        s0 = BooleanOperationsUtils.ExecuteBooleanOperation(s0, temp[i], BooleanOperationsType.Union);
                    }

                    solids.Add(s0.ToProtoType());
                }
            }

            foreach (RevitLinkInstance rli in new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).WhereElementIsNotElementType())
            {
                Autodesk.Revit.DB.Document link = rli.GetLinkDocument();

                foreach (Autodesk.Revit.DB.Element e in new FilteredElementCollector(link)
                         .WherePasses(orFilter)
                         .WhereElementIsNotElementType()
                         .Where(x =>
                                x.Parameters.Cast<Autodesk.Revit.DB.Parameter>()
                                .First(p => p.Id.IntegerValue == (int)BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).HasValue))
                {
                    Transform tr = rli.GetTotalTransform();

                    string comments = e.Parameters.Cast<Autodesk.Revit.DB.Parameter>().First(p => p.Id.IntegerValue == (int)BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();

                    if (comments.ToLower() != "opening")
                    {
                        continue;
                    }

                    if (e is Instance)
                    {
                        Instance instance = e as Instance;
                        tr = tr.Multiply(instance.GetTotalTransform());
                    }

                    IList<Autodesk.Revit.DB.Solid> temp = new List<Autodesk.Revit.DB.Solid>();

                    foreach (var go in e.get_Geometry(new Options()))
                    {
                        if (go is GeometryInstance)
                        {
                            GeometryInstance geoInstance = go as GeometryInstance;

                            foreach (var gi in geoInstance.SymbolGeometry)
                            {
                                if (gi is Autodesk.Revit.DB.Solid)
                                {
                                    Autodesk.Revit.DB.Solid s = gi as Autodesk.Revit.DB.Solid;
                                    s = SolidUtils.CreateTransformed(s, tr);
                                    temp.Add(s);
                                }
                            }
                        }
                        else
                        {
                            if (go is Autodesk.Revit.DB.Solid)
                            {
                                Autodesk.Revit.DB.Solid s = go as Autodesk.Revit.DB.Solid;
                                s = SolidUtils.CreateTransformed(s, tr);
                                temp.Add(s);
                            }
                        }
                    }

                    if (temp.Count > 0)
                    {
                        Autodesk.Revit.DB.Solid s0 = temp[0];
                        for (int i = 1; i < temp.Count; ++i)
                        {
                            s0 = BooleanOperationsUtils.ExecuteBooleanOperation(s0, temp[i], BooleanOperationsType.Union);
                        }

                        solids.Add(s0.ToProtoType());
                    }
                }
            }

            foreach (var s in solids)
            {
                IList<Autodesk.DesignScript.Geometry.Point> points = new List<Autodesk.DesignScript.Geometry.Point>();

                foreach (var v in s.Vertices)
                {
                    points.Add(v.PointGeometry);
                }

                points = Autodesk.DesignScript.Geometry.Point.PruneDuplicates(points);

                Autodesk.DesignScript.Geometry.Plane plane = Autodesk.DesignScript.Geometry.Plane.ByBestFitThroughPoints(points);

                plane = Autodesk.DesignScript.Geometry.Plane.ByOriginNormal(points.Last(), plane.Normal);

                IList<Autodesk.DesignScript.Geometry.Point> temp = new List<Autodesk.DesignScript.Geometry.Point>();

                foreach (var p in points)
                {
                    foreach (var q in p.Project(plane, plane.Normal))
                    {
                        temp.Add(q as Autodesk.DesignScript.Geometry.Point);
                    }

                    foreach (var q in p.Project(plane, plane.Normal.Reverse()))
                    {
                        temp.Add(q as Autodesk.DesignScript.Geometry.Point);
                    }
                }

                temp = Autodesk.DesignScript.Geometry.Point.PruneDuplicates(temp);

                CoordinateSystem cs = CoordinateSystem.ByPlane(plane);

                IList<Autodesk.DesignScript.Geometry.Point> relative = new List<Autodesk.DesignScript.Geometry.Point>();

                foreach (var p in temp)
                {
                    relative.Add(p.Transform(cs.Inverse()) as Autodesk.DesignScript.Geometry.Point);
                }

                var min = Autodesk.DesignScript.Geometry.Point.ByCoordinates(relative.Min(p => p.X),
                    relative.Min(p => p.Y),
                    relative.Min(p => p.Z));

                var max = Autodesk.DesignScript.Geometry.Point.ByCoordinates(relative.Max(p => p.X),
                   relative.Max(p => p.Y),
                   relative.Max(p => p.Z));

                double width = max.X - min.X;
                double height = max.Y - min.Y;

                min = min.Transform(cs) as Autodesk.DesignScript.Geometry.Point;
                max = max.Transform(cs) as Autodesk.DesignScript.Geometry.Point;

                plane = Autodesk.DesignScript.Geometry.Plane.ByOriginNormal(Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(min, max).PointAtParameter(0.5), plane.Normal);

                Autodesk.DesignScript.Geometry.Rectangle rectangle = Autodesk.DesignScript.Geometry.Rectangle.ByWidthLength(plane,
                    width,
                    height);

                output.Add(rectangle);

                plane.Dispose();
                cs.Dispose();
                min.Dispose();
                max.Dispose();
            }

            Utils.Log(string.Format("OpeningUtils.GetRectangularOpenings completed.", ""));

            return output;
        }

        #endregion


    }

}
