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


namespace CivilConnection
{
    /// <summary>
    /// Collection of ShapePoints
    /// </summary>
    [Newtonsoft.Json.JsonConverter(typeof(ShapePointArrayConverter))]
    public class ShapePointArray
    {
        #region PRIVATE PROPERTIES


        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the points.
        /// </summary>
        /// <value>
        /// The points.
        /// </value>
        /// 
        [Newtonsoft.Json.JsonProperty("Points")]
        public IList<ShapePoint> Points { get; set; }

        /// <summary>
        /// Gets the count.
        /// </summary>
        /// <value>
        /// The count.
        /// </value>
        [Newtonsoft.Json.JsonIgnore]
        public int Count { get { return this.Points.Count; } }

        #endregion

        #region CONSTRUCTOR
        /// <exclude />
        internal ShapePointArray() { }

        #endregion

        #region PRIVATE METHODS

        /// <exclude />
        private ShapePointArray Copy()
        {
            Utils.Log(string.Format("ShapePointArray.Copy started...", ""));

            ShapePointArray spa = new ShapePointArray();

            spa.Points = new List<ShapePoint>();

            foreach (ShapePoint p in this.Points)
            {
                var sp = p.Copy(spa.Points.Count);
                spa.Points.Add(sp);
            }

            Utils.Log(string.Format("ShapePointArray.Copy completed.", ""));

            return spa;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapePointArray"/> class.
        /// </summary>
        /// <param name="shapePoints">The shape points.</param>
        [Newtonsoft.Json.JsonConstructor]
        internal ShapePointArray(IList<ShapePoint> shapePoints)
        {
            this.Points = new List<ShapePoint>();

            for (int i = 0; i < shapePoints.Count; ++i)
            {
                this.Points.Add(shapePoints[i].Copy(i));
            }
        }
        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Returns a ShapePointArray object
        /// </summary>
        /// <param name="shapePointList">The list of ShapePoints.</param>
        /// <returns></returns>
        public static ShapePointArray ByShapePointList(IList<ShapePoint> shapePointList)
        {
            return new ShapePointArray(shapePointList);
        }

        //[IsVisibleInDynamoLibrary(false)]
        /// <summary>
        /// Adds the specified shape point.
        /// </summary>
        /// <param name="shapePoint">The shape point.</param>
        /// <returns></returns>
        public ShapePointArray Add(ShapePoint shapePoint)
        {
            Utils.Log(string.Format("ShapePointArray.Add started...", ""));

            ShapePointArray spa = this.Copy();

            var sp = shapePoint.Copy();
            sp.SetId(spa.Points.Count);
            spa.Points.Add(sp);

            Utils.Log(string.Format("ShapePointArray.Add completed.", ""));

            return spa;
        }

        //[IsVisibleInDynamoLibrary(false)]
        /// <summary>
        /// Renumbers the ShapePoints in the instance.
        /// </summary>
        public ShapePointArray Renumber()
        {
            Utils.Log(string.Format("ShapePointArray.Renumber started...", ""));

            ShapePointArray spa = this.Copy();

            if (spa.Points.Count > 0)
            {
                for (int i = 0; i < spa.Points.Count; ++i)
                {
                    spa.Points[i].SetId(i);
                }
            }
            else
            {
                var message = "The collection is empty";

                Utils.Log(string.Format("ERROR: ShapePointArray.Renumber {0}", message));

                throw new Exception(message);
            }

            Utils.Log(string.Format("ShapePointArray.Renumber completed.", ""));

            return spa;
        }

        //[IsVisibleInDynamoLibrary(false)]
        /// <summary>
        /// Removes the ShapePoint at index.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public ShapePointArray RemoveAtIndex(int index)
        {
            Utils.Log(string.Format("ShapePointArray.RemoveAtIndex started...", ""));

            ShapePointArray spa = this.Copy();

            if (index < 0 || index > spa.Count - 1)
            {
                throw new Exception("Invalid index");
            }

            spa.Points.RemoveAt(index);
            spa.Renumber();

            Utils.Log(string.Format("ShapePointArray.RemoveAtIndex completed.", ""));

            return spa;
        }

        //[IsVisibleInDynamoLibrary(false)]
        /// <summary>
        /// Sorts the by station.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public ShapePointArray SortByStation(int index)
        {
            Utils.Log(string.Format("ShapePointArray.SortByStation started...", ""));

            ShapePointArray spa = this.Copy();

            spa.Points = spa.Points.OrderBy(p => p.Station).ToList();
            spa.Renumber();

            Utils.Log(string.Format("ShapePointArray.SortByStation completed.", ""));

            return spa;
        }

        //[IsVisibleInDynamoLibrary(false)]
        /// <summary>
        /// Splits ShpaePoints into subset of ShapePoints by featureline.
        /// </summary>
        /// <returns></returns>
        public IList<ShapePointArray> SplitByFeatureline()
        {
            return this.Points.GroupBy(p => p.Featureline).Select(g => new ShapePointArray(g.ToList())).ToList();
        }

        //[IsVisibleInDynamoLibrary(false)]
        /// <summary>
        /// Reverses this instance.
        /// </summary>
        /// <returns></returns>
        public ShapePointArray Reverse()
        {
            Utils.Log(string.Format("ShapePointArray.Reverse started...", ""));

            ShapePointArray spa = this.Copy();

            spa.Points = spa.Points.Reverse<ShapePoint>().ToList();
            spa.Renumber();

            Utils.Log(string.Format("ShapePointArray.Reverse completed.", ""));

            return spa;
        }

        /// <summary>
        /// Joins two ShapePoints objects into a new one.
        /// </summary>
        /// <param name="other">The other ShapePointArray.</param>
        /// <returns></returns>
        //[IsVisibleInDynamoLibrary(false)]
        public ShapePointArray Join(ShapePointArray other)
        {
            Utils.Log(string.Format("ShapePointArray.Join started...", ""));

            ShapePointArray spa = this.Copy();

            foreach (ShapePoint p in other.Points)
            {
                var sp = p.Copy(spa.Points.Count);
                spa.Points.Add(sp);
            }

            Utils.Log(string.Format("ShapePointArray.Join completed.", ""));

            return spa;
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("ShapePointArray(Count={0})", this.Count);
        }
        #endregion
    }
}
