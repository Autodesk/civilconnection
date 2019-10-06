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
using DynamoServices;
using Revit.Elements;
using RevitServices.Persistence;
using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;
using ADSK_Parameters = CivilConnection.UtilsObjectsLocation.ADSK_Parameters;

namespace CivilConnection
{
    /// <summary>
    /// Shape Point Class
    /// </summary>
    //[IsVisibleInDynamoLibrary(false)]
    [RegisterForTrace()]
    [Newtonsoft.Json.JsonConverter(typeof(MultiPointConverter))]
    public class MultiPoint
    {
        #region PRIVATE PROPERTIES


        #endregion

        #region PUBLIC PROPERTIES

        /// <summary>
        /// Gets or sets the shape points.
        /// </summary>
        /// <value>
        /// The shape points.
        /// </value>
        /// 
        [Newtonsoft.Json.JsonProperty("ShapePoints")]
        public ShapePointArray ShapePoints { get; set; }

        #endregion

        #region CONSTRUCTOR

        internal MultiPoint() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="MultiPoint"/> class.
        /// </summary>
        /// <param name="shapePointArray">A ShapePointArray object.</param>
        [Newtonsoft.Json.JsonConstructor]
        internal MultiPoint(ShapePointArray shapePointArray)
        {
            this.ShapePoints = shapePointArray;
        }

        #endregion

        #region PRIVATE METHODS

        /// <exclude />
        private Element ToHorizontalFloor(FloorType floorType, Level level)
        {
            Utils.Log(string.Format("MultiPoint.ToHorizontalFloor started...", ""));

            try
            {
                if (!SessionVariables.ParametersCreated)
                {
                    UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
                }

                PolyCurve outline = PolyCurve.ByPoints(this.ShapePoints.Points.Select(p => p.RevitPoint).ToList(), true);

                outline = PolyCurve.ByJoinedCurves(outline.PullOntoPlane(Plane.XY()).Explode().Cast<Curve>().ToList());

                var output = Floor.ByOutlineTypeAndLevel(outline, floorType, level);

                output.SetParameterByName(ADSK_Parameters.Instance.MultiPoint.Name, SerializeJSON());

                Utils.Log(string.Format("MultiPoint.ToHorizontalFloor completed.", ""));

                return output;
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: MultiPoint.ToHorizontalFloor {0}", ex.Message));
                throw ex;
            }

            
        }

        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Returns a MultiPoint by a collection of shape points.
        /// </summary>
        /// <param name="shapePointArray">The shape points.</param>
        /// <returns></returns>
        public static MultiPoint ByShapePointArray(ShapePointArray shapePointArray)
        {
            return new MultiPoint(shapePointArray);
        }

        /// <summary>
        /// Returns a MultiPoint by a collection of shape points.
        /// </summary>
        /// <param name="element">The Revit element.</param>
        /// <returns></returns>
        public static MultiPoint ByElement(Element element)
        {
            string parameter = element.GetParameterValueByName(ADSK_Parameters.Instance.MultiPoint.Name) as string;

            if (parameter != "" && parameter != null)
            {
                return Newtonsoft.Json.JsonConvert.DeserializeObject<MultiPoint>(parameter);
            }

            return null;
        }

        /// <summary>
        /// Converts the MultiPoint into a floor of the specified type.
        /// </summary>
        /// <param name="floorType">Type of the floor.</param>
        /// <param name="level">The level.</param>
        /// <param name="structural">if set to <c>true</c> [structural].</param>
        /// <returns></returns>
        public Element ToFloor(FloorType floorType, Level level, bool structural = true)
        {
            Utils.Log(string.Format("MultiPoint.ToFloor started...", ""));

            try
            {
                if (!SessionVariables.ParametersCreated)
                {
                    UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
                }

                PolyCurve outline = PolyCurve.ByPoints(this.ShapePoints.Points.Select(p => p.RevitPoint).ToList(), true);

                if (null == outline)
                {
                    System.Windows.Forms.MessageBox.Show("Outline is null");
                }

                Element output = null;

                try
                {
                    output = SlopedFloor.ByOutlineTypeAndLevel(outline, floorType, level, structural);
                }
                catch (Exception ex)
                {
                    Utils.Log(string.Format("ERROR: MultiPoint.ToFloor {0}", ex.Message));

                    output = Floor.ByOutlineTypeAndLevel(outline, floorType, level);
                }

                output.SetParameterByName(ADSK_Parameters.Instance.MultiPoint.Name, this.SerializeJSON());
                output.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
                output.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);

                Utils.Log(string.Format("MultiPoint.ToFloor completed.", ""));

                return output;
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: MultiPoint.ToFloor {0}", ex.Message));
                throw ex;
            }


        }

        /// <summary>
        /// Converts the MultiPoint into an adaptive component of the specified type.
        /// </summary>
        /// <param name="familyType">Type of the family.</param>
        /// <returns></returns>
        public AdaptiveComponent ToAdaptiveComponent(FamilyType familyType)
        {
            Utils.Log(string.Format("MultiPoint.ToAdaptiveComponent started...", ""));

            AdaptiveComponent output = null;

            try
            {
                if (!SessionVariables.ParametersCreated)
                {
                    UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
                }

                output = AdaptiveComponent.ByPoints(new Point[][] { this.ShapePoints.Points.Select(p => p.RevitPoint).ToArray() }, familyType)[0];
                output.SetParameterByName(ADSK_Parameters.Instance.MultiPoint.Name, this.SerializeJSON());
                output.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
                output.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);
            }
            catch (Exception ex)
            {
                Utils.Log(ex.Message);
            }

            Utils.Log(string.Format("MultiPoint.ToAdaptiveComponent completed.", ""));

            return output;
        }

        /// <summary>
        /// Serializes this instance.
        /// </summary>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public string SerializeXML()
        {
            Utils.Log(string.Format("MultiPoint.SerializeXML started...", ""));

            XmlSerializer serializer = new XmlSerializer(typeof(MultiPoint));
            string xml = "";
            using (StringWriter sw = new StringWriter())
            {
                using (XmlWriter xw = XmlWriter.Create(sw))
                {
                    serializer.Serialize(xw, this);
                    xml = serializer.ToString();
                }
            }

            Utils.Log(string.Format("MultiPoint.SerializeXML completed.", ""));

            return xml;
        }

        /// <summary>
        /// Serializes to json.
        /// </summary>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public string SerializeJSON()
        {
            Utils.Log(string.Format("MultiPoint.SerializeJSON started...", ""));

            string serialized = "";

            try
            {
                serialized = Newtonsoft.Json.JsonConvert.SerializeObject(this);
            }
            catch (Exception ex)
            {
                serialized = ex.Message;
            }

            Utils.Log(serialized);

            Utils.Log(string.Format("MultiPoint.SerializeJSON completed.", ""));

            return serialized;
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("MultiPoint({0})",
                this.ShapePoints);
        }

        #endregion
    }
}
