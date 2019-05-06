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
using Autodesk.DesignScript.Runtime;
using Autodesk.Revit.DB;

using Revit.GeometryConversion;

using RevitServices.Persistence;
using RevitServices.Transactions;


namespace CivilConnection
{
    /// <summary>
    /// AbstratMEPCurve object Type. Base class for Revti MEP Curve objects.
    /// </summary>
    /// <seealso cref="Revit.Elements.Element" />
    [DynamoServices.RegisterForTrace()]
    [IsVisibleInDynamoLibrary(false)]
    public abstract class AbstractMEPCurve : Revit.Elements.Element
    {
        #region PRIVATE MEMBERS

        /// <summary>
        /// Gets the internal mep curve.
        /// </summary>
        /// <value>
        /// The internal mep curve.
        /// </value>
        internal Autodesk.Revit.DB.MEPCurve InternalMEPCurve
        {
            get;
            private set;
        }

        #endregion

        #region PUBLIC PROPERTIES

        /// <summary>
        /// A reference to the element
        /// </summary>
        public override Autodesk.Revit.DB.Element InternalElement
        {
            get { return InternalMEPCurve; }
        }

        /// <summary>
        /// Gets or sets the location.
        /// </summary>
        /// <value>
        /// The location.
        /// </value>
        public Autodesk.DesignScript.Geometry.Curve Location
        {
            get
            {
                TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);
                DocumentManager.Regenerate();
                var pos = InternalMEPCurve.Location as LocationCurve;
                TransactionManager.Instance.TransactionTaskDone();
                return pos.Curve.ToProtoType();
            }
            set
            {
                TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);
                DocumentManager.Regenerate();
                var pos = InternalMEPCurve.Location as LocationCurve;
                pos.Curve = value.ToRevitType();
                TransactionManager.Instance.TransactionTaskDone();
            }
        }

        /// <summary>
        /// Gets the MEP system.
        /// </summary>
        /// <value>
        /// The system.
        /// </value>
        public Revit.Elements.Element System
        {
            get
            {
                TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);
                DocumentManager.Regenerate();
                var sys = InternalMEPCurve.MEPSystem as MEPSystem;
                TransactionManager.Instance.TransactionTaskDone();
                return Revit.Elements.ElementSelector.ByUniqueId(sys.UniqueId);
            }
        }

        /// <summary>
        /// Gets the width.
        /// </summary>
        /// <value>
        /// The width.
        /// </value>
        public double Width
        {
            get
            {
                TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);
                DocumentManager.Regenerate();
                double d = InternalMEPCurve.Width;
                TransactionManager.Instance.TransactionTaskDone();
                return UtilsObjectsLocation.FeetToMm(d);
            }
        }

        /// <summary>
        /// Gets the height.
        /// </summary>
        /// <value>
        /// The height.
        /// </value>
        public double Height
        {
            get
            {
                TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);
                DocumentManager.Regenerate();
                double d = InternalMEPCurve.Height;
                TransactionManager.Instance.TransactionTaskDone();
                return UtilsObjectsLocation.FeetToMm(d);
            }
        }

        /// <summary>
        /// Gets the level offset.
        /// </summary>
        /// <value>
        /// The level offset.
        /// </value>
        public double LevelOffset
        {
            get
            {
                TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);
                DocumentManager.Regenerate();
                double d = InternalMEPCurve.LevelOffset;
                TransactionManager.Instance.TransactionTaskDone();
                return d;
            }
        }

        #endregion

        #region CONSTRUCTORS

        /// <summary>
        /// Internals the set mep curve.
        /// </summary>
        /// <param name="fi">The MEPCurve instance</param>
        protected void InternalSetMEPCurve(Autodesk.Revit.DB.MEPCurve fi)
        {
            this.InternalMEPCurve = fi;
            this.InternalElementId = fi.Id;
            this.InternalUniqueId = fi.UniqueId;
        }

        #endregion

        #region PRIVATE METHODS

        /// <summary>
        /// Internals the type of the set mep curve.
        /// </summary>
        /// <param name="type">The type.</param>
        protected void InternalSetMEPCurveType(Autodesk.Revit.DB.MEPCurveType type)
        {
            if (InternalMEPCurve.GetTypeId().IntegerValue.Equals(type.Id.IntegerValue))
                return;

            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            InternalMEPCurve.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).Set(type.Id);

            TransactionManager.Instance.TransactionTaskDone();
        }

        /// <summary>
        /// Internals the set position.
        /// </summary>
        /// <param name="start">The start point.</param>
        /// <param name="end">The end point.</param>
        protected void InternalSetPosition(XYZ start, XYZ end)
        {
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            var lp = InternalMEPCurve.Location as LocationCurve;
            if (lp != null && !lp.Curve.ToProtoType().IsAlmostEqualTo(Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(start.ToPoint(), end.ToPoint()))) lp.Curve = Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(start.ToPoint(), end.ToPoint()).ToRevitType();

            TransactionManager.Instance.TransactionTaskDone();
        }

        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// Connectors on the MEPCurve.
        /// </summary>
        /// <returns></returns>
        public CivilConnection.MEP.Connector[] Connectors()
        {
            Utils.Log(string.Format("AbstractMEPCurve.Connectors started...", ""));

            CivilConnection.MEP.Connector[] connectors = new CivilConnection.MEP.Connector[InternalMEPCurve.ConnectorManager.Connectors.Size];
            int i = 0;

            foreach (Autodesk.Revit.DB.Connector c in InternalMEPCurve.ConnectorManager.Connectors)
            {
                connectors[i] = new CivilConnection.MEP.Connector(c);
                i = i + 1;

            }

            Utils.Log(string.Format("AbstractMEPCurve.Connectors completed.", ""));

            return connectors;
        }

        /// <summary>
        /// Returns a text that represents this instance.
        /// </summary>
        /// <returns>
        /// A text that represents this instance.
        /// </returns>
        public override string ToString()
        {
            if (InternalMEPCurve != null && InternalMEPCurve.IsValidObject)
                return string.Format("{2}(System={0}, Name={1})", InternalMEPCurve.MEPSystem.Name, InternalMEPCurve.Name, this.Name);

            return string.Format("{2}(System={0}, Name={1})", "empty", "empty", this.Name);
        }

        #endregion
    }
}
