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

using RevitServices.Persistence;
using RevitServices.Transactions;


namespace CivilConnection.MEP
{
    /// <summary>
    /// Connector object type.
    /// </summary>
    /// <seealso cref="Revit.Elements.Element" />
    [DynamoServices.RegisterForTrace]
    [IsVisibleInDynamoLibrary(false)]
    [SupressImportIntoVM]
    public class Connector : Revit.Elements.Element
    {
        #region PRIVATE PROPERTIES

        /// <summary>
        /// Gets the internal connector.
        /// </summary>
        /// <value>
        /// The internal connector.
        /// </value>
        internal Autodesk.Revit.DB.Connector InternalConnector
        {
            get;
            private set;
        }

        #endregion

        #region PUBLIC PROPERTIES

        /// <summary>
        /// Gets the origin.
        /// </summary>
        /// <value>
        /// The origin.
        /// </value>
        public XYZ Origin { get; private set; }

        /// <summary>
        /// Gets the mep system.
        /// </summary>
        /// <value>
        /// The mep system.
        /// </value>
        public MEPSystem MEPSystem { get; private set; }

        /// <summary>
        /// Gets the connector manager.
        /// </summary>
        /// <value>
        /// The connector manager.
        /// </value>
        public ConnectorManager ConnectorManager { get; private set; }

        /// <summary>
        /// Gets all refs.
        /// </summary>
        /// <value>
        /// All refs.
        /// </value>
        public ConnectorSet AllRefs { get; private set; }

        /// <summary>
        /// Gets the domain.
        /// </summary>
        /// <value>
        /// The domain.
        /// </value>
        public Autodesk.Revit.DB.Domain Domain { get; private set; }

        /// <summary>
        /// Gets the direction.
        /// </summary>
        /// <value>
        /// The direction.
        /// </value>
        public XYZ Direction { get; private set; }

        /// <summary>
        /// Gets the owner.
        /// </summary>
        /// <value>
        /// The owner.
        /// </value>
        public Autodesk.Revit.DB.Element Owner { get; private set; }

        /// <summary>
        /// Gets the type of the connector.
        /// </summary>
        /// <value>
        /// The type of the connector.
        /// </value>
        public Autodesk.Revit.DB.ConnectorType ConnectorType { get; private set; }

        /// <summary>
        /// Gets the internal coordiante system.
        /// </summary>
        /// <value>
        /// The internal coordiante system.
        /// </value>
        public Autodesk.Revit.DB.Transform InternalCoordianteSystem { get; private set; }

        /// <summary>
        /// Gets a value indicating whether this instance is connected.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is connected; otherwise, <c>false</c>.
        /// </value>
        public bool IsConnected { get; private set; }

        /// <summary>
        /// Gets the shape.
        /// </summary>
        /// <value>
        /// The shape.
        /// </value>
        public Autodesk.Revit.DB.ConnectorProfileType Shape { get; private set; }

        /// <summary>
        /// Gets the connector identifier.
        /// </summary>
        /// <value>
        /// The connector identifier.
        /// </value>
        public int ConnectorId { get; private set; }

        /// <summary>
        /// A reference to the element
        /// </summary>
        public override Autodesk.Revit.DB.Element InternalElement
        {
            get
            {
                return Owner;
            }
        }

        #endregion

        #region PRIVATE METHODS

        /// <summary>
        /// Internals the set connector.
        /// </summary>
        /// <param name="fi">The Revit connector</param>
        protected void InternalSetConnector(Autodesk.Revit.DB.Connector fi)
        {
            this.InternalConnector = fi;
            this.ConnectorId = fi.Id;
            this.InternalElementId = fi.Owner.Id;
            this.InternalUniqueId = fi.Owner.UniqueId;
            this.AllRefs = fi.AllRefs;
            this.ConnectorManager = fi.ConnectorManager;
            this.ConnectorType = fi.ConnectorType;
            this.Domain = fi.Domain;
            this.IsConnected = fi.IsConnected;
            this.MEPSystem = fi.MEPSystem;
            this.Origin = fi.Origin;
            this.Owner = fi.Owner;
            this.Shape = fi.Shape;
            this.InternalCoordianteSystem = fi.CoordinateSystem;
        }

        /// <summary>
        /// Method to set the angle.
        /// </summary>
        /// <param name="angle">The angle.</param>
        protected void InternalSetAngle(double angle)
        {
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            InternalConnector.Angle = angle;

            TransactionManager.Instance.TransactionTaskDone();
        }

        /// <summary>
        ///  Method to set the height.
        /// </summary>
        /// <param name="height">The height.</param>
        protected void InternalSetHeight(double height)
        {
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            if (this.Shape == ConnectorProfileType.Rectangular)
            {
                InternalConnector.Height = height; 
            }

            TransactionManager.Instance.TransactionTaskDone();
        }

        /// <summary>
        /// Internals the set origin.
        /// </summary>
        /// <param name="origin">The origin.</param>
        protected void InternalSetOrigin(Autodesk.Revit.DB.XYZ origin)
        {
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            InternalConnector.Origin = origin;

            TransactionManager.Instance.TransactionTaskDone();
        }

        /// <summary>
        /// Method to set the radius.
        /// </summary>
        /// <param name="radius">The radius.</param>
        protected void InternalSetRadius(double radius)
        {
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            if (this.Shape == ConnectorProfileType.Round)
            {
                InternalConnector.Radius = radius; 
            }

            TransactionManager.Instance.TransactionTaskDone();
        }

        /// <summary>
        /// Method to set the width.
        /// </summary>
        /// <param name="width">The width.</param>
        protected void InternalSetWidth(double width)
        {
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            if (this.Shape == ConnectorProfileType.Rectangular)
            {
                InternalConnector.Width = width; 
            }

            TransactionManager.Instance.TransactionTaskDone();
        }

        /// <summary>
        /// Initialize a Connector
        /// </summary>
        /// <param name="instance">The instance.</param>
        private void InitObject(Autodesk.Revit.DB.Connector instance)
        {
            InternalSetConnector(instance);
        }

        #endregion

        #region CONSTRUCTOR
        /// <summary>
        /// Initializes a new instance of the <see cref="Connector"/> class.
        /// </summary>
        /// <param name="instance">The instance.</param>
        public Connector(Autodesk.Revit.DB.Connector instance)
        {
            SafeInit(() => InitObject(instance));
        }

        #endregion

        #region PUBLIC METHODS

        #endregion
    }
}
