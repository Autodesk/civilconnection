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


using RevitServices.Persistence;
using RevitServices.Transactions;
using System;


namespace CivilConnection.MEP
{
    /// <summary>
    /// Fitting obejct type.
    /// </summary>
    /// <seealso cref="Revit.Elements.AbstractFamilyInstance" />
    [DynamoServices.RegisterForTrace]
    //[IsVisibleInDynamoLibrary(false)]
    public class Fitting : Revit.Elements.AbstractFamilyInstance
    {
        #region PRIVATE PROPERTIES


        #endregion

        #region PUBLIC PROPERTIES


        #endregion

        #region CONSTRUCTOR
        /// <summary>
        /// Wrap an existing Fitting.
        /// </summary>
        /// <param name="instance">The instance.</param>
        protected Fitting(Autodesk.Revit.DB.FamilyInstance instance)
        {
            SafeInit(() => InitFitting(instance));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Fitting"/> class.
        /// </summary>
        /// <param name="c1">The first connector.</param>
        /// <param name="c2">The second connector.</param>
        /// <param name="partType">Type of the part.</param>
        internal Fitting(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.Connector c2, string partType)
        {
            if (partType == "Elbow")
            {
                InitElbowObject(c1, c2);
            }

            if (partType == "Union")
            {
                InitUnionObject(c1, c2);
            }

            if (partType == "Transition")
            {
                InitTransitionObject(c1, c2);
            }

            if (partType == "Connection")
            {
                InitConnection(c1, c2);
            }
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="Fitting"/> class.
        /// </summary>
        /// <param name="c1">The first connector.</param>
        /// <param name="c2">The second connector.</param>
        /// <param name="c3">The third connector.</param>
        internal Fitting(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.Connector c2, Autodesk.Revit.DB.Connector c3)
        {
            InitTeeObject(c1, c2, c3);
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="Fitting"/> class.
        /// </summary>
        /// <param name="c1">The first connector.</param>
        /// <param name="c2">The second connector.</param>
        /// <param name="c3">The third connector.</param>
        /// <param name="c4">The fourth connector.</param>
        internal Fitting(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.Connector c2, Autodesk.Revit.DB.Connector c3, Autodesk.Revit.DB.Connector c4)
        {
            InitCrossObject(c1, c2, c3, c4);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Fitting"/> class.
        /// </summary>
        /// <param name="c1">The connector.</param>
        /// <param name="curve">The curve.</param>
        internal Fitting(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.MEPCurve curve)
        {
            InitTakeoffObject(c1, curve);
        }

        #endregion

        #region PRIVATE METHODS
        /// <summary>
        /// Initializes the fitting.
        /// </summary>
        /// <param name="instance">The instance.</param>
        private void InitFitting(Autodesk.Revit.DB.FamilyInstance instance)
        {
            InternalSetFamilyInstance(instance);
        }

        /// <summary>
        /// Initializes the elbow object.
        /// </summary>
        /// <param name="c1">The first connector.</param>
        /// <param name="c2">The second connector.</param>
        private void InitElbowObject(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.Connector c2)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.FamilyInstance>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetFamilyInstance(oldFam);
                return;
            }

            //Phase 2- There was no existing point, create one
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.Revit.DB.FamilyInstance fi;

            if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
            {
                fi = null;
            }
            else
            {
                fi = DocumentManager.Instance.CurrentDBDocument.Create.NewElbowFitting(c1, c2);
            }

            InternalSetFamilyInstance(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }

        /// <summary>
        /// Initializes the union object.
        /// </summary>
        /// <param name="c1">The first connector.</param>
        /// <param name="c2">The second connector.</param>
        private void InitUnionObject(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.Connector c2)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.FamilyInstance>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetFamilyInstance(oldFam);
                return;
            }

            //Phase 2- There was no existing point, create one
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.Revit.DB.FamilyInstance fi;

            if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
            {
                fi = null;
            }
            else
            {
                fi = DocumentManager.Instance.CurrentDBDocument.Create.NewUnionFitting(c1, c2);
            }

            InternalSetFamilyInstance(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }

        /// <summary>
        /// Initializes the connection.
        /// </summary>
        /// <param name="c1">The first connector.</param>
        /// <param name="c2">The second connector.</param>
        private void InitConnection(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.Connector c2)
        {
            //Phase 2- There was no existing point, create one
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
            {
                return;
            }
            else
            {
                c1.ConnectTo(c2);
            }

            TransactionManager.Instance.TransactionTaskDone();
        }

        /// <summary>
        /// Initializes the transition object.
        /// </summary>
        /// <param name="c1">The first connector.</param>
        /// <param name="c2">The second connector.</param>
        private void InitTransitionObject(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.Connector c2)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.FamilyInstance>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetFamilyInstance(oldFam);
                return;
            }

            //Phase 2- There was no existing point, create one
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.Revit.DB.FamilyInstance fi;

            if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
            {
                fi = null;
            }
            else
            {
                fi = DocumentManager.Instance.CurrentDBDocument.Create.NewTransitionFitting(c1, c2);
            }

            InternalSetFamilyInstance(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }

        /// <summary>
        /// Initializes the tee object.
        /// </summary>
        /// <param name="c1">The first connector.</param>
        /// <param name="c2">The second connector.</param>
        /// <param name="c3">The third connector.</param>
        private void InitTeeObject(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.Connector c2, Autodesk.Revit.DB.Connector c3)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.FamilyInstance>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetFamilyInstance(oldFam);
                return;
            }

            //Phase 2- There was no existing point, create one
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.Revit.DB.FamilyInstance fi;

            if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
            {
                fi = null;
            }
            else
            {
                fi = DocumentManager.Instance.CurrentDBDocument.Create.NewTeeFitting(c1, c2, c3);
            }

            InternalSetFamilyInstance(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }

        /// <summary>
        /// Initializes the cross object.
        /// </summary>
        /// <param name="c1">The first connector.</param>
        /// <param name="c2">The second connector.</param>
        /// <param name="c3">The third connector.</param>
        /// <param name="c4">The fourth connector.</param>
        private void InitCrossObject(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.Connector c2, Autodesk.Revit.DB.Connector c3, Autodesk.Revit.DB.Connector c4)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.FamilyInstance>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetFamilyInstance(oldFam);
                return;
            }

            //Phase 2- There was no existing point, create one
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.Revit.DB.FamilyInstance fi;

            if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
            {
                fi = null;
            }
            else
            {
                fi = DocumentManager.Instance.CurrentDBDocument.Create.NewCrossFitting(c1, c2, c3, c4);
            }

            InternalSetFamilyInstance(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }

        /// <summary>
        /// Initializes the takeoff object.
        /// </summary>
        /// <param name="c1">The conenctor.</param>
        /// <param name="curve">The curve.</param>
        private void InitTakeoffObject(Autodesk.Revit.DB.Connector c1, Autodesk.Revit.DB.MEPCurve curve)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.FamilyInstance>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetFamilyInstance(oldFam);
                return;
            }

            //Phase 2- There was no existing point, create one
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.Revit.DB.FamilyInstance fi;

            if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
            {
                fi = null;
            }
            else
            {
                fi = DocumentManager.Instance.CurrentDBDocument.Create.NewTakeoffFitting(c1, curve);
            }

            InternalSetFamilyInstance(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }


        /// <summary>
        /// Connectorses this instance.
        /// </summary>
        /// <returns></returns>
        protected CivilConnection.MEP.Connector[] Connectors()
        {
            Utils.Log(string.Format("Fitting.Connectors started...", ""));

            Autodesk.Revit.DB.FamilyInstance fi = InternalElement as Autodesk.Revit.DB.FamilyInstance;

            CivilConnection.MEP.Connector[] connectors = new CivilConnection.MEP.Connector[fi.MEPModel.ConnectorManager.Connectors.Size];

            int i = 0;

            foreach (Autodesk.Revit.DB.Connector c in fi.MEPModel.ConnectorManager.Connectors)
            {
                connectors[i] = new CivilConnection.MEP.Connector(c);
                i = i + 1;
            }

            Utils.Log(string.Format("Fitting.Connectors completed.", ""));

            return connectors;
        }
        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// Creates an elbow fitting.
        /// </summary>
        /// <param name="curve1">The curve1.</param>
        /// <param name="curve2">The curve2.</param>
        /// <returns></returns>
        public static Fitting Elbow(AbstractMEPCurve curve1, AbstractMEPCurve curve2)
        {
            Utils.Log(string.Format("Fitting.Elbow started...", ""));

            CivilConnection.MEP.Connector s = null;
            CivilConnection.MEP.Connector e = null;

            bool found = false;

            foreach (CivilConnection.MEP.Connector c1 in curve1.Connectors())
            {
                foreach (CivilConnection.MEP.Connector c2 in curve2.Connectors())
                {
                    if (c1.Domain == c2.Domain && c1.ConnectorType == c2.ConnectorType)
                    {
                        if (c1.Origin.IsAlmostEqualTo(c2.Origin))
                        {
                            s = c1;
                            e = c2;
                            found = true;
                            break;
                        }
                    }
                }

                if (found)
                {
                    break;
                }
            }

            try
            {
                return new Fitting(s.InternalConnector, e.InternalConnector, "Elbow");
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR Fitting.Elbow: {0}", ex.Message));

                Utils.Log(string.Format("Fitting.Elbow completed.", ""));

                return new Fitting(s.InternalConnector, e.InternalConnector, "Connection");
            }
        }

        /// <summary>
        /// Creates an union fitting.
        /// </summary>
        /// <param name="curve1">The curve1.</param>
        /// <param name="curve2">The curve2.</param>
        /// <returns></returns>
        public static Fitting Union(AbstractMEPCurve curve1, AbstractMEPCurve curve2)
        {
            Utils.Log(string.Format("Fitting.Union started...", ""));

            CivilConnection.MEP.Connector s = null;
            CivilConnection.MEP.Connector e = null;

            bool found = false;

            foreach (CivilConnection.MEP.Connector c1 in curve1.Connectors())
            {
                foreach (CivilConnection.MEP.Connector c2 in curve2.Connectors())
                {
                    if (c1.Domain == c2.Domain && c1.ConnectorType == c2.ConnectorType)
                    {
                        if (c1.Origin.IsAlmostEqualTo(c2.Origin))
                        {
                            s = c1;
                            e = c2;
                            found = true;
                            break;
                        }
                    }
                }

                if (found)
                {
                    break;
                }
            }

            try
            {
                return new Fitting(s.InternalConnector, e.InternalConnector, "Union");
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR Fitting.Union: {0}", ex.Message));

                Utils.Log(string.Format("Fitting.Union completed.", ""));

                return new Fitting(s.InternalConnector, e.InternalConnector, "Connection");
            }

        }

        /// <summary>
        /// Creates a transition fitting.
        /// </summary>
        /// <param name="curve1">The curve1.</param>
        /// <param name="curve2">The curve2.</param>
        /// <returns></returns>
        public static Fitting Transition(AbstractMEPCurve curve1, AbstractMEPCurve curve2)
        {
            Utils.Log(string.Format("Fitting.Transition started...", ""));

            CivilConnection.MEP.Connector s = null;
            CivilConnection.MEP.Connector e = null;

            bool found = false;

            foreach (CivilConnection.MEP.Connector c1 in curve1.Connectors())
            {
                foreach (CivilConnection.MEP.Connector c2 in curve2.Connectors())
                {
                    if (c1.Domain == c2.Domain && c1.ConnectorType == c2.ConnectorType)
                    {
                        if (c1.Origin.IsAlmostEqualTo(c2.Origin))
                        {
                            s = c1;
                            e = c2;
                            found = true;
                            break;
                        }
                    }
                }

                if (found)
                {
                    break;
                }
            }

            try
            {
                Utils.Log(string.Format("Fitting.Transition completed.", ""));

                return new Fitting(s.InternalConnector, e.InternalConnector, "Transition");
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR Fitting.Transition: {0}", ex.Message));

                Utils.Log(string.Format("Fitting.Transition completed.", ""));

                return new Fitting(s.InternalConnector, e.InternalConnector, "Connection");
            }
        }

        #endregion
    }
}
