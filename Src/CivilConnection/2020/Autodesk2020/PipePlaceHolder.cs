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
using Autodesk.Revit.DB.Mechanical;

using Revit.GeometryConversion;

using RevitServices.Persistence;
using RevitServices.Transactions;
using Revit.Elements;
using System.Collections.Generic;


namespace CivilConnection.MEP
{
    /// <summary>
    /// PipePlaceHolder obejct type.
    /// </summary>
    /// <seealso cref="CivilConnection.AbstractMEPCurve" />
    [DynamoServices.RegisterForTrace()]
    [IsVisibleInDynamoLibrary(false)]
    public class PipePlaceHolder : AbstractMEPCurve
    {
        #region PRIVATE PROPERTIES


        #endregion

        #region PUBLIC PROPERTIES


        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Initializes a new instance of the <see cref="PipePlaceHolder"/> class.
        /// </summary>
        /// <param name="instance">The instance.</param>
        protected PipePlaceHolder(Autodesk.Revit.DB.Plumbing.Pipe instance)
        {
            SafeInit(() => InitPipe(instance));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PipePlaceHolder"/> class.
        /// </summary>
        /// <param name="pipeType">Type of the pipe.</param>
        /// <param name="systemType">Type of the system.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <param name="level">The level.</param>
        internal PipePlaceHolder(Autodesk.Revit.DB.Plumbing.PipeType pipeType, Autodesk.Revit.DB.Plumbing.PipingSystemType systemType, XYZ start, XYZ end,
            Autodesk.Revit.DB.Level level)
        {
            InitPipe(pipeType, systemType, start, end, level);
        }

        #endregion

        #region PRIVATE METHODS
        /// <summary>
        /// Initialize a Pipe element.
        /// </summary>
        /// <param name="instance">The instance.</param>
        private void InitPipe(Autodesk.Revit.DB.Plumbing.Pipe instance)
        {
            InternalSetMEPCurve(instance);
        }

        /// <summary>
        /// Initialize a Pipe element.
        /// </summary>
        /// <param name="pipeType">Type of the pipe.</param>
        /// <param name="systemType">Type of the system.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <param name="level">The level.</param>
        private void InitPipe(Autodesk.Revit.DB.Plumbing.PipeType pipeType, Autodesk.Revit.DB.Plumbing.PipingSystemType systemType, XYZ start, XYZ end,
            Autodesk.Revit.DB.Level level)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.MEPCurve>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetMEPCurve(oldFam);
                InternalSetMEPCurveType(pipeType);
                InternalSetPosition(start, end);
                return;
            }

            //Phase 2- There was no existing point, create one
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.Revit.DB.MEPCurve fi;

            if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
            {
                fi = null;
            }
            else
            {
                fi = Autodesk.Revit.DB.Plumbing.Pipe.CreatePlaceholder(DocumentManager.Instance.CurrentDBDocument, systemType.Id, pipeType.Id, level.Id, start, end);
            }

            InternalSetMEPCurve(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }

        /// <summary>
        /// Internals the type of the set piping system.
        /// </summary>
        /// <param name="type">The type.</param>
        private void InternalSetPipingSystemType(Autodesk.Revit.DB.Plumbing.PipingSystemType type)
        {
            if (InternalMEPCurve.MEPSystem.GetTypeId().IntegerValue.Equals(type.Id.IntegerValue))
                return;

            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            InternalMEPCurve.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).Set(type.Id);

            TransactionManager.Instance.TransactionTaskDone();
        }

        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Creates a PipePlaceholder by two points.
        /// </summary>
        /// <param name="pipeType">Type of the pipe.</param>
        /// <param name="systemType">Type of the system.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <param name="level">The level.</param>
        /// <returns></returns>
        public static PipePlaceHolder ByPoints(Revit.Elements.Element pipeType, Revit.Elements.Element systemType, Autodesk.DesignScript.Geometry.Point start, Autodesk.DesignScript.Geometry.Point end, Revit.Elements.Level level)
        {
            Utils.Log(string.Format("PipePlaceHolder.ByPoints started...", ""));

            var oType = pipeType.InternalElement as Autodesk.Revit.DB.Plumbing.PipeType;
            var oSystemType = systemType.InternalElement as Autodesk.Revit.DB.Plumbing.PipingSystemType;
            var totalTransform = RevitUtils.DocumentTotalTransform();

            var nstart = start.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var s = nstart.ToXyz();
            var nend = end.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var e = nend.ToXyz();
            var l = level.InternalElement as Autodesk.Revit.DB.Level;

            if (nstart != null)
            {
                nstart.Dispose();
            }
            if (nend != null)
            {
                nend.Dispose();
            }

            Utils.Log(string.Format("PipePlaceHolder.ByPoints completed.", ""));

            return new PipePlaceHolder(oType, oSystemType, s, e, l);
        }

        #endregion
    }
}
