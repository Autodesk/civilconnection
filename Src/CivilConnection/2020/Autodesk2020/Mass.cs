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
using System.IO;
using System.Collections.Generic;
using System;
using System.Linq;

namespace CivilConnection
{
    /// <summary>
    /// Mass obejct type.
    /// </summary>
    /// <seealso cref="Revit.Elements.AbstractFamilyInstance" />
    [DynamoServices.RegisterForTrace]
    public class Mass : Revit.Elements.AbstractFamilyInstance
    {
        #region PRIVATE PROPERTIES
        /// <summary>
        /// Reference to the Element
        /// </summary>
        internal Autodesk.Revit.DB.FamilyInstance InternalFamilyInstance
        {
            get;
            private set;
        }

        #endregion

        #region PUBLIC PROPERTIES

        /// <summary>
        /// Reference to the Element
        /// </summary>
        public override Autodesk.Revit.DB.Element InternalElement
        {
            get
            {
                return InternalFamilyInstance;
            }
        }

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Wrap an existing Mass.
        /// </summary>
        /// <param name="instance">The family instance.</param>
        protected Mass(Autodesk.Revit.DB.FamilyInstance instance)
        {
            SafeInit(() => InitMass(instance));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Mass"/> class.
        /// </summary>
        /// <param name="fs">The fs.</param>
        /// <param name="pos">The position.</param>
        internal Mass(Autodesk.Revit.DB.FamilySymbol fs, XYZ pos)
        {
            SafeInit(() => InitMass(fs, pos));
        }

        #endregion

        #region PRIVATE METHODS
        /// <summary>
        /// Initializes the mass.
        /// </summary>
        /// <param name="instance">The instance.</param>
        private void InitMass(Autodesk.Revit.DB.FamilyInstance instance)
        {
            InternalSetFamilyInstance(instance);
        }

        /// <summary>
        /// Initializes the mass.
        /// </summary>
        /// <param name="fs">The fs.</param>
        /// <param name="pos">The position.</param>
        private void InitMass(Autodesk.Revit.DB.FamilySymbol fs, XYZ pos)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.FamilyInstance>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetFamilyInstance(oldFam);
                InternalSetFamilySymbol(fs);
                InternalSetPosition(pos);
                return;
            }

            //Phase 2- There was no existing point, create one
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.Revit.DB.FamilyInstance fi;

            if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
            {
                fi = DocumentManager.Instance.CurrentDBDocument.FamilyCreate.NewFamilyInstance(pos, fs,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            }
            else
            {
                fi = DocumentManager.Instance.CurrentDBDocument.Create.NewFamilyInstance(
                    pos, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            }

            InternalSetFamilyInstance(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }

        /// <summary>
        /// Internals the set family instance.
        /// </summary>
        /// <param name="fi">The family instance.</param>
        protected void InternalSetFamilyInstance(Autodesk.Revit.DB.FamilyInstance fi)
        {
            this.InternalFamilyInstance = fi;
            this.InternalElementId = fi.Id;
            this.InternalUniqueId = fi.UniqueId;
        }

        /// <summary>
        /// Method to set position.
        /// </summary>
        /// <param name="fi">The fi.</param>
        private void InternalSetPosition(XYZ fi)
        {
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            var lp = InternalFamilyInstance.Location as LocationPoint;
            if (lp != null && !lp.Point.IsAlmostEqualTo(fi)) lp.Point = fi;

            TransactionManager.Instance.TransactionTaskDone();
        }

        /// <summary>
        /// Closes the family.
        /// </summary>
        /// <param name="name">The name.</param>
        private static void CloseFamily(string name)
        {
            Utils.Log(string.Format("Mass.CloseFamily started...", ""));

            var app = DocumentManager.Instance.CurrentUIApplication.Application;

            string path = Path.Combine(Path.GetTempPath(), name);

            if (File.Exists(path))
            {
                foreach (Document doc in app.Documents)
                {
                    if (doc.PathName == path)
                    {
                        doc.Close(false);
                        break;
                    }
                }

                File.Delete(path);
            }

            Utils.Log(string.Format("Mass.CloseFamily completed.", ""));
        }

        /// <summary>
        /// Utility function that creates model curves in the document.
        /// </summary>
        /// <param name="doc">The Document</param>
        /// <param name="cl">The CurveLoop.</param>
        private static void CreateModelCurves(Document doc, CurveLoop cl)
        {
            Utils.Log(string.Format("Mass.CreateModelCurves started...", ""));

            Plane plane = cl.GetPlane();

            Autodesk.Revit.DB.SketchPlane sp = Autodesk.Revit.DB.SketchPlane.Create(doc, cl.GetPlane());

            if (doc.IsFamilyDocument)
            {
                foreach (Curve c in cl)
                {
                    doc.FamilyCreate.NewModelCurve(c, sp);
                }
            }
            else
            {
                foreach (Curve c in cl)
                {
                    doc.Create.NewModelCurve(c, sp);
                }
            }

            Utils.Log(string.Format("Mass.CreateModelCurves completed.", ""));
        }

        /// <summary>
        /// Utility function that creates model curves in the document.
        /// </summary>
        /// <param name="doc">The Document</param>
        /// <param name="c">The curve.</param>
        private static Autodesk.Revit.DB.ModelCurve CreateModelCurve(Document doc, Curve c)
        {
            Utils.Log(string.Format("Mass.CreateModelCurve started...", ""));

            Autodesk.Revit.DB.ModelCurve output = null;

            Plane plane = null;

            try
            {
                XYZ normal = c.GetEndPoint(0).CrossProduct(c.GetEndPoint(1)).Normalize();

                plane = Plane.CreateByNormalAndOrigin(normal, c.GetEndPoint(0));

                Autodesk.Revit.DB.SketchPlane sp = Autodesk.Revit.DB.SketchPlane.Create(doc, plane);

                if (doc.IsFamilyDocument)
                {
                    doc.FamilyCreate.NewModelCurve(c, sp);
                }
                else
                {
                    doc.Create.NewModelCurve(c, sp);
                }
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: Mass.CreateModelCurve {0}", ex.Message));

                return output;
            }

            Utils.Log(string.Format("Mass.CreateModelCurve completed.", ""));

            return output;
        }

        /// <summary>
        /// Closes the open family document and returns the family name with extension
        /// </summary>
        /// <param name="app">A reference to the application.</param>
        /// <param name="famName">The title of the family document to close including extension.</param>
        private static string CloseDocument(Autodesk.Revit.ApplicationServices.Application app, string famName)
        {
            TransactionManager.Instance.ForceCloseTransaction();

            foreach (Document d in app.Documents)
            {
                if (d.Title == famName)
                {
                    Utils.Log(string.Format("Closing document...{0}", ""));

                    d.Close(false);

                    Utils.Log(string.Format("Document Closed: {0}", famName));
                }
            }

            return famName;
        }


        /// <summary>
        /// Returns a reference to a Family document with the specified template and family name.
        /// </summary>
        /// <param name="app">A reference to the application.</param>
        /// <param name="famName">The title of the family document to close including extension.</param>
        /// <param name="familyTemplate">The mass template path.</param>
        /// <param name="family">The reference to the Family object in the model.</param>
        /// <param name="famDoc">The output family document.</param>
        /// <param name="rvtFI">A placeholder for the family instance in the Revit model.</param>
        private static int GetFamilyDocument(Autodesk.Revit.ApplicationServices.Application app,
            string famName,
            string familyTemplate,
            out Autodesk.Revit.DB.Family family,
            out Document famDoc,
            out Autodesk.Revit.DB.FamilyInstance rvtFI)
        {
            Utils.Log(string.Format("Mass.GetFamilyDocument started...", ""));

            family = null;

            famDoc = null;

            rvtFI = null;

            string famPath = Path.Combine(Path.GetTempPath(), famName);

            try
            {
                famPath = Path.Combine(Path.GetDirectoryName(DocumentManager.Instance.CurrentDBDocument.PathName), famName);
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: Mass.GetFamilyDocument {0}", ex.Message));
            }

            foreach (Autodesk.Revit.DB.Family f in DocumentManager.Instance.ElementsOfType<Autodesk.Revit.DB.Family>())
            {
                if (f.Name + ".rfa" == famName)
                {
                    family = f;

                    Utils.Log(string.Format("Family Found: {0}", family.Id.IntegerValue));

                    break;
                }
            }

            if (family != null)
            {
                ElementId famId = family.Id;

                foreach (Autodesk.Revit.DB.FamilyInstance rfi in new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
                       .OfClass(typeof(Autodesk.Revit.DB.FamilyInstance))
                       .WhereElementIsNotElementType()
                       .Cast<Autodesk.Revit.DB.FamilyInstance>()
                       .Where(x => x.Symbol.Family.Id.IntegerValue.Equals(famId.IntegerValue)))
                {
                    rvtFI = rfi;

                    Utils.Log(string.Format("Family Instance Found: {0}", rfi.Id.IntegerValue));

                    break;
                }
            }

            if (null != family)
            {
                TransactionManager.Instance.ForceCloseTransaction();

                Utils.Log(string.Format("Closing Transactions...", ""));

                famDoc = DocumentManager.Instance.CurrentDBDocument.EditFamily(family);

                Utils.Log(string.Format("Editing Family {0}...", family.Name));
            }
            else
            {
                try
                {
                    Utils.Log(string.Format("New Family {0}...", famPath));

                    famDoc = app.NewFamilyDocument(familyTemplate);

                    var sao = new SaveAsOptions();
                    sao.OverwriteExistingFile = true;
                    sao.Compact = true;
                    sao.MaximumBackups = 1;
                    famDoc.SaveAs(famPath, sao);

                    Utils.Log(string.Format("Family Ready...", ""));
                }
                catch (Exception ex)
                {
                    Utils.Log(string.Format("ERROR: Mass.GetFamilyDocument {0}", ex.Message));

                    return 0;
                }
            }

            Utils.Log(string.Format("Mass.GetFamilyDocument completed.", ""));

            return 1;
        }

        /// <summary>
        /// Removes all CurveElements and FreeFormElements in the family document.
        /// </summary>
        /// <param name="famDoc">The family document.</param>
        private static void CleanupFamilyDocument(Document famDoc)
        {
            Utils.Log(string.Format("Mass.CleanupFamilyDocument started...", ""));

            IList<ElementId> toDelete = new List<ElementId>();

            toDelete = new FilteredElementCollector(famDoc).OfClass(typeof(Autodesk.Revit.DB.Form)).WhereElementIsNotElementType().ToElementIds().ToList();

            foreach (ElementId id in new FilteredElementCollector(famDoc).OfClass(typeof(Autodesk.Revit.DB.CurveElement)).WhereElementIsNotElementType().ToElementIds())
            {
                toDelete.Add(id);
            }

            toDelete = toDelete.Concat(new FilteredElementCollector(famDoc)
               .OfClass(typeof(Autodesk.Revit.DB.FreeFormElement))
               .WhereElementIsNotElementType()
               .ToElementIds())
               .ToList();

            Utils.Log(string.Format("Removing {0} elements...", toDelete.Count));

            foreach (ElementId id in toDelete)
            {
                try
                {
                    famDoc.Delete(id);
                }
                catch (Exception ex)
                {
                    Utils.Log(string.Format("ERROR: Mass.ByCrossSections {0}", ex.Message));

                    continue;
                }
            }

            Utils.Log(string.Format("Mass.CleanupFamilyDocument completed..", ""));
        }

        /// <summary>
        /// Trying to join the solids in the list.
        /// </summary>
        /// <param name="famDoc">The family document.</param>
        /// <param name="solids">The list of Solid to try to join.</param>
        private static void TryJoinSolids(Document famDoc, IList<Solid> solids)
        {
            Utils.Log(string.Format("Mass.TryJoinSolids started...", ""));

            if (solids.Count > 0)
            {
                Solid s = solids[0];

                solids.RemoveAt(0);

                Utils.Log(string.Format("First Solid to process...", ""));

                foreach (Solid sol in solids)
                {
                    Utils.Log(string.Format("Joining Solids...", ""));

                    s = BooleanOperationsUtils.ExecuteBooleanOperation(s, sol, BooleanOperationsType.Union);

                    Utils.Log(string.Format("Success!", ""));
                }

                if (s != null)
                {
                    Utils.Log(string.Format("One Solid to process...", ""));

                   FreeFormElement.Create(famDoc, s);

                    Utils.Log(string.Format("Free Form Element Created!", ""));
                }
            }

            Utils.Log(string.Format("Mass.TryJoinSolids completed.", ""));
        }

        /// <summary>
        /// Saves the family
        /// </summary>
        /// <param name="famDoc">The family document.</param>
        /// <param name="famPath">The path where to save the family.</param>
        private static void SaveFamily(Document famDoc, string famPath)
        {
            Utils.Log(string.Format("Mass.SaveFamily started...", ""));

            if (famDoc.IsReadOnly)
            {
                Utils.Log(string.Format("ERROR: Family is Read-Only", ""));
                return;
            }
            else
            {
                Utils.Log(string.Format("Family is NOT Read-Only", ""));
            }

            var sao = new SaveAsOptions();
            sao.OverwriteExistingFile = true;
            sao.Compact = true;
            sao.MaximumBackups = 1;

            try
            {
                famDoc.SaveAs(famPath, sao);
            }
            catch (ArgumentNullException ex)
            {
                Utils.Log(string.Format("ArgumentNullException: {0}", ex.Message));
            }
            catch (Autodesk.Revit.Exceptions.OutdatedDirectlyOpenedCentralException ex)
            {
                Utils.Log(string.Format("OutdatedDirectlyOpenedCentralException: {0}", ex.Message));
            }
            catch (Autodesk.Revit.Exceptions.CentralModelException ex)
            {
                Utils.Log(string.Format("CentralModelException: {0}", ex.Message));
            }
            catch (Autodesk.Revit.Exceptions.FileAccessException ex)
            {
                Utils.Log(string.Format("FileAccessException: {0}", ex.Message));
            }
            catch (Autodesk.Revit.Exceptions.FileNotFoundException ex)
            {
                Utils.Log(string.Format("FileNotFoundException: {0}", ex.Message));
            }
            catch (Autodesk.Revit.Exceptions.ForbiddenForDynamicUpdateException ex)
            {
                Utils.Log(string.Format("ForbiddenForDynamicUpdateException: {0}", ex.Message));
            }
            catch (Autodesk.Revit.Exceptions.InsufficientResourcesException ex)
            {
                Utils.Log(string.Format("InsufficientResourcesException: {0}", ex.Message));
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
            {
                Utils.Log(string.Format("InvalidOperationException: {0}", ex.Message));
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException ex)
            {
                Utils.Log(string.Format("OperationCanceledException: {0}", ex.Message));
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException ex)
            {
                Utils.Log(string.Format("ArgumentException: {0}", ex.Message));
            }

            Utils.Log(string.Format("Family Saved {0}", famPath));

            famDoc.Close(false);

            Utils.Log(string.Format("Family Closed!", ""));

            Utils.Log(string.Format("Mass.SaveFamily completed.", ""));
        }

        /// <summary>
        /// Updates the family and the family instance.
        /// </summary>
        /// <param name="famPath">The path to the family to reload.</param>
        /// <param name="rvtFI">The existing Revit Family Instance.</param>
        /// <param name="found">A boolean value that states if the family was existing or created anew.</param>
        private static Revit.Elements.FamilyInstance UpdateFamilyInstance(string famPath, Autodesk.Revit.DB.FamilyInstance rvtFI, bool found)
        {
            Utils.Log(string.Format("Mass.UpdateFamilyInstance started...", ""));

            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.Revit.DB.Family family = null;

            Revit.Elements.FamilyInstance fi = null;

            DocumentManager.Instance.CurrentDBDocument.LoadFamily(famPath, new RevitFamilyLoadOptions(), out family);

            Revit.Elements.FamilyType fs = Revit.Elements.FamilyType.ByFamilyNameAndTypeName(family.Name, family.Name);

            Utils.Log(string.Format("Family Loaded: {0}", family.Id.IntegerValue));

            if (!found && rvtFI == null)
            {
                Utils.Log(string.Format("Creating new Family Instance...", ""));

                Autodesk.DesignScript.Geometry.Point point = Autodesk.DesignScript.Geometry.Point.Origin();

                fi = Revit.Elements.FamilyInstance.ByPoint(fs, point);

                Utils.Log(string.Format("Family Instance Created: {0}", fi.InternalElement.Id.IntegerValue));
            }
            else
            {
                // DocumentManager.Instance.CurrentDBDocument.LoadFamily(famPath, new RevitFamilyLoadOptions(), out family);

                // fs = Revit.Elements.FamilyType.ByFamilyNameAndTypeName(family.Name, family.Name);

                // fi = Revit.Elements.InternalUtilities.ElementQueries.OfFamilyType(fs).First() as Revit.Elements.FamilyInstance;

                fi = Revit.Elements.FamilyInstance.ByFamilyType(fs).First();

                if (fi == null)
                {
                    fi = rvtFI.ToDSType(true) as Revit.Elements.FamilyInstance;

                    Utils.Log(string.Format("Family Query returned null...", ""));
                }

                Utils.Log(string.Format("Family Instance Updated: {0}", rvtFI.Id.IntegerValue));
            }

            TransactionManager.Instance.TransactionTaskDone();

            Utils.Log(string.Format("Mass.UpdateFamilyInstance completed.", ""));

            return fi;
        }

        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Creates a free form mass family by cross sections on the fly and inserts it in the project in Revit local coordinates.
        /// </summary>
        /// <param name="crossSections">The cross sections.</param>
        /// <param name="name">The name.</param>
        /// <param name="familyTemplate">The mass template path.</param>
        /// <param name="append">Append the geometry definition to the current geometry in the Family.</param>
        /// <param name="join">If true attempets to join the geometries.</param>
        /// <param name="rebar">Can host rebar.</param>
        /// <returns></returns>
        public static Revit.Elements.Element ByCrossSections(Autodesk.DesignScript.Geometry.Curve[][] crossSections, string name, string familyTemplate, bool append = false, bool join = false, bool rebar = true)
        {
            Utils.Log(string.Format("Mass.ByCrossSections started...", ""));

            Autodesk.Revit.ApplicationServices.Application app = DocumentManager.Instance.CurrentUIApplication.Application;

            string famName = string.Format("{0}.rfa", name);

            string famPath = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), famName);  // Revit 2020 changed the path to the temp at a session level

            Autodesk.Revit.DB.Family family = null;

            Document famDoc = null;

            Revit.Elements.Element fi = null;

            Autodesk.Revit.DB.FamilyInstance rvtFI = null;

            bool found = false;

            CloseDocument(app, famName);

            int familyFound = GetFamilyDocument(app, famName, familyTemplate, out family, out famDoc, out rvtFI);

            if (familyFound == 0)
            {
                return null;
            }

            if (rvtFI != null)
            {
                found = true;
            }

            if (famDoc != null)
            {
                using (Transaction f = new Transaction(famDoc, "Mass"))
                {
                    var fho = f.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new RevitFailuresPreprocessor());
                    f.SetFailureHandlingOptions(fho);

                    f.Start();

                    try
                    {
                        if (!append)
                        {
                            CleanupFamilyDocument(famDoc);
                        }

                        var output = new List<Solid>();

                        Options opts = new Options();

                        for (int i = 0; i < crossSections.Length - 1; i++)
                        {
                            Utils.Log(string.Format("Processing Cross Section...", ""));

                            var refArrArr = new ReferenceArrayArray();

                            foreach (var profile in new Autodesk.DesignScript.Geometry.Curve[][] { crossSections[i], crossSections[i + 1] })
                            {
                                Utils.Log(string.Format("Processing Profile...", ""));

                                var refArr = new ReferenceArray();

                                Autodesk.DesignScript.Geometry.Point p = null;
                                Autodesk.DesignScript.Geometry.Vector normal = null;

                                foreach (var c in profile)
                                {
                                    Utils.Log(string.Format("Processing Curve...", ""));

                                    var curve = c.ToRevitType();
                                    p = c.StartPoint;
                                    normal = c.Normal;
                                    var sp = Autodesk.Revit.DB.SketchPlane.Create(famDoc, Plane.CreateByNormalAndOrigin(normal.ToXyz(), p.ToXyz()));
                                    var mc = famDoc.FamilyCreate.NewModelCurve(curve, sp);
                                    var r = new Reference(mc);
                                    refArr.Append(r);

                                    Utils.Log(string.Format("Curve Added!", ""));
                                }

                                if (p != null)
                                {
                                    p.Dispose();
                                }

                                if (normal != null)
                                {
                                    normal.Dispose();
                                }

                                refArrArr.Append(refArr);

                                Utils.Log(string.Format("Profile Added!", ""));
                            }

                            var formTemp = famDoc.FamilyCreate.NewLoftForm(true, refArrArr);

                            Utils.Log(string.Format("Loft Created!", ""));

                            foreach (GeometryObject go in formTemp.get_Geometry(opts))
                            {
                                if (go is Solid)
                                {
                                    Solid solid = go as Solid;

                                    if (solid.Volume > 0)
                                    {
                                        Utils.Log(string.Format("Loft Solid Extracted!", ""));
                                    }
                                }
                            }

                        }

                        if (join)
                        {
                            TryJoinSolids(famDoc, output);
                        }

                        if (rebar)
                        {
                            famDoc.OwnerFamily.get_Parameter(BuiltInParameter.FAMILY_CAN_HOST_REBAR).Set(1);
                            Utils.Log(string.Format("Family can host rebar.", ""));
                        }
                    }
                    catch (Exception ex)
                    {
                        Utils.Log(string.Format("ERROR: Mass.ByCrossSections {0}", ex.Message));

                        throw new Exception(string.Format("CivilConnection\nLoft Form failed\n\n{0}", ex.Message));
                    }

                    f.Commit();
                }

                SaveFamily(famDoc, famPath);
            }

            fi = UpdateFamilyInstance(famPath, rvtFI, found);

            Utils.Log(string.Format("Mass.ByCrossSections completed.", ""));

            return fi;
        }

        /// <summary>
        /// Returns a FamilyInstance by a Dynamo solid in Revit local coordinates.
        /// </summary>
        /// <param name="solid">The Dynamo solid in Revit local coordinates.</param>
        /// <param name="name">The name of the family type.</param>
        /// <param name="familyTemplate">the path to the RFT file to use as a template.</param>
        /// <param name="material">The material to assign to the Revit family type.</param>
        /// <param name="isVoid">If true it will create a void that can be used to cut other Revit elements.</param>
        /// <param name="rebar">Can host rebar.</param>
        /// <returns></returns>
        public static Revit.Elements.FamilyInstance BySolid(Autodesk.DesignScript.Geometry.Solid solid, 
            string name, 
            string familyTemplate, 
            Revit.Elements.Material material,
            bool isVoid = false,
            bool rebar = true
            )
        {
            Utils.Log(string.Format("Mass.BySolid started...", ""));

            Autodesk.Revit.DB.Family family = null;

            Revit.Elements.FamilyInstance fi = null;

            Autodesk.Revit.DB.FamilyInstance rvtFI = null;

            Document famDoc = null;

            bool found = false;

            // WARNING: This Dynamo method returns the solid poisitoned at it's bounding box minimum point. Not good for CivilConnection.

            var doc = DocumentManager.Instance.CurrentDBDocument;

            TransactionManager.Instance.ForceCloseTransaction();

            var famName = string.Format("{0}.rfa", name);

            string famPath = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), famName);  // Revit 2020 changed the path to the temp at a session level

            Autodesk.Revit.ApplicationServices.Application app = DocumentManager.Instance.CurrentUIApplication.Application;

            CloseDocument(app, famName);

            int familyFound = GetFamilyDocument(app, famName, familyTemplate, out family, out famDoc, out rvtFI);

            if (familyFound == 0)
            {
                return null;
            }

            if (rvtFI != null)
            {
                found = true;
            }

            if (famDoc != null)
            {
                Utils.Log(string.Format("Start processing...", ""));

                using (Transaction f = new Transaction(famDoc, "Mass"))
                {
                    var fho = f.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new RevitFailuresPreprocessor());
                    f.SetFailureHandlingOptions(fho);

                    f.Start();

                    try
                    {
                        bool newFFE = true;

                        foreach (ElementId eid in new FilteredElementCollector(famDoc).OfClass(typeof(FreeFormElement)).ToElementIds())
                        {
                            FreeFormElement ffe = famDoc.GetElement(eid) as FreeFormElement;

                            if (ffe != null)
                            {
                                foreach (var item in solid.ToRevitType(TessellatedShapeBuilderTarget.Solid, TessellatedShapeBuilderFallback.Abort, material.InternalElement.Id))
                                {
                                    if (item is Solid)
                                    {
                                        Utils.Log(string.Format("Updating Solid...", ""));

                                        Solid s = item as Solid;

                                        ffe.UpdateSolidGeometry(s);

                                        if (isVoid)
                                        {
                                            ffe.Parameters.Cast<Autodesk.Revit.DB.Parameter>().First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.ELEMENT_IS_CUTTING)).Set(1);

                                            famDoc.OwnerFamily.Parameters.Cast<Autodesk.Revit.DB.Parameter>()
                                                .First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.FAMILY_ALLOW_CUT_WITH_VOIDS)).Set(1);
                                        }

                                        newFFE = false;

                                        Utils.Log(string.Format("Solid Updated.", ""));
                                    }

                                    break; // Only the first solid
                                }
                            }
                        }

                        if (newFFE)
                        {
                            foreach (var item in solid.ToRevitType(TessellatedShapeBuilderTarget.Solid, TessellatedShapeBuilderFallback.Abort, material.InternalElement.Id))
                            {
                                // For all the solids

                                if (item is Solid)
                                {
                                    Utils.Log(string.Format("New Solid found...", ""));

                                    Solid s = item as Solid;

                                    Autodesk.Revit.DB.FreeFormElement form = FreeFormElement.Create(famDoc, s);

                                    if (isVoid)
                                    {
                                        form.Parameters.Cast<Autodesk.Revit.DB.Parameter>().First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.ELEMENT_IS_CUTTING)).Set(1);
                                        famDoc.OwnerFamily.Parameters.Cast<Autodesk.Revit.DB.Parameter>().First(x => x.Id.IntegerValue.Equals((int)BuiltInParameter.FAMILY_ALLOW_CUT_WITH_VOIDS)).Set(1);
                                    }

                                    Utils.Log(string.Format("Solid created.", ""));
                                }
                            } 
                        }

                        if (rebar)
                        {
                            famDoc.OwnerFamily.get_Parameter(BuiltInParameter.FAMILY_CAN_HOST_REBAR).Set(1);
                            Utils.Log(string.Format("Family can host rebar.", ""));
                        }
                    }
                    catch (Exception ex)
                    {
                        Utils.Log(string.Format("ERROR: Mass.BySolid {0}", ex.Message));

                        throw new Exception(string.Format("CivilConnection\nLoft Form failed\n\n{0}", ex.Message));
                    }

                    f.Commit();
                }

                SaveFamily(famDoc, famPath);

                Utils.Log(string.Format("Processing completed.", ""));
            }

            fi = UpdateFamilyInstance(famPath, rvtFI, found);

            Utils.Log(string.Format("Mass.BySolid completed.", ""));

            return fi;
        }

        /// <summary>
        /// Creates a free form mass family by cross sections and path on the fly and inserts it in the project in Revit local coordinates.
        /// </summary>
        /// <param name="pathPoints"></param>
        /// <param name="crossSections"></param>
        /// <param name="name"></param>
        /// <param name="familyTemplate"></param>
        /// <param name="append"></param>
        /// <param name="createForm"></param>
        /// <param name="rebar">Can host Rebar.</param>
        /// <returns></returns>
        public static Revit.Elements.FamilyInstance ByPathCrossSections(Autodesk.DesignScript.Geometry.Point[] pathPoints, 
            Autodesk.DesignScript.Geometry.Curve[][] crossSections, 
            string name, string familyTemplate,
            bool append = false, 
            bool createForm = false,
            bool rebar = true)
        {
            Utils.Log(string.Format("Mass.ByPathCrossSections started...", ""));

            TransactionManager.Instance.ForceCloseTransaction();

            string famName = string.Format("{0}.rfa", name);

            string famPath = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), famName);  // Revit 2020 changed the path to the temp at a session level

            Autodesk.Revit.ApplicationServices.Application app = DocumentManager.Instance.CurrentUIApplication.Application;

            Autodesk.Revit.DB.Family family = null;

            Revit.Elements.FamilyInstance fi = null;

            Autodesk.Revit.DB.FamilyInstance rvtFI = null;

            Document famDoc = null;

            bool found = false;

            CloseDocument(app, famName);

            int familyFound = GetFamilyDocument(app, famName, familyTemplate, out family, out famDoc, out rvtFI);

            if (familyFound == 0)
            {
                return null;
            }

            if (rvtFI != null)
            {
                found = true;
            }

            if (famDoc != null)
            {
                using (Transaction f = new Transaction(famDoc, "Mass"))
                {
                    var fho = f.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new RevitFailuresPreprocessor());
                    f.SetFailureHandlingOptions(fho);

                    f.Start();

                    try
                    {
                        if (!append)
                        {
                            CleanupFamilyDocument(famDoc);
                        }

                        // Path
                        var refArrPath = new ReferenceArray();

                        var refPointArray = new ReferencePointArray();

                        foreach (var p in pathPoints)
                        {
                            XYZ point = p.ToRevitType();

                            refPointArray.Append(famDoc.FamilyCreate.NewReferencePoint(point));
                        }

                        var mcPath = famDoc.FamilyCreate.NewCurveByPoints(refPointArray);

                        refArrPath.Append(new Reference(mcPath));

                        var output = new List<Autodesk.DesignScript.Geometry.Solid>();

                        var toDel = new List<ElementId>();

                        Options opts = new Options();


                        // Profiles

                        Autodesk.DesignScript.Geometry.Point pt = null;
                        Autodesk.DesignScript.Geometry.Vector normal = null;

                        var refArrArr = new ReferenceArrayArray();

                        foreach (var profile in crossSections)
                        {
                            var refArr = new ReferenceArray();

                            foreach (var c in profile)
                            {
                                pt = c.StartPoint;
                                normal = c.Normal;
                                var curve = c.ToRevitType();
                                var sp = Autodesk.Revit.DB.SketchPlane.Create(famDoc, Plane.CreateByNormalAndOrigin(normal.ToXyz(), pt.ToXyz()));
                                var mc = famDoc.FamilyCreate.NewModelCurve(curve, sp);
                                var r = new Reference(mc);
                                refArr.Append(r);
                            }

                            refArrArr.Append(refArr);
                        }

                        if (pt != null)
                        {
                            pt.Dispose();
                        }

                        if (normal != null)
                        {
                            normal.Dispose();
                        }

                        if (createForm)
                        {
                            var formTemp = famDoc.FamilyCreate.NewSweptBlendForm(true, refArrPath, refArrArr);
                        }

                        if (rebar)
                        {
                            famDoc.OwnerFamily.get_Parameter(BuiltInParameter.FAMILY_CAN_HOST_REBAR).Set(1);
                            Utils.Log(string.Format("Family can host rebar.", ""));
                        }
                    }
                    catch (Exception ex)
                    {
                        Utils.Log(string.Format("ERROR: Mass.ByPathCrossSections {0}", ex.Message));

                        throw new Exception(string.Format("CivilConnection\nLoft Form failed\n\n{0}", ex.Message));
                    }

                    f.Commit();
                }

                SaveFamily(famDoc, famPath);
            }

            fi = UpdateFamilyInstance(famPath, rvtFI, found);

            Utils.Log(string.Format("Mass.ByPathCrossSections completed.", ""));

            return fi;
        }


        /// <summary>
        /// Creates a free form mass family by cross sections on the fly and inserts it in the project in Revit local coordinates.
        /// </summary>
        /// <param name="crossSections">The cross sections.</param>
        /// <param name="name">The name.</param>
        /// <param name="familyTemplate">The mass template path.</param>
        /// <param name="append">Append the geoemtry definition to the current geometry in the Family.</param>
        /// <param name="rebar">Can host rebar.</param>
        /// <returns></returns>
        public static Revit.Elements.Element ByLoftCrossSections(Autodesk.DesignScript.Geometry.Curve[][] crossSections, string name, string familyTemplate, bool append = false, bool rebar = true)
        {
            Utils.Log(string.Format("Mass.ByLoftCrossSections started...", ""));

            TransactionManager.Instance.ForceCloseTransaction();

            string famName = string.Format("{0}.rfa", name);

            string famPath = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), famName);  // Revit 2020 changed the path to the temp at a session level

            Autodesk.Revit.ApplicationServices.Application app = DocumentManager.Instance.CurrentUIApplication.Application;

            Autodesk.Revit.DB.Family family = null;

            Revit.Elements.FamilyInstance fi = null;

            Autodesk.Revit.DB.FamilyInstance rvtFI = null;

            Document famDoc = null;

            bool found = false;

            CloseDocument(app, famName);

            int familyFound = GetFamilyDocument(app, famName, familyTemplate, out family, out famDoc, out rvtFI);

            if (familyFound == 0)
            {
                return null;
            }

            if (rvtFI != null)
            {
                found = true;
            }

            if (famDoc != null)
            {
                using (Transaction f = new Transaction(famDoc, "Mass"))
                {
                    var fho = f.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new RevitFailuresPreprocessor());
                    f.SetFailureHandlingOptions(fho);

                    f.Start();

                    try
                    {
                        if (!append)
                        {
                            CleanupFamilyDocument(famDoc);
                        }

                        var output = new List<Solid>();

                        var toDel = new List<ElementId>();

                        Options opts = new Options();

                        var profiles = new List<CurveLoop>();

                        for (int i = 0; i < crossSections.Length; i++)
                        {
                            Utils.Log(string.Format("Processing Cross Section...", ""));

                            var profile = new CurveLoop();

                            foreach (var c in crossSections[i])
                            {
                                var curve = c.ToRevitType();
                                profile.Append(curve);
                            }

                            profiles.Add(profile);

                            Utils.Log(string.Format("Profile Added!", ""));
                        }

                        ElementId gsid = new FilteredElementCollector(famDoc)
                            .OfClass(typeof(GraphicsStyle))
                            .WhereElementIsNotElementType()
                            .Cast<GraphicsStyle>()
                            .First(x => x.GraphicsStyleType == GraphicsStyleType.Projection)
                            .Id;

                        ElementId matid = new FilteredElementCollector(famDoc)
                            .OfClass(typeof(Autodesk.Revit.DB.Material))
                            .WhereElementIsNotElementType()
                            .Cast<Autodesk.Revit.DB.Material>()
                            .FirstOrDefault()
                            .Id;

                        Solid solid = GeometryCreationUtilities.CreateLoftGeometry(profiles, new SolidOptions(gsid, matid));

                        Autodesk.Revit.DB.FreeFormElement form = FreeFormElement.Create(famDoc, solid);

                        Utils.Log(string.Format("Loft Geometries created.", ""));

                        if (rebar)
                        {
                            famDoc.OwnerFamily.get_Parameter(BuiltInParameter.FAMILY_CAN_HOST_REBAR).Set(1);
                            Utils.Log(string.Format("Family can host rebar.", ""));
                        }

                    }
                    catch (Exception ex)
                    {
                        Utils.Log(string.Format("ERROR: Mass.ByLoftCrossSections {0}", ex.Message));

                        throw new Exception(string.Format("CivilConnection\nLoft Form failed\n\n{0}", ex.Message));
                    }

                    f.Commit();
                }

                SaveFamily(famDoc, famPath);
            }

            fi = UpdateFamilyInstance(famPath, rvtFI, found);

            Utils.Log(string.Format("Mass.ByLoftCrossSections completed.", ""));

            return fi;
        }

        // TODO: Check if the family template supports FreeFormElements

        /// <summary>
        /// Creates a free form mass family by cross sections on the fly and inserts it in the project in Revit local coordinates.
        /// </summary>
        /// <param name="shapes">The AppliedSubassemblyShape that represents the cross sections.</param>
        /// <param name="stations">The sequence of stations that defines the creases along the alignment. If null, the loft will be continuous.</param>
        /// <param name="name">The name.</param>
        /// <param name="familyTemplate">The mass template path.</param>
        /// <param name="append">Append the geoemtry definition to the current geometry in the Family.</param>
        /// <param name="rebar">Can host rebar.</param>
        /// <returns></returns>
        public static Revit.Elements.Element ByShapesCreaseStations(string familyTemplate, string name, AppliedSubassemblyShape[] shapes, double[] stations = null, bool append = false, bool rebar = true)
        {
            #region FAIL GRACEFULLY

            Utils.Log(string.Format("Mass.ByShapesCreaseStations started...", ""));

            // Check that the shapes share the same name
            if (shapes.GroupBy(x => x.Name).Count() > 1)
            {
                string message = "The AppliedSubassemblyShape must have the same name.";

                Utils.Log(string.Format("ERROR: Mass.ByShapesCreaseStations {0}", message));

                throw new Exception(message);
            }

            shapes = shapes.OrderBy(x => x.Station).ToArray();  // make sure the shapes are sorted by station

            shapes = shapes.GroupBy(x => Math.Round(x.Station, 8)).Select(g => g.First()).ToArray();  // make sure there are no overalpping shapes
            
            #endregion

            #region FAMILY HOUSEKEEPING

            TransactionManager.Instance.ForceCloseTransaction();

            string famName = string.Format("{0}.rfa", name);

            string famPath = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), famName);  // Revit 2020 changed the path to the temp at a session level

            Autodesk.Revit.ApplicationServices.Application app = DocumentManager.Instance.CurrentUIApplication.Application;

            Autodesk.Revit.DB.Family family = null;

            Revit.Elements.Element fi = null;

            Autodesk.Revit.DB.FamilyInstance rvtFI = null;

            bool found = false;

            Document famDoc = null;

            CloseDocument(app, famName);

            int familyFound = GetFamilyDocument(app, famName, familyTemplate, out family, out famDoc, out rvtFI);

            if (familyFound == 0)
            {
                return null;
            }

            if (rvtFI != null)
            {
                found = true;
            }

            #endregion

            #region GEOMETRY COMPUTATION

            if (famDoc != null)
            {
                using (Transaction f = new Transaction(famDoc, "Mass"))
                {
                    var fho = f.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new RevitFailuresPreprocessor());
                    f.SetFailureHandlingOptions(fho);

                    f.Start();

                    #region DELETE EXISTING GEOMETRY
                    try
                    {
                        if (!append)
                        {
                            CleanupFamilyDocument(famDoc);
                        }
                    }
                    catch (Exception ex)
                    {
                        Utils.Log(string.Format("ERROR: Mass.ByShapesCreaseStations {0}", ex.Message));

                        throw new Exception(string.Format("CivilConnection\nLoft Form failed\n\n{0}", ex.Message));
                    }

                    #endregion

                    var output = new List<Solid>();

                    Options opts = new Options();

                    // Only the shapes that satisfy the stations pairs will be processed
                    IList<IList<AppliedSubassemblyShape>> processShapes = new List<IList<AppliedSubassemblyShape>>();

                    if (stations != null || stations.Length > 1)
                    {
                        stations = stations.OrderBy(x => x).ToArray();

                        for (int i = 0; i < stations.Length - 1; ++i)
                        {
                            double min = stations[i];
                            double max = stations[i + 1];

                            IList<AppliedSubassemblyShape> list = new List<AppliedSubassemblyShape>();

                            foreach (AppliedSubassemblyShape sh in shapes)
                            {
                                if ((Math.Round(min, 6) <= Math.Round(sh.Station, 3) && Math.Round(sh.Station, 6) <= Math.Round(max, 3)) ||
                                    (Math.Abs(sh.Station - min) < 0.0001) || (Math.Abs(sh.Station - max) < 0.0001))
                                {
                                    list.Add(sh);
                                }
                            }

                            if (list.Count > 1)
                            {
                                Utils.Log(string.Format("Min Station {0}", min));
                                Utils.Log(string.Format("Max Station {0}", max));

                                processShapes.Add(list.OrderBy(x => x.Station).ToList());
                            }
                        }
                    }
                    else
                    {
                        processShapes.Add(shapes.ToList());
                    }

                    Utils.Log(string.Format("Segments ready: {0}", processShapes.Count));

                    var cs = RevitUtils.DocumentTotalTransform();

                    IList<GenericForm> familyForms = new List<GenericForm>();  // to make it accept FreeFormElements and Forms at the same time

                    foreach (var segment in processShapes)
                    {
                        var profiles = new List<CurveLoop>();

                        try
                        {
                            Utils.Log(string.Format("Min Station {0}", segment.Min(x => x.Station)));
                            Utils.Log(string.Format("Max Station {0}", segment.Max(x => x.Station)));

                            var crossSections = segment.Select(x => x.Geometry.Transform(cs).Explode().Select(c => c as Autodesk.DesignScript.Geometry.Curve)).ToArray();

                            Utils.Log(string.Format("Cross Sections ready: {0}", crossSections.Length));

                            if (crossSections.Length > 1)
                            {
                                for (int i = 0; i < crossSections.Length; i++)
                                {
                                    Utils.Log(string.Format("Processing Cross Section: {0}", crossSections[i].Count()));

                                    var profile = new CurveLoop();

                                    foreach (var c in crossSections[i])
                                    {
                                        var curve = c.ToRevitType();
                                        profile.Append(curve);
                                    }

                                    profiles.Add(profile);

                                    Utils.Log(string.Format("Profile Added!", ""));
                                }

                                ElementId gsid = new FilteredElementCollector(famDoc)
                                    .OfClass(typeof(GraphicsStyle))
                                    .WhereElementIsNotElementType()
                                    .Cast<GraphicsStyle>()
                                    .First(x => x.GraphicsStyleType == GraphicsStyleType.Projection)
                                    .Id;

                                ElementId matid = new FilteredElementCollector(famDoc)
                                    .OfClass(typeof(Autodesk.Revit.DB.Material))
                                    .WhereElementIsNotElementType()
                                    .Cast<Autodesk.Revit.DB.Material>()
                                    .FirstOrDefault()
                                    .Id;

                                if (profiles.Count > 1)
                                {
                                    try
                                    {
                                        Solid solid = GeometryCreationUtilities.CreateLoftGeometry(profiles, new SolidOptions(gsid, matid));

                                        Autodesk.Revit.DB.FreeFormElement form = FreeFormElement.Create(famDoc, solid);

                                        familyForms.Add(form);

                                        // Attempt LoftForm it is failing after creating the first model curve
                                        // Possisbly need to regenerate to get the mode curve geometries

                                        Utils.Log(string.Format("Form Created!", ""));
                                    }
                                    catch (Exception ex)
                                    {
                                        Utils.Log(string.Format("ERROR 1: Mass.ByShapesCreaseStations Profiles {0}", ex.Message));

                                        foreach (CurveLoop cl in profiles)
                                        {
                                            CreateModelCurves(famDoc, cl);
                                        }

                                        // 20190610 -- START

                                        for (int i = 0; i < profiles.Count - 1; i++)
                                        {
                                            Utils.Log(string.Format("Processing Cross Section...", ""));

                                            var refArrArr = new ReferenceArrayArray();

                                            foreach (var profile in new CurveLoop[] { profiles[i], profiles[i + 1] })
                                            {
                                                Utils.Log(string.Format("Processing Profile...", ""));

                                                var refArr = new ReferenceArray();

                                                foreach (var curve in profile)
                                                {
                                                    Utils.Log(string.Format("Processing Curve...", ""));

                                                    var sp = Autodesk.Revit.DB.SketchPlane.Create(famDoc, profile.GetPlane());
                                                    var mc = famDoc.FamilyCreate.NewModelCurve(curve, sp);
                                                    var r = new Reference(mc);
                                                    refArr.Append(r);

                                                    Utils.Log(string.Format("Curve Added!", ""));
                                                }

                                                refArrArr.Append(refArr);

                                                Utils.Log(string.Format("Profile Added!", ""));
                                            }

                                            var formTemp = famDoc.FamilyCreate.NewLoftForm(true, refArrArr);

                                            Utils.Log(string.Format("Loft Created!", ""));

                                            foreach (GeometryObject go in formTemp.get_Geometry(opts))
                                            {
                                                if (go is Solid)
                                                {
                                                    Solid solid = go as Solid;

                                                    if (solid.Volume > 0)
                                                    {
                                                        Utils.Log(string.Format("Loft Solid Extracted!", ""));
                                                    }
                                                }
                                            }
                                        }

                                        // 20190610 -- END
                                    }
                                }
                                else
                                {
                                    Utils.Log(string.Format("ERROR 2: Not enough profiles for a loft", ""));
                                }
                            }
                            else
                            {
                                Utils.Log(string.Format("ERROR 3: Not enough profiles for a loft", ""));
                            }

                            // Dispose 20191118
                            foreach (var loop in crossSections)
                            {
                                foreach (var item in loop)
                                {
                                    if (item != null)
                                    {
                                        item.Dispose();
                                    }
                                }
                            }

                            if (rebar)
                            {
                                famDoc.OwnerFamily.get_Parameter(BuiltInParameter.FAMILY_CAN_HOST_REBAR).Set(1);
                                Utils.Log(string.Format("Family can host rebar.", ""));
                            }
                        }
                        catch (Exception ex)
                        {
                            Utils.Log(string.Format("ERROR 4: Mass.ByShapesCreaseStations {0}", ex.Message));
                        }
                    }

                    if (familyForms.Count > 0)
                    {
                        if (familyForms.Count == processShapes.Count)
                        {
                            Utils.Log(string.Format("All Loft Geometries were created.", ""));
                        }
                        else
                        {
                            Utils.Log(string.Format("ERROR 5: Some Loft Geometries were created, but some were not!", ""));
                        }
                    }

                    f.Commit();
                }

                SaveFamily(famDoc, famPath);

            }

            #endregion

            #region FAMILY LOADING AND PLACEMENT
            
            fi = UpdateFamilyInstance(famPath, rvtFI, found);

            Utils.Log(string.Format("Mass.ByShapesCreaseStations completed.", ""));

            return fi;

            #endregion
        }

        /// <summary>
        /// Creates a free form mass family by cross sections on the fly and inserts it in the project in Revit local coordinates.
        /// </summary>
        /// <param name="closedCurves">The closed curves that represents the cross sections.</param>
        /// <param name="stations">The sequence of stations that defines the creases along the alignment. If null, the loft will be continuous.</param>
        /// <param name="name">The name.</param>
        /// <param name="familyTemplate">The mass template path.</param>
        /// <param name="alignment">The alignment used to calculate the stations.</param>
        /// <param name="append">Append the geoemtry definition to the current geometry in the Family.</param>
        /// <param name="rebar">Can host rebar.</param>
        /// <returns></returns>
        public static Revit.Elements.Element ByClosedCurvesCreaseStations(
            string familyTemplate, 
            string name, 
            Alignment alignment, 
            Autodesk.DesignScript.Geometry.PolyCurve[] closedCurves, 
            double[] stations = null, 
            bool append = false,
            bool rebar = true)
        {
            #region FAIL GRACEFULLY

            Utils.Log(string.Format("Mass.ByClosedCurvesCreaseStations started...", ""));

            // Check that the curves are planar
            if (!closedCurves.All(x => x.IsPlanar))
            {
                string message = "The curves must be planar.";

                Utils.Log(string.Format("ERROR: Mass.ByClosedCurvesCreaseStations {0}", message));

                throw new Exception(message);
            }

            var aeccAlignment = alignment.InternalElement as Autodesk.AECC.Interop.Land.AeccAlignment;

            closedCurves = closedCurves.OrderBy(x => alignment.GetStationOffsetElevation(x.StartPoint)["Station"]).ToArray();  // make sure the shapes are sorted by station

            #endregion

            #region FAMILY HOUSEKEEPING

            TransactionManager.Instance.ForceCloseTransaction();

            string famName = string.Format("{0}.rfa", name);

            string famPath = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), famName);  // Revit 2020 changed the path to the temp at a session level

            Autodesk.Revit.ApplicationServices.Application app = DocumentManager.Instance.CurrentUIApplication.Application;

            Autodesk.Revit.DB.Family family = null;

            Revit.Elements.Element fi = null;

            Autodesk.Revit.DB.FamilyInstance rvtFI = null;

            bool found = false;

            Document famDoc = null;

            CloseDocument(app, famName);

            int familyFound = GetFamilyDocument(app, famName, familyTemplate, out family, out famDoc, out rvtFI);

            if (familyFound == 0)
            {
                return null;
            }

            if (rvtFI != null)
            {
                found = true;
            }

            #endregion

            #region GEOMETRY COMPUTATION

            if (famDoc != null)
            {
                using (Transaction f = new Transaction(famDoc, "Mass"))
                {
                    var fho = f.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new RevitFailuresPreprocessor());
                    f.SetFailureHandlingOptions(fho);

                    f.Start();

                    #region DELETE EXISTING GEOMETRY
                    try
                    {
                        if (!append)
                        {
                            CleanupFamilyDocument(famDoc);
                        }
                    }
                    catch (Exception ex)
                    {
                        Utils.Log(string.Format("ERROR: Mass.ByClosedCurvesCreaseStations {0}", ex.Message));

                        throw new Exception(string.Format("CivilConnection\nLoft Form failed\n\n{0}", ex.Message));
                    }

                    #endregion

                    var output = new List<Solid>();

                    Options opts = new Options();

                    // Only the shapes that satisfy the stations pairs will be processed
                    IList<IList<Autodesk.DesignScript.Geometry.PolyCurve>> processShapes = new List<IList<Autodesk.DesignScript.Geometry.PolyCurve>>();

                    if (stations != null || stations.Length > 1)
                    {
                        stations = stations.OrderBy(x => x).ToArray();

                        for (int i = 0; i < stations.Length - 1; ++i)
                        {
                            double min = Math.Round(stations[i], 9);
                            double max = Math.Round(stations[i + 1], 9);

                            IList<Autodesk.DesignScript.Geometry.PolyCurve> list = new List<Autodesk.DesignScript.Geometry.PolyCurve>();

                            double previous = -2 * min;

                            foreach (Autodesk.DesignScript.Geometry.PolyCurve sh in closedCurves)
                            {
                                double station = Convert.ToDouble(alignment.GetStationOffsetElevation(sh.StartPoint)["Station"]);

                                if ((Math.Round(min, 6) <= Math.Round(station, 3) && Math.Round(station, 6) <= Math.Round(max, 3)) ||
                                   (Math.Abs(station - min) < 0.0001) || (Math.Abs(station - max) < 0.0001))
                                {
                                    if (Math.Abs(previous - station) > 0.00001)
                                    {
                                        list.Add(sh);
                                        previous = Math.Round(station, 5);
                                    }
                                }
                            }

                            if (list.Count > 1)
                            {
                                Utils.Log(string.Format("Min Station {0}", min));
                                Utils.Log(string.Format("Max Station {0}", max));

                                processShapes.Add(list.OrderBy(x => alignment.GetStationOffsetElevation(x.StartPoint)["Station"]).ToList());
                            }
                        }
                    }
                    else
                    {
                        processShapes.Add(closedCurves.ToList());
                    }

                    Utils.Log(string.Format("Segments ready: {0}", processShapes.Count));

                    var cs = RevitUtils.DocumentTotalTransform();

                    IList<GenericForm> familyForms = new List<GenericForm>();  // to make it accept FreeFormElements and Forms at the same time

                    foreach (var segment in processShapes)
                    {
                        var profiles = new List<CurveLoop>();

                        try
                        {
                            var crossSections = segment.Select(x => x.Transform(cs).Explode().Select(c => c as Autodesk.DesignScript.Geometry.Curve)).ToArray();

                            Utils.Log(string.Format("Cross Sections ready: {0}", crossSections.Length));

                            if (crossSections.Length > 1)
                            {
                                for (int i = 0; i < crossSections.Length; i++)
                                {
                                    Utils.Log(string.Format("Processing Cross Section: {0}", crossSections[i].Count()));

                                    var profile = new CurveLoop();

                                    foreach (var c in crossSections[i])
                                    {
                                        var curve = c.ToRevitType();
                                        profile.Append(curve);
                                    }

                                    profiles.Add(profile);

                                    Utils.Log(string.Format("Profile Added!", ""));
                                }

                                ElementId gsid = new FilteredElementCollector(famDoc)
                                    .OfClass(typeof(GraphicsStyle))
                                    .WhereElementIsNotElementType()
                                    .Cast<GraphicsStyle>()
                                    .First(x => x.GraphicsStyleType == GraphicsStyleType.Projection)
                                    .Id;

                                ElementId matid = new FilteredElementCollector(famDoc)
                                    .OfClass(typeof(Autodesk.Revit.DB.Material))
                                    .WhereElementIsNotElementType()
                                    .Cast<Autodesk.Revit.DB.Material>()
                                    .FirstOrDefault()
                                    .Id;

                                if (profiles.Count > 1)
                                {
                                    try
                                    {
                                        Solid solid = GeometryCreationUtilities.CreateLoftGeometry(profiles, new SolidOptions(gsid, matid));

                                        Autodesk.Revit.DB.FreeFormElement form = FreeFormElement.Create(famDoc, solid);

                                        familyForms.Add(form);

                                        // Attempt LoftForm it is failing after creating the first model curve
                                        // Possibly need to regenerate to get the mode curve geometries

                                        Utils.Log(string.Format("Form Created!", ""));
                                    }
                                    catch (Exception ex)
                                    {
                                        Utils.Log(string.Format("ERROR 1: Mass.ByClosedCurvesCreaseStations Profiles {0}", ex.Message));

                                        foreach (CurveLoop cl in profiles)
                                        {
                                            CreateModelCurves(famDoc, cl);
                                        }

                                        // 20190610 -- START

                                        for (int i = 0; i < profiles.Count - 1; i++)
                                        {
                                            Utils.Log(string.Format("Processing Cross Section...", ""));

                                            var refArrArr = new ReferenceArrayArray();

                                            foreach (var profile in new CurveLoop[] { profiles[i], profiles[i + 1] })
                                            {
                                                Utils.Log(string.Format("Processing Profile...", ""));

                                                var refArr = new ReferenceArray();

                                                foreach (var curve in profile)
                                                {
                                                    Utils.Log(string.Format("Processing Curve...", ""));

                                                    var sp = Autodesk.Revit.DB.SketchPlane.Create(famDoc, profile.GetPlane());
                                                    var mc = famDoc.FamilyCreate.NewModelCurve(curve, sp);
                                                    var r = new Reference(mc);
                                                    refArr.Append(r);

                                                    Utils.Log(string.Format("Curve Added!", ""));
                                                }

                                                refArrArr.Append(refArr);

                                                Utils.Log(string.Format("Profile Added!", ""));
                                            }

                                            var formTemp = famDoc.FamilyCreate.NewLoftForm(true, refArrArr);

                                            Utils.Log(string.Format("Loft Created!", ""));

                                            foreach (GeometryObject go in formTemp.get_Geometry(opts))
                                            {
                                                if (go is Solid)
                                                {
                                                    Solid solid = go as Solid;

                                                    if (solid.Volume > 0)
                                                    {
                                                        Utils.Log(string.Format("Loft Solid Extracted!", ""));
                                                    }
                                                }
                                            }

                                        }

                                        // 20190610 -- END
                                    }
                                }
                                else
                                {
                                    Utils.Log(string.Format("ERROR 2: Not enough profiles for a loft", ""));
                                }

                                // Dispose 20191118
                                foreach (var loop in crossSections)
                                {
                                    foreach (var item in loop)
                                    {
                                        if (item != null)
                                        {
                                            item.Dispose();
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Utils.Log(string.Format("ERROR 3: Not enough profiles for a loft", ""));
                            }

                            if (rebar)
                            {
                                famDoc.OwnerFamily.get_Parameter(BuiltInParameter.FAMILY_CAN_HOST_REBAR).Set(1);
                                Utils.Log(string.Format("Family can host rebar.", ""));
                            }
                        }
                        catch (Exception ex)
                        {
                            Utils.Log(string.Format("ERROR 4: Mass.ByClosedCurvesCreaseStations {0}", ex.Message));
                        }
                    }

                    if (familyForms.Count > 0)
                    {
                        if (familyForms.Count == processShapes.Count)
                        {
                            Utils.Log(string.Format("All Loft Geometries were created.", ""));
                        }
                        else
                        {
                            Utils.Log(string.Format("ERROR 5: Some Loft Geometries were created, but some were not!", ""));
                        }
                    }

                    f.Commit();

                    //cs.Dispose();
                }

                SaveFamily(famDoc, famPath);
            }

            #endregion

            #region FAMILY LOADING AND PLACEMENT

            fi = UpdateFamilyInstance(famPath, rvtFI, found);

            Utils.Log(string.Format("Mass.ByClosedCurvesCreaseStations completed.", ""));

            return fi;

            #endregion
        }

        #endregion
    }
}
