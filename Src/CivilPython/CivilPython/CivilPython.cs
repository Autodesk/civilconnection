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

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using IronPython.Hosting;
using Microsoft.Scripting.Hosting;
using System.Reflection;
using WF = System.Windows.Forms;
using Autodesk.Windows;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing;
using System.Xml;
using Autodesk.Civil.ApplicationServices;


namespace CivilPython
{
    public class CivilPython
    {
        private void PythonConsole(bool cmdLine)
        {
            Dictionary<string, object> options = new Dictionary<string, object>();
            options["Debug"] = true;

            var ipy = Python.CreateRuntime(options);
            var engine = Python.GetEngine(ipy);

            try
            {
                string path = "";

                System.Version version = Application.Version;

                string ver = version.ToString();

                string release = "2020";

                // 20190910 --- Rob Todd Suggestion to catch AutoCAD installation different from C drive

                release = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetProcessesByName("acad").FirstOrDefault().MainModule.FileName);

                //switch (ver)
                //{
                //    case "20":
                //        release = "2016";
                //        break;
                //    case "21.0.0.0":
                //        release = "2017";
                //        break;
                //    case "22.0.0.0":
                //        release = "2018";
                //        break;
                //    case "23.0.0.0":
                //        release = "2019";
                //        break;
                //    case "23.1.0.0":
                //        release = "2020";
                //        break;
                //}

                if (!cmdLine)
                {
                    WF.OpenFileDialog ofd = new WF.OpenFileDialog();
                    ofd.Filter = "Python Script (*.py) | *.py";

                    if (ofd.ShowDialog() == WF.DialogResult.OK)
                    {
                        path = ofd.FileName;
                    }
                }
                else
                {
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    Editor ed = doc.Editor;

                    short fd = (short)Application.GetSystemVariable("FILEDIA");

                    // Ask the user to select a .py file

                    PromptOpenFileOptions pfo = new PromptOpenFileOptions("Select Python script to load");
                    pfo.Filter = "Python script (*.py)|*.py";
                    pfo.PreferCommandLine = (cmdLine || fd == 0);
                    PromptFileNameResult pr = ed.GetFileNameForOpen(pfo);
                    path = pr.StringResult;
                }

                if (path != "")
                {
                    //ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\acmgd.dll", release)));
                    //ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\acdbmgd.dll", release)));
                    //ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\accoremgd.dll", release)));
                    //ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\ACA\AecBaseMgd.dll", release)));
                    //ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\ACA\AecPropDataMgd.dll", release)));
                    //ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\C3D\AeccDbMgd.dll", release)));
                    //ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\C3D\AeccPressurePipesMgd.dll", release)));
                    //ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\acdbmgdbrep.dll", release)));

                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"{0}\acmgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"{0}\acdbmgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"{0}\accoremgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"{0}\ACA\AecBaseMgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"{0}\ACA\AecPropDataMgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"{0}\C3D\AeccDbMgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"{0}\C3D\AeccPressurePipesMgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"{0}\acdbmgdbrep.dll", release)));

                    ScriptSource source = engine.CreateScriptSourceFromFile(path);
                    CompiledCode compiledCode = source.Compile();
                    ScriptScope scope = engine.CreateScope();
                    compiledCode.Execute(scope);
                }
            }
            catch (System.Exception ex)
            {
                string message = engine.GetService<ExceptionOperations>().FormatException(ex);
                WF.MessageBox.Show(message);
            }

            return;
        }

        [CommandMethod("Python")]
        public void PythonLoadUI()
        {
            PythonConsole(false);
        }

        [CommandMethod("-Python")]
        public void PythonCmdLine()
        {
            PythonConsole(true);
        }

        [CommandMethod("-ReplaceSolid")]
        public void PythonScriptCmdLine()
        {
            // PythonScript(true);
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.DatabaseServices.Database db = doc.Database;
            Editor ed = doc.Editor;

            using (doc.LockDocument())
            {
                using (db)
                {
                    using (Transaction t = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = t.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = t.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        short fd = (short)Application.GetSystemVariable("FILEDIA");

                        PromptStringOptions pso = new PromptStringOptions("Insert First Handle");
                        pso.AllowSpaces = true;

                        PromptResult pr = ed.GetString(pso);
                        string handle1 = pr.StringResult.Replace("\"", "");

                        pso = new PromptStringOptions("Insert Second Handle");
                        pso.AllowSpaces = true;

                        pr = ed.GetString(pso);
                        string handle2 = pr.StringResult.Replace("\"", "");

                        ObjectId oid1 = db.GetObjectId(false, new Handle(Convert.ToInt64(handle1, 16)), 0);
                        ObjectId oid2 = db.GetObjectId(false, new Handle(Convert.ToInt64(handle2, 16)), 0);

                        DBObject ent1 = t.GetObject(oid1, OpenMode.ForWrite);
                        DBObject ent2 = t.GetObject(oid2, OpenMode.ForWrite);

                        ent2.SwapIdWith(oid1, true, true);

                        ent1.Erase();

                        Application.SetSystemVariable("FILEDIA", fd);

                        t.Commit();
                    }
                }
            }
        }

        [CommandMethod("-ExportLandFeatureLinesToXml")]
        public void ExportLandFeatureLinesToXml()
        {
            string path = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "LandFeatureLinesReport.xml");

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            XmlDocument xmlDoc = new XmlDocument();

            XmlElement docElement = xmlDoc.CreateElement("Document");
            xmlDoc.AppendChild(docElement);

            XmlElement featurelines = xmlDoc.CreateElement("FeatureLines");
            docElement.AppendChild(featurelines);

            Document doc = Application.DocumentManager.MdiActiveDocument;
            docElement.SetAttribute("Name", doc.Name);

            using (doc.LockDocument())
            {
                using (Database db = doc.Database)
                {
                    using (Transaction t = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = t.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = t.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                        foreach (ObjectId oid in btr)
                        {
                            DBObject obj = t.GetObject(oid, OpenMode.ForRead);

                            if ((obj is Autodesk.Civil.DatabaseServices.FeatureLine))
                            {
                                Autodesk.Civil.DatabaseServices.FeatureLine f = obj as Autodesk.Civil.DatabaseServices.FeatureLine;

                                XmlElement featureline = xmlDoc.CreateElement("FeatureLine");
                                featureline.SetAttribute("Name", f.Name);
                                if (f.StyleName != null)
                                {
                                    featureline.SetAttribute("Style", f.StyleName);
                                }
                                else
                                {
                                    featureline.SetAttribute("Style", "");
                                }
                                featureline.SetAttribute("Handle", f.Handle.ToString());
                                featurelines.AppendChild(featureline);

                                XmlElement points = xmlDoc.CreateElement("Points");
                                featureline.AppendChild(points);

                                foreach (Autodesk.AutoCAD.Geometry.Point3d p3d in f.GetPoints(Autodesk.Civil.FeatureLinePointType.AllPoints))
                                {
                                    XmlElement point = xmlDoc.CreateElement("Point");
                                    point.SetAttribute("X", p3d.X.ToString());
                                    point.SetAttribute("Y", p3d.Y.ToString());
                                    point.SetAttribute("Z", p3d.Z.ToString());
                                    points.AppendChild(point);
                                }
                            }
                        }
                    }
                }
            }

            xmlDoc.Save(path);
        }

        [CommandMethod("-ExportCorridorFeatureLinesToXML")]
        public void ExportCorridorFeatureLinesToXml()
        {
            string path = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "CorridorFeatureLines.xml");

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            XmlDocument xmlDoc = new XmlDocument();

            XmlElement docElement = xmlDoc.CreateElement("Document");
            xmlDoc.AppendChild(docElement);

            XmlElement corridors = xmlDoc.CreateElement("Corridors");
            docElement.AppendChild(corridors);

            Document doc = Application.DocumentManager.MdiActiveDocument;
            CivilDocument cdoc = CivilApplication.ActiveDocument;

            docElement.SetAttribute("Name", doc.Name);

            short fd = (short)Application.GetSystemVariable("FILEDIA");

            PromptStringOptions pso = new PromptStringOptions("\nInsert Corridor Handle");
            pso.AllowSpaces = false;

            PromptResult pr = doc.Editor.GetString(pso);
            string handle1 = pr.StringResult.Replace("\"", "");

            Application.SetSystemVariable("FILEDIA", fd);

            using (doc.LockDocument())
            {
                using (Database db = doc.Database)
                {
                    using (Transaction t = db.TransactionManager.StartTransaction())
                    {
                        ObjectId oid1 = ObjectId.Null;

                        try
                        {
                            if (!string.IsNullOrEmpty(handle1) && !string.IsNullOrWhiteSpace(handle1))
                            {
                                oid1 = db.GetObjectId(false, new Handle(Convert.ToInt64(handle1, 16)), 0);
                            }
                        }
                        catch { }

                        if (oid1 == ObjectId.Null)
                        {
                            foreach (ObjectId oid in cdoc.CorridorCollection)
                            {
                                if (oid == ObjectId.Null)
                                {
                                    continue;
                                }

                                Autodesk.Civil.DatabaseServices.Corridor corr = null;

                                try
                                {
                                    corr = t.GetObject(oid, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Corridor;
                                }
                                catch (System.Exception ex)
                                {
                                    System.Windows.Forms.MessageBox.Show(string.Format("ERROR: {0}", ex.Message));
                                }

                                if (corr == null)
                                {
                                    continue;
                                }

                                XmlElement corridor = xmlDoc.CreateElement("Corridor");
                                corridors.AppendChild(corridor);
                                corridor.SetAttribute("Name", corr.Name);

                                XmlElement baselines = xmlDoc.CreateElement("Baselines");
                                corridor.AppendChild(baselines);

                                int blCounter = 0;

                                foreach (Autodesk.Civil.DatabaseServices.Baseline b in corr.Baselines)
                                {
                                    XmlElement baseline = xmlDoc.CreateElement("Baseline");
                                    baselines.AppendChild(baseline);
                                    baseline.SetAttribute("Name", b.Name);
                                    baseline.SetAttribute("Index", blCounter.ToString());

                                    XmlElement featurelines = xmlDoc.CreateElement("FeatureLines");
                                    baseline.AppendChild(featurelines);

                                    foreach (string cn in b.MainBaselineFeatureLines.FeatureLineCollectionMap.CodeNames())
                                    {
                                        try
                                        {
                                            foreach (Autodesk.Civil.DatabaseServices.CorridorFeatureLine cfl in b.MainBaselineFeatureLines.FeatureLineCollectionMap[cn])
                                            {
                                                try
                                                {
                                                    XmlElement featureline = xmlDoc.CreateElement("FeatureLine");
                                                    featureline.SetAttribute("Code", cn);
                                                    if (cfl.StyleName != null)
                                                    {
                                                        featureline.SetAttribute("Style", cfl.StyleName);
                                                    }
                                                    else
                                                    {
                                                        featureline.SetAttribute("Style", "");
                                                    }

                                                    featurelines.AppendChild(featureline);

                                                    XmlElement points = xmlDoc.CreateElement("Points");
                                                    featureline.AppendChild(points);

                                                    double totalStation = 0;

                                                    foreach (Autodesk.Civil.DatabaseServices.FeatureLinePoint cflp in cfl.FeatureLinePoints)
                                                    {
                                                        double offset = cflp.Offset;
                                                        double station = cflp.Station;
                                                        Autodesk.AutoCAD.Geometry.Point3d p3d = cflp.XYZ;

                                                        XmlElement point = xmlDoc.CreateElement("Point");
                                                        point.SetAttribute("X", p3d.X.ToString());
                                                        point.SetAttribute("Y", p3d.Y.ToString());
                                                        point.SetAttribute("Z", p3d.Z.ToString());
                                                        point.SetAttribute("Station", station.ToString());
                                                        point.SetAttribute("Offset", offset.ToString());
                                                        point.SetAttribute("IsBreak", cflp.IsBreak ? "1" : "0");

                                                        var reg = b.BaselineRegions.Cast<Autodesk.Civil.DatabaseServices.BaselineRegion>()
                                                       .First(x => x.StartStation < station && x.EndStation > station
                                                           || Math.Abs(x.StartStation - station) < 0.0001
                                                           || Math.Abs(x.EndStation - station) < 0.0001);

                                                        point.SetAttribute("RegionIndex", b.BaselineRegions.IndexOf(reg).ToString());
                                                        points.AppendChild(point);
                                                        totalStation += station;
                                                    }

                                                    //double s = Convert.ToDouble(points.ChildNodes[points.ChildNodes.Count / 2].Attributes["Station"].Value);
                                                    // double s = totalStation / points.ChildNodes.Count;
                                                    double o = Convert.ToDouble(points.FirstChild.Attributes["Offset"].Value);
                                                    featureline.SetAttribute("Side", o < 0 ? "-1" : "1");

                                                    //Autodesk.Civil.DatabaseServices.BaselineRegion reg = null; //  b.BaselineRegions.Cast<Autodesk.Civil.DatabaseServices.BaselineRegion>().First(x => x.StartStation < s && x.EndStation > s);

                                                    //foreach (var br in b.BaselineRegions)
                                                    //{
                                                    //    if (br.StartStation < s && br.EndStation > s)
                                                    //    {
                                                    //        reg = br;
                                                    //        break;
                                                    //    }
                                                    //}

                                                    //if (reg != null)
                                                    //{
                                                    //    if (b.BaselineRegions.Contains(reg))
                                                    //    {
                                                    //        try
                                                    //        {
                                                    //            //featureline.SetAttribute("RegionIndex", b.BaselineRegions.IndexOf(reg).ToString());
                                                    //            featureline.SetAttribute("Side", o < 0 ? "-1" : "1");
                                                    //        }
                                                    //        catch (System.Exception)
                                                    //        {
                                                    //            //featureline.SetAttribute("RegionIndex", "-1");  // this is wrong on purpose
                                                    //            featureline.SetAttribute("Side", "1");
                                                    //        }
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        //featureline.SetAttribute("RegionIndex", "-1");  // this is wrong on purpose
                                                    //        featureline.SetAttribute("Side", "1");
                                                    //    }
                                                    //}
                                                    //else 
                                                    //{
                                                    //    //featureline.SetAttribute("RegionIndex", "-1");  // this is wrong on purpose
                                                    //    featureline.SetAttribute("Side", "1");
                                                    //}

                                                    //if (!featureline.HasAttribute("RegionIndex"))
                                                    //{
                                                    //    featureline.SetAttribute("RegionIndex", "-1");  // this is wrong on purpose
                                                    //}
                                                    if (!featureline.HasAttribute("Side"))
                                                    {
                                                        featureline.SetAttribute("Side", "1");  // this is wrong on purpose
                                                    }
                                                }
                                                catch (System.Exception ex)
                                                {
                                                    System.Windows.Forms.MessageBox.Show(string.Format("ERROR 1: {0}", ex.Message));
                                                }
                                            }
                                        }
                                        catch (System.Exception ex)
                                        {
                                            // System.Windows.Forms.MessageBox.Show(string.Format("ERROR 2: {0}", ex.Message));
                                        }
                                    }

                                    ++blCounter;
                                }
                            }
                        }
                        else
                        {
                            Autodesk.Civil.DatabaseServices.Corridor corr = null;

                            try
                            {
                                corr = t.GetObject(oid1, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Corridor;
                            }
                            catch (System.Exception ex)
                            {
                                System.Windows.Forms.MessageBox.Show(string.Format("ERROR: {0}", ex.Message));
                            }

                            if (corr == null)
                            {
                                System.Windows.Forms.MessageBox.Show(string.Format("ERROR: {0}", "Cannot find the specified corridor."));
                                return;
                            }

                            XmlElement corridor = xmlDoc.CreateElement("Corridor");
                            corridors.AppendChild(corridor);
                            corridor.SetAttribute("Name", corr.Name);

                            path = path.Replace(".xml", string.Format("_{0}.xml", corr.Name));

                            XmlElement baselines = xmlDoc.CreateElement("Baselines");
                            corridor.AppendChild(baselines);

                            int blCounter = 0;

                            foreach (Autodesk.Civil.DatabaseServices.Baseline b in corr.Baselines)
                            {
                                XmlElement baseline = xmlDoc.CreateElement("Baseline");
                                baselines.AppendChild(baseline);
                                baseline.SetAttribute("Name", b.Name);
                                baseline.SetAttribute("Index", blCounter.ToString());

                                XmlElement featurelines = xmlDoc.CreateElement("FeatureLines");
                                baseline.AppendChild(featurelines);

                                foreach (string cn in b.MainBaselineFeatureLines.FeatureLineCollectionMap.CodeNames())
                                {
                                    try
                                    {
                                        foreach (Autodesk.Civil.DatabaseServices.CorridorFeatureLine cfl in b.MainBaselineFeatureLines.FeatureLineCollectionMap[cn])
                                        {
                                            try
                                            {
                                                XmlElement featureline = xmlDoc.CreateElement("FeatureLine");
                                                featureline.SetAttribute("Code", cn);
                                                if (cfl.StyleName != null)
                                                {
                                                    featureline.SetAttribute("Style", cfl.StyleName);
                                                }
                                                else
                                                {
                                                    featureline.SetAttribute("Style", "");
                                                }
                                                featurelines.AppendChild(featureline);

                                                XmlElement points = xmlDoc.CreateElement("Points");
                                                featureline.AppendChild(points);

                                                foreach (Autodesk.Civil.DatabaseServices.FeatureLinePoint cflp in cfl.FeatureLinePoints)
                                                {
                                                    double offset = cflp.Offset;
                                                    double station = cflp.Station;
                                                    Autodesk.AutoCAD.Geometry.Point3d p3d = cflp.XYZ;

                                                    XmlElement point = xmlDoc.CreateElement("Point");
                                                    point.SetAttribute("X", p3d.X.ToString());
                                                    point.SetAttribute("Y", p3d.Y.ToString());
                                                    point.SetAttribute("Z", p3d.Z.ToString());
                                                    point.SetAttribute("Station", station.ToString());
                                                    point.SetAttribute("Offset", offset.ToString());
                                                    point.SetAttribute("IsBreak", cflp.IsBreak ? "1" : "0");

                                                    var reg = b.BaselineRegions.Cast<Autodesk.Civil.DatabaseServices.BaselineRegion>()
                                                        .First(x => x.StartStation < station && x.EndStation > station 
                                                            || Math.Abs(x.StartStation - station) < 0.0001 
                                                            || Math.Abs(x.EndStation - station) < 0.0001);

                                                    point.SetAttribute("RegionIndex", b.BaselineRegions.IndexOf(reg).ToString());

                                                    points.AppendChild(point);
                                                }

                                                //double s = Convert.ToDouble(points.ChildNodes[points.ChildNodes.Count / 2].Attributes["Station"].Value);
                                                double o = Convert.ToDouble(points.FirstChild.Attributes["Offset"].Value);

                                                //var reg = b.BaselineRegions.Cast<Autodesk.Civil.DatabaseServices.BaselineRegion>().First(x => x.StartStation < s && x.EndStation > s);

                                                //featureline.SetAttribute("RegionIndex", b.BaselineRegions.IndexOf(reg).ToString());
                                                featureline.SetAttribute("Side", o < 0 ? "-1" : "1");
                                            }
                                            catch (System.Exception ex)
                                            {
                                                //System.Windows.Forms.MessageBox.Show(string.Format("ERROR 3: {0}", ex.Message));
                                            }
                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                        //System.Windows.Forms.MessageBox.Show(string.Format("ERROR 4: {0}", ex.Message));
                                    }
                                }

                                ++blCounter;
                            }
                        }
                    }
                }
            }

            xmlDoc.Save(path);
        }

        [CommandMethod("-ExportSubassemblyShapesToXML")]
        public void ExportSubassemblyShapesToXml()
        {
            string path = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "CorridorShapes.xml");

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            Document doc = Application.DocumentManager.MdiActiveDocument;
            CivilDocument cdoc = CivilApplication.ActiveDocument;

            short fd = (short)Application.GetSystemVariable("FILEDIA");

            string corridorHandle = "";
            int bi = -1;
            int ri = -1;

            try
            {
                PromptStringOptions pso = new PromptStringOptions("\nEnter Corridor Handle");
                PromptResult resCorr = doc.Editor.GetString(pso);
                corridorHandle = resCorr.StringResult;

                PromptIntegerOptions pio = new PromptIntegerOptions("\nEnter Baseline Index");
                PromptIntegerResult resBI = doc.Editor.GetInteger(pio);
                bi = resBI.Value;

                PromptIntegerOptions pior = new PromptIntegerOptions("\nEnter BaselineRegion Index");
                PromptIntegerResult resRg = doc.Editor.GetInteger(pior);
                ri = resRg.Value;
            }
            catch { }

            Application.SetSystemVariable("FILEDIA", fd);

            XmlDocument xmlDoc = new XmlDocument();

            XmlElement docElement = xmlDoc.CreateElement("Document");
            xmlDoc.AppendChild(docElement);

            XmlElement corridors = xmlDoc.CreateElement("Corridors");
            docElement.AppendChild(corridors);

            docElement.SetAttribute("Name", doc.Name);

            using (doc.LockDocument())
            {
                using (Database db = doc.Database)
                {
                    using (Transaction t = db.TransactionManager.StartTransaction())
                    {
                        foreach (ObjectId oid in cdoc.CorridorCollection)
                        {
                            bool toRebuild = false;

                            Autodesk.Civil.DatabaseServices.Corridor corr = t.GetObject(oid, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Corridor;

                            if (!string.IsNullOrWhiteSpace(corridorHandle) && !string.IsNullOrEmpty(corridorHandle))
                            {
                                if (corr.Handle.ToString() != corridorHandle)
                                {
                                    continue;
                                }
                            }

                            if (false) // DEBUG set to false to avoid rebuilding corridors without Subassemly PKTs
                            {

                                foreach (Autodesk.Civil.DatabaseServices.Baseline b in corr.Baselines)
                                {
                                    foreach (Autodesk.Civil.DatabaseServices.BaselineRegion r in b.BaselineRegions)
                                    {
                                        int rIndex = b.BaselineRegions.IndexOf(r);

                                        if (rIndex > 0 && r.AssemblyId != b.BaselineRegions[rIndex - 1].AssemblyId && r.StartStation - b.BaselineRegions[rIndex - 1].EndStation < 0.001)
                                        {
                                            if (!toRebuild)
                                            {
                                                toRebuild = true;
                                            }

                                            if (r.SortedStations()[1] - r.StartStation > 0.001)
                                            {
                                                r.AddStation(r.StartStation + 0.001, "Extra Station");  // Need to rebuild the corridor !!!
                                            }
                                        }
                                    }
                                }

                                if (toRebuild)
                                {
                                    corr.UpgradeOpen();
                                    corr.Rebuild();
                                    corr.DowngradeOpen();
                                }
                            }

                            XmlElement corridor = xmlDoc.CreateElement("Corridor");
                            corridors.AppendChild(corridor);
                            corridor.SetAttribute("Name", corr.Name);

                            XmlElement baselines = xmlDoc.CreateElement("Baselines");
                            corridor.AppendChild(baselines);

                            int blCounter = 0;

                            foreach (Autodesk.Civil.DatabaseServices.Baseline b in corr.Baselines)
                            {
                                if (bi != -1)
                                {
                                    if (blCounter != bi)
                                    {
                                        ++blCounter;
                                        continue;
                                    }
                                }

                                XmlElement baseline = xmlDoc.CreateElement("Baseline");
                                baselines.AppendChild(baseline);
                                baseline.SetAttribute("Name", b.Name);
                                baseline.SetAttribute("Index", blCounter.ToString());

                                XmlElement regions = xmlDoc.CreateElement("Regions");
                                baseline.AppendChild(regions);

                                int rCounter = 0;

                                foreach (Autodesk.Civil.DatabaseServices.BaselineRegion r in b.BaselineRegions)
                                {
                                    int rIndex = b.BaselineRegions.IndexOf(r);

                                    if (ri != -1)
                                    {
                                        if (ri != rCounter)
                                        {
                                            ++rCounter;
                                            continue;
                                        }
                                    }

                                    XmlElement region = xmlDoc.CreateElement("Region");
                                    regions.AppendChild(region);
                                    region.SetAttribute("Name", r.Name);

                                    region.SetAttribute("Index", rIndex.ToString());

                                    // WARNING: Baselines can have multiple regions and differnet assemblies associated.
                                    // In this case only the first assembly has the starting station in the first region.
                                    // The other regions will not have the starting station if it is equal to
                                    // the last station of the previous region.

                                    // SOLUTION: Use multiple regions on the same baseline IF AND ONLY IF the regions have gaps in between.
                                    // If the regions have to be contiguous and the assemblies are differnet it is better
                                    // to model a baseline with a single region for each assembly instead.

                                    foreach (Autodesk.Civil.DatabaseServices.AppliedAssembly aa in r.AppliedAssemblies)
                                    {
                                        Autodesk.Civil.DatabaseServices.Assembly assembly = null;

                                        if (r.AssemblyId != ObjectId.Null)
                                        {
                                            try
                                            {
                                                assembly = t.GetObject(r.AssemblyId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Assembly;
                                            }
                                            catch { }
                                        }

                                        string assemblyName = "";

                                        if (assembly != null)
                                        {
                                            assemblyName = assembly.Name;
                                        }

                                        foreach (Autodesk.Civil.DatabaseServices.AppliedSubassembly asa in aa.GetAppliedSubassemblies())
                                        {
                                            Autodesk.Civil.DatabaseServices.Subassembly subassembly = null;

                                            if (asa.SubassemblyId != ObjectId.Null)
                                            {
                                                try
                                                {
                                                    subassembly = t.GetObject(asa.SubassemblyId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Subassembly;
                                                }
                                                catch { }
                                            }

                                            string subassemblyName = "";
                                            string handle = "";
                                            string station = asa.OriginStationOffsetElevationToBaseline.X.ToString();

                                            if (subassembly != null)
                                            {
                                                subassemblyName = subassembly.Name;
                                                handle = asa.SubassemblyId.Handle.ToString();
                                            }

                                            int shapeCounter = 0;

                                            foreach (Autodesk.Civil.DatabaseServices.CalculatedShape cs in asa.Shapes)
                                            {
                                                XmlElement shape = xmlDoc.CreateElement("Shape");
                                                region.AppendChild(shape);
                                                shape.SetAttribute("Corridor", corr.Name);
                                                shape.SetAttribute("BaselineIndex", blCounter.ToString());
                                                shape.SetAttribute("RegionIndex", rCounter.ToString());
                                                shape.SetAttribute("AssemblyName", assemblyName);
                                                shape.SetAttribute("SubassemblyName", subassemblyName);
                                                shape.SetAttribute("Handle", handle);
                                                shape.SetAttribute("ShapeIndex", shapeCounter.ToString());
                                                shape.SetAttribute("Station", station);

                                                XmlElement codes = xmlDoc.CreateElement("Codes");
                                                shape.AppendChild(codes);
                                                foreach (string cd in cs.CorridorCodes)
                                                {
                                                    XmlElement code = xmlDoc.CreateElement("Code");
                                                    codes.AppendChild(code);
                                                    code.SetAttribute("Name", cd);
                                                }

                                                IList<Autodesk.AutoCAD.Geometry.Point3d> points = new List<Autodesk.AutoCAD.Geometry.Point3d>();

                                                foreach (Autodesk.Civil.DatabaseServices.CalculatedLink cl in cs.CalculatedLinks)
                                                {
                                                    foreach (Autodesk.Civil.DatabaseServices.CalculatedPoint cp in cl.CalculatedPoints)
                                                    {
                                                        Autodesk.AutoCAD.Geometry.Point3d soe = cp.StationOffsetElevationToBaseline;
                                                        Autodesk.AutoCAD.Geometry.Point3d p3d = b.StationOffsetElevationToXYZ(soe);

                                                        if (!points.Contains(p3d))
                                                        {
                                                            points.Add(p3d);
                                                        }
                                                    }
                                                }

                                                foreach (Autodesk.AutoCAD.Geometry.Point3d p3d in points)
                                                {
                                                    XmlElement point = xmlDoc.CreateElement("Point");
                                                    shape.AppendChild(point);
                                                    point.SetAttribute("X", p3d.X.ToString());
                                                    point.SetAttribute("Y", p3d.Y.ToString());
                                                    point.SetAttribute("Z", p3d.Z.ToString());
                                                }

                                                ++shapeCounter;
                                            }
                                        }
                                    }

                                    ++rCounter;
                                }

                                ++blCounter;
                            }
                        }

                        t.Commit();
                    }
                }
            }

            xmlDoc.Save(path);
        }

        [CommandMethod("-ExportSubassemblyLinksToXML")]
        public void ExportSubassemblyLinksToXml()
        {
            string path = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "CorridorLinks.xml");

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            Document doc = Application.DocumentManager.MdiActiveDocument;
            CivilDocument cdoc = CivilApplication.ActiveDocument;

            short fd = (short)Application.GetSystemVariable("FILEDIA");

            string corridorHandle = "";
            int bi = -1;
            int ri = -1;

            try
            {
                PromptStringOptions pso = new PromptStringOptions("\nEnter Corridor Handle");
                PromptResult resCorr = doc.Editor.GetString(pso);
                corridorHandle = resCorr.StringResult;

                PromptIntegerOptions pio = new PromptIntegerOptions("\nEnter Baseline Index");
                PromptIntegerResult resBI = doc.Editor.GetInteger(pio);
                bi = resBI.Value;

                PromptIntegerOptions pior = new PromptIntegerOptions("\nEnter BaselineRegion Index");
                PromptIntegerResult resRg = doc.Editor.GetInteger(pior);
                ri = resRg.Value;
            }
            catch { }

            Application.SetSystemVariable("FILEDIA", fd);

            XmlDocument xmlDoc = new XmlDocument();

            XmlElement docElement = xmlDoc.CreateElement("Document");
            xmlDoc.AppendChild(docElement);

            XmlElement corridors = xmlDoc.CreateElement("Corridors");
            docElement.AppendChild(corridors);

            docElement.SetAttribute("Name", doc.Name);

            using (doc.LockDocument())
            {
                using (Database db = doc.Database)
                {
                    using (Transaction t = db.TransactionManager.StartTransaction())
                    {
                        foreach (ObjectId oid in cdoc.CorridorCollection)
                        {
                            Autodesk.Civil.DatabaseServices.Corridor corr = t.GetObject(oid, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Corridor;

                            if (!string.IsNullOrWhiteSpace(corridorHandle) && !string.IsNullOrEmpty(corridorHandle))
                            {
                                if (corr.Handle.ToString() != corridorHandle)
                                {
                                    continue;
                                }
                            }

                            bool toRebuild = false;

                            foreach (Autodesk.Civil.DatabaseServices.Baseline b in corr.Baselines)
                            {
                                foreach (Autodesk.Civil.DatabaseServices.BaselineRegion r in b.BaselineRegions)
                                {
                                    int rIndex = b.BaselineRegions.IndexOf(r);

                                    if (rIndex > 0 && r.AssemblyId != b.BaselineRegions[rIndex - 1].AssemblyId && r.StartStation - b.BaselineRegions[rIndex - 1].EndStation < 0.001)
                                    {
                                        if (!toRebuild)
                                        {
                                            toRebuild = true;
                                        }

                                        if (r.SortedStations()[1] - r.StartStation > 0.001)
                                        {
                                            r.AddStation(r.StartStation + 0.001, "Extra Station");  // Need to rebuild the corridor !!!
                                        }
                                    }
                                }
                            }

                            if (toRebuild)
                            {
                                corr.UpgradeOpen();
                                corr.Rebuild();
                                corr.DowngradeOpen();
                            }

                            XmlElement corridor = xmlDoc.CreateElement("Corridor");
                            corridors.AppendChild(corridor);
                            corridor.SetAttribute("Name", corr.Name);

                            XmlElement baselines = xmlDoc.CreateElement("Baselines");
                            corridor.AppendChild(baselines);

                            int blCounter = 0;

                            foreach (Autodesk.Civil.DatabaseServices.Baseline b in corr.Baselines)
                            {
                                if (bi != -1)
                                {
                                    if (blCounter != bi)
                                    {
                                        ++blCounter;
                                        continue;
                                    }
                                }

                                XmlElement baseline = xmlDoc.CreateElement("Baseline");
                                baselines.AppendChild(baseline);
                                baseline.SetAttribute("Name", b.Name);
                                baseline.SetAttribute("Index", blCounter.ToString());

                                XmlElement regions = xmlDoc.CreateElement("Regions");
                                baseline.AppendChild(regions);

                                int rCounter = 0;

                                foreach (Autodesk.Civil.DatabaseServices.BaselineRegion r in b.BaselineRegions)
                                {
                                    int rIndex = b.BaselineRegions.IndexOf(r);

                                    if (ri != -1)
                                    {
                                        if (rCounter != ri)
                                        {
                                            ++rCounter;
                                            continue;
                                        }
                                    }

                                    XmlElement region = xmlDoc.CreateElement("Region");
                                    regions.AppendChild(region);
                                    region.SetAttribute("Name", r.Name);
                                    region.SetAttribute("Index", rIndex.ToString());

                                    // WARNING: Baselines can have multiple regions and differnet assemblies associated.
                                    // In this case only the first assembly has the starting station in the first region.
                                    // The other regions will not have the starting station if it is equal to
                                    // the last station of the previous region.

                                    // SOLUTION: Use multiple regions on the same baseline IF AND ONLY IF the regions have gaps in between.
                                    // If the regions have to be contiguous and the assemblies are differnet it is better
                                    // to model a baseline with a single region for each assembly instead.

                                    foreach (Autodesk.Civil.DatabaseServices.AppliedAssembly aa in r.AppliedAssemblies)
                                    {
                                        Autodesk.Civil.DatabaseServices.Assembly assembly = null;

                                        if (r.AssemblyId != ObjectId.Null)
                                        {
                                            try
                                            {
                                                assembly = t.GetObject(r.AssemblyId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Assembly;
                                            }
                                            catch { }
                                        }

                                        string assemblyName = "";

                                        if (assembly != null)
                                        {
                                            assemblyName = assembly.Name;
                                        }

                                        foreach (Autodesk.Civil.DatabaseServices.AppliedSubassembly asa in aa.GetAppliedSubassemblies())
                                        {
                                            Autodesk.Civil.DatabaseServices.Subassembly subassembly = null;

                                            if (asa.SubassemblyId != ObjectId.Null)
                                            {
                                                try
                                                {
                                                    subassembly = t.GetObject(asa.SubassemblyId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Subassembly;
                                                }
                                                catch { }
                                            }

                                            string subassemblyName = "";
                                            string handle = "";
                                            string station = asa.OriginStationOffsetElevationToBaseline.X.ToString();

                                            if (subassembly != null)
                                            {
                                                subassemblyName = subassembly.Name;
                                                handle = asa.SubassemblyId.Handle.ToString();
                                            }

                                            int linkCounter = 0;

                                            foreach (Autodesk.Civil.DatabaseServices.CalculatedLink cl in asa.Links)
                                            {
                                                XmlElement shape = xmlDoc.CreateElement("Link");
                                                region.AppendChild(shape);
                                                shape.SetAttribute("Corridor", corr.Name);
                                                shape.SetAttribute("BaselineIndex", blCounter.ToString());
                                                shape.SetAttribute("RegionIndex", rCounter.ToString());
                                                shape.SetAttribute("AssemblyName", assemblyName);
                                                shape.SetAttribute("SubassemblyName", subassemblyName);
                                                shape.SetAttribute("Handle", handle);
                                                shape.SetAttribute("LinkIndex", linkCounter.ToString());
                                                shape.SetAttribute("Station", station);

                                                XmlElement codes = xmlDoc.CreateElement("Codes");
                                                shape.AppendChild(codes);
                                                foreach (string cd in cl.CorridorCodes)
                                                {
                                                    XmlElement code = xmlDoc.CreateElement("Code");
                                                    codes.AppendChild(code);
                                                    code.SetAttribute("Name", cd);
                                                }

                                                IList<Autodesk.AutoCAD.Geometry.Point3d> points = new List<Autodesk.AutoCAD.Geometry.Point3d>();

                                                foreach (Autodesk.Civil.DatabaseServices.CalculatedPoint cp in cl.CalculatedPoints)
                                                {
                                                    Autodesk.AutoCAD.Geometry.Point3d soe = cp.StationOffsetElevationToBaseline;
                                                    Autodesk.AutoCAD.Geometry.Point3d p3d = b.StationOffsetElevationToXYZ(soe);

                                                    if (!points.Contains(p3d))
                                                    {
                                                        points.Add(p3d);
                                                    }
                                                }

                                                foreach (Autodesk.AutoCAD.Geometry.Point3d p3d in points)
                                                {
                                                    XmlElement point = xmlDoc.CreateElement("Point");
                                                    shape.AppendChild(point);
                                                    point.SetAttribute("X", p3d.X.ToString());
                                                    point.SetAttribute("Y", p3d.Y.ToString());
                                                    point.SetAttribute("Z", p3d.Z.ToString());
                                                }

                                                ++linkCounter;
                                            }
                                        }
                                    }

                                    ++rCounter;
                                }

                                ++blCounter;
                            }
                        }
                    }
                }
            }

            xmlDoc.Save(path);
        }

        [CommandMethod("-ExportSurfaceToXML")]
        public void ExportSurfaceToXml()
        {
            string path = Path.Combine(Environment.GetEnvironmentVariable("TMP", EnvironmentVariableTarget.User), "Surface.xml");

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            Document doc = Application.DocumentManager.MdiActiveDocument;
            CivilDocument cdoc = CivilApplication.ActiveDocument;

            short fd = (short)Application.GetSystemVariable("FILEDIA");

            string surfHandle = "";

            try
            {
                PromptStringOptions pso = new PromptStringOptions("\nEnter Surface Handle");
                PromptResult resSurf = doc.Editor.GetString(pso);
                surfHandle = resSurf.StringResult;
            }
            catch { }

            Application.SetSystemVariable("FILEDIA", fd);

            XmlDocument xmlDoc = new XmlDocument();

            XmlElement docElement = xmlDoc.CreateElement("Document");
            xmlDoc.AppendChild(docElement);

            XmlElement surfaces = xmlDoc.CreateElement("Surfaces");
            docElement.AppendChild(surfaces);

            docElement.SetAttribute("Name", doc.Name);

            using (doc.LockDocument())
            {
                using (Database db = doc.Database)
                {
                    using (Transaction t = db.TransactionManager.StartTransaction())
                    {
                        foreach (ObjectId oid in cdoc.GetSurfaceIds())
                        {
                            Autodesk.Civil.DatabaseServices.TinSurface surf = t.GetObject(oid, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.TinSurface;

                            if (!string.IsNullOrWhiteSpace(surfHandle) && !string.IsNullOrEmpty(surfHandle))
                            {
                                if (surf.Handle.ToString() != surfHandle)
                                {
                                    continue;
                                }
                            }

                            var visibleTriangles = surf.GetTriangles(false);

                            XmlElement surface = xmlDoc.CreateElement("Surface");
                            surfaces.AppendChild(surface);
                            surface.SetAttribute("Name", surf.Name);

                            XmlElement vertices = xmlDoc.CreateElement("Vertices");
                            surface.AppendChild(vertices);

                            Dictionary<Autodesk.Civil.DatabaseServices.TinSurfaceTriangle, List<string>> triangleDict = new Dictionary<Autodesk.Civil.DatabaseServices.TinSurfaceTriangle, List<string>>();

                            for (int vCounter = 0; vCounter < surf.Vertices.Count; ++vCounter)
                            {
                                Autodesk.Civil.DatabaseServices.TinSurfaceVertex v = surf.Vertices.ElementAt(vCounter);

                                Autodesk.AutoCAD.Geometry.Point3d p3d = v.Location;

                                foreach (var tri in v.Triangles)
                                {
                                    if (!visibleTriangles.Contains(tri))
                                    {
                                        continue;
                                    }

                                    if (triangleDict.Keys.Contains(tri))
                                    {
                                        if (!triangleDict[tri].Contains(vCounter.ToString()) && triangleDict[tri].Count < 3)
                                        {
                                            triangleDict[tri].Add(vCounter.ToString());
                                        }
                                    }
                                    else
                                    {
                                        triangleDict.Add(tri, new List<string>() { vCounter.ToString() });
                                    }
                                }

                                XmlElement vertex = xmlDoc.CreateElement("Vertex");
                                vertices.AppendChild(vertex);
                                vertex.SetAttribute("X", p3d.X.ToString());
                                vertex.SetAttribute("Y", p3d.Y.ToString());
                                vertex.SetAttribute("Z", p3d.Z.ToString());
                                vertex.SetAttribute("id", vCounter.ToString());
                            }

                            XmlElement triangles = xmlDoc.CreateElement("Triangles");
                            surface.AppendChild(triangles);

                            foreach (var tri in triangleDict.Keys)
                            {
                                XmlElement tria = xmlDoc.CreateElement("Triangle");
                                triangles.AppendChild(tria);

                                for (int i = 0; i < 3; ++i)
                                {
                                    tria.SetAttribute(string.Format("V{0}", i), triangleDict[tri][i]);
                                }
                            }
                        }
                    }
                }
            }

            xmlDoc.Save(path);
        }
    }
}
