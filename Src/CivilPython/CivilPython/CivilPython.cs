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

                string ver = version.ToString().Split(new char[] {'.'})[0];

                string release = "2020";

                switch (ver)
                {
                    case "20":
                        release = "2016";
                        break;
                    case "21":
                        release = "2017";
                        break;
                    case "22":
                        release = "2018";
                        break;
                    case "23":
                        release = "2019";
                        break;
                    case "24":
                        release = "2020";
                        break;
                }

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
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\acmgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\acdbmgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\accoremgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\ACA\AecBaseMgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\ACA\AecPropDataMgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\C3D\AeccDbMgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\C3D\AeccPressurePipesMgd.dll", release)));
                    ipy.LoadAssembly(Assembly.LoadFrom(string.Format(@"C:\Program Files\Autodesk\AutoCAD {0}\acdbmgdbrep.dll", release)));


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

                        // As the user to select a .py file

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
            string path = Path.Combine(Path.GetTempPath(), "LandFeatureLinesReport.xml");

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            XmlDocument xmlDoc = new XmlDocument();

            XmlElement docElement = xmlDoc.CreateElement("Document");
            xmlDoc.AppendChild(docElement);
            docElement.SetAttribute("Name", docElement.Name);

            XmlElement featurelines = xmlDoc.CreateElement("FeatureLines");
            docElement.AppendChild(featurelines);

            Document doc = Application.DocumentManager.MdiActiveDocument;

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
                                featureline.SetAttribute("Style", f.StyleName);
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
    }
}
