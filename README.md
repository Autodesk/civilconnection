# CivilConnection and CivilPython
CivilConnection is a Dynamo for Revit package for Autodesk® Revit and Autodesk® Civil 3D.

CivilPython is a command for Autodesk® Civil 3D that enables the execution of Python scripts accessing AutoCAD and Civil 3D .NET APIs.

## CivilConnection Main Features
Here are some features enabled by CivilConnection:
* Enables the exchange of information between Civil 3D, Dynamo and Revit.
* Reads the information embedded in rich linear objects such as Civil 3D alignments, corridors or feature lines and creates proxy elements in Dynamo. In turn the proxy elements can be used to establish dynamic relationships to drive the creation of discrete Revit elements (i.e. single point family instances, line based objects such as structural framing or MEP segments, complex objects such as adaptive components, floors, walls, Revit link instances).
* Provides features to update the location, orientation and metadata of Revit elements against a Civil 3D input.
* Reads the shapes and links of the Civil 3D Corridors and create and update modifiable Revit families without any tessellation. This enables to further the detailing in Revit, use parts and rebar, assign custom materials to the objects and preparing it for the construction phase.
* Creates basic AutoCAD entities such as layers, points, line, arcs, polylines 2d and 3d, region, solids.
* Creates basic Civil 3D entities such as point groups or alignments to enable the creation of TIN surfaces or providing targets for corridors.
* Performs boolean operations between AutoCAD solids; this functionalities are used to add details to the Civil 3D models with discrete elements or performing subtraction that preserve the individual solids involved.
* Imports geometry elements generated in Dynamo into Civil 3D via SAT Export / Import.
* Imports and updates the solids in the geometry of Revit elements into Civil 3D via the link Element functionality.
* Gets or sets the parameters of the subassemblies in a corridor and force the corridor to rebuild.
* Sends commands to the Civil 3D command line (this feature can be used to launch CivilPython).

## CivilPython Main Features
Here are some features enabled by CivilPython:
* Uses the IronPython 2.7 that comes with Dynamo for Revit (in alternative you can download IronPython separately).
* Uses .NET API with the same interpreted approach without the need to compile the code into a .DLL assembly.
* The .py code can be developed in any external Python IDE.
* The code can be copied and pasted into a Python Script node in Dynamo for Civil 3D (2020 onward)
* Provides a user interface command and a command line version that can be used to select Python scripts interactively or to specify their path on the file system via CivilConnection.
* CivilPython runs for Civil 3D 2016 onward.

## How to use
See ./Doc/Linear Structures Workflow Guide.pdf

In ./Compiled there are the releases of the CivilConnection Dynamo package as well as CivilPython ready to use.

## Release
The release numbers for the package correspond to the major release numbers of Civil 3D and Revit it will run against

**Note that CivilConnection and Civil 3D with different release numbers are not compatible.**

* Civil 3D and Revit 2017 - CivilConnection 2017 - Autodesk2017.dll [End of Life]
* Civil 3D and Revit 2018 - CivilConnection 2018 - Autodesk2018.dll
* Civil 3D and Revit 2019 - CivilConnection 2019 - Autodesk2019.dll
* Civil 3D and Revit 2020 - CivilConnection 2020 - Autodesk2020.dll
* Civil 3D and Revit 2021 - CivilConnection 2021 - Autodesk2021.dll

## License
See [LICENSING.md](LICENSING.md)

## Contributions
Read our contribution guidelines here: [CONTRIBUTING.md](CONTRIBUTING.md)
