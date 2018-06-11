# Simple-Building-Geometry-Generator-for-EnergyPlus

Coding Log for epc+ver-0-0-3_params

Main Updates:
- Added perimeter 2 as a option to represent "corner" zones
- Added ASHRAE 90.1 2010 construction templates to allow generating geometry in together with constructions. 
- Added general idf objects to match Open Studio exported geometry files. 
- Current generated idf objects are:
*Version
*SimulaitonControl
*Building 
*RunPeriod
*ScheduleTypeLimits
*Material
*Material:NoMass
*Material:AirGap
*WindowMaterial:SimpleGlazingSystem
*WindowMaterial:Glazing
*Construction
*GlobalGeometryRules
*Zone
*BuildingSurface:Detailed
*FenestrationSurface:Detailed
*Sizing:Parameters


- Building the interface to IDFexporter. Tested the work flow of working with IDFexporter in creating complete EnergyPlus models. Current workflow requires a extra step to open generated idf file in OpenStudio and then export, in order to sort the objects orders to work with IDFexporter. 


Next:
- Rewrite the entire code using geomeppy, write a py class if necesary and merge to geoeppy
- Add schedules, people, equipment, light, thermostats, infiltration, outdoor air, sizing zone, ideal load hvac objects to contruct a load calc E+ model.
 


01/18/2018 Updates:

- Build and tested interface with params. Tested workflow of working with params in creating complete EnergyPlus models. Current epc+ Nongeo Generator creates idf objects equivalent to the "general.imf", "geometry.pxt" and construction templates. 
- Templates needed from params are location(location and designday objects in EnergyPlus, this could be replaced by the ddy file usually shipped with weather file, therefore Optional), zoneloads, zonehvac and system objects. 
- EnergyPlus IDD file set to 8.5 in eppy to in order to work with params. Current version of params writes idf object following EnergyPlus 8.5 IDD file. One need to update the EnergyPlus idf from version 8.5 to 8.8 if needed.
- Changed the exported file name as "NongeoXport"+datetime stamp
- Added utility to change exported file extension to be ".pxt", in order to be directly import to params.

 

Added the following objects to the exports:
*ShadowCalculation
*SurfaceConvectionAlgorithm:Inside
*SurfaceConvectionAlgorithm:Outside
*HeatBalanceAlgorithm
*SurfaceProperty:OtherSideConditionsModel
*ConvergenceLimits

Updated the follwoing objects:
*ScheduleTypeLimits, added 7 type limits objects
 
