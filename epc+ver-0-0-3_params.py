"""
# Copyright (c) 2018 Yiyuan Jia
# =======================================================================
#  Distributed under the MIT License.
#  (See accompanying file LICENSE or copy at
#  http://opensource.org/licenses/MIT)
# =======================================================================
"""
__authors__ = "Yiyuan Jia"
__copyright__ = "Copyright 2018, Yiyuan Jia, Affiliated Engineers, Inc."
__credits__ = [""]
__license__ = "MIT"
__version__ = "0.1"
__maintainer__ = "Yiyuan Jia"
__email__ = "yjia@aeieng.com"
__status__ = "BETA"


from xlpython import *
import numpy as np
from eppy import modeleditor
from eppy.modeleditor import IDF
from StringIO import StringIO
import datetime
import time
import os 
import shutil



iddfile = "c:/Anaconda2/Lib/site-packages/eppy/resources/iddfiles/Energy+V8_5_0.idd"
IDF.setiddname(iddfile)


@xlfunc
def LoadZoneName(myIDFfile,ID):
    fname=myIDFfile
    my_idf = IDF(fname)
    zones = my_idf.idfobjects['ZONE'] 
#    myZones = np.array([zone.Name for zone in zones])
    zoneName = zones[int(ID)-1].Name
    return zoneName


@xlfunc
def LoadZoneTag(myIDFfile,ID):
    fname=myIDFfile
    my_idf = IDF(fname)
    zones = my_idf.idfobjects['ZONE'] 
#    myZones = np.array([zone.Name for zone in zones])
    zoneTag = zones[int(ID)-1].Name[-9:-4]
    return zoneTag     

@xlfunc
def WriteIDF(myIDFfile, Bname):
    fname=myIDFfile
    my_idf = IDF(fname)
    building = my_idf.idfobjects['BUILDING']
    building[0].Name = Bname
    my_idf.save
    

#####This funtion is used to generate non-goemetric EnergyPlus geometry input files
#Argument of this funtions are
#Number of Zones
#Building Name 
    
@xlsub
@xlarg("sheet", vba="Sheet5")

def GeoGen(sheet):
    num_of_zones = sheet.Range("A7").Value
    construction_template = sheet.Range("AN7").Value
                                       
    #Initialize a blank idf file    
    ts = time.time()
#st = datetime.datetime.fromtimestamp(ts).strftime('%Y%m%d %H%M%S')
    st = datetime.datetime.fromtimestamp(ts).strftime('%Y%m%d-%H%M')                            
    blankstr = ""
    new_idf = IDF(StringIO(blankstr))
    new_idf.idfname = "NongeoXport "+st+".idf"
                              
    #Import Constructoin Templates                           
    const_fname="ahshrae901_construction_templates.idf"
    my_const = IDF(const_fname)
    
    #Copy All Construction Templates into New idf file
    materials = my_const.idfobjects['material'.upper()]
    for i in range(len(materials)):
        new_idf.copyidfobject(materials[i])
    
    materials_nomass = my_const.idfobjects['material:nomass'.upper()]
    for i in range(len(materials_nomass)):
        new_idf.copyidfobject(materials_nomass[i])
        
    materials_airgap = my_const.idfobjects['material:airgap'.upper()]
    for i in range(len(materials_airgap)):
        new_idf.copyidfobject(materials_airgap[i])
        
    window_materials_sgs = my_const.idfobjects['windowmaterial:simpleglazingsystem'.upper()]
    for i in range(len(window_materials_sgs)):
        new_idf.copyidfobject(window_materials_sgs[i])
    
    window_materials_glz = my_const.idfobjects['windowmaterial:glazing'.upper()]
    for i in range(len(window_materials_glz)):
        new_idf.copyidfobject(window_materials_glz[i])
        
    constructions = my_const.idfobjects['construction'.upper()]
    for i in range(len(constructions)):
        new_idf.copyidfobject(constructions[i])
    
    #Assign Surface Construction Names Based on User Templates Selections
    interior_floor_const = "Interior Floor"
    interior_wall_const = "Interior Wall"
    interior_ceiling_const = "Interior Ceiling"
    exterior_floor_const = "ExtSlabCarpet 4in ClimateZone 1-8"
    if construction_template =="ASHRAE 90.1 2010 Climate Zone 1":                
        exterior_wall_const = "ASHRAE 90.1-2010 ExtWall Mass ClimateZone 1"
        exterior_window_const = "ASHRAE 90.1-2010 ExtWindow Metal ClimateZone 1"
        exterior_roof_const = "ASHRAE 90.1-2010 ExtRoof IEAD ClimateZone 1"
        
    elif construction_template =="ASHRAE 90.1 2010 Climate Zone 2":                
        exterior_wall_const = "ASHRAE 90.1-2010 ExtWall Mass ClimateZone 2"
        exterior_window_const = "ASHRAE 90.1-2010 ExtWindow Metal ClimateZone 2"
        exterior_roof_const = "ASHRAE 90.1-2010 ExtRoof IEAD ClimateZone 2-8"
    
    elif construction_template =="ASHRAE 90.1 2010 Climate Zone 3":                
        exterior_wall_const = "ASHRAE 90.1-2010 ExtWall Mass ClimateZone 3"
        exterior_window_const = "ASHRAE 90.1-2010 ExtWindow Metal ClimateZone 3"
        exterior_roof_const = "ASHRAE 90.1-2010 ExtRoof IEAD ClimateZone 2-8"   

    elif construction_template =="ASHRAE 90.1 2010 Climate Zone 4":                
        exterior_wall_const = "ASHRAE 90.1-2010 ExtWall Mass ClimateZone 4"
        exterior_window_const = "ASHRAE 90.1-2010 ExtWindow Metal ClimateZone 4-6"
        exterior_roof_const = "ASHRAE 90.1-2010 ExtRoof IEAD ClimateZone 2-8"   

    elif construction_template =="ASHRAE 90.1 2010 Climate Zone 5":                
        exterior_wall_const = "ASHRAE 90.1-2010 ExtWall Mass ClimateZone 5"
        exterior_window_const = "ASHRAE 90.1-2010 ExtWindow Metal ClimateZone 4-6"
        exterior_roof_const = "ASHRAE 90.1-2010 ExtRoof IEAD ClimateZone 2-8"
        
    elif construction_template =="ASHRAE 90.1 2010 Climate Zone 6":                
        exterior_wall_const = "ASHRAE 90.1-2010 ExtWall Mass ClimateZone 6"
        exterior_window_const = "ASHRAE 90.1-2010 ExtWindow Metal ClimateZone 4-6"
        exterior_roof_const = "ASHRAE 90.1-2010 ExtRoof IEAD ClimateZone 2-8"        
        
    elif construction_template =="ASHRAE 90.1 2010 Climate Zone 7":                
        exterior_wall_const = "ASHRAE 90.1-2010 ExtWall Mass ClimateZone 7-8"
        exterior_window_const = "ASHRAE 90.1-2010 ExtWindow Metal ClimateZone 7-8"
        exterior_roof_const = "ASHRAE 90.1-2010 ExtRoof IEAD ClimateZone 2-8"   
        
    elif construction_template =="ASHRAE 90.1 2010 Climate Zone 8":                
        exterior_wall_const = "ASHRAE 90.1-2010 ExtWall Mass ClimateZone 7-8"
        exterior_window_const = "ASHRAE 90.1-2010 ExtWindow Metal ClimateZone 7-8"
        exterior_roof_const = "ASHRAE 90.1-2010 ExtRoof IEAD ClimateZone 2-8" 

        
    ####Start Writing idf Objects
    ##General idf Objects for IDFXporter
    
    #Define Version Objects
    version = new_idf.newidfobject("version".upper())
    version.Version_Identifier = 8.5
    
    #Define Simulation Control Objects
    simctrl = new_idf.newidfobject("simulationcontrol".upper())
    simctrl.Do_Zone_Sizing_Calculation = "Yes"
    simctrl.Do_System_Sizing_Calculation = "Yes"
    simctrl.Do_Plant_Sizing_Calculation = "Yes"
    simctrl.Run_Simulation_for_Sizing_Periods = "No"
    simctrl.Run_Simulation_for_Weather_File_Run_Periods = "Yes"
    
    #Define Building Object
    building = new_idf.newidfobject("building".upper())
    building.Name = "My Building"
    building.North_Axis = 0
    
    #Define Run Periods Object
    run_period = new_idf.newidfobject("runperiod".upper())
    run_period.Name = "Run Period 1"
    run_period.Begin_Month = 1
    run_period.Begin_Day_of_Month = 1
    run_period.End_Month = 12
    run_period.End_Day_of_Month = 31
    run_period.Day_of_Week_for_Start_Day = "Thursday"
    run_period.Use_Weather_File_Rain_Indicators = "Yes"
    run_period.Use_Weather_File_Snow_Indicators = "Yes"
    
    #Define Schedule Type Limits Objects
    sch_tp_limit_1 = new_idf.newidfobject("scheduletypelimits".upper())
    sch_tp_limit_1.Name = "Any Number"
    sch_tp_limit_2 = new_idf.newidfobject("scheduletypelimits".upper())
    sch_tp_limit_2.Name = "Fraction"
    sch_tp_limit_3 = new_idf.newidfobject("scheduletypelimits".upper())
    sch_tp_limit_3.Name = "Temperature"
    sch_tp_limit_4 = new_idf.newidfobject("scheduletypelimits".upper())
    sch_tp_limit_4.Name = "On/Off"
    sch_tp_limit_5 = new_idf.newidfobject("scheduletypelimits".upper())
    sch_tp_limit_5.Name = "Control Type"
    sch_tp_limit_6 = new_idf.newidfobject("scheduletypelimits".upper())
    sch_tp_limit_6.Name = "Humidity"
    sch_tp_limit_7 = new_idf.newidfobject("scheduletypelimits".upper())
    sch_tp_limit_7.Name = "Number"        
    
    ##Additional idf Objects from Params
    
    #Define Shadow Calculation Objects
    shadow_calc = new_idf.newidfobject("shadowcalculation".upper())
    shadow_calc.Calculation_Method = "AverageOverDaysInFrequency"
    
    #Define SurfaceConvectionAlgorithm:Inside Objects
    surf_conv_algorithm_in = new_idf.newidfobject("SurfaceConvectionAlgorithm:Inside".upper())
    surf_conv_algorithm_in.Algorithm = "TARP"
    
    #Define SurfaceConvectionAlgorithm:Outside Objects
    surf_conv_algorithm_out = new_idf.newidfobject("SurfaceConvectionAlgorithm:Outside".upper())
    surf_conv_algorithm_out.Algorithm = "DOE-2"    
    
    #Define HeatBalanceAlgorithm Objects
    hb_algorithm = new_idf.newidfobject("HeatBalanceAlgorithm".upper())
    hb_algorithm.Algorithm = "ConductionTransferFunction"
    
    #Define SurfaceProperty:OtherSideConditionsModel Objects
    surf_prop_oscm = new_idf.newidfobject("SurfaceProperty:OtherSideConditionsModel".upper())
    surf_prop_oscm.Name = "GapConvectionModel"
    
    #Define ConvergenceLimits Objects
    conv_limits = new_idf.newidfobject("ConvergenceLimits".upper())
    conv_limits.Minimum_System_Timestep = 0
    
    
    
    
    #Define Global Geometry Rules Objects
    rule = new_idf.newidfobject("globalgeometryrules".upper())
    rule.Starting_Vertex_Position = "UpperLeftCorner"
    rule.Vertex_Entry_Direction = "Counterclockwise"
    rule.Coordinate_System = "Relative"
    rule.Daylighting_Reference_Point_Coordinate_System = "Relative"
    rule.Rectangular_Surface_Coordinate_System = "Relative"
    
    #Define SizingParameters Objects
    szparams = new_idf.newidfobject("sizing:parameters".upper())
    szparams.Heating_Sizing_Factor = 1.25
    szparams.Cooling_Sizing_Factor = 1.15

   
    ####Non-Geometric Generator Starts
    
        
    for id in range(7,int(7+num_of_zones)):
        zone_name = sheet.Range("B"+str(id)).Value
#        zone_name = "Thermal Zone: " + zone_name
#        zone_name = zone_name
        zone_origin_x = sheet.Range("C"+str(id)).Value
        zone_origin_y = sheet.Range("D"+str(id)).Value
        zone_origin_z = sheet.Range("E"+str(id)).Value
#        prmtr_or_not =  sheet.Range("F"+str(id)).Value                      
        roof_or_not = sheet.Range("M"+str(id)).Value
        zone_height = sheet.Range("H"+str(id)).Value
        zone_length= sheet.Range("I"+str(id)).Value
        zone_width = sheet.Range("G"+str(id)).Value/sheet.Range("I"+str(id)).Value
        prmtr1_normal = sheet.Range("K"+str(id)).Value
        prmtr2_normal = sheet.Range("L"+str(id)).Value                       
        grndflr_or_not = sheet.Range("N"+str(id)).Value                      
        wind2wall_ratio = sheet.Range("O"+str(id)).Value
        wind_sill_height = sheet.Range("P"+str(id)).Value
        
        #Define Zone Objects
        zone = new_idf.newidfobject("zone".upper())
        zone.Name = zone_name
        zone.X_Origin = zone_origin_x
        zone.Y_Origin = zone_origin_y
        zone.Z_Origin = zone_origin_z
        zone.Ceiling_Height = zone_height
    
        #Define BuildingSurface:Detailed Objects, default as adiabatic, and interior constructions 
        surface_1 = new_idf.newidfobject("buildingsurface:detailed".upper())
        surface_1.Name = zone_name+" Surface 1"
        surface_1.Surface_Type = "Floor"
        surface_1.Construction_Name = interior_floor_const
        surface_1.Zone_Name = zone_name
        surface_1.Outside_Boundary_Condition = "Adiabatic"
        surface_1.Sun_Exposure = "NoSun"
        surface_1.Wind_Exposure = "NoWind"
        surface_1.Vertex_1_Xcoordinate = zone_length
        surface_1.Vertex_1_Ycoordinate = zone_width
        surface_1.Vertex_1_Zcoordinate = 0
        surface_1.Vertex_2_Xcoordinate = zone_length
        surface_1.Vertex_2_Ycoordinate = 0
        surface_1.Vertex_2_Zcoordinate = 0   
        surface_1.Vertex_3_Xcoordinate = 0
        surface_1.Vertex_3_Ycoordinate = 0
        surface_1.Vertex_3_Zcoordinate = 0     
        surface_1.Vertex_4_Xcoordinate = 0
        surface_1.Vertex_4_Ycoordinate = zone_width
        surface_1.Vertex_4_Zcoordinate = 0   
        
        surface_2 = new_idf.newidfobject("buildingsurface:detailed".upper())
        surface_2.Name = zone_name+" Surface 2"
        surface_2.Surface_Type = "Wall"
        surface_2.Construction_Name = interior_wall_const
        surface_2.Zone_Name = zone_name
        surface_2.Outside_Boundary_Condition = "Adiabatic"
        surface_2.Sun_Exposure = "NoSun"
        surface_2.Wind_Exposure = "NoWind"
        surface_2.Vertex_1_Xcoordinate = 0
        surface_2.Vertex_1_Ycoordinate = zone_width
        surface_2.Vertex_1_Zcoordinate = zone_height
        surface_2.Vertex_2_Xcoordinate = 0
        surface_2.Vertex_2_Ycoordinate = zone_width
        surface_2.Vertex_2_Zcoordinate = 0   
        surface_2.Vertex_3_Xcoordinate = 0
        surface_2.Vertex_3_Ycoordinate = 0
        surface_2.Vertex_3_Zcoordinate = 0     
        surface_2.Vertex_4_Xcoordinate = 0
        surface_2.Vertex_4_Ycoordinate = 0
        surface_2.Vertex_4_Zcoordinate = zone_height
        
        surface_3 = new_idf.newidfobject("buildingsurface:detailed".upper())
        surface_3.Name = zone_name+" Surface 3"
        surface_3.Surface_Type = "Wall"
        surface_3.Construction_Name = interior_wall_const
        surface_3.Zone_Name = zone_name
        surface_3.Outside_Boundary_Condition = "Adiabatic"
        surface_3.Sun_Exposure = "NoSun"
        surface_3.Wind_Exposure = "NoWind"
        surface_3.Vertex_1_Xcoordinate = zone_length
        surface_3.Vertex_1_Ycoordinate = zone_width
        surface_3.Vertex_1_Zcoordinate = zone_height
        surface_3.Vertex_2_Xcoordinate = zone_length
        surface_3.Vertex_2_Ycoordinate = zone_width
        surface_3.Vertex_2_Zcoordinate = 0   
        surface_3.Vertex_3_Xcoordinate = 0
        surface_3.Vertex_3_Ycoordinate = zone_width
        surface_3.Vertex_3_Zcoordinate = 0     
        surface_3.Vertex_4_Xcoordinate = 0
        surface_3.Vertex_4_Ycoordinate = zone_width
        surface_3.Vertex_4_Zcoordinate = zone_height
      
        surface_4 = new_idf.newidfobject("buildingsurface:detailed".upper())
        surface_4.Name = zone_name+" Surface 4"
        surface_4.Surface_Type = "Wall"
        surface_4.Construction_Name = interior_wall_const
        surface_4.Zone_Name = zone_name
        surface_4.Outside_Boundary_Condition = "Adiabatic"
        surface_4.Sun_Exposure = "NoSun"
        surface_4.Wind_Exposure = "NoWind"
        surface_4.Vertex_1_Xcoordinate = zone_length
        surface_4.Vertex_1_Ycoordinate = 0
        surface_4.Vertex_1_Zcoordinate = zone_height
        surface_4.Vertex_2_Xcoordinate = zone_length
        surface_4.Vertex_2_Ycoordinate = 0
        surface_4.Vertex_2_Zcoordinate = 0   
        surface_4.Vertex_3_Xcoordinate = zone_length
        surface_4.Vertex_3_Ycoordinate = zone_width
        surface_4.Vertex_3_Zcoordinate = 0     
        surface_4.Vertex_4_Xcoordinate = zone_length
        surface_4.Vertex_4_Ycoordinate = zone_width
        surface_4.Vertex_4_Zcoordinate = zone_height
        
        surface_5 = new_idf.newidfobject("buildingsurface:detailed".upper())
        surface_5.Name = zone_name+" Surface 5"
        surface_5.Surface_Type = "Wall"
        surface_5.Construction_Name = interior_wall_const
        surface_5.Zone_Name = zone_name
        surface_5.Outside_Boundary_Condition = "Adiabatic"
        surface_5.Sun_Exposure = "NoSun"
        surface_5.Wind_Exposure = "NoWind"
        surface_5.Vertex_1_Xcoordinate = 0
        surface_5.Vertex_1_Ycoordinate = 0
        surface_5.Vertex_1_Zcoordinate = zone_height
        surface_5.Vertex_2_Xcoordinate = 0
        surface_5.Vertex_2_Ycoordinate = 0
        surface_5.Vertex_2_Zcoordinate = 0   
        surface_5.Vertex_3_Xcoordinate = zone_length
        surface_5.Vertex_3_Ycoordinate = 0
        surface_5.Vertex_3_Zcoordinate = 0     
        surface_5.Vertex_4_Xcoordinate = zone_length
        surface_5.Vertex_4_Ycoordinate = 0
        surface_5.Vertex_4_Zcoordinate = zone_height
        
        surface_6 = new_idf.newidfobject("buildingsurface:detailed".upper())
        surface_6.Name = zone_name+" Surface 6"
        surface_6.Zone_Name = zone_name
        surface_6.Vertex_1_Xcoordinate = zone_length
        surface_6.Vertex_1_Ycoordinate = 0
        surface_6.Vertex_1_Zcoordinate = zone_height
        surface_6.Vertex_2_Xcoordinate = zone_length
        surface_6.Vertex_2_Ycoordinate = zone_width
        surface_6.Vertex_2_Zcoordinate = zone_height   
        surface_6.Vertex_3_Xcoordinate = 0
        surface_6.Vertex_3_Ycoordinate = zone_width
        surface_6.Vertex_3_Zcoordinate = zone_height     
        surface_6.Vertex_4_Xcoordinate = 0
        surface_6.Vertex_4_Ycoordinate = 0
        surface_6.Vertex_4_Zcoordinate = zone_height
    
        #Change Surface Type if Roof
        if roof_or_not == "y":
            surface_6.Surface_Type = "Roof"
            surface_6.Construction_Name = exterior_roof_const
            surface_6.Outside_Boundary_Condition = "Outdoors"
            surface_6.Sun_Exposure = "SunExposed"
            surface_6.Wind_Exposure = "WindExposed"
        else:
            surface_6.Surface_Type = "Ceiling"
            surface_6.Construction_Name = interior_ceiling_const
            surface_6.Outside_Boundary_Condition = "Adiabatic"
            surface_6.Sun_Exposure = "NoSun"
            surface_6.Wind_Exposure = "NoWind"
        
        #Change Surface Type if Ground Floor
        if grndflr_or_not == "y":
            surface_1.Construction_Name = exterior_floor_const
            surface_1.Outside_Boundary_Condition = "Ground"
        else:
            pass
    
    
        #Change Surface Type if Perimeter 1 Normal is non-zero
        if prmtr1_normal == 0:
            surface_3.Construction_Name = exterior_wall_const
            surface_3.Outside_Boundary_Condition = "Outdoors"
            surface_3.Sun_Exposure = "SunExposed"
            surface_3.Wind_Exposure = "WindExposed"
            if wind2wall_ratio>0:
                sub_surface = new_idf.newidfobject("FenestrationSurface:Detailed".upper())
                sub_surface.Name = zone_name+ " Sub Surface 3"
                sub_surface.Surface_Type = "Window"
                sub_surface.Construction_Name = exterior_window_const
                sub_surface.Building_Surface_Name = surface_3.Name
                sub_surface.Vertex_1_Xcoordinate = zone_length-0.0254
                sub_surface.Vertex_1_Ycoordinate = zone_width
                sub_surface.Vertex_1_Zcoordinate = zone_length*zone_height*wind2wall_ratio/(zone_length-0.0254*2)+wind_sill_height
                sub_surface.Vertex_2_Xcoordinate = zone_length-0.0254
                sub_surface.Vertex_2_Ycoordinate = zone_width
                sub_surface.Vertex_2_Zcoordinate = wind_sill_height
                sub_surface.Vertex_3_Xcoordinate = 0.0254
                sub_surface.Vertex_3_Ycoordinate = zone_width
                sub_surface.Vertex_3_Zcoordinate = wind_sill_height
                sub_surface.Vertex_4_Xcoordinate = 0.0254
                sub_surface.Vertex_4_Ycoordinate = zone_width
                sub_surface.Vertex_4_Zcoordinate = sub_surface.Vertex_1_Zcoordinate
            else:
                pass
        elif prmtr1_normal == 90:
            surface_4.Construction_Name = exterior_wall_const
            surface_4.Outside_Boundary_Condition = "Outdoors"
            surface_4.Sun_Exposure = "SunExposed"
            surface_4.Wind_Exposure = "WindExposed"
            if wind2wall_ratio>0:
                sub_surface = new_idf.newidfobject("FenestrationSurface:Detailed".upper())
                sub_surface.Name = zone_name+ " Sub Surface 4"
                sub_surface.Surface_Type = "Window"
                sub_surface.Construction_Name = exterior_window_const
                sub_surface.Building_Surface_Name = surface_4.Name
                sub_surface.Vertex_1_Xcoordinate = zone_length
                sub_surface.Vertex_1_Ycoordinate = 0.0254
                sub_surface.Vertex_1_Zcoordinate = zone_width*zone_height*wind2wall_ratio/(zone_width-0.0254*2)+wind_sill_height
#                sub_surface.Vertex_1_Zcoordinate = zone_length*zone_height*wind2wall_ratio/sub_surface.Vertex_1_Xcoordinate+wind_sill_height
                sub_surface.Vertex_2_Xcoordinate = zone_length
                sub_surface.Vertex_2_Ycoordinate = 0.0254
                sub_surface.Vertex_2_Zcoordinate = wind_sill_height
                sub_surface.Vertex_3_Xcoordinate = zone_length
                sub_surface.Vertex_3_Ycoordinate = zone_width-0.0254
                sub_surface.Vertex_3_Zcoordinate = wind_sill_height
                sub_surface.Vertex_4_Xcoordinate = zone_length
                sub_surface.Vertex_4_Ycoordinate = zone_width-0.0254
                sub_surface.Vertex_4_Zcoordinate = sub_surface.Vertex_1_Zcoordinate
            else:
                pass
        elif prmtr1_normal == 180:  
            surface_5.Construction_Name = exterior_wall_const
            surface_5.Outside_Boundary_Condition = "Outdoors"
            surface_5.Sun_Exposure = "SunExposed"
            surface_5.Wind_Exposure = "WindExposed"
            if wind2wall_ratio>0:
                sub_surface = new_idf.newidfobject("FenestrationSurface:Detailed".upper())
                sub_surface.Name = zone_name+ " Sub Surface 5"
                sub_surface.Surface_Type = "Window"
                sub_surface.Construction_Name = exterior_window_const
                sub_surface.Building_Surface_Name = surface_5.Name
                sub_surface.Vertex_1_Xcoordinate = 0.0254
                sub_surface.Vertex_1_Ycoordinate = 0
                sub_surface.Vertex_1_Zcoordinate = zone_length*zone_height*wind2wall_ratio/(zone_length-0.0254*2)+wind_sill_height
#                sub_surface.Vertex_1_Zcoordinate = zone_length*zone_height*wind2wall_ratio/sub_surface.Vertex_1_Xcoordinate+wind_sill_height
                sub_surface.Vertex_2_Xcoordinate = 0.0254
                sub_surface.Vertex_2_Ycoordinate = 0
                sub_surface.Vertex_2_Zcoordinate = wind_sill_height
                sub_surface.Vertex_3_Xcoordinate = zone_length-0.0254
                sub_surface.Vertex_3_Ycoordinate = 0
                sub_surface.Vertex_3_Zcoordinate = wind_sill_height
                sub_surface.Vertex_4_Xcoordinate = zone_length-0.0254
                sub_surface.Vertex_4_Ycoordinate = 0
                sub_surface.Vertex_4_Zcoordinate = sub_surface.Vertex_1_Zcoordinate
            else:
                pass
        elif prmtr1_normal == 270:    
            surface_2.Construction_Name = exterior_wall_const
            surface_2.Outside_Boundary_Condition = "Outdoors"
            surface_2.Sun_Exposure = "SunExposed"
            surface_2.Wind_Exposure = "WindExposed"
            if wind2wall_ratio>0:
                sub_surface = new_idf.newidfobject("FenestrationSurface:Detailed".upper())
                sub_surface.Name = zone_name+ " Sub Surface 2"
                sub_surface.Surface_Type = "Window"
                sub_surface.Construction_Name = exterior_window_const
                sub_surface.Building_Surface_Name = surface_2.Name
                sub_surface.Vertex_1_Xcoordinate = 0
                sub_surface.Vertex_1_Ycoordinate = zone_width-0.0254
                sub_surface.Vertex_1_Zcoordinate = zone_width*zone_height*wind2wall_ratio/(zone_width-2*0.0254)+wind_sill_height
#                sub_surface.Vertex_1_Zcoordinate = zone_length*zone_height*wind2wall_ratio/sub_surface.Vertex_1_Xcoordinate+wind_sill_height
                sub_surface.Vertex_2_Xcoordinate = 0
                sub_surface.Vertex_2_Ycoordinate = zone_width-0.0254
                sub_surface.Vertex_2_Zcoordinate = wind_sill_height
                sub_surface.Vertex_3_Xcoordinate = 0
                sub_surface.Vertex_3_Ycoordinate = 0.0254
                sub_surface.Vertex_3_Zcoordinate = wind_sill_height
                sub_surface.Vertex_4_Xcoordinate = 0
                sub_surface.Vertex_4_Ycoordinate = 0.0254
                sub_surface.Vertex_4_Zcoordinate = sub_surface.Vertex_1_Zcoordinate
            else:
                pass
        else:
            pass
            
        #Change Surface Type if Perimeter 2 Normal is non-zero

        if prmtr2_normal == 0:
            surface_3.Construction_Name = exterior_wall_const
            surface_3.Outside_Boundary_Condition = "Outdoors"
            surface_3.Sun_Exposure = "SunExposed"
            surface_3.Wind_Exposure = "WindExposed"
            if wind2wall_ratio>0:
                sub_surface = new_idf.newidfobject("FenestrationSurface:Detailed".upper())
                sub_surface.Name = zone_name+ " Sub Surface 3"
                sub_surface.Surface_Type = "Window"
                sub_surface.Construction_Name = exterior_window_const
                sub_surface.Building_Surface_Name = surface_3.Name
                sub_surface.Vertex_1_Xcoordinate = zone_length-0.0254
                sub_surface.Vertex_1_Ycoordinate = zone_width
                sub_surface.Vertex_1_Zcoordinate = zone_length*zone_height*wind2wall_ratio/(zone_length-0.0254*2)+wind_sill_height
                sub_surface.Vertex_2_Xcoordinate = zone_length-0.0254
                sub_surface.Vertex_2_Ycoordinate = zone_width
                sub_surface.Vertex_2_Zcoordinate = wind_sill_height
                sub_surface.Vertex_3_Xcoordinate = 0.0254
                sub_surface.Vertex_3_Ycoordinate = zone_width
                sub_surface.Vertex_3_Zcoordinate = wind_sill_height
                sub_surface.Vertex_4_Xcoordinate = 0.0254
                sub_surface.Vertex_4_Ycoordinate = zone_width
                sub_surface.Vertex_4_Zcoordinate = sub_surface.Vertex_1_Zcoordinate
            else:
                pass
        elif prmtr2_normal == 90:
            surface_4.Construction_Name = exterior_wall_const
            surface_4.Outside_Boundary_Condition = "Outdoors"
            surface_4.Sun_Exposure = "SunExposed"
            surface_4.Wind_Exposure = "WindExposed"
            if wind2wall_ratio>0:
                sub_surface = new_idf.newidfobject("FenestrationSurface:Detailed".upper())
                sub_surface.Name = zone_name+ " Sub Surface 4"
                sub_surface.Surface_Type = "Window"
                sub_surface.Construction_Name = exterior_window_const
                sub_surface.Building_Surface_Name = surface_4.Name
                sub_surface.Vertex_1_Xcoordinate = zone_length
                sub_surface.Vertex_1_Ycoordinate = 0.0254
                sub_surface.Vertex_1_Zcoordinate = zone_width*zone_height*wind2wall_ratio/(zone_width-0.0254*2)+wind_sill_height
#                sub_surface.Vertex_1_Zcoordinate = zone_length*zone_height*wind2wall_ratio/sub_surface.Vertex_1_Xcoordinate+wind_sill_height
                sub_surface.Vertex_2_Xcoordinate = zone_length
                sub_surface.Vertex_2_Ycoordinate = 0.0254
                sub_surface.Vertex_2_Zcoordinate = wind_sill_height
                sub_surface.Vertex_3_Xcoordinate = zone_length
                sub_surface.Vertex_3_Ycoordinate = zone_width-0.0254
                sub_surface.Vertex_3_Zcoordinate = wind_sill_height
                sub_surface.Vertex_4_Xcoordinate = zone_length
                sub_surface.Vertex_4_Ycoordinate = zone_width-0.0254
                sub_surface.Vertex_4_Zcoordinate = sub_surface.Vertex_1_Zcoordinate
            else:
                pass
        elif prmtr2_normal == 180:  
            surface_5.Construction_Name = exterior_wall_const
            surface_5.Outside_Boundary_Condition = "Outdoors"
            surface_5.Sun_Exposure = "SunExposed"
            surface_5.Wind_Exposure = "WindExposed"
            if wind2wall_ratio>0:
                sub_surface = new_idf.newidfobject("FenestrationSurface:Detailed".upper())
                sub_surface.Name = zone_name+ " Sub Surface 5"
                sub_surface.Surface_Type = "Window"
                sub_surface.Construction_Name = exterior_window_const
                sub_surface.Building_Surface_Name = surface_5.Name
                sub_surface.Vertex_1_Xcoordinate = 0.0254
                sub_surface.Vertex_1_Ycoordinate = 0
                sub_surface.Vertex_1_Zcoordinate = zone_length*zone_height*wind2wall_ratio/(zone_length-0.0254*2)+wind_sill_height
#                sub_surface.Vertex_1_Zcoordinate = zone_length*zone_height*wind2wall_ratio/sub_surface.Vertex_1_Xcoordinate+wind_sill_height
                sub_surface.Vertex_2_Xcoordinate = 0.0254
                sub_surface.Vertex_2_Ycoordinate = 0
                sub_surface.Vertex_2_Zcoordinate = wind_sill_height
                sub_surface.Vertex_3_Xcoordinate = zone_length-0.0254
                sub_surface.Vertex_3_Ycoordinate = 0
                sub_surface.Vertex_3_Zcoordinate = wind_sill_height
                sub_surface.Vertex_4_Xcoordinate = zone_length-0.0254
                sub_surface.Vertex_4_Ycoordinate = 0
                sub_surface.Vertex_4_Zcoordinate = sub_surface.Vertex_1_Zcoordinate
            else:
                pass
        elif prmtr2_normal == 270:    
            surface_2.Construction_Name = exterior_wall_const
            surface_2.Outside_Boundary_Condition = "Outdoors"
            surface_2.Sun_Exposure = "SunExposed"
            surface_2.Wind_Exposure = "WindExposed"
            if wind2wall_ratio>0:
                sub_surface = new_idf.newidfobject("FenestrationSurface:Detailed".upper())
                sub_surface.Name = zone_name+ " Sub Surface 2"
                sub_surface.Surface_Type = "Window"
                sub_surface.Construction_Name = exterior_window_const
                sub_surface.Building_Surface_Name = surface_2.Name
                sub_surface.Vertex_1_Xcoordinate = 0
                sub_surface.Vertex_1_Ycoordinate = zone_width-0.0254
                sub_surface.Vertex_1_Zcoordinate = zone_width*zone_height*wind2wall_ratio/(zone_width-2*0.0254)+wind_sill_height
#                sub_surface.Vertex_1_Zcoordinate = zone_length*zone_height*wind2wall_ratio/sub_surface.Vertex_1_Xcoordinate+wind_sill_height
                sub_surface.Vertex_2_Xcoordinate = 0
                sub_surface.Vertex_2_Ycoordinate = zone_width-0.0254
                sub_surface.Vertex_2_Zcoordinate = wind_sill_height
                sub_surface.Vertex_3_Xcoordinate = 0
                sub_surface.Vertex_3_Ycoordinate = 0.0254
                sub_surface.Vertex_3_Zcoordinate = wind_sill_height
                sub_surface.Vertex_4_Xcoordinate = 0
                sub_surface.Vertex_4_Ycoordinate = 0.0254
                sub_surface.Vertex_4_Zcoordinate = sub_surface.Vertex_1_Zcoordinate
            else:
                pass
        else:
            pass

            
    new_idf.save()    
    this_idf = new_idf.idfname
#    dir_path = os.path.dirname(os.path.realpath(this_idf))
#    
#    os.chdir('..')
#    shutil.copy2(dir_path, '/templates/) 
    # complete target filename given
#    shutil.copy2('/src/file.ext', '/dst/dir') # target filename is /dst/dir/file.ext
    pre, ext = os.path.splitext(this_idf)
    os.rename(this_idf, pre+".pxt")

    
#@xlsub
#@xlarg("sheet", vba="Sheet5")
#def my_other_macro(sheet):
#    sheet.Range("A3").Value
#                  
            
