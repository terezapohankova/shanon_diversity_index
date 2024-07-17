# -*- coding: utf-8 -*-

import arcpy
import math



class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Shanon Diversity Index"
        self.alias = "SDI"

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Shanon Diversity Index"
        self.description = "Calculation of Shanon Diversity Index on https://fragstats.org/index.php/fragstats-metrics/patch-based-metrics/diversity-metrics/l4-shannons-diversity-index"
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
       
       
        # Input SHP layer with intersected grid and land patches
        param0 = arcpy.Parameter(
            displayName="Input Layer",
            name="input_layer",
            datatype=["DEShapeFile","DEFeatureClass"],
            parameterType="Required",
            direction="Input"
            )
        
     
        # Output SHP Layer
        param1 = arcpy.Parameter(
            displayName="Output Layer",
            name="out_features",
            datatype=["DEShapeFile","DEFeatureClass"],
            parameterType="Required",
            direction="Output")
        
    
        # Name of field with information about gridcell ID
        param2 = arcpy.Parameter(
            displayName="GRID ID",
            name="GRID_id",
            datatype="Field",
            parameterType="Required",
            direction="Input")
        
        
        param1.parameterDependencies = [param0.name]
        param1.schema.clone = True
        
        params = [param0, param1, param2]
        
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        
        # get text value of parameters
        INPUT_LAYER = parameters[0].valueAsText
        OUTPUT_LAYER = parameters[1].valueAsText
        GRID_ID = parameters[2].valueAsText
        
        # Define new field names
        SHAPE_AREA_FIELD_NAME = "AREA_M"  # shape area of one patch in sq m
        
        # Define datype for fields
        FIELD_TYPE = "FLOAT"
    
        # Create copy of input layer to preserve original data
        input_layer_copy = arcpy.Copy_management(
            INPUT_LAYER, 
            OUTPUT_LAYER)
        
        # add field for calculation for perimeter of the land patch    
        arcpy.management.AddField(
            input_layer_copy, 
            SHAPE_AREA_FIELD_NAME, 
            FIELD_TYPE)
        
        """# add field for calculation for gridcell area
        arcpy.management.AddField(
            input_layer_copy, 
            GRID_AREA_FIELD_NAME, 
            FIELD_TYPE)"""
        
        # calculate area of land patch and write it into column
        arcpy.management.CalculateGeometryAttributes(
            input_layer_copy,
            [[SHAPE_AREA_FIELD_NAME, "AREA"]],
            length_unit = "METERS",
            area_unit = "SQUARE_METERS"
        )
        
      # sum patch perimeter and patch area for each ID
        # write it into separate table and store in memory
        summary_table = r"in_memory/summary_table"
        arcpy.analysis.Statistics(
            OUTPUT_LAYER,
            summary_table,
            [[SHAPE_AREA_FIELD_NAME, "SUM"]],
            GRID_ID
        )
            
        # Join the memory-help summary statistics back to the output layer
        # based on Grid ID
        arcpy.management.JoinField(
            OUTPUT_LAYER,
            GRID_ID,
            summary_table,
            GRID_ID,
            [f"SUM_{SHAPE_AREA_FIELD_NAME}"],
        )
        
        arcpy.conversion.TableToExcel(
         summary_table,
         "summary.xlsx"
        )
        
        
        return


