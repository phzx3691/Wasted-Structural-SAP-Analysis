import numpy as np
import pandas as pd
import comtypes.client # package required to use the API. It should already be on your machine.
import openpyxl

def find_below(point_labels, point_coordinates, Area_Point_Names, area_point_Names_All, bay_size_z):
    out = []
    element_names_1 = np.where(np.array(point_labels) == Area_Point_Names[1][0])[0][0]
    element_names_2 = np.where(np.array(point_labels) == Area_Point_Names[1][1])[0][0]
    element_names_3 = np.where(np.array(point_labels) == Area_Point_Names[1][2])[0][0]
    element_names_4 = np.where(np.array(point_labels) == Area_Point_Names[1][3])[0][0]
    
    for u in range(4):
        Deletion_point_1 = np.array([0, 0, -bay_size_z*u, 0]) + point_coordinates[element_names_1]
        Deletion_point_2 = np.array([0, 0, -bay_size_z*u, 0]) + point_coordinates[element_names_2]
        Deletion_point_3 = np.array([0, 0, -bay_size_z*u, 0]) + point_coordinates[element_names_3]
        Deletion_point_4 = np.array([0, 0, -bay_size_z*u, 0]) + point_coordinates[element_names_4]
        
        m = 0
        while True:
            if np.array(point_coordinates)[m,0] == Deletion_point_1[0]:
                if np.array(point_coordinates)[m,1] == Deletion_point_1[1]:
                    if np.array(point_coordinates)[m,2] == Deletion_point_1[2]:
                        break
            m += 1
            
        Labels_of_Area_Points_1 = point_labels[m]
        
        m = 0
        while True:
            if np.array(point_coordinates)[m,0] == Deletion_point_2[0]:
                if np.array(point_coordinates)[m,1] == Deletion_point_2[1]:
                    if np.array(point_coordinates)[m,2] == Deletion_point_2[2]:
                        break
            m += 1
            
        Labels_of_Area_Points_2 = point_labels[m]
        
        m = 0
        while True:
            if np.array(point_coordinates)[m,0] == Deletion_point_3[0]:
                if np.array(point_coordinates)[m,1] == Deletion_point_3[1]:
                    if np.array(point_coordinates)[m,2] == Deletion_point_3[2]:
                        break
            m += 1
            
        Labels_of_Area_Points_3 = point_labels[m]
        
        m = 0
        while True:
            if np.array(point_coordinates)[m,0] == Deletion_point_4[0]:
                if np.array(point_coordinates)[m,1] == Deletion_point_4[1]:
                    if np.array(point_coordinates)[m,2] == Deletion_point_4[2]:
                        break
            m += 1
            
        Labels_of_Area_Points_4 = point_labels[m]
        
        m = 0
        while True:
            if np.array(area_point_Names_All[m][0]) == Labels_of_Area_Points_1 or np.array(area_point_Names_All[m][0]) == Labels_of_Area_Points_2 or np.array(area_point_Names_All[m][0]) == Labels_of_Area_Points_3 or np.array(area_point_Names_All[m][0]) == Labels_of_Area_Points_4:
                if np.array(area_point_Names_All[m][1]) == Labels_of_Area_Points_1 or np.array(area_point_Names_All[m][1]) == Labels_of_Area_Points_2 or np.array(area_point_Names_All[m][1]) == Labels_of_Area_Points_3 or np.array(area_point_Names_All[m][1]) == Labels_of_Area_Points_4:
                    if np.array(area_point_Names_All[m][2]) == Labels_of_Area_Points_1 or np.array(area_point_Names_All[m][2]) == Labels_of_Area_Points_2 or np.array(area_point_Names_All[m][2]) == Labels_of_Area_Points_3 or np.array(area_point_Names_All[m][2]) == Labels_of_Area_Points_4:
                        if np.array(area_point_Names_All[m][3]) == Labels_of_Area_Points_1 or np.array(area_point_Names_All[m][3]) == Labels_of_Area_Points_2 or np.array(area_point_Names_All[m][3]) == Labels_of_Area_Points_3 or np.array(area_point_Names_All[m][3]) == Labels_of_Area_Points_4:
                            break
            m += 1
        out.append(m)
    return out
sap_object = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")
sap_model = sap_object.SapModel
for i in range(1):# range(sap_model.AreaObj.GetNameList()[0]):
    # Connect Python to the SAP model
    
    
    sap_model.InitializeNewModel
    sap_model.File.OpenFile('Model_Original.sdb')
    
    # Unlock the model / make sure it is unlocked:
    sap_model.SetModelIsLocked(False)
    
    # # To use the metric system, do these below:
    # kN_mm_C = 6 #6 is from the documentation!
    # sap_model.SetPresentUnits(kN_mm_C)
    
    # ...
    
    point_labels = sap_model.PointObj.GetNameList()[1]
    
    point_coordinates = []
    
    for n in range( len(point_labels)):
        point_coordinates.append( sap_model.PointObj.GetCoordCartesian(point_labels[n]))
        
    area_point_Names_All = []
    area_labels = sap_model.AreaObj.GetNameList()[1]
    for n in range( len(area_labels)):
        area_point_Names_All.append( list(sap_model.AreaObj.GetPoints(area_labels[n])[1]))
    
    
    bay_size_x = 24
    bay_size_y = 24
    bay_size_z = 12
    levels = 4
    roof_live = .10 #k/sf
    roof_dead = 150*3*0.9/1000 #k/sf
     
    Area_Objects_Top = []
    Num_of_Areas = []
    
    i = 14 #22 #14
    
    Area_Point_Names = sap_model.AreaObj.GetPoints(sap_model.AreaObj.GetNameList()[1][i])

    Area_Point_Coord = [sap_model.PointObj.GetCoordCartesian(Area_Point_Names[1][0]),
                        sap_model.PointObj.GetCoordCartesian(Area_Point_Names[1][1]),
                        sap_model.PointObj.GetCoordCartesian(Area_Point_Names[1][2]),
                        sap_model.PointObj.GetCoordCartesian(Area_Point_Names[1][3])]
    
    if Area_Point_Coord[0][2] == bay_size_z*levels:
        #Area_Objects_Top.append([sap_model.AreaObj.GetNameList()[1][i], Area_Point_Coord])
        #sap_model.AreaObj.SetSelected(sap_model.AreaObj.GetNameList()[1][i], True)
        #sap_model.PointObj.SetSelected(Area_Point_Names[1][0], True)
        
        Point_Connectivity_Areas = np.array([sum(1 for x in list(sap_model.PointObj.GetConnectivity(Area_Point_Names[1][0])[1]) if x == 5),
                            sum(1 for x in list(sap_model.PointObj.GetConnectivity(Area_Point_Names[1][1])[1]) if x == 5),
                            sum(1 for x in list(sap_model.PointObj.GetConnectivity(Area_Point_Names[1][2])[1]) if x == 5),
                            sum(1 for x in list(sap_model.PointObj.GetConnectivity(Area_Point_Names[1][3])[1]) if x == 5)])

        Num_of_Areas = np.min(Point_Connectivity_Areas)
        Num_of_Areas_idx = np.argmin(Point_Connectivity_Areas)
        Min_Area_Point_Label = Area_Point_Names[1][Num_of_Areas_idx]
        
        sap_model.AreaObj.SetLoadUniformToFrame(sap_model.AreaObj.GetNameList()[1][i], "Live", roof_live,            10, 2, True, "Global")
        sap_model.AreaObj.SetLoadUniformToFrame(sap_model.AreaObj.GetNameList()[1][i], "Dead", roof_dead, 10, 2, True, "Global")
        
        sap_model.Analyze.RunAnalysis()
        sap_model.DesignSteel.StartDesign()
        
        number_failed = sap_model.DesignSteel.VerifyPassed()
        
        if number_failed[0] > 0:
            save_filepath = 'Roof_Model_' + str(i) + '.sdb'
            #sap_model.File.Save(save_filepath)
            
            sap_model.SetModelIsLocked(False)
        
            if Num_of_Areas == 1:
                Labels_to_delete = list(sap_model.PointObj.GetConnectivity(Min_Area_Point_Label))[2]
                
                Area_Stack = find_below(point_labels, point_coordinates, Area_Point_Names, area_point_Names_All, bay_size_z)
                
                sap_model.FrameObj.Delete(Labels_to_delete[0], 0)
                sap_model.FrameObj.Delete(Labels_to_delete[1], 0)
                sap_model.FrameObj.Delete(Labels_to_delete[2], 0)
                sap_model.AreaObj.Delete(Labels_to_delete[3], 0)
            
                deletion_type = 1
            
            if Num_of_Areas == 2:
                idx = np.where(Point_Connectivity_Areas == 2)[0]
                element_names_1 = list(sap_model.PointObj.GetConnectivity(Area_Point_Names[1][idx[0]])[2])
                element_names_2 = list(sap_model.PointObj.GetConnectivity(Area_Point_Names[1][idx[1]])[2])
                
                
                Labels_to_delete = [element for element in element_names_1 if element in element_names_2]
                
                Area_Stack = find_below(point_labels, point_coordinates, Area_Point_Names, area_point_Names_All, bay_size_z)
                
                sap_model.FrameObj.Delete(Labels_to_delete[0], 0)
                sap_model.AreaObj.Delete(Labels_to_delete[1], 0)
            
                deletion_type = 2
                
            if Num_of_Areas == 3 or Num_of_Areas == 4:
                
                Area_Stack = find_below(point_labels, point_coordinates, Area_Point_Names, area_point_Names_All, bay_size_z)
                
                sap_model.AreaObj.Delete(sap_model.AreaObj.GetNameList()[1][i], 0)
                
                deletion_type = 3
                         
            sap_model.AreaObj.SetLoadUniformToFrame(area_labels[Area_Stack[1]], "Live", roof_live,           10, 2, True, "Global")
            sap_model.AreaObj.SetLoadUniformToFrame(area_labels[Area_Stack[1]], "Dead", roof_dead, 10, 2, True, "Global")
            
            sap_model.Analyze.RunAnalysis()
            sap_model.DesignSteel.StartDesign()
        
            number_failed = sap_model.DesignSteel.VerifyPassed()
        
            if number_failed[0] > 0:
                save_filepath = 'Roof_Model_' + str(i) + '.sdb'
                #sap_model.File.Save(save_filepath)
                sap_model.SetModelIsLocked(False)
                
                Area_Objects_Next = []
                
                Area_Point_Names = sap_model.AreaObj.GetPoints(area_labels[Area_Stack[1]])

                Area_Point_Coord = [sap_model.PointObj.GetCoordCartesian(Area_Point_Names[1][0]),
                                    sap_model.PointObj.GetCoordCartesian(Area_Point_Names[1][1]),
                                    sap_model.PointObj.GetCoordCartesian(Area_Point_Names[1][2]),
                                    sap_model.PointObj.GetCoordCartesian(Area_Point_Names[1][3])]
                                
                Point_Connectivity_Areas = np.array([sum(1 for x in list(sap_model.PointObj.GetConnectivity(Area_Point_Names[1][0])[1]) if x == 5),
                                    sum(1 for x in list(sap_model.PointObj.GetConnectivity(Area_Point_Names[1][1])[1]) if x == 5),
                                    sum(1 for x in list(sap_model.PointObj.GetConnectivity(Area_Point_Names[1][2])[1]) if x == 5),
                                    sum(1 for x in list(sap_model.PointObj.GetConnectivity(Area_Point_Names[1][3])[1]) if x == 5)])
    
                Num_of_Areas = np.min(Point_Connectivity_Areas)
                Num_of_Areas_idx = np.argmin(Point_Connectivity_Areas)
                Min_Area_Point_Label = Area_Point_Names[1][Num_of_Areas_idx]
            
                if deletion_type == 1:
                    
                    Deletion_point = np.array([sap_model.PointObj.GetCoordCartesian(Min_Area_Point_Label)])
                    
                    #Deletion_point -= np.array([0, 0, bay_size_z, 0])
                    
                    m = 0
                    while True:
                        if np.array(point_coordinates)[m,0] == Deletion_point[0,0]:
                            if np.array(point_coordinates)[m,1] == Deletion_point[0,1]:
                                if np.array(point_coordinates)[m,2] == Deletion_point[0,2]:
                                    break
                        m += 1
                        
                    Labels_to_delete = list(sap_model.PointObj.GetConnectivity(point_labels[m] ))[2]
                    
                    sap_model.FrameObj.Delete(Labels_to_delete[0], 0)
                    sap_model.FrameObj.Delete(Labels_to_delete[1], 0)
                    sap_model.FrameObj.Delete(Labels_to_delete[2], 0)
                    sap_model.AreaObj.Delete(Labels_to_delete[3], 0)
                
                if deletion_type == 2:
                    
                    idx = np.where(Point_Connectivity_Areas == 2)[0]
                    
                    Labels_to_delete = [element for element in element_names_1 if element in element_names_2]
                    
                    element_names_1 = np.where(np.array(point_labels) == Area_Point_Names[1][idx[0]])[0][0]
                    element_names_2 = np.where(np.array(point_labels) == Area_Point_Names[1][idx[1]])[0][0]
                    
                    Deletion_point_1 = np.array([0, 0, -bay_size_z, 0]) + point_coordinates[element_names_1]
                    Deletion_point_2 = np.array([0, 0, -bay_size_z, 0]) + point_coordinates[element_names_2]
                    
                    m = 0
                    while True:
                        if np.array(point_coordinates)[m,0] == Deletion_point_1[0]:
                            if np.array(point_coordinates)[m,1] == Deletion_point_1[1]:
                                if np.array(point_coordinates)[m,2] == Deletion_point_1[2]:
                                    break
                        m += 1
                        
                    Labels_to_delete_1 = list(sap_model.PointObj.GetConnectivity(point_labels[m] ))[2]
                    
                    m = 0
                    while True:
                        if np.array(point_coordinates)[m,0] == Deletion_point_2[0]:
                            if np.array(point_coordinates)[m,1] == Deletion_point_2[1]:
                                if np.array(point_coordinates)[m,2] == Deletion_point_2[2]:
                                    break
                        m += 1
                        
                    Labels_to_delete_2 = list(sap_model.PointObj.GetConnectivity(point_labels[m] ))[2]
    
                    
                    sap_model.FrameObj.Delete(Labels_to_delete[0], 0)
                    sap_model.AreaObj.Delete(Labels_to_delete[1], 0)
                
                if deletion_type == 3:
                    
                    m = find_below(point_labels, point_coordinates, Area_Point_Names, area_point_Names_All, bay_size_z)
                    
                    sap_model.AreaObj.Delete(sap_model.AreaObj.GetNameList()[1][m], 0)
        
                sap_model.AreaObj.SetLoadUniformToFrame(area_labels[Area_Stack[2]], "Live", roof_live,           10, 2, True, "Global")
                sap_model.AreaObj.SetLoadUniformToFrame(area_labels[Area_Stack[2]], "Dead", roof_dead, 10, 2, True, "Global")
                
                sap_model.Analyze.RunAnalysis()
                sap_model.DesignSteel.StartDesign()
            
                number_failed = sap_model.DesignSteel.VerifyPassed()
    
    
    #if sap_model.AreaObj.GetNameList()[1][i] == '26':
    #    break        




