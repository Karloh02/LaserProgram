import ezdxf
import math


def laser_calculator(directory_user, thickness_user, material_user):
    sheet_thickness = ["1", "1.5", "2", "2.5", "3", "4", "5", "6", "8", "10", "12", "15"]

    #carbon sheets parameters
    time_cut_carb = [0.35, 0.37, 0.40, 0.42, 0.45, 0.51, 0.58, 0.66, 0.85, 1.1, 1.42, 2.08]
    time_hole_carb = [0.0002, 0.0005, 0.0006, 0.0007, 0.0008, 0.0023, 0.0035, 0.0043, 0.0087, 0.0286, 0.0624, 0.298]

    #stainless sheets parameters
    time_cut_stain = [0.25, 0.29, 0.33, 0.39, 0.45, 0.6, 0.81, 1.08, 1.96, 3.54, 6.4]
    time_hole_stain = [0.00072, 0.00087, 0.00122, 0.00149, 0.00176, 0.0026, 0.0039, 0.0052, 0.00823, 0.013, 0.0273]

    #parameters to calculate
    total_perimeter = 0             #Total perimeter of the component
    bends = 0                       #Total number of bends                         
    interior_profiles = 0           #Total number of interior profiles - Holes

    #user inputs
    file = directory_user 
    material = material_user
    thickness = str(thickness_user)

    #reading the dxf file - saving it as a modelspace 
    dwg = ezdxf.readfile(file)
    msp = dwg.modelspace()

    #this for function will iterate trough all the modelspace parameters 
    for e in msp:
        
        #This part will determine how many lines there are in the file. 
        if e.dxf.layer == "IV_BEND_DOWN" or e.dxf.layer == "IV_BEND":
            bends += 1

        #This part will read the layers, and calculate the perimeter of those respective layers.
        #If it is not a defined bend line, it will measured into the perimeter
        else: #e.dxf.layer == "OUTER" or e.dxf.layer == "IV_ARC_CENTERS" or e.dxf.layer == "IV_INTERIOR_PROFILES" or e.dxf.layer == "IV_TOOL_CENTER" or e.dxf.layer == "IV_TOOL_CENTER_DOWN" or e.dxf.layer == "IV_FEATURE_PROFILES" or e.dxf.layer == "AM_KONT" or e.dxf.layer == "Visible (ISO)" or e.dxf.layer == "0": 

            #this part server the pupose to read how many holes there are in the part
            if e.dxf.layer == "IV_INTERIOR_PROFILES":
                interior_profiles += 1
            
            #This part reads the distance "dl" of lines 
            if e.dxftype() == "LINE":
                
                dl = abs(math.sqrt((e.dxf.start[0] - e.dxf.end[0])**2 + (e.dxf.start[1] - e.dxf.end[1])**2)) 
                total_perimeter += dl
            
            #This part reads the distance "dc" of circles
            elif e.dxftype() == "CIRCLE": 

                dc = 2*math.pi*e.dxf.radius
                total_perimeter += dc
            
            #This part reads the distance "da" of arcs
            elif e.dxftype() == "ARC":
                
                end = e.dxf.end_angle
                start = e.dxf.start_angle
                angle = abs(start - end)

                if end == 0:
                    end = 360 
                    angle = abs(start - end)
                
                da = abs(2*math.pi*e.dxf.radius*angle/360)
                total_perimeter += da

            #This part read the distance "ds" of splines - an aproximation since splines are curved.
            elif e.dxftype() == "SPLINE":

                points = e.control_points
                
                for i in range(len(points) - 1):
                    ds = math.sqrt((points[i][0] - points[i + 1][0])**2 + (points[i][1] - points[i + 1][1])**2)
                    total_perimeter += ds

            #This part reads the distance "dp" of polylines - an aproximation since polylines can be curved 
            elif e.dxftype() == "POLYLINE":
                
                poly_points = e.points()

                polyline_vectors = []
                for k in poly_points:
                    polyline_vectors.append(ezdxf.math.Vec3(k))
                
                for f in range(len(polyline_vectors) - 1):
                    dp = math.sqrt((polyline_vectors[f][0] - polyline_vectors[f + 1][0])**2 + (polyline_vectors[f][1] - polyline_vectors[f + 1][1])**2)
                    total_perimeter += dp*1.05
                
                total_perimeter += (math.sqrt((polyline_vectors[0][0] - polyline_vectors[len(polyline_vectors) - 1][0])**2 + (polyline_vectors[0][1] - polyline_vectors[len(polyline_vectors) - 1][1])**2))*1.05

    #Where the time of the cut is calculated
    total_perimeter = round(total_perimeter)  

                
    if material == "Carbono":
        cut_time = round((total_perimeter/1000)*time_cut_carb[sheet_thickness.index(thickness)] + time_hole_carb[sheet_thickness.index(thickness)]*(interior_profiles + 1), 2)

    elif material == "Inox":
        cut_time = round((total_perimeter/1000)*time_cut_stain[sheet_thickness.index(thickness)] + time_hole_stain[sheet_thickness.index(thickness)]*(interior_profiles + 1), 2)
        
    #The function returns the time it takes to cut the part and, the number of bends; information useful in the future.

    return(cut_time, bends)