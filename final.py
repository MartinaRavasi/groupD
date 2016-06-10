from openpyxl import *

#define a function that compute the interpolation between two values
def interpolate(x1, x2, y1, y2, x_inputss):
    """
    this function gives y_inputss computing linear interpolation between (x1,y1) and (x2,y2) and finding the corresponding value for x_inputss
    """
    y_inputss = (y1 - y2)*(x_inputss - x2)/(x1-x2) + y2;
    return y_inputss
    
#define a funcition that gives back distribution losses 
def distributionLosses(input_duct, input_insulation, input_leakage, input_conditioning, input_stories, input_load):
    """
this function takes as input:
input_duct: Duct location "Attic"/"Basement"/"Crawlspace"/"Conditioned space" string
input_insulation: Insulation [m^2*K/W]: float
input_leakage: Supply/return leakage: 5/11 integer
input_conditioning: Conditioning "C" / "H/F" / "H/HP" string
input_stories: Number of stories: integer
input_load: Total building load [W]: float

and gives as output power losses[W]
return float
"""
	#We need to get the Fc coefficient (we call it coeff). In order to do that we need to read it from a table

    #it takes the excel file from the folder where is located (we need table 6 from chapter 17 of ashrae)
    ExcelFile = load_workbook("C:\Users\martina\Dropbox\IoT\\table6.xlsx");
    #choose the right sheet of the file
    WindowData = ExcelFile.get_sheet_by_name("Typical_Duct_Loss");
    #read the table from excel starting from stories. It select the starting cell as [1][2:] (that is C2) and goes on on its row.
    stories_cells = WindowData.rows[1][2:];
    #create empty lists. treshold[] will be filled with resistances
    treshold = {};
    
    #We load all the data of the table in a smart dictionary (of subdictionaries of subdictionaries etc) which at the end will contain all the values of Fc.
    # Fc will be then accessed in this way: duct_losses[input_stories][input_leakage][input_insulation][input_duct][input_conditioning]
    duct_losses = {};
    #It gives back 0 if the duct is conditioned space as table suggest
    if (input_duct == "Conditioned space"):
        return 0.0;
    #with for cycle read the value scanning columns
    for cell in stories_cells:
        #find the column index because it needs to create a relation between the variable's indexes 
        column_index = stories_cells.index(cell);
        #starting from stories it reads the values
        stories = int(cell.value)
        leakage = int(WindowData.rows[2][2+column_index].value)
        insulation = float(WindowData.rows[3][2+column_index].value)
        #Build a dictionary which has the resistances considered in the table as keys (useful for subsequent interpolation). 
        #The trick of using a dictionary and considering its keys only is in order to have unique values.
        treshold[insulation] = "word";
        #it reads the rest of the column(these cells contain the values for the duct loss factor)
        col = WindowData.columns[2+column_index][4:];
        #now it focus on these columns and it explores the rows 
        for row in col:
            #find the row index because it needs to create a relation between the variable's indexes 
            row_index = col.index(row);
            #it reads values for duct loction and working condition of the corresponding cells
            duct = WindowData.columns[0][4+row_index].value.encode('utf-8')
            conditioning = WindowData.columns[1][4+row_index].value.encode('utf-8')
            #It creates lists behind list if doesn't already find the read value. 
            #So at the first cycle everything will be empty and it creates a structure of
            #lists, then it fills it in every cycle.
            if (not stories in duct_losses):
                duct_losses[stories] = {}
            if (not leakage in duct_losses[stories]):
                duct_losses[stories][leakage] = {}
            if (not insulation in duct_losses[stories][leakage]):
                duct_losses[stories][leakage][insulation] = {}
            if (not duct in duct_losses[stories][leakage][insulation]):
                duct_losses[stories][leakage][insulation][duct] = {}
            if (not conditioning in duct_losses[stories][leakage][insulation][duct]):
                duct_losses[stories][leakage][insulation][duct][conditioning] = {}
            value = WindowData.columns[2+column_index][4+row_index].value;
            #if duct loss/gain factor is a long type transform it in float
            if (isinstance(value, long )):
                duct_losses[stories][leakage][insulation][duct][conditioning] = float(value)
            else:
                duct_losses[stories][leakage][insulation][duct][conditioning] = float(value.encode('utf-8'))
                
    #in this part it interpolate (restistance-duct losses) if the insert resistance is a value between two listed resistances and gives back duct loss/gain factor
    treshold = treshold.keys()
    
    #if the insert resistance is smaller than the smallest one it select the smallest resistance and gives duct loss/gain factor
    #if the insert resistance is bigger than the biggest one it select the biggest resistance and gives duct loss/gain factor 
    input_insulation = max(treshold[0], min(input_insulation, treshold[len(treshold)-1]))
    
    for i in range(0,len(treshold)-1):
        if(input_insulation <= treshold[i+1] and input_insulation >= treshold[i]):
            left = duct_losses[input_stories][input_leakage][treshold[i]][input_duct][input_conditioning];
            right = duct_losses[input_stories][input_leakage][treshold[i+1]][input_duct][input_conditioning];
            coeff=interpolate( treshold[i], treshold[i+1], left, right, input_insulation)
   
	#in the end compute distribution loss (W)        
    return coeff * input_load;

#validation with the ahrae values should gives 0.13 and 0.27 for ""Attic", 1.4 , 5, "H/F", 1, 1" and ""Attic", 1.4 , 5, "C", 1, 1"
c = distributionLosses("Attic", 1.2 , 5, "H/F", 1, 1);
print c;
d = distributionLosses("Attic", 1.2 , 5, "C", 1, 1);
print d;
