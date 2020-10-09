from pyomo.environ import *
import os
import math
from openpyxl import *

wb=load_workbook(filename='14A.V-Mobile_CLASS.xlsx')
ws=wb["Data"]

model=AbstractModel()

#Size
model.nc=Param(initialize=3)#Number of carriers
model.nd=Param(initialize=5)#Destinations
model.ni=Param(initialize=3)#Price intervals
model.nt=Param(initialize=2)#Months

#Sets
model.c=RangeSet(1, model.nc)#Carriers
model.d=RangeSet(1, model.nd)#Destinations
model.i=RangeSet(1, model.ni)#Price intervals
model.t=RangeSet(1, model.nt)#Months

#Data
#Price
def p_init(model, c,d,i,t):
    return (ws.cell(17+(3*(c-1))+i, 1+(5*(t-1))+d).value)
model.p=Param(model.c, model.d, model.i ,model.t, initialize=p_init)

#Penalty
def pen_init(model, c,d,t):
    return (ws.cell(30+c, 3*(t-1)+m+1).value)
model.pen=Param(model.c, model.d, model.t, initialize=pen_init)



