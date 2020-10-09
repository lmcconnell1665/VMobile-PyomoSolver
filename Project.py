from pyomo.environ import *
import os
import math
from openpyxl import *

wb=load_workbook(filename='14A.V-Mobile_CLASS.xlsx')
ws=wb["Data"]

model=AbstractModel()

#Size
model.nc=Param(initialize=3)
model.nd=Param(initialize=5)
model.ni=Param(initialize=3)
model.nt=Param(initialize=2)

#Sets
model.c=RangeSet(1, model.nc)
model.d=RangeSet(1, model.nd)
model.i=RangeSet(1, model.ni)
model.t=RangeSet(1, model.nt)

#Param

