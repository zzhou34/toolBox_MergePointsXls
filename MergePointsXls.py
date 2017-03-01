import arcpy, xlrd
arcpy.env.workspace = arcpy.GetParameter(0)
workspace=arcpy.GetParameter(0)+'\\'
spatialRef="GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]];-400 -400 1000000000;-100000 10000;-100000 10000;8.98315284119522E-09;0.001;0.001;IsHighPrecision"

coordinate_list = arcpy.ListFiles('*.xls')
for coordinate in coordinate_list:
  name=coordinate[:-4]
  excel=xlrd.open_workbook(workspace+coordinate)
  table=excel.sheet_by_name('coordinates')
  big_list=[]
  nrows=table.nrows
  for i in range(1,nrows):
    list=[]
    list.append(table.cell(i,0).value)
    list.append(table.cell(i,1).value)
    big_list.append(list)
  pointGeometryList=[]
  pnt=arcpy.Point()
  for coord in big_list:
    pnt.X=float(coord[1])
    pnt.Y=float(coord[0])
    pointGeometry=arcpy.PointGeometry(pnt,spatialRef)
    pointGeometryList.append(pointGeometry)
    #print (len(pointGeometryList))
  arcpy.CopyFeatures_management(pointGeometryList,workspace+name+'.shp')

coordinate_list = arcpy.ListFiles('*.xlsx')
for coordinate in coordinate_list:
  name = coordinate[:-5]
  excel = xlrd.open_workbook(workspace + coordinate)
  table = excel.sheet_by_name(name)
  big_list = []
  nrows = table.nrows
  for i in range(2, 100):
    list = []
    list.append(table.cell(i, 6).value)
    list.append(table.cell(i, 7).value)
    big_list.append(list)
  pointGeometryList = []
  pnt = arcpy.Point()
  for coord in big_list:
    pnt.X = float(coord[1])
    pnt.Y = float(coord[0])
    pointGeometry = arcpy.PointGeometry(pnt, spatialRef)
    pointGeometryList.append(pointGeometry)
  # print (len(pointGeometryList))
  arcpy.CopyFeatures_management(pointGeometryList, workspace + name + '.shp')

shp_list=arcpy.ListFiles('*.shp')
arcpy.Merge_management(inputs=shp_list,output=workspace+'all_dots.shp')
