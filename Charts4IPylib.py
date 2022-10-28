import sys
import clr
import json
from System import AppDomain
from Spotfire.Dxp.Data import *

###################################################################################
# Use the following script elements to include this library 
#
#import sys
#import __builtin__
#from System.IO import Path
#sys.path.append(Path.Combine(Path.GetTempPath(),'ScriptSync','ConfigureCharts4'))
#__builtin__.Document = Document
#__builtin__.Application = Application
#import Charts4IPylib
#
###################################################################################

loadAssemblies = ['Charts', 'Common', 'Publisher', 'ChemistryService']
for asm in AppDomain.CurrentDomain.GetAssemblies():
    if asm.GetName().Name in loadAssemblies:
        clr.AddReference(asm.FullName)
    elif asm.GetName().Name == 'Newtonsoft.Json' and asm.GetName().Version.Major == 12:
        clr.AddReference(asm.FullName)

from Common import ColumnIdService
from Charts import ChartsModel
from Newtonsoft.Json import JsonConvert
from Newtonsoft.Json.Linq import JObject
from Publisher import PublisherValueRendererSettings
from ChemistryService import ChemistryService, ChemistryServiceFxn


def ColumnNamesToIds(visual, columnNames):
    idService = Application.GetService[ColumnIdService]()
    columnIds = []
    for columnName in columnNames:
        if visual.DataTable.Columns.Contains(columnName):
            column = visual.DataTable.Columns[columnName]
            columnIds.append(idService.GetID(column))
    return columnIds


def SetTableColumns(visual, columnNames):
    idService = Application.GetService[ColumnIdService]()
    columns = []
    columnIds = []
    for columnName in columnNames:
        if visual.DataTable.Columns.Contains(columnName):
            column = visual.DataTable.Columns[columnName]
            columns.append(column)
            columnIds.append(idService.GetID(column))

    removedIds = []
    for column in visual.DataTable.Columns:
        if column not in columns:
            removedIds.append(idService.GetID(column))

    #unfixed columns
    tableId = visual.DataTable.Id.ToString()
    unfixedKey = tableId+".table-visualization.table.unfixedColumnIds"
    visual.SetKeyValue(unfixedKey,json.dumps(columnIds))
    removedKey = tableId+".table-visualization.table.removedColumnIds"
    visual.SetKeyValue(removedKey,json.dumps(removedIds))


def SetTableTranspose(visual, transposed):
    tableId = visual.DataTable.Id.ToString()
    key = tableId+".table-visualization.table.transposed"
    visual.SetKeyValue(key,transposed.ToString().ToLower())


def FixTableRows(visual, rowIdxs):
    tableId = visual.DataTable.Id.ToString()
    key = tableId+".table-visualization.table.fixedRowIdxs"
    visual.SetKeyValue(key, json.dumps(rowIdxs))


def ClearFixedTableRows(visual):
    tableId = visual.DataTable.Id.ToString()
    key = tableId+".table-visualization.table.fixedRowIdxs"
    visual.SetKeyValue(key, "[]")


def FixTableColumns(visual, columnNames):
    tableId = visual.DataTable.Id.ToString()
    fixedKey = tableId+".table-visualization.table.fixedColumnIds"
    unfixedKey = tableId+".table-visualization.table.unfixedColumnIds"
    columnIds = ColumnNamesToIds(visual, columnNames)
    if visual.ContainsKey(unfixedKey):
        unfixed = json.loads(visual.GetKeyValue(unfixedKey))
        unfixed = [x for x in unfixed if x not in columnIds]
        visual.SetKeyValue(unfixedKey, json.dumps(unfixed))    
    visual.SetKeyValue(fixedKey, json.dumps(columnIds))


def ClearFixedTableColumns(visual):
    tableId = visual.DataTable.Id.ToString()
    fixedKey = tableId+".table-visualization.table.fixedColumnIds"
    unfixedKey = tableId+".table-visualization.table.unfixedColumnIds"
    if visual.ContainsKey(fixedKey) and visual.ContainsKey(unfixedKey):
        fixed = json.loads(visual.GetKeyValue(fixedKey))
        unfixed = json.loads(visual.GetKeyValue(unfixedKey))
        [fixed.append(x) for x in unfixed if x not in fixed]
        visual.SetKeyValue(unfixedKey,json.dumps(fixed))

    visual.SetKeyValue(fixedKey, None)


def SetTableColumnWidth(visual, columnName, width, tposeHeight):
    idService = Application.GetService[ColumnIdService]()
    columnId = None
    tableId = visual.DataTable.Id.ToString()
    dimKey = tableId+".table-visualization.table.columnDimensions"
    columnDim = {"width":width,"transposedHeight":tposeHeight}

    if visual.DataTable.Columns.Contains(columnName):
        column = visual.DataTable.Columns[columnName]
        columnId = idService.GetID(column)
        if visual.ContainsKey(dimKey):
            dims = json.loads(visual.GetKeyValue(dimKey))
            dims[columnId] = columnDim
        else:
            dims = {columnId: columnDim}
        
        visual.SetKeyValue(dimKey,json.dumps(dims))


def SetTableRowHeight(visual, height):
    tableId = visual.DataTable.Id.ToString()
    key = tableId+".table-visualization.table.rowHeight"
    visual.SetKeyValue(key, str(height))
    

def SetTableTransposeColumnWidth(visual, width):
    tableId = visual.DataTable.Id.ToString()
    key = tableId+".table-visualization.table.transposedColumnWidth"
    visual.SetKeyValue(key, str(width))


def ClearRadarColumns(visual):
    key = visual.DataTable.Id.ToString()+'.radar-visualization.radar.axes'
    visual.SetKeyValue(key,'[]')


def AddRadarColumn(visual, columnName, logScale=None, inverted=None, minVal=None, maxVal=None):
    key = visual.DataTable.Id.ToString()+'.radar-visualization.radar.axes'
    if visual.ContainsKey(key):
        axesJson = json.loads(visual.GetKeyValue(key))
    else:
        axesJson = []
    idService = Application.GetService[ColumnIdService]()
    if visual.DataTable.Columns.Contains(columnName):
        column = visual.DataTable.Columns[columnName]
        columnId = idService.GetID(column)
        if columnId:
            axisJson = {'columnId':columnId}
            axisJson['logScale'] = True if logScale else False
            axisJson['inverted'] = True if inverted else False
            if (minVal is not None) and (maxVal is not None):
                axisJson['min'] = minVal
                axisJson['max'] = maxVal
            axesJson.append(axisJson)
            visual.SetKeyValue(key, json.dumps(axesJson))


def ClearMPOColumns(visual):
    key = visual.DataTable.Id.ToString()+'.mpo-visualization.mpo.axes'
    visual.SetKeyValue(key,'[]')


def AddMPOColumn(visual, columnName, logScale=None, inverted=None, minVal=None, maxVal=None):
    key = visual.DataTable.Id.ToString()+'.mpo-visualization.mpo.axes'
    if visual.ContainsKey(key):
        axesJson = json.loads(visual.GetKeyValue(key))
    else:
        axesJson = []
    idService = Application.GetService[ColumnIdService]()
    if visual.DataTable.Columns.Contains(columnName):
        column = visual.DataTable.Columns[columnName]
        columnId = idService.GetID(column)
        if columnId:
            axisJson = {'columnId':columnId}
            axisJson['logScale'] = True if logScale else False
            axisJson['inverted'] = True if inverted else False
            if (minVal is not None) and (maxVal is not None):
                axisJson['min'] = minVal
                axisJson['max'] = maxVal
            axesJson.append(axisJson)
            visual.SetKeyValue(key, json.dumps(axesJson))


def SetSortColumn(visual, columnName, order):
    #order = "asc" or "desc"
    idService = Application.GetService[ColumnIdService]()
    if visual.DataTable.Columns.Contains(columnName):
        column = visual.DataTable.Columns[columnName]
        columnId = idService.GetID(column)
        sortJson = {'id':columnId,'order':order}
        jobj = JsonConvert.DeserializeObject(json.dumps(sortJson))
        visual.SetSortCriteria(jobj)


def SetWhereClause(visual, expression):
    visual.WhereClauseExpression = expression


def SetDataTable(visual, tableName):
    dataMgr = Application.GetService[DataManager]()
    if dataMgr.Tables.Contains(tableName):
        visual.DataTable = dataMgr.Tables[tableName]


def SetMarking(visual, markingName):
    dataMgr = Application.GetService[DataManager]()
    if dataMgr.Markings.Contains(markingName):
        visual.Marking = dataMgr.Markings[markingName]


def SetRenderer(visual, columnName, rendererName):
    if visual.DataTable.Columns.Contains(columnName):
        column = visual.DataTable.Columns[columnName]
        visual.SetColumnRenderer(column, rendererName)


def SetColoring(visual, coloring):
    jobj = JsonConvert.DeserializeObject(coloring)
    visual.Coloring.FromJson(jobj)


def SetRendererSettings(visual, columnName, key, value):
    if visual.DataTable.Columns.Contains(columnName):
        column = visual.DataTable.Columns[columnName]
        settings = visual.GetColumnRendererSettings(column)
        if settings and settings.GetType() == clr.GetClrType(PublisherValueRendererSettings):
            settings.SetValue(key, value)
            return True
    return False


def SetRendererSettings(visual, columnName, jsonObject):
    if visual.DataTable.Columns.Contains(columnName):
        column = visual.DataTable.Columns[columnName]
        settings = visual.GetColumnRendererSettings(column)
        if settings and settings.GetType() == clr.GetClrType(PublisherValueRendererSettings):
            settings.CurrentSettings = json.dumps(jsonObject)
            return True
    return False


def GetRendererSettings(visual, columnName):
    column = visual.DataTable.Columns[columnName]
    settings = visual.GetColumnRendererSettings(column)
    if settings and settings.GetType() == clr.GetClrType(PublisherValueRendererSettings):
        return settings.CurrentSettings


def CreateColumnPropertyIfNeeded(propertyName):
    dataMgr = Application.GetService[DataManager]()
    if not dataMgr.Properties.ContainsProperty(DataPropertyClass.Column, propertyName):
        attributes = DataPropertyAttributes.IsPersistent |DataPropertyAttributes.IsPropagated |DataPropertyAttributes.IsEditable |DataPropertyAttributes.IsSearchable | DataPropertyAttributes.IsVisible         
        dataProperty = DataProperty.CreateCustomPrototype(propertyName, DataType.String, attributes)
        dataMgr.Properties.AddProperty(DataPropertyClass.Column, dataProperty)
        
        
def SetDepictionTemplates(templates):
    dataMgr = Application.Document.Data
    propertyName = 'LeadDiscoveryChemCharts.DepictionTemplates'
    if not dataMgr.Properties.ContainsProperty(DataPropertyClass.Document, propertyName):
        dataProperty = DataProperty.CreateCustomPrototype(propertyName, None, DataType.String, DataProperty.DefaultAttributes)
        dataMgr.Properties.AddProperty(DataPropertyClass.Document, dataProperty);                                                                                    

    if dataMgr.Properties.ContainsProperty(DataPropertyClass.Document, propertyName):
        Application.Document.Properties[propertyName] =	templates
        
        
def AddStructureSearch(tableName, columnName, searchType, query, options, resultName = 'Structure search'):
    if Application.Document.Data.Tables.Contains(tableName):
        dataTable = Application.Document.Data.Tables[tableName]
        if dataTable and dataTable.Columns.Contains(columnName):
            idSvc = Application.GetService[ColumnIdService]()
            chemSvc = Application.GetService[ChemistryService]()
            column = dataTable.Columns[columnName]
            columnId = idSvc.GetID(column)            
            chemTk = chemSvc.DefaultTkFor(ChemistryServiceFxn.structureSearch)
            request = {'column': columnId, 'options': options, 'query':query, 'resultName':resultName, 'searchType': searchType}
            result = chemTk.StructureSearch(dataTable, JsonConvert.DeserializeObject(json.dumps(request)))
           
           
def AddTablePlot(page, dataTable, transpose):
    tableVis = page.Visuals.AddNew[ChartsModel]()
    tableVis.SetKeyValue('visualization','table-visualization')
    
    if transpose:
        key = dataTable.Id.ToString()+".table-visualization.table.transposed"
        tableVis.SetKeyValue(key, 'true')
        
    tableVis.DataTable = dataTable
    tableVis.ConfigureColumns()
    tableVis.Marking = Document.Data.Markings.DefaultMarkingReference
    tableVis.SetActiveVisual()
    #tableVis.SetColumnRenderer(dataTable.Columns['Analog'], 'RDKit')
    return tableVis
    
    
