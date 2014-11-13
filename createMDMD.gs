//Constants
NUMBER_HEADER = [1, 1]
NAME_HEADER = [1, 2]
LINEAL_HEADER= [1, 3]

NUMFROZENCOLUMNS = 3
NUMFROZENROWS = 2

  
/**
***************************************************************************************
* createMDMD
* @description  
***************************************************************************************
*/
function createMDMD() {
  //Leer titulo de la hoja nueva
  var app = UiApp.createApplication().setTitle('Create new MDMD');
  
  var stepControl = app.createTextBox()
                .setText('Step 1')
                .setId('stepControl')
                .setName('step')
                .setEnabled(false)
                //.setHorizontalAlignment(UiApp.HorizontalAlignment.CENTER);
  
  //Param GRID
  var paramGrid = app.createGrid(2, 2).setId('paramGrid');
  
  paramGrid.setWidget(0, 0, app.createLabel('New MDMD Name:'));
  paramGrid.setWidget(0, 1, app.createTextBox().setName('name').setId('nameTextBox'));
  
  paramGrid.setWidget(1, 0, app.createLabel('Number of criterias:'));
  paramGrid.setWidget(1, 1, app.createTextBox().setName('optNumber').setId('optNumberTextBox'));
  
  //Criteria GRID
  var critGrid = app.createGrid(0,0).setId('critGrid');
  
  //Compose form
  var handler = app.createServerHandler('createMDMD_StepHandler')
                     .addCallbackElement(stepControl)
                     .addCallbackElement(paramGrid)
                     .addCallbackElement(critGrid);
  var mybutton = app.createButton('Update')
                      .setId('button')
                      .addClickHandler(handler);
  
  var mypanel = app.createVerticalPanel();
  mypanel.add(stepControl);
  mypanel.add(paramGrid);
  var scroll = app.createScrollPanel();
  scroll.add(critGrid);
  mypanel.add(scroll);
  mypanel.add(mybutton);
  app.add(mypanel);
  SpreadsheetApp.getActiveSpreadsheet().show(app);
}


/**
***************************************************************************************
* createMDMD_StepHandler
* @description  
***************************************************************************************
*/
function createMDMD_StepHandler(eventInfo) { 
  //Leer numero de opciones
  var app = UiApp.getActiveApplication();
  var step = eventInfo.parameter.step;
  var stepControl = app.getElementById("stepControl");
  
  if (step == 'Step 1') {
    Logger.log(step);
    stepControl.setText("Step 2")
    
    var paramGrid = app.getElementById("paramGrid")
    var name = eventInfo.parameter.name;
    var nameTextBox = app.getElementById("nameTextBox")
                         .setEnabled(false)
    var optNumber = parseInt(eventInfo.parameter.optNumber)
    var optNumberTextBox = app.getElementById("optNumberTextBox")
                              .setEnabled(false)
    Logger.log("Name: " + name)  
    Logger.log("optNumber: " + optNumber)
    
    //Una a una ir preguntando titulo, max/min y ponderacion
    var critGrid = app.getElementById("critGrid").resize(optNumber + 1, 4);
    var criteriaArray = new Array(new Array());
    
    critGrid.setWidget(0, 0, app.createLabel("#"));
    critGrid.setWidget(0, 1, app.createLabel("Criteria Name"));
    critGrid.setWidget(0, 2, app.createLabel("Max/Min"));
    critGrid.setWidget(0, 3, app.createLabel("Ponderate"));
    
    for (var i = 1; i <= optNumber; i++) {
      critGrid.setWidget(i, 0, app.createLabel(i));
      critGrid.setWidget(i, 1, app.createTextBox().setName('criterianame' + i).setId('criterianameTextBox' + i));
      critGrid.setWidget(i, 2, app.createTextBox().setName('maxmin' + i ).setId('maxmintextbox' + i));
      critGrid.setWidget(i, 3, app.createTextBox().setName('ponderate' + i).setId('ponderatetextbox' + i));
    }
    
    var mybutton = app.getElementById("button")
                      .setText('Do it!');
    
    return app;
  }
  
  if (step == 'Step 2') {
    Logger.log(step);
    stepControl.setText("Done")
  
    var mybutton = app.getElementById("button")
                      .setText('Close');
    
    var name = eventInfo.parameter.name; 
    var optNumber = parseInt(eventInfo.parameter.optNumber);
    
    //Degub
    /*try {
      SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name))
      SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet()
    } catch (err){     }*/
    
    // Make a new sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
    GSUtils.resize(sheet, 2, (optNumber * 2) + 3);
    
    // Make headers
    
    sheet.getRange(NUMBER_HEADER[0], NUMBER_HEADER[1]).setValue("#")
    sheet.getRange(NAME_HEADER[0], NAME_HEADER[1]).setValue("Name")
    sheet.getRange(LINEAL_HEADER[0], LINEAL_HEADER[1]).setValue("Lineal")
    
    // Add criterias
    for (var column = 1; column <= optNumber; column++) {
      var FIRST_COLUMN = 3 - 1 + (2*column)
      var SECOND_COLUMN = FIRST_COLUMN + 1
      
      sheet.getRange(1, FIRST_COLUMN).setValue(eventInfo.parameter['criterianame' + column])
      sheet.getRange(1, SECOND_COLUMN).setValue(eventInfo.parameter['criterianame' + column] + " " + "norm")
      
      sheet.getRange(2, FIRST_COLUMN).setValue(eventInfo.parameter['maxmin' + column])
      sheet.getRange(2, SECOND_COLUMN).setValue(eventInfo.parameter['ponderate' + column])
    }
   
    // Add first option
    GSUtils.addRowAtTheEnd(sheet, "1")
    
    // Freeze for the first time!
    while (sheet.getFrozenRows() != NUMFROZENROWS) sheet.setFrozenRows(NUMFROZENROWS)
    while (sheet.getFrozenColumns() != NUMFROZENCOLUMNS) sheet.setFrozenColumns(NUMFROZENCOLUMNS)
  
    updateMDMD()
    
    return app;
  }
  
  //If nothing else
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}



/**
***************************************************************************************
* updateMDMD
* @description  
***************************************************************************************
*/
function updateMDMD() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var NUMFROZENCOLUMNS = sheet.getFrozenColumns()
  var NUMFROZENROWS = sheet.getFrozenRows()
  var optNumber = (sheet.getLastColumn()-NUMFROZENCOLUMNS)/2;
  var FIRSTROW = NUMFROZENROWS + 1
  var LASTROW = sheet.getLastRow()
  
  //Add linear sum
  var formula = "SUM(                                                                                     \
                     INDIRECT(                                                                            \
                              CONCATENATE(                                                                \
                                          ADDRESS("+FIRSTROW+", "+LINEAL_HEADER[1]+", 1); \
                                          \":\";                                                          \
                                          ADDRESS("+LASTROW+", "+LINEAL_HEADER[1]+", 1)  \
                                         )                                                                \
                             )                                                                            \
                    )"
  sheet.getRange(LINEAL_HEADER[0] + 1, LINEAL_HEADER[1]).setFormula(formula.replace(/\s+/g, ' '))
  
  //Iterate over options
  for (var row = FIRSTROW; row <= LASTROW; row++){
    
    // Add linear
    var formula = "=SUM(0)"
       
    sheet.getRange(row, LINEAL_HEADER[1]).setFormula(formula)
    
    //Iterate over criterias
    for (var column = 1; column <= optNumber; column++) {
      var FIRST_COLUMN = 3 - 1 + (2*column)
      var SECOND_COLUMN = FIRST_COLUMN + 1
      
      var maxminCoords = [2, FIRST_COLUMN];
      var rowdataCurrCoords = [row, FIRST_COLUMN];
      var rowdataFirsCoords = [FIRSTROW, FIRST_COLUMN];
      var rowdataLastCoords = [LASTROW, FIRST_COLUMN];
      var pondCoords = [2, SECOND_COLUMN];
      var rowpondCurrCords = [row, SECOND_COLUMN];
      
      //Update lineal
      var formula = sheet.getRange(row, LINEAL_HEADER[1]).getFormula();
      formula = formula.replace(/\)$/g, "; \
                              INDIRECT(ADDRESS("+pondCoords[0]+", "+pondCoords[1]+")) \
                              * \
                              INDIRECT(ADDRESS("+rowpondCurrCords[0]+", "+rowpondCurrCords[1]+")) \
                             )")
      sheet.getRange(row, LINEAL_HEADER[1]).setFormula(formula.replace(/\s+/g, ' '));
      
      // Formula!
      var formula = "=IF(                                                                               \
                         EQ(                                                                            \
                            INDIRECT(ADDRESS("+maxminCoords[0]+", "+maxminCoords[1]+"));             \
                            \"max\"                                                                     \
                           )                                                                            \
                         ;                                                                              \
                         INDIRECT(ADDRESS("+rowdataCurrCoords[0]+", "+rowdataCurrCoords[1]+"))                  \
                          /                                                                                        \
                          SUM(                                                                                     \
                              INDIRECT(                                                                            \
                                       CONCATENATE(                                                                \
                                                   ADDRESS("+rowdataFirsCoords[0]+", "+rowdataFirsCoords[1]+"); \
                                                   \":\";                                                          \
                                                   ADDRESS("+rowdataLastCoords[0]+", "+rowdataLastCoords[1]+")  \
                                                  )                                                                \
                                      )                                                                            \
                              )                                                                                    \
                         ;                                                                                         \
                          (                                                                                                      \
                           1                                                                                                     \
                           /                                                                                                     \
                           INDIRECT(ADDRESS("+rowdataCurrCoords[0]+", "+rowdataCurrCoords[1]+"))                              \
                          )                                                                                                      \
                          /                                                                                                      \
                          SUM(                                                                                                   \
                              ARRAYFORMULA(                                                                                      \
                                           1                                                                                     \
                                           /                                                                                     \
                                           INDIRECT(                                                                             \
                                                    CONCATENATE(                                                                 \
                                                                ADDRESS("+rowdataFirsCoords[0]+", "+rowdataFirsCoords[1]+");  \
                                                                \":\";                                                           \
                                                                ADDRESS("+rowdataLastCoords[0]+", "+rowdataLastCoords[1]+")   \
                                                               )                                                                 \
                                                   )                                                                             \
                                          )                                                                                      \
                              )                                                                                                  \
                        )"
                       
                        
      sheet.getRange(row, SECOND_COLUMN).setFormula(formula.replace(/\s+/g, ' '))
    }
    
  }
    
  //Prettyprint
  prettyprint()
}



/**
***************************************************************************************
* prettyprint
* @description  
***************************************************************************************
*/
function prettyprint() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (sheet != null) {
    
    var NUMFROZENCOLUMNS = sheet.getFrozenColumns()
    var NUMFROZENROWS = sheet.getFrozenRows()
    var FIRSTROW = NUMFROZENROWS + 1
    var LASTROW = sheet.getLastRow() 
    var LASTCOL = sheet.getLastColumn()
    
    //Colorear las columnas de dash con negro
    for (var column = 1; column <=sheet.getMaxColumns(); column ++) {
      sheet.getRange(1, column, 1, 1).setBackgroundColor("darkblue")
      .setFontColor("white")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
      
      sheet.getRange(2, column, 1, 1).setBackgroundColor("CornflowerBlue")
      .setFontColor("white")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    }
    
    for (var column = NUMFROZENCOLUMNS; column <=sheet.getMaxColumns(); column = column + 2) {
      sheet.getRange(FIRSTROW, column, LASTROW - FIRSTROW, 1).setNumberFormat("0.00");
    }
  }
}   