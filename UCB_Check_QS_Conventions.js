// ****************************************************************************************************
// *
// *   UCB_Check_QS_Conventions
// *   Conventions check report 
// *   Check of model, object and connection attributes, process flow, process environment
// *   (c) Bayrische Hypo- und Vereinsbank AG, ein Unternehmen der Unicredit Group
// *
// ****************************************************************************************************
// Version 1.0.0 (2015-06-07 | Kerstin Drescher + Steffen Ploetz)
// Initial version

// Version 1.0.1 (2015-07-22 | Kerstin Drescher + Steffen Ploetz)
// rework of logging and changes in some checks

// Version 1.1 (2015-08-27 | Kerstin Drescher)
// rework and changes in some check, correction of errors

// Version 1.1.1 (2015-09-01 | Stephan Maget)
// rework and changes in some check, correction of errors

// Version 1.1.2 (2015-09-23 | Stephan Maget)
// rework and changes in some check, correction of errors

// Version 1.1.3 (2015-10-19 | Stephan Maget)
// rework and changes in some check, correction of errors

// Version 1.1.4 (2020-05-03 | Stephan Maget)
// bug fixed in uniques check (CheckUniqueOccurence)

// Version 1.1.5 (2020-04-29 | Pieter Jacobs)
// additional 'LAW262EPCAttributes' rules

// Version 1.1.6 (2020-05-14 | Pieter Jacobs)
// additional 'LAW262AssignedBCD' rules

// Version 1.1.7 (2020-06-05 | Pieter Jacobs)
// additional 'LAW262Object' rules

// Version 1.1.8 (2020-06-25 | Pieter Jacobs)
// several small bug fixes

// Version 1.1.9 (2020-06-29 | Steffen Ploetz)
// several missing rules (7) added and mis-configured rueles (4) updated

// Version 1.2.0 (2020-07-21 | Steffen Ploetz)
// several bug fixes: getAttrNum() function calls; IsEqualConnectedSymbol/IsEqualConnectedSymbolWithValue/IsEqualContainingBCDWithValue processing

// Version 1.2.1 (2020-07-21 | Steffen Ploetz)
// moved attribute compare for "AssignedBCDWithValue"/"IsEqual"/"CheckValue" from individual code to checkAttrDependence()

// Version 1.2.2 (2020-10-02 | Steffen Ploetz)
// additional attribute compare "IsEqualConnectedSymbolWithDiagramValue"

// Global declarations
var g_reportLcid     = 0;
var g_selectedModels = null;
var g_methodFilter   = null;
var g_outfile        = null;
var g_debugMode      = true;

//prepare logging
var g_currentLogType = 6;     //Logger.LOG_INFO          
var g_countErrors    = 0; 
var g_countSuccessfulModels = 0; 
var g_loggerClasses         = new Array();
var g_fileTransferFolder    = ""; 


//default fonts for general output
var g_stndFont = null;
var g_boldFont = null; 

var g_xlSheet = null; 
var g_xlWorkbook = null; 
var g_countColumns = 7; 
var g_rowNumber=0;
var g_cs = null; 
var g_hlinkStyle = null; 

var g_config = null; 

//attr configuration
var g_attrModelLanguage = "1ee748d0-6eb7-11e1-44a6-00300571cf1f"; 
var g_usedLanguage = 0; 

var g_currentModel = null; 

__usertype_logClass = function()
{
    this.logType = 0; 
    this.logTypeName = ""; 
    this.cellStyle = null; 
}

var g_countErrors   = 0; 
var g_countWarnings = 0; 
var g_countCritical = 0; 
var g_countNotice   = 0; 

var g_objsToCheck = new Array(); 
/// <summary> Caller for main flow control. </summary>
/// <returns> NOT EVALUATED BY CALLER. </returns>

var isMacro = Context.getProperty("macroRun");
if(!isMacro){
    isMacro = false;
}

var g_sheetCount = 0;

main();

/// <summary> Main flow control. Creates a loop through all selected models to create one narrative output document per model. </summary>
/// <returns> NOT EVALUATED BY CALLER. </returns>
function main()
{
    g_xlSheet = null; 
    g_xlWorkbook = null; 
    g_usedLanguage = Context.getSelectedLanguage();
    
    if (g_usedLanguage == 0) 
        return ""; 
   var fileName = createFileName("UCB Check QA Conventions"); 
   g_xlWorkbook = createExcelOutputFile(fileName); 
   
    var result = null;

    g_reportLcid = Context.getSelectedLanguage();

    var loggerClassError = new __usertype_logClass(); 
    loggerClassError.logType = 1; 
    loggerClassError.logTypeName = "Kritisch"; 
    g_loggerClasses.push(loggerClassError); 
    
    var loggerClassWarning = new __usertype_logClass(); 
    loggerClassWarning.logType = 2; 
    loggerClassWarning.logTypeName = "Fehler"; 
    g_loggerClasses.push(loggerClassWarning); 
    
    var loggerClassInfo = new __usertype_logClass(); 
    loggerClassInfo.logType = 3; 
    loggerClassInfo.logTypeName = "Warnung"; 
    g_loggerClasses.push(loggerClassInfo); 
    
    var loggerClassDebug = new __usertype_logClass(); 
    loggerClassDebug.logType = 4; 
    loggerClassDebug.logTypeName = "Notiz"; 
    g_loggerClasses.push(loggerClassDebug); 
    
    var loggerClassDebug = new __usertype_logClass(); 
    loggerClassDebug.logType = 5; 
    loggerClassDebug.logTypeName = "Info"; 
    g_loggerClasses.push(loggerClassDebug); 
    
    var loggerClassDebug = new __usertype_logClass(); 
    loggerClassDebug.logType = 6; 
    loggerClassDebug.logTypeName = "Detail"; 
    g_loggerClasses.push(loggerClassDebug); 

    // Test calling environment.
    if (g_reportLcid != 1031 && g_reportLcid != 1033)
    {
        Dialogs.MsgBox("Derzeit werden nur \'Deutsch(Deutschland)\' und \'Englich(USA)\' unterstützt. /\r\n" +
                       "Currently only \'German (Germany)\' and \'English (USA)\' are supported.", vbOKOnly,
                       "Ungültige Sprachauswahl / Invalid language selection");
        Context.setScriptError(Constants.ERR_CANCEL);
        return;
    }

    // Get Filter.
    g_methodFilter = ArisData.getActiveDatabase().ActiveFilter();

    // Get and test all selected models.
    g_selectedModels = ArisData.getSelectedModels();
    if (g_selectedModels.length < 1)
    {
        Dialogs.MsgBox("Es wurde kein Modell zur Auswertung ausgewähle. /\r\n" +
                       "No model was selected to be evaluated.", vbOKOnly,
                       "Ungültige Modellauswahl / Invalid model selection");
        Context.setScriptError(Constants.ERR_CANCEL);
        return;
    }
    
	g_fileTransferFolder = Common_GetServerToClientTransferFolder(Context.getSelectedFile());

    var totals = [];
    g_config = new xmlConfig("TEST_conventionsCheck_config.xml"); 
    var rootNode = g_config.readConfigXml(); 
    if (rootNode != null)
    {
        //pj1 - ignore dialogs if run from macro
        if(!isMacro)
            g_currentLogType = Dialogs.showDialog(new levelSelectionDialog(), Constants.DIALOG_TYPE_ACTION, getString("SELECT_ERROR_LEVEL")); 
        else
            g_currentLogType = 4;

        if (g_currentLogType != -1 )
        {
            var countModels = 0;
            var currentModel = null;
            var countSuccessful = 0;
            var lastTargetFileName = "";
            
           // g_outfile = Context.createOutputObject(Context.getSelectedFormat(), Context.getSelectedFile());
            // g_outfile = Context.createOutputObject();
            // g_outfile.OutputLn("Checked models: ", "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_BOLD, 0); 
            for (var couM=0; couM<g_selectedModels.length; couM++)
            {
                g_countErrors   = 0;
                g_countWarnings = 0;
                g_countCritical = 0;
                g_countNotice   = 0;
                // g_outfile.OutputLn(g_selectedModels[couM].Name(g_reportLcid), "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_BOLD, 0); 
                doConventionsCheck(g_selectedModels[couM])// if (doConventionsCheck(g_selectedModels[couM])== "")
                    // g_outfile.OutputLn(" - " + getString("TXT_NOLANGUAGESELECTED"), "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_BOLD, 0);
                totals.push([g_selectedModels[couM], g_countErrors, g_countWarnings, g_countCritical, g_countNotice]);
            }
			// g_outfile.WriteReport()
        }
    }   
    
    createResultSheet(totals); 
   //show result (number of errors, warnings, etc.) in the logfile of the report (txt)
   // g_outfile.OutputLn("Result: ", "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_BOLD, 0); 
   // 
   // g_outfile.OutputLn("\tNumber of critical errors: " + String(g_countCritical), "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_BOLD, 0); 
   // g_outfile.OutputLn("\tNumber of errors: " + String(g_countErrors), "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_BOLD, 0); 
   // g_outfile.OutputLn("\tNumber of warnings: " + String(g_countWarnings), "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_BOLD, 0); 
   // g_outfile.OutputLn("\tNumber of notes: " + String(g_countNotice), "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_BOLD, 0); 
   // g_outfile.OutputLn("", "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_BOLD, 0);     
                           
   var out = new java.io.FileOutputStream(g_fileTransferFolder + "\\" + fileName); 
   g_xlWorkbook.write(out);
   out.close(); 
   Context.addOutputFileName(fileName); 
   Context.setSelectedFile (fileName) 
}

function getUsedLanguage(model)
{
    var oFilter = model.Database().ActiveFilter();
    
    var atMaintainedLanguageNum = oFilter.UserDefinedAttributeTypeNum(g_attrModelLanguage);
    
    var nLcid =0; 
    
    var attrLang = model.Attribute(atMaintainedLanguageNum, g_reportLcid);
    if (attrLang != null && attrLang.IsMaintained())
    {
        if (attrLang.getValue() == "1ee748d1-6eb7-11e1-44a6-00300571cf1f" || attrLang.getValue() == "Deutsch" || attrLang.getValue() == "german") 
            nLcid = 1031;
        else
            nLcid = 1033;
    }

    if (nLcid == 0)
        nLcid = Dialogs.showDialog(new languageSelectionDialog(model.Name(g_reportLcid)), Constants.DIALOG_TYPE_ACTION, getString("SELECT_MODEL_LANGUAGE")); 

    return nLcid; 
}

function levelSelectionDialog()
{
    var logLevelSelected = -1; 
    this.getPages = function()
    {
        var iDialogTemplate = Dialogs.createNewDialogTemplate(600, 200, getString("SELECT_ERROR_LEVEL"));
        iDialogTemplate.Text(10, 30, 200, 40, getString("TXT_CHECKDETAILLEVEL")); 
        iDialogTemplate.DropListBox(220, 30, 130, 50,["6 - Detail", "5 - Info", "4 - Notiz", "3 - Warnung", "2 - Fehler", "1 - Kritisch"], "DRP_DETAILLEVEL",0); 
        iDialogTemplate.Text(30, 80, 470, 250, getString("TXT_DESCRIPTIONDETAILLEVEL")); 
        return [iDialogTemplate]; 
    }
    this.init = function(aPages)
    {
        aPages[0].getDialogElement("DRP_DETAILLEVEL").setSelection(2); // [SM]: auf Warnung default umgestellt 
    }
    this.onClose = function(pageNumber, bOk)
    {
        
        if (bOk == false)
            logLevelSelected= -1;
        else
        {
            switch(this.dialog.getPage(0).getDialogElement("DRP_DETAILLEVEL").getSelectedIndex())
            {
                case 0:
                logLevelSelected= 6; 
                break;
                case 1: 
                logLevelSelected= 5;
                break;
                case 2: 
                logLevelSelected = 4; 
                break;
                case 3: 
                logLevelSelected =  3; 
                break;
                case 4: 
                logLevelSelected = 2; 
                break;
                case 5: 
                logLevelSelected = 1;
                break;
                default: 
                logLevelSelected = -1;
            }
        }
    }
    // returns true if the page is in a valid state. In this case OK, Finish, or Next is enabled.
    // called each time a dialog value is changed by the user (button pressed, list selection, text field value, table entry, radio button,...)
    // pageNumber: the current page number, 0-based
    this.isInValidState = function(pageNumber)
    {
        return true;
    }
    
    // returns true if the "Finish" or "Ok" button should be visible on this page.
    // pageNumber: the current page number, 0-based
    // optional. if not present: always true
    this.canFinish = function(pageNumber)
    {
        return true;
    }
    
    this.getResult = function()
    {
        return logLevelSelected; 
    }
}

function languageSelectionDialog(modelName)
{
    var selectedLanguage = 0; 
    this.getPages = function()
    {
        var iDialogTemplate = Dialogs.createNewDialogTemplate(200, 150, getString("SELECT_MODEL_LANGUAGE"));
        iDialogTemplate.Text(30, 30, 400, 70, GetText("TXT_MODELLANGUAGEINFO", new Array(modelName))); 
        iDialogTemplate.DropListBox(30, 100, 200, 40,[getString("GERMAN"), getString("ENGLISH")], "DRP_Languages",0); 
        return [iDialogTemplate]; 
    }
    
    this.onClose = function(pageNumber, bOk)
    {
        if (bOk == false)
            selectedLanguage  = 0; 
        else
        {
             if (this.dialog.getPage(0).getDialogElement("DRP_Languages").getSelectedIndex() ==0)
                selectedLanguage =  1031; 
            else
                selectedLanguage = 1033; 
        }
    }
    // returns true if the page is in a valid state. In this case OK, Finish, or Next is enabled.
    // called each time a dialog value is changed by the user (button pressed, list selection, text field value, table entry, radio button,...)
    // pageNumber: the current page number, 0-based
    this.isInValidState = function(pageNumber)
    {
        return true;
    }

    // returns true if the "Finish" or "Ok" button should be visible on this page.
    // pageNumber: the current page number, 0-based
    // optional. if not present: always true
    this.canFinish = function(pageNumber)
    {
        return true;
    }
 
    this.getResult = function()
    {
        return selectedLanguage; 
    }
}

// The 'currentModel' is one of the selected models - filtered by report context to EPC.
function doConventionsCheck(currentModel, xmlConfigRootNode)
{
    g_xlSheet = g_xlWorkbook.createSheet(); 
    
    prepareWorksheet(g_xlSheet, g_xlWorkbook, 1, currentModel);
    
    g_currentModel = currentModel; 
    
    var isLaw262Relevant = g_currentModel.Attribute(g_methodFilter.UserDefinedAttributeTypeNum("09f6c370-6f4c-11e1-44a6-00300571cf1f"), g_usedLanguage).getValue();
    var isLaw262Relevant = (isLaw262Relevant=="Ja"||isLaw262Relevant=="Yes")?true:false;
    
    //if(!isLaw262Relevant){
        LogMessage("debug", "", "", "epc checks", "", 6); 
        var epcAttributeRules = g_config.getRulesOfRuleKind("EPCAttributes"); 
        if (epcAttributeRules != null)
            doConventionsCheckEPCAttributes(epcAttributeRules, currentModel);
    //}
    
    if(isLaw262Relevant){
        LogMessage("debug", "", "", "law 262 epc checks", "", 6); 
        var epcAttributeRules = g_config.getRulesOfRuleKind("LAW262EPCAttributes"); 
        if (epcAttributeRules != null)
            doConventionsCheckEPCAttributes(epcAttributeRules, currentModel);
    }
    
    //if(!isLaw262Relevant){
        LogMessage("debug", "", "", "check epc structure", "", 6); 
        var epcStructureRules = g_config.getRulesOfRuleKind("EPCStructure"); 
        if (epcStructureRules != null)
            doConventionsCheckEPCStructure(epcStructureRules, currentModel); 
    //}
    
    //if(!isLaw262Relevant){
        LogMessage("debug", "", "", "check model periphery", "", 6); 
        var modelPeripheryStructureRules = g_config.getRulesOfRuleKind("ModelPeriphery"); 
        if (modelPeripheryStructureRules != null)
            doConventionsCheckModelPeriphery(modelPeripheryStructureRules, currentModel); 
    //}
    
    //if(!isLaw262Relevant){
        LogMessage("debug", "", "", "check bcd attributes", "", 6); 
        var bcdAttributeRules = g_config.getRulesOfRuleKind("AssignedBCD"); 
        if (bcdAttributeRules !=  null)
            doConventionsCheckBCDModels(bcdAttributeRules, currentModel); 
    //}
    
    if(isLaw262Relevant){
        LogMessage("debug", "", "", "check law 262 bcd attributes", "", 6); 
        var bcdAttributeRules = g_config.getRulesOfRuleKind("LAW262AssignedBCD"); 
        if (bcdAttributeRules !=  null)
            doConventionsCheckBCDModels(bcdAttributeRules, currentModel); 
    }
    
    //if(!isLaw262Relevant){
        LogMessage("debug", "", "", "check object attributes", "", 6); 
        var objAttrRules = g_config.getRulesOfRuleKind("Object"); 
        if (objAttrRules != null)
            doConventionsCheckObjectAttributes(objAttrRules, currentModel, "Object"); 
    //}
    
    if(isLaw262Relevant){
        LogMessage("debug", "", "", "check law 262 object attributes", "", 6); 
        var objAttrRules = g_config.getRulesOfRuleKind("LAW262Object"); 
        if (objAttrRules != null)
            doConventionsCheckObjectAttributes(objAttrRules, currentModel, "LAW262Object"); 
    }
    
    //if(!isLaw262Relevant){
        LogMessage("debug", "", "", "check object occurences", "", 6); 
        var objOccRules = g_config.getRulesOfRuleKind("ObjectEnvironment"); 
        if (objOccRules != null)
            doConventionsCheckObjectOccurences(objOccRules, currentModel); 
    //}

    //if(!isLaw262Relevant){
        LogMessage("debug", "", "", "check connections", "", 6); 
        var cxnRules = g_config.getRulesOfRuleKind("Cxns"); 
        if (cxnRules != null)
            doConventionsCheckConnections(cxnRules, currentModel); 
    //}
    
   return;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///  Conventions check for EPCAttributes
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//check all rules with ruleKind = "EPCAttributes"
function doConventionsCheckEPCAttributes(epcRules, currentModel)
{

   for (var cou=0; cou<epcRules.size(); cou++)
   {
        var ruleNode     = epcRules.get(cou);
        var ruleShortcut = g_config.getRuleShortcut(ruleNode); // SP: Debug helper.
        var attrRule     = g_config.getAttrRuleNode(ruleNode);
        doAttributeCheck(currentModel, ruleNode, attrRule, null); 
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///  Conventions check for attributes of the objects which have occurences in the selected model
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//check all rules with ruleKind = "Object"
function doConventionsCheckObjectAttributes(objAttrRules, currentModel, ruleGroupName)
{
   for (var cou=0; cou<objAttrRules.size(); cou++)
   {
        
        var ruleNode = objAttrRules.get(cou); 
        
        var attrRule = g_config.getAttrRuleNode(ruleNode); 
        var symbolTypeList = new Array(); 
        symbolTypeList = ruleNode.getAttributeValue("typelist").split(",");
        var symbolTypeNums = new Array(); 
        for (var couS=0; couS<symbolTypeList.length; couS++)
        {
            var symbolTypeNum = 0; 
            if (vbIsNumeric(symbolTypeList[couS]))
                symbolTypeNum = symbolTypeList[couS]; 
            else if ( vbLen(symbolTypeList[couS] == 36))
                symbolTypeNum = g_methodFilter.UserDefinedSymbolTypeNum(symbolTypeList[couS]);

            if (symbolTypeNum == 0)
                LogMessage(getLogTypeForAttributeItemKind(currentModel), "", "Alle Regeln '" + ruleGroupName + "'", symbolTypeList[couS], "Kann Symbol Typ nicht ermitteln.", 1); 
            else
                symbolTypeNums.push(symbolTypeNum);
        }
        var filteredOccList = currentModel.ObjOccListBySymbol(symbolTypeNums);
        if (filteredOccList.length > 0)
        {
            var curObj; 
            var objDefListToCheck = new Array(); 
            for (var couO=0; couO<filteredOccList.length; couO++)
            {
                objDefListToCheck.push(filteredOccList[couO].ObjDef()); 
            }
            
            objDefListToCheck = ArisData.Unique(objDefListToCheck); 
            for (couO = 0; couO < objDefListToCheck.length; couO++)
            {
                if (g_debugMode == true)
                {
                    var ruleShortcut = g_config.getRuleShortcut(ruleNode);           // SP: Debug helper.
                    var curItemTyp   = objDefListToCheck[couO].Type();               // SP: Debug helper.
                    var curItemName  = objDefListToCheck[couO].Name(g_usedLanguage); // SP: Debug helper.
                    var AttrRuleType = g_config.getRuleType(attrRule);               // SP: Debug helper.
                    var compareType  = g_config.getAttrCompType(attrRule);           // SP: Debug helper.
                }
                doAttributeCheck(objDefListToCheck[couO], ruleNode, attrRule, currentModel);
            }
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///  Conventions check for assigned bcd models
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function doConventionsCheckBCDModels(bcdRules, superiorEPC)
{
    var riskObj; 
    var riskModels = null;
    var bcds = new Array(); 
    
    //find all assigned bcds of current model
    for (var couO = 0; couO < superiorEPC.ObjOccListFilter(Constants.OT_RISK).length; couO++)
    {
        riskObj =  superiorEPC.ObjOccListFilter(Constants.OT_RISK)[couO].ObjDef(); 
        riskModels = riskObj.AssignedModels(); 
        if (riskModels.length > 0)
        {
            for (var couM=0; couM<riskModels.length; couM++)
            {
                if (riskModels[couM].TypeNum() == Constants.MT_BUSY_CONTR_DGM)
                    bcds.push(riskModels[couM]); 
            }
        }
    }
    bcds = ArisData.Unique(bcds); 
    for (var couB = 0; couB < bcds.length; couB++)
    {
       for (var cou = 0; cou < bcdRules.size(); cou++)
       {
            var ruleNode     = bcdRules.get(cou); 
            var attrRule     = g_config.getAttrRuleNode(ruleNode);
            var typelist     = ruleNode.getAttributeValue("typelist");
            if (typelist != null && typelist != "")
            {
                if (typelist.indexOf(",") >= 0)
                    symbolTypes = typelist.split(",");
                else
                    symbolTypes = [typelist];
                var symbolTypeNum= 0; 
                var symbols = new Array(); 
                for (var couS=0; couS<symbolTypes.length; couS++)
                {
                    if (vbIsNumeric(symbolTypes[couS]))
                        symbolTypeNum = symbolTypes[couS];     
                    else if (vbLen(symbolTypes[couS] == 36))
                        symbolTypeNum = g_methodFilter.UserDefinedSymbolTypeNum(symbolTypes[couS]);

                    if (symbolTypeNum == 0)
                        LogMessage(getLogTypeForAttributeItemKind(bcds[couB]), "", "Regel '" + g_config.getRuleShortcut(ruleNode) + "'", symbolTypeList[couS], "Kann Symbol Typ nicht ermitteln.", 1); 
                    else
                        symbols.push(symbolTypeNum); 
                }
                var occList = bcds[couB].ObjOccListBySymbol(symbols); 
                for (var couO = 0; couO < occList.length; couO++)
                {
                    if (g_debugMode == true)
                    {
                        var ruleShortcut = g_config.getRuleShortcut(ruleNode);          // SP: Debug helper.
                        var curItemTyp   = occList[couO].ObjDef().Type();               // SP: Debug helper.
                        var curItemName  = occList[couO].ObjDef().Name(g_usedLanguage); // SP: Debug helper.
                        var attrRuleType = g_config.getRuleType(attrRule);              // SP: Debug helper.
                        var compareType  = g_config.getAttrCompType(attrRule);          // SP: Debug helper.
                    }
                    doAttributeCheck(occList[couO].ObjDef(), ruleNode, attrRule, bcds[couB]); 
                }
            }
            else
            {
                if (g_debugMode == true)
                {
                    var ruleShortcut = g_config.getRuleShortcut(ruleNode);  // SP: Debug helper.
                    var curItemTyp   = bcds[couB].Type();                   // SP: Debug helper.
                    var curItemName  = bcds[couB].Name(g_usedLanguage);     // SP: Debug helper.
                    var attrRuleType = g_config.getRuleType(attrRule);      // SP: Debug helper.
                    var compareType  = g_config.getAttrCompType(attrRule);  // SP: Debug helper.
                }
                doAttributeCheck(bcds[couB], ruleNode, attrRule, superiorEPC);
            }
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///  Conventions check for Attributes (epc and bcd, object, cxn)
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Check the attribute, containend in node AttrRule. 
// Function can be used for all items (models, objects, connections).
// SuperiorModel can be null for epc attributes.
function doAttributeCheck(currentItem, ruleNode, attrRule, superiorModel)
{
    var checkType        = "";
    var checkCategory    = "";
    var errorMessage     = "";
    var attrTypeNum      = 0;
    var minLength        = 0;
    var maxLength        = 0;
    var attrValue        = "";
    var compareType      = "";
    var attrValueType    = "";
    var ruleShortcut     = "";
    var showErrorIfEmpty = true;
    var ruleDescription  = "";

    if (attrRule != null)
    {
        attrTypeNum      = g_config.getAttributeTypeNum(ruleNode); 
        checkType        = String(__toString(g_config.getRuleType(attrRule)).trim()); // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
        checkCategory    = g_config.getCheckCategory(ruleNode); 
        errorMessage     = g_config.getErrorMessage(ruleNode); 
        ruleShortcut     = g_config.getRuleShortcut(ruleNode);
        showErrorIfEmpty = g_config.getShowErrorIfEmpty(attrRule); 
        ruleDescription  = g_config.getRuleDescription(ruleNode); 
        
        if (checkCategory <= g_currentLogType)
        {
            LogMessage(getLogTypeForAttributeItemKind(currentItem), ruleShortcut, ruleDescription,
                       GetText("TXT_ATTRIBUTE", new Array(currentItem.Attribute(attrTypeNum, g_usedLanguage).Type())),
                       "Attributwert: " + currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue(), 6); 
            switch (checkType)
            {
            case String("IsMaintained"): // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
                checkAttributeMaintained(currentItem, attrTypeNum, checkCategory, errorMessage, ruleShortcut, ruleDescription); 
                break;
            case String("CheckLength"): 
            
                minLength = g_config.getAttrMinLength(attrRule); 
                maxLength = g_config.getAttrMaxLength(attrRule); 
                if (minLength != null || maxLength != null)
                    checkAttributeLength(currentItem, attrTypeNum, minLength, maxLength, checkCategory, errorMessage, ruleShortcut, showErrorIfEmpty, ruleDescription); 
                break; 
            case String("CheckLineBreak"):
                checkLineBreak(currentItem, attrTypeNum, checkCategory, errorMessage, ruleShortcut, ruleDescription); 
                break; 
            case String("CheckValue"):
                startAt     = g_config.getAttrValueStartAt(attrRule); // Falls back to 0.
                endAt       = g_config.getAttrValueEndAt(attrRule);   // Falls back to 0.
                compareType = g_config.getAttrCompType(attrRule); 
                var attrValueNode      = attrRule.getChild("AttrValue");
                var compareAttrTypeNum = g_config.getAttributeTypeNum(attrValueNode, "compareToAttr");

                if (compareType == "IsEqualConnectedSymbol" || compareType == "IsEqualConnectedSymbolWithValue" || compareType == "IsEqualConnectedSymbolWithDiagramValue")
                {
                    //look for a connected symbol in the superior model (where the item has its occurences -> pick the first match)
                    var connectedSymbolTypeNum = g_config.getConnectedSymbolType(attrRule);
                    var pickConnectedObjDef = null;
                    
                    if (connectedSymbolTypeNum != 0 && compareAttrTypeNum != 0)
                    {
                        
                        var occs = currentItem.OccListInModel(superiorModel); 
                        var connectedObjOccs = null; 
                        for (var couO = 0; couO < occs.length; couO++)
                        {
                            if (occs[couO].Model().TypeNum() == Constants.MT_EEPC || occs[couO].Model().TypeNum() == Constants.MT_BUSY_CONTR_DGM)
                            {
                                connectedObjOccs = occs[couO].getConnectedObjOccs(connectedSymbolTypeNum); 
                                if (connectedObjOccs.length == 1)
                                {
                                    pickConnectedObjDef = connectedObjOccs[0].ObjDef();
                                    break;
                                }
                                else 
                                    LogMessage(getLogTypeForAttributeItemKind(currentItem), GetText("TXT_CHECKTYPEOBJECTATTRIBUTE"), ruleShortcut, ruleDescription,
                                               GetText("ERR_CONNECTEDOBJECTNOTUNIQUE", new Array(currentItem.Name(g_usedLanguage))), checkCategory); 
                            }
                        }
                    }
                    if (pickConnectedObjDef != null)
                    {
                        if (compareType == "IsEqualConnectedSymbol")
                        {
                            // Adjust the NEEDED attribute value before compare.
                            var prefix = (attrRule.getChild("AttrValue").getAttributeValue("prefix") == null ? "": String(attrRule.getChild("AttrValue").getAttributeValue("prefix"))); 
                            if (g_usedLanguage == 1033 && prefix == "Kontrolltest ") 
                            {
                                prefix = "Control test "
                            }
                            var suffix = (attrRule.getChild("AttrValue").getAttributeValue("appendix") == null ? "" : String(attrRule.getChild("AttrValue").getAttributeValue("appendix")));
                            attrValue = prefix + pickConnectedObjDef.Attribute(compareAttrTypeNum, g_usedLanguage).GetValue(true) + suffix;  
                            // if (attrValue=="Kontrolltest no control available"){attrValue="no control test"} // [SM] Sonderfall für "no control available" -> gibt's nicht mehr
                        }
                        else if (compareType == "IsEqualConnectedSymbolWithValue")
                        {
                            var compareValueNeeded = attrValueNode.getAttributeValue("compareToValue");
                            attrValue = g_config.getAttrValue(attrRule);
                            if (currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() == attrValue)
                            {
                                if (pickConnectedObjDef.Attribute(compareAttrTypeNum, g_usedLanguage).IsMaintained() == false)
                                {
                                    if (compareValueNeeded != "")
                                        attrValue = "trigger-an-error";
                                }
                                else if (pickConnectedObjDef.Attribute(compareAttrTypeNum, g_usedLanguage).GetValue(true) != compareValueNeeded)
                                  attrValue = "trigger-an-error";  
                            }
                            else
                            {
                                // Don't test this QA check, if prerequisits aren't met.
                                break;
                            }
                        }
                        else // (compareType == IsEqualConnectedSymbolWithDiagramValue)
                        {
                            // The currentItem and the pickConnectedObjDef will be compared by one single attribute utilizing: attrTypeNum.
                            // The prerequisit criterion is based on superiorModel and the config values "compareToAttr" and "compareToValue".

                            compareValueNeeded      = currentItem.Attribute(attrTypeNum, g_usedLanguage).GetValue(true);
                            var diagramCompareValue = superiorModel.Attribute(compareAttrTypeNum, g_usedLanguage).GetValue(true);
                            var diagramValueNeeded  = attrValueNode.getAttributeValue("compareToValue");

                            if (diagramCompareValue == diagramValueNeeded)
                            {
                                var prefix         = (attrRule.getChild("AttrValue").getAttributeValue("prefix") == null ? "": String(attrRule.getChild("AttrValue").getAttributeValue("prefix"))); 
                                var suffix         = (attrRule.getChild("AttrValue").getAttributeValue("appendix") == null ? "" : String(attrRule.getChild("AttrValue").getAttributeValue("appendix")));
                                attrValue          = prefix + pickConnectedObjDef.Attribute(attrTypeNum, g_usedLanguage).GetValue(true) + suffix;

                                if (pickConnectedObjDef.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == false)
                                {
                                    if (compareValueNeeded != "")
                                        attrValue = "trigger-an-error";
                                }
                                else if (attrValue != compareValueNeeded)
                                  attrValue = "trigger-an-error";  
                            }
                            else
                            {
                                // Don't test this QA check, if prerequisits aren't met.
                                break;
                            }
                        }

                        // The 'currentItem' 'attrTypeNum' and the 'pickConnectedObjDef' have been tested.
                        // No need to adjust the NEEDED attribute value before compare.
                        // Fall through to the default 'checkAttrValue()' function with 'currentItem' 'attrTypeNum'.
                    }
                }
                else if (compareType == "IsEqualContainingBCDWithValue")
                {
                    compareAttrTypeNum = g_config.getAttributeTypeNum(attrValueNode, "compareToAttr");
                    var compareValueNeeded = attrValueNode.getAttributeValue("compareToValue");
                    attrValue = g_config.getAttrValue(attrRule);
                    if (currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() == attrValue)
                    {
                        if (superiorModel.Attribute(compareAttrTypeNum, g_usedLanguage).GetValue(true) != compareValueNeeded)
                          attrValue = "trigger-an-error";  
                    }
                    else
                    {
                        // Don't test this QA check, if prerequisits aren't met.
                        break;
                    }
                }
                else
                {
                    if (compareAttrTypeNum == 0)
                        attrValue = g_config.getAttrValue(attrRule); 
                    else if (compareAttrTypeNum > 0) 
                        attrValue = currentItem.Attribute(compareAttrTypeNum, g_usedLanguage).GetValue(true); 
                }
                
                attrValueType = g_config.getAttrValueType(attrRule);
                checkAttrValue(currentItem, attrTypeNum, attrValue, compareType, attrValueType, startAt, endAt, checkCategory, errorMessage, ruleShortcut, showErrorIfEmpty, ruleDescription); 
                break;
            case String("CheckDate"): 
                attrValue   = g_config.getAttrValue(attrRule); 
                compareType = g_config.getAttrCompType(attrRule); 
                checkAttrDate(currentItem, attrTypeNum, attrValue, compareType, checkCategory, errorMessage, ruleShortcut, ruleDescription); 
                break;
            case String("CheckValueByRegExp"):
            case String("CheckDependence"):
                var regExpString = g_config.getRegExp(attrRule); 
                compareType = g_config.getAttrCompType(attrRule); 
                if (regExpString != null)
                {
                    // The primary rule is aRegEx rule.
                    var depAttributeNode = g_config.getDependentAttributeNode(attrRule);
                    if (depAttributeNode == null)
                    {
                        checkAttrValueByRegExp(currentItem, attrTypeNum, regExpString, null,             checkCategory, errorMessage, ruleShortcut, showErrorIfEmpty, ruleDescription); 
                    }
                    else
                        checkAttrValueByRegExp(currentItem, attrTypeNum, regExpString, depAttributeNode, checkCategory, errorMessage, ruleShortcut, showErrorIfEmpty, ruleDescription); 
                }
                else
                {
                    var depAttributeNode = g_config.getDependentAttributeNode(attrRule); 
                    if (depAttributeNode!= null)
                    {
                        
                        var depType = g_config.getDependenceType(depAttributeNode); 
                        //if depType = "IsMaintained" there is no dependent attr value and comparison type
                        var depCompType = ""; 
                        var depAttrTypeNum =0; 
                        var attrValue = ""; 
                        depAttrTypeNum = g_config.getAttributeTypeNum(depAttributeNode); 
                        if (depType == "CheckValue" || depType == "CheckValueOfModelAttr")
                        {
                            depCompType = g_config.getAttrCompType(depAttributeNode); 
                            attrValue = g_config.getAttrValue(attrRule); 
                            var depAttrValue = g_config.getAttrValue(depAttributeNode);
                        }
                        if (depType == "CheckValueOfModelAttr")
                        {
                            depItem = g_currentModel; 
                        }
                        else
                            depItem = null; 
    
                        if (attrValue == null)
                        {
                            //                  2.Check-Item, 2.Check-AttrTypeNum, 2.Check-CompType, 2.Check-AttrValue, 1.Check-Item, 1.Check-AttrTypeNum, 1.Check-CompType, 1.Check-AttrValue, ...
                            checkAttrDependence(currentItem,  attrTypeNum,         "",               null,              depItem,      depAttrTypeNum,      depCompType,      depAttrValue,
                                                checkCategory, errorMessage, ruleShortcut, ruleDescription); 
                        }
                        else
                        {
                            //                  2.Check-Item, 2.Check-AttrTypeNum, 2.Check-CompType, 2.Check-AttrValue, 1.Check-Item, 1.Check-AttrTypeNum, 1.Check-CompType, 1.Check-AttrValue, ...
                            checkAttrDependence(currentItem,  attrTypeNum,         compareType,      attrValue,         depItem,      depAttrTypeNum,      depCompType,      depAttrValue,
                                                checkCategory, errorMessage, ruleShortcut, ruleDescription); 
                        }
                    }
                }
                break;
            
            case String("EqualToSuperiorModel"):
               if (superiorModel != null)
               {
                   startAt = g_config.getAttrValueStartAt(attrRule); // Falls back to 0.
                   endAt = g_config.getAttrValueEndAt(attrRule); // Falls back to 0.
                   if (endAt > 0)
                       attrValue = superiorModel.Attribute(attrTypeNum, g_usedLanguage).getValue().substring(startAt, endAt);
                   else
                       attrValue = superiorModel.Attribute(attrTypeNum, g_usedLanguage).getValue().substring(startAt); 
                   compareType   = "IsEqual"; 
                   attrValueType = "String"; 
                   checkAttrValue(currentItem, attrTypeNum, attrValue, compareType, attrValueType, startAt, endAt, checkCategory, errorMessage, ruleShortcut, showErrorIfEmpty, ruleDescription); // [SM]: Aufrufparameter showErrorIfEmpty eingepflegt, da sonst Beschreibung = 'undefined'
               }
               break;
            case String("EqualToSuperiorObject"):
                var superiorObjDefList = currentItem.getSuperiorObjDefs();
                var superiorObj = null;
                if (superiorObjDefList.length > 1)
                    LogMessage(getLogTypeForAttributeItemKind(currentItem), ruleShortcut, ruleDescription, GetText("TXT_BCDCHECK"),
                               GetText("ERR_BCDMORETHANONESUPERIOR", new Array(currentItem.Name(g_usedLanguage))), 1); 
                else if (superiorObjDefList.length == 1)
                    superiorObj = superiorObjDefList[0]; 
                else
                    LogMessage(getLogTypeForAttributeItemKind(currentItem), ruleShortcut, ruleDescription, GetText("TXT_BCDCHECK"),
                               GetText("ERR_BCDLESSTHANONESUPERIOR", new Array(currentItem.Name(g_usedLanguage))), 1);

                startAt = g_config.getAttrValueStartAt(attrRule); // Falls back to 0.
                endAt = g_config.getAttrValueEndAt(attrRule); // Falls back to 0.
                //read the attribute value of the superior object without linebreaks because otherwise the comparison will in most cases fail
                if (endAt > 0)
                    attrValue = superiorObj.Attribute(attrTypeNum, g_usedLanguage).GetValue(true).substring(startAt, endAt);
                else
                    attrValue = superiorObj.Attribute(attrTypeNum, g_usedLanguage).GetValue(true).substring(startAt); 
                
                compareType   = "IsEqual";
                attrValueType = "String";
                checkAttrValue(currentItem, attrTypeNum, attrValue, compareType, attrValueType, startAt, endAt, checkCategory, errorMessage, ruleShortcut, showErrorIfEmpty, ruleDescription); 
                break;
            case String("AssignedBCDWithValue"):
                var superiorObjDefList = currentItem.getSuperiorObjDefs(); 
                if (superiorObjDefList.length > 1)
                    LogMessage(getLogTypeForAttributeItemKind(currentItem), ruleShortcut, ruleDescription, GetText("TXT_BCDCHECK"),
                               GetText("ERR_BCDMORETHANONESUPERIOR", new Array(currentItem.Name(g_usedLanguage))), 1); 
                else if (superiorObjDefList.length == 1)
                    var superiorObj = superiorObjDefList[0]; 
                else
                    LogMessage(getLogTypeForAttributeItemKind(currentItem), ruleShortcut, ruleDescription, GetText("TXT_BCDCHECK"),
                               GetText("ERR_BCDLESSTHANONESUPERIOR", new Array(currentItem.Name(g_usedLanguage))), 1);

                startAt = g_config.getAttrValueStartAt(attrRule); // Falls back to 0.
                endAt = g_config.getAttrValueEndAt(attrRule); // Falls back to 0.
                //read the attribute value of the superior object without linebreaks because otherwise the comparison will in most cases fail
                if (endAt > 0)
                    attrValue = superiorObj.Attribute(attrTypeNum, g_usedLanguage).GetValue(true).substring(startAt, endAt);
                else
                    attrValue = superiorObj.Attribute(attrTypeNum, g_usedLanguage).GetValue(true).substring(startAt); 
                
                compareType   = g_config.getAttrCompType(attrRule);
                attrValueType = "String";

                var compareValueNeeded = g_config.getAttrValue(attrRule); 

                if (compareType == "IsEqual" && attrValue == compareValueNeeded)
                {
                    var depAttributeNode = g_config.getDependentAttributeNode(attrRule); 
                    if (depAttributeNode!= null)
                    {
                        
                        var depType = g_config.getDependenceType(depAttributeNode); 
                        //if depType = "IsMaintained" there is no dependent attr value and comparison type
                        var depCompType = ""; 
                        var depCompAttrValue = ""; 
                        var depAttrTypeNum = g_config.getAttributeTypeNum(depAttributeNode);
                        var assAttrValue = "";
                        
                        if (depAttrTypeNum > 0)
                            assAttrValue = currentItem.Attribute(depAttrTypeNum, g_usedLanguage).GetValue(true);
    
                        if (depType == "CheckValue")
                        {
                            depCompType = g_config.getAttrCompType(depAttributeNode); 
                            depCompAttrValue = g_config.getAttrValue(depAttributeNode);
                            
                            if (attrValue == null)
                            {
                                //                  2.Check-Item, 2.Check-AttrTypeNum, 2.Check-CompType, 2.Check-AttrValue, 1.Check-Item, 1.Check-AttrTypeNum, 1.Check-CompType, 1.Check-AttrValue, ...
                                checkAttrDependence(currentItem,  depAttrTypeNum,      depCompType,      depCompAttrValue,  superiorObj,  attrTypeNum,         "",               null,
                                                    checkCategory, errorMessage, ruleShortcut, ruleDescription); 
                            }
                            else
                            {
                                //                  2.Check-Item, 2.Check-AttrTypeNum, 2.Check-CompType, 2.Check-AttrValue, 1.Check-Item, 1.Check-AttrTypeNum, 1.Check-CompType, 1.Check-AttrValue, ...
                                checkAttrDependence(currentItem,  depAttrTypeNum,      depCompType,      depCompAttrValue,  superiorObj,  attrTypeNum,         compareType,      attrValue,
                                                    checkCategory, errorMessage, ruleShortcut, ruleDescription); 
                            }
                        }                    
                        else if (attrValue != assAttrValue)
                        {
                            var attrName = currentItem.Attribute(attrTypeNum, g_usedLanguage).Type();
                            LogMessage(getLogTypeForAttributeItemKind(currentItem), ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errorMessage, checkCategory); 
                        }
                    }
                }
                break;
            case String("EqualToAssignedModel"): 
                startAt = g_config.getAttrValueStartAt(attrRule); // Falls back to 0.
                endAt = g_config.getAttrValueEndAt(attrRule); // Falls back to 0.

                if (currentItem.AssignedModels().length > 0)
                {
                    attrValue = currentItem.AssignedModels()[0].Attribute(attrTypeNum, g_usedLanguage).getValue(); 

                    compareType = g_config.getAttrCompType(attrRule); 
                    checkAttrValue(currentItem, attrTypeNum, attrValue, compareType, attrValueType, startAt, endAt, checkCategory, errorMessage, ruleShortcut, showErrorIfEmpty, ruleDescription); // [SM]: Aufrufparameter showErrorIfEmpty eingepflegt, da sonst Beschreibung = 'undefined'
                 }
                 break; 
            case String("checkCustom"): 
                custCheck = g_config.getAttrValue_customCheck(attrRule); 
                checkCustomAttr(currentItem, attrTypeNum, checkCategory, errorMessage, ruleShortcut, ruleDescription, custCheck);
                break; 
            default:
                Dialogs.MsgBox(getString("ERR_WRONGATTRCHECKTYPE") + checkType, Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION")); 
            }
        }
    }
    else
        Dialogs.MsgBox(GetText("ERR_RULEINCOMPLETE", new Array(g_config.getRuleName(ruleNode))), Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION")) ; 
}

function checkAttributeMaintained(currentItem, attrTypeNum, checkCategory, errMessage, ruleShortcut, ruleDescription)
{
    var logType= getLogTypeForAttributeItemKind(currentItem);
    if (currentItem.KindNum() == Constants.CID_OBJDEF)
        errMessage = errMessage.replace("{0}", "'" + currentItem.Name(g_usedLanguage) + "'").replace("{1}", currentItem.Type()) ; 
    if (currentItem.KindNum() == Constants.CID_MODEL)
        errMessage = errMessage.replace("{0}", "'" + currentItem.Name(g_usedLanguage) + "'").replace("{1}", currentItem.Type()) ; 
       
    if (currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == false || currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() == "")
        LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(currentItem.Attribute(attrTypeNum, g_usedLanguage).Type())), errMessage, checkCategory); 
    else
        LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(currentItem.Attribute(attrTypeNum, g_usedLanguage).Type())), errMessage + " - OK", 5);
}

function checkLineBreak(currentItem, attrTypeNum, checkCategory, errMessage, ruleShortcut, ruleDescription)
{
    if (checkCategory > g_currentLogType)
        return; 
    
    var logType = getLogTypeForAttributeItemKind(currentItem)
    if (currentItem.KindNum() == Constants.CID_OBJDEF)
        errMessage = errMessage.replace("{0}", "'" + currentItem.Name(g_usedLanguage)+"'").replace("{1}", currentItem.Type()) ; 
    else if (currentItem.KindNum() == Constants.CID_MODEL)
        errMessage = errMessage.replace("{0}", "'" + currentItem.Name(g_usedLanguage)+"'"); 
         
    if (currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained())
    {
        var attributeValue = currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue();        // Attribut mit Zeilenumbrüchen holen
        attributeValue = attributeValue + "";                               // Sicherstellen, dass Attribut vom Typ string ist
        // attributeValue = attributeValue.replace(/^- |\n- |\. \r\n|\.\r\n|:\r\n|: \r\n|\)\r\n|\) \r\n/gi , "");   // Alle richtig verwendeten Spiegelstriche entfernen
        // attributeValue = attributeValue.replace(/^- |\n- |- ([\w\"\.\:\(\)\\\/, -]*)\r\n|\. \r\n|\.\r\n|:\r\n|: \r\n/gi , "");   // Alle richtig verwendeten Spiegelstriche entfernen
        attributeValue = attributeValue.replace(/- ([\w\"\.\:\(\)\\\/,?!\% -]*)\r\n|^- |\n- |\. \r\n|\.\r\n|:\r\n|: \r\n/gi , "");
        
        var searchLineBreaks = attributeValue.search(/\n/);                    // Nach übrig gebliebenen, d.h. unerlaubten Umbrüchen suchen

        if (searchLineBreaks > -1)
            LogMessage(logType , ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(currentItem.Attribute(attrTypeNum, g_usedLanguage).Type())), errMessage, checkCategory); 
        else
            LogMessage(logType , ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(currentItem.Attribute(attrTypeNum, g_usedLanguage).Type())), errMessage + " - OK", 5); 
    }
    
}

function checkAttributeLength(currentItem, attrTypeNum, minLength, maxLength, checkCategory, errMessage, ruleShortcut, showErrorIfEmpty, ruleDescription)
{
    var logType = getLogTypeForAttributeItemKind(currentItem); 
    if (currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained())
    {
        var attrValue= currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue(); 
        errMessage = errMessage.replace("{0}", attrValue); 
        var bError = false;
        if (minLength !=  null && maxLength ==  null)
        {
            if (vbLen(attrValue) < minLength)
            {
                LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(currentItem.Attribute(attrTypeNum, g_usedLanguage).Type())), errMessage, checkCategory); 
                bError = true; 
            }
        }
        else if (maxLength != null && minLength == null)
        {
            if (vbLen(attrValue) > maxLength)
            {
                LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(currentItem.Attribute(attrTypeNum, g_usedLanguage).Type())) , errMessage, checkCategory); 
                bError = true; 
            }
        }
        else if (minLength != null && maxLength != null)
        {
            if (vbLen(attrValue) > maxLength || vbLen(attrValue) < minLength)
            {
                LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(currentItem.Attribute(attrTypeNum, g_usedLanguage).Type())), errMessage, checkCategory); 
                bError = true; 
            }
        }
        //check okay -> no error
        if (bError == false)
        {
            LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(currentItem.Attribute(attrTypeNum, g_usedLanguage).Type())), errMessage + " - OK", 5); 
        }
    }    
    else
    {
        if (showErrorIfEmpty)
            LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(currentItem.Attribute(attrTypeNum, g_usedLanguage).Type())), errMessage + " - " + getString("TXT_ATTRIBUTENOTMAINTAINED"), checkCategory); 
    }
}

//parameters: 
// currentItem: model, objDef, cxnDef
// attrTypeNum: a valid aris attribute type num, used to determite attribute type-name from currentItem (==> attrName) and attribute value from currentItem (==> attrValue)
// attrValueNeeded: the value which is used for the comparison (can contain a comma separated list of values)
// compType: the comparison type
// attrValueType: the data type (number, date or string)
// startAt: a number for a substring comparison (startAt Position, or 0 if not needed, or 999 if all text before the forst number [0-9] is to skip)
// endAt: a number for a substring comparison (endAt Position or 0 if not needed)
// checkCategory: error category
// errMessage: the error text out of the xml (evtl. replacements with item name/type needed to create a meaningful message)
function checkAttrValue(currentItem, attrTypeNum, attrValueNeeded, compType, attrValueType, startAt, endAt, checkCategory, errMessage, ruleShortcut, showErrorIfEmpty, ruleDescription)
{
    if (checkCategory > g_currentLogType)
        return; 
    
    var logType = getLogTypeForAttributeItemKind(currentItem); 
    var attrValue=""; 
    var attrName = currentItem.Attribute(attrTypeNum, g_usedLanguage).Type(); 
        
    if (currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained())
    {
            if (startAt == 999)
            {
                startAt = currentItem.Attribute(attrTypeNum, g_usedLanguage).GetValue(true).search(/[0-9]/);
                if (startAt == -1){startAt = 0} 
            }     

            if (endAt > 0)
                attrValue = currentItem.Attribute(attrTypeNum, g_usedLanguage).GetValue(true).substring(startAt, endAt); 
            else
                attrValue = currentItem.Attribute(attrTypeNum, g_usedLanguage).GetValue(true).substring(startAt); 

        var bError = false; 
        var comparisonType = "";
        comparisonType = String(compType); // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
        var bFound=false;
        
        if (attrValue != "")
        {
            valuesNeeded = attrValueNeeded.split(",");
            var usedWords = ""; 
            switch (comparisonType)
            {
                case String("IsEqual"): // SP: The compare value must be a JS object, not a Java object - like java.lang.String!
                case String("IsEqualConnectedSymbol"):
                case String("IsEqualConnectedSymbolWithValue"):
                case String("IsEqualConnectedSymbolWithDiagramValue"):
                case String("IsEqualContainingBCDWithValue"):
                    bFound=false;
                    for (var cou = 0; cou < valuesNeeded.length; cou++)
                    {
                        if (vbIsNumeric(valuesNeeded[cou]))
                        {
                            if (__toNumeric(attrValue) == valuesNeeded[cou])
                            {
                                bFound = true;
                                break; 
                            }
                            usedWords = valuesNeeded[cou]; 
                        }
                        else
                        {
                            if(attrValue == attrValueNeeded)
                            {
                                bFound = true; 
                                break;
                            }
                            usedWords = attrValue; 
                        }
                    }
                    if (bFound==false)
                        bError = true; 
                    break;
                case String("IsGreater"):
                        if (__toNumeric(attrValue) < __toNumeric(attrValueNeeded))
                            bError = true; 
                        usedWords = attrValue; 
                    break;
                case String("IsSmaller"):
                    if (__toNumeric(attrValue) > __toNumeric(attrValueNeeded))
                            bError = true; 
                        usedWords = attrValue; 
                    break;
                case String("IsNotEqual"):
                    bFound=false; 
                    for (var cou=0; cou<valuesNeeded.length; cou++)
                    {
                        if (attrValue == String(valuesNeeded[cou]))
                        {
                            bFound = true;
                            break; 
                        }
                    }
                    if (bFound)
                        bError = true; 
                    usedWords = ""; 
                        
                    break;
                case String("Contains"):
                
                    if (String(attrValueNeeded).indexOf(",") > 0)
                    {
                        var allowedValues = attrValueNeeded.split(","); 
                        var valueFound = false; 
                        for (var cou=0; cou<allowedValues.length; cou++)
                        {
                            var allowedValue = allowedValues[cou]+""; 
                            allowedValue = vbTrim(allowedValue); 
                            //allowedValue = allowedValue.replace("\"", "\\b"); 
                            //allowedValue = allowedValue.replace("\"", "\\b");
                            allowedValue = allowedValue.replace("\"", ""); 
                            allowedValue = allowedValue.replace("\"","");  
                            var regExp = new RegExp(allowedValue);
                            if (attrValue.search(regExp) > -1)
                            {
                                if (attrValue.length()==allowedValue.length)
                                {
                                    valueFound=true; 
                                    usedWords = attrValue; 
                                    break; 
                                }
                            }
                        }
                        if (valueFound == false)
                        {
                            bError = true; 
                            usedWords = attrValue; 
                        }
                    }
                    else
                    {
                        if (String(attrValueNeeded).indexOf(attrValue) < 0) 
                            bError = true; 
                        usedWords = attrValue; 
                    }
                    break;
                case String("ContainsNot"):
                    var regExp; 
                    if (attrValueNeeded.indexOf(",")> 0)
                    {
                        var unallowedValues = attrValueNeeded.split(","); 
                        var valueFound = false; 
                        
                        for (var cou=0; cou<unallowedValues.length; cou++)
                        {

                            if (unallowedValues[cou].replace("\"", "").replace("\"", "").length > 2)
                            {
                                var unallowedValue = unallowedValues[cou]+""; 
                                unallowedValue = unallowedValues[cou].replace("\"", "\\b"); 
                                unallowedValue = unallowedValue.replace("\"", "");      // [SM]: urspr. unallowedValue = unallowedValue.replace("\"", "\\b"); regex funktioniert mit\b am Ende nicht richtig
                                unallowedValue = unallowedValue.replace(/\./g, "\\."); //  [SM]: Punkte müssen escaped werden in RegEx steht der Punkt für ein beliebiges Zeichen
                                regExp = new RegExp(unallowedValue);
                                if (attrValue.search(regExp) > -1)
                                {
                                    valueFound=true; 
                                    usedWords = usedWords + unallowedValues[cou] +", "; 
                                    //break;
                                }
                            }
                            else
                            {
                                if (attrValue.search(/[ ]{2,}/) > -1)//search for double spaces
                                {
                                    valueFound=true;
                                    usedWords = usedWords + "doppelte Leerzeichen oder Leerzeichen vor Zeilenumbruch, "; 
                                    //break; 
                                }
                            }
                        }
                    }
                    else
                    {
                        if (attrValueNeeded.replace("\"", "").length() > 2)
                        {
                            attrValueNeeded = attrValueNeeded.replace("\"", "\\b"); 
                            attrValueNeeded = attrValueNeeded.replace("\"", "\\b"); 
                            regExp = new RegExp(attrValueNeeded); 
                             if (attrValue.search(regExp)> 0)
                                {
                                    valueFound=true; 
                                    usedWords = attrValueNeeded; 
                                }
                        }
                        else
                        {
                            if (attrValue.search(/[ ]{2,}/) > -1)//search for double spaces
                            {
                                valueFound=true;
                                usedWords = usedWords + "doppelte Leerzeichen, "; 
                                //break; 
                            }
                            if (attrValue.search(/[!]{1,}/) > -1)//search for exclamation marks
                            {
                                valueFound=true;
                                usedWords = usedWords + "Ausrufezeichen!, "; 
                                //break; 
                            }
                        }
                    }
                    if (valueFound)
                        bError = true; 
                    break; 
                default:
                    Dialogs.MsgBox("Wrong attr configuration in comparison type ('comp' = " + comparisonType +")", Constants.MSGBOX_BTN_OK, "Configuration error");
            } 
        }
        else
        {
            if (showErrorIfEmpty) // it is not always necessary to show an error if the attribute is empty (e.g. if the attribute maintaining already has been checked)
                bError = true; 
        }
        if (currentItem.KindNum() == Constants.CID_OBJDEF)
            errMessage = errMessage.replace("{0}", currentItem.Name(g_usedLanguage)).replace("{1}", currentItem.Type()).replace("{2}", usedWords); 
        else if (currentItem.KindNum() == Constants.CID_CXNDEF)
            errMessage = errMessage.replace("{0}", currentItem.SourceObjDef().Name(g_usedLanguage)).replace("{1}", currentItem.TargetObjDef().Name(g_usedLanguage)) + " Kantenrolle: "+ usedWords; 
        else if(currentItem.KindNum() == Constants.CID_MODEL)
        {
            if (currentItem.TypeNum() == Constants.MT_BUSY_CONTR_DGM)
                errMessage = errMessage.replace("{0}", currentItem.Name(g_usedLanguage)).replace("{1}", attrValue); 
            else
                errMessage = errMessage.replace("{0}", attrValue); 
        }
        if (bError)
        {
            LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errMessage, checkCategory); 
        }
        else
            LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errMessage + " - OK", 5); 
    }
    else
    {
        if (showErrorIfEmpty) // it is not always necessary to show an error if the attribute is empty (e.g. if the attribute maintaining already has been checked)
        {
            if (currentItem.KindNum() == Constants.CID_OBJDEF)
                errMessage = errMessage.replace("{0}", currentItem.Name(g_usedLanguage)).replace("{1}", currentItem.Type()) + ": "; 
            else if (currentItem.KindNum() == Constants.CID_CXNDEF)
                errMessage = errMessage.replace("{0}", currentItem.SourceObjDef().Name(g_usedLanguage)).replace("{1}", currentItem.TargetObjDef().Name(g_usedLanguage)); 
            else if(currentItem.KindNum() == Constants.CID_MODEL)
            {
                if (currentItem.TypeNum() == Constants.MT_BUSY_CONTR_DGM)
                    errMessage = errMessage.replace("{0}", currentItem.Name(g_usedLanguage)); 
            }
            errMessage = errMessage + " - " + getString("TXT_ATTRIBUTENOTMAINTAINED");    
            LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errMessage, checkCategory); 
        }
    }
}

function checkAttrDate(currentItem, attrTypeNum, compValue, compType, checkCategory, errorMessage, ruleShortcut, ruleDescription)
{
    if (checkCategory > g_currentLogType)
        return; 
    
    var logType = getLogTypeForAttributeItemKind(currentItem); 
    var attrName = currentItem.Attribute(attrTypeNum, g_usedLanguage).Type(); 
    if (currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained())
    {
        var attrValue = currentItem.Attribute(attrTypeNum, g_usedLanguage).getValueStd(); // [SM] .getValueStd() eingepflegt, da ansonsten Invalid Date
        
        var attrDate = new Date(attrValue);
        var now = new Date();
        if (compType == "IsSmaller")
        {
            if ((now.getTime() - attrDate.getTime()) / 1000 / 60 / 60 /24 > compValue) //convert milliseconds to days
                LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errorMessage, checkCategory); 
        }
        else if (compType == "IsGreater")
        {
            if ((now.getTime() - attrDate.getTime()) / 1000 / 60 / 60 /24 < compValue) //convert milliseconds to days
                LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errorMessage, checkCategory); 
        }
        else
            LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errorMessage + " - OK", 5); 
    }
}

// Check only primary attribute or primary and secondary attribute - only the primary attribute must be a RegEx check, the secondary attribute can be a simple or a RegEx check.
//parameters: 
// currentItem: the item that holds the attribute(s) to check - model, objDef, cxnDef
// attrTypeNum: a valid aris attribute type num, used to determite attribute type-name from currentItem (==> attrName) and attribute value from currentItem (==> attrValue)
// regExpString: the RegEx strint to use for the primary attribute check
// dependentAttributeNode: the optional attribute check configuration node for the secondary attribute check - can be null, is applied to the (same) currentItem
// checkCategory: the check category for output
// ruleShortcut: the rule ID
// showErrorIfEmpty: set to false, if primary attributes can not be accepted as not maintained/empty 
// ruleDescription: the rule description for output
function checkAttrValueByRegExp(currentItem, attrTypeNum, regExpString, dependentAttributeNode, checkCategory, errorMessage, ruleShortcut, showErrorIfEmpty, ruleDescription)
{
    var logType = getLogTypeForAttributeItemKind(currentItem); 
    var bSuccess; 
    var bError = false; 
    var attrName = currentItem.Attribute(attrTypeNum, g_usedLanguage).Type(); 
    var depAttrValue = ""; 

    try
    {
        var primaryRegExp = new RegExp(regExpString); 
        if (currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() && currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue().length() > 0)
        {
            if (dependentAttributeNode == null)
            {
                // no dependent attribute (no secondary attribute check to apply)
                bError  =!(helperCheckRegExp(primaryRegExp, currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue())); 
            }
            else
            {
                // dependent attribute (secondary attribute check is to apply)
                var depAttrTypeNum = g_config.getAttributeTypeNum(dependentAttributeNode);
                var isCaseSensitive = g_config.AttributeCheckIsCaseSensitive(dependentAttributeNode);

                if (g_config.getRegExp(dependentAttributeNode) != null)
                {
                    // dependent attribute with regexp - combined evaluation for both checks
                    var secondaryRegExp = new RegExp(g_config.getRegExp(dependentAttributeNode)); 
                    
                    depAttrValue = currentItem.Attribute(depAttrTypeNum, g_usedLanguage).getValue(); 
                    if (depAttrValue != "")
                    {
                        // secondary check
                        var secondaryCheckResult = secondaryRegExp.exec(depAttrValue);
                        
                        // primary check
                        var primaryCheckResult = primaryRegExp.exec(currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue());
                        
                        // combined result evaluation
                        if (primaryCheckResult != null && secondaryCheckResult != null)
                        {
                            if (primaryCheckResult.length != secondaryCheckResult.length)//wenn ein unterschiedlich großes Array gefunden wurde, können die ausdrücke nicht gleich sein
                                bError = true; 
                            else
                            {
                                for (var cou = 1; cou < primaryCheckResult.length; cou++)
                                {
                                    if (isCaseSensitive)
                                    {
                                        //if (primaryCheckResult[cou] != secondaryCheckResult[cou]) 
                                        if (encodefilename(primaryCheckResult[cou]) != secondaryCheckResult[cou])    // [SM]: encodefilname wird hier nur benutzt, um äöüß zu durch E-Mail-Kompatible Werte zu ersetzen
                                        {
                                            bError = true; 
                                            break; 
                                        }
                                    }
                                    else
                                    {
                                         //if (primaryCheckResult[cou].toLowerCase() != secondaryCheckResult[cou].toLowerCase())
                                        if (encodefilename(primaryCheckResult[cou].toLowerCase()) != secondaryCheckResult[cou].toLowerCase()) // [SM]: encodefilname wird hier nur benutzt, um äöüß zu durch E-Mail-Kompatible Werte zu ersetzen
                                        {
                                            bError = true; 
                                            break; 
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // throw "error while executing regular expression"; 
                        }
                    }
                }
                else
                {
                    // dependent attribute without regexp) - secondary check evaluation before primary check evaluation
                    var depType        = g_config.getDependenceType(dependentAttributeNode); 
                    
                    if (depType == "CheckValue")
                    {
                        // secondary check
                        var depAttrValue = g_config.getAttrValue(dependentAttributeNode);
                        var depCompType = g_config.getAttrCompType(dependentAttributeNode); 
                        if (currentItem.Attribute(depAttrTypeNum, g_usedLanguage).getValue().toLowerCase() == depAttrValue.toLowerCase())
                        {
                            // primary check
                            bError = !(helperCheckRegExp(primaryRegExp, currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue()));  
                        }
                    }
                    else // "IsMaintained"
                    {
                        // secondary check
                        if (currentItem.Attribute(depAttrTypeNum, g_usedLanguage).IsMaintained())
                        {
                            // primary check
                            bError = !(helperCheckRegExp(primaryRegExp, currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue())); 
                        }
                    }
                }
            }
        }
        else
        {
            if (showErrorIfEmpty) // it is not always necessary to show an error if the attribute is empty (e.g. if the attribute maintaining already has been checked)
            {
                bError = true;            
                errorMessage = errorMessage + " - " + getString("TXT_ATTRIBUTENOTMAINTAINED"); 
            }
        }
        
        if (currentItem.KindNum() == Constants.CID_OBJDEF)
            errorMessage = errorMessage.replace("{0}", currentItem.Name(g_usedLanguage)).replace("{1}", currentItem.Type()); 
        
        if (bError)
        {
            var errorMessageDetails = ""; 
            if (depAttrValue != "")
                errorMessageDetails = currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() + " - " + depAttrValue;
            else
                errorMessageDetails = currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue(); 

           
            LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errorMessage +": " + errorMessageDetails, checkCategory); 
        }
        else
            LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errorMessage + " - OK", 5); 
    }
    catch (err)
    {
        Dialogs.MsgBox(getString("ERR_REGEXPFAILED") + " (" + err + "): " + regExpString + " AttrValue: " + currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue(), Constants.MSGBOX_BTN_OK, "Error reading xml");
    }
}

function checkAttrDependence(secondaryCheckItem, secondaryCheckAttrTypeNum, secondaryCheckCompType, secondaryCheckAttrValue, primaryCheckItem, primaryCheckAttrTypeNum, primaryCheckCompType, primaryCheckAttrValue,
                             checkCategory, errorMessage, ruleShortcut, ruleDescription)
{
    var bAttrCondition = false; 
    var bError = false; 
    var logType = getLogTypeForAttributeItemKind(secondaryCheckItem); 
    var attrName = secondaryCheckItem.Attribute(secondaryCheckAttrTypeNum, g_usedLanguage).Type(); 
    
    //the dependent attribute can be an attribute of the superior model or an attribute of the current item
    if (primaryCheckItem == null)
        primaryCheckItem = secondaryCheckItem; 
    
    if (vbLen(String(primaryCheckCompType)) > 0)
    {
        switch (String(primaryCheckCompType)) // SP: The compared value must be a JS object, not a Java object - like java.lang.String! //comparisonType of dependent Attribute
        {
            case String("IsEqual"): // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
                if (primaryCheckItem.Attribute(primaryCheckAttrTypeNum, g_usedLanguage).getValue() == primaryCheckAttrValue)
                    bAttrCondition = true; 
                break;
            case String("IsNotEqual"):
                if (primaryCheckItem.Attribute(primaryCheckAttrTypeNum, g_usedLanguage).getValue() != primaryCheckAttrValue)
                    bAttrCondition = true; 
                break; 
            case String("Contains"):
               var allowedValues = primaryCheckAttrValue.split(","); 
               for (var cou=0; cou<allowedValues.length; cou++)
               {
                   if (primaryCheckItem.Attribute(primaryCheckAttrTypeNum, g_usedLanguage).getValue() == String(allowedValues[cou]))
                   {
                       bAttrCondition = true; 
                       break; 
                   }
               }
               break; 
           default:
               Dialogs.MsgBox(getString("ERR_WRONGCOMPTYPE") + String(primaryCheckCompType), Constants.MSGBOX_BTN_OK, "Error reading xml"); 
        }
    }
    else // no comparison type defined and no dependentValue defined -> check only if the attribute is maintained
    {
        if (primaryCheckItem.Attribute(primaryCheckAttrTypeNum, g_usedLanguage).IsMaintained())
            bAttrCondition = true; 
    }
    
    if  (bAttrCondition)
    {
        if (secondaryCheckAttrValue == null)//if the attr Value is not defined, we check the attribute for maintenance
        {
            if (!(secondaryCheckItem.Attribute(secondaryCheckAttrTypeNum, g_usedLanguage).IsMaintained()))
                bError = true; 
        }
        else
        {
            if (secondaryCheckItem.Attribute(secondaryCheckAttrTypeNum, g_usedLanguage).IsMaintained())
            {   
                switch (String(secondaryCheckCompType)) // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
                {
                    case String("IsEqual"): // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
                        if (secondaryCheckItem.Attribute(secondaryCheckAttrTypeNum, g_usedLanguage).getValue() != secondaryCheckAttrValue)
                            bError = true; 
                        break; 
                    case String("IsNotEqual"):
                        if (secondaryCheckItem.Attribute(secondaryCheckAttrTypeNum, g_usedLanguage).getValue() == secondaryCheckAttrValue)
                            bError = true; 
                        break; 
                    default: //secondaryCheckCompType = "" --> condition is true, if the attribute is maintained -> no Error
                        bError= false; 
                }
            }
            else
            {
                //if an attr Value is needed, we have an error, no further check needed
                if (secondaryCheckAttrValue != "" && secondaryCheckCompType == "IsEqual")
                    bError = true;      
            }
        }
    }
    
    if (secondaryCheckItem.KindNum() == Constants.CID_OBJDEF)
        errorMessage = errorMessage.replace("{0}", secondaryCheckItem.Name(g_usedLanguage)).replace("{1}", secondaryCheckItem.Type()); 
    else if (secondaryCheckItem.KindNum() == Constants.CID_CXNDEF)
        errorMessage = errorMessage.replace("{0}", secondaryCheckItem.SourceObjDef().Name(g_usedLanguage)).replace("{1}", secondaryCheckItem.TargetObjDef().Name(g_usedLanguage)); 
    else if(secondaryCheckItem.KindNum() == Constants.CID_MODEL)
    {
        if (secondaryCheckItem.TypeNum() == Constants.MT_BUSY_CONTR_DGM)
            errorMessage = errorMessage.replace("{0}", secondaryCheckItem.Name(g_usedLanguage)); 
    }   
    
    if (bError)
        LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errorMessage, checkCategory);
    else
        LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errorMessage + " - OK", 5);
}

//pj1:new way to handle custom complex attribute checks
function checkCustomAttr(currentItem, attrTypeNum, checkCategory, errorMessage, ruleShortcut, ruleDescription, custCheck)
{
    var bError = false;
    // SP-XX
    // var attrName = getAttrNum(currentItem, ruleShortcut, ruleDescription, attrTypeNum)
    
    var attrName  = currentItem.Attribute(attrTypeNum, g_usedLanguage).Type(); 
    var custCheck = ("" + custCheck).split("|");
    var logType   = getLogTypeForAttributeItemKind(currentItem);

    if (currentItem.KindNum() == Constants.CID_OBJDEF)
        errorMessage = errorMessage.replace("{0}", "'" + currentItem.Name(g_usedLanguage) + "'").replace("{1}", currentItem.Type()) ; 
       
    var test = custCheck[0]; // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
    
    switch(test){
        case String("IsMaintainedAndAdditionallyMainained"): // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
            if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() != "")
            {
                var targetAttr = getAttrNum(currentItem, ruleShortcut, ruleDescription, custCheck[1]);
                if(currentItem.Attribute(targetAttr, g_usedLanguage).IsMaintained() == false || currentItem.Attribute(targetAttr, g_usedLanguage).getValue() == "")
                {
                    bError = true;
                }
            }
            break;
        case String("IsMaintainedAndAdditionallyMainainedOneofAll"):
            if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() != "")
            {
                var isfound = false;
                //if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() != "")
                //{
                    for(var i = 1; i < custCheck.length; i++)
                    {
                        var targetAttr = getAttrNum(currentItem, ruleShortcut, ruleDescription, custCheck[i]);
                        if(currentItem.Attribute(targetAttr, g_usedLanguage).IsMaintained() == true)
                        {
                            isfound = true;
                            break;
                        }
                    }
                //}
                if(!isfound){
                    bError = true;
                }
            }
            break;
        case String("IsMaintainedAndAdditionallyNotMaintained"):
            if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() != "")
            {
                var targetAttr = getAttrNum(currentItem, ruleShortcut, ruleDescription, custCheck[1]);
                if(currentItem.Attribute(targetAttr, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(targetAttr, g_usedLanguage).getValue() != "")
                {
                    bError = true;
                }
            }
            break;
        case String("IsMaintainedWithValueAndAdditionallyMainained"):
            if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() != "")
            {
                var expectedValue = custCheck[1]
                var targetAttr = getAttrNum(currentItem, ruleShortcut, ruleDescription, custCheck[2]);
                if(currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() == expectedValue){
                    if(currentItem.Attribute(targetAttr, g_usedLanguage).IsMaintained() == false || currentItem.Attribute(targetAttr, g_usedLanguage).getValue() == "")
                    {
                        bError = true;
                    }
                }
            }
            break;
        case String("IsMaintainedWithValueAndAdditionallyNotMainained"):
            if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() != "")
            {
                var expectedValue = custCheck[1]
                var targetAttr = getAttrNum(currentItem, ruleShortcut, ruleDescription, custCheck[2]);
                if(currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() == expectedValue){
                    if(currentItem.Attribute(targetAttr, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(targetAttr, g_usedLanguage).getValue() != "")
                    {
                        bError = true;
                    }
                }
            }
            break;
        // ----
        case String("IsNotMaintainedAndAdditionallyMainained"):
            if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == false || currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() == "")
            {
                var targetAttr = getAttrNum(currentItem, ruleShortcut, ruleDescription, custCheck[1]);
                if(currentItem.Attribute(targetAttr, g_usedLanguage).IsMaintained() == false || currentItem.Attribute(targetAttr, g_usedLanguage).getValue() == "")
                {
                    if(currentItem.Attribute(targetAttr, g_usedLanguage).IsMaintained() == false || currentItem.Attribute(targetAttr, g_usedLanguage).getValue() == "")
                    {
                        bError = true;
                    }
                }
            }
            break;
        case String("IsNotMaintainedAndAdditionallyMainainedOneOfAll"):
            if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == false || currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() == "")
            {
                var isfound = false;
                if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() != "")
                {
                    for(var i = 1; i < custCheck.length; i++){
                        var targetAttr = getAttrNum(currentItem, ruleShortcut, ruleDescription, custCheck[i]);
                        if(currentItem.Attribute(targetAttr, g_usedLanguage).IsMaintained() == true){
                            isfound = true;
                            break;
                        }
                    }
                }
                if(!isfound){
                    bError = true;
                }
            }
            break;
        case String("IsNotMaintainedAndAdditionallyNotMainained"):
            if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == false || currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() == "")
            {
                // SP-XX
                // var expectedValue = custCheck[1]
                var targetAttr = getAttrNum(currentItem, ruleShortcut, ruleDescription, custCheck[2]);
                if(currentItem.Attribute(targetAttr, g_usedLanguage).IsMaintained() == false || currentItem.Attribute(targetAttr, g_usedLanguage).getValue() == "")
                {
                    if(currentItem.Attribute(targetAttr, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(targetAttr, g_usedLanguage).getValue() != "")
                    {
                        bError = true;
                    }
                }
            }
            break;
        case String("checkRiskLawType"):
            if(currentItem.Attribute(attrTypeNum, g_usedLanguage).IsMaintained() == true && currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() != "")
            {
                var expectedValue = custCheck[1]
                var targetAttr = getAttrNum(currentItem, ruleShortcut, ruleDescription, custCheck[2]);
                if(currentItem.Attribute(attrTypeNum, g_usedLanguage).getValue() == expectedValue)
                {
                    var targetVal = "";
                    if(expectedValue == "Att.Verf. LAW262")
                    {
                        targetVal = "Ja";
                    }

                    risks = currentItem.getConnectedObjs([Constants.OT_RISK]);
                    risks.forEach(
                        function(curRisk)
                        {
                            if(curRisk.Attribute(targetAttr, g_usedLanguage).getValue() != custCheck[3])
                            {
                                bError = true;
                            }
                        }
                    )
                }
            }
            break;
        default:
             Dialogs.MsgBox(getString("ERR_WRONGATTRIBUTECKTYPE") + checkType, Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION"));
            break;
    }
        
    if (bError)
        LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errorMessage, checkCategory); 
    else
        LogMessage(logType, ruleShortcut, ruleDescription, GetText("TXT_ATTRIBUTE", new Array(attrName)), errorMessage + " - OK", 5); 
       
    function getAttrNum(affectedItem, affectedRuleShortcut, affectedRuleDescription, attrId)
    {
        var attrTypeNum = 0; 
        if (attrId != null)
        {
            if (vbIsNumeric(attrId)> 0)
                attrTypeNum = __toLong(attrId); 
            else if (vbLen(attrId) == 36) //GUID
                attrTypeNum = g_methodFilter.UserDefinedAttributeTypeNum(attrId); 

            if (attrTypeNum == 0)
                LogMessage(getLogTypeForAttributeItemKind(affectedItem), affectedRuleShortcut, affectedRuleDescription, attrId, "Kann Attribut Typ nicht ermitteln.", 1); 
        }
        else
        {
            var logTypeForItemKind = "UNKNOWN";
            try
            {
                logTypeForItemKind = getLogTypeForAttributeItemKind(affectedItem);
            }
            catch (err)
            {
            }
            LogMessage(logTypeForItemKind, affectedRuleShortcut, affectedRuleDescription, attrId, "Kein Attribut Typ angegeben.", 1); 
        }
        
        return attrTypeNum; 
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///  Conventions check for EPCStructure
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//check all rules with ruleKind = "EPCStructure"
//parameters: 
// epcRules: The XML node containing the collection epc rules.
// currentModel: The model to check.
function doConventionsCheckEPCStructure(epcRules, currentModel)
{
   for (var cou = 0; cou < epcRules.size(); cou++)
   {
        var ruleXmlNode = epcRules.get(cou); 
        var strucRuleXmlNode = g_config.getStructureRule(ruleXmlNode); 
        doStructureCheck(currentModel, ruleXmlNode, strucRuleXmlNode); 
    }
}

//Check the structure constraints, containend in node StructureRule.
//parameters: 
// currentModel: The model to check.
// ruleXmlNode: The XML node containing a single epc rule's rule data.
// strucRuleXmlNode: The XML sub-node, containing the rurrent rule's structure check data.
function doStructureCheck(currentModel, ruleXmlNode, strucRuleXmlNode)
{
    var checkItemKind = "";
    var checkTypeList = "";
    var checkCategory=""; 
    var ruleDescription =""; 
    var errorMessage =""; 
    var checkType="";
    var checkCorrelation="";
    var ruleShortcut = ""; 
    
    if (strucRuleXmlNode != null)
    {
        checkItemKind    = ruleXmlNode.getAttributeValue("kind");
        checkTypeList    = ruleXmlNode.getAttributeValue("typelist");
        checkTypeFilter  = ruleXmlNode.getAttributeValue("typefilter");
        checkCategory    = g_config.getCheckCategory(ruleXmlNode);
        ruleDescription   = g_config.getRuleDescription(ruleXmlNode);
        errorMessage     = g_config.getErrorMessage(ruleXmlNode); 
        checkType        = String(g_config.getRuleType(strucRuleXmlNode)); // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
        checkCorrelation = String(g_config.getRuleCorrelation(strucRuleXmlNode));
        ruleShortcut     = g_config.getRuleShortcut(ruleXmlNode); 
        
        if (checkCategory <= g_currentLogType)
        {
            switch (checkType)
            {
            case String("ConnectionCountOnEachElement"): // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
                checkConnectionCountOnEachElement(currentModel, strucRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCorrelation, checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                break;
            case String("ConnectionCountOnElementSequence"):
                checkConnectionCountOnElementSequence(currentModel, strucRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCorrelation, checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                break;
            case String("AssignmentCount"):
                checkAssignmentCount(currentModel, strucRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCorrelation, checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                break;
            default:
                Dialogs.MsgBox(getString("ERR_WRONGSTRUCTCHECKTYPE") + checkType, Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION"));
                break;
            }
        }
    }
    else
    {       
        Dialogs.MsgBox(GetText("ERR_RULEINCOMPLETE", new Array(g_config.getRuleName(ruleXmlNode))), Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION")) ; 
    }
}

//Filter the 'items of indicated model' that are matching the constraints.
// currentModel: The model to filter items from.
// checkItemKind: The item kind to apply on checkTypeList filter (CID_OBJOCC or CID_OBJDEF).
// checkTypeList: The list of types to filter for (Item-Number or Symbol-Number).
// checkTypeFilter: The kind of filter to apply (white/black).
// returns: The array of filtered model items.
function helperFilterModelItems(currentModel, checkItemKind, checkTypeList, checkTypeFilter)
{
    // Prepare type numbers to check.
    var checkTypes;
    if (checkTypeList.indexOf("*") >= 0)
        checkTypes = ["*"];
    else if (checkTypeList.indexOf(",") >= 0)
        checkTypes = checkTypeList.split(",");
    else
        checkTypes = [checkTypeList];
    
    // Convert type GUIDs to type numbers.
    for (var checkTypeIndex = 0; checkTypeIndex < checkTypes.length; checkTypeIndex++)
        checkTypes[checkTypeIndex] = helperConvertTypeGuidToTypeNum(checkTypes[checkTypeIndex], checkItemKind);
    
    // Filter items to check.
    var filteredItems = new Array();
    for (var typeIndex = 0; typeIndex < checkTypes.length; typeIndex++)
    {
        var matchingItems = null;
        
        if (checkTypeFilter == "black")
        {
            var itemsToExclude = new Array();
            
            if (checkTypes[typeIndex] == "*")
                matchingItems = currentModel.ObjOccList();
            else
            {
                if (checkItemKind == "CID_OBJOCC")
                    matchingItems = currentModel.ObjOccListFilter(-1, checkTypes[typeIndex]);
                else // (checkItemKind == "CID_OBJDEF")
                    matchingItems = currentModel.ObjOccListFilter(checkTypes[typeIndex]);
            }
            for (var itemIdx = 0; itemIdx < matchingItems.length; itemIdx++)
                itemsToExclude.push(matchingItems[itemIdx]);
            
            var itemsToReduce = currentModel.ObjOccList();
            
            for (var intReduceIdx = 0; intReduceIdx < itemsToReduce.length; intReduceIdx++)
            {
                var addItem = true;
                itemsToReduce[intReduceIdx]
                for (var intExcludeIdx = 0; intExcludeIdx < itemsToExclude.length; intExcludeIdx++)
                {
                    if (itemsToReduce[intReduceIdx] == itemsToExclude[intExcludeIdx])
                    {
                        addItem = false;
                        break;
                    }
                }
                if (addItem == true)
                    filteredItems.push(itemsToReduce[intReduceIdx]);
            }
        }
        else
        {
            if (checkTypes[typeIndex] == "*")
                matchingItems = currentModel.ObjOccList();
            else
            {
                if (checkItemKind == "CID_OBJOCC")
                    matchingItems = currentModel.ObjOccListFilter(-1, checkTypes[typeIndex]);
                else // (checkItemKind == "CID_OBJDEF")
                    matchingItems = currentModel.ObjOccListFilter(checkTypes[typeIndex]);
            }
            for (var itemIdx = 0; itemIdx < matchingItems.length; itemIdx++)
                filteredItems.push(matchingItems[itemIdx]);
        }
    }
    return filteredItems;
}

//Convert the 'items to process' into unique object definition array.
// itemsToProcess: The items to malke unique and convert to object definition.
// itemKind: The item kind of the 'itemsToProcess' (CID_OBJOCC or CID_OBJDEF).
// returns: The unique array object definitions.
function helperToUniqueObjDefArray(itemsToProcess, itemKind)
{
    // Make items to process unique and convert ObjOcc to ObjDef, if required.
    var itemDefs = new Array();
    if (itemKind == "CID_OBJOCC")
    {
        for (var itemIndex = 0; itemIndex < itemsToProcess.length; itemIndex++)
        {
            var itemDef = itemsToProcess[itemIndex].ObjDef();
            var contained = false;
            for (var itemDefIndex = 0; itemDefIndex < itemDefs.length; itemDefIndex++)
            {
                if (itemDefs[itemDefIndex].GUID() == itemDef.GUID())
                {
                    contained = true;
                    break;
                }
            }
            if (contained == false)
                itemDefs.push(itemDef);
        }
    }
    else
    {
        for (var itemIndex = 0; itemIndex < itemsToProcess.length; itemIndex++)
        {
            itemDefs.push(itemsToProcess[itemIndex]);
        }
    }
    
    return itemDefs;
}

//Check the 'number of connection' constraints for each element matching the checkItemKind/checkTypeList/checkTypeFilter constraints.
// currentModel: The model to check.
// strucRuleXmlNode: The XML sub-node, containing the structure check.
// checkItemKind: The item kind to apply on checkTypeList filter (CID_OBJOCC or CID_OBJDEF).
// checkTypeList: The list of types to filter for (Item-Number or Symbol-Number).
// checkTypeFilter: The kind of filter to apply (white/black).
// checkCorrelation: The definition how the results of multiple checks on the same item are to be aggregated (fallback is 'and').
// checkCategory: The the definition of the failure effect.
// ruleDescription: The rule description text.
// errorMessage: The error message text.
function checkConnectionCountOnEachElement(currentModel, strucRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCorrelation, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var errorFoundInModel=false; 
    
    // Collect items to check.
    var itemsToCheck = helperFilterModelItems(currentModel, checkItemKind, checkTypeList, checkTypeFilter);
    
    var testNodes = strucRuleXmlNode.getChildren ("ConnectionCounts");
    for (var itemIndex = 0; itemIndex < itemsToCheck.length; itemIndex++)
    {
        var resultCheck = true;
        if (checkCorrelation == "or")
            resultCheck = false;
        else
            resultCheck = true;

        for (var testIdx = 0; testIdx < testNodes.size(); testIdx++)
        {
            var inNum  = testNodes.get(testIdx).getAttributeValue("in");
            var outNum = testNodes.get(testIdx).getAttributeValue("out");
            var testCorrelation = testNodes.get(testIdx).getAttributeValue("correlation");
            var testRelationship = testNodes.get(testIdx).getAttributeValue("relationship");
            
            var inoutRelationship = Constants.EDGES_ALL;
            if (testRelationship == "structure")
                inoutRelationship = Constants.EDGES_STRUCTURE;
            if (testRelationship == "nonstructure")
                inoutRelationship = Constants.EDGES_NONSTRUCTURE;
                
            var inReal  = itemsToCheck[itemIndex].Cxns(Constants.EDGES_IN, inoutRelationship).length;
            var outReal = itemsToCheck[itemIndex].Cxns(Constants.EDGES_OUT, inoutRelationship).length;
            
            var subResultIn = false;
            if ((inNum == "0" || inNum == "0...1" || inNum == "0...n") && inReal == 0)
                subResultIn = true;
            if ((inNum == "0...1" || inNum == "0...n" || inNum == "1" || inNum == "1...n") && inReal == 1)
                subResultIn = true;
            if ((inNum == "0...n" || inNum == "1...n" || inNum == "2...n") && inReal > 1)
                subResultIn = true;
            
            var subResultOut = false;
            if ((outNum == "0" || outNum == "0...1" || outNum == "0...n") && outReal == 0)
                subResultOut = true;
            if ((outNum == "0...1" || inNum == "0...n" || outNum == "1" || outNum == "1...n") && outReal == 1)
                subResultOut = true;
            if ((outNum == "0...n" || outNum == "1...n" || outNum == "2...n") && outReal > 1)
                subResultOut = true;
            
            var subResultTest = false;
            if (testCorrelation == "and")
                subResultTest = (subResultIn == true && subResultOut == true ? true : false);
            else
                subResultTest = (subResultIn == true || subResultOut == true ? true : false);

            if (checkCorrelation == "or")
                resultCheck = (resultCheck == true || subResultTest == true ? true : false);               
            else
                resultCheck = (resultCheck == true && subResultTest == true ? true : false);
        }
        
        //pj1:ignore unconnected objects above first event
        if(!resultCheck){
            if(itemsToCheck[itemIndex].ObjDef().TypeNum() == Constants.OT_DOC_KNWLDG){
                var connections = itemsToCheck[itemIndex].CxnOccList();
                if(connections.length == 0){
                    var eventList = itemsToCheck[itemIndex].Model().ObjOccListFilter(Constants.OT_EVT);
                    
                    eventList.sort(function(a, b)
                    {
                        return a.Y() - b.Y();
                    });
                    
                    if(eventList.length > 0){
                        if(itemsToCheck[itemIndex].Y() < eventList[0].Y()){
                            resultCheck = true;
                        }
                    }
                }
            }
        }
        
        if (resultCheck == false)
        {
            var itemDef    = itemsToCheck[itemIndex].ObjDef();
            var itemName   = itemDef.Name(g_usedLanguage, true);
            
            var predecCxn  = itemsToCheck[itemIndex].Cxns(Constants.EDGES_IN, Constants.EDGES_STRUCTURE);
            var predecName = (predecCxn != null && predecCxn.length > 0 ? predecCxn[0].SourceObjOcc().ObjDef().Name(g_usedLanguage, true) : "-");
            var succCxn    = itemsToCheck[itemIndex].Cxns(Constants.EDGES_OUT, Constants.EDGES_STRUCTURE);
            var succName   = (succCxn != null && succCxn.length > 0 ? succCxn[0].TargetObjOcc().ObjDef().Name(g_usedLanguage, true) : "-");
            
            LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CONNECTIONCOUNTCHECK"),
                errorMessage.replace("{0}", itemDef.Type()).replace("{1}", itemName).replace("{2}", predecName).replace("{3}", succName),
                checkCategory); 
            errorFoundInModel = true; 
        }
    }
    if (!(errorFoundInModel))
       LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CONNECTIONCOUNTCHECK"), ruleDescription + " - OK", 5); 

}

//Check the 'number of connection' constraints for each element matching the checkItemKind/checkTypeList/checkTypeFilter constraints.
// currentModel: The model to check.
// strucRuleXmlNode: The XML sub-node, containing the structure check.
// checkItemKind: The item kind to apply on checkTypeList filter (CID_OBJOCC or CID_OBJDEF).
// checkTypeList: The list of types to filter for (Item-Number or Symbol-Number).
// checkTypeFilter: The kind of filter to apply (white/black).
// checkCorrelation: The definition how the results of multiple checks on the same item are to be aggregated (fallback is 'and').
// checkCategory: The the definition of the failure effect.
// ruleDescription: The rule description text.
// errorMessage: The error message text.
function checkConnectionCountOnElementSequence(currentModel, strucRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCorrelation, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    
    var errorFoundInModel=false; 
    
    // Prepare type numbers to check.
    var checkTypeSequences;
    if (checkTypeList.indexOf(",") >= 0)
        checkTypeSequences = checkTypeList.split(",");
    else
        checkTypeSequences = [checkTypeList];

    for (var sequenceIdx = 0; sequenceIdx < checkTypeSequences.length; sequenceIdx++)
    {
        var checkTypes;
        if (checkTypeList.indexOf(" ") >= 0)
            checkTypes = checkTypeSequences[sequenceIdx].split(" ");
        else
            checkTypes = [checkTypeSequences[sequenceIdx]];
    
        // Convert type GUIDs to type numbers.
        for (var checkTypeIndex = 0; checkTypeIndex < checkTypes.length; checkTypeIndex++)
            checkTypes[checkTypeIndex] = helperConvertTypeGuidToTypeNum(checkTypes[checkTypeIndex], checkItemKind);
        
        if (checkTypes.length != 2)
        {
            Dialogs.MsgBox(GetText("ERR_RULEINCOMPLETE", new Array(ruleDescription)), Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION"));
            return;
        }
    
        // Collect items to check.
        var itemsToCheck = new Array();
        var matchingItems = null;
        
        if (checkTypeFilter == "black")
        {
            var itemsToExclude = new Array();
            
            if (checkTypes[0] == "*")
                matchingItems = currentModel.ObjOccList();
            else
            {
                if (checkItemKind == "CID_OBJOCC")
                    matchingItems = currentModel.ObjOccListFilter(-1, checkTypes[0]);
                else // (checkItemKind == "CID_OBJDEF")
                    matchingItems = currentModel.ObjOccListFilter(checkTypes[0]);
            }
            for (var itemIdx = 0; itemIdx < matchingItems.length; itemIdx++)
                itemsToExclude.push(matchingItems[itemIdx]);
            
            var itemsToReduce = currentModel.ObjOccList();
            
            for (var intReduceIdx = 0; intReduceIdx < itemsToReduce.length; intReduceIdx++)
            {
                var addItem = true;
                itemsToReduce[intReduceIdx]
                for (var intExcludeIdx = 0; intExcludeIdx < itemsToExclude.length; intExcludeIdx++)
                {
                    if (itemsToReduce[intReduceIdx] == itemsToExclude[intExcludeIdx])
                    {
                        addItem = false;
                        break;
                    }
                }
                if (addItem != true)
                    continue;
                
                var connectedItems = itemsToReduce[intReduceIdx].getConnectedObjOccs([parseInt(checkTypes[1])], Constants.EDGES_OUT);
                if (connectedItems.length > 0)
                    itemsToCheck.push(itemsToReduce[intReduceIdx]);
            }
        }
        else
        {
            if (checkTypes[0] == "*")
                matchingItems = currentModel.ObjOccList();
            else
            {
                if (checkItemKind == "CID_OBJOCC")
                    matchingItems = currentModel.ObjOccListFilter(-1, checkTypes[0]);
                else // (checkItemKind == "CID_OBJDEF")
                    matchingItems = currentModel.ObjOccListFilter(checkTypes[0]);
            }
            for (var itemIdx = 0; itemIdx < matchingItems.length; itemIdx++)
            {
                var connectedItems = matchingItems[itemIdx].getConnectedObjOccs([parseInt(checkTypes[1])], Constants.EDGES_OUT);
                if (connectedItems.length > 0)
                    itemsToCheck.push(matchingItems[itemIdx]);
            }
        }
        
        var testNodes = strucRuleXmlNode.getChildren ("ConnectionCounts");
        for (var itemIndex = 0; itemIndex < itemsToCheck.length; itemIndex++)
        {
            var resultCheck = true;
            if (checkCorrelation == "or")
                resultCheck = false;
            else
                resultCheck = true;

            var inNumStart  = testNodes.get(0).getAttributeValue("in");
            var outNumStart = testNodes.get(0).getAttributeValue("out");
            var testCorrelationStart = testNodes.get(0).getAttributeValue("correlation");
            var testRelationshipStart = testNodes.get(0).getAttributeValue("relationship");
            
            var inoutRelationshipStart = Constants.EDGES_ALL;
            if (testRelationshipStart == "structure")
                inoutRelationshipStart = Constants.EDGES_STRUCTURE;
            if (testRelationshipStart == "nonstructure")
                inoutRelationshipStart = Constants.EDGES_NONSTRUCTURE;
                
            var inRealStart  = itemsToCheck[itemIndex].Cxns(Constants.EDGES_IN, inoutRelationshipStart).length;
            var outRealStart = itemsToCheck[itemIndex].Cxns(Constants.EDGES_OUT, inoutRelationshipStart).length;
            
            var subResultInStart = false;
            if ((inNumStart == "0" || inNumStart == "0...1" || inNumStart == "0...n") && inRealStart == 0)
                subResultInStart = true;
            if ((inNumStart == "0...1" || inNumStart == "0...n" || inNumStart == "1" || inNumStart == "1...n") && inRealStart == 1)
                subResultInStart = true;
            if ((inNumStart == "0...n" || inNumStart == "1...n" || inNumStart == "2...n") && inRealStart > 1)
                subResultInStart = true;
            
            var subResultOutStart = false;
            if ((outNumStart == "0" || outNumStart == "0...1" || outNumStart == "0...n") && outRealStart == 0)
                subResultOutStart = true;
            if ((outNumStart == "0...1" || inNumStart == "0...n" || outNumStart == "1" || outNumStart == "1...n") && outRealStart == 1)
                subResultOutStart = true;
            if ((outNumStart == "0...n" || outNumStart == "1...n" || outNumStart == "2...n") && outRealStart > 1)
                subResultOutStart = true;
            
            var subResultTestStart = false;
            if (testCorrelationStart == "and")
                subResultTestStart = (subResultInStart == true && subResultOutStart == true ? true : false);
            else
                subResultTestStart = (subResultInStart == true || subResultOutStart == true ? true : false);

            if (checkCorrelation == "or")
                resultCheck = (resultCheck == true || subResultTestStart == true ? true : false);               
            else
                resultCheck = (resultCheck == true && subResultTestStart == true ? true : false);
            
            var connectedItems = itemsToCheck[itemIndex].getConnectedObjOccs([parseInt(checkTypes[1])], Constants.EDGES_OUT);
            for (var connectedItemIndex = 0; connectedItemIndex < connectedItems.length; connectedItemIndex++)
            {
                var inNumEnd  = testNodes.get(1).getAttributeValue("in");
                var outNumEnd = testNodes.get(1).getAttributeValue("out");
                var testCorrelationEnd = testNodes.get(1).getAttributeValue("correlation");
                var testRelationshipEnd = testNodes.get(1).getAttributeValue("relationship");
                
                var inoutRelationshipEnd = Constants.EDGES_ALL;
                if (testRelationshipEnd == "structure")
                    inoutRelationshipEnd = Constants.EDGES_STRUCTURE;
                if (testRelationshipEnd == "nonstructure")
                    inoutRelationshipEnd = Constants.EDGES_NONSTRUCTURE;
                    
                var inRealEnd  = connectedItems[connectedItemIndex].Cxns(Constants.EDGES_IN, inoutRelationshipEnd).length;
                var outRealEnd = connectedItems[connectedItemIndex].Cxns(Constants.EDGES_OUT, inoutRelationshipEnd).length;
            
                var subResultInEnd = false;
                if ((inNumEnd == "0" || inNumEnd == "0...1" || inNumEnd == "0...n") && inRealEnd == 0)
                    subResultInEnd = true;
                if ((inNumEnd == "0...1" || inNumEnd == "0...n" || inNumEnd == "1" || inNumEnd == "1...n") && inRealEnd == 1)
                    subResultInEnd = true;
                if ((inNumEnd == "0...n" || inNumEnd == "1...n" || inNumEnd == "2...n") && inRealEnd > 1)
                    subResultInEnd = true;
                
                var subResultOutEnd = false;
                if ((outNumEnd == "0" || outNumEnd == "0...1" || outNumEnd == "0...n") && outRealEnd == 0)
                    subResultOutEnd = true;
                if ((outNumEnd == "0...1" || inNumEnd == "0...n" || outNumEnd == "1" || outNumEnd == "1...n") && outRealEnd == 1)
                    subResultOutEnd = true;
                if ((outNumEnd == "0...n" || outNumEnd == "1...n" || outNumEnd == "2...n") && outRealEnd > 1)
                    subResultOutEnd = true;
                
                var subResultTestEnd = false;
                if (testCorrelationEnd == "and")
                    subResultTestEnd = (subResultInEnd == true && subResultOutEnd == true ? true : false);
                else
                    subResultTestEnd = (subResultInEnd == true || subResultOutEnd == true ? true : false);
    
                var combinationResult = false;
                if (checkCorrelation == "or")
                    combinationResult = (resultCheck == true || subResultTestEnd == true ? true : false);               
                else
                    combinationResult = (resultCheck == true && subResultTestEnd == true ? true : false);
                
                if (combinationResult == false)
                {
                    var itemDef    = itemsToCheck[itemIndex].ObjDef();
                    var itemName   = itemDef.Name(g_usedLanguage, true);
                    
                    var predecCxn  = itemsToCheck[itemIndex].Cxns(Constants.EDGES_IN, Constants.EDGES_STRUCTURE);
                    var predecName = (predecCxn != null && predecCxn.length > 0 ? predecCxn[0].SourceObjOcc().ObjDef().Name(g_usedLanguage, true) : "-");
                    var succCxn    = itemsToCheck[itemIndex].Cxns(Constants.EDGES_OUT, Constants.EDGES_STRUCTURE);
                    var succName   = (succCxn != null && succCxn.length > 0 ? succCxn[0].TargetObjOcc().ObjDef().Name(g_usedLanguage, true) : "-");
                    
                    LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CONNECTIONCOUNTCHECK"),
                        errorMessage.replace("{0}", itemDef.Type()).replace("{1}", itemName).replace("{2}", predecName).replace("{3}", succName),
                        checkCategory); 
                        
                    errorFoundInModel = true; 
                }
               
            }
        }
    }
    if (!(errorFoundInModel))
       LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CONNECTIONCOUNTCHECK"), ruleDescription + " - OK", 5); 
}

//Check the 'number of assignment' constraints for each element matching the checkItemKind/checkTypeList/checkTypeFilter constraints.
// currentModel: The model to check.
// strucRuleXmlNode: The XML sub-node, containing the structure check.
// checkItemKind: The item kind to apply on checkTypeList filter (CID_OBJOCC or CID_OBJDEF).
// checkTypeList: The list of types to filter for (Item-Number or Symbol-Number).
// checkTypeFilter: The kind of filter to apply (white/black).
// checkCorrelation: The definition how the results of multiple checks on the same item are to be aggregated (fallback is 'and').
// checkCategory: The the definition of the failure effect.
// ruleDescription: The rule description text.
// errorMessage: The error message text.
function checkAssignmentCount(currentModel, strucRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCorrelation, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var errorFoundInModel = false; 
    
    // Collect items to check. We assume 'checkItemKind' is 'CID_OBJOCC'.
    var itemsToCheck = helperFilterModelItems(currentModel, checkItemKind, checkTypeList, checkTypeFilter);
    
    var testNodes     = strucRuleXmlNode.getChildren ("AssignmentCounts");
    var requestedNum  = testNodes.get(0).getAttributeValue("number");
    var type          = testNodes.get(0).getAttributeValue("type");
    var checkEvents   = testNodes.get(0).getAttributeValue("checkEvents");
    var checkReverse  = testNodes.get(0).getAttributeValue("checkReverseInterface");

    // Make items to check unique and convert ObjOcc to ObjDef, if required.
    var itemDefsToCheck = helperToUniqueObjDefArray(itemsToCheck, checkItemKind);
    itemDefsToCheck = helperClearOutInterfaceCollections(itemDefsToCheck);
    
    for (var itemIndex = 0; itemIndex < itemDefsToCheck.length; itemIndex++)
    {
        var numOfAssignments = itemDefsToCheck[itemIndex].AssignedModels(type).length;
        
        var result= false;
        if ((requestedNum == "0" || requestedNum == "0...1" || requestedNum == "0...n") && numOfAssignments == 0)
            result = true;
        if ((requestedNum == "0...1" || requestedNum == "0...n" || requestedNum == "1" || requestedNum == "1...n") && numOfAssignments == 1)
            result = true;
        if ((requestedNum == "0...n" || requestedNum == "1...n" || requestedNum == "2...n") && numOfAssignments > 1)
            result = true;

        var itemNameForException   = itemDefsToCheck[itemIndex].Name(g_usedLanguage, true);
        if(itemNameForException == "Prozessende" || itemNameForException == "End of Process" || itemNameForException == "Neustart Prozess" || itemNameForException == "Re-start process")
            result = true;  // [SM] weitere Ausnahmen hinzugefügt
        
        if (result == false)
        {
            var itemDef    = itemDefsToCheck[itemIndex];
            var itemName   = itemDef.Name(g_usedLanguage, true);
            
            LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_ASSIGNMENTCOUNTCHECK"),
                errorMessage.replace("{0}", itemDef.Type()).replace("{1}", itemName), checkCategory); 
        }
    }
    if (checkEvents == "true")
    {
        for (var itemIndex = 0; itemIndex < itemDefsToCheck.length; itemIndex++)
        {
            // Recall interface occs from interface def and check predecessor/successor events.
            var currentInerfaceOccsToCheck = itemDefsToCheck[itemIndex].OccList([currentModel]);
            
            for (var currentInerfaceOccIndex = 0; currentInerfaceOccIndex < currentInerfaceOccsToCheck.length; currentInerfaceOccIndex++)
            {
                // Only recalled interface occs from current model are to investigate.
                if (currentInerfaceOccsToCheck[currentInerfaceOccIndex].Model().GUID() != currentModel.GUID())
                    continue;
                
                var assignedModels = itemDefsToCheck[itemIndex].AssignedModels(type);
                var predecessorEvents = currentInerfaceOccsToCheck[currentInerfaceOccIndex].getConnectedObjOccs ([Constants.ST_EV], Constants.EDGES_IN);
                var successorEvents   = currentInerfaceOccsToCheck[currentInerfaceOccIndex].getConnectedObjOccs ([Constants.ST_EV], Constants.EDGES_OUT);
            
                for (var assignedIndex = 0; assignedIndex < assignedModels.length; assignedIndex++)
                {
                    var predecessorMatch = true;
                    var successorMatch   = true;
                    var mismatchObjDef   = null;
                    
                    for (var predecessorIdx = 0; predecessorIdx < predecessorEvents.length; predecessorIdx++)
                    {
                        if (predecessorEvents[predecessorIdx].ObjDef().OccList([assignedModels[assignedIndex]]).length < 1)
                        {
                            mismatchObjDef = predecessorEvents[predecessorIdx].ObjDef();
                            predecessorMatch = false;
                        }
                    }
                    if (predecessorMatch == false)
                    {
                        var itemDef      = itemDefsToCheck[itemIndex];
                        var itemName     = itemDef.Name(g_usedLanguage, true);
                        var mismatchName = (mismatchObjDef != null ? mismatchObjDef.Name(g_usedLanguage, true) : "-");
                        
                        LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_ASSIGNMENTCOUNTCHECK"),
                            errorMessage.replace("{0}", itemDef.Type()).replace("{1}", itemName).replace("{2}", mismatchName), checkCategory);
                    }

                    for (var successorIdx = 0; successorIdx < successorEvents.length; successorIdx++)
                    {
                        if (successorEvents[successorIdx].ObjDef().OccList([assignedModels[assignedIndex]]).length < 1)
                        {
                            mismatchObjDef = successorEvents[successorIdx].ObjDef();
                            successorMatch = false;
                        }
                    }
                    if (successorMatch   == false)
                    {
                        var itemDef      = itemDefsToCheck[itemIndex];
                        var itemName     = itemDef.Name(g_usedLanguage, true);
                        var mismatchName = (mismatchObjDef != null ? mismatchObjDef.Name(g_usedLanguage, true) : "-");
                        
                        LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_ASSIGNMENTCOUNTCHECK"),
                            errorMessage.replace("{0}", itemDef.Type()).replace("{1}", itemName).replace("{2}", mismatchName), checkCategory);
                    }
                }
            }
        }
    }
    if (checkReverse == "true")
    {
        for (var itemIndex = 0; itemIndex < itemDefsToCheck.length; itemIndex++)
        {
            var assignedModels = itemDefsToCheck[itemIndex].AssignedModels(type);
            // var superiourList = currentModel.getSuperiorObjDefs();
            var superiourList = helperClearOutInterfaceCollections(currentModel.getSuperiorObjDefs());   // [SM] helperFunktion eingefügt wegen Sammelschnittstellen
    
            
            var superiour     = null;
            if (superiourList.length > 0)
                superiour = superiourList[0];
            
            for (var assignedIndex = 0; assignedIndex < assignedModels.length; assignedIndex++)
            {
                if (superiour != null)
                {
                    var reverseMatch = true;
                
                    if (superiour.OccList([assignedModels[assignedIndex]]).length < 1)
                        reverseMatch = false;
                    
                    if (reverseMatch == false)
                    {
                        var itemDef      = itemDefsToCheck[itemIndex];
                        var itemName     = itemDef.Name(g_usedLanguage, true);
                        var mismatchName = assignedModels[assignedIndex].Name(g_usedLanguage, true);
                        
                        LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_ASSIGNMENTCOUNTCHECK"),
                            errorMessage.replace("{0}", itemDef.Type()).replace("{1}", itemName).replace("{2}", mismatchName), checkCategory);
                    }
                }
                else
                {
                    var itemDef      = itemDefsToCheck[itemIndex];
                    var itemName     = itemDef.Name(g_usedLanguage, true);
                    var mismatchName = assignedModels[assignedIndex].Name(g_usedLanguage, true);
                    
                    LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_ASSIGNMENTCOUNTCHECK"),
                        errorMessage.replace("{0}", itemDef.Type()).replace("{1}", itemName).replace("{2}", mismatchName), checkCategory);
                }
            }
        }
    }
    if (!(errorFoundInModel))
        LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_ASSIGNMENTCOUNTCHECK"), ruleDescription + " - OK", 5); 

}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///  Conventions check for ModelPeriphery
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//check all rules with ruleKind = "ModelPeriphery"
function doConventionsCheckModelPeriphery(modelPeripheryRules, currentModel)
{
   for (var cou = 0; cou < modelPeripheryRules.size(); cou++)
   {
        var ruleXmlNode = modelPeripheryRules.get(cou); 
        var peripheryRuleXmlNode = g_config.getModelPeripheryRule(ruleXmlNode); 
        doModelPeripheryCheck(currentModel, ruleXmlNode, peripheryRuleXmlNode); 
    }
}

//Check the structure constraints, containend in node StructureRule.
// currentModel: The model to check.
// ruleXmlNode: The XML node containing all rule data.
// peripheryRuleXmlNode: The XML sub-node, containing the model periphery check.
function doModelPeripheryCheck(currentModel, ruleXmlNode, peripheryRuleXmlNode)
{
    var checkItemKind = "";
    var checkTypeList = "";
    var checkCategory=""; 
    var ruleDescription =""; 
    var errorMessage =""; 
    var checkType="";
    var ruleShortcut = ""; 
    
    if (peripheryRuleXmlNode != null)
    {
        checkItemKind    = ruleXmlNode.getAttributeValue("kind");
        checkTypeList    = ruleXmlNode.getAttributeValue("typelist");
        checkTypeFilter  = ruleXmlNode.getAttributeValue("typefilter");
        checkCategory    = g_config.getCheckCategory(ruleXmlNode);
        ruleDescription   = g_config.getRuleDescription(ruleXmlNode);
        errorMessage     = g_config.getErrorMessage(ruleXmlNode); 
        checkType        = String(g_config.getRuleType(peripheryRuleXmlNode)); // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
        ruleShortcut     = g_config.getRuleShortcut(ruleXmlNode); 
        
        if (checkCategory <= g_currentLogType)
        {
            switch (checkType)
            {
            case String("SuperiorOccurence"): // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
                if (!(checkSuperiorOccurence(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)))
                    LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"), ruleDescription + " - OK", 5);
                break;
            case String("SuperiorOccurenceConnectionType"):
                if (checkSuperiorOccurenceConnectionType(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut) == false)
                {
                    LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"), ruleDescription + " - OK", 5);
                }
                break;
            case String("SuperiorOccurenceConnectedRiskCategoryMatchesBCDs"):
                if (!(checkSuperiorOccurenceConnectedRiskCategoryMatchesBCDs(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)))
                    LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"), ruleDescription +" - OK", 5);
                break;
            case String("SuperiorStandardSymbol"):
                if (!(checkSuperiorStandardSymbol(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)))
                    LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"), ruleDescription + " - OK", 5);
                break;
            case String("SuperiorName"):
                if (!(checkSuperiorName(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)))
                    LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"), ruleDescription + " - OK", 5);
                break;
            case String("SuperiorZadChapter"):
                if (!(checkSuperiorZadChapter(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)))
                    LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"), ruleDescription + " - OK", 5);
                break;
            default:
                    Dialogs.MsgBox(getString("ERR_WRONGPERIPHERYCHECKTYPE") + checkType, Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION")); 
            }
        }
    }
    else
    {       
        Dialogs.MsgBox(GetText("ERR_RULEINCOMPLETE", new Array(g_config.getRuleName(ruleXmlNode))), Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION")) ; 
    }
}

//Check the superiour object's occurencies within a specific model type.
// currentModel: The model to check.
// peripheryRuleXmlNode: The XML sub-node, containing the periphery check.
// checkItemKind: The item kind to apply on checkTypeList filter (CID_OBJOCC or CID_OBJDEF).
// checkTypeList: The list of types to filter for (Item-Number or Symbol-Number).
// checkTypeFilter: The kind of filter to apply (white/black).
// checkCategory: The the definition of the failure effect.
// ruleDescription: The rule description text.
// errorMessage: The error message text.
function checkSuperiorOccurence(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var errorFoundInModel = false; 
    
    var superiorSymbolType     = helperConvertTypeGuidToTypeNum(peripheryRuleXmlNode.getAttributeValue("superiorSymbolType"), "CID_OBJOCC");
    var superiorModelTypeList  = peripheryRuleXmlNode.getAttributeValue("superiorModelTypelist");
    var sameGroup              = peripheryRuleXmlNode.getAttributeValue("sameGroup");
    var sameModelName          = peripheryRuleXmlNode.getAttributeValue("sameModelName");
    // var superiourDefList       = currentModel.getSuperiorObjDefs();
    var superiourDefList       = helperClearOutInterfaceCollections(currentModel.getSuperiorObjDefs());   // [SM] helperFunktion eingefügt wegen Sammelschnittstellen
    
    
    // Prepare type numbers to check.
    var checkModelTypes;
    if (superiorModelTypeList.indexOf("*") >= 0)
        checkModelTypes = ["*"];
    else if (superiorModelTypeList.indexOf(",") >= 0)
        checkModelTypes = superiorModelTypeList.split(",");
    else
        checkModelTypes = [superiorModelTypeList];
    
    // Convert type GUIDs to type numbers.
    for (var checkTypeIndex = 0; checkTypeIndex < checkModelTypes.length; checkTypeIndex++)
        checkModelTypes[checkTypeIndex] = helperConvertTypeGuidToTypeNum(checkModelTypes[checkTypeIndex], "CID_MODEL");
    
    if (superiourDefList.length == 0)
    {
        var itemName     = currentModel.Name(g_usedLanguage, true);
        
        LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"),
            errorMessage.replace("{0}", currentModel.Type()).replace("{1}", itemName).replace("{2}", "Keine Modell verfügbar!"), checkCategory);
        return true;
    }
    
    if (checkModelTypes.length == 1 && sameModelName == "true")
    {
        Dialogs.MsgBox(getString("ERR_WRONGPERIPHERYCHECKTYPE") + checkType + "\n\r\n\r'sameModelName' kann nur bei mehreren Einträgen in 'superiorModelTypelist' gesetzt werden.",
            Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION"));
        sameModelName == "false";
    }
    
    for (var superiorIndex = 0; superiorIndex < superiourDefList.length; superiorIndex++)
    {
        var p_sGroupName="";
        var p_sSuperiourOccGroupName="";
        var modeltypeDemand = 0;
        var modelasgnMatch  = 0;
        var modeltypeMatch  = 0;
        var leadingModNames = "";
        var mismatchName    = "Kein Modell verfügbar!";
        var currentSuperiourOccList = superiourDefList[superiorIndex].OccList();
        for (var checkModelTypeIndex = 0; checkModelTypeIndex < checkModelTypes.length; checkModelTypeIndex++)
        {
            for (var currentSuperiorOccIndex = 0; currentSuperiorOccIndex < currentSuperiourOccList.length; currentSuperiorOccIndex++)
            {
                if (currentSuperiourOccList[currentSuperiorOccIndex].Model().TypeNum() == checkModelTypes[checkModelTypeIndex])
                {
                    modelasgnMatch += 1;
                    if (checkModelTypeIndex == 0)
                    {
                        // Prevent duplicates.
                        var groupAndName = currentSuperiourOccList[currentSuperiorOccIndex].Model().Group().Name(g_usedLanguage, true) + "/" +
                                           currentSuperiourOccList[currentSuperiorOccIndex].Model().Name(g_usedLanguage, true)
                        if (leadingModNames.indexOf(groupAndName) < 0)
                        {
                            modeltypeDemand += 1;
                            if (leadingModNames != "")
                                leadingModNames  += "~~~";
                            leadingModNames += groupAndName;
                        }
                        
                        mismatchName     = groupAndName;
                    }
                    
                    
                    if (sameGroup == "true" && sameModelName == "true")
                    {
                        // Check the 'same model name' constraint for a model of any ---additional--- model type (checkModelTypeIndex > 0) against any model of the ---initial--- model type.
                        // var groupAndName = currentSuperiourOccList[currentSuperiorOccIndex].Model().Group().Name(g_usedLanguage, true) + "/" +
                        //                   currentSuperiourOccList[currentSuperiorOccIndex].Model().Name(g_usedLanguage, true)
                                           
                        // Änderung wegen neuer ProcessTree
                        var groupAndName = currentSuperiourOccList[currentSuperiorOccIndex].Model().Group().Group().Name(g_usedLanguage, true) + "/" +
                                           currentSuperiourOccList[currentSuperiorOccIndex].Model().Name(g_usedLanguage, true)                   
                        
                        if (checkModelTypeIndex > 0)
                        {
                            if (leadingModNames.indexOf(groupAndName) >= 0)
                                modeltypeMatch += 1;
                            else
                                mismatchName += " - Risiko-Matrix: " +groupAndName ; 
                        }
                    }
                    else if (sameGroup == "true")
                    {
                        // if (superiourDefList[superiorIndex].Group().Name(g_usedLanguage) == currentSuperiourOccList[currentSuperiorOccIndex].Model().Group().Name(g_usedLanguage))
                        // if (currentModel.Group().Name(g_usedLanguage) == currentSuperiourOccList[currentSuperiorOccIndex].Model().Group().Name(g_usedLanguage))
                        
                        // Änderung wegen neuer ProcessTree [SM]
                        var p_sGroupName = currentModel.Group().Name(g_usedLanguage)+"";
                        var p_sSuperiourOccGroupName = currentSuperiourOccList[currentSuperiorOccIndex].Model().Group().Name(g_usedLanguage)+"";
                        
                        if (p_sGroupName.length > 7 && p_sSuperiourOccGroupName.length > 7){
                            if (p_sGroupName.substring(0, 8) == p_sSuperiourOccGroupName.substring(0, 8))
                            {
                                modeltypeMatch += 1;
                            }
                        }
                    }
                    else // (sameGroup != "true")
                    {
                        if (sameModelName == "true")
                        {
                            // Check the 'same model name' constraint for a model of any ---additional--- model type (checkModelTypeIndex > 0) against any model of the ---initial--- model type.
                            if (checkModelTypeIndex > 0 && leadingModNames.indexOf("/" + currentSuperiourOccList[currentSuperiorOccIndex].Model().Name(g_usedLanguage, true)) >= 0)
                                modeltypeMatch += 1;
                        }
                        else
                            modeltypeMatch += 1;
                    }
                }
            }
        }
        
        // Every superiour object's occurence matches the requiremets.
        // if (modeltypeMatch < modelasgnMatch)
        // Any superiour object's occurence matches the requiremets.
        if (sameModelName == "true" && modeltypeDemand > modeltypeMatch ||
            sameModelName != "true" && modeltypeMatch < 1)
        {
            var itemName     = currentModel.Name(g_usedLanguage, true);
            
            LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"),
                errorMessage.replace("{0}", currentModel.Type()).replace("{1}", itemName).replace("{2}", mismatchName ), checkCategory);
            errorFoundInModel = true; 
            return errorFoundInModel;
        }
    }
    return errorFoundInModel; 
}

//Check the superiour object's occurencies within a specific model type and the occurence of a specific connection.
// currentModel: The model to check.
// peripheryRuleXmlNode: The XML sub-node, containing the periphery check.
// checkItemKind: The item kind to apply on checkTypeList filter (CID_OBJOCC or CID_OBJDEF).
// checkTypeList: The list of types to filter for (Item-Number or Symbol-Number).
// checkTypeFilter: The kind of filter to apply (white/black).
// checkCategory: The the definition of the failure effect.
// ruleDescription: The rule description text.
// errorMessage: The error message text.
function checkSuperiorOccurenceConnectionType(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var errorFoundInModel = false; 
    var superiorSymbolType     = helperConvertTypeGuidToTypeNum(peripheryRuleXmlNode.getAttributeValue("superiorSymbolType"), "CID_OBJOCC");
    var superiorModelTypeList  = peripheryRuleXmlNode.getAttributeValue("superiorModelTypelist");
    var superiorConnectionType = peripheryRuleXmlNode.getAttributeValue("superiorConnectionType");
    var superiourDefList       = currentModel.getSuperiorObjDefs();
    
    // Prepare type numbers to check.
    var checkModelTypes;
    if (superiorModelTypeList.indexOf("*") >= 0)
        checkModelTypes = ["*"];
    else if (superiorModelTypeList.indexOf(",") >= 0)
        checkModelTypes = superiorModelTypeList.split(",");
    else
        checkModelTypes = [superiorModelTypeList];
    
    // Convert type GUIDs to type numbers.
    for (var checkTypeIndex = 0; checkTypeIndex < checkModelTypes.length; checkTypeIndex++)
        checkModelTypes[checkTypeIndex] = helperConvertTypeGuidToTypeNum(checkModelTypes[checkTypeIndex], "CID_MODEL");
    
    if (superiourDefList.length == 0)
    {
        var itemName     = currentModel.Name(g_usedLanguage, true);
        
        LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"),
            GetText("ERR_NOSUPERIOROBJDEF", new Array(itemName)), checkCategory);
        errorFoundInModel = true; 
        return errorFoundInModel;
    }
    
    for (var superiorIndex = 0; superiorIndex < superiourDefList.length; superiorIndex++)
    {
        var mismatchName     = getString("ERR_NOMODELAVAILABLE");
        var currentSuperiourOccList = superiourDefList[superiorIndex].OccList();
        for (var currentSuperiorOccIndex = 0; currentSuperiorOccIndex < currentSuperiourOccList.length; currentSuperiorOccIndex++)
        {
            for (var checkModelTypeIndex = 0; checkModelTypeIndex < checkModelTypes.length; checkModelTypeIndex++)
            {
                if (currentSuperiourOccList[currentSuperiorOccIndex].Model().TypeNum() == checkModelTypes[checkModelTypeIndex])
                {
                    mismatchName = currentSuperiourOccList[currentSuperiorOccIndex].Model().Name(g_usedLanguage, true);
                    var connectionMatch = false;
                    if (checkModelTypes[checkModelTypeIndex] == 220)
                    {
                        // Attention: Matrix models don't provide CxnOcc!!!
                        // ================================================
                    
                        // 1. Find the conterpart objects, that could be connected with currentSuperiourOcc within currentModel.
                        var realCounterparts = helperGetConnectedObjDefListForObjOcc(currentSuperiourOccList[currentSuperiorOccIndex]);
                        
                        // 2. Check the connection between counterpart objects and currentSuperiourOcc within currentModel (matrix).
                        var currentMatrix = currentSuperiourOccList[currentSuperiorOccIndex].Model().getMatrixModel();
                        if (currentMatrix != null)
                        {
                            var matrixContentCells = helperGetMatrixCells(currentMatrix, superiourDefList[superiorIndex], realCounterparts, superiorConnectionType);
                            if (matrixContentCells.length > 0)
                                connectionMatch = true;
                        }
                    }/*
                    else // (checkModelTypes[checkModelTypeIndex] != 220)
                    {
                        var realCounterpartOccs = new Array();
                        var connections = currentSuperiourOccList[currentSuperiorOccIndex].Cxns();
                        for (var cxnIndex = 0; cxnIndex < connections.length; cxnIndex++)
                        {
                            if (connections[cxnIndex].Cxn().TypeNum() == superiorConnectionType)
                            {
                                if (connections[cxnIndex].SourceObjOcc().ObjDef().GUID() == currentSuperiourOccList[currentSuperiorOccIndex].ObjDef().GUID())
                                    realCounterpartOccs.push(connections[cxnIndex].TargetObjOcc());
                                else
                                    realCounterpartOccs.push(connections[cxnIndex].SourceObjOcc());
                            }
                        }
                    }*/
                    
                    if (connectionMatch == false)
                    {
                        var itemName     = currentModel.Name(g_usedLanguage, true);
                        
                        LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"),
                            errorMessage.replace("{0}", currentModel.Type()).replace("{1}", itemName).replace("{2}", mismatchName), checkCategory);
                        
                        errorFoundInModel = true; 
                    }
                }
            }
        }
    } 
   
    return errorFoundInModel; 
}

//Check the superiour object's occurencies within a specific model type and the occurence of a specific connection.
// currentModel: The model to check.
// peripheryRuleXmlNode: The XML sub-node, containing the periphery check.
// checkItemKind: The item kind to apply on checkTypeList filter (CID_OBJOCC or CID_OBJDEF).
// checkTypeList: The list of types to filter for (Item-Number or Symbol-Number).
// checkTypeFilter: The kind of filter to apply (white/black).
// checkCategory: The the definition of the failure effect.
// ruleDescription: The rule description text.
// errorMessage: The error message text.
function checkSuperiorOccurenceConnectedRiskCategoryMatchesBCDs(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var errorFoundInModel = false; 
    
    var superiorSymbolType     = helperConvertTypeGuidToTypeNum(peripheryRuleXmlNode.getAttributeValue("superiorSymbolType"), "CID_OBJOCC");
    var superiorModelTypeList  = peripheryRuleXmlNode.getAttributeValue("superiorModelTypelist");
    var superiorConnectionType = peripheryRuleXmlNode.getAttributeValue("superiorConnectionType");
    var equalRiskCategoryFor   = peripheryRuleXmlNode.getAttributeValue("equalRiskCategoryFor");
    // var superiourDefList       = currentModel.getSuperiorObjDefs();
    var superiourDefList = helperClearOutInterfaceCollections(currentModel.getSuperiorObjDefs());   // [SM] helperFunktion eingefügt wegen Sammelschnittstellen
    
    // Prepare type numbers to check.
    var checkModelTypes;
    if (superiorModelTypeList.indexOf("*") >= 0)
        checkModelTypes = ["*"];
    else if (superiorModelTypeList.indexOf(",") >= 0)
        checkModelTypes = superiorModelTypeList.split(",");
    else
        checkModelTypes = [superiorModelTypeList];
    
    // Convert type GUIDs to type numbers.
    for (var checkTypeIndex = 0; checkTypeIndex < checkModelTypes.length; checkTypeIndex++)
        checkModelTypes[checkTypeIndex] = helperConvertTypeGuidToTypeNum(checkModelTypes[checkTypeIndex], "CID_MODEL");
    
    if (superiourDefList.length == 0)
    {
        var itemName     = currentModel.Name(g_usedLanguage, true);
        
        LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"),
            GetText("ERR_NOSUPERIOROBJDEF", new Array(itemName)), checkCategory);
        errorFoundInModel = true; 
        return errorFoundInModel;
    }
    
    // Determine all technical terms, that define a risk category for any of the contained risks.
    var foundRiskCategoryTermDefs = new Array();
    if (equalRiskCategoryFor == "Risk")
    {
        var foundRisks = currentModel.ObjOccListFilter(Constants.OT_RISK);
        for (var riskIndex = 0; riskIndex < foundRisks.length; riskIndex++)
        {
            var currentBCDs = foundRisks[riskIndex].ObjDef().AssignedModels(Constants.MT_BUSY_CONTR_DGM);
            for (var bcdIndex = 0; bcdIndex < currentBCDs.length; bcdIndex++)
            {
                // ToDo: Nachfragen, ob die Abfrage von Ausprägungen der FBs reicht, oder ob das Risiko ermittelt und die Kanten verfolgt werden müssen.
                var containedTechTerms = currentBCDs[bcdIndex].ObjDefListFilter(Constants.OT_TECH_TRM);
                for (var techTermIndex = 0; techTermIndex < containedTechTerms.length; techTermIndex++)
                {
                    foundRiskCategoryTermDefs.push(containedTechTerms[techTermIndex]);
                }
            }
        }
    }
    // Determine all technical terms, that define a risk category for any of the contained control.
    else
    {
        var isControlAttribute = helperConvertTypeGuidToTypeNum("{7aa94170-8d7c-11e2-1a33-abe82ad1c06e}", "CID_ATTRDEF");
        // var foundPossibleControls = currentModel.ObjOccListFilter(Constants.OT_FUNC);
		var foundControls = currentModel.ObjOccListFilter(-1, Constants.ST_CONTR);
		
        //for (var controlIndex = 0; controlIndex < foundPossibleControls.length; controlIndex++)
		for (var controlIndex = 0; controlIndex < foundControls.length; controlIndex++)
        {
            //var attr = foundPossibleControls[controlIndex].ObjDef().Attribute(isControlAttribute, g_usedLanguage, true);
            //if (attr != null && attr.IsMaintained() == true && attr.GetValue(true) == "ja")
            //{
                //var currentOccs = foundPossibleControls[controlIndex].ObjDef().OccList();
				var currentOccs = foundControls[controlIndex].ObjDef().OccList();
                for (var currentOccIndex = 0; currentOccIndex < currentOccs.length; currentOccIndex++)
                {
                    if (currentOccs[currentOccIndex].Model().TypeNum() == Constants.MT_BUSY_CONTR_DGM)
                    {
                        var containedTechTerms = currentOccs[currentOccIndex].Model().ObjDefListFilter(Constants.OT_TECH_TRM);
                        for (var techTermIndex = 0; techTermIndex < containedTechTerms.length; techTermIndex++)
                        {
                            foundRiskCategoryTermDefs.push(containedTechTerms[techTermIndex]);
                        }
                    }
                }
            //}
        }
    }
    
    foundRiskCategoryTermDefs = ArisData.Unique(foundRiskCategoryTermDefs);
    
    /* OLD: There is no need to evaluate the Matrix model, instead it's enough to determine the connections of type ''!
    for (var superiorIndex = 0; superiorIndex < superiourDefList.length; superiorIndex++)
    {
        var mismatchName     = "Kein Modell verfügbar!";
        var currentSuperiourOccList = superiourDefList[superiorIndex].OccList();
        for (var currentSuperiorOccIndex = 0; currentSuperiorOccIndex < currentSuperiourOccList.length; currentSuperiorOccIndex++)
        {
            for (var checkModelTypeIndex = 0; checkModelTypeIndex < checkModelTypes.length; checkModelTypeIndex++)
            {
                if (currentSuperiourOccList[currentSuperiorOccIndex].Model().TypeNum() == checkModelTypes[checkModelTypeIndex])
                {
                    mismatchName = currentSuperiourOccList[currentSuperiorOccIndex].Model().Name(g_usedLanguage, true);
                    var matchingRiskCategoryRiskDefs = new Array();
                    if (checkModelTypes[checkModelTypeIndex] == 220)
                    {
                        // Attention: Matrix models don't provide CxnOcc!!!
                        // ================================================
                    
                        // 1. Find the conterpart objects, that could be connected with currentSuperiourOcc within currentModel.
                        var realCounterparts = helperGetConnectedObjDefListForObjOcc(currentSuperiourOccList[currentSuperiorOccIndex]);
                        
                        // 2. Check the connection between counterpart objects and currentSuperiourOcc within currentModel (matrix).
                        var currentMatrix = currentSuperiourOccList[currentSuperiorOccIndex].Model().getMatrixModel();
                        if (currentMatrix != null)
                        {
                            matchingRiskCategoryRiskDefs = helperGetMatrixConnectedCounterparts(currentMatrix, superiourDefList[superiorIndex], realCounterparts, superiorConnectionType);
                        }
                    }
                    //else // (checkModelTypes[checkModelTypeIndex] != 220)
                    //{
                    //    var realCounterpartOccs = new Array();
                    //    var connections = currentSuperiourOccList[currentSuperiorOccIndex].Cxns();
                    //    for (var cxnIndex = 0; cxnIndex < connections.length; cxnIndex++)
                    //    {
                    //        if (connections[cxnIndex].Cxn().TypeNum() == superiorConnectionType)
                    //        {
                    //            if (connections[cxnIndex].SourceObjOcc().ObjDef().GUID() == currentSuperiourOccList[currentSuperiorOccIndex].ObjDef().GUID())
                    //                realCounterpartOccs.push(connections[cxnIndex].TargetObjOcc());
                    //            else
                    //                realCounterpartOccs.push(connections[cxnIndex].SourceObjOcc());
                    //        }
                    //    }
                    //}
                    matchingRiskCategoryRiskDefs = ArisData.Unique(matchingRiskCategoryRiskDefs);
                    
                    var connectionMatch = true;
                    var mismatchDetailName = "";
                    if (matchingRiskCategoryRiskDefs.length >= foundRiskCategoryTermDefs.lenght)
                    {
                        for (var objDefIndex = 0; objDefIndex < matchingRiskCategoryRiskDefs.length; objDefIndex++)
                        {
                            if (helpNameMatch(matchingRiskCategoryRiskDefs[objDefIndex], foundRiskCategoryTermDefs) == false)
                            {
                                connectionMatch = false;
                                if (mismatchDetailName == "")
                                    mismatchDetailName = matchingRiskCategoryRiskDefs[objDefIndex].Name(g_usedLanguage, true);
                                else
                                    mismatchDetailName += ", " + matchingRiskCategoryRiskDefs[objDefIndex].Name(g_usedLanguage, true);
                            }
                        }
                    }
                    else
                    {
                        for (var objDefIndex = 0; objDefIndex < foundRiskCategoryTermDefs.length; objDefIndex++)
                        {
                            if (helpNameMatch(foundRiskCategoryTermDefs[objDefIndex], matchingRiskCategoryRiskDefs) == false)
                            {
                                connectionMatch = false;
                                if (mismatchDetailName == "")
                                    mismatchDetailName = foundRiskCategoryTermDefs[objDefIndex].Name(g_usedLanguage, true);
                                else
                                    mismatchDetailName += ", " + foundRiskCategoryTermDefs[objDefIndex].Name(g_usedLanguage, true);
                            }
                        }
                    }
                    
                    if (connectionMatch == false)
                    {
                        var itemName     = currentModel.Name(g_usedLanguage, true);
                        
                        LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"),
                            errorMessage.replace("{0}", currentModel.Type()).replace("{1}", itemName).replace("{2}", mismatchName).replace("{3}", mismatchDetailName), checkCategory);
                        errorFoundInModel = true; 
                    }
                }
            }
        }
    }*/
    for (var superiorIndex = 0; superiorIndex < superiourDefList.length; superiorIndex++)
    {
        // var currentSuperiourOccList = superiourDefList[superiorIndex].CxnListFilter(Constants.EDGES_INOUT, superiorConnectionType);
        var matchingRiskCategoryRiskDefs = superiourDefList[superiorIndex].getConnectedObjs ([Constants.OT_RISK], Constants.EDGES_INOUT, [superiorConnectionType]);
        
        var connectionMatch = true;
        var mismatchDetailName = "";
        // if (matchingRiskCategoryRiskDefs.length >= foundRiskCategoryTermDefs.length)  // [SM]: auskommentiert, da IMMER beide Seiten (nur in BCD/nur in Matrix) betrachtet werden müssen
        //{
            for (var objDefIndex = 0; objDefIndex < matchingRiskCategoryRiskDefs.length; objDefIndex++)
            {
                if ((helpNameMatch(matchingRiskCategoryRiskDefs[objDefIndex], foundRiskCategoryTermDefs) == false) && (matchingRiskCategoryRiskDefs[objDefIndex].Name(1031, true) != "Kein Risiko/ keine Kontrolle vorhanden"))
                {
                    connectionMatch = false;
                    if (mismatchDetailName == "")
                        mismatchDetailName = matchingRiskCategoryRiskDefs[objDefIndex].Name(g_usedLanguage, true);
                    else
                        mismatchDetailName += ", " + matchingRiskCategoryRiskDefs[objDefIndex].Name(g_usedLanguage, true);
						mismatchDetailName += " (nur in Matrix)";   // [SM]: Zusatz, damit man weiß wo der Fehler liegt
                }
            }
        //}
        //else
        //{
            for (var objDefIndex = 0; objDefIndex < foundRiskCategoryTermDefs.length; objDefIndex++)
            {
                if (helpNameMatch(foundRiskCategoryTermDefs[objDefIndex], matchingRiskCategoryRiskDefs) == false)
                {
                    connectionMatch = false;
                    if (mismatchDetailName == "")
                        mismatchDetailName = foundRiskCategoryTermDefs[objDefIndex].Name(g_usedLanguage, true);
                    else
                        mismatchDetailName += ", " + foundRiskCategoryTermDefs[objDefIndex].Name(g_usedLanguage, true);
						mismatchDetailName += " (nur in BCD)";      // [SM]: Zusatz, damit man weiß wo der Fehler liegt
                }
            }
		//}
        
        if (connectionMatch == false)
        {
            var itemName     = superiourDefList[superiorIndex].Name(g_usedLanguage, true);
            
            LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"),
                errorMessage.replace("{0}", superiourDefList[superiorIndex].Type()).replace("{1}", itemName).replace("{3}", mismatchDetailName), checkCategory);
            errorFoundInModel = true; 
        }
    }
    
    return errorFoundInModel;     
}


//Check the superiour object's default symbol.
// currentModel: The model to check.
// peripheryRuleXmlNode: The XML sub-node, containing the periphery check.
// checkItemKind: The item kind to apply on checkTypeList filter (CID_OBJOCC or CID_OBJDEF).
// checkTypeList: The list of types to filter for (Item-Number or Symbol-Number).
// checkTypeFilter: The kind of filter to apply (white/black).
// checkCategory: The definition of the failure effect.
// ruleDescription: The rule description text.
// errorMessage: The error message text.
function checkSuperiorStandardSymbol(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var errorFoundInModel = false; 
    
    var superiorType  = helperConvertTypeGuidToTypeNum(peripheryRuleXmlNode.getAttributeValue("superiorSymbolType"), "CID_OBJOCC");
    var attrTypeNum   = helperConvertTypeGuidToTypeNum("{317e1987-7df8-11da-0c60-cf8e338f9a0b}", "CID_ATTRDEF");
    var prcsTypeNum   = 72;
    var fprcsTypeNum  = helperConvertTypeGuidToTypeNum("{f30c7df0-3137-11e1-44a6-00300571cf1f}", "CID_OBJOCC");
    // var superiourDefList = currentModel.getSuperiorObjDefs();
    var superiourDefList = helperClearOutInterfaceCollections(currentModel.getSuperiorObjDefs());   // [SM] helperFunktion eingefügt wegen Sammelschnittstellen
    
    if (superiourDefList.length == 0)
        return;
    
    for (var superiorIndex = 0; superiorIndex < superiourDefList.length; superiorIndex++)
    {
        var match = false;
        var enterprise = superiourDefList[superiorIndex].Attribute(attrTypeNum, g_usedLanguage, true).GetValue(true);
        if (enterprise == "UniCredit Bank AG")
        {
            if (superiourDefList[superiorIndex].getDefaultSymbolNum() == prcsTypeNum)
                match = true;
        }
        else if (vbLen(enterprise) != 0)
        {
            if (superiourDefList[superiorIndex].getDefaultSymbolNum() == fprcsTypeNum)
                match = true;
        }
        else
        {
             enterprise = getString("TXT_ATTRIBUTENOTMAINTAINED"); // bei ungefüllten Attribut -> Fehler
        }

        if (match == false)
        {
            var itemName     = currentModel.Name(g_usedLanguage, true);
            LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"),
                errorMessage.replace("{0}", currentModel.Type()).replace("{1}", itemName).replace("{2}", enterprise), checkCategory);
            errorFoundInModel = true; 
            return errorFoundInModel;
        }
    }
    return errorFoundInModel; 
}

//Check the superiour object's name is equal to the process name.
// currentModel: The model to check.
// peripheryRuleXmlNode: The XML sub-node, containing the periphery check.
// checkItemKind: The item kind to apply on checkTypeList filter (CID_OBJOCC or CID_OBJDEF).
// checkTypeList: The list of types to filter for (Item-Number or Symbol-Number).
// checkTypeFilter: The kind of filter to apply (white/black).
// checkCategory: The the definition of the failure effect.
// ruleDescription: The rule description text.
// errorMessage: The error message text.
function checkSuperiorName(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var errorFoundInModel = false; 
    
    var superiorType  = helperConvertTypeGuidToTypeNum(peripheryRuleXmlNode.getAttributeValue("superiorSymbolType"), "CID_OBJOCC");
    
    //var superiourDefList = currentModel.getSuperiorObjDefs();
    var superiourDefList = helperClearOutInterfaceCollections(currentModel.getSuperiorObjDefs());   // [SM] helperFunktion eingefügt wegen Sammelschnittstellen
    
    if (superiourDefList.length == 0)
        return;
    
    for (var superiorIndex = 0; superiorIndex < superiourDefList.length; superiorIndex++)
    {
        var match = false;
        if (superiourDefList[superiorIndex].Name(g_usedLanguage, true) == currentModel.Name(g_usedLanguage, true))
            match = true;
        
        if (match == false)
        {
            var itemName     = currentModel.Name(g_usedLanguage, true);
            
            LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"),
                errorMessage.replace("{0}", currentModel.Type()).replace("{1}", itemName), checkCategory);
                errorFoundInModel = true; 
            return errorFoundInModel;
        }
    }
    return errorFoundInModel
}

// [SM] Clear out Interface colletions from object list (checks, if first 3 Characters="ZAD")
function helperClearOutInterfaceCollections(aSuperiorObjectList){
	var superiorIndex;	// counter
    for (var superiorIndex = 0; superiorIndex < aSuperiorObjectList.length; superiorIndex++){
		
//        if (aSuperiorObjectList[superiorIndex].Name(g_usedLanguage, true).length > 2){
            // Wenn der Name des übergeordneten Objekts mit "ZAD" beginnt -> Sammelschnittstelle
            if (aSuperiorObjectList[superiorIndex].Name(g_usedLanguage, true).substring(0,3) == "ZAD"){
                aSuperiorObjectList.splice(superiorIndex,1);
                superiorIndex--;
            }
//        }
    }
	// Bereinigten Array zurück melden
	return aSuperiorObjectList;
}

//Check the superiour object's name is equal to the process name.
// currentModel: The model to check.
// peripheryRuleXmlNode: The XML sub-node, containing the periphery check.
// checkItemKind: The item kind to apply on checkTypeList filter (CID_OBJOCC or CID_OBJDEF).
// checkTypeList: The list of types to filter for (Item-Number or Symbol-Number).
// checkTypeFilter: The kind of filter to apply (white/black).
// checkCategory: The the definition of the failure effect.
// ruleDescription: The rule description text.
// errorMessage: The error message text.
function checkSuperiorZadChapter(currentModel, peripheryRuleXmlNode, checkItemKind, checkTypeList, checkTypeFilter, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var errorFoundInModel = false; 
    var superiorType  = helperConvertTypeGuidToTypeNum(peripheryRuleXmlNode.getAttributeValue("superiorSymbolType"), "CID_OBJOCC");
    // var superiourDefList = currentModel.getSuperiorObjDefs();
    var superiourDefList = helperClearOutInterfaceCollections(currentModel.getSuperiorObjDefs());   // [SM] helperFunktion eingefügt wegen Sammelschnittstellen
    
    if (superiourDefList.length == 0)
        return;
    
    for (var superiorIndex = 0; superiorIndex < superiourDefList.length; superiorIndex++)
    {
        var match = false;
        var modelZadChaper = "";
        var superiorZadChaper = "";
        
        var attr = currentModel.Attribute(Constants.AT_USER_ATTR1, g_usedLanguage, true);
        if (attr != null && attr.IsMaintained() == true)
            modelZadChaper = attr.GetValue(true);
        attr = superiourDefList[superiorIndex].Attribute(Constants.AT_USER_ATTR1, g_usedLanguage, true);
        if (attr != null && attr.IsMaintained() == true)
            superiorZadChaper = attr.GetValue(true);
        
        if (modelZadChaper != "" && superiorZadChaper != "")
        {
            if (modelZadChaper.length <= superiorZadChaper.length)
            {
                if (superiorZadChaper.indexOf(modelZadChaper) >= 0)
                    match = true;
            }
            else
            {
                if (modelZadChaper.indexOf(superiorZadChaper) >= 0)
                    match = true;
            }
        }else if (modelZadChaper == "" && superiorZadChaper == ""){ // [SM]: wenn beide nicht befüllt sind, dann sind sie natürlich auch gleich und match = true
            match = true;
        }
        
        if (match == false)
        {
            var itemName     = currentModel.Name(g_usedLanguage, true);
            
            LogMessage(getString("TXT_CHECKTYPEMODELPERIPHERY"), ruleShortcut, ruleDescription, getString("TXT_SUPERIORCHECK"),
                errorMessage.replace("{0}", currentModel.Type()).replace("{1}", itemName), checkCategory);
                errorFoundInModel = true; 
            return errorFoundInModel;
        }
    }
    return errorFoundInModel; 
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///  Conventions check for Object Environment (Object occurences and connections to other objects
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//check all rules with ruleKind = "ObjectEnvironment"
function doConventionsCheckObjectOccurences(objectOccRules, currentModel)
{
   for (var cou = 0; cou < objectOccRules.size(); cou++)
   {
        var ruleXmlNode = objectOccRules.get(cou); 
        var occurenceRuleXmlNode = g_config.getObjectOccurenceRule(ruleXmlNode); 
        doObjectEnvironmentCheck(currentModel, ruleXmlNode, occurenceRuleXmlNode); 
    }
}

//Check the structure constraints, containend in node StructureRule.
// currentModel: The model to check.
// ruleXmlNode: The XML node containing all rule data.
// peripheryRuleXmlNode: The XML sub-node, containing the model periphery check.
function doObjectEnvironmentCheck(currentModel, ruleXmlNode, occurenceRuleXmlNode)
{
    var checkItemKind = "";
    var checkTypeList = "";
    var checkCategory=""; 
    var ruleDescription =""; 
    var errorMessage =""; 
    var checkType="";
    var ruleShortcut=""; 
    
    if (occurenceRuleXmlNode != null)
    {
        checkItemKind    = ruleXmlNode.getAttributeValue("kind");
        checkTypeList    = ruleXmlNode.getAttributeValue("typelist");
        checkCategory    = g_config.getCheckCategory(ruleXmlNode);
        ruleDescription   = g_config.getRuleDescription(ruleXmlNode);
        errorMessage     = g_config.getErrorMessage(ruleXmlNode); 
        checkType        = String(g_config.getRuleType(occurenceRuleXmlNode)); // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
        ruleShortcut     = g_config.getRuleShortcut(ruleXmlNode); 
        
        if (checkCategory <= g_currentLogType)
        {
            switch (checkType)
            {
            case String("CheckConnectionInModel"): // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
                CheckConnectionInModel(currentModel, occurenceRuleXmlNode, checkTypeList, checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                break;
            case String("CheckUniqueOccurence"):
                CheckUniqueOccurence(currentModel, occurenceRuleXmlNode, checkTypeList, checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                break;
            case String("CheckRiskControlFunction"):
                CheckRiskControlFunction(currentModel, checkTypeList, checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                break;
            case String("CheckObjectLocationInProcessGroup"):
                CheckObjectLocationInProcessGroup(currentModel, checkTypeList, checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                break;
            case String("checkDefaultSymbolType"):
                checkDefaultSymbolType(currentModel, checkTypeList, checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                break;
            default:
                    Dialogs.MsgBox(getString("ERR_WRONGOBJECTENVIRONMENTCHECKTYPE") + checkType, Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION")); 
            }
        }
    }
    else
    {       
        Dialogs.MsgBox(GetText("ERR_RULEINCOMPLETE", new Array(g_config.getRuleName(ruleXmlNode))), Constants.MSGBOX_BTN_OK, getString("TITL_ERRCONFIGURATION")) ; 
    }
}

//the function checks if the configured connections are used (only allowed symbol types as source/target object and allowed connection type)
function CheckConnectionInModel(currentModel, peripheryRuleXmlNode, checkTypeList, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var connectedSymbolType     = helperConvertTypeGuidToTypeNum(peripheryRuleXmlNode.getAttributeValue("connectedSymbolType"), "CID_OBJOCC");
    var occModelTypeList  = peripheryRuleXmlNode.getAttributeValue("occurenceModelTypeList");
    var connectionTypeNum = peripheryRuleXmlNode.getAttributeValue("connectionType"); 
    // Prepare type numbers to check.
    var checkModelTypes;
    if (occModelTypeList.indexOf(",") >= 0)
        checkModelTypes = occModelTypeList.split(",");
    else
        checkModelTypes = [occModelTypeList];
    
    // Convert type GUIDs to type numbers.
    for (var checkTypeIndex = 0; checkTypeIndex < checkModelTypes.length; checkTypeIndex++)
        checkModelTypes[checkTypeIndex] = helperConvertTypeGuidToTypeNum(checkModelTypes[checkTypeIndex], "CID_MODEL");
    
    var objOccListToCheck=null; 
    if (IsInArray(checkModelTypes, currentModel.TypeNum()))
    {
        objOccListToCheck = currentModel.ObjOccListBySymbol(checkTypeList); 
        if (objOccListToCheck.length > 0)
            CheckConnectionsForSymbolList(currentModel, objOccListToCheck, connectionTypeNum, connectedSymbolType, errorMessage, checkCategory, ruleShortcut, ruleDescription); // [SM]: ruleDescription als Parameter ergänzt
        
    }
    else
    {
        //look for assignments of configured model type in the current model
        for (var couOccs=0; couOccs<currentModel.ObjDefList().length; couOccs++)
        {
            var obj = currentModel.ObjDefList()[couOccs]; 
            if (obj.AssignedModels().length > 0)
            {
                for (var couM=0; couM<obj.AssignedModels().length; couM++)
                {
                    if (IsInArray(checkModelTypes, obj.AssignedModels()[couM].TypeNum()))
                    {
                        //get the object occs with the configured symbol types
                        objOccListToCheck =obj.AssignedModels()[couM].ObjOccListBySymbol(checkTypeList);
                        CheckConnectionsForSymbolList(obj.AssignedModels()[couM], objOccListToCheck, connectionTypeNum, connectedSymbolType, errorMessage, checkCategory, ruleShortcut, ruleDescription); // [SM]: ruleDescription als Parameter ergänzt
                        
                    }
                }
            }
            
        }
    }
}

function CheckConnectionsForSymbolList(currentModel, objOccListToCheck, connectionTypeNum, connectedSymbolType, errorMessage, checkCategory, ruleShortcut, ruleDescription) // [SM]: ruleDescription als Parameter ergänzt
{
    var bConnectionFound = false; 
    var occToCheck = null;  
    var cxnOccList = null; 
    for(var couO=0; couO<objOccListToCheck.length; couO++)
    {
        bConnectionFound = false; 
        occToCheck = objOccListToCheck[couO]; 
        cxnOccList = occToCheck.CxnOccList(); 
        if (cxnOccList.length> 0)
        {
            for(var couC=0; couC<cxnOccList.length; couC++)
            {
                if (cxnOccList[couC].CxnDef().TypeNum() == connectionTypeNum)
                {
                    var connectedOcc;
                    if (cxnOccList[couC].SourceObjOcc().IsEqual(occToCheck))
                        connectedOcc = cxnOccList[couC].TargetObjOcc(); 
                    else
                        connectedOcc = cxnOccList[couC].SourceObjOcc(); 
                    
                    if (connectedOcc.SymbolNum() == connectedSymbolType)
                        bConnectionFound = true; 
                }
            }
        }
        if (!(bConnectionFound))
        {
            errorMessage = errorMessage.replace("{0}", occToCheck.ObjDef().Name(g_usedLanguage)); 
           LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, "Model '" + currentModel.Name(g_usedLanguage)+"': " + getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessage, checkCategory); 
        }
    }
}
//Anzahl der Ausprägungen wird überprüft. Manche Symboltypen dürfen nur einmalig in einem Modell verwendet werden. 
//Beispiel: eine Kontrolle darf zwar mehrfach in einer EPK aber nicht in mehreren EPKs vorkommen
// es werden die im aktuellen Modell vorkommenden Objektausprägungen überprüft und zusätzlich die in den hinterlegten BCDs
function CheckUniqueOccurence(currentModel, occRuleXmlNode, checkTypeList, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var checkTypes = new Array(); 
    checkTypes = checkTypeList.split(",");
    var occModelTypeList  = occRuleXmlNode.getAttributeValue("occurenceModelTypeList");
    // Prepare type numbers to check.
    var checkModelTypes;
    if (occModelTypeList.indexOf(",") >= 0)
        checkModelTypes = occModelTypeList.split(",");
    else
        checkModelTypes = [occModelTypeList];
    
    // Convert type GUIDs to type numbers.
    for (var checkTypeIndex = 0; checkTypeIndex < checkModelTypes.length; checkTypeIndex++)
        checkModelTypes[checkTypeIndex] = helperConvertTypeGuidToTypeNum(checkModelTypes[checkTypeIndex], "CID_MODEL");
    
    var objOccListToCheck=null;
    var couOcc=0; 
    var bError = false; 
    g_objsToCheck = new Array(); 
    objOccListToCheck = currentModel.ObjOccListBySymbol(checkTypeList); 
    var errorMessageDetails=""; 
    var countOccs= 0; 
    if (objOccListToCheck.length > 0)
    {

        
        for (couOcc=0; couOcc<objOccListToCheck.length; couOcc++)
        {
            countOccs=0; 
            //if the current model type is part of the model types to check, it is necessary to increment the counter already
            if (IsInArray(checkModelTypes, currentModel.TypeNum()))
                countOccs++; 
            //check if this is the first occurence of the current object
            isNewObject = helperAddObjToList(objOccListToCheck[couOcc].ObjDef()); 
            
            
            if (isNewObject)
            {
                countOccs = countOccs + helperCountNumberOfOccurencesInOtherModels(objOccListToCheck[couOcc], currentModel, checkModelTypes); 
                if (countOccs > 1)
                {
                    errorMessageDetails = errorMessage.replace("{0}", objOccListToCheck[couOcc].ObjDef().Name(g_usedLanguage)).replace("{1}", String(countOccs)); 
                    LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessageDetails, checkCategory); 
                }
                else
                {   //check for multiple occurences in the current model for symbol types function and risk 
                    //pj1 - added symbol system function actual
                    if (IsInArray(checkModelTypes, currentModel.TypeNum()))
                    {
                        if (objOccListToCheck[couOcc].SymbolNum() == Constants.ST_FUNC || objOccListToCheck[couOcc].SymbolNum() == Constants.ST_SYS_FUNC_ACT || 
                            objOccListToCheck[couOcc].SymbolNum() == Constants.ST_RISK_1 || objOccListToCheck[couOcc].SymbolNum() == Constants.ST_CONTR)
                        {
                            if (objOccListToCheck[couOcc].ObjDef().OccListInModel(currentModel).length > 1)
                            {
                            countOccs= objOccListToCheck[couOcc].ObjDef().OccListInModel(currentModel).length; 
                            errorMessageDetails = errorMessage.replace("{0}", objOccListToCheck[couOcc].ObjDef().Name(g_usedLanguage)).replace("{1}", String(countOccs)); 
                            LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription,  getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessageDetails, checkCategory); 
        
                            }
                            else
                            {
                                errorMessageDetails = errorMessage.replace("{0}", objOccListToCheck[couOcc].ObjDef().Name(g_usedLanguage)).replace("{1}", String(countOccs)) + " - OK"; 
                                LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessageDetails, 5); 
        
                            }
                        }
                        else
                        {
                            errorMessageDetails = errorMessage.replace("{0}", objOccListToCheck[couOcc].ObjDef().Name(g_usedLanguage)).replace("{1}", String(countOccs)) + " - OK"; 
                            LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessageDetails, 5); 
                        }
                    }
                    else
                    {
                        errorMessageDetails = errorMessage.replace("{0}", objOccListToCheck[couOcc].ObjDef().Name(g_usedLanguage)).replace("{1}", String(countOccs)) + " - OK"; 
                        LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessageDetails, 5); 

                    }
                }
            }
            
        }
    }    
    //die hinterlegten BCDs müssen nur geprüft werden, wenn die Regel nicht für Risiken definiert wurde. Diese wurden bereits über die Ausprägungen in der EPK geprüft
    
    if (!(IsInArray(checkTypes, Constants.ST_RISK_1)))
    {
        
        //look for assignments of the risks in the current model
        for (var couOccs=0; couOccs<currentModel.ObjDefListByTypes([Constants.OT_RISK]).length; couOccs++)
        {
            bError = false; 
            // var obj = currentModel.ObjDefList()[couOccs];    [SM]: Fehler, nur Risiken im Array berücksichtigen und nicht alle Ausprägungen
            var obj = currentModel.ObjDefListByTypes([Constants.OT_RISK])[couOccs];     // [SM]: Fehler bereinigt: nur Risiken im Array berücksichtigen und nicht alle Ausprägungen
            if (obj.AssignedModels().length > 0)
            {
                for (var couM=0; couM<obj.AssignedModels().length; couM++)
                {
                    if (obj.AssignedModels()[couM].TypeNum() == Constants.MT_BUSY_CONTR_DGM)
                    {
                       objOccListToCheck = obj.AssignedModels()[couM].ObjOccListBySymbol(checkTypeList);
                       for (couOcc=0; couOcc<objOccListToCheck.length; couOcc++)
                       {
                           countOccs=0; 
                           //check if this is the first occurence of the current object
                            isNewObject = helperAddObjToList(objOccListToCheck[couOcc].ObjDef()); 
                            if (isNewObject)
                            {
                                countOccs = helperCountNumberOfOccurencesInOtherModels(objOccListToCheck[couOcc], currentModel, checkModelTypes); 
                                if (countOccs > 1 || countOccs == 0)
                                {
                                    errorMessageDetails = errorMessage.replace("{0}", objOccListToCheck[couOcc].ObjDef().Name(g_usedLanguage)).replace("{1}", String(countOccs)); 
                                    LogMessage(getString("TXT_CHECKTYPEOBJECTOCCURENCES"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessageDetails, checkCategory); 
                                }
                                else
                                {
                                    errorMessageDetails = errorMessage.replace("{0}", objOccListToCheck[couOcc].ObjDef().Name(g_usedLanguage)).replace("{1}", String(countOccs)) + " - OK"; 
                                    LogMessage(getString("TXT_CHECKTYPEOBJECTOCCURENCES"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessageDetails, 5); 

                                }
                            }
                       }
                    }
                }
            }
            
        }
    }
    
}

function CheckRiskControlFunction(currentModel, checkTypeList, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    if (checkTypeList != String(Constants.ST_RISK_1) && checkTypeList != String(Constants.ST_RISK))
    {
        LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), getString("ERR_WRONGCONFIGURATIONFORRISKCONTROLCHECK"), checkCategory); 
    }
    else
    {
        var errorMessageDetails=""; 
        var objOccListToCheck = currentModel.ObjOccListBySymbol(checkTypeList); 
        if (objOccListToCheck.length > 0)
        {
            for(var cou=0; cou<objOccListToCheck.length; cou++)
            {
                var connectedFunctionOccList = objOccListToCheck[cou].getConnectedObjOccs([Constants.ST_FUNC, Constants.ST_SYS_FUNC_ACT]); 
                if (connectedFunctionOccList.length != 1)
                {
                    LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), GetText("ERR_NOUNIQUECONNECTEDFUNCTION", new Array(objOccListToCheck[cou].ObjDef().Name(g_usedLanguage))), checkCategory); 
                }
                else
                {
                    var connectedControlOccList = connectedFunctionOccList[0].getConnectedObjOccs([Constants.ST_CONTR]); 
                    // if a control is connected we have to compare the connected controls to the control which exists in the assigned bcd of the risk
                    if (connectedControlOccList.length > 0)
                    {
                       connectedControlCxns = objOccListToCheck[cou].ObjDef().CxnListFilter(Constants.EDGES_OUT, Constants.CT_IS_REDU_BY); 
                       for (var couFuncControls=0; couFuncControls<connectedControlOccList.length; couFuncControls++)
                       {
                        var isSameControl = false; 
                            for (var couRiskControls = 0; couRiskControls<connectedControlCxns.length; couRiskControls++)
                            {
                                if (connectedControlCxns[couRiskControls].TargetObjDef().IsEqual(connectedControlOccList[couFuncControls].ObjDef()))
                                {
                                    isSameControl=true; 
                                    break; 
                                }
                            }
                            if (isSameControl)
                            {
                                errorMessageDetails = errorMessage.replace("{0}", objOccListToCheck[cou].ObjDef().Name(g_usedLanguage)).replace("{1}", connectedControlCxns[couRiskControls].TargetObjDef().Name(g_usedLanguage)); 
                                LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessageDetails, checkCategory); 
                            }
                            else
                            {
                                //errorMessageDetails = errorMessage.replace("{0}", objOccListToCheck[cou].ObjDef().Name(g_usedLanguage)).replace("{1}", connectedControlCxns[couRiskControls].TargetObjDef().Name(g_usedLanguage)) + " - OK"; // [SM] Target-Objekte kann es in diesem Fall nicht geben... 
                                errorMessageDetails = errorMessage.replace("{0}", objOccListToCheck[cou].ObjDef().Name(g_usedLanguage)).replace("{1}", connectedControlOccList[couFuncControls].ObjDef().Name(g_usedLanguage)) + " - OK"; 
                                LogMessage(getString("TXT_CHECKTYPERISKCONTROL"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessageDetails, 5); 

                            }
                       }
                    }
                }
            }
        }
    }
    
}

function CheckObjectLocationInProcessGroup(currentModel, checkTypeList, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var checkTypes = new Array(); 
    checkTypes = checkTypeList.split(",");
    var objectsToCheck; 
    for (var couT=0; couT<checkTypes.length; couT++)
    {
        checkTypes[couT] = helperConvertTypeGuidToTypeNum(checkTypes[couT], "CID_OBJOCC");
        var objOccListToCheck = currentModel.ObjOccListBySymbol(checkTypes[couT]); 
        objectsToCheck = new Array(); 
        for (var couO=0; couO<objOccListToCheck.length; couO++)
        {
            if (checkTypes[couT] == Constants.ST_PRCS_IF)
            {
                if (objOccListToCheck[couO].ObjDef().Name(1031) == "Prozessende")
                    objectsToCheck.push(objOccListToCheck[couO].ObjDef()); 
            }
            else
                objectsToCheck.push(objOccListToCheck[couO].ObjDef()); 
        }
        
        objectsToCheck = ArisData.Unique(objectsToCheck); 
        var detailErrorMessage = errorMessage; 
        for (var cou=0; cou<objectsToCheck.length; cou++)
        {
            var obj = objectsToCheck[cou]; 
            var objPath = obj.Group().Path(1031); 
            if (objPath.indexOf("2. Prozesssicht")> 0)
            {
                detailErrorMessage = errorMessage.replace("{0}", obj.Name(g_usedLanguage)).replace("{1}", obj.Type()); 
                LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), detailErrorMessage, checkCategory); 
            }
            else
            {
                detailErrorMessage = errorMessage.replace("{0}", obj.Name(g_usedLanguage)).replace("{1}", obj.Type())+ " - OK"; 
                LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), detailErrorMessage, 5); 
            }
        }
    }
   
}

function checkDefaultSymbolType(currentModel, checkTypeList, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var checkTypes = new Array(); 
    checkTypes = checkTypeList.split(",");
    var objectsToCheck; 
    for (var couT=0; couT<checkTypes.length; couT++)
    {
        checkTypes[couT] = helperConvertTypeGuidToTypeNum(checkTypes[couT], "CID_OBJOCC");
        var objOccListToCheck = currentModel.ObjOccListBySymbol(checkTypes[couT]); 
        objectsToCheck = new Array(); 
        for (var couO=0; couO<objOccListToCheck.length; couO++)
        {
            objectsToCheck.push(objOccListToCheck[couO].ObjDef()); 
        }
        
        objectsToCheck = ArisData.Unique(objectsToCheck); 
        var errorMessage_Details=""; 
        var obj = null; 
        for (couO=0; couO<objectsToCheck.length; couO++)
        {
            obj = objectsToCheck[couO]; 
            if (obj.getDefaultSymbolNum() != checkTypes[couT])
            {
                errorMessage_Details = errorMessage.replace("{0}", obj.Name(g_usedLanguage)).replace("{1}", String(obj.getDefaultSymbolNum())); 
                LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessage_Details, checkCategory); 
            }
            else
            {
                errorMessage_Details = errorMessage.replace("{0}", obj.Name(g_usedLanguage)).replace("{1}", String(obj.getDefaultSymbolNum()))+ " - OK"; 
                LogMessage(getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKTYPEOBJECTCONNECTIONS"), errorMessage_Details, 5); 
            }
        }
    }
}
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///  Conventions check for Connections and connection attributes in the flow model
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//check all rules with ruleKind = "Cxns"
function doConventionsCheckConnections(cxnRules, currentModel)
{
  for (var cou = 0; cou < cxnRules.size(); cou++)
   {
        var ruleXmlNode = cxnRules.get(cou); 
        var attrTypeNum = g_config.getAttributeTypeNum(ruleXmlNode); 
        var checkItemKind    = ruleXmlNode.getAttributeValue("kind");
        var checkTypeList    = ruleXmlNode.getAttributeValue("typelist");
        var checkCategory    = g_config.getCheckCategory(ruleXmlNode);
        var ruleDescription   = g_config.getRuleDescription(ruleXmlNode);
        var errorMessage     = g_config.getErrorMessage(ruleXmlNode); 
        var ruleShortcut     = g_config.getRuleShortcut(ruleXmlNode); 
        var allowedFuncSymbolTypes; 
        var attrCompany = 0; 
        var attrValue = ""; 
        
        if (checkCategory <= g_currentLogType)
        {
            if (attrTypeNum == 0)
            {
                var cxnRuleXmlNode = g_config.getCxnRule(ruleXmlNode); 
                checkType        = String(g_config.getRuleType(cxnRuleXmlNode)); // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
            
                switch (checkType)
                {
                case String("CheckNumberOfCxns"): // SP: The compared value must be a JS object, not a Java object - like java.lang.String!
                    
                    var targetSymbolTypes = cxnRuleXmlNode.getAttributeValue("targetSymbolTypes"); 
                    CheckNumberOfCxnsForEachFunction(currentModel, checkTypeList, targetSymbolTypes, checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                    break; 
                case String("CheckNumberOfCxns_ExternalPerson"):
                    allowedFuncSymbolTypes = new Array(); 
                    allowedFuncSymbolTypes[0] = g_methodFilter.UserDefinedSymbolTypeNum("30619aa0-ef1b-11db-262b-000d607a2205"); 
                    CheckNumberOfCxnsExternalRessource(currentModel, checkTypeList, allowedFuncSymbolTypes, 0, "", checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                    break; 
                case String("CheckNumberOfCxns_ExternalOrgUnits"):
                case String("CheckNumberOfCxns_ExternalPersonType"):
                    allowedFuncSymbolTypes = new Array(); 
                    allowedFuncSymbolTypes[0] = g_methodFilter.UserDefinedSymbolTypeNum("30619aa0-ef1b-11db-262b-000d607a2205");  
                    attrCompany = g_methodFilter.UserDefinedAttributeTypeNum("317e1987-7df8-11da-0c60-cf8e338f9a0b");  
                    attrValue = "UniCredit Bank AG"; 
                    CheckNumberOfCxnsExternalRessource(currentModel, checkTypeList, allowedFuncSymbolTypes, attrCompany, attrValue,  checkCategory, ruleDescription, errorMessage, ruleShortcut); 
                    break; 
                default:
                
                    
                }
            }
            else
            {
                var attrRule = g_config.getAttrRuleNode(ruleXmlNode);
                var checkTypes = checkTypeList.split(","); 
                for (var couT=0; couT<checkTypes.length; couT++)
                {
                    var cxnOccList = currentModel.CxnOccListFilter(checkTypes[couT]);    
                    for (var couC=0; couC<cxnOccList.length; couC++)
                    {
                        if (cxnOccList[couC].TargetObjOcc().SymbolNum() == Constants.ST_FUNC || cxnOccList[couC].TargetObjOcc().SymbolNum() == Constants.ST_FUNC_ACT) //only for connected functions needed
                            doAttributeCheck(cxnOccList[couC].CxnDef(), ruleXmlNode, attrRule, null); 
                    }
                }
                
            }
        }
            
        
    }
}

function CheckNumberOfCxnsForEachFunction(currentModel, cxnTypeList, targetSymbolTypes, checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var cxnTypes = new Array(); 
    
    cxnConfTypes = cxnTypeList.split(","); 
    
    var targetSymbolTypeList = new Array(); 
    if (targetSymbolTypeList.indexOf(",")> -1)
    {
        targetSymbolTypeList = targetSymbolTypes.split(","); 
        for (var couS=0; couS<targetSymbolTypeList.length; couS++)
            targetSymbolTypeList[couS] = helperConvertTypeGuidToTypeNum(targetSymbolTypeList[couS], "CID_OBJOCC"); 
    }
    else
        targetSymbolTypeList[0] = helperConvertTypeGuidToTypeNum(targetSymbolTypes, "CID_OBJOCC"); 
    
    for (var couT=0; couT<cxnConfTypes.length; couT++)
    {
        cxnTypes.push(__toNumeric(cxnConfTypes[couT])); 
    }
    
    var errorMessageDetails = ""; 
    for (couS=0; couS<targetSymbolTypeList.length; couS++)
    {
        for (var couO=0; couO<currentModel.ObjOccListFilter(-1, targetSymbolTypeList[couS]).length; couO++)
        {
            var func = currentModel.ObjOccListFilter(-1, targetSymbolTypeList[couS])[couO].ObjDef(); 
            if (func.CxnListFilter(Constants.EDGES_INOUT, cxnTypes).length ==0)
            {
                errorMessageDetails = errorMessage.replace("{0}", func.Name(g_usedLanguage)); 
                LogMessage(getString("TXT_CHECKCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKCONNECTIONS"), errorMessageDetails, checkCategory); 
            }
        }
    }
    
}

function CheckNumberOfCxnsExternalRessource(currentModel, symbolTypeList, allowedFuncSymbolTypes, attrCompany, unallowedAttrValue,  checkCategory, ruleDescription, errorMessage, ruleShortcut)
{
    var symbolTypes = new Array(); 
    if (symbolTypeList.indexOf(",")> 0)
    {
        symbolTypes = symbolTypeList.split(","); 
        for (var couS=0; couS<symbolTypes.length; couS++)
        {
            symbolTypes[couS] = helperConvertTypeGuidToTypeNum(symbolTypes[couS], "CID_OBJOCC"); 
        }
    }
    else
        symbolTypes[0] = helperConvertTypeGuidToTypeNum(symbolTypeList, "CID_OBJOCC"); 
    
    for (var cou=0; cou<symbolTypes.length; cou++)
    {
        var resourceOccList = currentModel.ObjOccListFilter(-1, symbolTypes[cou]); 
        if (attrCompany != 0)
        {
         resourceOccList = helperGetExternalRessources(resourceOccList, attrCompany, unallowedAttrValue); 
        }
        var errorMessageDetails= ""; 
        for (var couO=0; couO<resourceOccList.length; couO++)
        {
            var cxnList = resourceOccList[couO].Cxns(Constants.EDGES_OUT, Constants.EDGES_ALL); 
            for (var couC=0; couC<cxnList.length; couC++)
            {
                if (cxnList[couC].CxnDef().TypeNum() == Constants.CT_EXEC_1 || cxnList[couC].CxnDef().TypeNum() == Constants.CT_EXEC_2)
                {
                 if ((!(IsInArray(allowedFuncSymbolTypes, cxnList[couC].TargetObjOcc().SymbolNum()))) && currentModel.Group().Path(g_usedLanguage).substring(0,14) != "Hauptgruppe\\8.")
                 {
                    errorMessageDetails = errorMessage.replace("{0}", cxnList[couC].TargetObjOcc().ObjDef().Name(g_usedLanguage)).replace("{1}", cxnList[couC].SourceObjOcc().ObjDef().Name(g_usedLanguage));  
                    LogMessage(getString("TXT_CHECKCONNECTIONS"), ruleShortcut, ruleDescription, getString("TXT_CHECKCONNECTIONS"),  errorMessageDetails, checkCategory);  
                 }
                }
            }
        }
    }
    
}
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///  HELPER
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function helperCheckRegExp(regExp, valueToCheck)
{
    try
    {


        //if (regExp.test(valueToCheck))
        //{
            var regExpResult = valueToCheck.match(regExp);

            if (regExpResult != null)
            {
             if (regExpResult.length > 0)
             {
                if (regExpResult[0].length != valueToCheck.length()) // nur ein Teil des Strings entspricht der Vorgabe des RegExp´s --> Fehler zurückgeben
                {
                    return false; 
                
                }
                else
                    return true; //das Ergebnis des RegExpResult ist ein Array, dessen erstes Element die gleiche Länge hat wie der ursprüngliche String --> das Format des Strings ist korrekt
             } 
             else
                return false; 

            }
            else 

                return false; // zu prüfender String passt nicht zum Regular Expression, es wird kein Array als Ergebnis zurückgeliefert. 
       
       // }
       // else
       //     return true; 
    }
    catch (err)
    {
        return false; 
    }
}

//some resources (org units, person types) are marked as external with a special attribute value (e.g. "company" != "Unicredit Bank AG")
//resourceOccList: the list of resource occs in one process (one symbol type)
//companyAttr: the attr type num which is used as label for external or internal
//internalAttrValue: the attr value which labels a resource as internal 
function helperGetExternalRessources(resourceOccList, companyAttr, internalAttrValue)
{
    var newResourceOccList = new Array(); 
    
    for (var cou=0; cou<resourceOccList.length; cou++)
    {
        if (resourceOccList[cou].ObjDef().Attribute(companyAttr, 1031).getValue() != internalAttrValue)
            newResourceOccList.push(resourceOccList[cou]); 
    }
    
    return newResourceOccList; 
}
//add an object def to a list and check for uniqueness
//an obj def can occure more than once while a process check within the current model or within assigned models (e.g. in epc and in assigned bcds)
//it is important to check it only once
function helperAddObjToList(objDef)
{
    var countObjs = g_objsToCheck.length; 
    g_objsToCheck.push(objDef); 
    g_objsToCheck = ArisData.Unique(g_objsToCheck); 
    if (g_objsToCheck.length > countObjs)
        return true; 
    else 
        return false; 
    
}
function helperCountNumberOfOccurencesInOtherModels(objOcc, currentModel, allowedModelTypes)
{
    var countOccs = 0; 
    for (var couO=0; couO<objOcc.ObjDef().OccList().length; couO++)
    {
        var curOcc =     objOcc.ObjDef().OccList()[couO]; 
        if (!(curOcc.Model().IsEqual(currentModel)) && IsInArray(allowedModelTypes, curOcc.Model().TypeNum()))
            countOccs++; 
    }
    return countOccs; 
}
//Convert type GUID to type number or return input unchanged if conversion fails.
// inputType: The type GUID to convert or the type number to keep unchanged.
// itemKind: The item kind to apply on 'inputType' conversion (CID_ATTRDEF, CID_OBJOCC or CID_MODEL).
// returns: The type number.
function helperConvertTypeGuidToTypeNum(inputType, itemKind)
{
    if (String(inputType).indexOf("{") >= 0)
    {
        var typeGUID = String(inputType).replace("{", "").replace("}", "");
        if (itemKind == "CID_ATTRDEF")
        {
              if (g_methodFilter.UserDefinedAttributeTypeNum(typeGUID) != -1)
                return g_methodFilter.UserDefinedAttributeTypeNum(typeGUID);
        }
        else if (itemKind == "CID_OBJOCC")
        {
              if (g_methodFilter.UserDefinedSymbolTypeNum(typeGUID) != -1)
                return g_methodFilter.UserDefinedSymbolTypeNum(typeGUID);
        }
        else if (itemKind == "CID_MODEL")
        {
              if (g_methodFilter.UserDefinedModelTypeNum(typeGUID) != -1)
                return g_methodFilter.UserDefinedModelTypeNum(typeGUID);
        }
    }
    return inputType;
}

//Find the conterpart object definitions, that are connected with indicated object occurency within object occurencie's model.
//This is circumstantial, ugly, time- and memory consuming compared to 'currentObjOcc.getConnectedObjOccs()'! But this is the only way MATRIX models can be evaluated!
// currentObjOcc: The object occurency  to find connected object definitions for, that have occurences within object occurencie's model.
// returns: The conterpart object definition array.
function helperGetConnectedObjDefListForObjOcc(currentObjOcc)
{
    var realCounterparts = new Array();
    var potentialCounterparts = currentObjOcc.Model().ObjDefList();
    for (var potentialCounterpartsIndex = 0; potentialCounterpartsIndex < potentialCounterparts.length; potentialCounterpartsIndex++)
    {
        var reverseConnectedObjects = potentialCounterparts[potentialCounterpartsIndex].getConnectedObjs(currentObjOcc.ObjDef().TypeNum());
        for (var reverseConnectedObjectsIndex = 0; reverseConnectedObjectsIndex < reverseConnectedObjects.length; reverseConnectedObjectsIndex++)
        {
           if (reverseConnectedObjects[reverseConnectedObjectsIndex].GUID() == currentObjOcc.ObjDef().GUID())
               realCounterparts.push(potentialCounterparts[potentialCounterpartsIndex]);
        }
    }
    potentialCounterparts = null;
    return realCounterparts;
}

//Determine those matrix cells, that connect 'definingObjDef' with any of the 'possibleCounterpartObjDefs' by a connection of type 'connectionType'.
// currentMatrix: The matrix, the requested cells are contained in.
// definingObjDef: The object that must be connected.
// possibleCounterpartObjDefs: The possible couterpart objects to connect to.
// connectionType: The type, 'definingObjDef' must be connected to any of the 'possibleCounterpartObjDefs'.
// returns: The matrix cells array.
function helperGetMatrixCells(currentMatrix, definingObjDef, possibleCounterpartObjDefs, connectionType)
{
    var matchingMatrixContentCells = new Array();
    var matrixContentCells = currentMatrix.getContentCells();
    for (var index = 0; index < possibleCounterpartObjDefs.length; index++)
    {
        for (var matrixContentCellIndex = 0; matrixContentCellIndex < matrixContentCells.length; matrixContentCellIndex++)
        {
            if (matrixContentCells[matrixContentCellIndex].getColumnHeader().getDefinition() != null && matrixContentCells[matrixContentCellIndex].getRowHeader().getDefinition() != null)
            {
                if ((matrixContentCells[matrixContentCellIndex].getColumnHeader().getDefinition().GUID() == possibleCounterpartObjDefs[index].GUID() &&
                     matrixContentCells[matrixContentCellIndex].getRowHeader().getDefinition().GUID() == definingObjDef.GUID()) ||
                    (matrixContentCells[matrixContentCellIndex].getColumnHeader().getDefinition().GUID() == definingObjDef.GUID() &&
                     matrixContentCells[matrixContentCellIndex].getRowHeader().getDefinition().GUID() == possibleCounterpartObjDefs[index].GUID()))
                {
                    var cellCxnList = matrixContentCells[matrixContentCellIndex].getCxns();
                    for (var cellCxnIndex = 0; cellCxnIndex < cellCxnList.length; cellCxnIndex++)
                    {
                        if (cellCxnList[cellCxnIndex].TypeNum() == connectionType)
                        {
                            matchingMatrixContentCells.push( matrixContentCells[matrixContentCellIndex]);
                            break;
                        }
                    }
                }
            }
            
        }
    }
    return matchingMatrixContentCells;
}

//Determine those counterpart object definitions of the 'possibleCounterpartObjDefs', that are connected to 'definingObjDef' by a connection of type 'connectionType'.
// currentMatrix: The matrix, the requested cells are contained in.
// definingObjDef: The object that must be connected.
// possibleCounterpartObjDefs: The possible couterpart objects to connect to.
// connectionType: The type, 'definingObjDef' must be connected to any of the 'possibleCounterpartObjDefs'.
// returns: The connected counterpart object definitions.
function helperGetMatrixConnectedCounterparts(currentMatrix, definingObjDef, possibleCounterpartObjDefs, connectionType)
{
    var matchingCounterpartObjDefs = new Array();
    var matrixContentCells = currentMatrix.getContentCells();
    for (var index = 0; index < possibleCounterpartObjDefs.length; index++)
    {
        for (var matrixContentCellIndex = 0; matrixContentCellIndex < matrixContentCells.length; matrixContentCellIndex++)
        {
            if (matrixContentCells[matrixContentCellIndex].getColumnHeader().getDefinition() != null && matrixContentCells[matrixContentCellIndex].getRowHeader().getDefinition() != null)
            {
                if ((matrixContentCells[matrixContentCellIndex].getColumnHeader().getDefinition().GUID() == possibleCounterpartObjDefs[index].GUID() &&
                     matrixContentCells[matrixContentCellIndex].getRowHeader().getDefinition().GUID() == definingObjDef.GUID()) ||
                    (matrixContentCells[matrixContentCellIndex].getColumnHeader().getDefinition().GUID() == definingObjDef.GUID() &&
                     matrixContentCells[matrixContentCellIndex].getRowHeader().getDefinition().GUID() == possibleCounterpartObjDefs[index].GUID()))
                {
                    var cellCxnList = matrixContentCells[matrixContentCellIndex].getCxns();
                    for (var cellCxnIndex = 0; cellCxnIndex < cellCxnList.length; cellCxnIndex++)
                    {
                        if (cellCxnList[cellCxnIndex].TypeNum() == connectionType)
                        {
                            matchingCounterpartObjDefs.push(possibleCounterpartObjDefs[index]);
                            break;
                        }
                    }
                }
            }
        }
    }
    return matchingCounterpartObjDefs;
}

//Test whether the 'objDef' name matches with any 'objDefArray' name.
// objDef: Item who's name is to test.
// objDefArray: Array of items to test for a matching name.
// returns: true on success, or false otherwise.
function helpNameMatch(objDef, objDefArray)
{
    for (var index = 0; index < objDefArray.length; index++)
        if (objDef.Name(g_usedLanguage, true) == objDefArray[index].Name(g_usedLanguage, true))
            return true;
    return false;
}

function getLogTypeForAttributeItemKind(currentItem)
{
    var logType=""; 
    if (currentItem.KindNum() == Constants.CID_MODEL)
        logType =  getString("TXT_CHECKTYPEMODELATTRIBUTE");
    else if (currentItem.KindNum() == Constants.CID_OBJDEF)
        logType = getString("TXT_CHECKTYPEOBJECTATTRIBUTE"); 
    else if (currentItem.KindNum() == Constants.CID_CXNDEF)
        logType = getString("TXT_CHECKTYPECXNATTRIBUTE"); 
        
    return logType;
}

function GetText(textId, replacementTexts)
{
     var retVal = getString(textId);
     
     //abort on invalid textId  (7.0.x -> ???   7.1.x -> String not found...)
     if (retVal=="???" || retVal.substring(0,16)=="String not found")
     {
         return "[" + textId + "]";
     }
     
     //replace wildcards
     if (replacementTexts != null)
     {
         for (var cou=0; cou < replacementTexts.length; cou++)
         {
             retVal = retVal.replace(eval("/\\{"+cou+"\\}/g"), replacementTexts[cou]);
         }
     }
     
     //return final string
     return retVal;
}

function createFileName(modelName)
{
    var currentDate = new Date();
    modelName = encodefilename(modelName); 
    var fileName = ""; 
     fileName = Common_GetRightSubstring(__toString(currentDate.getFullYear()), 2) +
         Common_GetRightSubstring("00" + __toString(currentDate.getMonth() + 1), 2) +
         Common_GetRightSubstring("00" + __toString(currentDate.getDate()), 2) +
         Common_GetRightSubstring("00" + __toString(currentDate.getHours()), 2) +
         Common_GetRightSubstring("00" + __toString(currentDate.getMinutes()), 2) + "_" +
         "QS_" + encodefilename(modelName) +
         ".xlsx";
    return fileName;     
}


// <summary> Remove characters, not suitable for a file name. </summary>
// <param name="sFile"> The raw file name. </param>
// <returns> The suitable filename. </returns>
// [SM]: wird auch benutzt für E-Mail-encoding, da ähnliche Regeln gelten !!!
function encodefilename(sfile)
{
  var stmp = "";

  stmp = __toString(vbReplace(sfile, " ", "_"));
  stmp = __toString(vbReplace(stmp, "/", "_"));
  stmp = __toString(vbReplace(stmp, "+", "_"));
  stmp = __toString(vbReplace(stmp, "§", "_"));
  stmp = __toString(vbReplace(stmp, "$", "_"));
  stmp = __toString(vbReplace(stmp, "%", "_"));
  stmp = __toString(vbReplace(stmp, "&", "_"));
  stmp = __toString(vbReplace(stmp, "‘", "_"));
  stmp = __toString(vbReplace(stmp, "…", "_"));

  stmp = __toString(vbReplace(stmp, "Ä", "Ae"));
  stmp = __toString(vbReplace(stmp, "ä", "ae"));
  stmp = __toString(vbReplace(stmp, "Ö", "Oe"));
  stmp = __toString(vbReplace(stmp, "ö", "oe"));
  stmp = __toString(vbReplace(stmp, "Ü", "Ue"));
  stmp = __toString(vbReplace(stmp, "ü", "ue"));
  stmp = __toString(vbReplace(stmp, "ß", "ss"));

  stmp = __toString(vbReplace(stmp, "\\", "_"));
  stmp = __toString(vbReplace(stmp, ":", "."));
  stmp = __toString(vbReplace(stmp, "*", "."));
  stmp = __toString(vbReplace(stmp, "?", "."));
  stmp = __toString(vbReplace(stmp, "\"", "\'"));
  stmp = __toString(vbReplace(stmp, "‘", "\'"));
  stmp = __toString(vbReplace(stmp, "<", "-"));
  stmp = __toString(vbReplace(stmp, ">", "-"));
  stmp = __toString(vbReplace(stmp, "|", "-"));

  stmp = __toString(vbReplace(stmp, String.fromCharCode(8211), "-"));

  stmp = __toString(vbReplace(stmp, vbCr, ""));
  stmp = __toString(vbReplace(stmp, vbLf, ""));

  return stmp;
}

//helper function to find an element in an array
function IsInArray(arr, el){
    for (var cou=0; cou<arr.length; cou++){
        if (el == arr[cou])
            return true; 
    }
    return false; 
}

////////////////////////////////////////////////////////////////////////////////////////////////////////
// Excel Helper functions
////////////////////////////////////////////////////////////////////////////////////////////////////////
function writeHeader(sheet, xlworkbook, model)
{
    var cs = xlworkbook.createCellStyle(); //new org.apache.poi.xssf.usermodel.CellStyle();
    cs.setFont(g_boldFont); 
    cs.setBorderBottom(1); 
    cs.setBorderTop(1); 
    cs.setBorderLeft(1); 
    cs.setBorderRight(1);
    cs.setRightBorderColor(0); 
    cs.setLeftBorderColor(0); 
    cs.setTopBorderColor(0); 
    cs.setBottomBorderColor(0); 
    cs.setFillPattern(Constants.NO_FILL); 
    cs.setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
    cs.setHidden(false); 
    cs.setIndention(0); 
    cs.setLocked(false); 
    cs.setRotation(0); 
    cs.setWrapText(false); 
    var headerRow = sheet.createRow(0);
    
    var c; 
    for (var col = 0; col < 7; col++)
    {
        c = headerRow.createCell(col);
        c.setCellStyle(cs); 
       
    }
    headerRow.getCell(0).setCellValue(""); 
    headerRow.getCell(1).setCellValue(model.Name(g_usedLanguage));
    g_rowNumber++; 
    
    var headerRow = sheet.createRow(1); 
    var cs = xlworkbook.createCellStyle(); //new org.apache.poi.xssf.usermodel.CellStyle();
    cs.setFont(g_boldFont); 
    cs.setBorderBottom(1); 
    cs.setBorderTop(1); 
    cs.setBorderLeft(1); 
    cs.setBorderRight(1);
    cs.setRightBorderColor(0); 
    cs.setLeftBorderColor(0); 
    cs.setTopBorderColor(0); 
    cs.setBottomBorderColor(0); 
    cs.setFillPattern(Constants.NO_FILL); 
    cs.setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
    cs.setHidden(false); 
    cs.setIndention(0); 
    cs.setLocked(false); 
    cs.setRotation(0); 
    cs.setWrapText(true); 
    var c; 
    for (var col = 0; col < 7; col++)
    {
        c = headerRow.createCell(col);
        c.setCellStyle(cs); 
       
    }

    headerRow.getCell(0).setCellValue(getString("TXT_SHORTCUT")); 
    headerRow.getCell(1).setCellValue(getString("TXT_DETAILLEVEL"));
    headerRow.getCell(2).setCellValue(getString("TXT_CHECK"));
    headerRow.getCell(3).setCellValue(getString("TXT_CHECKCONTEXT"));
    headerRow.getCell(4).setCellValue(getString("TXT_QSDESCRIPTION"));
    headerRow.getCell(5).setCellValue(getString("TXT_QSMESSAGE"));
    headerRow.getCell(6).setCellValue(getString("TXT_COMMENT"));  
    g_rowNumber++; 
    
    sheet.setColumnWidth(0, 10*256); 
    sheet.setColumnWidth(1, 10*256); 
    sheet.setColumnWidth(2, 20*256); 
    sheet.setColumnWidth(3, 20*256); 
    sheet.setColumnWidth(4, 60*256);
    sheet.setColumnWidth(5, 40*256); 
    sheet.setColumnWidth(6, 40*256); 

    sheet.setColumnHidden(0, true); 
    sheet.setColumnHidden(4, true);     
}

function createExcelOutputFile()
{
    var newWorkbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(); 
    return newWorkbook; 
}

// <summary>
// Prepares the worksheet based on the specified mode.
// Headers and headlines will be created, and ranges will be formatted.
// </summary>
function prepareWorksheet(xlworksheet, xlworkbook, rowDataRecordCount, model)
{
    xlworkbook.setSheetName(g_sheetCount, ((g_sheetCount + 1) + "-" + model.Name(g_usedLanguage).replace("/", "-")));
    g_sheetCount++;

    //pagesetup.zoom = false;
    xlworksheet.setZoom(1, 1);
    g_rowNumber = 0; 
    
    // CREATE REQURED CELLS AND SET STANDARD CELL STYLE
    {
        g_stndFont = xlworkbook.createFont();
        g_stndFont.setFontHeightInPoints(8);
        g_stndFont.setFontName("Arial");
        g_stndFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        //g_stndFont.setColor(Constants.C_BLACK);
        // REMARK: Setting the font color is not possible as the documented constants are not available

        g_boldFont = xlworkbook.createFont();
        g_boldFont.setFontHeightInPoints(8);
        g_boldFont.setFontName("Arial");
        g_boldFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
        
        
        var fontDetail = xlworkbook.createFont();
        fontDetail.setFontHeightInPoints(8);
        fontDetail.setFontName("Arial");
        fontDetail.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        
        var fontColor = new java.awt.Color(vbRgb(146,208,80)); 
        var xssfColor = new org.apache.poi.xssf.usermodel.XSSFColor(fontColor); 
        fontDetail.setColor(xssfColor); 
        
        var fontInfo = xlworkbook.createFont();
        fontInfo.setFontHeightInPoints(8);
        fontInfo.setFontName("Arial");
        fontInfo.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        fontColor = new java.awt.Color(vbRgb(146,208,80)); 
        xssfColor = new org.apache.poi.xssf.usermodel.XSSFColor(fontColor); 
        fontInfo.setColor(xssfColor); 
        
        var fontNotice = xlworkbook.createFont();
        fontNotice.setFontHeightInPoints(8);
        fontNotice.setFontName("Arial");
        fontNotice.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        fontColor = new java.awt.Color(vbRgb(0,176,80)); 
        xssfColor = new org.apache.poi.xssf.usermodel.XSSFColor(fontColor); 
        fontNotice.setColor(xssfColor); 
        
        var fontWarning = xlworkbook.createFont();
        fontWarning.setFontHeightInPoints(8);
        fontWarning.setFontName("Arial");
        fontWarning.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        fontColor = new java.awt.Color(vbRgb(255,153,0)); 
        xssfColor = new org.apache.poi.xssf.usermodel.XSSFColor(fontColor); 
        fontWarning.setColor(xssfColor); 
        
        var fontError = xlworkbook.createFont(); 
        fontError.setFontHeightInPoints(8);
        fontError.setFontName("Arial");
        fontError.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        fontColor = new java.awt.Color(vbRgb(255,0,0));  
        xssfColor = new org.apache.poi.xssf.usermodel.XSSFColor(fontColor); 
        fontError.setColor(xssfColor); 
        
        var fontCritical = xlworkbook.createFont(); 
        fontCritical.setFontHeightInPoints(8);
        fontCritical.setFontName("Arial");
        fontCritical.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        fontColor = new java.awt.Color(vbRgb(192,0,0)); 
        xssfColor = new org.apache.poi.xssf.usermodel.XSSFColor(fontColor); 
        fontCritical.setColor(xssfColor); 
        
        g_cs = g_xlWorkbook.createCellStyle(); //new org.apache.poi.xssf.usermodel.CellStyle();
        g_cs.setFont(g_stndFont); 
        g_cs.setBorderBottom(1); 
        g_cs.setBorderTop(1); 
        g_cs.setBorderLeft(1); 
        g_cs.setBorderRight(1);
        g_cs.setRightBorderColor(0); 
        g_cs.setLeftBorderColor(0); 
        g_cs.setTopBorderColor(0); 
        g_cs.setBottomBorderColor(0); 
        g_cs.setFillPattern(Constants.NO_FILL); 
        g_cs.setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
        g_cs.setHidden(false); 
        g_cs.setIndention(0); 
        g_cs.setLocked(false); 
        g_cs.setRotation(0); 
        g_cs.setWrapText(true); 
        
        var cs_critical = g_xlWorkbook.createCellStyle(); //new org.apache.poi.xssf.usermodel.CellStyle();
        with (cs_critical)
        {
            setFont(fontCritical); 
            setBorderBottom(1); 
            setBorderTop(1); 
            setBorderLeft(1); 
            setBorderRight(1);
            setRightBorderColor(0); 
            setLeftBorderColor(0); 
            setTopBorderColor(0); 
            setBottomBorderColor(0); 
            setFillPattern(Constants.NO_FILL); 
            setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
            setHidden(false); 
            setIndention(0); 
            setLocked(false); 
            setRotation(0); 
            setWrapText(true);
        }
        
        var cs_error = g_xlWorkbook.createCellStyle(); //new org.apache.poi.xssf.usermodel.CellStyle();
        with (cs_error)
        {
            setFont(fontError); 
            setBorderBottom(1); 
            setBorderTop(1); 
            setBorderLeft(1); 
            setBorderRight(1);
            setRightBorderColor(0); 
            setLeftBorderColor(0); 
            setTopBorderColor(0); 
            setBottomBorderColor(0); 
            setFillPattern(Constants.NO_FILL); 
            setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
            setHidden(false); 
            setIndention(0); 
            setLocked(false); 
            setRotation(0); 
            setWrapText(true); 
        }
        var cs_warning = g_xlWorkbook.createCellStyle(); 
        with (cs_warning)
        {
            setFont(fontWarning); 
            setBorderBottom(1); 
            setBorderTop(1); 
            setBorderLeft(1); 
            setBorderRight(1);
            setRightBorderColor(0); 
            setLeftBorderColor(0); 
            setTopBorderColor(0); 
            setBottomBorderColor(0); 
            setFillPattern(Constants.NO_FILL); 
            setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
            setHidden(false); 
            setIndention(0); 
            setLocked(false); 
            setRotation(0); 
            setWrapText(true); 
        }
        
        var cs_info = g_xlWorkbook.createCellStyle(); 
        with (cs_info)
        {
            setFont(fontInfo); 
            setBorderBottom(1); 
            setBorderTop(1); 
            setBorderLeft(1); 
            setBorderRight(1);
            setRightBorderColor(0); 
            setLeftBorderColor(0); 
            setTopBorderColor(0); 
            setBottomBorderColor(0); 
            setFillPattern(Constants.NO_FILL); 
            setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
            setHidden(false); 
            setIndention(0); 
            setLocked(false); 
            setRotation(0); 
            setWrapText(true); 
        }
        
        var cs_detail = g_xlWorkbook.createCellStyle(); 
        with (cs_detail)
        {
            setFont(fontDetail); 
            setBorderBottom(1); 
            setBorderTop(1); 
            setBorderLeft(1); 
            setBorderRight(1);
            setRightBorderColor(0); 
            setLeftBorderColor(0); 
            setTopBorderColor(0); 
            setBottomBorderColor(0); 
            setFillPattern(Constants.NO_FILL); 
            setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
            setHidden(false); 
            setIndention(0); 
            setLocked(false); 
            setRotation(0); 
            setWrapText(true); 
        }
        
        var cs_notice = g_xlWorkbook.createCellStyle(); 
        with (cs_notice)
        {
            setFont(fontNotice); 
            setBorderBottom(1); 
            setBorderTop(1); 
            setBorderLeft(1); 
            setBorderRight(1);
            setRightBorderColor(0); 
            setLeftBorderColor(0); 
            setTopBorderColor(0); 
            setBottomBorderColor(0); 
            setFillPattern(Constants.NO_FILL); 
            setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
            setHidden(false); 
            setIndention(0); 
            setLocked(false); 
            setRotation(0); 
            setWrapText(true); 
        }
        

        g_loggerClasses[0].cellStyle = cs_critical; 
        g_loggerClasses[1].cellStyle = cs_error; 
        g_loggerClasses[2].cellStyle = cs_warning; 
        g_loggerClasses[3].cellStyle = cs_notice; 
        g_loggerClasses[4].cellStyle = cs_info; 
        g_loggerClasses[5].cellStyle = cs_detail; 
        
        
        //cell style for outptu of hyperlinks
        g_hlinkStyle = g_xlWorkbook.createCellStyle();
        var hlinkfont = g_xlWorkbook.createFont();
        hlinkfont.setUnderline(1);
        hlinkfont.setFontHeightInPoints(8);
        hlinkfont.setFontName("Arial");
        hlinkfont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        
        fontColor = new java.awt.Color(vbRgb(0,0,255));  
        xssfColor = new org.apache.poi.xssf.usermodel.XSSFColor(fontColor); 
        hlinkfont.setColor(xssfColor); 
        with (g_hlinkStyle)
        {
            setFont(hlinkfont); 
            setBorderBottom(1); 
            setBorderTop(1); 
            setBorderLeft(1); 
            setBorderRight(1);
            setRightBorderColor(0); 
            setLeftBorderColor(0); 
            setTopBorderColor(0); 
            setBottomBorderColor(0); 
            setFillPattern(Constants.NO_FILL); 
            setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
            setHidden(false); 
            setIndention(0); 
            setLocked(false); 
            setRotation(0); 
            setWrapText(true); 
        }

    }

    writeHeader(xlworksheet, xlworkbook, model);
}

function createResultSheet(totals)
{
    var titlFont = g_xlWorkbook.createFont();
        titlFont.setFontHeightInPoints(36);
        titlFont.setFontName("Arial");
        titlFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
        
    var titlCS = g_xlWorkbook.createCellStyle(); //new org.apache.poi.xssf.usermodel.CellStyle();
    titlCS.setFont(titlFont); 
    titlCS.setBorderBottom(0); 
    titlCS.setBorderTop(0); 
    titlCS.setBorderLeft(0); 
    titlCS.setBorderRight(0);
    titlCS.setFillPattern(Constants.NO_FILL); 
    titlCS.setDataFormat(Constants.XL_CELL_DATAFORMAT_GENERAL); 
    titlCS.setHidden(false); 
    titlCS.setIndention(0); 
    titlCS.setLocked(false); 
    titlCS.setRotation(0); 
    titlCS.setWrapText(false); 
        
    var resultSheet = g_xlWorkbook.createSheet("Deckblatt"); 
    g_xlWorkbook.setSheetOrder("Deckblatt", 0); 
    
    var r = null; 
    var c = null; 
    r = resultSheet.createRow(2); 
    c = r.createCell(1); 
    c.setCellStyle(titlCS); 
    c = r.createCell(2); 
    c.setCellStyle(titlCS); 
    
    for (var row=4; row <(5 + (totals.length*8)); row++)
    {    
        r = resultSheet.createRow(row); 
        c = r.createCell(1); 
        c.setCellStyle(g_cs); 
        c = r.createCell(2); 
        c.setCellStyle(g_cs); 
    }
    resultSheet.setColumnWidth(1, 25*256); 
    resultSheet.setColumnWidth(2, 35*256); 
    
    var currentDate = new Date();
    var dateString = Common_GetRightSubstring("00" + __toString(currentDate.getDate()), 2) + "."+ Common_GetRightSubstring("00" + __toString(currentDate.getMonth() + 1), 2)  +"."+__toString(currentDate.getFullYear()); 
    var timeString = Common_GetRightSubstring("00" + __toString(currentDate.getHours()), 2)+":"+Common_GetRightSubstring("00" + __toString(currentDate.getMinutes()),2)+":"+Common_GetRightSubstring("00" + __toString(currentDate.getSeconds()), 2);
    
    resultSheet.getRow(2).getCell(1).setCellValue("QS Report"); 
    // resultSheet.getRow(2).getCell(1).setCellValue("Datum: "); 
    resultSheet.getRow(2).getCell(2).setCellValue(dateString);
    resultSheet.getRow(4).getCell(1).setCellValue("Uhrzeit: "); 
    resultSheet.getRow(4).getCell(2).setCellValue(timeString);
    resultSheet.getRow(5).getCell(1).setCellValue("Benutzerkennung: "); 
    resultSheet.getRow(5).getCell(2).setCellValue(g_currentModel.Database().ActiveUser().Name(g_usedLanguage));
    var startRow = 6;
    totals.forEach(function(curMod, i){
        startRow++; 
        resultSheet.getRow(startRow).getCell(1).setCellValue("Sheet Number:"); 
        resultSheet.getRow(startRow).getCell(2).setCellValue(i+1);
        startRow++;
        resultSheet.getRow(startRow).getCell(1).setCellValue("Modellname: "); 
        resultSheet.getRow(startRow).getCell(2).setCellValue(curMod[0].Name(g_usedLanguage));
        
        startRow++;
        resultSheet.getRow(startRow).getCell(1).setCellValue("ZAD Kapitel Nr.: "); 
        resultSheet.getRow(startRow).getCell(2).setCellValue(curMod[0].Attribute(g_methodFilter.UserDefinedAttributeTypeNum("317ba881-7df8-11da-0c60-cf8e338f9a0b"), g_usedLanguage).getValue()); 
        startRow++;
        resultSheet.getRow(startRow).getCell(1).setCellValue("Anzahl Warnungen: "); 
        resultSheet.getRow(startRow).getCell(2).setCellValue(__toString(curMod[2]));
        startRow++;
        resultSheet.getRow(startRow).getCell(1).setCellValue("Anzahl Fehler: "); 
        resultSheet.getRow(startRow).getCell(2).setCellValue(__toString(curMod[1]));
        startRow++;
        resultSheet.getRow(startRow).getCell(1).setCellValue("Anzahl Kritische Fehler: "); 
        resultSheet.getRow(startRow).getCell(2).setCellValue(__toString(curMod[3]));
        startRow++;
    })
}

//logging class
function LogMessage(checkType, ruleShortCut, ruleDescription, checkContext, outputText, logLevel)
{
    if (logLevel <= g_currentLogType)
    {
        var r = g_xlSheet.createRow(g_rowNumber); 
        var c= null; 
        var col0CellStyle = null; 
        for (var cou=0; cou<g_countColumns; cou++)
        {
            c = r.createCell(cou); 
            if (cou==1)
                c.setCellStyle(g_loggerClasses[logLevel-1].cellStyle); 
            else
                c.setCellStyle(g_cs); 
        }
        
        try
        {
            r.getCell(0).setCellValue(ruleShortCut); 
            r.getCell(1).setCellValue(g_loggerClasses[logLevel-1].logTypeName); 
            r.getCell(2).setCellValue(checkType); 
            r.getCell(3).setCellValue(checkContext);
            r.getCell(4).setCellValue(ruleDescription); 
        
            if (outputText.indexOf("http") < 0)
                r.getCell(5).setCellValue(outputText); 
            else
            {
                var createHelper = g_xlWorkbook.getCreationHelper();
                r.getCell(5).setCellValue(outputText.substring(0, outputText.indexOf("http")-1)); 
                var textArr = outputText.split(" "); 
                for (cou=0; cou<textArr.length; cou++)
                {
                  if(textArr[cou].indexOf("http") > -1)
                  {
                      var link = createHelper.createHyperlink(1);
                      // link.setAddress(String(textArr[cou])); encodeURI, da URL teilweise unerlaubte Zeichen enthält...-> Java Exception
                      // link.setAddress(encodeURI(String(textArr[cou])));
                      link.setAddress(String(textArr[cou]).replace("{", "%7b").replace("}", "%7d")); // geschweifte Klammern müssen escaped werden, führt ansonsten zu Java Exception
                      r.getCell(5).setHyperlink(link);
                  }    
                }
                r.getCell(5).setCellStyle(g_hlinkStyle); 
            }
        }
        catch(err)
        {
            Dialogs.MsgBox(err, Constants.MSGBOX_BTN_OK, "Error writing xls"); 
        }
        g_rowNumber++; 
        
        if (logLevel == 1)
            g_countCritical++; 
        else if (logLevel == 2)
            g_countErrors++; 
        else if (logLevel == 3)
            g_countWarnings++; 
        else if (logLevel == 4)
            g_countNotice++; 
    }
}

/////////////////////////////////////////////////////////////////////////////////////////
///// XML Class to Read the xml configuration file
////////////////////////////////////////////////////////////////////////////////////////
function xmlConfig(xmlFileName)
{
    var _xmlFileName = xmlFileName; 
    var _rootNode; 
    
    this.readConfigXml = function()
    {
        var xmlConfigFile = Context.getFile(_xmlFileName, Constants.LOCATION_SCRIPT); 
        var xmlInputStream = new java.io.ByteArrayInputStream(xmlConfigFile); 
        var saxBuilder = this.createBuilder();
        
        if (saxBuilder != null)
        {
            try
            {               
                var xmlDoc = saxBuilder.build(xmlInputStream); 
                var _rootNode = xmlDoc.getRootElement(); 
            }
            catch(err)
            {
                Dialogs.MsgBox(err, Constants.MSGBOX_BTN_OK, "Error reading xml"); 
            }
        }
        return _rootNode; 
    }
    
    this.getRuleKind = function(ruleKind)
    {  return _rootNode.getChild(ruleKind);    }
    
    this.getRulesOfRuleKind = function(ruleKind)
    {
        var rootNode = this.readConfigXml(); 
        var ruleKindNodes = rootNode.getChildren("RuleKind"); 
        for (var cou=0; cou<ruleKindNodes.size(); cou++)
        {
            if (ruleKindNodes.get(cou).getAttributeValue("type") == ruleKind)
                return ruleKindNodes.get(cou).getChildren(); 
        }
        return null;     
    }
    
    this.getRuleName = function(ruleNode)
    {   return ruleNode.getAttributeValue("name");    }
    this.getRuleShortcut = function(ruleNode)
    {   return ruleNode.getAttributeValue("shortcut");  }
    
    this.getRuleDescription = function(ruleNode)
    {   
        if( ruleNode.getChildren("RuleDescription").size() > 1)
        {
            for (var cou=0; cou<ruleNode.getChildren("RuleDescription").size(); cou++)
            {
                if (ruleNode.getChildren().get(cou).getAttributeValue("lang") == g_reportLcid)
                    return ruleNode.getChildren().get(cou).getText(); 
            }
        }
        else
            return ruleNode.getChildText("RuleDescription"); 
    }
    this.getAttributeTypeNum = function(ruleNode, attrName)
    {
        if (attrName == null)
            attrName = "attr"; 
        else
            attrName = "compareToAttr"; 
        if (ruleNode.getAttribute(attrName) != null)
        {
            var attrValue = ruleNode.getAttributeValue(attrName); 
            var attrTypeNum = 0; 
            if (attrValue != null)
            {
                if (vbIsNumeric(attrValue)> 0)
                    attrTypeNum = __toLong(attrValue); 
                else if (vbLen(attrValue) == 36) //GUID
                    attrTypeNum = g_methodFilter.UserDefinedAttributeTypeNum(attrValue); 
            }
        }
        else
            attrTypeNum = 0; 
        return attrTypeNum; 
    }
    this.getConnectedSymbolType = function(attrRuleNode)
    {
        var connectedSymbolTypeNum = 0; 
        var connectedSymbol=""; 
        if (attrRuleNode.getChild("AttrValue") != null)
            connectedSymbol = attrRuleNode.getChild("AttrValue").getAttributeValue("connectedSymbol"); 
        if (vbIsNumeric(connectedSymbol))
            connectedSymbolTypeNum = __toLong(connectedSymbol); 
        else
        {
            if (vbLen(connectedSymbol) == 35)
                connectedSymbolTypeNum = g_methodFilter.UserDefinedSymbolTypeNum(connectedSymbol); 
        }
        return connectedSymbolTypeNum;
    }
    
    this.AttributeCheckIsCaseSensitive = function (attrRuleNode)
    {
        var isCaseSensitive = attrRuleNode.getAttributeValue("caseSensitive"); 
        if (isCaseSensitive != null)
            return (isCaseSensitive.toLowerCase() =="true" ? true : false); 
        else
            return false; 

    }
    
    this.getCheckCategory = function(ruleNode)
    {   return categoryValue = __toLong(ruleNode.getAttributeValue("category"));     }
    
    this.getAttrRuleNode = function(ruleNode)
    {  return ruleNode.getChild("AttrRule");    }
    
    this.getStructureRule = function(ruleNode)
    {  return ruleNode.getChild("StructureRule");    }
    
    this.getModelPeripheryRule = function(ruleNode)
    {  return ruleNode.getChild("PeripheryRule");    }
    
    this.getObjectOccurenceRule = function(ruleNode)
    {   return ruleNode.getChild("OccurenceRule"); }
    
    this.getCxnRule = function(ruleNode)
    {   return ruleNode.getChild("CxnRule"); }
    this.getRuleType = function(ruleNode)
    {  var result = ruleNode.getAttributeValue("type");    return (result != null ? result.trim() : result);    }
    
    this.getShowErrorIfEmpty = function(attrRuleNode)
    { if (attrRuleNode.getAttributeValue("showErrorIfEmpty") == null)
        return true; 
      else
      {
          if (attrRuleNode.getAttributeValue("showErrorIfEmpty") == "true")
              return true; 
          else
              return false;
      }
    }
    
    this.getRuleCorrelation = function(ruleNode)
    {  var result = ruleNode.getAttributeValue("correlation");    return (result != null ? result.trim() : result);    }
    
    this.getAttrMinLength = function(attrRuleNode)
    {   return attrRuleNode.getChildText("MinLength");    }
    
    this.getAttrMaxLength = function(attrRuleNode)
    {   return attrRuleNode.getChildText("MaxLength");    }
    
    this.getAttrValue = function(attrRuleNode)
    {
        var neededValue;
        if (attrRuleNode.getChild("AttrValue") != null)
        {
            if (attrRuleNode.getChild("AttrValue").getChildren().size() > 0)
            {
                 //language dependent attr values
                 var langNodeName = "Lang"+__toString(g_usedLanguage);
                 if (attrRuleNode.getChild("AttrValue").getChildText(langNodeName) != null)
                    neededValue= attrRuleNode.getChild("AttrValue").getChildText(langNodeName).trim(); 
                 else
                     neededValue =""; 
            }
            else
            {    
                if (attrRuleNode.getChild("AttrValue").getAttributeValue("type") == "number" || attrRuleNode.getChild("AttrValue").getAttributeValue("type") == "date")
                    neededValue = __toNumeric(attrRuleNode.getChildText("AttrValue").trim()); 
                else //type is string
                    neededValue = attrRuleNode.getChildText("AttrValue").trim(); 
            }
        }
        else 
            neededValue = null; 
        return neededValue; 
    }
    
    this.getAttrValueStartAt = function(attrRuleNode)
    {
        if (attrRuleNode.getChild("AttrValue") == null)
            return 0;
        var value = attrRuleNode.getChild("AttrValue").getAttributeValue("startAt"); 
        if (value != null)
            return __toNumeric(value); 
        else
            return 0; 
    }
    
    this.getAttrValueEndAt = function(attrRuleNode)
    {
        if (attrRuleNode.getChild("AttrValue") == null)
            return 0;
        var value = attrRuleNode.getChild("AttrValue").getAttributeValue("endAt"); 
        if (value != null)
            return __toNumeric(value); 
        else
            return 0; 
    }
    
    this.getRegExp = function(attrRuleNode)
    {   var result= attrRuleNode.getChildText("RegExp"); return (result != null ? result.trim() : result);    }
    
    this.getDependentAttributeNode = function(attrRuleNode)
    {   return attrRuleNode.getChild("DependentAttribute");    }
    
    this.getDependenceType = function(attrRuleNode)
    {   return attrRuleNode.getAttributeValue("depType"); }
    
    this.getAttrValueType = function(attrRuleNode)
    {   return attrRuleNode.getChild("AttrValue").getAttributeValue("type");    }
    
    
    this.getAttrCompType = function(attrRuleNode)
    {
        if (attrRuleNode.getChild("AttrValue") == null)
            return comp = ""; 
        else
            return attrRuleNode.getChild("AttrValue").getAttributeValue("comp");    
    }
    
    this.getConnectedSymbolType = function(attrRuleNode)
    {
        var connectedSymbol = attrRuleNode.getChild("AttrValue").getAttributeValue("connectedSymbol"); 
        var symbolTypeNum= 0; 
         if (vbIsNumeric(connectedSymbol)> 0)
            symbolTypeNum = __toLong(connectedSymbol); 
         else if (vbLen(connectedObject) == 36) //GUID
            symbolTypeNum = g_methodFilter.UserDefinedSymbolTypeNum(connectedSymbol); 

        return symbolTypeNum; 
    }
    this.getRuleDescription = function(ruleNode)
    {
        var ruleDescriptionNodes = ruleNode.getChildren("RuleDescription"); 
        for (var cou = 0; cou < ruleDescriptionNodes.size(); cou++)
        {
            // if (g_reportLcid == __toLong(ruleDescriptionNodes.get(cou).getAttributeValue("lang"))) nur deutsch
            if (__toLong(ruleDescriptionNodes.get(cou).getAttributeValue("lang")) == 1031)
                return ruleDescriptionNodes.get(cou).getText(); 
        }
        return null; 
    }
    
    this.getErrorMessage = function(ruleNode)
    {
        var errorMessageNodes = ruleNode.getChildren("ErrorMessage"); 
        for (var cou = 0; cou < errorMessageNodes.size(); cou++)
        {
            // if (g_reportLcid == __toLong(errorMessageNodes.get(cou).getAttributeValue("lang"))) nur deutsch
            if (__toLong(errorMessageNodes.get(cou).getAttributeValue("lang")) == 1031)
                return errorMessageNodes.get(cou).getText(); 
        }       
        return null; 
    }
    
    this.getAttrValue_customCheck = function(attrRuleNode)
    {
        return attrRuleNode.getAttributeValue("customCheck")
    }
    
    this.createBuilder = function() 
    {
        try
        {

            var builder = new org.jdom.input.SAXBuilder(true);
            return builder; 
           
        }
        catch(err)
        {
            Dialogs.MsgBox("Reading xml configuration failed: "+ err, Constants.MSGBOX_RESULT_OK, "Error creating sax builder"); 
            return null; 
        }                       
    }
    
}


