//================================================================================//
//==============================] kay.pop.shop GOs [==============================//
//===============================] danilo  montes [===============================//
//================================================================================//

// globals
let ss = SpreadsheetApp.getActiveSpreadsheet();
let ui = SpreadsheetApp.getUi();
const Ranges = Object.freeze({
  // actual Ranges
  OrderInfo          : Symbol('orderInfo'),
  SetInfo            : Symbol('setInfo'),
  ShippingInfo       : Symbol('shippingInfo'),
  JoinerInfo         : Symbol('joinerInfo'),

  // buttons
  AddSetCost         : Symbol('addSetCost'),
  AddShippingCost    : Symbol('addShippingCost'),
  AddOrder           : Symbol('addOrder'),
  AddMember          : Symbol('addMember'),
  RemoveSetCost      : Symbol('removeSetCost'),
  RemoveShippingCost : Symbol('removeShippingCost'),
  RemoveOrder        : Symbol('removeOrder'),
  RemoveMember       : Symbol('removeMember')
});
let sheetInfo = {
  "activeSheet" : ss.getActiveSheet().getSheetId(),
  "ranges" : {
    "orderInfo"    : getNamedRangeActual(makeFullRangeName("orderInfo")),
    "setInfo"      : getNamedRangeActual(makeFullRangeName("setInfo")),
    "shippingInfo" : getNamedRangeActual(makeFullRangeName("shippingInfo")),
    "joinerInfo"   : getNamedRangeActual(makeFullRangeName("joinerInfo"))
  },
  "colors" : {
    "border"     : "#580d8b",
    "background" : "#aa63da",
    "mainHeader" : "#f0a8f0",
    "subheader"  : "#f9c3f9"
  },
  "costs" : {
    "sets"     : {},
    "shipping" : {}
  }
};
initCosts();



//================================================================//
//==========================] triggers [==========================//
//================================================================//

function onOpen(e) {
  // create menus
  ui
    .createMenu('kay.pop.shop')
    .addItem('Create New GO', 'createGO')
    // .addItem('Add/Remove Based On Selection', 'addOrRemove')
    .addItem('DONT CLICK :)', 'addOrRemove')
    .addSeparator()
    // .addItem('Add Member/Item', 'addMember')
    // .addItem('Add Order', 'addOrder')
    .addItem('Add Set Cost', 'addSetCost')
    .addItem('Add Shipping Cost', 'addShippingCost')
    .addSeparator()
    // .addItem('Change Sheet Order', 'sortSheets')
    .addItem('Archive Order', 'archiveOrder')
    .addToUi();

  // update joiner info
}


function onEdit(e) {
  let editedRange = getRangeEnumFromRange(e.range);
  // TODO

  switch (editedRange) {
    case Ranges.OrderInfo:
      ui.alert('orderInfo edited');
      // update joiner info
      break;

    case Ranges.SetInfo:
      ui.alert('setInfo edited');
      // update internal set info
      // update joiner info
      break;

    case Ranges.ShippingInfo:
      ui.alert('shippingInfo edited');
      // update internal shipping info
      // update joiner info
      break;

    case Ranges.JoinerInfo:
      ui.alert('joinerInfo edited');
      // warn to not edit joiner info
      break;

    default: return;
  }
}



//================================================================//
//========================] menu items [==========================//
//================================================================//

/**
 * Creates a sheet for a new group order.
 */
function createGO() {
  let response = ui.prompt(
      'Create New GO',
      'Enter GO Name',
      ui.ButtonSet.OK_CANCEL
    );
  
  switch(response.getSelectedButton()) {
    case ui.Button.OK:
      let goTitle = response.getResponseText();

      // no error check for duplicate name since it's already handled well by Google Script
      ss.insertSheet(goTitle, {template: ss.getSheetByName('template')}) 
        .getRange('C2')
        .setValue(goTitle);

    case ui.Button.CANCEL:
    case ui.Button.CLOSE:
      break;

    default:
      ui.alert('Error creating GO');
  }
}


/**
 * Contextually performs action in active cell.
 */
function addOrRemove() {
  let activeRange = ss.getActiveRange();
  if (activeRange.getNumRows() != 1 || activeRange.getNumColumns() != 1) {
    ui.alert('Please select only the cell you want to activate.');
    return;
  }

  let selectedRange = getRangeEnum();
  switch(selectedRange) {
    case Ranges.AddOrder           : addOrder(); break;
    case Ranges.AddMember          : addMember(); break;
    case Ranges.AddSetCost         : addSetCost(); break;
    case Ranges.AddShippingCost    : addShippingCost(); break;
    case Ranges.RemoveSetCost      : removeSetCost(); break;
    case Ranges.RemoveShippingCost : removeShippingCost(); break;
    case Ranges.RemoveMember       : removeMember(); break;
    case Ranges.RemoveOrder        : removeOrder(); break;
    default: ui.alert('Please select a red or blue text colored cell.');
  }
}

//======================================================//
//=====================] member [=======================//
//======================================================//

/**
 * Adds a member/item to the order list.
 */
function addMember() {
  updateSheetInfoObject();
  // TODO
  // prompt for number of members to add

  // let namedRange = makeFullRangeName("orderInfo");
  // let dataRange = ss.getRangeByName(namedRange);

  // ss.removeNamedRange(namedRange);

  // let newDataRange = dataRange; // increase rows in data range baesd on response

  // ss.setNamedRange(namedRange, newDataRange);
}


/**
 * Removes a member/item from the order list.
 */
function removeMember() {
  updateSheetInfoObject();
  ss.getActiveSheet().deleteRow(ss.getCurrentCell().getRow());
  sheetInfo.ranges.orderInfo = getNamedRangeActual(makeFullRangeName("orderInfo"));
  updateJoinerInfo();
}

//======================================================//
//=====================] order [========================//
//======================================================//

/**
 * Adds an order to the order list.
 */
function addOrder() {
  updateSheetInfoObject();
  // TODO
}


/**
 * Removes an order to the order list.
 */
function removeOrder() {
  updateSheetInfoObject();
  // TODO
}

//======================================================//
//======================] set [=========================//
//======================================================//

/**
 * Adds a set cost.
 */
function addSetCost() {
  updateSheetInfoObject();

  let response = ui.prompt(
      'Add Set Cost',
      'Enter name of cost and price, separated by a comma\n' +
      '(ex: Album Signed, 100)',
      ui.ButtonSet.OK_CANCEL
  );

  if (response.getResponseText() == "") {
    ui.alert("Please enter a valid name.");
    return;
  }

  switch (response.getSelectedButton()) {
    case ui.Button.OK:
      let setInfo = sheetInfo.ranges.setInfo.getRange();
      let lastRow = setInfo.getLastRow();
      let sheet = ss.getActiveSheet();

      let parts = response.getResponseText().split(",");
      let name = parts[0];
      let cost = parts[1].trim();
      if (!isNumeric(cost)) {
        ui.alert("Please enter a valid price.");
        break;
      }

      let correctFormatting = SpreadsheetApp.newTextStyle()
                              .setUnderline(false)
                              .build();

      sheet.insertRowBefore(lastRow)
           .getRange(lastRow, setInfo.getColumn(), 1, 3)
           .setValues([
              [name, cost, "delete set"]
           ])
           .setTextStyle(correctFormatting)
           .setHorizontalAlignment('left');

      sheetInfo.ranges.setInfo = getNamedRangeActual(makeFullRangeName("setInfo"));
    
    case ui.Button.CLOSE:
    case ui.Button.CANCEL:
      break;

    default: ui.alert('Error adding set cost.')
  }
}


/**
 * Removes a set cost.
 */
function removeSetCost() {
  updateSheetInfoObject();
  ss.getActiveSheet().deleteRow(ss.getCurrentCell().getRow());
  sheetInfo.ranges.setInfo = getNamedRangeActual(makeFullRangeName("setInfo"));
  updateJoinerInfo();
}

//======================================================//
//====================] shipping [======================//
//======================================================//

/**
 * Adds a shipping cost.
 */
function addShippingCost() {
  updateSheetInfoObject();

  let response = ui.prompt(
      'Add Shipping Cost',
      'Enter name of cost and price, separated by a comma\n' +
      '(ex: girl boss fee, 1000)',
      ui.ButtonSet.OK_CANCEL
  );

  if (response.getResponseText() == "") {
    ui.alert("Please enter a valid name.");
    return;
  }

  switch (response.getSelectedButton()) {
    case ui.Button.OK:
      let shippingInfo = sheetInfo.ranges.shippingInfo.getRange();
      let lastRow = shippingInfo.getLastRow();
      let sheet = ss.getActiveSheet();

      let parts = response.getResponseText().split(",");
      let name = parts[0];
      let cost = parts[1].trim();
      if (!isNumeric(cost)) {
        ui.alert("Please enter a valid price.");
        break;
      }
      
      sheet.insertRowBefore(lastRow)
           .getRange(lastRow, shippingInfo.getColumn(), 1, 3)
           .setValues([
              [name, cost, "delete shipping"]
           ])
           .setHorizontalAlignment('left');

      sheetInfo.ranges.shippingInfo = getNamedRangeActual(makeFullRangeName("shippingInfo"));
      sheetInfo.ranges.shippingInfo.getRange()
                                   .setBorder(null, true, true, true, null, null, 
                                              sheetInfo.colors.border,
                                              SpreadsheetApp.BorderStyle.SOLID_THICK);
    
    case ui.Button.CLOSE:
    case ui.Button.CANCEL:
      break;

    default: ui.alert('Error adding shipping cost.')
  }
}


/**
 * Removes a shipping cost.
 */
function removeShippingCost() {
  updateSheetInfoObject();
  ss.getActiveSheet().deleteRow(ss.getCurrentCell().getRow());
  sheetInfo.ranges.shippingInfo = getNamedRangeActual(makeFullRangeName("shippingInfo"));
  updateJoinerInfo();
}

//======================================================//
//======================] etc [=========================//
//======================================================//

/**
 * Changes the order of the sheets.
 */
function sortSheets() {
  // TODO
  // https://www.labnol.org/sort-sheets-in-google-spreadsheet-230512
}


/**
 * Archives the currently active order sheet,
 * sending it to the archive spreadsheet.
 */
function archiveOrder() {
  if (ss.getSheets().length == 3) {
    ui.alert('Since this is the last GO in the sheet, remember to hide the "template" and "template info" sheets after you create a new GO (right-click on sheet and click "Hide Sheet").');
  }

  let orderToArchive = ss.getActiveSheet();
  let destination = SpreadsheetApp.openById('1QuQ9txGkRZ2rkyPMtLnVSt2ArSPDc3_vbLtavbO0BtM');
  let name = orderToArchive.getName();

  try {
    orderToArchive.copyTo(destination).setName(name);
  } catch (e) {
    let allArchivedSheets = destination.getSheets();
    let archivedOrder = allArchivedSheets[allArchivedSheets.length-1];
    ui.alert(`There is already an archived order with the name "${name}." A copy was made the name "${archivedOrder.getName()}".`);
  }
  ss.deleteActiveSheet();
}

//================================================================//
//===========================] utils [============================//
//================================================================//

/**
 * Create the objects that hold the set and shipping costs.
 */
function initCosts() {
  updateSetCosts();
  updateShippingCosts();
}


/**
 * Updates the joinerInfo range, run after any modification is made
 * to the orderInfo, shippingInfo, or setInfo ranges.
 */
function updateJoinerInfo() {
  // TODO
}


/**
 * Adds joiner to the By Joiner category.
 * 
 * @param {string} joiner 
 */
function addJoiner(joiner) {
  // TODO
  let sheet = SpreadsheetApp.getActiveSheet();
}


/**
 * Updates the internal set cost dict.
 */
function updateSetCosts() {
  let costs = {};
  let setsNamedRange = sheetInfo.ranges.setInfo;
  let values = setsNamedRange.getRange().getValues();

  for (let i = 1; i < values.length - 1; i++) {
    costs[values[i][0]] = values[i][1];
  }

  sheetInfo.costs.sets = costs;
}


/**
 * Updates the internal shipping cost dict.
 */
function updateShippingCosts() {
  let costs = {};
  let shippingNamedRange = sheetInfo.ranges.shippingInfo;
  let values = shippingNamedRange.getRange().getValues();

  for (let i = 1; i < values.length - 1; i++) {
    costs[values[i][0]] = values[i][1];
  }

  sheetInfo.costs.shipping = costs;
}



//================================================================//
//==========================] helpers [===========================//
//================================================================//

/**
 * Gets the corresponding Ranges enum for a given active cell.
 * 
 * @returns {Enum}
 */
function getRangeEnum() {
  switch (ss.getCurrentCell().getValue()) {
    case "add order"         : return Ranges.AddOrder;
    case "add m/i"           : return Ranges.AddMember;
    case "Add Set Cost"      : return Ranges.AddSetCost;
    case "Add Shipping Cost" : return Ranges.AddShippingCost;
    case "delete set"        : return Ranges.RemoveSetCost;
    case "delete shipping"   : return Ranges.RemoveShippingCost;
    case "remove"            : return Ranges.RemoveMember;
    case "remove order"      : return Ranges.RemoveOrder;
    default                  : return null;
  }
}


/**
 * Gets the corresponding Ranges enum for a given edited cell.
 * 
 * @param {Range} range
 * @returns {Enum}
 */
function getRangeEnumFromRange(range) {
  let col = range.getColumn();
  let row = range.getRow();
  // Logger.log(col);
  // Logger.log(row);

  let orderInfoRange    = sheetInfo.ranges.orderInfo.getRange();
  let setInfoRange      = sheetInfo.ranges.setInfo.getRange();
  let shippingInfoRange = sheetInfo.ranges.shippingInfo.getRange();
  let joinerInfoRange   = sheetInfo.ranges.joinerInfo.getRange();
  
  if (
      (col >= orderInfoRange.getColumn() &&
       col <= orderInfoRange.getLastColumn()) &&
      (row >= orderInfoRange.getRow() &&
       row <= orderInfoRange.getLastRow())
    ) return Ranges.OrderInfo;

  if (
      (col >= setInfoRange.getColumn() &&
       col <= setInfoRange.getLastColumn()) &&
      (row >= setInfoRange.getRow() &&
       row <= setInfoRange.getLastRow())
    ) return Ranges.SetInfo;
    

  if (
      (col >= shippingInfoRange.getColumn() &&
       col <= shippingInfoRange.getLastColumn()) &&
      (row >= shippingInfoRange.getRow() &&
       row <= shippingInfoRange.getLastRow())
    ) return Ranges.ShippingInfo;

  if (
      (col >= joinerInfoRange.getColumn() &&
       col <= joinerInfoRange.getLastColumn()) &&
      (row >= joinerInfoRange.getRow() &&
       row <= joinerInfoRange.getLastRow())
    ) return Ranges.JoinerInfo;
  
  return null;
}


/**
 * If needed, updates the named ranges in the `sheetInfo` object to be the active sheet's ranges.
 */
function updateSheetInfoObject() {
  let activeSheet = ss.getActiveSheet().getSheetId();
  if (activeSheet == sheetInfo.activeSheet) {
    // Logger.log(`in sheet: ${getSheetById(activeSheet)}`)
    return;
  }

  sheetInfo.activeSheet = activeSheet;
  for (let range in sheetInfo.ranges) {
    if (sheetInfo.ranges.hasOwnProperty(range)) {
      sheetInfo.ranges[range] = ss.getRangeByName(makeFullRangeName(range));
    }
  }

  initCosts();
}


/**
 * Removes the named ranges created by duplicating the template sheet.
 */
function deleteTestRanges() {
  ss.getNamedRanges().forEach((range) => {
      if (range.getName().includes("!")) range.remove();
  });
  return;
}


/**
 * Creates a string of the full NamedRange name for a given type of range.
 * 
 * @param {string} rangeType - type of range to create string for
 * @returns {string}
 */
function makeFullRangeName(rangeType) {
  let active = ss.getActiveSheet();
  if (active.getName() == 'template') {
    return rangeType;
  }
  // sheet name is surrounded by single quotes since 
  // NamedRange.getName() has quotes around the sheetname
  return `'${active.getName()}'!${rangeType}`; 
}


/////////// https://stackoverflow.com/a/26682689
/**
 * Gets sheet by a given id (why isn't this a native function).
 * 
 * @param {int} id - id of the sheet to locate
 * @returns {Sheet}
 */
function getSheetById(id) {
  return ss.getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}
///////////


/**
 * Gets the actual NamedRange instead of just the Range like
 * ss.getRangeByName does.
 * 
 * @param {string} fullRangeName - the entire name of the range, including the sheet name and !
 * @returns {NamedRange}
 */
function getNamedRangeActual(fullRangeName) {
  // Logger.log(fullRangeName);
  return ss.getNamedRanges().filter(
    function(namedRange) {
        // Logger.log(namedRange.getName());
        return namedRange.getName() == fullRangeName;
      }
  )[0];
}


/**
 * Determines if a given input is a number.
 * 
 * @param {string} str - the input string to check
 * @returns {boolean}
 */
function isNumeric(str) {
  if (typeof str !== "string") return false; 
  return !isNaN(str) && !isNaN(parseFloat(str));
}