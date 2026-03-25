type CellValue = string | number | boolean;

function main(workbook: ExcelScript.Workbook) {
  // declared sheet names
  const poSheetVar = "PO Rec";
  const journalSheetVar = "Journal";

  const poSheet: ExcelScript.Worksheet = workbook.getWorksheet(poSheetVar);
  const journalSheet: ExcelScript.Worksheet = workbook.getWorksheet(journalSheetVar);
  if (!poSheet || !journalSheet) return;

  // getting range and values from both sheets
  const poUsedRange: ExcelScript.Range = poSheet.getUsedRange();
  if (!poUsedRange) return;

  let poValues = poUsedRange.getValues() as CellValue[][];
  if (poValues.length === 0) return;

  // type Record = { PO Number: 1 }
  const headerMap: Record<string, number> = {}

  // getting index of all columns required
  for (let col = 0; col < poValues[0].length; col++) {
    const header = poValues[0][col];
    if (typeof header === "string") {
      headerMap[header.trim().toLowerCase()] = col;
    }
  }

  // takes the required columns index
  const poNumberCol: number = headerMap["po number"] ?? -1;
  const quarterCol: number = headerMap["quarter"] ?? -1;
  const allocationCol: number = headerMap["allocation"] ?? -1;
  const finalDeltaCol: number = headerMap["final delta"] ?? -1;
  const firstDeliveryMonthCol: number = headerMap["fy first delivery month"] ?? -1;
  const lastDeliveryMonthCol: number = headerMap["fy last delivery month"] ?? -1;
  const remainingCol: number = headerMap["remaining"] ?? -1;

  if ([poNumberCol, quarterCol, allocationCol, finalDeltaCol, firstDeliveryMonthCol, lastDeliveryMonthCol, remainingCol].includes(-1)) return;

  /**
   * Finds the row index of VOIS cell in column PO Number
   * Sets First and Last delivery month values to -ve increasing order
  */
  let voisRowIndex = -1
  // row = 0 = header
  for (let row = 1; row < poValues.length; row++) {
    const poCell = poValues[row][poNumberCol] // returns [row][B]
    if (typeof poCell === "string" && poCell.toLowerCase() === "vois") {
      voisRowIndex = row;
      break;
    }
  }

  if (voisRowIndex === -1) return;

  // write to FY FDM, LDM
  for (let row = 1; row <= voisRowIndex; row++) {
    const diff = row - voisRowIndex;
    poValues[row][firstDeliveryMonthCol] = diff;
    poValues[row][lastDeliveryMonthCol] = diff;
  }

  for (let row = voisRowIndex + 1; row < poValues.length; row++) {
    // get cell as string
    const quarterCell = poValues[row][quarterCol] as string;
    // get month numbers
    let monthNumber: { first: number; second: number } = extractMonthNumber(quarterCell);
    // update values

    if (monthNumber.first === -1 || monthNumber.second === -1) continue;
    poValues[row][firstDeliveryMonthCol] = monthNumber.first;
    poValues[row][lastDeliveryMonthCol] = monthNumber.second;
  }
  // update all the data && not updating rn
  //   poUsedRange.setValues(poValues)
  /**
   * Part one is done.
   * Second part starts here.
  */

  // getting range and values journal sheet
  const journalUsedRange: ExcelScript.Range = journalSheet.getUsedRange();
  if (!journalUsedRange) return;

  let journalValues = journalUsedRange.getValues() as CellValue[][];
  if (journalValues.length < 2) {
    poUsedRange.setValues(poValues);
    return;
  }

  // type Record = { PO Number: 1 }
  const headerMapJournal: Record<string, number> = {}
  // getting index of all columns required
  // the header row index is == 1 
  for (let col = 0; col < journalValues[1].length; col++) {
    const header = journalValues[1][col];
    if (typeof header === "string") {
      headerMapJournal[header.trim().toLowerCase()] = col;
    }
  }

  // takes the required columns index
  const po3PTransfersCol: number = headerMapJournal["po for 3rd party transfers"] ?? -1;
  const amtToDebitCol: number = headerMapJournal["amount to debit"] ?? -1;
  const amtToCreditCol: number = headerMapJournal["amount to credit"] ?? -1;
  const wbsNum2Col: number = headerMapJournal["wbs number2"] ?? -1;

  if ([po3PTransfersCol, amtToDebitCol, amtToCreditCol, wbsNum2Col].includes(-1)) {
    poUsedRange.setValues(poValues);
    return;
  }

  /**
   * build credit map: wbsnumber2 = amount to credit
   * take the value to subtract first: amtToDebit
   * then go to PO Rec sheet and find the column Final Delta
   * if value < 0: skip
   * if value > 0 && value > amtToDebit => 
   *    STEP1: update allocation+=amtToDebit, 
   *    STEP2: remaining -= allocation, 
   *    STEP3: po3rdparttranser = povalue
   *    STEP4: update credit map
   * 
   * if values > 0 && value < amtToDebit => 
   *    STEP1: update allocation += value, 
   *    STEP2: remaining = 0, 
   *    STEP3: po3rdpartytransfer = poValue
   *    STEP4: add new row 
   *    STEP5: autoFill from above
   *    STEP6: update amtToDebit 
   *    STEP7: using amtToCredit - previous all amount to debit (so basically create an array storing all amt to debit until amtToDebit <= value situation) 
   */

  // creating mapper for remembering total credit to complete 
  let creditDebitMap: Record<string, number> = {};
  for (let row = 2; row < journalValues.length; row++) {
    const wbsKey = journalValues[row][wbsNum2Col] as string;
    const amtToCredit = toNumber(journalValues[row][amtToCreditCol]);
    if (wbsKey !== "" && amtToCredit > 0) {
      creditDebitMap[wbsKey] = amtToCredit;
    }
  }

  var deltaStartIndex = 1;

  // header row index == 1; starting loop from row == 2
  for (let row = 2; row < journalValues.length; row++) {
    /****************************************************
     * outer loop goes through the journel sheet row by row
     * updating creadit map here
     * taking amtToDebit
     *****************************************************/
    if (journalValues[row][po3PTransfersCol] !== "") continue;

    const amtToDebit = toNumber(journalValues[row][amtToDebitCol]);
    if (amtToDebit <= 0) continue;

    // const amtToCredit = journalValues[row][amtToCreditCol] as number;

    const wbsKey = journalValues[row][wbsNum2Col] as string;

    if (!(wbsKey in creditDebitMap)) {
      creditDebitMap[wbsKey] = amtToDebit;
    }

    let remainingDebitForRow: number = amtToDebit;

    // creditDebitMap[wbsNum2] = journalValues[row][amtToDebitCol]
    for (let deltaRow = deltaStartIndex; deltaRow < poValues.length && remainingDebitForRow > 0; deltaRow++) {
      /****************************************************
       * inner loop goes through the PO Rec starting with delatStartIndex and stays in that index range
       * following 3 conditions
       * taking amtToDebit
       *****************************************************/
      const available = toNumber(poValues[deltaRow][finalDeltaCol] as CellValue);

      if (available <= 0) {
        // update global index
        deltaStartIndex = deltaRow + 1;
        // skip the row where available balance is low
        continue;

      }
      if (available >= remainingDebitForRow) {
        // STEP1: get allocation and update it
        const currentAllocation = toNumber(poValues[deltaRow][allocationCol] as CellValue);
        poValues[deltaRow][allocationCol] = currentAllocation + remainingDebitForRow;

        // STEP2: get remaning and final delta and update them
        // const finalDelta = Number(poValues[deltaRow][finalDeltaCol]) || 0;
        // const remaining = Number(poValues[deltaRow][remainingCol]) || 0;
        poValues[deltaRow][finalDeltaCol] = available - remainingDebitForRow;
        // poValues[deltaRow][finalDeltaCol] = remaining - amtToDebit;
        poValues[deltaRow][remainingCol] = poValues[deltaRow][finalDeltaCol]

        // STEP3: update PO for 3rd Party transfers
        const poNumberCell = poValues[deltaRow][poNumberCol];
        journalValues[row][po3PTransfersCol] = poNumberCell;

        // STEP4: update credit map

        creditDebitMap[wbsKey] = Math.max(creditDebitMap[wbsKey] - remainingDebitForRow, 0);


        remainingDebitForRow = 0;

        // updating global index and 
        deltaStartIndex = deltaRow;

      } else {
        // STEP1: getting allocation and updating allocation += available
        const currentAllocation = toNumber(poValues[deltaRow][allocationCol] as CellValue);
        poValues[deltaRow][allocationCol] = currentAllocation + available;

        // STEP2: putting the value of remaining as zero (and final delta value as well.)
        poValues[deltaRow][finalDeltaCol] = 0;
        // poValues[deltaRow][finalDeltaCol] = remaining - amtToDebit;
        poValues[deltaRow][remainingCol] = 0;

        // STEP2.5: updating current amtToDebit
        journalValues[row][amtToDebitCol] = available

        // STEP3: po for 3rdparty transfer = poValue
        // STEP3: update PO for 3rd Party transfers
        const poNumberCell = poValues[deltaRow][poNumberCol];
        journalValues[row][po3PTransfersCol] = poNumberCell;

        // STEP6&7: update amtToDebit && credit map
        creditDebitMap[wbsKey] = creditDebitMap[wbsKey] - available;
        const remainingForWbs = creditDebitMap[wbsKey];

        if (remainingForWbs > 0) {
          const newRow: CellValue[] = [...journalValues[row]];
          newRow[po3PTransfersCol] = "";
          newRow[amtToDebitCol] = remainingForWbs;
          journalValues.splice(row + 1, 0, newRow);
        }

        // updating global index
        remainingDebitForRow = 0;
        deltaStartIndex = deltaRow + 1;
      }
    }
  }

  // final update
  poUsedRange.setValues(poValues);

  const journalStartRow = journalUsedRange.getRowIndex();
  const journalStartCol = journalUsedRange.getColumnIndex();
  const journalColCount = journalUsedRange.getColumnCount();
  const journalRowCount = journalValues.length;

  const journalTargetRange: ExcelScript.Range = journalSheet.getRangeByIndexes(
    journalStartRow,
    journalStartCol,
    journalRowCount,
    journalColCount
  );

  journalTargetRange.setValues(journalValues);
}

/**
 * Takes a string of date and return the month range
 * @params "jan'24 - feb'24"
 * @returns "{ first= 10, second= 11}"
 */
function extractMonthNumber(input: string): { first: number; second: number } {
  // trim whitespaces
  const trimmed: string = input.trim();

  // split on "-" and remove any whitespace inbetween
  // returns a list of strings
  const parts: string[] = trimmed.split("-").map(p => p.trim())
  // ["jan'24", "feb'24"]

  const fMonth = getMonthNumber(parts[0].toLowerCase());
  const sMonth = parts.length > 1 ? getMonthNumber(parts[1].toLowerCase()) : fMonth // get second month, incase not present return first month

  return { first: fMonth, second: sMonth };
}

/**
 * Takes a string of date and return the number based on financial year
 * @param "jan'24"
 * @returns 10
 */
function getMonthNumber(input: string): number {

  const monthMap = {
    "apr": 1, "may": 2, "jun": 3, "jul": 4, "aug": 5, "sep": 6, "oct": 7, "nov": 8, "dec": 9,
    "jan": 10, "feb": 11, "mar": 12
  }

  const match = input.toLowerCase().match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)
  // jan'24 == jan

  if (!match) return -1;
  return monthMap[match[0]] ?? -1;
}

/**
 * Cause the Number() won't work properly
 * @param CellValue
 * @returns Number
 */
function toNumber(value: CellValue): number {
  if (typeof value === "number") {
    return value;
  }
  if (typeof value === "string") {
    const n: number = Number(value);
    return isNaN(n) ? 0 : n;
  }
  return 0;
}
