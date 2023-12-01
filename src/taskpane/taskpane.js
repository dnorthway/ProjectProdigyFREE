/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("SavePO").onclick = SavePO;
        document.getElementById("SaveSCCO").onclick = SaveSCCO;
        document.getElementById("SaveRFP").onclick = SaveRFP;
        document.getElementById("SaveRFI").onclick = SaveRFI;
        document.getElementById("SaveTransmittal").onclick = SaveTransmittal;
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
    }
});
async function SavePO() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("POData");
        const range = sheet.getRange("A2:AK2");

        range.insert(Excel.InsertShiftDirection.down);
        sheet.getRange("D2").copyFrom("PO!D4", Excel.RangeCopyType.values);
        sheet.getRange("E2").copyFrom("PO!B6", Excel.RangeCopyType.values);
        sheet.getRange("F2").copyFrom("PO!C18", Excel.RangeCopyType.values);
        sheet.getRange("G2").copyFrom("PO!J18", Excel.RangeCopyType.values);
        sheet.getRange("H2").copyFrom("PO!A25", Excel.RangeCopyType.values);
        sheet.getRange("I2").copyFrom("PO!A26", Excel.RangeCopyType.values);
        sheet.getRange("J2").copyFrom("PO!A27", Excel.RangeCopyType.values);
        sheet.getRange("K2").copyFrom("PO!A28", Excel.RangeCopyType.values);
        sheet.getRange("L2").copyFrom("PO!A29", Excel.RangeCopyType.values);
        sheet.getRange("M2").copyFrom("PO!A30", Excel.RangeCopyType.values);
        sheet.getRange("N2").copyFrom("PO!A31", Excel.RangeCopyType.values);
        sheet.getRange("O2").copyFrom("PO!H25", Excel.RangeCopyType.values);
        sheet.getRange("P2").copyFrom("PO!H26", Excel.RangeCopyType.values);
        sheet.getRange("Q2").copyFrom("PO!H27", Excel.RangeCopyType.values);
        sheet.getRange("R2").copyFrom("PO!H28", Excel.RangeCopyType.values);
        sheet.getRange("S2").copyFrom("PO!H29", Excel.RangeCopyType.values);
        sheet.getRange("T2").copyFrom("PO!H30", Excel.RangeCopyType.values);
        sheet.getRange("U2").copyFrom("PO!H31", Excel.RangeCopyType.values);
        sheet.getRange("V2").copyFrom("PO!I25", Excel.RangeCopyType.values);
        sheet.getRange("W2").copyFrom("PO!I26", Excel.RangeCopyType.values);
        sheet.getRange("X2").copyFrom("PO!I27", Excel.RangeCopyType.values);
        sheet.getRange("Y2").copyFrom("PO!I28", Excel.RangeCopyType.values);
        sheet.getRange("Z2").copyFrom("PO!I29", Excel.RangeCopyType.values);
        sheet.getRange("AA2").copyFrom("PO!I30", Excel.RangeCopyType.values);
        sheet.getRange("AB2").copyFrom("PO!I31", Excel.RangeCopyType.values);
        sheet.getRange("AC2").copyFrom("PO!J34", Excel.RangeCopyType.values);
        sheet.getRange("AD2").copyFrom("PO!H47", Excel.RangeCopyType.values);
        sheet.getRange("AE2").copyFrom("PO!H47", Excel.RangeCopyType.values);
        sheet.getRange("AF2").copyFrom("PO!I4", Excel.RangeCopyType.values);
        sheet.getRange("AG2").copyFrom("PO!B36", Excel.RangeCopyType.values);
        sheet.getRange("AH2").copyFrom("PO!G6", Excel.RangeCopyType.values);
        sheet.getRange("AI2").copyFrom("PO!G14", Excel.RangeCopyType.values);
        sheet.getRange("AJ2").copyFrom("PO!J37", Excel.RangeCopyType.values);
        sheet.getRange("AK2").copyFrom("PO!G40", Excel.RangeCopyType.values);
        sheet.getRange("AL2").copyFrom("PO!I40", Excel.RangeCopyType.values);
        sheet.getRange("C1").formulas = [["=POData!D1"]];
        sheet.getRange("C2:C2000").copyFrom("POData!C1", Excel.RangeCopyType.formulas);

    });
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("PO");
        sheet.getRange("D4").values = [[""]];
        sheet.getRange("I4").values = [[""]];
        sheet.getRange("C18").values = [[""]];
        sheet.getRange("A25").values = [[""]];
        sheet.getRange("A26").values = [[""]];
        sheet.getRange("A27").values = [[""]];
        sheet.getRange("A28").values = [[""]];
        sheet.getRange("A29").values = [[""]];
        sheet.getRange("A30").values = [[""]];
        sheet.getRange("A31").values = [[""]];
        sheet.getRange("H25").values = [[""]];
        sheet.getRange("H26").values = [[""]];
        sheet.getRange("H27").values = [[""]];
        sheet.getRange("H28").values = [[""]];
        sheet.getRange("H29").values = [[""]];
        sheet.getRange("H30").values = [[""]];
        sheet.getRange("H31").values = [[""]];
        sheet.getRange("I25").values = [[""]];
        sheet.getRange("I26").values = [[""]];
        sheet.getRange("I27").values = [[""]];
        sheet.getRange("I28").values = [[""]];
        sheet.getRange("I29").values = [[""]];
        sheet.getRange("I30").values = [[""]];
        sheet.getRange("I31").values = [[""]];
        sheet.getRange("H47").values = [[""]];
        sheet.getRange("B36").values = [[""]];
        sheet.getRange("H25").values = [[""]];
        sheet.getRange("I25").values = [[""]];
        sheet.getRange("G40").values = [[""]];
        sheet.getRange("I40").values = [[""]];
        sheet.getRange("B6").values = [[""]];
    });
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("POLog");
        let table16 = sheet.tables.getItem("Table16");

        let sortFields = [
            {
                key: 0,
                ascending: false
            }
        ];
        table16.sort.apply(sortFields);

        sheet.getRange("A10").formulas = [["=POData!D2"]];
        sheet.getRange("A11:A2000").copyFrom("POLog!A10", Excel.RangeCopyType.formulas);
        sheet.getRange("B10").formulas = [["=POData!E2"]];
        sheet.getRange("B11:B2000").copyFrom("POLog!B10", Excel.RangeCopyType.formulas);
        sheet.getRange("C10").formulas = [["=POData!F2"]];
        sheet.getRange("C11:C2000").copyFrom("POLog!C10", Excel.RangeCopyType.formulas);
        sheet.getRange("D10").formulas = [["=POData!AC2"]];
        sheet.getRange("D11:D2000").copyFrom("POLog!D10", Excel.RangeCopyType.formulas);
        sheet.getRange("E10").formulas = [["=POData!AD2"]];
        sheet.getRange("E11:E2000").copyFrom("POLog!E10", Excel.RangeCopyType.formulas);
        sheet.getRange("F10").formulas = [["=POData!AE2"]];
        sheet.getRange("F11:F2000").copyFrom("POLog!F10", Excel.RangeCopyType.formulas);
        sheet.getRange("G10").formulas = [["=POData!AL2"]];
        sheet.getRange("G11:G2000").copyFrom("POLog!G10", Excel.RangeCopyType.formulas);

        sheet.tables.sort

    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("POView");
        sheet.activate();
        sheet.getRange("D4").values = [[""]];

    });
};


async function SaveSCCO() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("SCCOData");
        const range = sheet.getRange("A2:AS2");

        range.insert(Excel.InsertShiftDirection.down);
        sheet.getRange("D2").copyFrom("SCCO!C8", Excel.RangeCopyType.values);
        sheet.getRange("E2").copyFrom("SCCO!H5", Excel.RangeCopyType.values);
        sheet.getRange("F2").copyFrom("SCCO!F7", Excel.RangeCopyType.values);
        sheet.getRange("G2").copyFrom("SCCO!C7", Excel.RangeCopyType.values);
        sheet.getRange("H2").copyFrom("SCCO!C9", Excel.RangeCopyType.values);
        sheet.getRange("I2").copyFrom("SCCO!B17", Excel.RangeCopyType.values);
        sheet.getRange("J2").copyFrom("SCCO!A29", Excel.RangeCopyType.values);
        sheet.getRange("K2").copyFrom("SCCO!A30", Excel.RangeCopyType.values);
        sheet.getRange("L2").copyFrom("SCCO!A31", Excel.RangeCopyType.values);
        sheet.getRange("M2").copyFrom("SCCO!A32", Excel.RangeCopyType.values);
        sheet.getRange("N2").copyFrom("SCCO!A33", Excel.RangeCopyType.values);
        sheet.getRange("O2").copyFrom("SCCO!A34", Excel.RangeCopyType.values);
        sheet.getRange("P2").copyFrom("SCCO!A35", Excel.RangeCopyType.values);
        sheet.getRange("Q2").copyFrom("SCCO!A36", Excel.RangeCopyType.values);
        sheet.getRange("R2").copyFrom("SCCO!A37", Excel.RangeCopyType.values);
        sheet.getRange("S2").copyFrom("SCCO!A38", Excel.RangeCopyType.values);
        sheet.getRange("T2").copyFrom("SCCO!F29", Excel.RangeCopyType.values);
        sheet.getRange("U2").copyFrom("SCCO!F30", Excel.RangeCopyType.values);
        sheet.getRange("V2").copyFrom("SCCO!F31", Excel.RangeCopyType.values);
        sheet.getRange("W2").copyFrom("SCCO!F32", Excel.RangeCopyType.values);
        sheet.getRange("X2").copyFrom("SCCO!F33", Excel.RangeCopyType.values);
        sheet.getRange("Y2").copyFrom("SCCO!F34", Excel.RangeCopyType.values);
        sheet.getRange("Z2").copyFrom("SCCO!F35", Excel.RangeCopyType.values);
        sheet.getRange("AA2").copyFrom("SCCO!F36", Excel.RangeCopyType.values);
        sheet.getRange("AB2").copyFrom("SCCO!F37", Excel.RangeCopyType.values);
        sheet.getRange("AC2").copyFrom("SCCO!F38", Excel.RangeCopyType.values);
        sheet.getRange("AD2").copyFrom("SCCO!G29", Excel.RangeCopyType.values);
        sheet.getRange("AE2").copyFrom("SCCO!G30", Excel.RangeCopyType.values);
        sheet.getRange("AF2").copyFrom("SCCO!G31", Excel.RangeCopyType.values);
        sheet.getRange("AG2").copyFrom("SCCO!G32", Excel.RangeCopyType.values);
        sheet.getRange("AH2").copyFrom("SCCO!G33", Excel.RangeCopyType.values);
        sheet.getRange("AI2").copyFrom("SCCO!G34", Excel.RangeCopyType.values);
        sheet.getRange("AJ2").copyFrom("SCCO!G35", Excel.RangeCopyType.values);
        sheet.getRange("AK2").copyFrom("SCCO!G36", Excel.RangeCopyType.values);
        sheet.getRange("AL2").copyFrom("SCCO!G37", Excel.RangeCopyType.values);
        sheet.getRange("AM2").copyFrom("SCCO!G38", Excel.RangeCopyType.values);
        sheet.getRange("AN2").copyFrom("SCCO!F39", Excel.RangeCopyType.values);
        sheet.getRange("AO2").copyFrom("SCCO!G39", Excel.RangeCopyType.values);
        sheet.getRange("AP2").copyFrom("SCCO!F17", Excel.RangeCopyType.values);
        sheet.getRange("AQ2").copyFrom("SCCO!G17", Excel.RangeCopyType.values);
        sheet.getRange("AR2").copyFrom("SCCO!H47", Excel.RangeCopyType.values);
        sheet.getRange("AS2").copyFrom("SCCO!G40", Excel.RangeCopyType.values);
        sheet.getRange("C1").formulas = [["=SCCOData!D1"]];
        sheet.getRange("C2:C2000").copyFrom("SCCOData!C1", Excel.RangeCopyType.formulas);

    });
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("SCCO");
        sheet.getRange("C7").values = [[""]];
        sheet.getRange("C8").values = [[""]];
        sheet.getRange("H5").values = [[""]];
        sheet.getRange("C9").values = [[""]];
        sheet.getRange("B17").values = [[""]];
        sheet.getRange("A29").values = [[""]];
        sheet.getRange("A30").values = [[""]];
        sheet.getRange("A31").values = [[""]];
        sheet.getRange("A32").values = [[""]];
        sheet.getRange("A33").values = [[""]];
        sheet.getRange("A34").values = [[""]];
        sheet.getRange("A35").values = [[""]];
        sheet.getRange("A36").values = [[""]];
        sheet.getRange("A37").values = [[""]];
        sheet.getRange("A38").values = [[""]];
        sheet.getRange("F29").values = [[""]];
        sheet.getRange("F30").values = [[""]];
        sheet.getRange("F31").values = [[""]];
        sheet.getRange("F32").values = [[""]];
        sheet.getRange("F33").values = [[""]];
        sheet.getRange("F34").values = [[""]];
        sheet.getRange("F35").values = [[""]];
        sheet.getRange("F36").values = [[""]];
        sheet.getRange("F37").values = [[""]];
        sheet.getRange("F38").values = [[""]];
        sheet.getRange("H47").values = [[""]];
        sheet.getRange("F7").values = [[""]];

    });
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("SCCOLog");
        let table17 = sheet.tables.getItem("Table17");

        let sortFields = [
            {
                key: 0,
                ascending: false
            }
        ];
        table17.sort.apply(sortFields);
        sheet.getRange("A10").formulas = [["=SCCOData!D2"]];
        sheet.getRange("A11:A2000").copyFrom("SCCOLog!A10", Excel.RangeCopyType.formulas);
        sheet.getRange("B10").formulas = [["=SCCOData!E2"]];
        sheet.getRange("B11:B2000").copyFrom("SCCOLog!B10", Excel.RangeCopyType.formulas);
        sheet.getRange("C10").formulas = [["=SCCOData!F2"]];
        sheet.getRange("C11:C2000").copyFrom("SCCOLog!C10", Excel.RangeCopyType.formulas);
        sheet.getRange("D10").formulas = [["=SCCOData!G2"]];
        sheet.getRange("D11:D2000").copyFrom("SCCOLog!D10", Excel.RangeCopyType.formulas);
        sheet.getRange("E10").formulas = [["=SCCOData!I2"]];
        sheet.getRange("E11:E2000").copyFrom("SCCOLog!E10", Excel.RangeCopyType.formulas);
        sheet.getRange("F10").formulas = [["=SCCOData!AN2"]];
        sheet.getRange("F11:F2000").copyFrom("SCCOLog!F10", Excel.RangeCopyType.formulas);
        sheet.getRange("G10").formulas = [["=SCCOData!AO2"]];
        sheet.getRange("G11:G2000").copyFrom("SCCOLog!G10", Excel.RangeCopyType.formulas);
        sheet.getRange("H10").formulas = [["=SCCOData!AP2"]];
        sheet.getRange("H11:H2000").copyFrom("SCCOLog!H10", Excel.RangeCopyType.formulas);
        sheet.getRange("I10").formulas = [["=SCCOData!AQ2"]];
        sheet.getRange("I11:I2000").copyFrom("SCCOLog!I10", Excel.RangeCopyType.formulas);
        sheet.getRange("J10").formulas = [["=SCCOData!AR2"]];
        sheet.getRange("J11:J2000").copyFrom("SCCOLog!J10", Excel.RangeCopyType.formulas);
        sheet.tables.sort

    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("SCCOView");
        sheet.activate();
        sheet.getRange("D8").values = [[""]];

    });
};
async function SaveRFP() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("RFPData");
        const range = sheet.getRange("A2:EF2");

        range.insert(Excel.InsertShiftDirection.down);
        sheet.getRange("D2").copyFrom("RFP!B3", Excel.RangeCopyType.values);
        sheet.getRange("E2").copyFrom("RFP!D3", Excel.RangeCopyType.values);
        sheet.getRange("F2").copyFrom("RFP!B4", Excel.RangeCopyType.values);
        sheet.getRange("G2").copyFrom("RFP!A7", Excel.RangeCopyType.values);
        sheet.getRange("H2").copyFrom("RFP!A8", Excel.RangeCopyType.values);
        sheet.getRange("I2").copyFrom("RFP!A9", Excel.RangeCopyType.values);
        sheet.getRange("J2").copyFrom("RFP!A10", Excel.RangeCopyType.values);
        sheet.getRange("K2").copyFrom("RFP!A11", Excel.RangeCopyType.values);
        sheet.getRange("L2").copyFrom("RFP!A12", Excel.RangeCopyType.values);
        sheet.getRange("M2").copyFrom("RFP!A13", Excel.RangeCopyType.values);
        sheet.getRange("N2").copyFrom("RFP!A14", Excel.RangeCopyType.values);
        sheet.getRange("O2").copyFrom("RFP!E7", Excel.RangeCopyType.values);
        sheet.getRange("P2").copyFrom("RFP!E8", Excel.RangeCopyType.values);
        sheet.getRange("Q2").copyFrom("RFP!E9", Excel.RangeCopyType.values);
        sheet.getRange("R2").copyFrom("RFP!E10", Excel.RangeCopyType.values);
        sheet.getRange("S2").copyFrom("RFP!E11", Excel.RangeCopyType.values);
        sheet.getRange("T2").copyFrom("RFP!E12", Excel.RangeCopyType.values);
        sheet.getRange("U2").copyFrom("RFP!E13", Excel.RangeCopyType.values);
        sheet.getRange("V2").copyFrom("RFP!E14", Excel.RangeCopyType.values);
        sheet.getRange("W2").copyFrom("RFP!K7", Excel.RangeCopyType.values);
        sheet.getRange("X2").copyFrom("RFP!K8", Excel.RangeCopyType.values);
        sheet.getRange("Y2").copyFrom("RFP!K9", Excel.RangeCopyType.values);
        sheet.getRange("Z2").copyFrom("RFP!K10", Excel.RangeCopyType.values);
        sheet.getRange("AA2").copyFrom("RFP!K11", Excel.RangeCopyType.values);
        sheet.getRange("AB2").copyFrom("RFP!K12", Excel.RangeCopyType.values);
        sheet.getRange("AC2").copyFrom("RFP!K13", Excel.RangeCopyType.values);
        sheet.getRange("AD2").copyFrom("RFP!K14", Excel.RangeCopyType.values);
        sheet.getRange("AE2").copyFrom("RFP!A19", Excel.RangeCopyType.values);
        sheet.getRange("AF2").copyFrom("RFP!A20", Excel.RangeCopyType.values);
        sheet.getRange("AG2").copyFrom("RFP!A21", Excel.RangeCopyType.values);
        sheet.getRange("AH2").copyFrom("RFP!A22", Excel.RangeCopyType.values);
        sheet.getRange("AI2").copyFrom("RFP!A23", Excel.RangeCopyType.values);
        sheet.getRange("AJ2").copyFrom("RFP!A24", Excel.RangeCopyType.values);
        sheet.getRange("AK2").copyFrom("RFP!E19", Excel.RangeCopyType.values);
        sheet.getRange("AL2").copyFrom("RFP!E20", Excel.RangeCopyType.values);
        sheet.getRange("AM2").copyFrom("RFP!E21", Excel.RangeCopyType.values);
        sheet.getRange("AN2").copyFrom("RFP!E22", Excel.RangeCopyType.values);
        sheet.getRange("AO2").copyFrom("RFP!E23", Excel.RangeCopyType.values);
        sheet.getRange("AP2").copyFrom("RFP!E24", Excel.RangeCopyType.values);
        sheet.getRange("AQ2").copyFrom("RFP!K19", Excel.RangeCopyType.values);
        sheet.getRange("AR2").copyFrom("RFP!K20", Excel.RangeCopyType.values);
        sheet.getRange("AS2").copyFrom("RFP!K21", Excel.RangeCopyType.values);
        sheet.getRange("AT2").copyFrom("RFP!K22", Excel.RangeCopyType.values);
        sheet.getRange("AU2").copyFrom("RFP!K23", Excel.RangeCopyType.values);
        sheet.getRange("AV2").copyFrom("RFP!K24", Excel.RangeCopyType.values);
        sheet.getRange("AW2").copyFrom("RFP!A29", Excel.RangeCopyType.values);
        sheet.getRange("AX2").copyFrom("RFP!A30", Excel.RangeCopyType.values);
        sheet.getRange("AY2").copyFrom("RFP!A31", Excel.RangeCopyType.values);
        sheet.getRange("AZ2").copyFrom("RFP!A32", Excel.RangeCopyType.values);
        sheet.getRange("BA2").copyFrom("RFP!A33", Excel.RangeCopyType.values);
        sheet.getRange("BB2").copyFrom("RFP!A34", Excel.RangeCopyType.values);
        sheet.getRange("BC2").copyFrom("RFP!D29", Excel.RangeCopyType.values);
        sheet.getRange("BD2").copyFrom("RFP!D30", Excel.RangeCopyType.values);
        sheet.getRange("BE2").copyFrom("RFP!D31", Excel.RangeCopyType.values);
        sheet.getRange("BF2").copyFrom("RFP!D32", Excel.RangeCopyType.values);
        sheet.getRange("BG2").copyFrom("RFP!D33", Excel.RangeCopyType.values);
        sheet.getRange("BH2").copyFrom("RFP!D34", Excel.RangeCopyType.values);
        sheet.getRange("BI2").copyFrom("RFP!E29", Excel.RangeCopyType.values);
        sheet.getRange("BJ2").copyFrom("RFP!E30", Excel.RangeCopyType.values);
        sheet.getRange("BK2").copyFrom("RFP!E31", Excel.RangeCopyType.values);
        sheet.getRange("BL2").copyFrom("RFP!E32", Excel.RangeCopyType.values);
        sheet.getRange("BM2").copyFrom("RFP!E33", Excel.RangeCopyType.values);
        sheet.getRange("BN2").copyFrom("RFP!E34", Excel.RangeCopyType.values);
        sheet.getRange("BO2").copyFrom("RFP!K29", Excel.RangeCopyType.values);
        sheet.getRange("BP2").copyFrom("RFP!K30", Excel.RangeCopyType.values);
        sheet.getRange("BQ2").copyFrom("RFP!K31", Excel.RangeCopyType.values);
        sheet.getRange("BR2").copyFrom("RFP!K32", Excel.RangeCopyType.values);
        sheet.getRange("BS2").copyFrom("RFP!K33", Excel.RangeCopyType.values);
        sheet.getRange("BT2").copyFrom("RFP!K34", Excel.RangeCopyType.values);
        sheet.getRange("BU2").copyFrom("RFP!A38", Excel.RangeCopyType.values);
        sheet.getRange("BV2").copyFrom("RFP!A39", Excel.RangeCopyType.values);
        sheet.getRange("BW2").copyFrom("RFP!A40", Excel.RangeCopyType.values);
        sheet.getRange("BX2").copyFrom("RFP!A41", Excel.RangeCopyType.values);
        sheet.getRange("BY2").copyFrom("RFP!A42", Excel.RangeCopyType.values);
        sheet.getRange("BZ2").copyFrom("RFP!D38", Excel.RangeCopyType.values);
        sheet.getRange("CA2").copyFrom("RFP!D39", Excel.RangeCopyType.values);
        sheet.getRange("CB2").copyFrom("RFP!D40", Excel.RangeCopyType.values);
        sheet.getRange("CC2").copyFrom("RFP!D41", Excel.RangeCopyType.values);
        sheet.getRange("CD2").copyFrom("RFP!D42", Excel.RangeCopyType.values);
        sheet.getRange("CE2").copyFrom("RFP!E38", Excel.RangeCopyType.values);
        sheet.getRange("CF2").copyFrom("RFP!E39", Excel.RangeCopyType.values);
        sheet.getRange("CG2").copyFrom("RFP!E40", Excel.RangeCopyType.values);
        sheet.getRange("CH2").copyFrom("RFP!E41", Excel.RangeCopyType.values);
        sheet.getRange("CI2").copyFrom("RFP!E42", Excel.RangeCopyType.values);
        sheet.getRange("CJ2").copyFrom("RFP!K38", Excel.RangeCopyType.values);
        sheet.getRange("CK2").copyFrom("RFP!K39", Excel.RangeCopyType.values);
        sheet.getRange("CL2").copyFrom("RFP!K40", Excel.RangeCopyType.values);
        sheet.getRange("CM2").copyFrom("RFP!K41", Excel.RangeCopyType.values);
        sheet.getRange("CN2").copyFrom("RFP!K42", Excel.RangeCopyType.values);
        sheet.getRange("CO2").copyFrom("RFP!D46", Excel.RangeCopyType.values);
        sheet.getRange("CP2").copyFrom("RFP!E46", Excel.RangeCopyType.values);
        sheet.getRange("CQ2").copyFrom("RFP!H15", Excel.RangeCopyType.values);
        sheet.getRange("CR2").copyFrom("RFP!H25", Excel.RangeCopyType.values);
        sheet.getRange("CS2").copyFrom("RFP!H35", Excel.RangeCopyType.values);
        sheet.getRange("CT2").copyFrom("RFP!H43", Excel.RangeCopyType.values);
        sheet.getRange("CU2").copyFrom("RFP!H44", Excel.RangeCopyType.values);
        sheet.getRange("CV2").copyFrom("RFP!H48", Excel.RangeCopyType.values);
        sheet.getRange("CW2").copyFrom("RFP!H55", Excel.RangeCopyType.values);
        sheet.getRange("CX2").copyFrom("RFP!F4", Excel.RangeCopyType.values);
        sheet.getRange("CY2").copyFrom("RFP!C53", Excel.RangeCopyType.values);
        sheet.getRange("CZ2").copyFrom("RFP!C51", Excel.RangeCopyType.values);
        sheet.getRange("DA2").copyFrom("RFP!D7", Excel.RangeCopyType.values);
        sheet.getRange("DB2").copyFrom("RFP!D8", Excel.RangeCopyType.values);
        sheet.getRange("DC2").copyFrom("RFP!D9", Excel.RangeCopyType.values);
        sheet.getRange("DD2").copyFrom("RFP!D10", Excel.RangeCopyType.values);
        sheet.getRange("DE2").copyFrom("RFP!D11", Excel.RangeCopyType.values);
        sheet.getRange("DF2").copyFrom("RFP!D12", Excel.RangeCopyType.values);
        sheet.getRange("DG2").copyFrom("RFP!D13", Excel.RangeCopyType.values);
        sheet.getRange("DH2").copyFrom("RFP!D14", Excel.RangeCopyType.values);
        sheet.getRange("DI2").copyFrom("RFP!D19", Excel.RangeCopyType.values);
        sheet.getRange("DJ2").copyFrom("RFP!D20", Excel.RangeCopyType.values);
        sheet.getRange("DK2").copyFrom("RFP!D21", Excel.RangeCopyType.values);
        sheet.getRange("DL2").copyFrom("RFP!D22", Excel.RangeCopyType.values);
        sheet.getRange("DM2").copyFrom("RFP!D23", Excel.RangeCopyType.values);
        sheet.getRange("DN2").copyFrom("RFP!D24", Excel.RangeCopyType.values);
        sheet.getRange("DQ2").formulas = [["=VLOOKUP(D2,RFPStatus!$A$2:$CF$5148,5,FALSE)"]];
        sheet.getRange("DQ3:DQ2000").copyFrom("RFPData!DQ2", Excel.RangeCopyType.formulas);
        sheet.getRange("DR2").formulas = [["=VLOOKUP(D2,RFPStatus!$A$2:$CF$5148,6,FALSE"]];
        sheet.getRange("DR3:DR2000").copyFrom("RFPData!DR2", Excel.RangeCopyType.formulas);
        sheet.getRange("DS2").formulas = [["=VLOOKUP(D2,RFPStatus!$A$2:$CF$5148,7,FALSE)"]];
        sheet.getRange("DS3:DS2000").copyFrom("RFPData!DS2", Excel.RangeCopyType.formulas);
        sheet.getRange("DT2").formulas = [["=VLOOKUP(D2,RFPStatus!$A$2:$CF$5148,8,FALSE)"]];
        sheet.getRange("DT3:DT2000").copyFrom("RFPData!DT2", Excel.RangeCopyType.formulas);
        sheet.getRange("DU2").formulas = [["=VLOOKUP(D2,RFPStatus!$A$2:$CF$5148,9,FALSE)"]];
        sheet.getRange("DU3:DU2000").copyFrom("RFPData!DU2", Excel.RangeCopyType.formulas);
        sheet.getRange("DV2").formulas = [["=VLOOKUP(D2,RFPStatus!$A$2:$CF$5148,10,FALSE)"]];
        sheet.getRange("DV3:DV2000").copyFrom("RFPData!DV2", Excel.RangeCopyType.formulas);
        sheet.getRange("DW2").formulas = [["=CONCAT(BU2:BY2)"]];
        sheet.getRange("DW3:DW2000").copyFrom("RFPData!DW2", Excel.RangeCopyType.formulas);
        sheet.getRange("DX2").copyFrom("RFP!B49", Excel.RangeCopyType.values);
        sheet.getRange("A1").formulas = [["=RFPData!D1"]];
        sheet.getRange("A2:A2000").copyFrom("RFPData!A1", Excel.RangeCopyType.formulas);

    });
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("RFP");
        sheet.getRange("B3").values = [[""]];
        sheet.getRange("D3").values = [[""]];
        sheet.getRange("B4").values = [[""]];
        sheet.getRange("F4").values = [[""]];
        sheet.getRange("A7").values = [[""]];
        sheet.getRange("A8").values = [[""]];
        sheet.getRange("A9").values = [[""]];
        sheet.getRange("A10").values = [[""]];
        sheet.getRange("A11").values = [[""]];
        sheet.getRange("A12").values = [[""]];
        sheet.getRange("A13").values = [[""]];
        sheet.getRange("A14").values = [[""]];
        sheet.getRange("E7").values = [[""]];
        sheet.getRange("E8").values = [[""]];
        sheet.getRange("E9").values = [[""]];
        sheet.getRange("E10").values = [[""]];
        sheet.getRange("E11").values = [[""]];
        sheet.getRange("E12").values = [[""]];
        sheet.getRange("E13").values = [[""]];
        sheet.getRange("E14").values = [[""]];
        sheet.getRange("D7").formulas = [["=VLOOKUP(A7,Craft!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D8").formulas = [["=VLOOKUP(A8,Craft!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D9").formulas = [["=VLOOKUP(A9,Craft!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D10").formulas = [["=VLOOKUP(A10,Craft!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D11").formulas = [["=VLOOKUP(A11,Craft!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D12").formulas = [["=VLOOKUP(A12,Craft!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D13").formulas = [["=VLOOKUP(A13,Craft!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D14").formulas = [["=VLOOKUP(A14,Craft!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("A19").values = [[""]];
        sheet.getRange("A20").values = [[""]];
        sheet.getRange("A21").values = [[""]];
        sheet.getRange("A22").values = [[""]];
        sheet.getRange("A23").values = [[""]];
        sheet.getRange("A24").values = [[""]];
        sheet.getRange("D19").formulas = [["=VLOOKUP(A19,Equipment!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D20").formulas = [["=VLOOKUP(A20,Equipment!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D21").formulas = [["=VLOOKUP(A21,Equipment!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D22").formulas = [["=VLOOKUP(A22,Equipment!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D23").formulas = [["=VLOOKUP(A23,Equipment!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("D24").formulas = [["=VLOOKUP(A24,Equipment!$A$1:$D$1000,2,FALSE)"]];
        sheet.getRange("K7").values = [[""]];
        sheet.getRange("K8").values = [[""]];
        sheet.getRange("K9").values = [[""]];
        sheet.getRange("K10").values = [[""]];
        sheet.getRange("K11").values = [[""]];
        sheet.getRange("K12").values = [[""]];
        sheet.getRange("K13").values = [[""]];
        sheet.getRange("K14").values = [[""]];
        sheet.getRange("E19").values = [[""]];
        sheet.getRange("E20").values = [[""]];
        sheet.getRange("E21").values = [[""]];
        sheet.getRange("E22").values = [[""]];
        sheet.getRange("E23").values = [[""]];
        sheet.getRange("E24").values = [[""]];
        sheet.getRange("K19").values = [[""]];
        sheet.getRange("K20").values = [[""]];
        sheet.getRange("K21").values = [[""]];
        sheet.getRange("K22").values = [[""]];
        sheet.getRange("K23").values = [[""]];
        sheet.getRange("K24").values = [[""]];
        sheet.getRange("A29").values = [[""]];
        sheet.getRange("A30").values = [[""]];
        sheet.getRange("A31").values = [[""]];
        sheet.getRange("A32").values = [[""]];
        sheet.getRange("A33").values = [[""]];
        sheet.getRange("A34").values = [[""]];
        sheet.getRange("D29").values = [[""]];
        sheet.getRange("D30").values = [[""]];
        sheet.getRange("D31").values = [[""]];
        sheet.getRange("D32").values = [[""]];
        sheet.getRange("D33").values = [[""]];
        sheet.getRange("D34").values = [[""]];
        sheet.getRange("E29").values = [[""]];
        sheet.getRange("E30").values = [[""]];
        sheet.getRange("E31").values = [[""]];
        sheet.getRange("E32").values = [[""]];
        sheet.getRange("E33").values = [[""]];
        sheet.getRange("E34").values = [[""]];
        sheet.getRange("K29").values = [[""]];
        sheet.getRange("K30").values = [[""]];
        sheet.getRange("K31").values = [[""]];
        sheet.getRange("K32").values = [[""]];
        sheet.getRange("K33").values = [[""]];
        sheet.getRange("K34").values = [[""]];
        sheet.getRange("A38").values = [[""]];
        sheet.getRange("A39").values = [[""]];
        sheet.getRange("A40").values = [[""]];
        sheet.getRange("A41").values = [[""]];
        sheet.getRange("A42").values = [[""]];
        sheet.getRange("D38").values = [[""]];
        sheet.getRange("D39").values = [[""]];
        sheet.getRange("D40").values = [[""]];
        sheet.getRange("D41").values = [[""]];
        sheet.getRange("D42").values = [[""]];
        sheet.getRange("E38").values = [[""]];
        sheet.getRange("E39").values = [[""]];
        sheet.getRange("E40").values = [[""]];
        sheet.getRange("E41").values = [[""]];
        sheet.getRange("E42").values = [[""]];
        sheet.getRange("K38").values = [[""]];
        sheet.getRange("K39").values = [[""]];
        sheet.getRange("K40").values = [[""]];
        sheet.getRange("K41").values = [[""]];
        sheet.getRange("K42").values = [[""]];
        sheet.getRange("D46").values = [[""]];
        sheet.getRange("E46").values = [[""]];
        sheet.getRange("C51").values = [[""]];
        sheet.getRange("C53").values = [[""]];
        sheet.getRange("B49").values = [[""]];

    });
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("RFPLog");
        let table5 = sheet.tables.getItem("Table5");

        let sortFields = [
            {
                key: 0,
                ascending: false
            }
        ];
        table5.sort.apply(sortFields);

        sheet.getRange("A7").formulas = [["=RFPData!D2"]];
        sheet.getRange("A8:A2000").copyFrom("=RFPLog!A7", Excel.RangeCopyType.formulas);
        sheet.getRange("B7").formulas = [["=RFPData!DU2"]];
        sheet.getRange("B8:B2000").copyFrom("=RFPLog!B7", Excel.RangeCopyType.formulas);
        sheet.getRange("C7").formulas = [["=RFPData!E2"]];
        sheet.getRange("C8:C2000").copyFrom("=RFPLog!C7", Excel.RangeCopyType.formulas);
        sheet.getRange("D7").formulas = [["=RFPData!DT2"]];
        sheet.getRange("D8:D2000").copyFrom("=RFPLog!D7", Excel.RangeCopyType.formulas);
        sheet.getRange("E7").formulas = [["=RFPData!DW2"]];
        sheet.getRange("E8:E2000").copyFrom("=RFPLog!E7", Excel.RangeCopyType.formulas);
        sheet.getRange("F7").formulas = [["=RFPData!CX2"]];
        sheet.getRange("F8:F2000").copyFrom("=RFPLog!F7", Excel.RangeCopyType.formulas);
        sheet.getRange("G7").formulas = [["=RFPData!DS2"]];
        sheet.getRange("G8:G2000").copyFrom("=RFPLog!G7", Excel.RangeCopyType.formulas);
        sheet.getRange("H7").formulas = [["=RFPData!DQ2"]];
        sheet.getRange("H8:H2000").copyFrom("=RFPLog!H7", Excel.RangeCopyType.formulas);
        sheet.getRange("I7").formulas = [["=RFPData!DV2"]];
        sheet.getRange("I8:I2000").copyFrom("=RFPLog!I7", Excel.RangeCopyType.formulas);
        sheet.getRange("J7").formulas = [["=RFPData!CV2"]];
        sheet.getRange("J8:J2000").copyFrom("RFPLog!J7", Excel.RangeCopyType.formulas);
        sheet.getRange("K7").formulas = [["=RFPData!CW2"]];
        sheet.getRange("K8:K2000").copyFrom("=RFPLog!K7", Excel.RangeCopyType.formulas);
        sheet.getRange("L7").formulas = [["=RFPData!DR2"]];
        sheet.getRange("L8:L2000").copyFrom("=RFPlog!L7", Excel.RangeCopyType.formulas);
        sheet.tables.sort

    });
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("RFPStatus");
        const range = sheet.getRange("A2:AK2");

        range.insert(Excel.InsertShiftDirection.down);
        sheet.getRange("RFPStatus!E2").copyFrom("RFPStatus!AA1");
        sheet.getRange("RFPStatus!F2").copyFrom("RFPStatus!AB1");
        sheet.getRange("RFPStatus!A2").formulas = [["=RFPStatus!D2"]];
        sheet.getRange("RFPStatus!A3:A2000").copyFrom("RFPStatus!A2", Excel.RangeCopyType.formulas);
        sheet.getRange("RFPStatus!D2").formulas = [["=RFPData!D2"]];
        sheet.getRange("RFPStatus!D3:D2000").copyFrom("RFPStatus!D2", Excel.RangeCopyType.formulas);
        sheet.getRange("RFPStatus!K2").formulas = [["=RFPData!E2"]];
        sheet.getRange("RFPStatus!K3:K2000").copyFrom("RFPStatus!K2", Excel.RangeCopyType.formulas);

    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("RFPView");
        sheet.activate();
        sheet.getRange("B3").values = [[""]];

    });
};

async function SaveRFI() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("RFIData");
        const range = sheet.getRange("A2:T2");

        range.insert(Excel.InsertShiftDirection.down);
        sheet.getRange("D2").copyFrom("RFI!B9", Excel.RangeCopyType.values);
        sheet.getRange("F2").copyFrom("RFI!B15", Excel.RangeCopyType.values);
        sheet.getRange("G2").copyFrom("RFI!B17", Excel.RangeCopyType.values);
        sheet.getRange("H2").copyFrom("RFI!H17", Excel.RangeCopyType.values);
        sheet.getRange("I2").copyFrom("RFI!C19", Excel.RangeCopyType.values);
        sheet.getRange("J2").copyFrom("RFI!C21", Excel.RangeCopyType.values);
        sheet.getRange("K2").copyFrom("RFI!D23", Excel.RangeCopyType.values);
        sheet.getRange("L2").copyFrom("RFI!B25", Excel.RangeCopyType.values);
        sheet.getRange("M2").copyFrom("RFI!A31", Excel.RangeCopyType.values);
        sheet.getRange("N2").copyFrom("RFI!I6", Excel.RangeCopyType.values);
        sheet.getRange("O2").copyFrom("RFI!B38", Excel.RangeCopyType.values);
        sheet.getRange("P2").formulas = [["=VLOOKUP(D2,RFIResponse!$A$2:$CF$5148,6,FALSE)"]];
        sheet.getRange("P3:P2000").copyFrom("RFIData!P2", Excel.RangeCopyType.formulas);
        sheet.getRange("Q2").formulas = [["=VLOOKUP(D2,RFIResponse!$A$2:$CF$5148,7,FALSE)"]];
        sheet.getRange("Q3:Q2000").copyFrom("RFIData!Q2", Excel.RangeCopyType.formulas);
        sheet.getRange("R2").formulas = [["=VLOOKUP(D2,RFIResponse!$A$2:$CF$5148,5,FALSE)"]];
        sheet.getRange("R3:R2000").copyFrom("RFIData!R2", Excel.RangeCopyType.formulas);
        sheet.getRange("S2").formulas = [["=VLOOKUP(D2,RFIResponse!$A$2:$CF$5148,8,FALSE)"]];
        sheet.getRange("S3:S2000").copyFrom("RFIData!S2", Excel.RangeCopyType.formulas);
        sheet.getRange("A1").formulas = [["=RFIData!D1"]];
        sheet.getRange("A2:A2000").copyFrom("RFIData!A1", Excel.RangeCopyType.formulas);

    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("RFI");
        sheet.getRange("B9").values = [[""]];
        sheet.getRange("I6").values = [[""]];
        sheet.getRange("B15").values = [[""]];
        sheet.getRange("D23").values = [[""]];
        sheet.getRange("B25").values = [[""]];
        sheet.getRange("A31").values = [[""]];
        sheet.getRange("B38").values = [[""]];

    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("RFILog");
        let table10 = sheet.tables.getItem("Table10");

        let sortFields = [
            {
                key: 1,
                ascending: false
            }
        ];
        table10.sort.apply(sortFields);
        sheet.getRange("A7").formulas = [["=RFIData!D2"]];
        sheet.getRange("A8:A2000").copyFrom("RFILog!A7", Excel.RangeCopyType.formulas);
        sheet.getRange("B7").formulas = [["=RFIData!K2"]];
        sheet.getRange("B8:B2000").copyFrom("RFILog!B7", Excel.RangeCopyType.formulas);
        sheet.getRange("C7").formulas = [["=RFIData!F2"]];
        sheet.getRange("C8:C2000").copyFrom("RFILog!C7", Excel.RangeCopyType.formulas);
        sheet.getRange("D7").formulas = [["=RFIData!G2"]];
        sheet.getRange("D8:D2000").copyFrom("RFILog!D7", Excel.RangeCopyType.formulas);
        sheet.getRange("E7").formulas = [["=RFIData!L2"]];
        sheet.getRange("E8:E2000").copyFrom("RFILog!E7", Excel.RangeCopyType.formulas);
        sheet.getRange("F7").formulas = [["=RFIData!P2"]];
        sheet.getRange("F8:F2000").copyFrom("RFILog!F7", Excel.RangeCopyType.formulas);
        sheet.getRange("G7").formulas = [["=RFIData!N2"]];
        sheet.getRange("G8:G2000").copyFrom("RFILog!G7", Excel.RangeCopyType.formulas);
        sheet.tables.sort

    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("RFIResponse");
        const range = sheet.getRange("A2:AK2");

        range.insert(Excel.InsertShiftDirection.down);
        sheet.getRange("RFIResponse!A2").formulas = [["=RFIResponse!D2"]];
        sheet.getRange("RFIResponse!A3:A2000").copyFrom("RFIResponse!A2", Excel.RangeCopyType.formulas);
        sheet.getRange("RFIResponse!D2").formulas = [["=RFIData!D2"]];
        sheet.getRange("RFIResponse!D3:D2000").copyFrom("RFIResponse!D2", Excel.RangeCopyType.formulas);
        sheet.getRange("RFIResponse!I2").formulas = [["=RFIData!L2"]];
        sheet.getRange("RFIResponse!I3:I2000").copyFrom("RFIResponse!I2", Excel.RangeCopyType.formulas);
        sheet.getRange("RFIResponse!H2").copyFrom("RFIResponse!AA1");

    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("RFIVeiw");
        sheet.activate();
        sheet.getRange("B8").values = [[""]];

    });
};

async function SaveTransmittal() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("SubmittalData");
        const range = sheet.getRange("A2:BE2");

        range.insert(Excel.InsertShiftDirection.down);
        sheet.getRange("D2").copyFrom("Transmittal!F7", Excel.RangeCopyType.values);
        sheet.getRange("E2").copyFrom("Transmittal!D5", Excel.RangeCopyType.values);
        sheet.getRange("F2").copyFrom("Transmittal!A11", Excel.RangeCopyType.values);
        sheet.getRange("G2").copyFrom("Transmittal!A20", Excel.RangeCopyType.values);
        sheet.getRange("H2").copyFrom("Transmittal!A21", Excel.RangeCopyType.values);
        sheet.getRange("I2").copyFrom("Transmittal!A22", Excel.RangeCopyType.values);
        sheet.getRange("J2").copyFrom("Transmittal!A23", Excel.RangeCopyType.values);
        sheet.getRange("K2").copyFrom("Transmittal!A24", Excel.RangeCopyType.values);
        sheet.getRange("L2").copyFrom("Transmittal!A25", Excel.RangeCopyType.values);
        sheet.getRange("M2").copyFrom("Transmittal!A26", Excel.RangeCopyType.values);
        sheet.getRange("N2").copyFrom("Transmittal!A27", Excel.RangeCopyType.values);
        sheet.getRange("O2").copyFrom("Transmittal!A28", Excel.RangeCopyType.values);
        sheet.getRange("P2").copyFrom("Transmittal!A29", Excel.RangeCopyType.values);
        sheet.getRange("Q2").copyFrom("Transmittal!B20", Excel.RangeCopyType.values);
        sheet.getRange("R2").copyFrom("Transmittal!B21", Excel.RangeCopyType.values);
        sheet.getRange("S2").copyFrom("Transmittal!B22", Excel.RangeCopyType.values);
        sheet.getRange("T2").copyFrom("Transmittal!B23", Excel.RangeCopyType.values);
        sheet.getRange("U2").copyFrom("Transmittal!B24", Excel.RangeCopyType.values);
        sheet.getRange("V2").copyFrom("Transmittal!B25", Excel.RangeCopyType.values);
        sheet.getRange("W2").copyFrom("Transmittal!B26", Excel.RangeCopyType.values);
        sheet.getRange("X2").copyFrom("Transmittal!B27", Excel.RangeCopyType.values);
        sheet.getRange("Y2").copyFrom("Transmittal!B28", Excel.RangeCopyType.values);
        sheet.getRange("Z2").copyFrom("Transmittal!B29", Excel.RangeCopyType.values);
        sheet.getRange("AA2").copyFrom("Transmittal!C20", Excel.RangeCopyType.values);
        sheet.getRange("AB2").copyFrom("Transmittal!C21", Excel.RangeCopyType.values);
        sheet.getRange("AC2").copyFrom("Transmittal!C22", Excel.RangeCopyType.values);
        sheet.getRange("AD2").copyFrom("Transmittal!C23", Excel.RangeCopyType.values);
        sheet.getRange("AE2").copyFrom("Transmittal!C24", Excel.RangeCopyType.values);
        sheet.getRange("AF2").copyFrom("Transmittal!C25", Excel.RangeCopyType.values);
        sheet.getRange("AG2").copyFrom("Transmittal!C26", Excel.RangeCopyType.values);
        sheet.getRange("AH2").copyFrom("Transmittal!C27", Excel.RangeCopyType.values);
        sheet.getRange("AI2").copyFrom("Transmittal!C28", Excel.RangeCopyType.values);
        sheet.getRange("AJ2").copyFrom("Transmittal!C29", Excel.RangeCopyType.values);
        sheet.getRange("AK2").copyFrom("Transmittal!A35", Excel.RangeCopyType.values);
        sheet.getRange("AL2").copyFrom("Transmittal!A43", Excel.RangeCopyType.values);
        sheet.getRange("AM2").copyFrom("Transmittal!A44", Excel.RangeCopyType.values);
        sheet.getRange("AN2").copyFrom("Transmittal!A45", Excel.RangeCopyType.values);
        sheet.getRange("AO2").copyFrom("Transmittal!A46", Excel.RangeCopyType.values);
        sheet.getRange("AP2").copyFrom("Transmittal!A47", Excel.RangeCopyType.values);
        sheet.getRange("AQ2").copyFrom("Transmittal!A41", Excel.RangeCopyType.values);
        sheet.getRange("AR2").copyFrom("Transmittal!D8", Excel.RangeCopyType.values);
        sheet.getRange("BD2").copyFrom("Transmittal!E43", Excel.RangeCopyType.values);
        sheet.getRange("AS2").formulas = [["=VLOOKUP(D2,SubmittalStatus!$A$2:$CF$5148,7,FALSE)"]];
        sheet.getRange("AS3:AS4000").copyFrom("SubmittalData!AS2", Excel.RangeCopyType.formulas);
        sheet.getRange("AT2").formulas = [["=CONCAT(G2:P2)"]];
        sheet.getRange("AT3:AT4000").copyFrom("SubmittalData!AT2", Excel.RangeCopyType.formulas);
        sheet.getRange("AU2").formulas = [["=VLOOKUP(D2,SubmittalSchedule!$A$5:$CF$5148,3,FALSE)"]];
        sheet.getRange("AU3:AU4000").copyFrom("SubmittalData!AU2", Excel.RangeCopyType.formulas);
        sheet.getRange("AV2").formulas = [["=VLOOKUP(D2,SubmittalStatus!$A$2:$CF$5148,9,FALSE)"]];
        sheet.getRange("AV3:AV4000").copyFrom("SubmittalData!AV2", Excel.RangeCopyType.formulas);
        sheet.getRange("AW2").formulas = [["=VLOOKUP(D2,SubmittalStatus!$A$2:$CF$5148,10,FALSE)"]];
        sheet.getRange("AW3:AW4000").copyFrom("SubmittalData!AW2", Excel.RangeCopyType.formulas);
        sheet.getRange("AX2").formulas = [["=VLOOKUP(D2,SubmittalStatus!$A$2:$CF$5148,11,FALSE)"]];
        sheet.getRange("AX3:AX4000").copyFrom("SubmittalData!AX2", Excel.RangeCopyType.formulas);
        sheet.getRange("AY2").formulas = [["=CONCAT(AA2:AJ2)"]];
        sheet.getRange("AY3:AY4000").copyFrom("SubmittalData!AY2", Excel.RangeCopyType.formulas);
        sheet.getRange("AZ2").formulas = [["=SubmittalData!D2"]];
        sheet.getRange("AZ3:AZ4000").copyFrom("SubmittalData!AZ2", Excel.RangeCopyType.formulas);
        sheet.getRange("BA2").values = [["-"]];
        sheet.getRange("BA3:BA4000").copyFrom("SubmittalData!BA2", Excel.RangeCopyType.formulas);
        sheet.getRange("BB2").formulas = [["=SubmittalData!AR2"]];
        sheet.getRange("BB3:BB4000").copyFrom("SubmittalData!BB2", Excel.RangeCopyType.formulas);
        sheet.getRange("BC2").formulas = [["=CONCAT(AZ2:BB2)"]];
        sheet.getRange("BC3:BC4000").copyFrom("SubmittalData!BC2", Excel.RangeCopyType.formulas);
        sheet.getRange("A1").formulas = [["=SubmittalData!AY1"]];
        sheet.getRange("A2:A4000").copyFrom("SubmittalData!A1", Excel.RangeCopyType.formulas);


    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("Transmittal");
        sheet.getRange("F7").values = [[""]];
        sheet.getRange("A20").values = [[""]];
        sheet.getRange("A21").values = [[""]];
        sheet.getRange("A22").values = [[""]];
        sheet.getRange("A23").values = [[""]];
        sheet.getRange("A24").values = [[""]];
        sheet.getRange("A25").values = [[""]];
        sheet.getRange("A26").values = [[""]];
        sheet.getRange("A27").values = [[""]];
        sheet.getRange("A28").values = [[""]];
        sheet.getRange("A29").values = [[""]];
        sheet.getRange("B20").values = [[""]];
        sheet.getRange("B21").values = [[""]];
        sheet.getRange("B22").values = [[""]];
        sheet.getRange("B23").values = [[""]];
        sheet.getRange("B24").values = [[""]];
        sheet.getRange("B25").values = [[""]];
        sheet.getRange("B26").values = [[""]];
        sheet.getRange("B27").values = [[""]];
        sheet.getRange("B28").values = [[""]];
        sheet.getRange("B29").values = [[""]];
        sheet.getRange("A35").values = [[""]];
        sheet.getRange("A41").values = [[""]];
        sheet.getRange("A43").values = [[""]];
        sheet.getRange("A44").values = [[""]];
        sheet.getRange("A45").values = [[""]];
        sheet.getRange("A46").values = [[""]];
        sheet.getRange("A47").values = [[""]];
        sheet.getRange("E43").values = [[""]];
        sheet.getRange("D5").values = [[""]];
        sheet.getRange("C20").values = [["=VLOOKUP(F7,SubmittalSchedule!$A$2:$CE$9364,24,FALSE)"]];
        sheet.getRange("C21").values = [["=VLOOKUP(F7,SubmittalSchedule!$A$2:$CE$9364,25,FALSE)"]];
        sheet.getRange("C22").values = [["=VLOOKUP(F7,SubmittalSchedule!$A$2:$CE$9364,26,FALSE)"]];
        sheet.getRange("C23").values = [["=VLOOKUP(F7,SubmittalSchedule!$A$2:$CE$9364,27,FALSE)"]];
        sheet.getRange("C24").values = [["=VLOOKUP(F7,SubmittalSchedule!$A$2:$CE$9364,28,FALSE)"]];
        sheet.getRange("C25").values = [["=VLOOKUP(F7,SubmittalSchedule!$A$2:$CE$9364,29,FALSE)"]];
        sheet.getRange("C26").values = [["=VLOOKUP(F7,SubmittalSchedule!$A$2:$CE$9364,30,FALSE)"]];
        sheet.getRange("C27").values = [["=VLOOKUP(F7,SubmittalSchedule!$A$2:$CE$9364,31,FALSE)"]];
        sheet.getRange("C28").values = [["=VLOOKUP(F7,SubmittalSchedule!$A$2:$CE$9364,32,FALSE)"]];
        sheet.getRange("C29").values = [["=VLOOKUP(F7,SubmittalSchedule!$A$2:$CE$9364,33,FALSE)"]];


    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("SubmittalLog");
        let table12 = sheet.tables.getItem("Table12");

        let sortFields = [
            {
                key: 1,
                ascending: false
            }
        ];
        table12.sort.apply(sortFields);

        sheet.getRange("A11").formulas = [["=SubmittalData!BC2"]];
        sheet.getRange("A12:A4000").copyFrom("SubmittalLog!A11", Excel.RangeCopyType.formulas);
        sheet.getRange("B11").formulas = [["=SubmittalData!AY2"]];
        sheet.getRange("B12:B4000").copyFrom("SubmittalLog!B11", Excel.RangeCopyType.formulas);
        sheet.getRange("C11").formulas = [["=SubmittalData!AQ2"]];
        sheet.getRange("C12:C4000").copyFrom("SubmittalLog!C11", Excel.RangeCopyType.formulas);
        sheet.getRange("D11").formulas = [["=SubmittalData!AU2"]];
        sheet.getRange("D12:D4000").copyFrom("SubmittalLog!D11", Excel.RangeCopyType.formulas);
        sheet.getRange("E11").formulas = [["=SubmittalData!AS2"]];
        sheet.getRange("E12:E4000").copyFrom("SubmittalLog!E11", Excel.RangeCopyType.formulas);
        sheet.getRange("F11").formulas = [["=SubmittalData!AT2"]];
        sheet.getRange("F12:F4000").copyFrom("SubmittalLog!F11", Excel.RangeCopyType.formulas);
        sheet.getRange("G11").formulas = [["=SubmittalData!E2"]];
        sheet.getRange("G12:G4000").copyFrom("SubmittalLog!G11", Excel.RangeCopyType.formulas);
        sheet.getRange("H11").formulas = [["=SubmittalData!AV2"]];
        sheet.getRange("H12:H4000").copyFrom("SubmittalLog!H11", Excel.RangeCopyType.formulas);
        sheet.getRange("I11").formulas = [["=SubmittalData!AW2"]];
        sheet.getRange("I12:I4000").copyFrom("SubmittalLog!I11", Excel.RangeCopyType.formulas);
        sheet.getRange("J11").formulas = [["=SubmittalData!AX2"]];
        sheet.getRange("J12:J4000").copyFrom("SubmittalLog!J11", Excel.RangeCopyType.formulas);
        sheet.getRange("K11").formulas = [["=SubmittalData!BD2"]];
        sheet.getRange("K12:K4000").copyFrom("SubmittalLog!K11", Excel.RangeCopyType.formulas);
        sheet.tables.sortFields


    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("SubmittalStatus");
        const range = sheet.getRange("A2:AK2");

        range.insert(Excel.InsertShiftDirection.down);
        sheet.getRange("SubmittalStatus!D2").formulas = [["=SubmittalData!D2"]];
        sheet.getRange("SubmittalStatus!D3:D4000").copyFrom("SubmittalStatus!D2", Excel.RangeCopyType.formulas);
        sheet.getRange("SubmittalStatus!A2").formulas = [["=SubmittalStatus!D2"]];
        sheet.getRange("SubmittalStatus!A3:A4000").copyFrom("SubmittalStatus!A2", Excel.RangeCopyType.formulas);
        sheet.getRange("SubmittalStatus!E2").formulas = [["=SubmittalData!AQ2"]];
        sheet.getRange("SubmittalStatus!E3:E4000").copyFrom("SubmittalStatus!E2", Excel.RangeCopyType.formulas);
        sheet.getRange("SubmittalStatus!F2").formulas = [["=SubmittalData!AU2"]];
        sheet.getRange("SubmittalStatus!F3:F4000").copyFrom("SubmittalStatus!F2", Excel.RangeCopyType.formulas);
        sheet.getRange("SubmittalStatus!H2").formulas = [["=SubmittalData!E2"]];
        sheet.getRange("SubmittalStatus!H3:H4000").copyFrom("SubmittalStatus!H2", Excel.RangeCopyType.formulas);
        sheet.getRange("SubmittalStatus!L2").formulas = [["=SubmittalData!AY2"]];
        sheet.getRange("SubmittalStatus!L3:L4000").copyFrom("SubmittalStatus!L2", Excel.RangeCopyType.formulas);
        sheet.getRange("SubmittalStatus!K2").copyFrom("SubmittalStatus!AA1");

    });

    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("TransmittalView");
        sheet.activate();
        sheet.getRange("H7").values = [[""]];

    });

}
