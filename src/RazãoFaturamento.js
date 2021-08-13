class RazaoFaturamento {
    constructor(sheet) {
        this.sheet = sheet;
        this.sheet.getRange("B:M").setNumberFormat("#,##0.00");
        this.sheet.getRange("A:A").setNumberFormat("@");
        this.geraCodigoSemPonto();
        this.sheet.hideSheet();  
    }

    geraContaSemPonto() {
        let formula = `=arrayformula(IF(LEN(A3:A)<>0; SUBSTITUTE(A3:A; ".";""); ""))`;
        this.sheet.getRange("J2:J3").setValues([["Conta Contabil sem ponto"], [formula]]);
    }

    geraColunaComparativaComDataBase() {
        let formula = `=arrayformula(IF(LEN(A3:A)<>0; IFERROR(VLOOKUP(J3:J&C3:C;'${tableConciliacao.sheet.getName()}'!J${tableConciliacao.cellRangeQuery.getRow() + 1}:J${tableConciliacao.rangeQueryCompleto.getLastRow()}; 1; FALSE); "X"); ""))`;
        this.sheet.getRange("K2:K3").setValues([["De para"], [formula]]);
    }
}