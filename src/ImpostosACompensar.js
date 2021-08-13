class ImpostosACompensar {
    constructor(sheet) {
        this.sheet = sheet;
        this.sheet.getRange("K:L").setNumberFormat("#,##0.00");
        this.sheet.getRange("T:U").setNumberFormat("@");
        this.setaZero();
        this.criarColunasParaProcV();
        this.sheet.hideSheet();
    }


    setaZero() {
        let rangeJK = this.sheet.getRange(`K1:L${this.sheet.getDataRange().getLastRow()}`);
        let valores = rangeJK.getValues();

        let novosValores = valores.map((item) => {
            if (item[0] == "") {
                item[0] = 0;
            }
            if (item[1] == "") {
                item[1] = 0;
            }
            return item;
        });
        rangeJK.setValues(novosValores);
    }


    criarColunasParaProcV() {
        let valoresColunaA = this.sheet.getRange(`A3:A${this.sheet.getDataRange().getLastRow()}`).getValues();
        let novasColunas = [];
        valoresColunaA.forEach(linha => {
            let valor = linha[0];
            let stringSeparada = valor.split("-");
            novasColunas.push([stringSeparada[0], stringSeparada[1]]);
        });

        this.sheet.getRange(`T3:U${this.sheet.getDataRange().getLastRow()}`).setValues(novasColunas);
    }

    geraColunaComparativaComDataBase() {
        let formula = `=arrayformula(IF(LEN(A3:A)<>0; IFERROR(VLOOKUP(T3:T&U3:U;'${tableConciliacao.sheet.getName()}'!K${tableConciliacao.cellRangeQuery.getRow() + 1}:L${tableConciliacao.rangeQueryCompleto.getLastRow()}; 1; FALSE); "X"); ""))`;
        this.sheet.getRange("S2:S3").setValues([["De para"], [formula]]);
    }
}
