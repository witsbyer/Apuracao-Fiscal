class ContaContabil {
    constructor(sheet,filtro) {
        this.sheet = sheet;
        this.sheet.getRange("I:I").setNumberFormat("#,##0.00");
        this.sheet.getRange("A:A").setNumberFormat("@");
        this.geraCodigoSemPonto();
        this.sheet.hideSheet();  
        this.filtro = filtro;
        this.aplicarFiltros();
    }
    
    aplicarFiltros() {
        let valoresColunaA = this.filtro.getDataRange().getValues().map(linha => { return linha[0] });
        let valoresContas = this.sheet.getDataRange().getValues();
        let novosValoresContas = valoresContas((linha, index) => {
            let valor = linha[17];
            if (index < 2) {
                return linha;
            }
            if (valoresColunaA.includes(valor)) {
                return linha;
            }
            else {
                if (valor == '') {
                    return linha;
                }
            }
        });

        this.sheet.clear();
        this.sheet.getRange(1, 1, novosValoresColunas.length, 19).setValues(novosValoresContas);
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