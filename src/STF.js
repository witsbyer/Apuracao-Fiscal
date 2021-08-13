class SFT {
    constructor(sheet) {
        this.sheet = sheet;
        this.cellValoresComparacaoInicial = null;
        this.cellRangeQuery = null;
        this.sheet.getRange("I:I").setNumberFormat("@");
        this.sheet.getRange("A:A").setNumberFormat("@");
        this.sheet.hideSheet();
    }

    createConciliacao() {
        let conciliacao = ss.insertSheet().setName(`Conciliação De Apuração - ${information.cliente}`);
        return new Conciliacao(conciliacao, this);
    }

    createDataComplementImpostosARecolher() {
        let sheet = ss.insertSheet().setName("Valores para inserção");
        let cellInicial = sheet.getRange(1, 1);

        let formula = `=iferror(unique(query('${tableImpostos.sheet.getName()}'!A3:S; "select A, Q, R, P where S = 'X'"));"NULL")`;

        cellInicial.setValue(formula);

        if (cellInicial.getValue() !== "NULL") {
            let dataRange = sheet.getDataRange();

            let valores = dataRange.getDisplayValues();

            let arrayParaSet = [];

            valores.forEach(linha => {
                let filial= linha[0].substring(10) || "---";
                let contaContabil = linha[3] || "---";//se for uma string vazia, coloca 3 traços pra não dar problema futuro

                arrayParaSet.push([information.idCliente < 10 ? "0" + information.idCliente : information.idCliente, "✗", contaContabil,filial, "✗"])
            });

            this.sheet.getRange(this.sheet.getDataRange().getLastRow() + 1, 1, arrayParaSet.length, 7).setValues(arrayParaSet);
            deleteTableIfExist(sheet.getName());
            return true;
        }
        else {
            deleteTableIfExist(sheet.getName());
            return false;
        }
    }

    createDataComplementContaContabil() {
        let sheet = ss.insertSheet().setName("Contas para inserção");
        let cellInicial = sheet.getRange(1, 1);

        let formula = `=iferror(unique(query('${tableContaContabil.sheet.getName()}'!A3:K; "select C, D, J where K = 'X'"));"NULL")`;

        cellInicial.setValue(formula);

        if (cellInicial.getValue() !== "NULL") {
            let dataRange = sheet.getDataRange();

            let valores = dataRange.getDisplayValues();

            let arrayParaSet = [];

            valores.forEach(linha => {
                let filial = linha[1] || "---";
                let contaContabil = linha[0] || "---";

                arrayParaSet.push([information.idCliente < 10 ? "0" + information.idCliente : information.idCliente, "✗", "✗", filial, "✗", contaContabil])
            });

            this.sheet.getRange(this.sheet.getDataRange().getLastRow() + 1, 1, arrayParaSet.length, 7).setValues(arrayParaSet);
            deleteTableIfExist(sheet.getName());
            return true;
        }
        else {
            deleteTableIfExist(sheet.getName());
            return false;
        }
    }

    createDataComplement() {
        let verificadorImpostos = this.createDataComplementImpostosAReceber();
        let verificadorConta = this.createDataComplementContaContabil();
        if (verificadorConta || verificadorImpostos) {
            repetidor = true;
            contador++;
        } else {
            repetidor = false;
        }
    }
}
