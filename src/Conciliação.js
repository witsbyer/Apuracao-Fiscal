class Conciliacao {
    constructor(sheet, stf) {
        this.sheet = sheet;
        this.stf = stf;
        this.cellValoresComparacaoInicial = null;
        this.cellRangeQuery = null;
        this.rangeQueryCompleto = null;

        this.setNumberFormatInRange(this.sheet.getRange("G:I"), "#,##0.00");
        this.sheet.hideColumn(this.sheet.getRange("J1"));
        this.sheet.hideColumn(this.sheet.getRange("K1"));
    }

    montaEstrutura() {
        let cellQuery = this.sheet.getRange(1, 1);
        let query = `=query('${this.cadastroFilais.sheet.getName()}'!${this.cadastroFilais.sheet.getDataRange().getA1Notation()}; "select A, B, C, D, G, F where A = ${information.idCliente}")`
        setValueForRange(cellQuery, query);

        let cellCContabilUnique = this.sheet.getRange(1, 7);
        let formulaUnique = `=unique(F2:F)`;
        setValueForRange(cellCContabilUnique, formulaUnique);

        let cellContUnique = this.sheet.getRange(1, 8);
        let formulaContUnique = `=COUNTA(G1:G)`;
        setValueForRange(cellContUnique, formulaContUnique);

        let valorCont = cellContUnique.getValue();

        let valoresUnicos = cellCContabilUnique.offset(0, 0,valorCont, 1).getValues();

        this.sheet.clearContents();

        cellCContabilUnique = cellCContabilUnique.offset(0, -1, valorCont, 1);

        setValuesForRange(cellCContabilUnique, valoresUnicos);

        cellQuery = cellQuery.offset(valorCont + 4, 0, 1, 1);
        setValueForRange(cellQuery, query);

        let rangeSoma = cellCContabilUnique.offset(0, 1, 1, 1);

        for (let index = 1; index <= valorCont; index++) {
            let valeuRangeAoLado = rangeSoma.offset(0, -1).getValue();
            let formula = `=SUMIF(F${valorCont + 4}:F; "${valeuRangeAoLado}";G${valorCont + 4}:G)`;
            let formula2 = `=SUMIF(F${valorCont + 4}:F; "${valeuRangeAoLado}";H${valorCont + 4}:H)`;
            let formula3 = `=SUMIF(F${valorCont + 4}:F; "${valeuRangeAoLado}";I${valorCont + 4}:I)`;

            setValueForRange(rangeSoma, formula);
            setValueForRange(rangeSoma.offset(0, 1), formula2);
            setValueForRange(rangeSoma.offset(0, 2), formula3);
            rangeSoma = rangeSoma.offset(1, 0);
        }

        let totais = [
            [
                `=sum(G1:G${valorCont})`,
                `=sum(H1:H${valorCont})`,
                `=sum(I1:I${valorCont})`,
            ],
            [
                `=query('${tableContaContabil.sheet.getName()}'!${tableContaContabil.sheet.getDataRange().getA1Notation()}; "select SUM(I) LABEL SUM(I) ''")`,
                `=-query('${tableImpostos.sheet.getName()}'!${tableImpostos.sheet.getDataRange().getA1Notation()}; "select SUM(K)+SUM(L) LABEL SUM(K)+SUM(L) ''")`,
                `=${rangeSoma.offset(1, 0).getA1Notation()}-${rangeSoma.offset(1, 1).getA1Notation()}`
            ],
            [
                `=${rangeSoma.getA1Notation()}-${rangeSoma.offset(1, 0).getA1Notation()}`,
                `=${rangeSoma.offset(0, 1).getA1Notation()}-${rangeSoma.offset(1, 1).getA1Notation()}`,
                `=${rangeSoma.offset(0, 2).getA1Notation()}-${rangeSoma.offset(1, 2).getA1Notation()}`
            ]
        ]

        let formulaIdJuncao = `=arrayformula(if(len(A${cellQuery.getRow() + 1}:A)<>0; F${cellQuery.getRow() + 1}:F&E${cellQuery.getRow() + 1}:E; ))`
        let formulaIdJuncao2 = `=arrayformula(if(len(A${cellQuery.getRow() + 1}:A)<>0; B${cellQuery.getRow() + 1}:B&C${cellQuery.getRow() + 1}:C; ))`

        setValueForRange(rangeSoma.offset(0, -1), "Totais comparativos: ")
        setValueForRange(rangeSoma.offset(1, -1), "Totais outras tabelas: ")
        setValueForRange(rangeSoma.offset(2, -1), "Diferença: ")
        setValuesForRange(rangeSoma.offset(0, 0, 3, 3), totais);

        let rangeSaldos = cellQuery.offset(0, 6, 1, 3);
        setValuesForRange(rangeSaldos, [["Saldo conta", "Saldo impostos", "Diferença"]]);
        setValuesForRange(rangeSaldos.offset(0, 3, 2, 2), [["ID juncao", ""], [formulaIdJuncao, formulaIdJuncao2]]);
        this.cellValoresComparacaoInicial = rangeSaldos.offset(1, 0, 1, 1);
        this.cellRangeQuery = cellQuery;
    }

    preencheDadosComparacao() {
        let rangeQuery = this.sheet.getRange(`A${this.cellRangeQuery.getRow() + 1}:I${this.sheet.getDataRange().getLastRow()}`);
        let valores = rangeQuery.getDisplayValues();
        let arrayParaSet = [];

        valores.forEach((linha, index) => {
            let contaContabil = linha[4];
            let filial = linha[2];

            let queryItemConta = `=iferror(query('${tableContaContabil.sheet.getName()}'!${tableContaContabil.sheet.getDataRange().getA1Notation()}; "select sum(I) WHERE C = '${contaContabil}' and A = '${this.formataContaContabil(contaContabil)}' LABEL SUM(I) ''");0)`;
            let queryImpostosAReceber = `=-iferror(query('${tableImpostos.sheet.getName()}'!${tableImpostos.sheet.getDataRange().getA1Notation()}; "select sum(K)+sum(L) WHERE A contains '${filial}' LABEL sum(K)+sum(L) ''");0)`;
            let formulaDiferenca = `=iferror(MINUS(${rangeQuery.getCell(index + 1, 7).getA1Notation()};${rangeQuery.getCell(index + 1, 8).getA1Notation()});0)`

            arrayParaSet.push([queryItemConta, queryImpostosAReceber, formulaDiferenca]);
        });

        rangeQuery.offset(0, 6, valores.length, 3).setValues(arrayParaSet)

        this.rangeQueryCompleto = rangeQuery;
    }

    geraAnaliseContabilVsFiscal() {
        let sheetAnalise = ss.insertSheet().setName("Análise ITEM CONTÁBIL VS DIFERENÇAS");
        let query = `=query('${this.sheet.getName()}'!${this.rangeQueryCompleto.getA1Notation()}; "select * order by I")`
        setValueForRange(sheetAnalise.getRange(1, 1), query);
    }

    setNumberFormatInRange(range, type) {
        range.setNumberFormat(type);
    }

    formataContaContabil(contaCont) {
        if (contaCont !== "---") {
            let arrayCaract = contaCont.split('');
            let stringFormada = `${arrayCaract[0]}.${arrayCaract[1]}.${arrayCaract[2]}.${arrayCaract[3]}${arrayCaract[4]}${arrayCaract[5]}.${arrayCaract[6]}${arrayCaract[7]}${arrayCaract[8]}${arrayCaract[9]}`
            return stringFormada;
        } else {
            return "✗";
        }
    }

    formataCodigo(cod) {
        let zeros = '000000';
        let juncao = zeros + cod;
        if (isNaN(+cod)) {
            return cod;
        }
        if (juncao.length > 6) {
            let quantidade = juncao.length;
            let juncaoCerta = juncao.substring(quantidade - 6);

            return juncaoCerta;
        }
    }
}