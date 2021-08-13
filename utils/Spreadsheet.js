//NOTE Tudo relacionado com utilitário do Spreadsheet da google

const ss = SpreadsheetApp.getActiveSpreadsheet();
const dashBoard = ss.getSheetByName('Dashboard');
const dados = getRangesImportant();
const information = getInformation();
let tableSTF = new SFT(ss.getSheetByName('Data Base'));
let tableConciliacao = null;
let tableImpostos = null;
let tableContaContabil = null;
const filtros = ss.getSheetByName('Filtro de contas');

function getInformation() {
    return {
        cliente: dashBoard.getRange("B1").getValue().toString().substr(5),
        idCliente: dashBoard.getRange("B1").getValue().toString().substr(0, 2),
        ano: dashBoard.getRange("B2").getValue().toString(),
        mes: dashBoard.getRange("B3").getValue().toString(),
        cotacoes: {
            dolar: dashBoard.getRange("F2").getDisplayValue(),
            euro: dashBoard.getRange("F4").getDisplayValue()
        }
    }
}

function getRangesImportant() {
    return {
        rangePasso1: dashBoard.getRange("B8"),
        rangePasso2: dashBoard.getRange("B12"),
        rangePasso3: dashBoard.getRange("B16"),
        rangePasso4: dashBoard.getRange("B20"),
        rangePasso5: dashBoard.getRange("B24"),
        rangePasso6: dashBoard.getRange("B29"),
        rangePasso7: dashBoard.getRange("B33"),
        rangePasso8: dashBoard.getRange("B37"),
        rangePasso9: dashBoard.getRange("B41"),
        rangePasso10: dashBoard.getRange("B46"),
        rangePasso11: dashBoard.getRange("B50"),
        rangePasso12: dashBoard.getRange("B54"),
        rangePasso13: dashBoard.getRange("B58"),
    }
}

function setValuesOfVerification() {//Seta valores de verificação 
    let dado = ['Aguardando!'];
    dados.rangePasso1.setValue(dado).setBackground("#FFDB6D");
    dados.rangePasso2.setValue(dado).setBackground("#FFDB6D");
    dados.rangePasso3.setValue(dado).setBackground("#FFDB6D");
    dados.rangePasso4.setValue(dado).setBackground("#FFDB6D");
    dados.rangePasso5.setValue(dado).setBackground("#FFDB6D");
}

function clearSS() { //restaura o dashboard par ao padrão
    setValuesOfVerification();
    dados.rangePasso1.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso2.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso3.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso4.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso5.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso6.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso7.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso8.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso9.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso10.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso11.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso12.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso13.offset(0, 1, 1, 12).clearContent().clearFormat();
}

function clearBlockBlue() { //restaura o dashboard par ao padrão
    setValuesOfVerificationBlockBlue();
    dados.rangePasso1.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso2.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso3.offset(0, 1, 1, 12).clearContent().clearFormat();
    dados.rangePasso4.offset(0, 1, 1, 12).clearContent().clearFormat();
}

function setValuesOfVerificationBlockBlue() {//Seta valores de verificação 
    let dado = ['Aguardando!'];
    dados.rangePasso1.setValue(dado).setBackground("#FFDB6D");
    dados.rangePasso2.setValue(dado).setBackground("#FFDB6D");
    dados.rangePasso3.setValue(dado).setBackground("#FFDB6D");
    dados.rangePasso4.setValue(dado).setBackground("#FFDB6D");
}

function deleteTableIfExist(nome) {
    let verificador = ss.getSheetByName(nome);
    if (verificador != null) {
        ss.deleteSheet(verificador);
    }
}

function removeTableIfExist(table) {
    let sheets = ss.getSheets();
    sheets.forEach(sheet => {
        if (sheet.getName().includes(table)) {
            ss.deleteSheet(sheet);
        }
    });
}

function removeAllTable() { //Função apaga todas as tabelas para criar conciliação menos as indicadas a baixo
    let sheets = ss.getSheets();
    let regex = /Dashboard|Filtro|Data Base/;
    sheets.forEach(sheet => {
        if (!regex.exec(sheet.getName())) {
            ss.deleteSheet(sheet);
        }
    });
}

function getTablesForId(id) {
    let spreadSheet = SpreadsheetApp.openById(id);
    return spreadSheet.getSheets();
}

function copyTablesToDestination(tabelas, destino, tabela) {
    return tabelas[tabela].copyTo(destino);
}

function formataTable(type, range) {
    switch (type) {
        case 'OK':
            range.setBackground('#7DFF6D');
            break;
        case 'NO_OK':
            range.getCell(1, 1).setBackground('#FF5252');
            range.getCell(1, 2).setBackground('#FFDB6D');
            break;
        default:
            Logger.log('Opção não encontrada!');
            break;
    }
}

function setValuesForRange(range, dados) {//?seta vários valores
    range.setValues(dados);
}

function setValueForRange(range, dado) {//?seta um valor
    range.setValue(dado);
}

function setMesesFaltantes() {
    if (files.noOk.length > 0) {
        let hyperlink = '=HYPERLINK("' + files.noOk[0].folder.getUrl() + '"; "' + files.noOk[0].folder.getName() + '")';
        let meses = ["NÃO", hyperlink];
        let query = [meses];

        formataTable("NO_OK", dados.rangePasso1.offset(0, 0, 1, 2));
        setValuesForRange(dados.rangePasso1.offset(0, 0, 1, 2), query);
        return false;
    }
    else {
        formataTable("OK", dados.rangePasso1);
        setValueForRange(dados.rangePasso1, ["SIM"]);
        return true;
    }
}

function createTablesRequireds() {
    try {
        files.oks.forEach(arquivo => {

            let arquivoConvertido = convertXlsToSpreadSheet(arquivo.file);
            let tablesConvertidas = getTablesForId(arquivoConvertido.getId());

            let fileNameNoMime = arquivo.file.getName().replace(".xlsx", "");
            removeTableIfExist(fileNameNoMime, ss);

            let sheet = copyTablesToDestination(tablesConvertidas, ss, tablesConvertidas.length - 1);
            sheet.showColumns(1, sheet.getLastColumn())
            sheet.showRows(1, sheet.getLastRow());
            sheet.setName(fileNameNoMime);

            deleteFileById(arquivoConvertido.getId());

            if (fileNameNoMime.includes("Conta Contabil")) {
                tableContaContabil = new ContaContabil(sheet,filtros);
            }
            if (fileNameNoMime.includes("Impostos a Recolher")) {
                tableImpostosARecolher = new ImpostosARecolher(sheet, filtros);
            }

        });

        formataTable('OK', dados.rangePasso2);
        setValueForRange(dados.rangePasso2, ["SIM"]);
        setDefaultFiles(files)
        return true;

    } catch (e) {
        let query = [["Não", e.message + "(" + e.stack + ")"]]
        setValuesForRange(dados.rangePasso2.offset(0, 0, 1, 2), query);
        formataTable('NO_OK', dados.rangePasso2.offset(0, 0, 1, 2));
        setDefaultFiles(files)
        return false;
    }
}

function gerarTabelasdeComparacao() {
    try {
        if (tableSTF) {
            let rangeContas = tableSTF.getRange(3, 1, tableSTF.getDataRange().getLastRow() - 2);

            for (let index = 0; index < (rangeContas.getLastRow() - 2); index++) {
                let valueConta = rangeContas.getCell(index + 1, 1).getValue();

                let sheetContaContabilComparativo = ss.insertSheet();
                let sheetImpostosARecolherComparativo = ss.insertSheet();

                let query = `=query('${tableContaContabil.getName()}'!${tableContaContabil.getDataRange().getA1Notation()}; "select * where A = '${valueConta}'")`;
                let cellInicialQuery = sheetContaContabilComparativo.getDataRange().getCell(1, 1);
                let query2 = `=query('${tableTitulos.getName()}'!${tableImpostos.getDataRange().getA1Notation()}; "select * where N = '${valueConta.replace(/\./g, '')}'")`;
                let cellInicialQuery2 = sheetImpostosARecolherComparativo.getDataRange().getCell(1, 1);

                setValueForRange(cellInicialQuery, query);
                setValueForRange(cellInicialQuery2, query2);

                let valorVerificador = cellInicialQuery.offset(1, 0).getValue();
                let valorVerificador2 = cellInicialQuery2.offset(1, 0).getValue();

                if (valorVerificador === "" || valorVerificador2 === "") {
                    deleteTableIfExist(sheetContaContabilComparativo.getName());
                    deleteTableIfExist(sheetImpostosARecolherComparativo.getName());
                }
                else {
                    let dadosParacomparar = {
                        conta: valueConta,
                        sheetTitulosAReceber: null,
                        sheetItemContabil: null,
                        sheetConciliacao: null
                    }

                    sheetContaContabilComparativo.setName(`Conta Contabil Apoio - ${valueConta}`);
                    dadosParacomparar.sheetContaContabil = sheetContaContabilComparativo;

                    sheetImpostosARecolherComparativo.setName(`Impostos A Recolher Apoio - ${valueConta}`);
                    dadosParacomparar.sheetImpostosARecolher = sheetImpostosARecolherComparativo;

                    tablesForComparation.push(dadosParacomparar);
                }

            }
            formataTable('OK', dados.rangePasso3);
            setValueForRange(dados.rangePasso3, ["SIM"]);
            return true;
        }
        else {
            let query = [["Não", "Tabela de STF ausente!"]];
            setValuesForRange(dados.rangePasso3.offset(0, 0, 1, 2), query);
            formataTable('NO_OK', dados.rangePasso3.offset(0, 0, 1, 2));
            return false;
        }
    } catch (e) {
        let query = [["Não", e.message + "(" + e.stack + ")"]];
        setValuesForRange(dados.rangePasso3.offset(0, 0, 1, 2), query);
        formataTable('NO_OK', dados.rangePasso3.offset(0, 0, 1, 2));
        return false;
    }
}

function gerarConciliacoes() {
    try {
        tablesForComparation.forEach(table => {
            let sheetConciliacao = ss.insertSheet().setName(`Conciliação - ${table.conta}`);
            table.sheetConciliacao = sheetConciliacao;

            let queryContaContabil = `=QUERY('${table.sheetContaContabil.getName()}'!A:O; "SELECT C, D, I where C is not null order by D")`;
            let formulaContador = `=COUNTA(E2:E)`;

            let valoresJaComparados = [];

            let rangeSheetConciliacao = sheetConciliacao.getRange("A1:M");
            let cellInicial = rangeSheetConciliacao.getCell(1, 1);
            setValueForRange(cellInicial, queryContaContabil);

            let rangeQueryContaContabilComparacao = sheetConciliacao.getDataRange();

            let cellContador = rangeSheetConciliacao.getCell(1, 4);
            setValueForRange(cellContador, formulaContador);

            let cellValuesUnificados = rangeSheetConciliacao.getCell(1, 9).offset(0, 0, 1, 3);
            setValuesForRange(cellValuesUnificados, cellInicial.offset(0, 0, 1, 3).getValues());
            let indexParaVolta = 0; //NOTE ele é usado para determinar quantas linhas o offset precisa voltar na hora de unificar
            for (let index = 0; index < rangeQueryContaContabilComparacao.getNumRows() - 1; index++) {

                let cellComparacaoInicial = rangeSheetConciliacao.getCell(1, 5);
                let offset = rangeQueryContaContabilComparacao.offset(index + 1, 1, 1, 1);
                let valor = offset.getValue();
                let valorQuebrado = valor.split(' ');
                typeof valorQuebrado[1] === "undefined" ? valorQuebrado[1] = "" : valorQuebrado[1];
                typeof valorQuebrado[2] === "undefined" ? valorQuebrado[2] = "" : valorQuebrado[2];

                let queryContaContabilComparacao = `=query(A:C; "select * where B contains '${valorQuebrado[0]}' AND B CONTAINS '${valorQuebrado[1]}'")`;
                setValueForRange(cellComparacaoInicial, queryContaContabilComparacao);
                if (cellContador.getValue() == 1) {
                    let valuesOffset = cellComparacaoInicial.offset(1, 0, 1, 3).getValues();
                    setValuesForRange(cellValuesUnificados.offset(index + 1 - indexParaVolta, 0), valuesOffset);
                }
                else if (cellContador.getValue() > 1) {
                    let offsetParaUnificacao = cellComparacaoInicial.offset(1, 0, 1, 3);
                    let verificador = true;
                    let valoresAchados = [{
                        valor: offsetParaUnificacao.getCell(1, 3).getValue(),
                        desc: offsetParaUnificacao.getCell(1, 2).getValue(),
                        cod: offsetParaUnificacao.getCell(1, 1).getValue()
                    }];
                    let objFinalUnificado = { soma: 0, desc: "", cod: "" };
                    while (verificador) {
                        offsetParaUnificacao = offsetParaUnificacao.offset(1, 0);
                        let valorAchado = {
                            valor: offsetParaUnificacao.getCell(1, 3).getValue(),
                            desc: offsetParaUnificacao.getCell(1, 2).getValue(),
                            cod: offsetParaUnificacao.getCell(1, 1).getValue()
                        }
                        if (valorAchado.valor) {
                            valoresAchados.push(valorAchado);
                            let maiorValor = valoresAchados[0].valor;
                            valoresAchados.forEach(e => {
                                objFinalUnificado.soma += e.valor;
                                if (maiorValor <= e.valor) {
                                    objFinalUnificado.desc = e.desc;
                                    objFinalUnificado.cod = e.cod
                                    maiorValor = e.valor;
                                }
                            });
                        }
                        else {
                            verificador = false;
                        }
                    }
                    if (!valoresJaComparados.includes(valoresJaComparados.find(el => el.cod === objFinalUnificado.cod))) {
                        valoresJaComparados.push(objFinalUnificado);
                        setValuesForRange(cellValuesUnificados.offset(index + 1 - indexParaVolta, 0), [[objFinalUnificado.cod, objFinalUnificado.desc, objFinalUnificado.soma]]);
                    }
                    else {
                        indexParaVolta += 1; //NOTE cada vez que tem repetido, o programa soma a quantidade de "-1" linhas par anão ocorrer espaços brancos
                    }
                }
            }
        });

        formataTable('OK', dados.rangePasso4);
        setValueForRange(dados.rangePasso4, ["SIM"]);
        return true;
    } catch (e) {
        let query = [["Não", e.message + "(" + e.stack + ")"]];
        setValuesForRange(dados.rangePasso4.offset(0, 0, 1, 2), query);
        formataTable('NO_OK', dados.rangePasso4.offset(0, 0, 1, 2));
        return false;
    }
}

function trataConciliacoesForContaContabil() {
    try {
        tablesForComparation.forEach(sheet => {
            let rangeUnificados = sheet.sheetConciliacao.getRange("I1:K");
            let valuesUnificados = rangeUnificados.getValues();
            sheet.sheetConciliacao.clear();
            setValuesForRange(rangeUnificados.offset(0, -8), valuesUnificados);
        })
        formataTable('OK', dados.rangePasso4);
        setValueForRange(dados.rangePasso4, ["SIM"]);
        return true;
    } catch (e) {
        let query = [["Não", e.message + "(" + e.stack + ")"]];
        setValuesForRange(dados.rangePasso4.offset(0, 0, 1, 2), query);
        formataTable('NO_OK', dados.rangePasso4.offset(0, 0, 1, 2));
        return false;
    }
}

function trataConciliacoesForImpostosARecolher() {
    try {
        tablesForComparation.forEach(sheet => {
            let cellQueryComparacao = sheet.sheetConciliacao.getRange("F1");
            let rangeContaContabil = sheet.sheetConciliacao.getDataRange();

            for (let index = 0; index < rangeContaContabil.getLastRow() - 1; index++) {
                let offsetPercorreContaContabil = rangeContaContabil.offset(index + 1, 0);
                let descSplit = offsetPercorreItemContabil.getCell(1, 2).getValue().split(" ");

                let queryComparacao = `=query('${sheet.sheetImpostosARecolher.getName()}'!1:1000; "select SUM(K) + SUM(L) where A contains '${descSplit[0]}' AND A contains '${descSplit[1]}'")`;
                setValueForRange(cellQueryComparacao, queryComparacao);
                let valorComparacao = cellQueryComparacao.offset(1, 0).getValue();
                if (valorComparacao === "") {
                    valorComparacao = "Valor não encontrado";
                }

                setValueForRange(offsetPercorreContaContabil.offset(0, 3, 1, 1), valorComparacao);
            }

        })
        formataTable('OK', dados.rangePasso4);
        setValueForRange(dados.rangePasso4, ["SIM"]);
        return true;
    } catch (e) {
        let query = [["Não", e.message + "(" + e.stack + ")"]];
        setValuesForRange(dados.rangePasso4.offset(0, 0, 1, 2), query);
        formataTable('NO_OK', dados.rangePasso4.offset(0, 0, 1, 2));
        return false;
    }
}