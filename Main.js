//NOTE Função de inicio
var verificadorDePassagem = false;
var repetidor = true;
var contador = 0;

function main() {
    while (repetidor && contador < 2) {
        clearBlockBlue();
        removeAllTable();
        getFilesRequireds();

        verificadorDePassagem = setMesesFaltantes();//Vai definir se continua executando

        if (verificadorDePassagem) {
            verificadorDePassagem = createTablesRequireds();

        }
        if (verificadorDePassagem) {
            try {
                tableConciliacao = tableSTF.createConciliacao();
                tableConciliacao.montaEstrutura();
                tableConciliacao.preencheDadosComparacao();
                tableConciliacao.geraConciliacao();
                verificadorDePassagem = true
                formataTable('OK', dados.rangePasso3);
                setValueForRange(dados.rangePasso3, ["SIM"]);
            }
            catch (e) {
                verificadorDePassagem = false
                let query = [["Não", e.message + "(" + e.stack + ")"]]
                setValuesForRange(dados.rangePasso3.offset(0, 0, 1, 2), query);
                formataTable('NO_OK', dados.rangePasso3.offset(0, 0, 1, 2));
            }

        }
        if (verificadorDePassagem) {
            try {
                tableContaContabil.geraColunaComparativaComDataBase();
                tableARecolher.geraColunaComparativaComDataBase();
                tableSTF.createDataComplement();
                verificadorDePassagem = true
                formataTable('OK', dados.rangePasso4);
                setValueForRange(dados.rangePasso4, ["SIM"]);
            }
            catch (e) {
                verificadorDePassagem = false
                let query = [["Não", e.message + "(" + e.stack + ")"]]
                setValuesForRange(dados.rangePasso4.offset(0, 0, 1, 2), query);
                formataTable('NO_OK', dados.rangePasso4.offset(0, 0, 1, 2));
            }
        }
        else {
            repetidor = false
        }
    }
}
