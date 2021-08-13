//NOTE Tudo relacionado com utilitário do drive da google

const folderApuracao = DriveApp.getFolderById("1WNGCwdG2e6od3B7wxicciv4_mSdp1WQE"); //Pasta da Apuração em Automações
function getFolderByName(raiz, target) {//Função que através de uma pasta raiz, acha uma pasta procurada
    if (raiz != null) {
        var result = raiz.searchFolders('title = "' + target + '"');
        if (result.hasNext()) {
            return result.next();
        }
        else {
            return null;
        }
    }
    else {
        return null;
    }
}

function getFileByName(raiz, target) {//Função que através de uma pasta raiz, acha um arquivo procurado
    if (raiz != null) {
        var result = raiz.searchFiles('title = "' + target + '"');
        if (result.hasNext()) {
            return result.next();
        }
        else {
            return null;
        }
    }
    else {
        return null;
    }
}

function getFiles(raiz) {//Função que através de uma pasta raiz, devolve a lista de arquivos
    if (raiz != null) {
        var result = raiz.getFiles();
        if (result.hasNext()) {
            return result.next();
        }
        else {
            return null;
        }
    }
    else {
        return null;
    }
}

function convertXlsToSpreadSheet(file) {//converte arquivos excel para o tipo google sheet se necessário
    let fileBlob = file.getBlob();

    if (!fileBlob.isGoogleType()) {
        var newFile = {
            title: file.getName() + '_converted',
            mimeType: 'application/vnd.google-apps.spreadsheet' //  Added 
        };
        file = Drive.Files.insert(newFile, fileBlob, { convert: true });
    }

    return DriveApp.getFileById(file.id);
}

function deleteFileById(id) {//Deleta o arquivo achado por id
    DriveApp.getFileById(id).setTrashed(true);
}

function getfolderRaiz() {
    let folderFornecedor = getFolderByName(folderApuracao, information.cliente);
    let folderAno = getFolderByName(folderFornecedor, information.ano);
    let folderMes = getFolderByName(folderAno, information.mes);
    let folderApuracao = getFolderByName(folderMes);
    return folderApuracao;
}

function getFilesRequireds() {
    let folderApuracao = getfolderRaiz();

    let contaContabil = contaContabilExist(folderApuracao);
    let impostosARecolher = impostosARecolherexist(folderApuracao);

    if (contaContabil != null && impostosARecolher != null) {
        let contaContabilTemp = new FileAuxiliar(contaContabil, folderApuracao);
        let impostoARecolherTemp = new FileAuxiliar(impostosARecolher, folderApuracao);
        files.oks.push(contaContabilTemp, impostoARecolherTemp);
        files.status = true;
    }
    else {
        let fileTempNoOk = new FileAuxiliar(null, folderApuracao);
        files.noOk.push(fileTempNoOk);
        files.status = false;
    }
}

function contaContabilExist(folder) {
    if (folder != null) {
        let result = folder.searchFiles('title contains "Conta"');
        if (result.hasNext()) {
            return result.next();
        }
        else {
            return null;
        }
    }
    else {
        return null;
    }
}

function stfExist(folder) {
    if (folder != null) {
        let result = folder.searchFiles('title contains "STF"');
        if (result.hasNext()) {
            return result.next();
        }
        else {
            return null;
        }
    }
    else {
        return null;
    }
}
function impostosARecolherExist(folder) {
    if (folder != null) {
        let result = folder.searchFiles('title contains "Impostos"');
        if (result.hasNext()) {
            return result.next();
        }
        else {
            return null;
        }
    }
    else {
        return null;
    }
}