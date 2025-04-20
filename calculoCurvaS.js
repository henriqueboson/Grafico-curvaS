function calcularCurvaS() {
    var pastaProjetos = DriveApp.getFolderById('CODIGO-DA-PLANILHA');
    var arquivosProjetos = pastaProjetos.getFiles();
    var masterPlan = SpreadsheetApp.openById('CODIGO-DA-PLANILHA');
    
    var meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    var mesAtual = meses[new Date().getMonth()];
    var abaMasterPlan = masterPlan.getSheetByName(mesAtual);

    if (!abaMasterPlan) {
        Logger.log("Erro: Aba do mês '" + mesAtual + "' não encontrada no MasterPlan.");
        return;
    }

    // Pegar dados do Plano mensal
    var dadosMasterPlanH = abaMasterPlan.getRange('H:H').getValues().flat();
    var datasD = abaMasterPlan.getRange('D:D').getValues().flat();
    var datasF = abaMasterPlan.getRange('F:F').getValues().flat();
    var dataAtual = new Date();
    
    // Busca arquivos da pasta de projetos
    while (arquivosProjetos.hasNext()) {
        var arquivoProjeto = arquivosProjetos.next();
        var nomeArquivoDaVez = arquivoProjeto.getName();
        
        if (nomeArquivoDaVez.includes('___modelo')) continue;
        
        var planilhaProjeto = SpreadsheetApp.open(arquivoProjeto);
        var abaAux = planilhaProjeto.getSheetByName('aux');
        if (!abaAux) {
            Logger.log('Aba aux não encontrada no projeto: ' + nomeArquivoDaVez);
            continue;
        }
        
        // Limpa as colunas X, Y, Z
        abaAux.getRange('X9:Z').clearContent();
        
        var dadosPorSemana = {}; // Ex: { '2025-14': { total: 0, concluidas: 0 } }
        
        // Contabiliza atividades no MasterPlan para esse projeto
        for (var i = 0; i < dadosMasterPlanH.length; i++) {
            var nomeAtividadeH = dadosMasterPlanH[i];
            var dataD = datasD[i]; // Data final planejada
            var dataF = datasF[i]; // Data final real
            
            // Ignora se a linha estiver vazia ou não for texto
            if (!nomeAtividadeH || typeof nomeAtividadeH !== 'string') continue;
            
            // Ignora se não for desse projeto (nomeArquivoDaVez tem que estar no final)
            if (!nomeAtividadeH.trim().endsWith(nomeArquivoDaVez)) continue;
            
            // Ignora se a data final planejada for inválida
            if (!(dataD instanceof Date)) continue;
            
            var semanaAno = getAnoSemanaIso(dataD);
            
            // Se ainda não existe, inicializa a semana
            if (!dadosPorSemana[semanaAno]) {
                dadosPorSemana[semanaAno] = { total: 0, concluidas: 0 };
            }
            
            // Conta atividade total
            dadosPorSemana[semanaAno].total++;
            
            // Conta como concluída se tiver data final real válida
            if (dataF instanceof Date) {
                dadosPorSemana[semanaAno].concluidas++;
            }
        }
        
        // Ordenar as semanas em ordem crescente
        var semanasOrdenadas = Object.keys(dadosPorSemana).sort();
        
        // Preencher o aux (colunas X, Y, Z) de forma ACUMULADA
        var linha = 9;
        var acumuladoTotal = 0;
        var acumuladoConcluido = 0;
        
        var semanaAtual = getAnoSemanaIso(new Date()); // Pega a semana atual

        for (var s = 0; s < semanasOrdenadas.length; s++) {
            var semana = semanasOrdenadas[s];
            acumuladoTotal += dadosPorSemana[semana].total;
            acumuladoConcluido += dadosPorSemana[semana].concluidas;
        
            // Só plota concluídas até a semana atual
            if (semana <= semanaAtual) {
              abaAux.getRange('Z' + linha).setValue(acumuladoConcluido);   // Concluídas acumulado
            }
        
            abaAux.getRange('X' + linha).setValue(semana);               // Semana
            abaAux.getRange('Y' + linha).setValue(acumuladoTotal);       // Total acumulado
            linha++;
        }
        
    }
}

// Função para calcular semana ISO 8601
function getAnoSemanaIso(date) {
    var data = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    var diaSemana = data.getUTCDay() || 7;
    data.setUTCDate(data.getUTCDate() + 4 - diaSemana);
    var anoInicioSemana = data.getUTCFullYear();
    var primeiroDiaAno = new Date(Date.UTC(anoInicioSemana, 0, 1));
    var semana = Math.ceil((((data - primeiroDiaAno) / 86400000) + 1) / 7);
    return anoInicioSemana + '-' + (semana < 10 ? '0' + semana : semana);
}