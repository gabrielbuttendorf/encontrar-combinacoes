//Dados da planilha
const planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
const ultimaLinha = planilha.getLastRow()

//Pegar valores da planilha
var valorSoma = planilha.getRange('B2').getValue()
const colunaValores = planilha.getRange('A2:A' + ultimaLinha).getValues()

const valores = []
for (let i = 0; i < colunaValores.length; i++) {
  valores.push(colunaValores[i][0])
}

function encontrarCombinacoes(colunaValores, soma) {
  let combinacaoEncontrada = false
  const combinacoes = []
  const EPSILON = 0.0001
  const negativo = planilha.getRange('G1').getValue()

  function backtrack(indiceInicial, somaAtual, combinacaoAtual) {
    somaAtual = somaAtual - 0.000000001; //Evitar que não some caso ultrapasse.
    if (combinacaoEncontrada) {
      return
    }

    if (Math.abs(somaAtual - soma) < EPSILON) {
      combinacoes.push([...combinacaoAtual]);
      combinacaoEncontrada = true
      return
    }

    //Negativos e positivos
    if (negativo === true) {
      if (colunaValores == null) {
        return
      } else {
        for (let i = indiceInicial; i < colunaValores.length; i++) {
            combinacaoAtual.push(colunaValores[i])
            backtrack(i + 1, somaAtual + colunaValores[i], combinacaoAtual)
            combinacaoAtual.pop()
        }
      }
    //Somente positivos
    } else {
      if (colunaValores == null) {
        return
      } else {
        for (let i = indiceInicial; i < colunaValores.length; i++) {
          if (somaAtual + colunaValores[i] <= soma) {
            combinacaoAtual.push(colunaValores[i])
            backtrack(i + 1, somaAtual + colunaValores[i], combinacaoAtual)
            combinacaoAtual.pop()
          }
        }
      }
    }
  }

  backtrack(0, 0, [])
  return combinacoes
}

function exibir() {
  const soma = valorSoma
  const combinacoes = encontrarCombinacoes(valores, soma)
  if (combinacoes[0] == null || soma == '') {
    return 'NADA ENCONTRADO'
  } else {
    return combinacoes[0]
  }
}

function modificacaoPlanilha() {
  const celulaFormula = planilha.getRange('D2')
  
  //Limpar D2 para baixo
  const numRowsToClear = planilha.getLastRow() - 1
  planilha.getRange(celulaFormula.getRow(), celulaFormula.getColumn(), numRowsToClear, 1).clearContent()
  
  const resultado = exibir()
  const arrayResultado = []

  if (resultado === 'NADA ENCONTRADO') {
    celulaFormula.setValue(resultado)
  } else {
    for (let i = 0; i < resultado.length; i++) {
      arrayResultado.push([resultado[i]])
    }

    const celulasResultado = celulaFormula.offset(0, 0, arrayResultado.length, arrayResultado[0].length)
    celulasResultado.setValues(arrayResultado)
  }
}