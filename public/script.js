let resultado = [];
const downloadButton = document.querySelector('#downloadButton');
const login = document.querySelector('.login');
const detDataDiv = document.querySelector('.getData');
const barInBar = document.querySelector('.barInBar');
const progressText = document.querySelector('.progressText');
let orders;
let shipment;
let resultFinal;
let dateStart;
let dateEnd;
let percentBar = 0;

let access_token = null;

document.getElementById('btnConsultar').addEventListener('click', async (event) => {
  event.preventDefault();
  document.getElementById('btnConsultar').disabled = true;
  await consultar();
  document.getElementById('btnConsultar').disabled = false;
});

document.getElementById('downloadButton').addEventListener('click', async (event) => {
  event.preventDefault();
  await exportarParaExcel();
});

// Botão de login
document.getElementById('btnLogin').addEventListener('click', async () => {
  try {
    const response = await fetch('/auth/url');
    const { authUrl } = await response.json();

    if (!authUrl) {
      alert('Erro ao obter URL de autenticação');
      return;
    }

    window.open(authUrl, '_blank', 'width=600,height=700');
  } catch (err) {
    alert('Erro ao iniciar login: ' + err.message);
  }
});

document.getElementById('authCode').addEventListener('input', () => {
  if (document.getElementById('authCode').value.trim() !== "") {
    document.getElementById('btnGetToken').disabled = false;
  } else {
    document.getElementById('btnGetToken').disabled = true;
  }
})

// Obter access token
document.getElementById('btnGetToken').addEventListener('click', async () => {
  const code = document.getElementById('authCode').value.trim();
  if (!code) {
    alert('Informe o código de autorização');
    return;
  }

  try {
    const response = await fetch('/auth/token', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ code })
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || 'Erro ao obter token');
    }

    access_token = data.access_token;

    login.style.display = "none";
    detDataDiv.style.display = "flex";
  } catch (err) {
    alert(`Erro ao obter token: Code inválido ou já usado. Crie um code novo e tente novamente.`);
  }
});

function validarData(valueInput) {
  if (!valueInput) {
    alert("Por favor, preencha a data corretamente.");
    return false;
  }

  const data = new Date(valueInput);
  const dataAtual = new Date();

  if (isNaN(data.getTime())) {
    alert("Data inválida: Formato errado. preencha os campos de data corretamente.");
    return false;
  }

  if (data > dataAtual) {
    alert("Data inválida: Você não pode selecionar uma data posterior à data de hoje.");
    return false;
  }

  return data.getTime();
}


// Consultar pedido e nota
async function consultar() {
  resultado = [];
  dateStart = document.querySelector('#data-start-rel').value;
  dateEnd = document.querySelector('#data-end-rel').value;
  const stateDiv = document.querySelector('#state-div');
  let percentBar = 0;

  downloadButton.disabled = true;

  if (!validarData(dateStart) || !validarData(dateEnd)) {
    return; 
  } else if (validarData(dateEnd) < validarData(dateStart)) {
    alert("Data inválida: A data final não pode ser menor que a inicial. Coloque a mesma ou uma maior.");
    return;
  }

  if (!access_token) {
    alert('Faça o login e obtenha o token primeiro.');
    return;
  }

  stateDiv.textContent = 'Status: Gerando relatório';

  // liberações
  try {
    let restanteLib;
    let liberacoes = [];
    let offset = 0;
    const limit = 500;

    while (true) {
      progressText.textContent = `Buscando ordens`;
      const response = await fetch(`/liberacoes/${dateStart}/${dateEnd}/${offset}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ token: access_token })
      });

      const data = await response.json();
      const pagina = data.pedido.results || [];

      liberacoes.push(...pagina);

      if (pagina.length < limit) break; // Chegou na última página

      offset += limit;
    }
    orders = liberacoes;
    progressText.textContent = `Filtrando ordens`;
    const liberacoesFilter = liberacoes.filter(item => item.description !== 'marketplace_shipment' && !item.external_reference.startsWith("cashback_"));

    const quantLib = liberacoesFilter.length;
    const percentLib = (100 / quantLib) / 2;

    restanteLib = quantLib;

    for (const lib of liberacoesFilter) {
      progressText.textContent = `Buscando ordem: ${lib.order.id} (${percentBar.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}% | ${restanteLib} restantes)`;

      try {
        const response = await fetch(`/pedido/${lib.order.id}`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ token: access_token })
        });

        const data = await response.json();

        const dateFormated = moment(lib.money_release_date).tz('America/Sao_Paulo').format().slice(0, 10);
        const [ano, mes, dia] = dateFormated.split('-');
        const dataBR = `${dia}/${mes}/${ano}`;

        let libFrete = liberacoes.filter(item => item.description === 'marketplace_shipment' && data.shipping?.id === Number(item.external_reference));
        if (libFrete.length === 0) {
          libFrete = 0
        } else {
          libFrete = libFrete[0].transaction_amount;
        }

        resultado.push({
          date: dataBR,
          order: lib.order.id,
          idEnvio: data.shipping?.id,
          cliente: data.buyer.first_name,
          nfe: 0,
          precoProduto: Math.round(lib.transaction_amount * 100) / 100,
          taxa: Math.round((lib.transaction_amount - (lib.transaction_details.net_received_amount + libFrete)) * 100) / 100,
          valorLiq: Math.round((lib.transaction_details.net_received_amount + libFrete) * 100) / 100
        });

        percentBar = percentBar + percentLib;
        barInBar.style.width = `${percentBar}%`
        restanteLib = restanteLib - 1;
      } catch (e) {
        console.error('Erro ao consultar pedido:', e.message);
      }
    };

    restanteLib = quantLib;

    for (const res of resultado) {
      progressText.textContent = `Buscando NF-e da ordem: ${res.order} (${percentBar.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}% | ${restanteLib} restantes)`;

      try {
        const response = await fetch(`/nfe/${Number(res.order)}`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ token: access_token })
        });

        const data = await response.json();

        res.nfe = data.invoice_number;

        percentBar = percentBar + percentLib;
        barInBar.style.width = `${percentBar}%`
        restanteLib = restanteLib - 1;
      } catch (e) {
        console.error('Erro ao consultar pedido:', e.message);
      }
    }
  } catch (error) {
    console.error("Erro ao buscar liberações:", error);
  }
  const mapa = new Map();

  for (const item of resultado) {
    const chave = item.idEnvio;

    if (!mapa.has(chave)) {
      mapa.set(chave, {
        ...item,
        precoProduto: Number(item.precoProduto) || 0,
        valorLiq: Number(item.valorLiq) || 0,
        taxa: Number(item.taxa) || 0
      });
    } else {
      const acumulado = mapa.get(chave);
      acumulado.precoProduto += Number(item.precoProduto) || 0;
      acumulado.valorLiq += Number(item.valorLiq) || 0;
      acumulado.taxa += Number(item.taxa) || 0;
    }
  }
  resultFinal = Array.from(mapa.values());
  stateDiv.textContent = 'Status: Relatório gerado e disponível para download.';
  progressText.textContent = `Concluído`;
  downloadButton.disabled = false;
  return resultFinal;
}

async function exportarParaExcel() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Dados');

  const headerPlan = ["Data", "Nº Ordem", "ID de Envio", "Nome do Cliente", "NF-e", "Valor Bruto", "Taxa", "Valor Liquido"];

  // Adiciona o cabeçalho
  worksheet.addRow(headerPlan);

  // Adiciona as linhas de dados
  resultFinal.forEach(item => {
    worksheet.addRow([
      item.date,
      item.order,
      item.idEnvio,
      item.cliente,
      item.nfe,
      item.precoProduto,
      item.taxa,
      item.valorLiq
    ]);
  });

  // Formatar cabeçalho (linha 1) - cor, negrito, borda e centralizar
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{ argb:'BDD7EE' }
    };
    cell.font = { bold: true };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: {style:'thin'},
      left: {style:'thin'},
      bottom: {style:'thin'},
      right: {style:'thin'}
    };
  });

  // Centralizar toda a coluna E (NF-e)
  worksheet.getColumn(2).alignment = { horizontal: 'center' };
  worksheet.getColumn(3).alignment = { horizontal: 'center' };
  worksheet.getColumn(5).alignment = { horizontal: 'center' };
  worksheet.getColumn(6).alignment = { horizontal: 'center' };
  worksheet.getColumn(7).alignment = { horizontal: 'center' };
  worksheet.getColumn(8).alignment = { horizontal: 'center' };

  // Formatar colunas F, G e H como moeda brasileira (pt-BR)
  const formatoMoeda = 'R$ #,##0.00;[Red]-R$ #,##0.00';
  [6,7,8].forEach(colNumber => {
    worksheet.getColumn(colNumber).numFmt = formatoMoeda;
  });

  // Aplicar bordas finas em todas as células da planilha (exceto cabeçalho que já tem)
  worksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {
    if (rowNumber === 1) return; // pular cabeçalho
    row.eachCell((cell) => {
      cell.border = {
        top: {style:'thin'},
        left: {style:'thin'},
        bottom: {style:'thin'},
        right: {style:'thin'}
      };
    });
  });

  // Ajusta a largura das colunas automaticamente baseado no conteúdo
  worksheet.columns.forEach(column => {
    let maxLength = 10;
    column.eachCell({ includeEmpty: false }, cell => {
      const length = cell.value ? cell.value.toString().length : 0;
      if (length > maxLength) maxLength = length;
    });
    column.width = maxLength + 2;
  });

  // Gera o arquivo e baixa (browser)
  const buf = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  saveAs(blob, `dados-${dateStart}-a-${dateEnd}.xlsx`);
}