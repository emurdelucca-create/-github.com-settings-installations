const $ = id => document.getElementById(id);

async function carregarConfig() {
  const { webAppUrl, token } = await chrome.storage.local.get(['webAppUrl', 'token']);
  if (webAppUrl) $('webAppUrl').value = webAppUrl;
  if (token) $('token').value = token;
}

async function salvarConfig() {
  await chrome.storage.local.set({
    webAppUrl: $('webAppUrl').value.trim(),
    token: $('token').value.trim(),
  });
}

function log(msg, cls) {
  const el = $('status');
  el.textContent = msg;
  el.className = cls || '';
}

// O trabalho de verdade roda no background.js (service worker), que
// sobrevive ao popup fechar — o Chrome fecha o popup sozinho sempre que
// ele perde o foco, o que acontece justo quando abrimos a aba do
// concorrente em primeiro plano. O popup só dispara e depois fica lendo
// o status salvo em chrome.storage.local, então reabrir o popup no meio
// da captura mostra o progresso de onde ela realmente está.
function renderizarStatus(status) {
  const btn = $('capturar');
  if (!status) { log('Cadastre os concorrentes na planilha (coluna "Link Anúncio") antes de capturar.'); return; }

  btn.disabled = status.fase === 'rodando';
  log(status.mensagem || '', status.fase === 'ok' ? 'ok' : (status.fase === 'erro' ? 'err' : ''));

  if (status.debug) {
    $('debug').textContent = 'Debug (1º item, JSON bruto cortado):\n' + status.debug;
    $('debug').style.display = 'block';
  } else {
    $('debug').style.display = 'none';
  }
}

chrome.storage.onChanged.addListener(changes => {
  if (changes.espStatus) renderizarStatus(changes.espStatus.newValue);
});

$('salvar').addEventListener('click', async () => {
  await salvarConfig();
  log('✅ Configuração salva.', 'ok');
});

$('capturar').addEventListener('click', async () => {
  const webAppUrl = $('webAppUrl').value.trim();
  const token = $('token').value.trim();
  if (!webAppUrl || !token) {
    log('⚠ Preencha a URL do Web App e o Token antes de continuar.', 'err');
    return;
  }
  await salvarConfig();
  $('capturar').disabled = true;
  $('debug').style.display = 'none';
  log('▶ Iniciando... (pode fechar este popup, a captura continua sozinha — reabra pra ver o progresso)', '');
  chrome.runtime.sendMessage({ action: 'iniciarCaptura', webAppUrl, token });
});

(async function init() {
  await carregarConfig();
  const { espStatus } = await chrome.storage.local.get('espStatus');
  renderizarStatus(espStatus);
})();
