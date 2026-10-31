/**
 * CHÁ DO ISAAC — backend (Google Apps Script + Google Sheets)
 * A própria planilha é o banco de dados e o painel dos organizadores.
 * Abas: Itens (você edita), Convidados e Reservas (preenchidas pelo site).
 */
const CATS = ['Fraldas','Higiene do bebê','Roupinhas','Alimentação','Banho','Acessórios','Outros itens'];

function setup() { // Execute UMA vez, manualmente, para criar as abas.
  const ss = SpreadsheetApp.getActive();
  const defs = {
    Itens: ['id','nome','categoria','tamanho','quantidade_necessaria','ativo'],
    Convidados: ['token','nome_completo','telefone','presenca','quantidade_pessoas','mensagem','criado_em','atualizado_em'],
    Reservas: ['token','convidado','item_id','item','quantidade','criado_em','atualizado_em']
  };
  for (const n in defs) {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (s.getLastRow() === 0) s.appendRow(defs[n]);
    s.getRange(1,1,1,defs[n].length).setFontWeight('bold').setBackground('#dfe8c8');
    s.setFrozenRows(1);
  }
  ss.getSheetByName('Convidados').getRange('C:C').setNumberFormat('@');
  const it = ss.getSheetByName('Itens');
  if (it.getLastRow() < 2) {
    const seed = [
      ['Fralda tamanho P','Fraldas','Pacote',10],['Fralda tamanho M','Fraldas','Pacote',10],['Fralda tamanho G','Fraldas','Pacote',6],
      ['Lenço umedecido','Higiene do bebê','Pacote',8],['Pomada para assadura','Higiene do bebê','Tubo',4],['Algodão','Higiene do bebê','Pacote',4],
      ['Body manga curta','Roupinhas','Tamanho P',6],['Macacão','Roupinhas','Tamanho RN/P',5],['Meias e luvas','Roupinhas','Par',8],
      ['Mamadeira','Alimentação','',4],['Babador','Alimentação','',6],['Chupeta','Alimentação','',3],
      ['Sabonete líquido','Banho','',4],['Shampoo suave','Banho','',3],['Toalha com capuz','Banho','',3],
      ['Cobertor','Acessórios','',3],['Mordedor','Acessórios','',3],['Naninha','Acessórios','',2],
      ['Termômetro digital','Outros itens','',2],['Cartão-presente / Surpresa','Outros itens','',5]
    ];
    seed.forEach((r,i) => it.appendRow(['i'+(i+1), r[0], r[1], r[2], r[3], 'SIM']));
  }
}

const sh_ = n => SpreadsheetApp.getActive().getSheetByName(n);
function rows_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k,i) => [k, r[i]])));
}
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function state_(token) {
  const res = rows_('Reservas');
  const items = rows_('Itens').filter(i => String(i.ativo).toUpperCase() !== 'NÃO' && String(i.ativo).toUpperCase() !== 'NAO' && i.nome)
    .map(i => {
      const mine = res.filter(r => r.token === token && String(r.item_id) === String(i.id)).reduce((a,r) => a + Number(r.quantidade), 0);
      const tot = res.filter(r => String(r.item_id) === String(i.id)).reduce((a,r) => a + Number(r.quantidade), 0);
      return { id: String(i.id), nome: i.nome, categoria: i.categoria, tamanho: i.tamanho, necessaria: Number(i.quantidade_necessaria) || 0, reservada: tot, minha: mine };
    });
  let guest = null;
  if (token) {
    const g = rows_('Convidados').find(r => r.token === token);
    if (g) guest = { nome: g.nome_completo, telefone: String(g.telefone || ''), presenca: g.presenca, pessoas: g.quantidade_pessoas, mensagem: g.mensagem };
  }
  return { items, guest };
}

function doGet(e) { // Só devolve os dados do próprio token; nunca telefones/mensagens de terceiros.
  const t = String((e.parameter && e.parameter.t) || '');
  return out_(Object.assign({ ok: true }, state_(/^[0-9a-f-]{36}$/.test(t) ? t : '')));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000); // garante uma reserva por vez: verificação + gravação atômicas
    const d = JSON.parse(e.postData.contents);
    return out_(save_(d));
  } catch (err) {
    return out_({ ok: false, erro: 'Não foi possível concluir agora. Tente novamente.' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function save_(d) {
  if (d.website) return { ok: false, erro: 'Envio inválido.' }; // campo-isca contra robôs
  const token = String(d.token || '');
  if (!/^[0-9a-f-]{36}$/.test(token)) return { ok: false, erro: 'Identificação inválida. Recarregue a página.' };
  const nome = String(d.nome || '').trim().slice(0, 100);
  if (nome.length < 3) return { ok: false, erro: 'Informe seu nome completo.' };
  const sim = d.presenca === 'sim';
  if (!sim && d.presenca !== 'nao') return { ok: false, erro: 'Informe se você poderá comparecer.' };
  const pessoas = sim ? Math.min(20, Math.max(1, parseInt(d.pessoas) || 1)) : 0;
  const tel = String(d.telefone || '').slice(0, 30), msg = String(d.mensagem || '').slice(0, 500);

  const want = {};
  if (sim) (d.itens || []).forEach(x => { const q = parseInt(x.qtd); if (q > 0) want[String(x.id)] = (want[String(x.id)] || 0) + q; });
  if (sim && !Object.keys(want).length) return { ok: false, erro: 'Para confirmar presença, escolha ao menos um item para levar.' };

  const st = state_(token);
  for (const id in want) {
    const it = st.items.find(i => i.id === id);
    if (!it) return Object.assign({ ok: false, erro: 'Um dos itens não está mais na lista.' }, state_(token));
    const livre = it.necessaria - it.reservada + it.minha;
    if (want[id] > livre) return Object.assign({ ok: false,
      erro: livre <= 0 ? '"' + it.nome + '" acabou de ser reservado por outra pessoa.' : 'Só restam ' + livre + ' unidade(s) de "' + it.nome + '". Ajuste a quantidade.' }, state_(token));
  }

  const now = new Date(), cs = sh_('Convidados'), rs = sh_('Reservas');
  const toks = cs.getRange(1,1,cs.getLastRow(),1).getValues().map(r => r[0]);
  const idx = toks.indexOf(token);
  if (idx > 0) {
    const old = cs.getRange(idx+1,7).getValue();
    cs.getRange(idx+1,1,1,8).setValues([[token,nome,tel,sim?'Sim':'Não',pessoas,msg,old,now]]);
  } else cs.appendRow([token,nome,tel,sim?'Sim':'Não',pessoas,msg,now,now]);

  const rt = rs.getRange(1,1,rs.getLastRow(),1).getValues();
  for (let r = rt.length - 1; r >= 1; r--) if (rt[r][0] === token) rs.deleteRow(r + 1);
  Object.keys(want).forEach(id => {
    const it = st.items.find(i => i.id === id);
    rs.appendRow([token, nome, id, it.nome, want[id], now, now]);
  });
  SpreadsheetApp.flush();
  const fresh = state_(token);
  return Object.assign({ ok: true, token, presenca: sim, pessoas,
    reservas: fresh.items.filter(i => i.minha > 0).map(i => ({ nome: i.nome, tamanho: i.tamanho, qtd: i.minha })), data: now.toISOString() }, fresh);
}
