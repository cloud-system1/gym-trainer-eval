// Use the URL of the deployment updated with the included Code.gs.
window.GYM_API_URL = 'https://script.google.com/macros/s/AKfycbwiafaHcPG-WgypZ1_9EgqZ-nLb34p9U3uV-g3w7ukrRlbJLR1Y_hK87mdJMoPiDyU/exec';
window.gymRequest = async function(action, payload) {
  const url = new URL(window.GYM_API_URL);
  url.searchParams.set('t', Date.now());
  if (action) url.searchParams.set('action', action);
  const res = await fetch(url.toString(), payload ? {method:'POST', body:JSON.stringify(payload)} : {cache:'no-store'});
  const raw = await res.text();
  if (!res.ok) throw new Error('HTTP ' + res.status + ' — check web app access');
  let result;
  try { result = JSON.parse(raw); } catch (_) { throw new Error('Google returned a non-JSON page. Check deployment URL and Anyone access.'); }
  if (!result.ok) throw new Error(result.error || 'Backend request failed');
  return result;
};
window.gymHealth = async function() {
  const result = await window.gymRequest('health');
  if (result.version !== 'backup-v3') throw new Error('Outdated Apps Script deployment. Deploy the included Code.gs as a new version.');
  return result;
};
