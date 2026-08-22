/* NegosyoSheet — unlock gate. OWNER: set codes + real full-file URL (see PAYMENTS.md). */
'use strict';
const PRO_CODES = ['NEGOSYO-149', 'NS-DEMO'];
const FULL_URL = './full/NegosyoSheet.xlsx'; // set to release URL or keep path (see PAYMENTS.md)
const LS_KEY = 'ns_pro';

document.addEventListener('DOMContentLoaded', () => {
  const $ = (id) => document.getElementById(id);
  if (localStorage.getItem(LS_KEY) === '1') revealDownload();

  const openPay = () => { $('payModal').classList.remove('hidden'); $('codeMsg').textContent = ''; };
  $('proBtn').addEventListener('click', openPay);
  $('proBtn2').addEventListener('click', openPay);
  $('payClose').addEventListener('click', () => $('payModal').classList.add('hidden'));
  $('payModal').addEventListener('click', e => { if (e.target === e.currentTarget) e.currentTarget.classList.add('hidden'); });

  const tryCode = () => {
    const code = $('codeInput').value.trim().toUpperCase();
    if (PRO_CODES.map(c => c.toUpperCase()).includes(code)) {
      localStorage.setItem(LS_KEY, '1');
      $('codeMsg').textContent = '✓ Valid! I-download na ang full version.';
      $('codeMsg').className = 'code-msg ok';
      revealDownload();
    } else {
      $('codeMsg').textContent = 'Mali ang code — check ang confirmation ng bayad.';
      $('codeMsg').className = 'code-msg bad';
    }
  };
  $('codeBtn').addEventListener('click', tryCode);
  $('codeInput').addEventListener('keydown', e => { if (e.key === 'Enter') tryCode(); });

  function revealDownload() {
    $('dlLink').href = FULL_URL;
    $('dlBox').classList.remove('hidden');
  }
});
