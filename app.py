from __future__ import annotations

import cgi
import json
import socket
import sys
import threading
import uuid
import warnings
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pandas as pd

from core import analyze_balances, dataframe_to_excel, ledger_file_diagnostics, read_accounting_file, sorted_report


warnings.filterwarnings("ignore", message="'cgi' is deprecated.*", category=DeprecationWarning)

HOST = "127.0.0.1"
PORT = 8505
ANALYSES: dict[str, pd.DataFrame] = {}
LOGO_PATH = Path(__file__).with_name("logo_analisador_contabil.svg")


HTML = r"""<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Analisador Contábil</title>
  <link rel="icon" href="/logo" type="image/svg+xml">
  <style>
    :root {
      color-scheme: dark;
      --bg: #0B1220;
      --sidebar: #0B1220;
      --panel: #131D32;
      --panel-2: #17243B;
      --line: #24324A;
      --text: #F8FAFC;
      --muted: #94A3B8;
      --blue: #3B82F6;
      --cyan: #06B6D4;
      --green: #22C55E;
      --red: #EF4444;
      --purple: #8B5CF6;
      --yellow: #F59E0B;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background:
        radial-gradient(900px 520px at 18% 0%, rgba(59, 130, 246, .13), transparent 60%),
        radial-gradient(850px 520px at 86% 6%, rgba(6, 182, 212, .10), transparent 58%),
        var(--bg);
      color: var(--text);
      font-family: "Segoe UI", Arial, sans-serif;
      letter-spacing: 0;
    }
    .app { display: grid; grid-template-columns: 280px minmax(0, 1fr); min-height: 100vh; }
    aside {
      position: sticky;
      top: 0;
      height: 100vh;
      padding: 30px 16px 22px;
      background: linear-gradient(180deg, var(--sidebar), #091221);
      border-right: 1px solid var(--line);
    }
    main { padding: 40px 42px 48px; min-width: 0; max-width: 1360px; }
    .brand { display: flex; align-items: center; gap: 14px; margin: 0 0 34px; }
    .brand img {
      width: 52px; height: 52px; flex: 0 0 auto;
      filter: drop-shadow(0 12px 22px rgba(59,130,246,.20));
    }
    .brand strong { display: block; font-size: 21px; line-height: 1.08; letter-spacing: -.02em; }
    .brand span { display: block; margin-top: 6px; color: var(--muted); font-size: 12px; }
    .menu-icon {
      width: 18px; height: 18px; display: inline-grid; place-items: center; color: #B8C7DD;
    }
    .menu-icon svg { width: 18px; height: 18px; stroke: currentColor; fill: none; stroke-width: 2; }
    .menu-item {
      display: flex; align-items: center; gap: 12px;
      padding: 14px; margin: 7px 0; border-radius: 8px;
      color: #c8d2e5; font-weight: 650;
      cursor: pointer;
      border: 0;
      width: 100%;
      background: transparent;
      text-align: left;
    }
    .menu-item.active {
      background: rgba(47, 140, 255, .13);
      color: var(--text);
      box-shadow: inset 0 0 0 1px rgba(148, 163, 184, .13);
    }
    .side-card {
      position: absolute; left: 16px; right: 16px; bottom: 22px;
      padding: 16px; border: 1px solid var(--line); border-radius: 8px;
      background: linear-gradient(180deg, rgba(17, 29, 48, .86), rgba(10, 18, 31, .92));
    }
    .side-card small { display: block; color: var(--muted); margin-bottom: 9px; }
    .side-card strong { display: block; font-size: 14px; word-break: break-word; }
    .side-status {
      display: inline-flex; margin-top: 12px; padding: 7px 10px; border-radius: 7px;
      color: var(--green); background: rgba(16, 232, 138, .12);
      font-size: 12px; font-weight: 800;
    }
    .side-credit {
      margin-top: 16px;
      padding-top: 14px;
      border-top: 1px solid rgba(148, 163, 184, .14);
      color: var(--muted);
      font-size: 11px;
      line-height: 1.45;
    }
    .side-credit strong {
      display: block;
      color: #DDE6F5;
      font-size: 12px;
      margin-top: 3px;
    }
    .side-credit a {
      display: inline-block;
      color: var(--cyan);
      text-decoration: none;
      margin-top: 4px;
    }
    .side-credit a:hover { color: var(--green); }
    .title-row { display: flex; justify-content: space-between; align-items: flex-start; gap: 24px; margin-bottom: 30px; }
    h1 { margin: 0; font-size: 46px; line-height: 1.04; letter-spacing: -.035em; }
    .subtitle { margin: 12px 0 0; color: var(--muted); font-size: 17px; }
    .new-file {
      display: inline-flex; align-items: center; justify-content: center; gap: 9px;
      min-width: 190px; min-height: 54px; border-radius: 8px; border: 0;
      background: linear-gradient(135deg, #0abf6b, var(--green));
      color: #03120b; font-weight: 900; cursor: pointer;
      box-shadow: 0 16px 34px rgba(16, 232, 138, .18);
    }
    .upload-panel, .section-panel, .metric, .upload-card, .upload-aside {
      border: 1px solid var(--line);
      border-radius: 16px;
      background: linear-gradient(180deg, rgba(19, 29, 50, .96), rgba(11, 18, 32, .98));
      box-shadow: 0 24px 70px rgba(0, 0, 0, .24);
    }
    .upload-panel {
      padding: 0;
      border: 0;
      background: transparent;
      box-shadow: none;
      max-width: 1180px;
    }
    .upload-shell { display: grid; grid-template-columns: minmax(0, 1.35fr) minmax(300px, .65fr); gap: 20px; align-items: stretch; }
    .upload-card { padding: 30px; }
    .upload-aside { padding: 26px; position: relative; overflow: hidden; }
    .upload-aside::after {
      content: "";
      position: absolute;
      width: 220px; height: 220px; right: -90px; top: -90px;
      background: radial-gradient(circle, rgba(6,182,212,.24), transparent 62%);
    }
    .upload-kicker {
      display: inline-flex; align-items: center; gap: 8px;
      padding: 7px 10px; border: 1px solid rgba(59,130,246,.32); border-radius: 999px;
      color: #BFDBFE; background: rgba(59,130,246,.10); font-size: 12px; font-weight: 800;
      margin-bottom: 18px;
    }
    .upload-title { font-size: 31px; font-weight: 950; margin-bottom: 8px; letter-spacing: -.03em; }
    .upload-subtitle { color: var(--muted); margin-bottom: 24px; font-size: 15px; line-height: 1.6; max-width: 720px; }
    .upload-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 18px; }
    .file-drop {
      display: block;
      min-height: 178px;
      padding: 20px;
      border: 1px dashed rgba(148,163,184,.36);
      border-radius: 16px;
      background:
        linear-gradient(180deg, rgba(59,130,246,.06), rgba(6,182,212,.03)),
        rgba(11,18,32,.62);
      transition: border-color .16s ease, transform .16s ease, background .16s ease;
    }
    .file-drop:hover { border-color: rgba(59,130,246,.78); transform: translateY(-2px); background: rgba(59,130,246,.08); }
    .file-icon {
      width: 44px; height: 44px; border-radius: 12px;
      display: grid; place-items: center;
      color: #DBEAFE; background: rgba(59,130,246,.18);
      margin-bottom: 16px; font-size: 20px;
    }
    .file-drop strong { display: block; font-size: 16px; margin-bottom: 6px; }
    .file-drop span.text { display: block; color: var(--muted); font-size: 13px; line-height: 1.45; margin-bottom: 16px; }
    label { display: block; color: #dbe5f6; font-weight: 750; margin-bottom: 8px; }
    input, select, button {
      font: inherit;
      color: var(--text);
    }
    input[type="search"], select {
      width: 100%; min-height: 48px; padding: 10px 12px;
      border: 1px dashed var(--line); border-radius: 14px;
      background: rgba(6, 11, 20, .48);
    }
    input[type="file"] {
      width: 100%;
      color: var(--muted);
      font-size: 12px;
    }
    input::file-selector-button {
      border: 0; border-radius: 7px; padding: 8px 12px; margin-right: 10px;
      color: #dff7ff; background: rgba(47, 140, 255, .18);
      cursor: pointer;
    }
    .analyze {
      width: 100%; min-height: 56px; border: 0; border-radius: 12px;
      background: linear-gradient(135deg, var(--blue), #2563eb);
      color: white; font-weight: 900; cursor: pointer;
      box-shadow: 0 18px 42px rgba(59,130,246,.25);
    }
    .upload-points { display: grid; gap: 14px; margin-top: 22px; position: relative; z-index: 1; }
    .upload-point { display: grid; grid-template-columns: 32px 1fr; gap: 12px; align-items: start; }
    .upload-point b {
      width: 32px; height: 32px; border-radius: 10px;
      display: grid; place-items: center;
      background: rgba(34,197,94,.12); color: var(--green);
      border: 1px solid rgba(34,197,94,.25);
    }
    .upload-point strong { display: block; font-size: 14px; margin-bottom: 3px; }
    .upload-point span { color: var(--muted); font-size: 12px; line-height: 1.45; }
    .requirements-list { display: grid; gap: 14px; margin-top: 18px; position: relative; z-index: 1; }
    .requirement-card {
      border: 1px solid rgba(148,163,184,.18);
      border-radius: 12px;
      padding: 14px;
      background: rgba(11,18,32,.56);
    }
    .requirement-card strong { display: block; font-size: 14px; margin-bottom: 8px; }
    .requirement-card span { display: block; color: var(--muted); font-size: 12px; line-height: 1.48; }
    .about-details {
      margin-top: 16px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.55;
      border-top: 1px solid rgba(148,163,184,.13);
      padding-top: 14px;
    }
    .about-details summary { color: #DDE6F5; cursor: pointer; font-weight: 800; }
    .aside-title { font-size: 20px; font-weight: 900; letter-spacing: -.02em; position: relative; z-index: 1; }
    .aside-subtitle { color: var(--muted); margin-top: 8px; line-height: 1.55; position: relative; z-index: 1; }
    .hint { color: var(--muted); font-size: 13px; margin-top: 10px; }
    .status { color: var(--muted); margin-top: 12px; }
    .error { color: var(--yellow); font-weight: 800; }
    .metrics { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 20px; margin-bottom: 26px; }
    .metric { padding: 22px; min-height: 148px; transition: transform .16s ease, border-color .16s ease; }
    .metric:hover { transform: translateY(-2px); border-color: rgba(59, 130, 246, .55); }
    .metric span { display: block; color: #cbd5e1; font-size: 14px; margin-bottom: 16px; }
    .metric strong { display: block; font-size: 32px; line-height: 1.05; font-weight: 760; letter-spacing: -.03em; }
    .metric em { display: block; color: var(--muted); font-style: normal; font-size: 13px; margin-top: 14px; }
    .metric.period strong {
      color: #DDEAFE !important;
      font-size: 20px !important;
      line-height: 1.25;
      font-weight: 720;
      letter-spacing: -.015em;
    }
    .metric.period em { color: var(--cyan); }
    .section-panel { padding: 0; overflow: hidden; }
    .section-head {
      display: flex; justify-content: space-between; align-items: center; gap: 18px;
      padding: 26px 24px 22px;
    }
    .section-head h2 { margin: 0 0 6px; font-size: 27px; }
    .section-head p { margin: 0; color: var(--muted); }
    .tools { display: grid; grid-template-columns: minmax(260px, 360px) 130px 150px; gap: 12px; align-items: end; }
    .filter-btn, .export-btn {
      display: inline-flex; align-items: center; justify-content: center;
      min-height: 48px; border: 1px solid var(--line); border-radius: 8px;
      background: rgba(17, 29, 48, .9); color: var(--text); text-decoration: none; font-weight: 850;
      cursor: pointer;
    }
    .table { width: 100%; overflow-x: auto; }
    .table-grid { min-width: 1220px; }
    .table-head, .table-row {
      display: grid;
      grid-template-columns: 115px 130px minmax(270px, 1.8fr) 135px 125px 120px 130px 105px 86px;
      gap: 12px; align-items: center;
    }
    .table-head {
      padding: 16px 24px; font-size: 13px; font-weight: 900; color: #eef4ff;
      border-top: 1px solid rgba(148, 163, 184, .10);
      border-bottom: 1px solid var(--line);
      background: rgba(17, 29, 48, .45);
    }
    .table-row {
      padding: 14px 24px; font-size: 13px; color: #dde6f5;
      border-bottom: 1px solid rgba(148, 163, 184, .12);
    }
    .table-row:hover { background: rgba(47, 140, 255, .05); }
    .desc { overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
    .badge {
      display: inline-flex; align-items: center; justify-content: center;
      padding: 6px 10px; border-radius: 7px; font-size: 12px; font-weight: 850; white-space: nowrap;
    }
    .fornecedor { color: #e9d5ff; background: rgba(181, 108, 255, .18); }
    .cliente { color: #bdfcee; background: rgba(32, 217, 194, .16); }
    .conta { color: #bfdbfe; background: rgba(47, 140, 255, .16); }
    .credor { color: var(--green); background: rgba(16, 232, 138, .10); }
    .devedor { color: var(--red); background: rgba(255, 92, 102, .10); }
    .action-btn {
      min-height: 38px; border: 1px solid var(--line); border-radius: 8px;
      background: rgba(6, 11, 20, .55); color: #dbeafe; cursor: pointer;
    }
    .modal-backdrop {
      display: none; position: fixed; inset: 0; z-index: 20;
      background: rgba(2, 6, 23, .68); backdrop-filter: blur(10px);
      align-items: stretch; justify-content: flex-end;
    }
    .modal-backdrop.open { display: flex; }
    .drawer {
      width: min(620px, 100vw); min-height: 100vh; padding: 26px;
      border-left: 1px solid var(--line);
      background: linear-gradient(180deg, #131D32, #0B1220);
      box-shadow: -30px 0 70px rgba(0,0,0,.42);
      overflow-y: auto;
    }
    .drawer-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 18px; margin-bottom: 20px; }
    .drawer h3 { margin: 0 0 8px; font-size: 24px; line-height: 1.2; }
    .drawer p { margin: 0; color: var(--muted); }
    .close-btn {
      width: 42px; min-height: 42px; border-radius: 10px; border: 1px solid var(--line);
      background: rgba(11,18,32,.78); color: var(--text); cursor: pointer;
    }
    .detail-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin: 18px 0; }
    .detail-card {
      padding: 13px 14px; border: 1px solid var(--line); border-radius: 10px;
      background: rgba(11,18,32,.62);
    }
    .detail-card span { display: block; color: var(--muted); font-size: 12px; margin-bottom: 6px; }
    .detail-card strong { font-size: 14px; }
    .chart-card, .involved-card {
      margin-top: 14px; padding: 16px; border: 1px solid var(--line); border-radius: 12px;
      background: rgba(11,18,32,.62);
    }
    .chart-card h4, .involved-card h4 { margin: 0 0 12px; font-size: 15px; }
    .mini-chart { width: 100%; height: 180px; }
    .mini-table { width: 100%; border-collapse: collapse; font-size: 12px; }
    .mini-table th, .mini-table td { padding: 9px 6px; border-bottom: 1px solid rgba(148,163,184,.14); text-align: left; }
    .mini-table th { color: var(--muted); font-weight: 800; }
    .empty { padding: 34px; text-align: center; color: var(--muted); }
    .count { color: var(--muted); padding: 16px 24px 24px; font-size: 13px; }
    .info-content { padding: 26px 24px; }
    .info-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; margin-top: 16px; }
    .info-card {
      border: 1px solid rgba(148,163,184,.16);
      border-radius: 12px;
      padding: 16px;
      background: rgba(11,18,32,.48);
    }
    .info-card strong { display: block; margin-bottom: 8px; }
    .info-card span { display: block; color: var(--muted); font-size: 13px; line-height: 1.5; }
    .hidden { display: none; }
    @media (max-width: 1050px) {
      .app { grid-template-columns: 1fr; }
      aside { position: relative; height: auto; }
      .side-card { position: static; margin-top: 22px; }
      main { padding: 24px 18px 38px; }
      .title-row, .section-head { flex-direction: column; align-items: stretch; }
      .metrics, .upload-grid, .tools, .upload-shell, .info-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div class="app">
    <aside>
      <div class="brand">
        <img src="/logo" alt="Analisador Contabil">
        <div><strong>Analisador<br>Contábil</strong><span>Conferência automática</span></div>
      </div>
      <nav>
        <button class="menu-item active" data-page="Resumo"><span class="menu-icon"><svg viewBox="0 0 24 24"><rect x="4" y="4" width="6" height="6" rx="1"/><rect x="14" y="4" width="6" height="6" rx="1"/><rect x="4" y="14" width="6" height="6" rx="1"/><rect x="14" y="14" width="6" height="6" rx="1"/></svg></span>Resumo</button>
        <button class="menu-item" data-page="Inconsistências"><span class="menu-icon"><svg viewBox="0 0 24 24"><path d="M12 4l9 16H3L12 4z"/><path d="M12 9v5"/><path d="M12 18h.01"/></svg></span>Inconsistências</button>
        <button class="menu-item" data-page="Contas"><span class="menu-icon"><svg viewBox="0 0 24 24"><path d="M7 3h10a2 2 0 0 1 2 2v16H5V5a2 2 0 0 1 2-2z"/><path d="M8 8h8"/><path d="M8 12h8"/><path d="M8 16h5"/></svg></span>Contas</button>
        <button class="menu-item" data-page="Participantes"><span class="menu-icon"><svg viewBox="0 0 24 24"><path d="M16 11a4 4 0 1 0-8 0"/><path d="M4 21a8 8 0 0 1 16 0"/><path d="M18 8a3 3 0 0 1 3 3"/><path d="M22 21a6 6 0 0 0-4-5.7"/></svg></span>Participantes</button>
        <button class="menu-item" data-page="Configurações"><span class="menu-icon"><svg viewBox="0 0 24 24"><path d="M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8z"/><path d="M3 12h3"/><path d="M18 12h3"/><path d="M12 3v3"/><path d="M12 18v3"/><path d="M5.6 5.6l2.1 2.1"/><path d="M16.3 16.3l2.1 2.1"/><path d="M18.4 5.6l-2.1 2.1"/><path d="M7.7 16.3l-2.1 2.1"/></svg></span>Configurações</button>
        <button class="menu-item" data-page="Sobre"><span class="menu-icon"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 11v6"/><path d="M12 7h.01"/></svg></span>Sobre</button>
      </nav>
      <div class="side-card">
        <small>Arquivo analisado</small>
        <strong id="sideFile">Nenhum arquivo analisado</strong>
        <small id="sideTime" style="margin-top:10px">-</small>
        <span id="sideStatus" class="side-status">Aguardando arquivos</span>
        <div class="side-credit">
          Desenvolvido por
          <strong>Rayssa Mayara</strong>
          <a href="https://github.com/rayssamayarax" target="_blank" rel="noopener">github.com/rayssamayarax</a>
        </div>
      </div>
    </aside>
    <main>
      <div class="title-row">
        <div>
          <h1 id="pageTitle">Analisador Contábil</h1>
          <p id="pageSubtitle" class="subtitle">Conferência automática de saldos no razão SCI.</p>
        </div>
        <button id="newFileBtn" class="new-file hidden" type="button">⇧ Enviar novo arquivo</button>
      </div>

      <section id="uploadPanel" class="upload-panel">
        <div class="upload-shell">
          <div class="upload-card">
            <div class="upload-kicker">● Importação segura</div>
            <div class="upload-title">Envie os arquivos da análise</div>
            <div class="upload-subtitle">Formatos aceitos: CSV, XLSX e XLS.</div>
            <form id="uploadForm">
              <div class="upload-grid">
                <label class="file-drop" for="plan">
                  <div class="file-icon">▤</div>
                  <strong>Plano de contas</strong>
                  <span class="text">Código obrigatório. Nome, classificação, grupo, relatório, tipo e saldo/natureza são recomendados.</span>
                  <input id="plan" name="plan" type="file" accept=".csv,.xlsx,.xls" required>
                </label>
                <label class="file-drop" for="ledger">
                  <div class="file-icon">↻</div>
                  <strong>Razão SCI</strong>
                  <span class="text">Histórico e Saldo obrigatórios. Aceita Débito/Crédito ou Valor no arquivo exportado.</span>
                  <input id="ledger" name="ledger" type="file" accept=".csv,.xlsx,.xls" required>
                </label>
              </div>
              <button id="analyzeBtn" class="analyze" type="submit">Analisar Arquivos</button>
            </form>
            <div id="status" class="status"></div>
          </div>
          <div class="upload-aside">
            <div class="aside-title">Fluxo da conferência</div>
            <div class="aside-subtitle">Do arquivo completo até a fila limpa de revisão.</div>
            <div class="upload-points">
              <div class="upload-point"><b>1</b><div><strong>Lê o razão inteiro</strong><span>Identifica contas, datas, saldos diários e participantes.</span></div></div>
              <div class="upload-point"><b>2</b><div><strong>Cruza com o plano</strong><span>Define a natureza esperada, grupo, classificação e redutoras.</span></div></div>
              <div class="upload-point"><b>3</b><div><strong>Mostra só a revisão</strong><span>Resume sequências repetidas e destaca os casos que precisam de conferência.</span></div></div>
            </div>
            <div class="aside-title" style="margin-top:24px">Colunas dos arquivos</div>
            <div class="requirements-list">
              <div class="requirement-card">
                <strong>Plano de contas</strong>
                <span>Obrigatória: Código.</span>
                <span>Recomendadas: Nome, Classificação, Grupo, Relatório, Tipo e Saldo/Natureza.</span>
                <span>Aceita layout SCI de impressão de campos/Plano padrão.</span>
              </div>
              <div class="requirement-card">
                <strong>Razão SCI</strong>
                <span>Obrigatórias: Histórico e Saldo.</span>
                <span>Usa Chave, Contra, Débito e Crédito quando existirem.</span>
                <span>Aceita razão com coluna Valor no lugar de Débito/Crédito.</span>
              </div>
            </div>
            <details class="about-details">
              <summary>Sobre o app</summary>
              <p>O Analisador Contábil cruza o plano de contas com o razão SCI e mostra apenas saldos que precisam de conferência, incluindo contas com participantes, fornecedores e clientes.</p>
            </details>
          </div>
        </div>
      </section>

      <section id="dashboard" class="hidden">
        <div class="metrics">
          <div class="metric"><span>Total de Inconsistências</span><strong id="mIssues" style="color:var(--red)">0</strong><em>Casos que precisam de revisão</em></div>
          <div class="metric"><span>Contas Afetadas</span><strong id="mAccounts" style="color:var(--yellow)">0</strong><em>Contas do plano de contas</em></div>
          <div class="metric"><span>Participantes Afetados</span><strong id="mParticipants" style="color:var(--blue)">0</strong><em>Fornecedores e clientes</em></div>
          <div class="metric period"><span>Período Analisado</span><strong id="mPeriod">-</strong><em id="mPeriodHint">0 dias analisados</em></div>
        </div>

        <section class="section-panel">
          <div class="section-head">
            <div>
              <h2>Principais Inconsistências</h2>
              <p>Lista dos casos que precisam de verificação</p>
            </div>
            <div class="tools">
              <input id="search" type="search" placeholder="Buscar conta ou participante...">
              <select id="typeFilter" aria-label="Filtros">
                <option value="">Filtros</option>
                <option>Fornecedor</option>
                <option>Cliente</option>
                <option>Conta</option>
              </select>
              <a id="exportLink" class="export-btn" href="#">Exportar Excel</a>
            </div>
          </div>
          <div class="table">
            <div class="table-grid">
              <div class="table-head">
                <div>Tipo</div><div>Código</div><div>Descrição</div><div>Saldo Esperado</div>
                <div>Saldo Atual</div><div>Valor</div><div>1a Ocorrencia</div><div>Dias Afetados</div><div>Acoes</div>
              </div>
              <div id="tableBody"></div>
            </div>
          </div>
          <div id="rowCount" class="count"></div>
        </section>
      </section>

      <section id="infoPanel" class="section-panel hidden">
        <div class="section-head">
          <div>
            <h2 id="infoTitle">Sobre</h2>
            <p id="infoSubtitle">Informações do Analisador Contábil.</p>
          </div>
        </div>
        <div id="infoContent" class="info-content"></div>
      </section>
    </main>
  </div>

  <div id="detailModal" class="modal-backdrop" onclick="closeDetails(event)">
    <aside class="drawer" onclick="event.stopPropagation()">
      <div class="drawer-head">
        <div>
          <h3 id="detailTitle">Detalhes da inconsistência</h3>
          <p id="detailSubtitle">Evolução do saldo e dias impactados.</p>
        </div>
        <button class="close-btn" type="button" onclick="closeDetails()">×</button>
      </div>
      <div id="detailBadges"></div>
      <div class="detail-grid" id="detailGrid"></div>
      <div class="chart-card">
        <h4>Evolução do saldo</h4>
        <div id="detailChart" class="mini-chart"></div>
      </div>
      <div class="involved-card">
        <h4>Dias envolvidos</h4>
        <div id="detailRows"></div>
      </div>
    </aside>
  </div>

  <script>
    let rows = [];
    let analysisId = "";
    let ledgerFileName = "";

    const money = new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const form = document.getElementById("uploadForm");
    const statusBox = document.getElementById("status");

    document.querySelectorAll(".menu-item").forEach(item => {
      item.addEventListener("click", () => showPage(item.dataset.page));
    });

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const data = new FormData(form);
      ledgerFileName = document.getElementById("ledger").files[0]?.name || "Razao SCI";
      statusBox.textContent = "Analisando arquivos...";
      statusBox.className = "status";
      document.getElementById("analyzeBtn").disabled = true;

      try {
        const response = await fetch("/analyze", { method: "POST", body: data });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Nao foi possivel analisar os arquivos.");
        rows = payload.rows || [];
        analysisId = payload.analysis_id;
        renderDashboard(payload.summary || {});
      } catch (error) {
        statusBox.textContent = error.message;
        statusBox.className = "status error";
      } finally {
        document.getElementById("analyzeBtn").disabled = false;
      }
    });

    document.getElementById("newFileBtn").addEventListener("click", () => {
      rows = [];
      analysisId = "";
      form.reset();
      document.getElementById("dashboard").classList.add("hidden");
      document.getElementById("uploadPanel").classList.remove("hidden");
      document.getElementById("infoPanel").classList.add("hidden");
      document.getElementById("newFileBtn").classList.add("hidden");
      document.getElementById("pageTitle").textContent = "Analisador Contábil";
      document.getElementById("pageSubtitle").textContent = "Conferência automática de saldos no razão SCI.";
      document.getElementById("sideFile").textContent = "Nenhum arquivo analisado";
      document.getElementById("sideTime").textContent = "-";
      document.getElementById("sideStatus").textContent = "Aguardando arquivos";
      statusBox.textContent = "";
      setActiveMenu("Resumo");
    });

    document.getElementById("search").addEventListener("input", renderTable);
    document.getElementById("typeFilter").addEventListener("change", renderTable);

    function renderDashboard(summary) {
      document.getElementById("uploadPanel").classList.add("hidden");
      document.getElementById("dashboard").classList.remove("hidden");
      document.getElementById("infoPanel").classList.add("hidden");
      document.getElementById("newFileBtn").classList.remove("hidden");
      document.getElementById("pageTitle").textContent = "Resumo da Análise";
      document.getElementById("pageSubtitle").textContent = "Visão geral das inconsistências encontradas no razão contábil.";
      document.getElementById("mIssues").textContent = summary.issues ?? rows.length;
      document.getElementById("mAccounts").textContent = summary.accounts ?? "0";
      document.getElementById("mParticipants").textContent = summary.participants ?? "0";
      document.getElementById("mPeriod").textContent = summary.period || "-";
      document.getElementById("mPeriodHint").textContent = summary.period_hint || "0 dias analisados";
      document.getElementById("exportLink").href = `/export?id=${analysisId}`;
      document.getElementById("sideFile").textContent = ledgerFileName;
      document.getElementById("sideTime").textContent = `Atualizado em ${new Date().toLocaleString("pt-BR").slice(0, 17)}`;
      document.getElementById("sideStatus").textContent = "Análise concluída";
      setActiveMenu("Resumo");
      renderTable();
    }

    function showPage(page) {
      setActiveMenu(page);
      if (page === "Resumo") {
        document.getElementById("infoPanel").classList.add("hidden");
        if (rows.length) {
          document.getElementById("uploadPanel").classList.add("hidden");
          document.getElementById("dashboard").classList.remove("hidden");
          document.getElementById("newFileBtn").classList.remove("hidden");
          document.getElementById("pageTitle").textContent = "Resumo da Análise";
          document.getElementById("pageSubtitle").textContent = "Visão geral das inconsistências encontradas no razão contábil.";
        } else {
          document.getElementById("dashboard").classList.add("hidden");
          document.getElementById("uploadPanel").classList.remove("hidden");
          document.getElementById("newFileBtn").classList.add("hidden");
          document.getElementById("pageTitle").textContent = "Analisador Contábil";
          document.getElementById("pageSubtitle").textContent = "Conferência automática de saldos no razão SCI.";
        }
        return;
      }

      if (page === "Inconsistências" && rows.length) {
        document.getElementById("uploadPanel").classList.add("hidden");
        document.getElementById("infoPanel").classList.add("hidden");
        document.getElementById("dashboard").classList.remove("hidden");
        document.getElementById("newFileBtn").classList.remove("hidden");
        document.getElementById("pageTitle").textContent = "Inconsistências";
        document.getElementById("pageSubtitle").textContent = "Casos que exigem conferência no razão analisado.";
        return;
      }

      if (page === "Contas") {
        renderInfoPage("Contas", "Resumo das contas afetadas pela análise.", renderAccountsSummary());
      } else if (page === "Participantes") {
        renderInfoPage("Participantes", "Fornecedores e clientes tratados separadamente dentro do razão.", renderParticipantsSummary());
      } else if (page === "Configurações") {
        renderInfoPage("Configurações", "Regras e formatos aceitos pelo analisador.", renderSettingsInfo());
      } else if (page === "Sobre") {
        renderInfoPage("Sobre", "Informações do Analisador Contábil.", renderAboutInfo());
      } else {
        renderInfoPage(page, "Casos que exigem conferência no razão analisado.", emptyAnalysisHtml());
      }
    }

    function setActiveMenu(page) {
      document.querySelectorAll(".menu-item").forEach(item => {
        item.classList.toggle("active", item.dataset.page === page);
      });
    }

    function renderInfoPage(title, subtitle, html) {
      document.getElementById("uploadPanel").classList.add("hidden");
      document.getElementById("dashboard").classList.add("hidden");
      document.getElementById("infoPanel").classList.remove("hidden");
      document.getElementById("newFileBtn").classList.toggle("hidden", !rows.length);
      document.getElementById("pageTitle").textContent = title;
      document.getElementById("pageSubtitle").textContent = subtitle;
      document.getElementById("infoTitle").textContent = title;
      document.getElementById("infoSubtitle").textContent = subtitle;
      document.getElementById("infoContent").innerHTML = html;
    }

    function emptyAnalysisHtml() {
      return '<div class="empty">Envie o plano de contas e o razão SCI para liberar esta visão.</div>';
    }

    function renderAccountsSummary() {
      if (!rows.length) return emptyAnalysisHtml();
      const grouped = {};
      rows.forEach(row => {
        const key = row["Conta analisada"] || row._codigo || "";
        if (!grouped[key]) grouped[key] = { codigo: row["Codigo da conta"] || "", conta: key, ocorrencias: 0, dias: 0 };
        grouped[key].ocorrencias += 1;
        grouped[key].dias += Number(row["Dias impactados"] || 1);
      });
      return `<div class="info-grid">${Object.values(grouped).slice(0, 12).map(item => `
        <div class="info-card"><strong>${escapeHtml(item.conta)}</strong><span>Código: ${escapeHtml(item.codigo)}</span><span>Ocorrências: ${item.ocorrencias}</span><span>Dias afetados: ${item.dias}</span></div>
      `).join("")}</div>`;
    }

    function renderParticipantsSummary() {
      if (!rows.length) return emptyAnalysisHtml();
      const participants = rows.filter(row => row._tipo === "Fornecedor" || row._tipo === "Cliente");
      if (!participants.length) return '<div class="empty">Nenhum fornecedor ou cliente com inconsistência encontrada.</div>';
      return `<div class="info-grid">${participants.slice(0, 12).map(row => `
        <div class="info-card"><strong>${escapeHtml(row._descricao || row["Conta analisada"] || "")}</strong><span>Tipo: ${escapeHtml(row._tipo || "")}</span><span>Conta: ${escapeHtml(row._codigo || "")}</span><span>Dias afetados: ${escapeHtml(String(row["Dias impactados"] || "1"))}</span></div>
      `).join("")}</div>`;
    }

    function renderSettingsInfo() {
      return `
        <div class="info-grid">
          <div class="info-card"><strong>Plano de contas</strong><span>Obrigatória: Código.</span><span>Recomendadas: Nome, Classificação, Grupo, Relatório, Tipo e Saldo/Natureza.</span><span>Aceita CSV, XLSX, XLS e layout SCI de impressão de campos/Plano padrão.</span></div>
          <div class="info-card"><strong>Razão SCI</strong><span>Obrigatórias: Histórico e Saldo.</span><span>Usa Chave, Contra, Débito e Crédito quando existirem.</span><span>Aceita coluna Valor no lugar de Débito/Crédito.</span></div>
        </div>`;
    }

    function renderAboutInfo() {
      return `
        <div class="info-grid">
          <div class="info-card"><strong>Objetivo</strong><span>Analisar o razão completo de uma vez e mostrar somente contas ou participantes que precisam de revisão.</span></div>
          <div class="info-card"><strong>Autoria</strong><span>Desenvolvido por Rayssa Mayara.</span><span>github.com/rayssamayarax</span></div>
        </div>`;
    }

    function filteredRows() {
      const query = normalize(document.getElementById("search").value);
      const type = document.getElementById("typeFilter").value;
      return rows.filter(row => {
        if (type && row._tipo !== type) return false;
        if (!query) return true;
        return normalize(`${row._codigo} ${row._descricao} ${row["Conta analisada"] || ""}`).includes(query);
      });
    }

    function renderTable() {
      const visible = filteredRows();
      const body = document.getElementById("tableBody");
      if (!visible.length) {
        body.innerHTML = '<div class="empty">Nenhum caso encontrado para os filtros atuais.</div>';
        document.getElementById("rowCount").textContent = `Mostrando 0 de ${rows.length} resultados`;
        return;
      }
      body.innerHTML = visible.map((row, index) => renderRow(row, index)).join("");
      document.getElementById("rowCount").textContent = `Mostrando ${visible.length} de ${rows.length} resultados`;
    }

    function renderRow(row, index) {
      const expected = row._esperado || "";
      const current = row._atual || "";
      return `
        <div class="table-row">
          <div>${badge(row._tipo, classForType(row._tipo))}</div>
          <div>${escapeHtml(row._codigo || "")}</div>
          <div class="desc" title="${escapeAttr(row._descricao || "")}">${escapeHtml(row._descricao || "")}</div>
          <div>${badge(expected, expected.toLowerCase())}</div>
          <div>${badge(current, current.toLowerCase())}</div>
          <div>${formatMoney(row["Saldo final do dia"])}</div>
          <div>${escapeHtml(row["Data"] || "")}</div>
          <div>${escapeHtml(String(row["Dias impactados"] || "1"))}</div>
          <div><button class="action-btn" type="button" onclick="openDetails(${index})">Ver</button></div>
        </div>
      `;
    }

    function openDetails(index) {
      const row = filteredRows()[index];
      if (!row) return;
      document.getElementById("detailTitle").textContent = row["Conta analisada"] || "Detalhes da inconsistência";
      document.getElementById("detailSubtitle").textContent = row._descricao || "";
      document.getElementById("detailBadges").innerHTML = [
        badge(row._tipo, classForType(row._tipo)),
        badge(row._esperado || "", (row._esperado || "").toLowerCase()),
        badge(row._atual || "", (row._atual || "").toLowerCase())
      ].join(" ");

      const detailItems = [
        ["Código", row._codigo || row["Codigo da conta"] || ""],
        ["Natureza esperada", row._esperado || ""],
        ["Natureza encontrada", row._atual || ""],
        ["Primeira ocorrência", row["Data"] || ""],
        ["Última ocorrência", row["Data final da sequencia"] || row["Data"] || ""],
        ["Dias afetados", row["Dias impactados"] || "1"],
        ["Quantidade de lançamentos", row._launch_count || row["Dias impactados"] || "1"],
        ["Saldo encontrado", formatMoney(row["Saldo final do dia"])]
      ];
      document.getElementById("detailGrid").innerHTML = detailItems.map(([label, value]) => `
        <div class="detail-card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(String(value))}</strong></div>
      `).join("");

      document.getElementById("detailChart").innerHTML = buildChart(row._history || []);
      document.getElementById("detailRows").innerHTML = buildMiniTable(row._detail_rows || []);
      document.getElementById("detailModal").classList.add("open");
    }

    function closeDetails(event) {
      if (event && event.target.id !== "detailModal") return;
      document.getElementById("detailModal").classList.remove("open");
    }

    function buildChart(history) {
      if (!history.length) return '<div class="empty">Sem histórico para exibir.</div>';
      const values = history.map(item => Number(item.saldo || 0));
      const min = Math.min(...values);
      const max = Math.max(...values);
      const spread = Math.max(max - min, 1);
      const width = 560;
      const height = 160;
      const points = history.map((item, i) => {
        const x = history.length === 1 ? width / 2 : (i / (history.length - 1)) * width;
        const y = height - ((Number(item.saldo || 0) - min) / spread) * (height - 20) - 10;
        return `${x.toFixed(1)},${y.toFixed(1)}`;
      }).join(" ");
      const zeroY = max <= 0 || min >= 0 ? null : height - ((0 - min) / spread) * (height - 20) - 10;
      return `
        <svg viewBox="0 0 ${width} ${height}" width="100%" height="180" role="img">
          ${zeroY ? `<line x1="0" x2="${width}" y1="${zeroY}" y2="${zeroY}" stroke="#24324A" stroke-dasharray="5 5"/>` : ""}
          <polyline fill="none" stroke="#3B82F6" stroke-width="4" stroke-linecap="round" stroke-linejoin="round" points="${points}"/>
        </svg>
      `;
    }

    function buildMiniTable(items) {
      if (!items.length) return '<div class="empty">Sem dias envolvidos para exibir.</div>';
      return `
        <table class="mini-table">
          <thead><tr><th>Data</th><th>Saldo do dia</th></tr></thead>
          <tbody>
            ${items.map(item => `
              <tr>
                <td>${escapeHtml(item.data || "")}</td>
                <td>${formatMoney(item.saldo || 0)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      `;
    }

    function badge(text, cls) {
      return `<span class="badge ${cls}">${escapeHtml(text || "")}</span>`;
    }
    function classForType(type) {
      if (type === "Fornecedor") return "fornecedor";
      if (type === "Cliente") return "cliente";
      return "conta";
    }
    function formatMoney(value) {
      const number = Number(value || 0);
      const formatted = money.format(Math.abs(number));
      return number < 0 ? `(${formatted})` : formatted;
    }
    function normalize(value) {
      return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    }
    function escapeHtml(value) {
      return String(value).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#039;" }[c]));
    }
    function escapeAttr(value) {
      return escapeHtml(value).replace(/`/g, "&#096;");
    }
  </script>
</body>
</html>"""


class AppHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_html(HTML)
            return
        if parsed.path == "/logo":
            self.handle_logo()
            return
        if parsed.path == "/export":
            self.handle_export(parsed.query)
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Pagina nao encontrada")

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/analyze":
            self.send_error(HTTPStatus.NOT_FOUND, "Pagina nao encontrada")
            return
        self.handle_analyze()

    def handle_analyze(self) -> None:
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                    "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
                },
            )

            plan_item = form["plan"] if "plan" in form else None
            ledger_item = form["ledger"] if "ledger" in form else None
            if plan_item is None or ledger_item is None:
                raise ValueError("Envie o plano de contas e o razao diario.")

            plan_df = read_accounting_file(plan_item.file, plan_item.filename or "")
            ledger_df = read_accounting_file(ledger_item.file, ledger_item.filename or "")
            result, inconsistencies = analyze_balances(ledger_df, plan_df)
            report = sorted_report(enrich_report(inconsistencies if not inconsistencies.empty else result.head(0), result))
            diagnostics = ledger_file_diagnostics(ledger_df)

            analysis_id = uuid.uuid4().hex
            ANALYSES[analysis_id] = report.drop(columns=[column for column in report.columns if column.startswith("_")])

            payload = {
                "analysis_id": analysis_id,
                "rows": json.loads(report.to_json(orient="records", force_ascii=False)),
                "summary": build_summary(result, inconsistencies),
                "warnings": build_warnings(diagnostics),
            }
            self.send_json(payload)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def handle_export(self, query: str) -> None:
        analysis_id = parse_qs(query).get("id", [""])[0]
        result = ANALYSES.get(analysis_id)
        if result is None:
            self.send_error(HTTPStatus.NOT_FOUND, "Analise nao encontrada")
            return

        data = dataframe_to_excel(result)
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", 'attachment; filename="analise_saldos_diarios_sci.xlsx"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def handle_logo(self) -> None:
        if not LOGO_PATH.exists():
            self.send_error(HTTPStatus.NOT_FOUND, "Logo nao encontrado")
            return

        data = LOGO_PATH.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "image/svg+xml; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_html(self, html: str) -> None:
        data = html.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, format: str, *args: object) -> None:
        return


def enrich_report(report: pd.DataFrame, result: pd.DataFrame | None = None) -> pd.DataFrame:
    report = report.copy()
    if report.empty:
        return report

    report["_tipo"] = report.apply(classify_row, axis=1)
    report["_codigo"] = report.apply(display_code, axis=1)
    report["_descricao"] = report.apply(display_description, axis=1)
    report["_esperado"] = report["Natureza esperada"].map({"credora": "Credor", "devedora": "Devedor"}).fillna("Revisao")
    report["_atual"] = report.apply(current_balance_label, axis=1)
    report["_history"] = [[] for _ in range(len(report))]
    report["_detail_rows"] = [[] for _ in range(len(report))]
    report["_launch_count"] = report["Dias impactados"].fillna(1).astype(str)

    if result is not None and not result.empty:
        history_by_account = {}
        result_work = result.copy()
        result_work["_data_dt"] = pd.to_datetime(result_work["Data"], format="%d/%m/%Y", errors="coerce")
        for account, group in result_work.sort_values("_data_dt").groupby("Conta analisada", sort=False):
            history_by_account[account] = group

        histories = []
        detail_rows = []
        launch_counts = []
        for _, row in report.iterrows():
            account = row.get("Conta analisada", "")
            group = history_by_account.get(account, pd.DataFrame())
            if group.empty:
                histories.append([])
                detail_rows.append([])
                launch_counts.append(str(row.get("Dias impactados", "1") or "1"))
                continue

            start = pd.to_datetime(row.get("Data", ""), format="%d/%m/%Y", errors="coerce")
            end = pd.to_datetime(row.get("Data final da sequencia", row.get("Data", "")), format="%d/%m/%Y", errors="coerce")
            if pd.isna(start):
                start = group["_data_dt"].min()
            if pd.isna(end):
                end = start

            window_start = start - pd.Timedelta(days=45)
            history = group[group["_data_dt"].between(window_start, end)].tail(80)
            involved = group[group["_data_dt"].between(start, end)]
            if involved.empty:
                involved = group[group["_data_dt"].eq(start)]

            histories.append(
                [
                    {"data": item["Data"], "saldo": float(item["Saldo final do dia"])}
                    for _, item in history.iterrows()
                ]
            )
            detail_rows.append(
                [
                    {"data": item["Data"], "saldo": float(item["Saldo final do dia"])}
                    for _, item in involved.iterrows()
                ]
            )
            launch_counts.append(str(len(involved)))

        report["_history"] = histories
        report["_detail_rows"] = detail_rows
        report["_launch_count"] = launch_counts
    return report


def current_balance_label(row: pd.Series) -> str:
    side = str(row.get("Lado do saldo", "")).strip().upper()
    if side == "D":
        return "Devedor"
    if side == "C":
        return "Credor"

    balance = float(row.get("Saldo final do dia", 0) or 0)
    expected = str(row.get("Natureza esperada", "")).lower()
    if abs(balance) <= 0.004:
        return "Zerado"
    if expected == "credora":
        return "Credor" if balance > 0 else "Devedor"
    if expected == "devedora":
        return "Devedor" if balance > 0 else "Credor"
    return "Revisao"


def classify_row(row: pd.Series) -> str:
    text = " ".join(
        str(row.get(column, ""))
        for column in ["Codigo da conta", "Conta analisada", "Nome da conta no razao", "Nome no plano de contas", "Grupo"]
    ).lower()
    if "fornecedor" in text or str(row.get("Codigo da conta", "")) == "148":
        return "Fornecedor"
    if "cliente" in text:
        return "Cliente"
    return "Conta"


def display_code(row: pd.Series) -> str:
    code = str(row.get("Codigo da conta", "")).strip()
    name = str(row.get("Nome da conta no razao", "")).strip()
    first = name.split(" ", 1)[0].strip()
    if first.isdigit():
        return f"{code} - {first}"
    return code


def display_description(row: pd.Series) -> str:
    name = str(row.get("Nome da conta no razao", "")).strip()
    pieces = name.split(" ", 1)
    if pieces and pieces[0].isdigit() and len(pieces) > 1:
        return pieces[1].lstrip("- ").strip()
    return name or str(row.get("Conta analisada", ""))


def build_summary(result: pd.DataFrame, inconsistencies: pd.DataFrame) -> dict[str, object]:
    dates = pd.to_datetime(result["Data"], format="%d/%m/%Y", errors="coerce") if not result.empty else pd.Series(dtype="datetime64[ns]")
    period = "-"
    period_hint = "0 dias analisados"
    if not dates.empty and pd.notna(dates.min()) and pd.notna(dates.max()):
        period = f"{dates.min().strftime('%d/%m/%Y')} a {dates.max().strftime('%d/%m/%Y')}"
        period_hint = f"{dates.dropna().nunique()} dias analisados"

    enriched = enrich_report(inconsistencies) if not inconsistencies.empty else inconsistencies
    participants = int(enriched["_tipo"].isin(["Fornecedor", "Cliente"]).sum()) if not enriched.empty else 0

    return {
        "accounts": int(inconsistencies["Codigo da conta"].nunique() if not inconsistencies.empty else 0),
        "participants": participants,
        "period": period,
        "period_hint": period_hint,
        "issues": int(len(inconsistencies)),
    }


def build_warnings(diagnostics: dict[str, object]) -> list[str]:
    account_codes = diagnostics.get("account_codes", [])
    if not account_codes:
        return []
    return ["Contas encontradas como blocos de razao neste arquivo: " + ", ".join(str(code) for code in account_codes)]


def find_port(start: int) -> int:
    port = start
    while True:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            if sock.connect_ex((HOST, port)) != 0:
                return port
        port += 1


def main() -> None:
    requested_port = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else PORT
    port = find_port(requested_port)
    server = ThreadingHTTPServer((HOST, port), AppHandler)
    url = f"http://{HOST}:{port}"
    print(f"Servidor local aberto em {url}")
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    server.serve_forever()


if __name__ == "__main__":
    main()
