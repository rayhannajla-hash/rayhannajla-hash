# 🤖 AI Engineer Portfolio — Budhi

> Self-taught AI Engineer | LLM & Prompt Engineering | Local AI Infrastructure  
> 📍 Indonesia · Manufacturing & Touch Panel Design Domain

---

## 👤 About

Saya adalah Team Leader desain **resistive touch panel** di lingkungan manufaktur, yang secara mandiri membangun berbagai sistem dan tools berbasis AI untuk meningkatkan produktivitas tim dan mengeksplorasi kemungkinan otomasi.

Semua proyek dibangun dalam kondisi nyata:
- Laptop kerja dengan **IT restrictions** (no admin, no CMD, no internet untuk data sensitif)
- Data desain yang **confidential** — mendorong solusi offline/lokal
- Hardware terbatas: Intel i7 / ~16GB RAM / Intel Iris Xe

---

## 📊 Stats

| Metric | Value |
|--------|-------|
| Total Projects | 12+ |
| Domain Coverage | 5 area |
| Learning Approach | 100% Self-taught |
| Primary Focus | LLM, Prompt Engineering, Local AI |

---

## 🗂️ Project Index

### 🏗️ Local AI Infrastructure

| # | Project | Status | Stack |
|---|---------|--------|-------|
| 01 | [Virtual Design Engineer (VDE)](#01-virtual-design-engineer) | ✅ Done | Ollama, AnythingLLM, Qwen2.5 7B |
| 09 | [AINUN Telegram Bot](#09-ainun-telegram-bot) | ✅ Done | OpenClaw, Telegram API |
| 10 | [OpenClaw Self-hosted Gateway](#10-openclaw-self-hosted-gateway) | 🔄 In Progress | GCP, Docker |

### 🏭 Work Productivity Tools

| # | Project | Status | Stack |
|---|---------|--------|-------|
| 02 | [FreeCAD Touch Panel Macro Suite](#02-freecad-touch-panel-macro-suite) | ✅ Done | FreeCAD, Python, openpyxl |
| 07 | [Design Standard Hub](#07-design-standard-hub) | ✅ Done | Excel |
| 08 | [LLM Migration Script](#08-llm-migration-script) | ✅ Done | Python, LM Studio API |
| 06 | [PO Order Management System](#06-po-order-management-system) | ✅ Done | Python, xlsxwriter |

### 🌐 AI-Powered Web Tools (Single-file HTML)

| # | Project | Status | Stack |
|---|---------|--------|-------|
| 03 | [AppForge AI & CodeForge AI](#03-appforge-ai--codeforge-ai) | ✅ Done | HTML/JS, Multi-provider AI |
| 04 | [ExcelAI Template Generator](#04-excelai-template-generator) | ✅ Done | HTML/JS, SheetJS |
| 05 | [AI Document Editor](#05-ai-document-editor) | ✅ Done | HTML/JS, PDF.js, mammoth |

### 📊 Productivity Dashboards

| # | Project | Status | Stack |
|---|---------|--------|-------|
| 11 | [Task.html Dashboard](#11-taskhtml-team-dashboard) | ✅ Done | Alpine.js, HTML |
| 12 | [AI Engineer Roadmap Tool](#12-ai-engineer-roadmap-tool) | ✅ Done | HTML/JS |

### 📧 Email Automation

| # | Project | Status | Stack |
|---|---------|--------|-------|
| 13 | [Outlook Email Manager & Daily Summary](#13-outlook-email-manager--daily-summary) | 🔄 In Progress | Python, openpyxl, win32com |

---

## 📋 Project Details

---

### 01. Virtual Design Engineer

**Sistem AI lokal sepenuhnya offline untuk standar desain rahasia perusahaan**

Dibangun karena kebutuhan nyata: tim perlu konsultasi standar IPC/IEC dan dokumen internal, tapi data sensitif tidak boleh dikirim ke cloud.

**Fitur:**
- RAG dari dokumen PDF standar IPC/IEC
- Interface via AnythingLLM + Open WebUI
- Model: Qwen2.5 7B Instruct (stable), juga diuji Phi-3.1 Mini & Mistral 7B
- Zero cloud dependency — semua tetap lokal

**Tech:** `Ollama` `AnythingLLM` `Open WebUI` `LM Studio` `Qwen2.5 7B`

**Impact:** Tim bisa konsultasi standar teknis tanpa risiko kebocoran data ke cloud

---

### 02. FreeCAD Touch Panel Macro Suite

**Parametric CAD automation untuk resistive touch panel 340×190mm**

Kumpulan macro FreeCAD berbasis Python untuk otomasi desain panel sentuh. Dikembangkan dengan Claude Code di laptop Linux Mint pribadi karena IT restrictions di laptop kerja.

**Macro Coverage:**
- Dot Spacer Pattern distribution
- Bus Bar Generator (horizontal/vertikal)
- Electrode Grid dengan parameter Excel
- FPC Tail Area
- Export DXF/STEP untuk manufaktur

**Tech:** `FreeCAD` `Python` `openpyxl` `Claude Code` → Output: `DXF` `STEP`

**Impact:** Target tim meningkat dari 1 → 5 panel design/hari

---

### 03. AppForge AI & CodeForge AI

**Generator aplikasi & database berbasis AI multi-provider**

Dua tool web single-file HTML. AppForge untuk generate struktur aplikasi, CodeForge untuk generate schema database dan SQL dari deskripsi natural language.

**Fitur:**
- Multi-provider: Groq, OpenRouter, Gemini, Anthropic dalam satu UI
- Zero installation — satu file HTML, buka di browser
- API key tersimpan lokal (tidak ke server)

**Tech:** `HTML/JS` `Groq API` `OpenRouter` `Gemini API` `Anthropic API`

---

### 04. ExcelAI Template Generator

**AI-powered generator template Excel dari natural language**

Deskripsikan kebutuhan spreadsheet → AI generate template Excel siap pakai.

**Tech:** `HTML/JS` `Free-tier AI APIs` `SheetJS`

**Impact:** Template Excel dalam menit vs jam setup manual

---

### 05. AI Document Editor

**Editor dokumen cerdas multi-format dengan AI**

Upload PDF/Excel/Word → AI bisa rangkum, revisi, atau ekstrak informasi.

**Tech:** `HTML/JS` `PDF.js` `mammoth.js` `SheetJS` `Free AI APIs`

---

### 06. PO Order Management System

**Excel dashboard untuk 478 Purchase Order — workflow manufaktur Jepang**

Multi-sheet dashboard dengan summary, filtering, dan status tracking. Migrasi dari openpyxl ke xlsxwriter untuk stabilitas.

**Tech:** `Python` `xlsxwriter` `Excel`

**Impact:** Tracking 478 PO terstruktur menggantikan tracking manual

---

### 07. Design Standard Hub

**Konsolidasi standar desain tim dalam satu Excel master**

Single source of truth untuk semua standar desain tim. Fase 1 dari rencana 3-fase peningkatan produktivitas.

**Tech:** `Excel`

---

### 08. LLM Migration Script

**AI-powered extraction standar desain dari file Excel tersebar**

Script Python menggunakan LM Studio local API untuk batch processing dan klasifikasi standar dari file Excel yang tidak terstruktur.

**Tech:** `Python` `LM Studio API` `Qwen2.5 7B` `openpyxl`

**Highlight:** 100% lokal — tidak ada data keluar ke internet

---

### 09. AINUN Telegram Bot

**Personal AI assistant via Telegram**

Bot Telegram pribadi terhubung ke OpenClaw self-hosted. Nama: "AINUN", model: kimi-k2.5:cloud. Resolved: config corruption + pairing issues setelah forced power-off.

**Tech:** `OpenClaw` `Telegram Bot API` `kimi-k2.5` `Linux Mint`

---

### 10. OpenClaw Self-hosted Gateway

**AI gateway di Google Cloud Platform via Docker**

Perjalanan multi-platform: Railway → Render → Koyeb → Oracle Cloud → **GCP** (berhasil). VM + Docker deployment running, troubleshooting browser extension connection.

**Tech:** `GCP` `Docker` `OpenClaw` `Linux Mint`

**Status:** 🔄 VM & Docker ✅ | Browser extension 🔧

---

### 11. Task.html Team Dashboard

**Dashboard manajemen tugas tim — IT-restricted friendly**

Alpine.js single-file app dengan KPI cards, import/export JSON, column mapping fleksibel. Dirancang untuk berjalan tanpa instalasi di laptop dengan IT restrictions.

**Tech:** `Alpine.js` `HTML/CSS/JS`

---

### 12. AI Engineer Roadmap Tool

**Interactive learning roadmap dengan progress tracking**

Tool untuk memvisualisasikan dan melacak progress belajar AI Engineering. 4 fase, 20 topik, localStorage persistence.

**Tech:** `HTML/JS` `localStorage`

---

### 13. Outlook Email Manager & Daily Summary

**Otomasi email Outlook — copy ke Excel, sortir per penerima, summary harian**

Workflow 3 tahap untuk mengelola inbox Outlook secara otomatis setiap pagi:

**Flow:**
```
[Copy Email Baru dari Outlook] → [Sortir di Excel by Recipient] → [Buat Summary Pagi]
```

**Task Breakdown:**

| # | Task | Status | Deskripsi |
|---|------|--------|-----------|
| 1 | **Copy Email dari Outlook ke Excel** | 🔲 Todo | Baca inbox Outlook via win32com, copy subject/sender/date/body ke sheet Excel baru |
| 2 | **Sortir Email by Direct Recipient** | 🔲 Todo | Filter & kategorikan email berdasarkan penerima langsung: `to:budhiarso` vs email lain (CC/BCC/broadcast) |
| 3 | **Buat Summary Email Setiap Pagi** | 🔲 Todo | Generate ringkasan harian: jumlah email masuk, email direct to Budhiarso, topik utama, action items |

**Rencana Implementasi:**

- `email_fetcher.py` — Koneksi ke Outlook via `win32com.client`, tarik email baru sejak kemarin
- `email_sorter.py` — Parse field `To`, pisahkan direct email vs CC/broadcast
- `summary_generator.py` — Buat summary sheet di Excel + opsional kirim via Telegram bot (AINUN)
- Scheduler: Windows Task Scheduler / cron job untuk jalankan otomatis tiap pagi 07:00

**Tech:** `Python` `win32com` `openpyxl` `Outlook COM API`

**Target Impact:** Zero email terlewat, laporan harian siap sebelum jam kerja mulai

---

## ⚡ Skills

```
Prompt Engineering        ████████░░  Advanced
Python Scripting          ███████░░░  Intermediate
Local AI / LM Studio      ████████░░  Advanced
HTML/JS Single-file Apps  ████████░░  Advanced
Excel / xlsxwriter        ████████░░  Advanced
RAG Systems               ██████░░░░  Intermediate
Docker / GCP              ██████░░░░  Intermediate
FreeCAD / CAD Macro       ██████░░░░  Intermediate
Multi-provider LLM API    ████████░░  Advanced
IT-restricted Solutions   █████████░  Expert
```

---

## 🔑 Key Principles

1. **Local-first** — Data sensitif tidak pernah keluar ke cloud
2. **No-install solutions** — Single HTML file untuk IT-restricted environments
3. **Real-world problems** — Setiap project punya use case nyata di tempat kerja
4. **Free-tier friendly** — Maksimalkan AI gratis sebelum bayar

---

*Dokumentasi dibuat: Maret 2026 · Semua project dibangun secara self-taught*  
*"Portofolio beats degree." — Budhi 🚀*
