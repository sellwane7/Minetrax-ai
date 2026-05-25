# ⛏️ MINETRAX AI

### Mining Safety Intelligence Platform — South Africa

*AI-powered hazard detection, compliance auditing, and intelligent documentation built on the Mine Health and Safety Act (Act 29 of 1996)*

---

</div>

## 🧭 Overview

**Minetrax AI** is a web-based safety intelligence platform designed specifically for the South African mining industry. It combines AI language models, machine learning algorithms, and deep knowledge of MHSA Act 29 of 1996 to help mine safety officers, supervisors, and engineers:

- Predict and score hazard risk before incidents occur
- Audit compliance against mandatory MHSA safety protocols
- Generate formal, MHSA-compliant documentation in seconds
- Access 24/7 AI-powered expert safety advice

It covers **54 real mine operations** across **9 South African provinces** — from deep-level gold mines in Gauteng to open-pit diamond operations in Limpopo.

---

## 🖥️ Live Demo

> 🚀 **Try Minetrax AI live:** [https://minetrax.streamlit.app](https://minetrax.streamlit.app)
>
> *No sign-up required for now

---

## ✨ Features

### 📊 Safety Dashboard
- Real-time compliance score gauge (0–100%)
- Active mine overview: operations, zones, and mine profile
- MHSA quick-reference card panel (Sections 5, 11, 21, 23, 49, 54)
- Live gas anomaly monitor with Z-score statistical detection and linear regression trend forecasting

### ⚠️ Hazard Intelligence
- Input site parameters: depth, gas ppm, temperature, fatigue, maintenance status, incident history
- ML accident probability prediction (sigmoid-normalised weighted model)
- Severity classification: Minor / Moderate / Serious / Critical / Fatal
- Feature contribution visualisation (bar chart)
- AI-generated site-specific hazard report via Groq LLM

### ✦ Compliance Engine
- 7-protocol MHSA checklist with compliance scoring
- K-Means zone risk clustering (High / Medium / Low tiers)
- AI-generated compliance audit report with gap analysis and corrective actions
- Downloadable PDF and Word exports

### ◉ Report Forge
Generates 6 types of formal MHSA-compliant documents:

| Report Type | MHSA Reference |
|---|---|
| Incident / Accident Report | Section 23 |
| Near Miss Report | Section 11 |
| Risk Assessment | Section 11 |
| Toolbox Talk | Section 5 |
| Safety Observation Report | Chapter 4 |
| Compliance Audit Report | Act 29 of 1996 |

All documents include branded PDF output, metadata tables, and a formal signature block (Mine Manager, Safety Officer, H&S Representative, DMR Inspector).

### ◎ Safety Advisor
- Conversational AI chatbot with a 25+ year expert mining safety persona
- Context-aware: knows your active mine, province, and mine type
- Covers: MHSA Act, DMR inspections, Section 54 stop-work, ventilation, blasting, TMM, emergency response
- Chat history maintained within session; quick-start suggested questions

---

## 🤖 ML & AI Components

| Component | Method |
|---|---|
| Accident Probability | Weighted logistic regression (6 inputs, sigmoid output) |
| Gas Anomaly Detection | Z-score statistical outlier detection + linear regression trend |
| Zone Risk Clustering | Custom 2-iteration K-Means (3 tiers) |
| Safety Action Recommender | Q-Learning reinforcement learning Q-table |
| Report Generation | Groq LLM API (LLaMA-based, low-latency inference) |

---

## 🗂️ Province & Mine Coverage

| Province | Mine Type | Example Sites |
|---|---|---|
| Gauteng | Deep-Level Gold | TauTona, Mponeng, South Deep |
| Limpopo | Open Pit / Diamond & Platinum | Venetia, Mogalakwena |
| Mpumalanga | Underground Coal | Goedehoop, Isibonelo |
| North West | Platinum (Bushveld Complex) | Rustenburg, Marikana |
| Free State | Intermediate-Level Gold | Beatrix, Tshepong |
| Northern Cape | Iron Ore / Manganese | Sishen, Kolomela |
| KwaZulu-Natal | Anthracite / Heavy Mineral Sands | Somkhele, Richards Bay Minerals |
| Eastern Cape | Quarry / Limestone | AfriSam Rietfontein, PPC Riebeeck |
| Western Cape | Quarry / Industrial Minerals | AfriSam Piketberg, Cape Lime Works |

---

## ☁️ Deploy to Streamlit Cloud

1. Fork this repository
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo and select `app.py`
4. Under **Advanced settings → Secrets**, add:
   ```
   GROQ_API_KEY = "your_key_here"
   ```
5. Click **Deploy**

---

## 📦 Requirements

```
streamlit
groq
python-docx
fpdf2
Pillow
python-dotenv
```

## 📁 Project Structure

```
minetrax-ai/
├── app.py                  # Main Streamlit application
├── requirements.txt        # Python dependencies
├── .env.example            # Environment variable template
└── README.md
```
---

## 🔮 Roadmap

- [ ] Real-time IoT sensor integration via MQTT
- [ ] PWA wrapper for offline underground use
- [ ] Multi-language support (Zulu, Afrikaans, Sotho)
- [ ] Role-based authentication (Mine Manager / Safety Officer / Supervisor)

- [ ] Historical trend dashboards with persistent storage
- [ ] DMR Form 10.1 auto-fill from report data
- [ ] Geospatial mine zone risk mapping
- [ ] WhatsApp / SMS safety alert integration

- [ ] Computer vision PPE and hazard detection from camera feeds
- [ ] Digital twin mine simulation interface
- [ ] Multi-mine portfolio executive dashboard
- [ ] Smart PPE wearable biometric integration

---

## 📋 MHSA Compliance Reference

Minetrax AI is built on and references the following MHSA Act 29 of 1996 provisions:

| Section | Description |
|---|---|
| Section 5 | Employer general duty of care |
| Section 11 | Mandatory risk assessments |
| Section 21 | Pre-shift equipment inspections |
| Section 23 | Incident reporting to DMRE |
| Section 49 | PPE provision and training |
| Chapter 4 | Health and Safety Representatives |
| Section 54 | DMR Inspector powers and stop-work orders |

---

## ⚠️ Disclaimer

Minetrax AI is a decision-support tool intended to assist qualified mining safety professionals. It does not replace the statutory obligations of Mine Managers, Safety Officers, or Principal Officers under the Mine Health and Safety Act, Act 29 of 1996. All AI-generated reports must be reviewed by a competent person before use in formal MHSA processes. AI-generated content may contain errors and should not be submitted to the DMRE without verification.

---

<div align="center">

**⛏ Minetrax AI · Mining Safety Intelligence · South Africa**

*Built for the safety of South Africa's miners*

</div>
