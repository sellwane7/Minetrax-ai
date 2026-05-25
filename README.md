# ⛏ Minetrax AI — Mining Safety Intelligence Platform

> **Protecting South Africa's miners with AI-powered hazard detection, real-time compliance auditing, and intelligent documentation.**

---

## 🚀 Quick Start

```bash
git clone https://github.com/yourusername/harvestmind-ai.git
cd minetrax-ai
pip install -r requirements.txt
# Add GROQ_API_KEY to .env
streamlit run app.py
```

---

## 📋 What Is Minetrax AI?

Minetrax AI is a South African mining safety intelligence platform combining machine learning with regulatory expertise. It delivers instant risk assessments, MHSA Act 29 of 1996 compliant audit reports, and evidence-backed safety recommendations — with PDF and Word export including signature blocks for formal authorisation.

---

## ✨ What's New in This Version

- **Beautiful Welcome Screen** — Full-page hero with background image, custom SVG logo, and user story examples. Mine selection happens here before entering the platform.
- **Start Button Flow** — Users select their mine on the welcome page and click **Launch Platform** before seeing any modules.
- **Signature Blocks on PDFs** — Every PDF and Word report now includes a formal authorisation section with signature lines for Mine Manager, Safety Officer, H&S Representative, and DMR Inspector.
- **Improved Colour Contrast** — All labels, selects, sliders, and form elements now have high-contrast text (no more grey-on-grey issues).
- **Richer AI Prompts** — Hazard analysis now has 8 sections including PPE recommendations and incident reporting obligations. Compliance audits include DMR readiness assessment.
- **No API key mentions in the UI** — Status indicator is subtle; no disruptive warnings shown to end users.

---

## 🖥️ Features

### Welcome Page
- Full-screen hero with background photo + custom SVG logo
- Platform description and two real user stories (Thabo & Patricia)
- Mine + Province selection with Launch button
- Feature overview cards

### ◈ Safety Dashboard
- Live operational metrics
- Z-score gas anomaly detection + linear regression trend
- MHSA Act 29 quick reference cards
- Active mine overview

### ⚠️ Hazard Intelligence
- AI risk assessment (8-section analysis)
- Logistic regression accident probability (0–100%)
- Severity classification with visual breakdown
- PDF & Word export with signature block

### ✦ Compliance Engine
- MHSA Act 29 of 1996 automated audits
- K-Means zone clustering
- Gap analysis with MHSA section references
- DMR readiness assessment section
- PDF & Word export with signature block

### ◉ Report Forge
- Incident Reports, Risk Assessments, Audit Summaries, Near Miss Reports, Equipment Reports, Fatality Reports
- Image evidence with AI vision analysis
- Q-Learning action recommendations
- **PDF & Word always available — signature block included**

### ◎ Safety Advisor
- 24/7 AI mining safety Q&A
- Contextual to active mine, province, and mine type

---

## 🤖 Machine Learning Methods

| Method | Type | Where Used |
|--------|------|-----------|
| Logistic Regression | Supervised | Accident probability — Hazard Intelligence |
| Linear Regression | Supervised | Gas trend analysis — Dashboard |
| K-Means Clustering | Unsupervised | Zone risk classification — Compliance Engine |
| Z-Score Detection | Unsupervised Statistical | Gas anomaly detection — Dashboard |
| Q-Learning | Reinforcement Learning | Safety action recommendations — Report Forge |

---

## 📦 Installation

### Prerequisites
- Python 3.9+
- A [Groq API key](https://console.groq.com) (free tier available)

### Steps

```bash
# 1. Clone
git clone https://github.com/yourusername/harvestmind-ai.git
cd harvestmind-ai

# 2. Virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 3. Install
pip install -r requirements.txt

# 4. API key
echo "GROQ_API_KEY=your_key_here" > .env

# 5. Add background image
# Place mines.jpg in same directory as app.py

# 6. Run
streamlit run app.py
```

---

## ☁️ Deployment (Streamlit Community Cloud)

1. Push to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io) → New app
3. Connect repo, set main file to `app.py`
4. Advanced settings → Secrets → add: `GROQ_API_KEY = "your_key_here"`
5. Deploy ✓

---

## 📁 Project Structure

```
harvestmind-ai/
├── app.py              # Main Streamlit application
├── mines.jpg           # Hero background image (required)
├── requirements.txt    # Python dependencies
├── .env                # Local API key (NOT committed)
├── .gitignore
└── README.md
```

---

## 🔒 Security

- Never commit `.env` (already in `.gitignore`)
- Use Streamlit Secrets for production
- No user data stored between sessions

---

## 📜 Regulatory Framework

Built on **Mine Health and Safety Act, Act 29 of 1996 (MHSA)**:
- Section 5: Employer duty of care
- Section 11: Risk assessment obligations
- Section 23: Incident reporting (24-hour rule)
- Section 49: PPE requirements
- Chapter 4: Health and Safety Representatives
- Section 54: DMR Inspector powers

---

*Built with ⛏️ for the safety of South Africa's miners.*
