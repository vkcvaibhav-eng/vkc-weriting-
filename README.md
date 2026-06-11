# Agro Sandesh Gujarati Article Builder

Streamlit app for researching timely Gujarat agriculture topics and drafting Gujarati
Agro Sandesh-style extension articles with Gemini.

The default workflow helps you:

- Research current and prevailing topics for South Gujarat and wider Gujarat.
- Focus on agricultural acarology, agricultural entomology, or both.
- Select the best month-relevant topic using grounded Gemini research.
- Draft a Gujarati farmer-centric article inspired by Dr. M. S. Swaminathan's
  science-for-farmer-welfare communication style.
- Export the article as DOCX.

The original research-paper writer remains available from the sidebar workflow switch.

## Run Locally

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Streamlit Community Cloud

1. Push this folder to a GitHub repository.
2. Create a new app in Streamlit Community Cloud.
3. Set `streamlit_app.py` as the app file.
4. Add this secret in the Streamlit app settings:

```toml
GEMINI_API_KEY = "your Gemini API key"
```

Do not commit `.streamlit/secrets.toml`; use `.streamlit/secrets.toml.example`
only as a template.

## Notes

Gemini's Google Search grounding is used for current topic discovery and source
collection. Generated articles should still be reviewed by an agricultural expert,
especially before recommending pesticides, doses, waiting periods, or local advisories.
