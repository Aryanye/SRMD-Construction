# Construction Quality Inspector

Upload one or more construction-site photos for the same project batch and generate:

- one combined AI quality review for the batch
- clear priority improvements to make on the project
- a downloadable PDF report with the uploaded photos embedded

## Run

```bash
cd /Users/aryanshah/Desktop/Codex/const_qual_app
pip install -r requirements.txt
python3 -m streamlit run construction_quality_app.py
```

The app reads `OPENAI_API_KEY` from the repo root `.env` file.
