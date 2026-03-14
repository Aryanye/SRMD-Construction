# Construction Quality Inspector

Upload construction-site photos, add a project name and work type for each image, and generate:

- an AI quality review for each image
- an overall site summary
- a downloadable PDF report for each image

## Run

```bash
cd /Users/aryanshah/Desktop/Codex/const_qual_app
pip install -r requirements.txt
python3 -m streamlit run construction_quality_app.py
```

The app reads `OPENAI_API_KEY` from the repo root `.env` file.
