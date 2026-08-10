# AccountingAI — Live Deployment Package

## What this package contains

- `app.py` — production-oriented Streamlit app
- `requirements.txt` — Python dependencies
- `.streamlit/config.toml` — Streamlit runtime configuration
- `.streamlit/secrets.toml.example` — secret template
- `.gitignore` — prevents secrets and uploaded financial files from being committed

## Important cloud change

The original local version used Ollama/qwen3:8b for the AI fallback.
This live version uses the OpenAI API as an optional server-side fallback.

The deterministic accounting rulebook still runs first.
If no `OPENAI_API_KEY` is configured, unknown accounts are sent to manual review
instead of crashing the app.

## Local test

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud deployment

1. Create a GitHub repository, for example `accountingai`.
2. Upload this package to the repository.
3. Go to https://share.streamlit.io/ and connect GitHub.
4. Create an app.
5. Select the repository, `main` branch, and `app.py`.
6. Open Advanced settings → Secrets.
7. Add:

```toml
OPENAI_API_KEY = "YOUR_KEY"
OPENAI_MODEL = "gpt-5"
```

8. Deploy.
9. Test with the supplied Trial Balance test suite before sharing publicly.

## Security rules

- Never put the OpenAI API key in `app.py`.
- Never commit `.streamlit/secrets.toml`.
- Do not commit users' Trial Balance files.
- Treat uploaded financial data as confidential.
- Review privacy/terms requirements before public launch.

## Comparative analysis behavior

Comparative analysis is shown only when the user actually uploads a
comparative Trial Balance for that prepared run. A single current-year
Trial Balance does not generate fake 0%/100% movements.
