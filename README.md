# School Homework Agent

Automates the weekly homework workflow:

1. Sign in to the school portal.
2. Open the homework page.
3. Find and download homework PDFs.
4. Extract the questions.
5. Ask OpenAI to prepare student-ready answers.
6. Create one organized DOCX file.

## Security

Never commit:
- school username/password
- student-specific homework URLs
- OpenAI API keys
- cookies or browser sessions

Configure these only as private environment variables in Render.

## Environment variables

Required:
- `SCHOOL_LOGIN_URL`
- `SCHOOL_HOMEWORK_URL`
- `SCHOOL_LOGIN_ID`
- `SCHOOL_PASSWORD`
- `OPENAI_API_KEY`

Optional:
- `OPENAI_MODEL` (default: `gpt-5.6-luna`)

## Render

Suggested build command:

`npm install && npx playwright install chromium`

Start command:

`npm start`

The first run is intentionally diagnostic. School portals vary, so login selectors and PDF-link detection may need a small adjustment after we inspect the first Render log.

## Output

The generated Word file is written under `output/`.

Important: Render cron files are ephemeral. After login/download/solving is verified, add a delivery step (for example email or cloud storage) so the weekly DOCX is retained.
