# cash-ai-02-excel

Cloud Run API for Excel generation, using the same project layout as `cash-ai-02.zip`.

## Endpoint
- `POST /v1/pipeline`
- returns JSON: `{ok, b64, filename, content_type, generated_at}`

## Request
The request body is written to `output.json` and consumed by `app/pipeline/originals/colab101.py`.

## Response
The response format is close to the original Colab `download_excel` callback:
- `ok`
- `b64`
- `filename`

## Notes
- Colab-only display/callback behavior is disabled.
- The service focuses on Excel generation only.
