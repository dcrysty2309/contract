# ContractFlow

Run the local app:

```powershell
python app.py
```

Open:

```text
http://127.0.0.1:8000
```

The app generates a real `.docx` contract from the Word template files in this folder.

Extraction flow:

1. OpenAI Vision reads the uploaded PDF/image directly.
2. OpenAI returns structured JSON.
3. The UI shows the extracted JSON in an editable form so you can verify and fix fields before generating the contract.

The main flow uses `OPENAI_API_KEY` for Vision extraction.
If Vision fails or the account is out of quota, the app now falls back to the local OCR path for images and PDFs.
