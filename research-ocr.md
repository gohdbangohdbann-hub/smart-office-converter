# OCR Engine Research

## Azure Document Intelligence — official findings

Source: https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/language-support/ocr?view=doc-intel-4.0.0

The Azure Document Intelligence v4.0 documentation states that the Read model extracts printed and handwritten text, while the Layout model extracts text, tables, document structure, and selection marks. Arabic (`ar`), English (`en`), and French (`fr`) are listed among the supported printed-text languages. The documentation also states that universal deep-learning models extract multilingual text, including lines with mixed languages, and that a language code is optional; forcing a language when uncertain can return incomplete or incorrect text. Locale is also optional because the service can auto-detect the text language.

Decision direction: implement a provider-neutral OCR interface and use Azure Document Intelligence Layout as the first production-shaped adapter because it covers Arabic/mixed-language OCR plus tables and layout in one response. Keep provider selection behind a server-side environment configuration so the UI and unified result model remain unchanged. Do not expose provider keys to the browser.

Testing implication: benchmark Arabic, mixed Arabic/French/English, Arabic and Western numerals, PDF text/scanned/mixed pages, image preprocessing cases, and table extraction separately; confidence is a review signal rather than a correctness guarantee.
