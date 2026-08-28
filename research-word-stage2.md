# Word Stage 2 research

Source: https://learn.microsoft.com/en-us/javascript/api/word/word.table?view=word-js-preview

Microsoft's official Word.Table documentation states that Word.Table represents a real table in a Word document and can be inserted through Word.Body, Word.Range, Word.Paragraph, or Word.ContentControl. The documented TypeScript example uses `context.document.body.insertTable(rowCount, columnCount, "Start", data)` and then applies a built-in table style before `context.sync()`. The page lists table alignment and table values as supported properties. This supports implementing Word insertion as actual editable tables rather than pasted images or plain text.
