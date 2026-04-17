Please make the following fixes to this Flask proposal generator app:

1. LINE SPACING: All paragraphs in the generated .docx must use single spacing with 0pt space before and 0pt space after. The current output has 2.0 line spacing. Fix this app-wide in the document generation code.

2. INDENTATION: Fix all indentation in the generated .docx to match these exact measurements:
   - Roman numeral section headers (I, II, III, IV): left indent = 0" (starts at the 1" page margin)
   - Lettered sub-items (A, B, C...): left indent = 0.5"
   - Numbered sub-sub-items (1, 2, 3...): left indent = 1.0"

3. DELETE SECTION HEADERS: In the web form UI, add a delete button next to each of the four Roman numeral section headers (I through IV) so the user can remove a section entirely if not needed. If a section is deleted it should not appear in the .docx output at all. Include a way to restore a deleted section.

4. AUTO-NUMBERING IN WORD: When the .docx is opened in Microsoft Word and the user adds new list items after the exported content, the lettered sub-items (A, B, C) and numbered sub-sub-items (1, 2, 3) should continue sequentially. Fix the list style definitions in python-docx so Word recognizes and continues the numbering automatically.

Please make all four fixes and confirm which files were changed.
