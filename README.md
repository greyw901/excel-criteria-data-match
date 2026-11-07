# Excel Criteria Data Match

Matches rows between two Excel sources by criteria-label text (using
"nth occurrence of this text, top-to-bottom" to disambiguate repeated
labels) and writes a new workbook that mirrors the reference sheet's
exact row layout, with the matched amount ready to paste straight into
a real workbook.

See [`excel_criteria_data_match Instruction.txt`](excel_criteria_data_match%20Instruction.txt)
for full usage, developer notes, environment setup, and troubleshooting.

## Quick start

```
pip install openpyxl xlrd
python excel_criteria_data_match.py
```

## License

MIT - see [LICENSE](LICENSE).
