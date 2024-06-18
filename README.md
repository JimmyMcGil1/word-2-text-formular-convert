# Chuyển từ file world sang text format công thức theo Latex dùng syncfusion

# Link document

[https://help.syncfusion.com/file-formats/docio/working-with-mathematical-equation](https://help.syncfusion.com/file-formats/docio/working-with-mathematical-equation)

# Todo List

## File processing

- [x]  Read word document file
- [x]  Save file

## Document processing

- [x]  Parse math block/ math pattern
- [x]  Replace math block with math latex pattern

## Chuyển đổi block math sang latex math pattern

- [ ]  Accent (chỉ mới chuyển được widehat)

[https://help.syncfusion.com/file-formats/docio/working-with-mathematical-equation#equation-array](https://help.syncfusion.com/file-formats/docio/working-with-mathematical-equation#equation-array)

- [x]  Bar

[https://help.syncfusion.com/file-formats/docio/working-with-mathematical-equation#bar](https://help.syncfusion.com/file-formats/docio/working-with-mathematical-equation#bar)

- [x]  Box
- [x]  Border box
- [x]  Delimiter
- [ ]  Equation array

[https://help.syncfusion.com/file-formats/docio/working-with-mathematical-equation#equation-array](https://help.syncfusion.com/file-formats/docio/working-with-mathematical-equation#equation-array)

- [x]  Fraction
- [x]  Function
- [ ]  Group character
- [x]  Limit
- [x]  Matrix
- [x]  **N-Array**
- [x]  Radical
- [ ]  Phantom

[https://help.syncfusion.com/file-formats/docio/working-with-mathematical-equation#phantom](https://help.syncfusion.com/file-formats/docio/working-with-mathematical-equation#phantom)

- [x]  Script

# Usage

Sử dụng các hàm theo mục đích trong class Utils:

- ProcessToFile(inputFilePath: string, outputFilePath: string): void **:** xuất file txt từ file docx
- ReadFormula(document: WordDocument): MathSection[]: lấy ra các MathSection, ứng với từng section trong 1 WordDocument.
- ConvertDocument(mathSection: List<MathSection>): void **:** chuyển đổi math block trong math section nhận vào thành các math pattern trong latex.
- ExtractRawText(document: WordDocument, Mathsection[]): WordDocument **:** chuyển đổi các math block trong document thành math pattern trong latex. Định dạng này sẵn sàng render trên web bằng math-jax