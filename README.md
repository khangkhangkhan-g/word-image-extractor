<div align="center">

# Word Image Extractor

### Extract every image from Word documents, preserve quality, keep document order, and export everything into a clean ZIP.

[English](#english) · [Tiếng Việt](#tiếng-việt)

![Version](https://img.shields.io/badge/version-1.0.2-111827?style=flat-square)
![Platform](https://img.shields.io/badge/platform-Windows-111827?style=flat-square)
![Framework](https://img.shields.io/badge/.NET-8-111827?style=flat-square)
![UI](https://img.shields.io/badge/UI-WPF-111827?style=flat-square)
![Language](https://img.shields.io/badge/UI-English%20%2F%20Vietnamese-111827?style=flat-square)

</div>

---

# English

## What is Word Image Extractor?

**Word Image Extractor** is a Windows utility for extracting images from Microsoft Word documents into a structured ZIP archive.

Instead of manually saving pictures one by one, the app scans the document, finds image instances, preserves their original embedded quality whenever possible, keeps them in document order, and exports them with clean filenames and optional page-based folders.

It is designed for documents with just a few pictures as well as large Word files containing hundreds of pages and many embedded assets.

> The goal is simple: **add a Word file, choose how you want the images exported, and let the app organize the result for you.**

---

## Highlights

- Extract images from `.docx` and `.docm`
- Optional `.doc` support through Microsoft Word conversion
- Preserve original embedded image bytes in **Original / Best Quality** mode
- Export to PNG, JPEG, TIFF, or BMP
- Adjustable JPEG quality
- Keep images in document order
- Group output by Word page when accurate pagination is available
- Detect repeated image instances even when Word reuses the same internal asset
- Export image metadata
- Generate CSV and JSON manifests
- Support inline and floating pictures
- Optional header and footer image extraction
- EN / VI interface toggle
- Large-document processing with streaming-oriented output
- Progress indicator and cancel support
- Automatic fallback to the original file when conversion fails

---

## Why not just open `word/media/`?

A `.docx` file is internally a ZIP package, and its images usually live under:

```text
word/media/
```

But simply extracting that folder is not enough.

A single media file can appear multiple times in the document, while page numbers and visual order are determined by Word's layout engine.

Word Image Extractor separates:

```text
IMAGE ASSET
from
IMAGE INSTANCE
```

For example:

```text
word/media/image1.png

Used on:
- Page 1
- Page 7
- Page 12
```

The application can preserve those separate appearances in the exported result instead of treating them as only one image.

---

## Output quality

### Original / Best Quality

This is the recommended default.

The app copies the embedded image bytes directly whenever possible, which means it does **not** re-screenshot the picture from the Word page.

Example:

```text
Embedded source: 4032 × 3024 JPEG
Displayed in Word: 900 × 675

Exported result:
4032 × 3024 JPEG
```

If Word already compressed the image before the document was saved, the app cannot restore information that is no longer present in the file.

---

## Supported output formats

```text
Original / Best Quality
PNG
JPEG
TIFF
BMP
```

JPEG export supports adjustable quality.

If a conversion fails, the application keeps the original embedded image instead of dropping the asset.

---

## Page detection modes

### Accurate - Microsoft Word

Recommended when page accuracy matters.

Requires Microsoft Word to be installed.

Uses Word's own layout engine to determine where image instances appear in the document.

Best for:

- long reports
- theses
- documents with floating images
- complex tables
- section breaks
- documents where page order matters

### Portable - Document Order

Does not require Microsoft Word.

Preserves document order, but page numbers may not match Word exactly because `.docx` files do not store final pagination as a simple fixed field.

Best for:

- quick extraction
- machines without Word
- cases where sequence matters more than exact page numbers

---

## Example ZIP structure

### Group by page

```text
Document_Images.zip
├── README_FIRST.txt
├── images/
│   ├── Page_0001/
│   │   ├── 00001_P0001_I001.jpg
│   │   └── 00002_P0001_I002.png
│   ├── Page_0002/
│   │   └── 00003_P0002_I001.png
│   └── Static/
│       ├── Headers/
│       └── Footers/
└── metadata/
    ├── manifest.csv
    └── manifest.json
```

### Flat sequential mode

```text
images/
├── 00001_P0001_I001.jpg
├── 00002_P0001_I002.png
├── 00003_P0002_I001.png
└── ...
```

---

## Manifest files

The app can generate:

```text
metadata/manifest.csv
metadata/manifest.json
```

These files can contain information such as:

- sequence number
- page number
- image index
- original media name
- exported filename
- output format
- embedded resolution
- file size
- extraction warnings

This makes the output easier to audit or process in later automation workflows.

---

## Large documents

The application is designed to avoid loading every image into memory at once.

The intended processing flow is:

```text
Open document
      ↓
Index image instances
      ↓
Process one image
      ↓
Write to ZIP
      ↓
Release temporary data
      ↓
Continue
```

This makes the workflow more suitable for documents containing many pages or large embedded assets.

For very large images, the application may preserve the original file instead of performing a memory-heavy conversion.

---

## User interface

The interface is intentionally minimal.

Main controls include:

- Add file / drag and drop
- EN / VI language toggle
- Output format selector
- JPEG quality control
- Organization mode
- Accurate / Portable page detection
- Original-resolution preservation
- CSV manifest
- JSON manifest
- Header / footer extraction
- Progress bar
- Cancel
- Open output folder

---

## Quick start

### 1. Build and run

After extracting the project folder, run:

```text
BUILD_AND_RUN.bat
```

Requires the **.NET 8 SDK**.

### 2. Add a document

Drag a Word file into the application or use **Add file**.

### 3. Choose export settings

Recommended default:

```text
Format: Original / Best Quality
Organization: Group by page
Page mode: Accurate - Microsoft Word
Preserve original resolution: ON
Manifest CSV: ON
Manifest JSON: ON
```

### 4. Extract

Click **Extract Images** and choose the output location.

### 5. Open the result

The application creates a ZIP archive containing the extracted images and metadata.

---

## Portable Windows build

Run:

```text
PUBLISH_PORTABLE.bat
```

The published application is created under:

```text
publish/win-x64/
```

---

## Requirements

### Minimum

- Windows 10 or newer
- `.NET 8` runtime or SDK depending on build method

### For Accurate Page Mode

- Microsoft Word desktop application

### For legacy `.doc`

Microsoft Word is required because the application converts the old binary Word format before extraction.

---

## Supported Word objects

Current extraction logic focuses on picture-based content, including:

- inline pictures
- floating pictures
- pictures inside tables
- header images
- footer images
- raster assets
- supported vector assets such as SVG when available

Some Word content is fundamentally different from an embedded picture and may require rendering rather than extraction, for example:

- charts
- SmartArt
- OLE objects
- complex grouped shapes
- some legacy vector formats
- externally linked images

These should not be treated as guaranteed image assets unless the document actually contains extractable media.

---

## Accuracy and safety behavior

The application is designed to prefer a correct fallback over a misleading result.

Examples:

- If image conversion fails, preserve the original image.
- If accurate page mapping cannot be validated, preserve document order instead of inventing page numbers.
- If the same source asset is reused multiple times, keep separate image instances when possible.
- If an SVG and fallback raster image represent the same visual object, avoid unnecessary duplicates when possible.
- If processing is cancelled or fails before completion, temporary partial archives should not be treated as final output.

---

## Privacy

Word Image Extractor is designed as a local Windows utility.

The document and extracted images are processed locally on the user's machine.

The project does not require an online account or cloud service for normal extraction.

When Microsoft Word automation is used, documents are opened in a controlled read-only workflow with external update behavior minimized.

---

## Known limitations

- Exact Word pagination requires Microsoft Word.
- `.doc` support depends on Microsoft Word conversion.
- Images already compressed inside the Word file cannot be restored to their pre-compression quality.
- Some charts, SmartArt, OLE objects, or unusual drawing objects are not equivalent to standard embedded images.
- Extremely complex Word documents may expose layout differences between environments.
- Externally linked images may not be embedded inside the document package and therefore may not be available for direct extraction.

---

## Project structure

```text
WordImageExtractor/
├── WordImageExtractor.sln
├── WordImageExtractor.csproj
├── MainWindow.xaml
├── MainWindow.xaml.cs
├── Models/
├── Services/
├── Utilities/
├── BUILD_AND_RUN.bat
├── PUBLISH_PORTABLE.bat
├── START_HERE.txt
├── CHANGELOG.md
├── README.md
├── VALIDATION.txt
└── LICENSE
```

---

## Development philosophy

Word Image Extractor is built around four principles:

```text
Preserve quality
Preserve order
Avoid silent data loss
Keep the workflow simple
```

The application should prefer a transparent warning or safe fallback over silently dropping an image or producing a misleading page mapping.

---

## Troubleshooting

### The app cannot determine exact pages

Use **Accurate - Microsoft Word** mode and make sure Microsoft Word is installed.

### A converted image is missing

Check whether the original format was preserved as a fallback.

### A picture appears more than once

This may be intentional. The same embedded asset can appear multiple times in the Word document.

### `.doc` does not work

Legacy `.doc` conversion requires Microsoft Word desktop.

### Build fails

Make sure the .NET 8 SDK is installed, then run:

```text
dotnet --version
```

and rebuild with:

```text
BUILD_AND_RUN.bat
```

---

## Version

Current documented version:

```text
1.0.2
```

---

## Author

**Nguyen Khang**

- GitHub: [khangkhangkhan-g](https://github.com/khangkhangkhan-g)
- Facebook: [ngkph.m](https://www.facebook.com/ngkph.m)

---

## Copyright

**Copyright © 2026 Nguyen Khang. All Rights Reserved.**

See the included `LICENSE` file for the licensing terms distributed with the project.

---

<div align="center">

[↑ Back to top](#word-image-extractor)

</div>

---

# Tiếng Việt

## Word Image Extractor là gì?

**Word Image Extractor** là một ứng dụng Windows dùng để trích xuất hình ảnh từ tài liệu Microsoft Word và đóng gói chúng thành một file ZIP có cấu trúc rõ ràng.

Thay vì phải lưu từng ảnh thủ công, ứng dụng sẽ quét tài liệu, phát hiện các lần ảnh xuất hiện, cố gắng giữ nguyên chất lượng ảnh nhúng gốc, giữ đúng thứ tự trong tài liệu và xuất ảnh với tên file dễ quản lý.

Ứng dụng được thiết kế để xử lý cả file Word nhỏ lẫn tài liệu dài hàng trăm trang với số lượng ảnh lớn.

> Mục tiêu rất đơn giản: **thêm file Word, chọn cách xuất ảnh, và để ứng dụng tự xử lý phần còn lại.**

---

## Tính năng nổi bật

- Trích xuất ảnh từ `.docx` và `.docm`
- Hỗ trợ `.doc` thông qua Microsoft Word khi cần chuyển đổi
- Giữ nguyên byte ảnh nhúng với chế độ **Original / Best Quality**
- Chuyển sang PNG, JPEG, TIFF hoặc BMP
- Tùy chỉnh chất lượng JPEG
- Giữ đúng thứ tự ảnh trong tài liệu
- Có thể nhóm ảnh theo trang Word
- Phân biệt nhiều lần xuất hiện của cùng một ảnh nguồn
- Xuất metadata hình ảnh
- Tạo manifest CSV và JSON
- Hỗ trợ ảnh inline và floating
- Có thể lấy ảnh trong header/footer
- Toggle giao diện Anh / Việt
- Phù hợp hơn với tài liệu lớn nhờ cách xử lý theo luồng
- Có progress bar và nút Cancel
- Nếu convert ảnh thất bại, ứng dụng giữ lại file gốc thay vì bỏ mất ảnh

---

## Vì sao không chỉ lấy ảnh trong `word/media/`?

File `.docx` thực chất là một ZIP package và ảnh thường nằm trong:

```text
word/media/
```

Nhưng chỉ lấy folder này là chưa đủ.

Một file ảnh nội bộ có thể được sử dụng nhiều lần ở nhiều vị trí khác nhau trong tài liệu.

Ứng dụng phân biệt:

```text
IMAGE ASSET
và
IMAGE INSTANCE
```

Ví dụ:

```text
word/media/image1.png

Xuất hiện tại:
- Trang 1
- Trang 7
- Trang 12
```

Word Image Extractor có thể giữ các lần xuất hiện này như những instance riêng trong kết quả export.

---

## Chất lượng ảnh

### Original / Best Quality

Đây là chế độ mặc định được khuyên dùng.

Ứng dụng cố gắng copy trực tiếp byte ảnh đã được nhúng bên trong Word thay vì chụp lại ảnh từ giao diện Word.

Ví dụ:

```text
Ảnh nhúng: 4032 × 3024 JPEG
Hiển thị trong Word: 900 × 675

Kết quả export:
4032 × 3024 JPEG
```

Nếu Word đã nén ảnh trước khi tài liệu được lưu thì dữ liệu chất lượng cao hơn đã mất khỏi file và không thể phục hồi lại bằng extraction thông thường.

---

## Định dạng output

```text
Original / Best Quality
PNG
JPEG
TIFF
BMP
```

JPEG có thể điều chỉnh quality.

Nếu quá trình convert lỗi, ứng dụng giữ lại định dạng gốc thay vì làm mất ảnh.

---

## Chế độ xác định trang

### Accurate - Microsoft Word

Khuyên dùng khi số trang phải chính xác.

Yêu cầu Microsoft Word được cài trên máy.

Ứng dụng sử dụng layout engine của Word để xác định vị trí ảnh theo trang.

Phù hợp cho:

- luận văn
- báo cáo dài
- ảnh floating
- bảng phức tạp
- section break
- tài liệu cần giữ đúng page order

### Portable - Document Order

Không cần Microsoft Word.

Giữ đúng thứ tự tài liệu nhưng số trang có thể không trùng hoàn toàn với Word vì `.docx` không lưu pagination cuối cùng dưới dạng một field đơn giản.

Phù hợp cho:

- extract nhanh
- máy không có Word
- trường hợp cần đúng thứ tự hơn là đúng page tuyệt đối

---

## Cấu trúc ZIP mẫu

### Group by page

```text
Document_Images.zip
├── README_FIRST.txt
├── images/
│   ├── Page_0001/
│   │   ├── 00001_P0001_I001.jpg
│   │   └── 00002_P0001_I002.png
│   ├── Page_0002/
│   │   └── 00003_P0002_I001.png
│   └── Static/
│       ├── Headers/
│       └── Footers/
└── metadata/
    ├── manifest.csv
    └── manifest.json
```

### Flat sequential mode

```text
images/
├── 00001_P0001_I001.jpg
├── 00002_P0001_I002.png
├── 00003_P0002_I001.png
└── ...
```

---

## Manifest

Ứng dụng có thể tạo:

```text
metadata/manifest.csv
metadata/manifest.json
```

Các file này có thể chứa:

- số thứ tự
- số trang
- image index
- tên media gốc
- tên file export
- định dạng output
- resolution nhúng
- kích thước file
- warning trong quá trình extraction

Điều này giúp kiểm tra output hoặc dùng tiếp cho automation.

---

## File Word lớn

Ứng dụng được thiết kế để tránh load toàn bộ ảnh vào RAM cùng lúc.

Luồng xử lý dự kiến:

```text
Mở tài liệu
      ↓
Index image instances
      ↓
Xử lý một ảnh
      ↓
Ghi vào ZIP
      ↓
Giải phóng dữ liệu tạm
      ↓
Tiếp tục
```

Cách này phù hợp hơn với tài liệu dài hoặc chứa nhiều ảnh dung lượng lớn.

Với ảnh cực lớn, ứng dụng có thể giữ original thay vì cố convert và tiêu tốn quá nhiều RAM.

---

## Giao diện

Giao diện được giữ tối giản.

Các control chính:

- Add file / kéo thả
- Toggle EN / VI
- Chọn output format
- JPEG quality
- Chọn cách tổ chức file
- Accurate / Portable page mode
- Preserve original resolution
- Manifest CSV
- Manifest JSON
- Header / footer extraction
- Progress bar
- Cancel
- Open output folder

---

## Cách sử dụng nhanh

### 1. Build và chạy

Sau khi giải nén project:

```text
BUILD_AND_RUN.bat
```

Yêu cầu **.NET 8 SDK**.

### 2. Thêm file Word

Kéo file vào app hoặc bấm **Add file**.

### 3. Chọn setting

Setting được khuyên dùng:

```text
Format: Original / Best Quality
Organization: Group by page
Page mode: Accurate - Microsoft Word
Preserve original resolution: ON
Manifest CSV: ON
Manifest JSON: ON
```

### 4. Extract

Bấm **Extract Images** và chọn nơi lưu.

### 5. Mở kết quả

App sẽ tạo ZIP chứa hình ảnh và metadata.

---

## Build bản portable cho Windows

Chạy:

```text
PUBLISH_PORTABLE.bat
```

Output nằm tại:

```text
publish/win-x64/
```

---

## Yêu cầu hệ thống

### Tối thiểu

- Windows 10 trở lên
- `.NET 8` runtime hoặc SDK tùy cách chạy

### Accurate Page Mode

- Microsoft Word desktop

### Legacy `.doc`

Microsoft Word là bắt buộc vì ứng dụng cần convert format Word cũ trước khi trích xuất.

---

## Các loại object được xử lý

Extraction hiện tập trung vào picture-based content như:

- inline picture
- floating picture
- ảnh trong table
- ảnh header
- ảnh footer
- raster image
- một số vector asset như SVG khi có thể lấy trực tiếp

Một số object Word không tương đương với ảnh nhúng thông thường và có thể cần render riêng:

- chart
- SmartArt
- OLE object
- grouped shape phức tạp
- một số legacy vector format
- externally linked image

---

## Nguyên tắc an toàn dữ liệu

Ứng dụng ưu tiên fallback đúng hơn là output sai nhưng trông có vẻ hợp lệ.

Ví dụ:

- Convert lỗi thì giữ ảnh gốc.
- Không xác định được page chính xác thì giữ document order thay vì bịa số trang.
- Cùng một asset xuất hiện nhiều lần thì cố giữ các instance riêng.
- SVG có raster fallback thì cố tránh duplicate không cần thiết.
- File partial chưa hoàn thành không được xem là output cuối cùng.

---

## Privacy

Word Image Extractor được thiết kế như một Windows utility chạy local.

Tài liệu và hình ảnh được xử lý trên máy người dùng.

App không yêu cầu tài khoản online hoặc cloud service cho extraction thông thường.

Khi dùng Microsoft Word automation, tài liệu được mở theo workflow read-only và hạn chế external update khi có thể.

---

## Giới hạn hiện tại

- Exact pagination cần Microsoft Word.
- `.doc` cần Word để convert.
- Ảnh đã bị Word nén thì không thể phục hồi chất lượng trước khi nén.
- Chart, SmartArt, OLE hoặc drawing object đặc biệt không phải lúc nào cũng có media file trực tiếp.
- Tài liệu Word cực kỳ phức tạp có thể có khác biệt layout giữa môi trường.
- Externally linked images có thể không nằm trong package nên không thể extract như embedded image.

---

## Cấu trúc project

```text
WordImageExtractor/
├── WordImageExtractor.sln
├── WordImageExtractor.csproj
├── MainWindow.xaml
├── MainWindow.xaml.cs
├── Models/
├── Services/
├── Utilities/
├── BUILD_AND_RUN.bat
├── PUBLISH_PORTABLE.bat
├── START_HERE.txt
├── CHANGELOG.md
├── README.md
├── VALIDATION.txt
└── LICENSE
```

---

## Triết lý phát triển

Word Image Extractor xoay quanh bốn nguyên tắc:

```text
Giữ chất lượng
Giữ thứ tự
Không âm thầm làm mất dữ liệu
Giữ workflow đơn giản
```

Nếu có vấn đề, ứng dụng nên báo rõ hoặc fallback an toàn thay vì âm thầm bỏ ảnh hay gán sai trang.

---

## Troubleshooting

### App không xác định được đúng trang

Dùng **Accurate - Microsoft Word** và kiểm tra Microsoft Word đã được cài.

### Ảnh convert bị thiếu

Kiểm tra xem app có giữ lại original làm fallback hay không.

### Một hình xuất hiện nhiều lần

Điều này có thể đúng vì cùng một ảnh nguồn có thể được dùng nhiều lần trong Word.

### `.doc` không hoạt động

Legacy `.doc` cần Microsoft Word desktop để convert.

### Build lỗi

Kiểm tra .NET 8 SDK:

```text
dotnet --version
```

sau đó chạy lại:

```text
BUILD_AND_RUN.bat
```

---

## Phiên bản

Phiên bản hiện tại được mô tả trong README:

```text
1.0.2
```

---

## Tác giả

**Nguyen Khang**

- GitHub: [khangkhangkhan-g](https://github.com/khangkhangkhan-g)
- Facebook: [ngkph.m](https://www.facebook.com/ngkph.m)

---

## Bản quyền

**Copyright © 2026 Nguyen Khang. All Rights Reserved.**

Xem file `LICENSE` đi kèm project để biết điều khoản sử dụng.

---

<div align="center">

[↑ Back to top](#word-image-extractor)

</div>
