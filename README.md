# Word Image Extractor

A lightweight Windows utility for extracting embedded images from Microsoft Word documents and packaging them into a clean ZIP archive in document order.

## What it does

- Supports `.docx`, `.docm`, and legacy `.doc` files.
- Extracts each image occurrence, not only each unique media asset.
- Preserves the original embedded bytes by default for maximum available quality.
- Optional conversion to PNG, JPEG, TIFF, or BMP without intentional resizing.
- Uses Microsoft Word's own pagination engine for accurate page mapping when Word is installed.
- Safe fallback to document order when page mapping cannot be proven reliable.
- Handles inline and floating picture objects in the main document body.
- Optional one-time extraction of header/footer image assets.
- Creates CSV and JSON manifests with page, source asset, output path, dimensions, and status.
- Writes output incrementally to ZIP so large documents do not need to be held fully in RAM.
- English / Vietnamese UI toggle.

## Quality behavior

`Original - Best Quality` is the recommended mode. It copies the embedded image bytes directly from the Word package without re-encoding.

If the Word document itself already compressed an image before saving, the original pre-Word image cannot be reconstructed because those pixels are no longer stored in the document.

When converting formats, the utility preserves the decoded pixel dimensions and does not intentionally upscale. If a format cannot be decoded or converted, the app keeps the original image instead of dropping it.

## Page mapping safety

DOCX files do not store a simple page number for every image. Page numbers depend on Word's layout engine.

In `Accurate - Microsoft Word` mode, the utility opens the document read-only with macros disabled, asks Word for the page of each picture object, and applies those page numbers only when the picture-object count exactly matches the image-instance count found in the DOCX package.

If the counts do not match, the app does **not guess**. It keeps document order and records the mismatch as a warning.

## Output example

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

## Requirements

- Windows 10 or Windows 11.
- .NET 8 Desktop Runtime when using a framework-dependent build.
- Microsoft Word is optional for `.docx` / `.docm` extraction, but required for:
  - accurate page mapping;
  - legacy `.doc` conversion.

No Python and no external NuGet packages are required by the source project.

## Build

Install the .NET 8 SDK, then run:

```bat
BUILD_AND_RUN.bat
```

To create a self-contained Windows x64 build:

```bat
PUBLISH_PORTABLE.bat
```

The published files will be placed in:

```text
publish\win-x64\
```

## Large-document design

The extractor does not unpack the entire DOCX into memory. It reads the Word ZIP package directly, processes one image instance at a time, and writes each result directly into the destination ZIP.

For original-format extraction, very large images are streamed directly. Pixel-dimension probing is skipped for individual assets over 64 MB to avoid a large memory allocation solely for metadata.

Format conversion necessarily decodes one image at a time. If a conversion fails, the original bytes are preserved.

## Safety decisions

- Word documents are opened read-only for pagination.
- Word macro automation is forced off when COM automation is used.
- External linked images are never fetched automatically.
- A page number is never guessed after a mapping mismatch.
- A failed conversion never deletes the source image from the output.
- ZIP output is first written to a temporary partial file and moved into place only after completion.
- Cancellation removes the incomplete partial archive.

## Current scope

V1.0 focuses on embedded picture assets in the main Word story plus optional header/footer assets. Charts, SmartArt, OLE objects, and other non-picture Office objects are intentionally not rendered as images because doing so would require a different rendering pipeline and could reduce fidelity.

## Developer

**Nguyen Khang**  
Copyright © 2026 Nguyen Khang. All Rights Reserved.

Facebook: https://www.facebook.com/ngkph.m
